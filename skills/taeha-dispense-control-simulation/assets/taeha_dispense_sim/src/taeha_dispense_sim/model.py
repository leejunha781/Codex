from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field
from enum import Enum
import math
import random
from typing import Iterable


class ActuatorKind(str, Enum):
    SERVO = "servo"
    STEPPER = "stepper"
    BRUSHED_DC = "brushed_dc"
    BLDC_AC = "bldc_ac"


@dataclass(frozen=True)
class ActuatorProfile:
    kind: ActuatorKind = ActuatorKind.SERVO
    time_constant_s: float = 0.040
    command_deadband: float = 0.0
    command_quantum: float = 0.0
    max_slew_per_s: float = 30.0

    @staticmethod
    def for_kind(kind: ActuatorKind) -> "ActuatorProfile":
        presets = {
            ActuatorKind.SERVO: ActuatorProfile(kind, 0.035, 0.0, 0.0, 35.0),
            ActuatorKind.STEPPER: ActuatorProfile(kind, 0.020, 0.0, 0.005, 50.0),
            ActuatorKind.BRUSHED_DC: ActuatorProfile(kind, 0.085, 0.10, 0.0, 14.0),
            ActuatorKind.BLDC_AC: ActuatorProfile(kind, 0.050, 0.03, 0.0, 24.0),
        }
        return presets[kind]


@dataclass(frozen=True)
class ComponentProfile:
    density_mg_per_ul: float
    feedforward_flow_mg_s: float
    actual_flow_scale: float
    transport_delay_s: float
    fluid_time_constant_s: float
    dribble_fraction: float
    dribble_time_constant_s: float


@dataclass(frozen=True)
class SensorProfile:
    sample_period_s: float = 0.010
    delay_s: float = 0.120
    noise_std_mg: float = 1.0
    settle_time_s: float = 0.600


@dataclass(frozen=True)
class MaterialProfile:
    key: str
    reference_temperature_c: float
    temperature_viscosity_coefficient_per_c: float
    nozzle_factor: float
    component_a: ComponentProfile
    component_b: ComponentProfile
    sensor: SensorProfile
    min_target_mass_mg: float = 100.0
    max_target_mass_mg: float = 10000.0
    ratio_tolerance_percent: float = 5.0
    total_mass_learning_tolerance_percent: float = 10.0
    max_shot_time_s: float = 15.0


@dataclass
class TraceSample:
    time_s: float
    command_a: float
    command_b: float
    true_mass_a_mg: float
    true_mass_b_mg: float
    measured_mass_a_mg: float
    measured_mass_b_mg: float


@dataclass
class ShotResult:
    accepted: bool
    rejection_reason: str | None
    target_mass_a_mg: float
    target_mass_b_mg: float
    measured_mass_a_mg: float
    measured_mass_b_mg: float
    true_mass_a_mg: float
    true_mass_b_mg: float
    total_error_percent: float
    ratio_error_percent: float
    confidence: float
    trace: list[TraceSample] = field(default_factory=list)


@dataclass
class AdaptiveCompensator:
    profile_key: str
    gain_a: float = 1.0
    gain_b: float = 1.0
    confidence: float = 0.0
    accepted_samples: int = 0
    alpha: float = 0.18
    min_gain: float = 0.75
    max_gain: float = 1.25
    max_delta_per_shot: float = 0.025

    def planned_gains(self, profile_key: str) -> tuple[float, float]:
        if profile_key != self.profile_key:
            return (1.0, 1.0)
        confidence = _clamp(self.confidence, 0.0, 1.0)
        return (
            1.0 + confidence * (self.gain_a - 1.0),
            1.0 + confidence * (self.gain_b - 1.0),
        )

    def update(self, profile_key: str, result: ShotResult) -> bool:
        if profile_key != self.profile_key or not result.accepted:
            return False
        if min(result.target_mass_a_mg, result.target_mass_b_mg) <= 0.0:
            return False

        next_a = self.gain_a * (
            1.0
            + self.alpha
            * (result.target_mass_a_mg - result.measured_mass_a_mg)
            / result.target_mass_a_mg
        )
        next_b = self.gain_b * (
            1.0
            + self.alpha
            * (result.target_mass_b_mg - result.measured_mass_b_mg)
            / result.target_mass_b_mg
        )
        self.gain_a = _rate_and_range_limit(
            next_a, self.gain_a, self.max_delta_per_shot, self.min_gain, self.max_gain
        )
        self.gain_b = _rate_and_range_limit(
            next_b, self.gain_b, self.max_delta_per_shot, self.min_gain, self.max_gain
        )
        self.accepted_samples += 1
        self.confidence = min(1.0, self.accepted_samples / 12.0)
        return True


def _rate_and_range_limit(
    candidate: float, previous: float, max_delta: float, low: float, high: float
) -> float:
    candidate = max(previous - max_delta, min(previous + max_delta, candidate))
    return max(low, min(high, candidate))


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _temperature_flow_factor(profile: MaterialProfile, temperature_c: float) -> float:
    delta = temperature_c - profile.reference_temperature_c
    viscosity_ratio = math.exp(-profile.temperature_viscosity_coefficient_per_c * delta)
    return profile.nozzle_factor / math.sqrt(max(0.20, viscosity_ratio))


class _ComponentPlant:
    def __init__(
        self,
        component: ComponentProfile,
        actuator: ActuatorProfile,
        dt_s: float,
        flow_factor: float,
    ) -> None:
        self.component = component
        self.actuator = actuator
        self.dt_s = dt_s
        self.flow_factor = flow_factor
        self.command_state = 0.0
        self.actuator_response_state = 0.0
        self.flow_state_mg_s = 0.0
        self.mass_mg = 0.0
        self._closing_flow_mg_s = 0.0
        self._previous_delayed_command = 0.0
        delay_steps = max(0, round(component.transport_delay_s / dt_s))
        self._command_delay: deque[float] = deque([0.0] * (delay_steps + 1))

    def step(self, requested_command: float) -> None:
        requested_command = self._shape_command(requested_command)
        max_step = self.actuator.max_slew_per_s * self.dt_s
        self.command_state += _clamp(
            requested_command - self.command_state, -max_step, max_step
        )
        actuator_tau = max(self.dt_s, self.actuator.time_constant_s)
        self.actuator_response_state += (
            self.command_state - self.actuator_response_state
        ) * self.dt_s / actuator_tau
        self._command_delay.append(self.actuator_response_state)
        delayed_command = self._command_delay.popleft()

        target_flow = (
            self.component.feedforward_flow_mg_s
            * self.component.actual_flow_scale
            * self.flow_factor
            * delayed_command
        )
        tau = max(self.dt_s, self.component.fluid_time_constant_s)
        self.flow_state_mg_s += (target_flow - self.flow_state_mg_s) * self.dt_s / tau

        closing_started = (
            requested_command <= 0.0
            and delayed_command < self._previous_delayed_command - 1e-9
        )
        if closing_started:
            if self._closing_flow_mg_s == 0.0:
                self._closing_flow_mg_s = self.flow_state_mg_s * self.component.dribble_fraction
            dribble_tau = max(self.dt_s, self.component.dribble_time_constant_s)
            self._closing_flow_mg_s *= math.exp(-self.dt_s / dribble_tau)
        elif requested_command > 0.0:
            self._closing_flow_mg_s = 0.0

        delivered_flow = max(0.0, self.flow_state_mg_s + self._closing_flow_mg_s)
        self.mass_mg += delivered_flow * self.dt_s
        self._previous_delayed_command = delayed_command

    def _shape_command(self, command: float) -> float:
        command = _clamp(command, 0.0, 1.0)
        if command <= self.actuator.command_deadband:
            command = 0.0
        elif self.actuator.command_deadband > 0.0:
            command = (command - self.actuator.command_deadband) / (
                1.0 - self.actuator.command_deadband
            )
        if self.actuator.command_quantum > 0.0:
            command = round(command / self.actuator.command_quantum) * self.actuator.command_quantum
        return _clamp(command, 0.0, 1.0)


def _targets(total_target_mg: float, ratio_a_to_b: float) -> tuple[float, float]:
    if total_target_mg <= 0.0 or ratio_a_to_b <= 0.0:
        raise ValueError("target mass and A:B ratio must be positive")
    target_b = total_target_mg / (ratio_a_to_b + 1.0)
    return total_target_mg - target_b, target_b


def simulate_shot(
    profile: MaterialProfile,
    total_target_mg: float,
    ratio_a_to_b: float,
    temperature_c: float,
    compensator: AdaptiveCompensator | None = None,
    actuator_a: ActuatorKind = ActuatorKind.SERVO,
    actuator_b: ActuatorKind = ActuatorKind.SERVO,
    seed: int = 1,
    dt_s: float = 0.005,
    capture_trace: bool = False,
) -> ShotResult:
    if not profile.min_target_mass_mg <= total_target_mg <= profile.max_target_mass_mg:
        raise ValueError("target mass is outside the validated profile envelope")
    if dt_s <= 0.0 or dt_s > profile.sensor.sample_period_s:
        raise ValueError("dt_s must be positive and no larger than the sensor sample period")

    rng = random.Random(seed)
    target_a, target_b = _targets(total_target_mg, ratio_a_to_b)
    gain_a, gain_b = (1.0, 1.0)
    confidence = 0.0
    if compensator is not None:
        gain_a, gain_b = compensator.planned_gains(profile.key)
        confidence = compensator.confidence

    flow_factor = _temperature_flow_factor(profile, temperature_c)
    planned_flow_a = profile.component_a.feedforward_flow_mg_s * flow_factor
    planned_flow_b = profile.component_b.feedforward_flow_mg_s * flow_factor
    duration_a = target_a / max(1e-9, planned_flow_a) * gain_a
    duration_b = target_b / max(1e-9, planned_flow_b) * gain_b

    plant_a = _ComponentPlant(
        profile.component_a, ActuatorProfile.for_kind(actuator_a), dt_s, flow_factor
    )
    plant_b = _ComponentPlant(
        profile.component_b, ActuatorProfile.for_kind(actuator_b), dt_s, flow_factor
    )

    sensor_delay_steps = max(0, round(profile.sensor.delay_s / dt_s))
    sensed_a: deque[float] = deque([0.0] * (sensor_delay_steps + 1))
    sensed_b: deque[float] = deque([0.0] * (sensor_delay_steps + 1))
    measured_a = 0.0
    measured_b = 0.0
    trace: list[TraceSample] = []
    active_duration = max(duration_a, duration_b)
    end_time = min(
        profile.max_shot_time_s,
        active_duration + profile.sensor.settle_time_s + profile.sensor.delay_s,
    )

    t = 0.0
    while t <= end_time + 1e-12:
        command_a = 1.0 if t < duration_a else 0.0
        command_b = 1.0 if t < duration_b else 0.0
        plant_a.step(command_a)
        plant_b.step(command_b)
        sensed_a.append(plant_a.mass_mg)
        sensed_b.append(plant_b.mass_mg)
        delayed_a = sensed_a.popleft()
        delayed_b = sensed_b.popleft()
        measured_a = max(0.0, delayed_a + rng.gauss(0.0, profile.sensor.noise_std_mg))
        measured_b = max(0.0, delayed_b + rng.gauss(0.0, profile.sensor.noise_std_mg))
        if capture_trace and (
            not trace or t - trace[-1].time_s >= profile.sensor.sample_period_s - 1e-12
        ):
            trace.append(
                TraceSample(
                    time_s=t,
                    command_a=command_a,
                    command_b=command_b,
                    true_mass_a_mg=plant_a.mass_mg,
                    true_mass_b_mg=plant_b.mass_mg,
                    measured_mass_a_mg=measured_a,
                    measured_mass_b_mg=measured_b,
                )
            )
        t += dt_s

    measured_total = measured_a + measured_b
    total_error_percent = 100.0 * (measured_total - total_target_mg) / total_target_mg
    measured_ratio = measured_a / max(1e-9, measured_b)
    ratio_error_percent = 100.0 * (measured_ratio - ratio_a_to_b) / ratio_a_to_b

    rejection_reason: str | None = None
    if not all(math.isfinite(item) for item in (measured_a, measured_b, measured_total)):
        rejection_reason = "sensor_invalid"
    elif end_time >= profile.max_shot_time_s - 1e-9:
        rejection_reason = "shot_timeout"
    elif abs(total_error_percent) > profile.total_mass_learning_tolerance_percent:
        rejection_reason = "total_mass_out_of_learning_window"
    elif abs(ratio_error_percent) > profile.ratio_tolerance_percent:
        rejection_reason = "ratio_out_of_tolerance"

    return ShotResult(
        accepted=rejection_reason is None,
        rejection_reason=rejection_reason,
        target_mass_a_mg=target_a,
        target_mass_b_mg=target_b,
        measured_mass_a_mg=measured_a,
        measured_mass_b_mg=measured_b,
        true_mass_a_mg=plant_a.mass_mg,
        true_mass_b_mg=plant_b.mass_mg,
        total_error_percent=total_error_percent,
        ratio_error_percent=ratio_error_percent,
        confidence=confidence,
        trace=trace,
    )


def summarize(values: Iterable[float]) -> dict[str, float]:
    data = list(values)
    if not data:
        return {"mean": math.nan, "stdev": math.nan, "min": math.nan, "max": math.nan}
    mean = sum(data) / len(data)
    variance = sum((item - mean) ** 2 for item in data) / max(1, len(data) - 1)
    return {
        "mean": mean,
        "stdev": math.sqrt(variance),
        "min": min(data),
        "max": max(data),
    }
