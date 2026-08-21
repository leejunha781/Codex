from __future__ import annotations

from dataclasses import replace
import json
from pathlib import Path
import tempfile
import unittest

from taeha_dispense_sim.model import (
    ActuatorKind,
    AdaptiveCompensator,
    ShotResult,
    simulate_shot,
)
from taeha_dispense_sim.persistence import DualSlotCalibrationStore
from taeha_dispense_sim.runner import load_profile


ROOT = Path(__file__).resolve().parents[1]
PROFILE = load_profile(ROOT / "configs" / "material_profiles.json", "epoxy-demo-25c-nozzle-a")


class SimulationTests(unittest.TestCase):
    def test_bounded_adaptation_converges_and_never_leaves_limits(self) -> None:
        compensator = AdaptiveCompensator(PROFILE.key)
        errors: list[float] = []
        for shot in range(1, 41):
            result = simulate_shot(
                PROFILE,
                total_target_mg=1200.0,
                ratio_a_to_b=2.0,
                temperature_c=25.0,
                compensator=compensator,
                actuator_a=ActuatorKind.SERVO,
                actuator_b=ActuatorKind.STEPPER,
                seed=shot,
            )
            errors.append(abs(result.total_error_percent))
            compensator.update(PROFILE.key, result)
            self.assertGreaterEqual(compensator.gain_a, compensator.min_gain)
            self.assertLessEqual(compensator.gain_a, compensator.max_gain)
            self.assertGreaterEqual(compensator.gain_b, compensator.min_gain)
            self.assertLessEqual(compensator.gain_b, compensator.max_gain)
        self.assertLess(sum(errors[-10:]) / 10.0, sum(errors[:10]) / 10.0)

    def test_profile_isolation_blocks_cross_profile_learning(self) -> None:
        compensator = AdaptiveCompensator("another-profile", gain_a=1.2, gain_b=0.8)
        self.assertEqual(compensator.planned_gains(PROFILE.key), (1.0, 1.0))
        result = simulate_shot(PROFILE, 1200.0, 2.0, 25.0, compensator=compensator)
        self.assertFalse(compensator.update(PROFILE.key, result))

    def test_target_envelope_is_enforced(self) -> None:
        with self.assertRaises(ValueError):
            simulate_shot(PROFILE, 10.0, 2.0, 25.0)

    def test_high_total_error_with_valid_ratio_cannot_update_learning(self) -> None:
        biased_profile = replace(
            PROFILE,
            component_a=replace(PROFILE.component_a, actual_flow_scale=0.45),
            component_b=replace(PROFILE.component_b, actual_flow_scale=0.47),
        )
        compensator = AdaptiveCompensator(PROFILE.key)
        result = simulate_shot(
            biased_profile, 1200.0, 2.0, 25.0, compensator=compensator
        )
        self.assertEqual(result.rejection_reason, "total_mass_out_of_learning_window")
        self.assertFalse(compensator.update(PROFILE.key, result))

    def test_actuator_time_constants_change_delivered_mass(self) -> None:
        servo = simulate_shot(
            PROFILE, 1200.0, 2.0, 25.0,
            actuator_a=ActuatorKind.SERVO,
            actuator_b=ActuatorKind.SERVO,
        )
        dc = simulate_shot(
            PROFILE, 1200.0, 2.0, 25.0,
            actuator_a=ActuatorKind.BRUSHED_DC,
            actuator_b=ActuatorKind.BRUSHED_DC,
        )
        servo_total = servo.true_mass_a_mg + servo.true_mass_b_mg
        dc_total = dc.true_mass_a_mg + dc.true_mass_b_mg
        self.assertGreater(abs(servo_total - dc_total), 1.0)

    def test_confidence_blends_correction_gain(self) -> None:
        compensator = AdaptiveCompensator(
            PROFILE.key, gain_a=1.2, gain_b=0.8, confidence=0.0
        )
        self.assertEqual(compensator.planned_gains(PROFILE.key), (1.0, 1.0))
        compensator.confidence = 0.5
        self.assertEqual(compensator.planned_gains(PROFILE.key), (1.1, 0.9))

    def test_rejected_ratio_shot_cannot_update_learning(self) -> None:
        compensator = AdaptiveCompensator(PROFILE.key)
        rejected = ShotResult(
            accepted=False,
            rejection_reason="ratio_out_of_tolerance",
            target_mass_a_mg=800.0,
            target_mass_b_mg=400.0,
            measured_mass_a_mg=700.0,
            measured_mass_b_mg=500.0,
            true_mass_a_mg=700.0,
            true_mass_b_mg=500.0,
            total_error_percent=0.0,
            ratio_error_percent=-30.0,
            confidence=0.0,
        )
        self.assertFalse(compensator.update(PROFILE.key, rejected))
        self.assertEqual((compensator.gain_a, compensator.gain_b), (1.0, 1.0))
        self.assertEqual(compensator.accepted_samples, 0)


class PersistenceTests(unittest.TestCase):
    def test_crc_corruption_falls_back_to_older_slot(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            store = DualSlotCalibrationStore(directory)
            first = store.save(PROFILE.key, 1.01, 0.99, 0.4)
            second = store.save(PROFILE.key, 1.02, 0.98, 0.6)
            self.assertGreater(second.sequence, first.sequence)
            store.corrupt_latest(PROFILE.key)
            loaded = store.load(PROFILE.key)
            self.assertIsNotNone(loaded)
            self.assertEqual(loaded.sequence, first.sequence)

    def test_interrupted_write_preserves_current_record(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            store = DualSlotCalibrationStore(directory)
            first = store.save(PROFILE.key, 1.01, 0.99, 0.4)
            returned = store.save(PROFILE.key, 1.02, 0.98, 0.6, simulate_interruption=True)
            self.assertEqual(returned, first)
            self.assertEqual(store.load(PROFILE.key), first)


if __name__ == "__main__":
    unittest.main()
