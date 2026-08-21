"""Taeha material-dispensing control reference simulator."""

from .model import (
    ActuatorKind,
    AdaptiveCompensator,
    ComponentProfile,
    MaterialProfile,
    SensorProfile,
    ShotResult,
    simulate_shot,
)
from .persistence import DualSlotCalibrationStore

__all__ = [
    "ActuatorKind",
    "AdaptiveCompensator",
    "ComponentProfile",
    "DualSlotCalibrationStore",
    "MaterialProfile",
    "SensorProfile",
    "ShotResult",
    "simulate_shot",
]
