from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from statistics import mean

import matplotlib.pyplot as plt

from .model import (
    ActuatorKind,
    AdaptiveCompensator,
    ComponentProfile,
    MaterialProfile,
    SensorProfile,
    simulate_shot,
    summarize,
)
from .persistence import DualSlotCalibrationStore


def load_profile(path: Path, key: str) -> MaterialProfile:
    raw = json.loads(path.read_text(encoding="utf-8"))["profiles"][key]
    limits = raw["limits"]
    return MaterialProfile(
        key=key,
        reference_temperature_c=raw["reference_temperature_c"],
        temperature_viscosity_coefficient_per_c=raw[
            "temperature_viscosity_coefficient_per_c"
        ],
        nozzle_factor=raw["nozzle_factor"],
        component_a=ComponentProfile(**raw["component_a"]),
        component_b=ComponentProfile(**raw["component_b"]),
        sensor=SensorProfile(**raw["sensor"]),
        min_target_mass_mg=limits["min_target_mass_mg"],
        max_target_mass_mg=limits["max_target_mass_mg"],
        ratio_tolerance_percent=limits["ratio_tolerance_percent"],
        total_mass_learning_tolerance_percent=limits[
            "total_mass_learning_tolerance_percent"
        ],
        max_shot_time_s=limits["max_shot_time_s"],
    )


def run(args: argparse.Namespace) -> dict[str, object]:
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    profile = load_profile(Path(args.profiles), args.profile)
    compensator = AdaptiveCompensator(profile_key=profile.key)
    rows: list[dict[str, object]] = []

    for shot in range(1, args.shots + 1):
        result = simulate_shot(
            profile=profile,
            total_target_mg=args.target_mass_mg,
            ratio_a_to_b=args.ratio,
            temperature_c=args.temperature_c,
            compensator=compensator if args.adaptive else None,
            actuator_a=ActuatorKind(args.actuator_a),
            actuator_b=ActuatorKind(args.actuator_b),
            seed=args.seed + shot,
            capture_trace=shot == args.shots,
        )
        updated = compensator.update(profile.key, result) if args.adaptive else False
        rows.append(
            {
                "shot": shot,
                "accepted": result.accepted,
                "rejection_reason": result.rejection_reason or "",
                "measured_total_mg": result.measured_mass_a_mg + result.measured_mass_b_mg,
                "total_error_percent": result.total_error_percent,
                "ratio_error_percent": result.ratio_error_percent,
                "gain_a": compensator.gain_a,
                "gain_b": compensator.gain_b,
                "confidence": compensator.confidence,
                "adaptation_updated": updated,
            }
        )
        if shot == args.shots:
            _write_trace(output / "last_shot_trace.csv", result.trace)

    with (output / "shots.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    accepted_rows = [row for row in rows if row["accepted"]]
    summary = {
        "evidence_class": "SIMULATED",
        "profile": profile.key,
        "shots": args.shots,
        "accepted_shots": len(accepted_rows),
        "target_mass_mg": args.target_mass_mg,
        "target_ratio_a_to_b": args.ratio,
        "temperature_c": args.temperature_c,
        "actuator_a": args.actuator_a,
        "actuator_b": args.actuator_b,
        "adaptive": args.adaptive,
        "total_error_percent": summarize(
            float(row["total_error_percent"]) for row in accepted_rows
        ),
        "ratio_error_percent": summarize(
            float(row["ratio_error_percent"]) for row in accepted_rows
        ),
        "last_10_mean_abs_total_error_percent": mean(
            abs(float(row["total_error_percent"])) for row in rows[-10:]
        ),
        "final_gain_a": compensator.gain_a,
        "final_gain_b": compensator.gain_b,
        "final_confidence": compensator.confidence,
        "hardware_claims": "UNVERIFIED",
    }
    (output / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    _plot_results(output / "shot_convergence.png", rows)

    store = DualSlotCalibrationStore(output / "calibration_store")
    store.save(
        profile.key,
        compensator.gain_a,
        compensator.gain_b,
        compensator.confidence,
    )
    return summary


def _write_trace(path: Path, trace: list[object]) -> None:
    if not trace:
        return
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(trace[0].__dict__.keys()))
        writer.writeheader()
        for sample in trace:
            writer.writerow(sample.__dict__)


def _plot_results(path: Path, rows: list[dict[str, object]]) -> None:
    shots = [int(row["shot"]) for row in rows]
    total_error = [float(row["total_error_percent"]) for row in rows]
    ratio_error = [float(row["ratio_error_percent"]) for row in rows]
    gain_a = [float(row["gain_a"]) for row in rows]
    gain_b = [float(row["gain_b"]) for row in rows]

    figure, axes = plt.subplots(2, 1, figsize=(10, 7), sharex=True)
    axes[0].plot(shots, total_error, label="Total mass error (%)", color="#0b6e99")
    axes[0].plot(shots, ratio_error, label="A/B ratio error (%)", color="#d2691e")
    axes[0].axhline(0.0, color="black", linewidth=0.8)
    axes[0].set_ylabel("Error (%)")
    axes[0].grid(True, alpha=0.3)
    axes[0].legend()
    axes[1].plot(shots, gain_a, label="Gain A", color="#2e8b57")
    axes[1].plot(shots, gain_b, label="Gain B", color="#8a2be2")
    axes[1].set_xlabel("Shot")
    axes[1].set_ylabel("Bounded correction gain")
    axes[1].grid(True, alpha=0.3)
    axes[1].legend()
    figure.suptitle("Taeha dispenser simulated shot-to-shot convergence")
    figure.tight_layout()
    figure.savefig(path, dpi=160)
    plt.close(figure)


def build_parser() -> argparse.ArgumentParser:
    project_root = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--profiles", default=str(project_root / "configs" / "material_profiles.json")
    )
    parser.add_argument("--profile", default="epoxy-demo-25c-nozzle-a")
    parser.add_argument("--shots", type=int, default=40)
    parser.add_argument("--target-mass-mg", type=float, default=1200.0)
    parser.add_argument("--ratio", type=float, default=2.0)
    parser.add_argument("--temperature-c", type=float, default=25.0)
    parser.add_argument("--actuator-a", choices=[item.value for item in ActuatorKind], default="servo")
    parser.add_argument("--actuator-b", choices=[item.value for item in ActuatorKind], default="stepper")
    parser.add_argument("--seed", type=int, default=20260818)
    parser.add_argument("--adaptive", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--output", default="outputs/baseline")
    return parser


def main() -> None:
    summary = run(build_parser().parse_args())
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
