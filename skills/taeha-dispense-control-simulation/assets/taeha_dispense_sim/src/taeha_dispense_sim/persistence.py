from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import tempfile
import zlib


@dataclass(frozen=True)
class CalibrationRecord:
    magic: str
    schema_version: int
    profile_key: str
    sequence: int
    gain_a: float
    gain_b: float
    confidence: float
    crc32: int


def _payload(
    profile_key: str, sequence: int, gain_a: float, gain_b: float, confidence: float
) -> dict[str, object]:
    return {
        "magic": "THDS",
        "schema_version": 1,
        "profile_key": profile_key,
        "sequence": sequence,
        "gain_a": gain_a,
        "gain_b": gain_b,
        "confidence": confidence,
    }


def _crc(payload: dict[str, object]) -> int:
    canonical = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return zlib.crc32(canonical) & 0xFFFFFFFF


class DualSlotCalibrationStore:
    """Host model of a CRC-protected, monotonic dual-copy calibration store."""

    def __init__(self, directory: str | Path) -> None:
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)

    def load(self, profile_key: str) -> CalibrationRecord | None:
        valid = [record for record in (self._read("a"), self._read("b")) if record]
        valid = [record for record in valid if record.profile_key == profile_key]
        return max(valid, key=lambda item: item.sequence, default=None)

    def save(
        self,
        profile_key: str,
        gain_a: float,
        gain_b: float,
        confidence: float,
        simulate_interruption: bool = False,
    ) -> CalibrationRecord:
        current = self.load(profile_key)
        next_sequence = 1 if current is None else current.sequence + 1
        active_slot = None if current is None else self._slot_for_sequence(current.sequence)
        target_slot = "a" if active_slot != "a" else "b"
        payload = _payload(profile_key, next_sequence, gain_a, gain_b, confidence)
        record = CalibrationRecord(**payload, crc32=_crc(payload))
        target = self.directory / f"calibration_{target_slot}.json"

        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", delete=False, dir=self.directory, suffix=".tmp"
        ) as handle:
            json.dump(asdict(record), handle, sort_keys=True, indent=2)
            temp_path = Path(handle.name)
        if simulate_interruption:
            temp_path.unlink(missing_ok=True)
            if current is None:
                raise RuntimeError("simulated interrupted first write")
            return current
        temp_path.replace(target)
        verified = self._read(target_slot)
        if verified != record:
            raise IOError("calibration verification failed")
        return record

    def corrupt_latest(self, profile_key: str) -> None:
        current = self.load(profile_key)
        if current is None:
            raise FileNotFoundError("no calibration record")
        slot = self._slot_for_sequence(current.sequence)
        path = self.directory / f"calibration_{slot}.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        data["gain_a"] = float(data["gain_a"]) + 0.5
        path.write_text(json.dumps(data), encoding="utf-8")

    def _slot_for_sequence(self, sequence: int) -> str:
        for slot in ("a", "b"):
            record = self._read(slot)
            if record and record.sequence == sequence:
                return slot
        raise FileNotFoundError("sequence is not present in a valid slot")

    def _read(self, slot: str) -> CalibrationRecord | None:
        path = self.directory / f"calibration_{slot}.json"
        if not path.exists():
            return None
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            stored_crc = int(raw.pop("crc32"))
            if raw.get("magic") != "THDS" or int(raw.get("schema_version", -1)) != 1:
                return None
            if _crc(raw) != stored_crc:
                return None
            return CalibrationRecord(**raw, crc32=stored_crc)
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            return None
