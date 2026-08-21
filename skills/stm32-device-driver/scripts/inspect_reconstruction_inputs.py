#!/usr/bin/env python3
"""Inventory STM32 reconstruction artifacts and validate Intel HEX structure."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import struct
import zipfile
from pathlib import Path
from typing import Any


RELEVANT_EXTENSIONS = {
    ".pdf", ".hex", ".bin", ".elf", ".out", ".map", ".ewp", ".eww",
    ".icf", ".ioc", ".svd", ".csv", ".txt", ".log", ".sal", ".vcd", ".zip",
}

SOURCE_EXTENSIONS = {".c", ".h", ".cpp", ".hpp", ".s", ".asm", ".ewp", ".eww", ".icf", ".ioc"}
EXECUTABLE_EXTENSIONS = {".exe", ".dll", ".ocx", ".msi", ".bat", ".cmd", ".ps1"}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def inspect_pdf(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {"probe_available": False}
    try:
        from pypdf import PdfReader
    except ImportError:
        return result
    try:
        reader = PdfReader(path)
        metadata = reader.metadata or {}
        text = "\n".join((page.extract_text() or "") for page in reader.pages[:5])[:100000]
        result.update({
            "probe_available": True,
            "page_count": len(reader.pages),
            "title": str(metadata.get("/Title") or ""),
            "author": str(metadata.get("/Author") or ""),
            "text_probe": text,
        })
    except Exception as exc:  # malformed/encrypted PDFs remain inventory-visible
        result["probe_error"] = str(exc)
    return result


def role_for(path: Path, pdf: dict[str, Any] | None = None) -> str:
    name = path.name.lower()
    if path.suffix.lower() == ".hex":
        return "firmware-hex"
    if path.suffix.lower() in {".bin", ".elf", ".out", ".map"}:
        return "firmware-build"
    if path.suffix.lower() in {".ewp", ".eww", ".icf", ".ioc", ".svd"}:
        return "project-or-device"
    if path.suffix.lower() == ".zip":
        return "vendor-support-package"
    if path.suffix.lower() == ".pdf":
        if any(word in name for word in ("schematic", "circuit", "회로", "sch")):
            return "schematic-candidate"
        if re.search(r"(?:^|[^a-z])drv\d+", name) or any(word in name for word in ("servo", "drive", "motor", "서보", "모터", "드라이브")):
            return "motor-drive-doc-candidate"
        if re.search(r"proc[-_].*[-_]io(?:[-_.]|$)", name):
            return "schematic-candidate"
        if pdf and pdf.get("probe_available"):
            title = str(pdf.get("title", "")).lower()
            probe = (title + "\n" + str(pdf.get("text_probe", ""))[:30000]).lower()
            if "datasheet" in probe and ("motor driver" in probe or re.search(r"drv\d+", probe)):
                return "motor-drive-doc-candidate"
            if "document number" in probe and "sheet" in probe:
                return "schematic-candidate"
        return "pdf-unclassified"
    return "capture-or-note"


def inspect_zip(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as archive:
        entries = [entry for entry in archive.infolist() if not entry.is_dir()]
        source = [entry.filename for entry in entries if Path(entry.filename).suffix.lower() in SOURCE_EXTENSIONS]
        executable = [entry.filename for entry in entries if Path(entry.filename).suffix.lower() in EXECUTABLE_EXTENSIONS]
        return {
            "valid": True,
            "entry_count": len(entries),
            "source_candidate_count": len(source),
            "source_candidates": source[:50],
            "executable_or_script_count": len(executable),
            "executable_or_scripts": executable[:50],
            "classification": "source-bearing-package" if source else "binary-or-documentation-package",
            "executed": False,
        }


def segments(addresses: list[int]) -> list[dict[str, int]]:
    if not addresses:
        return []
    result: list[dict[str, int]] = []
    start = previous = addresses[0]
    for address in addresses[1:]:
        if address != previous + 1:
            result.append({"start": start, "end": previous, "bytes": previous - start + 1})
            start = address
        previous = address
    result.append({"start": start, "end": previous, "bytes": previous - start + 1})
    return result


def parse_hex(path: Path) -> dict[str, Any]:
    memory: dict[int, int] = {}
    base = 0
    eof_seen = False
    start_address: int | None = None
    records = 0

    for line_number, raw in enumerate(path.read_text(encoding="ascii").splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        if not line.startswith(":"):
            raise ValueError(f"line {line_number}: missing ':'")
        try:
            record = bytes.fromhex(line[1:])
        except ValueError as exc:
            raise ValueError(f"line {line_number}: invalid hexadecimal data") from exc
        if len(record) < 5 or len(record) != record[0] + 5:
            raise ValueError(f"line {line_number}: byte count mismatch")
        if sum(record) & 0xFF:
            raise ValueError(f"line {line_number}: checksum mismatch")

        count = record[0]
        offset = (record[1] << 8) | record[2]
        record_type = record[3]
        data = record[4:4 + count]
        records += 1

        if record_type == 0x00:
            for index, value in enumerate(data):
                address = base + offset + index
                if address in memory and memory[address] != value:
                    raise ValueError(f"line {line_number}: conflicting data at 0x{address:08X}")
                memory[address] = value
        elif record_type == 0x01:
            if count != 0:
                raise ValueError(f"line {line_number}: malformed EOF record")
            eof_seen = True
        elif record_type == 0x02:
            if count != 2:
                raise ValueError(f"line {line_number}: malformed segment address")
            base = int.from_bytes(data, "big") << 4
        elif record_type == 0x04:
            if count != 2:
                raise ValueError(f"line {line_number}: malformed linear address")
            base = int.from_bytes(data, "big") << 16
        elif record_type == 0x05:
            if count != 4:
                raise ValueError(f"line {line_number}: malformed start address")
            start_address = int.from_bytes(data, "big")
        elif record_type == 0x03:
            if count != 4:
                raise ValueError(f"line {line_number}: malformed CS:IP address")
            start_address = ((data[0] << 8 | data[1]) << 4) + (data[2] << 8 | data[3])

    addresses = sorted(memory)
    vector: dict[str, Any] | None = None
    candidates = [0x08000000]
    if addresses:
        candidates.append(addresses[0])
    for candidate in dict.fromkeys(candidates):
        if candidate and all(candidate + index in memory for index in range(8)):
            raw_vector = bytes(memory[candidate + index] for index in range(8))
            initial_sp, reset_vector = struct.unpack("<II", raw_vector)
            vector = {
                "base": candidate,
                "initial_sp": initial_sp,
                "reset_vector": reset_vector,
                "sp_sram_like": 0x20000000 <= initial_sp < 0x40000000,
                "thumb_bit_set": bool(reset_vector & 1),
            }
            break

    return {
        "valid": True,
        "records": records,
        "eof_seen": eof_seen,
        "data_bytes": len(memory),
        "minimum_address": addresses[0] if addresses else None,
        "maximum_address": addresses[-1] if addresses else None,
        "segments": segments(addresses),
        "declared_start_address": start_address,
        "vector_candidate": vector,
    }


def inspect(root: Path) -> dict[str, Any]:
    files: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in RELEVANT_EXTENSIONS:
            continue
        pdf = inspect_pdf(path) if path.suffix.lower() == ".pdf" else None
        item: dict[str, Any] = {
            "path": str(path.relative_to(root)),
            "role": role_for(path, pdf),
            "size": path.stat().st_size,
            "sha256": sha256(path),
        }
        if pdf is not None:
            pdf.pop("text_probe", None)
            item["pdf"] = pdf
        if path.suffix.lower() == ".hex":
            try:
                item["hex"] = parse_hex(path)
            except (OSError, UnicodeError, ValueError) as exc:
                item["hex"] = {"valid": False, "error": str(exc)}
                errors.append({"path": item["path"], "error": str(exc)})
        elif path.suffix.lower() == ".zip":
            try:
                item["zip"] = inspect_zip(path)
            except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
                item["zip"] = {"valid": False, "error": str(exc), "executed": False}
                errors.append({"path": item["path"], "error": str(exc)})
        files.append(item)

    roles = {item["role"] for item in files}
    return {
        "root": str(root),
        "files": files,
        "readiness": {
            "has_firmware_hex": "firmware-hex" in roles,
            "has_schematic_candidate": "schematic-candidate" in roles,
            "has_motor_drive_doc_candidate": "motor-drive-doc-candidate" in roles,
            "has_vendor_support_package": "vendor-support-package" in roles,
            "note": "PDF roles use filename plus optional text/metadata probes; verify circuit details visually. ZIP files are inventoried without extraction or execution.",
        },
        "errors": errors,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--format", choices=("json", "summary"), default="summary")
    args = parser.parse_args()
    root = args.root.resolve()
    if not root.is_dir():
        parser.error(f"not a directory: {root}")
    report = inspect(root)
    if args.format == "json":
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"Root: {report['root']}")
        for item in report["files"]:
            status = ""
            if "hex" in item:
                status = " valid" if item["hex"]["valid"] else " INVALID"
            print(f"{item['role']:26} {item['size']:10} {item['sha256']} {item['path']}{status}")
        print(json.dumps(report["readiness"], ensure_ascii=False))
        if report["errors"]:
            print(json.dumps(report["errors"], ensure_ascii=False, indent=2))
    return 2 if report["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
