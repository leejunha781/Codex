#!/usr/bin/env python3
"""Strip C2PA / Content Credentials metadata from LinkedIn infographic PNGs."""

from __future__ import annotations

import argparse
import struct
import sys
from pathlib import Path


def png_has_c2pa(path: Path) -> bool:
    raw = path.read_bytes()
    return any(marker in raw for marker in (b"caBX", b"c2pa", b"jumb", b"C2PA"))


def strip_with_pillow(path: Path) -> None:
    from PIL import Image

    with Image.open(path) as image:
        image.convert("RGB").save(path, format="PNG", optimize=True)


def strip_png_chunks(path: Path) -> None:
    """Fallback: rewrite PNG without caBX/C2PA auxiliary chunks."""
    data = path.read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"Not a PNG file: {path}")

    out = bytearray(data[:8])
    offset = 8
    while offset + 8 <= len(data):
        length = struct.unpack(">I", data[offset : offset + 4])[0]
        chunk_type = data[offset + 4 : offset + 8]
        chunk_end = offset + 12 + length
        if chunk_end > len(data):
            break
        if chunk_type not in (b"caBX", b"c2pa"):
            out.extend(data[offset:chunk_end])
        offset = chunk_end

    if offset < len(data):
        out.extend(data[offset:])
    path.write_bytes(bytes(out))


def strip_file(path: Path) -> bool:
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".png":
        print(f"SKIP (not PNG): {path}")
        return False
    if not png_has_c2pa(path):
        print(f"OK   no C2PA metadata: {path}")
        return False

    try:
        strip_with_pillow(path)
    except ImportError:
        strip_png_chunks(path)

    if png_has_c2pa(path):
        raise RuntimeError(f"C2PA metadata still present after strip: {path}")

    print(f"STRIP C2PA removed: {path}")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", help="PNG file or folder paths")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    targets: list[Path] = []
    for raw in args.paths:
        path = Path(raw)
        if path.is_dir():
            targets.extend(sorted(path.glob("*-infographic.png")))
        else:
            targets.append(path)

    if not targets:
        print("No infographic PNG files found.")
        return 0

    stripped = 0
    for target in targets:
        if args.dry_run:
            if png_has_c2pa(target):
                print(f"[dry-run] would strip C2PA from {target}")
                stripped += 1
            else:
                print(f"OK   no C2PA metadata: {target}")
            continue
        if strip_file(target):
            stripped += 1

    print(f"C2PA strip complete: {stripped} file(s) cleaned.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
