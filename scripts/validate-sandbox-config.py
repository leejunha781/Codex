#!/usr/bin/env python3
"""Validate committed Cursor sandbox and Auto-review policy files."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SANDBOX = ROOT / ".cursor" / "sandbox.json"
PERMISSIONS = ROOT / ".cursor" / "permissions.json"

ALLOWED_SANDBOX_TYPES = {"workspace_readwrite", "workspace_readonly", "insecure_none"}
SANDBOX_KEYS = {
    "type",
    "additionalReadwritePaths",
    "additionalReadonlyPaths",
    "disableTmpWrite",
    "enableSharedBuildCache",
    "networkPolicy",
}
NETWORK_KEYS = {"default", "allow", "deny"}


def fail(message: str) -> None:
    print(f"[FAIL] {message}")
    raise SystemExit(1)


def load_json(path: Path) -> Any:
    if not path.is_file():
        fail(f"missing file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        fail(f"{path} is not valid JSON: {exc}")
    return None


def validate_sandbox(data: Any) -> None:
    if not isinstance(data, dict):
        fail("sandbox.json must be an object")
    unknown = set(data) - SANDBOX_KEYS
    if unknown:
        fail(f"sandbox.json has unknown keys: {sorted(unknown)}")
    sandbox_type = data.get("type", "workspace_readwrite")
    if sandbox_type not in ALLOWED_SANDBOX_TYPES:
        fail(f"invalid sandbox type: {sandbox_type}")
    if sandbox_type == "insecure_none":
        fail("type insecure_none disables the sandbox; do not commit that")
    for key in ("additionalReadwritePaths", "additionalReadonlyPaths"):
        if key in data and not isinstance(data[key], list):
            fail(f"{key} must be an array of strings")
        if key in data and not all(isinstance(item, str) for item in data[key]):
            fail(f"{key} must contain only strings")
    for key in ("disableTmpWrite", "enableSharedBuildCache"):
        if key in data and not isinstance(data[key], bool):
            fail(f"{key} must be a boolean")
    policy = data.get("networkPolicy")
    if policy is None:
        return
    if not isinstance(policy, dict):
        fail("networkPolicy must be an object")
    unknown_net = set(policy) - NETWORK_KEYS
    if unknown_net:
        fail(f"networkPolicy has unknown keys: {sorted(unknown_net)}")
    default = policy.get("default", "deny")
    if default not in {"allow", "deny"}:
        fail(f"networkPolicy.default must be allow or deny: {default}")
    for key in ("allow", "deny"):
        if key in policy and not isinstance(policy[key], list):
            fail(f"networkPolicy.{key} must be an array")
        if key in policy and not all(isinstance(item, str) for item in policy[key]):
            fail(f"networkPolicy.{key} must contain only strings")


def validate_permissions(data: Any) -> None:
    if not isinstance(data, dict):
        fail("permissions.json must be an object")
    auto_run = data.get("autoRun")
    if not isinstance(auto_run, dict):
        fail("permissions.json must contain autoRun object")
    for key in ("allow_instructions", "block_instructions"):
        value = auto_run.get(key, [])
        if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
            fail(f"autoRun.{key} must be an array of strings")
        if not value and key == "block_instructions":
            fail("autoRun.block_instructions must not be empty in this repo")


def main() -> None:
    sandbox = load_json(SANDBOX)
    permissions = load_json(PERMISSIONS)
    validate_sandbox(sandbox)
    validate_permissions(permissions)
    print("[PASS] .cursor/sandbox.json")
    print("[PASS] .cursor/permissions.json")
    print("[PASS] Cursor sandbox policy files are valid.")


if __name__ == "__main__":
    main()
