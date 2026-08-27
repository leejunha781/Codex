#!/usr/bin/env python3
"""Configuration integrity check for the Codex AI Hub repository.

Validates the Cursor configuration that this repository ships:

* JSON config files parse (``.cursor/mcp.json``, ``.cursor/plugin.json``).
* ``plugin.json`` points ``rules``/``skills`` at real directories.
* ``mcp.json`` declares at least one MCP server, each with a URL or command.
* Every rule (``.cursor/rules/*.mdc``) has a valid frontmatter block with a
  non-empty ``description``.
* Every skill (``.cursor/skills/*/SKILL.md``) has frontmatter with ``name`` and
  ``description``, and its ``name`` matches its directory.
* Context-hub markdown files referenced by the project exist and are non-empty.

Only the Python standard library is used, so it runs on the default Cloud Agent
image without installing anything.

Exit code 0 means the workspace configuration is coherent; non-zero means at
least one problem was found (all problems are reported, not just the first).
"""
from __future__ import annotations

import glob
import json
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

errors: list[str] = []
checks: list[str] = []


def ok(msg: str) -> None:
    checks.append(msg)
    print(f"  \033[32mPASS\033[0m {msg}")


def fail(msg: str) -> None:
    errors.append(msg)
    print(f"  \033[31mFAIL\033[0m {msg}")


def read(path: str) -> str:
    with open(os.path.join(ROOT, path), encoding="utf-8") as fh:
        return fh.read()


def parse_frontmatter(text: str) -> dict[str, str] | None:
    """Parse a leading ``---`` delimited ``key: value`` frontmatter block.

    Returns ``None`` when no valid frontmatter block is present.
    """
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return None
    data: dict[str, str] = {}
    for line in lines[1:]:
        if line.strip() == "---":
            return data
        if not line.strip() or ":" not in line:
            continue
        key, _, value = line.partition(":")
        data[key.strip()] = value.strip()
    return None  # no closing delimiter


def load_json(path: str):
    raw = read(path)
    return json.loads(raw)


def check_json_files() -> None:
    print("[1/5] JSON config files parse")
    for rel in (".cursor/mcp.json", ".cursor/plugin.json"):
        try:
            load_json(rel)
            ok(f"{rel} is valid JSON")
        except FileNotFoundError:
            fail(f"{rel} is missing")
        except json.JSONDecodeError as exc:
            fail(f"{rel} is not valid JSON: {exc}")


def check_plugin() -> None:
    print("[2/5] plugin.json references real directories")
    try:
        plugin = load_json(".cursor/plugin.json")
    except Exception as exc:  # already reported in check_json_files
        fail(f".cursor/plugin.json unreadable: {exc}")
        return
    for key in ("rules", "skills"):
        rel = plugin.get(key)
        if not rel:
            fail(f"plugin.json missing '{key}' key")
            continue
        target = os.path.normpath(os.path.join(ROOT, ".cursor", rel))
        if os.path.isdir(target):
            ok(f"plugin.json '{key}' -> {rel} exists")
        else:
            fail(f"plugin.json '{key}' -> {rel} is not a directory")


def check_mcp() -> None:
    print("[3/5] mcp.json declares usable servers")
    try:
        mcp = load_json(".cursor/mcp.json")
    except Exception as exc:
        fail(f".cursor/mcp.json unreadable: {exc}")
        return
    servers = mcp.get("mcpServers")
    if not isinstance(servers, dict) or not servers:
        fail("mcp.json has no 'mcpServers' entries")
        return
    for name, cfg in servers.items():
        if isinstance(cfg, dict) and (cfg.get("url") or cfg.get("command")):
            ok(f"mcp server '{name}' has a url/command")
        else:
            fail(f"mcp server '{name}' lacks a url or command")


def check_rules() -> None:
    print("[4/5] rules have valid frontmatter")
    rule_files = sorted(glob.glob(os.path.join(ROOT, ".cursor/rules/*.mdc")))
    if not rule_files:
        fail("no rule files found under .cursor/rules/")
        return
    for path in rule_files:
        rel = os.path.relpath(path, ROOT)
        fm = parse_frontmatter(read(rel))
        if fm is None:
            fail(f"{rel} has no valid frontmatter block")
        elif not fm.get("description"):
            fail(f"{rel} frontmatter missing 'description'")
        else:
            ok(f"{rel} frontmatter valid")


def check_skills() -> None:
    print("[5/5] skills have valid frontmatter")
    skill_files = sorted(glob.glob(os.path.join(ROOT, ".cursor/skills/*/SKILL.md")))
    if not skill_files:
        fail("no skill files found under .cursor/skills/")
        return
    for path in skill_files:
        rel = os.path.relpath(path, ROOT)
        dirname = os.path.basename(os.path.dirname(path))
        fm = parse_frontmatter(read(rel))
        if fm is None:
            fail(f"{rel} has no valid frontmatter block")
            continue
        if not fm.get("description"):
            fail(f"{rel} frontmatter missing 'description'")
        name = fm.get("name")
        if not name:
            fail(f"{rel} frontmatter missing 'name'")
        elif name != dirname:
            fail(f"{rel} name '{name}' does not match directory '{dirname}'")
        elif fm.get("description"):
            ok(f"{rel} frontmatter valid (name={name})")


def check_context_hub() -> None:
    print("[bonus] context-hub files present and non-empty")
    for rel in ("AGENTS.md", "CLAUDE.md", "MEMORY.md", "README.md",
                "chatgpt-preferences.md", "docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md"):
        full = os.path.join(ROOT, rel)
        if os.path.isfile(full) and os.path.getsize(full) > 0:
            ok(f"{rel} present")
        else:
            fail(f"{rel} missing or empty")


def main() -> int:
    print(f"Codex AI Hub — configuration integrity check\nroot: {ROOT}\n")
    check_json_files()
    check_plugin()
    check_mcp()
    check_rules()
    check_skills()
    check_context_hub()

    print(
        f"\nSummary: {len(checks)} checks passed, {len(errors)} failed."
    )
    if errors:
        print("\nProblems:")
        for e in errors:
            print(f"  - {e}")
        return 1
    print("All configuration artifacts are valid.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
