#!/usr/bin/env bash
# Verify the ECC Cursor plugin adapter is complete and does not override hub rules.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PLUGIN="${ROOT}/plugins/ecc"
errors=0

fail() {
  echo "FAIL: $*" >&2
  errors=$((errors + 1))
}

pass() {
  echo "OK: $*"
}

require_file() {
  [[ -f "$1" ]] || fail "missing file $1"
}

require_file "${ROOT}/.cursor-plugin/marketplace.json"
require_file "${PLUGIN}/.cursor-plugin/plugin.json"
require_file "${PLUGIN}/SOURCE.json"
require_file "${PLUGIN}/LICENSE"
require_file "${ROOT}/.cursor/rules/00-core-agent.mdc"
require_file "${ROOT}/.cursor/skills/resume-jd-match/SKILL.md"

python3 - "${ROOT}" "${PLUGIN}" <<'PY' || exit 1
import json
import sys
from pathlib import Path

root = Path(sys.argv[1])
plugin = Path(sys.argv[2])
errors = 0

def fail(msg: str) -> None:
    global errors
    print(f"FAIL: {msg}")
    errors += 1

def ok(msg: str) -> None:
    print(f"OK: {msg}")

market = json.loads((root / ".cursor-plugin" / "marketplace.json").read_text())
plugin_json = json.loads((plugin / ".cursor-plugin" / "plugin.json").read_text())
source = json.loads((plugin / "SOURCE.json").read_text())

if market.get("name") != "codex-ai-hub":
    fail("marketplace name should be codex-ai-hub")
else:
    ok("marketplace name is codex-ai-hub")

plugins = market.get("plugins") or []
if not any(p.get("name") == "ecc" and p.get("source") == "ecc" for p in plugins):
    fail("marketplace does not list ecc from plugins/ecc")
else:
    ok("marketplace lists ecc")

if plugin_json.get("name") != "ecc":
    fail("plugin.json name should be ecc")
else:
    ok("plugin.json name is ecc")

if "hooks" in plugin_json or "rules" in plugin_json:
    fail("adapter should not enable ECC hooks or rules")
else:
    ok("adapter omits hooks and rules")

skills = list((plugin / "skills").glob("*/SKILL.md"))
agents = list((plugin / "agents").glob("ecc-*.md"))
commands = list((plugin / "commands").glob("*.md"))
if len(skills) < 200:
    fail(f"expected 200+ skills, found {len(skills)}")
else:
    ok(f"{len(skills)} skills")
if len(agents) < 50:
    fail(f"expected 50+ ecc- agents, found {len(agents)}")
else:
    ok(f"{len(agents)} prefixed agents")
if len(commands) < 50:
    fail(f"expected 50+ commands, found {len(commands)}")
else:
    ok(f"{len(commands)} commands")

hub_rules = list((root / ".cursor" / "rules").glob("*.mdc"))
if not any(p.name == "00-core-agent.mdc" for p in hub_rules):
    fail("hub core rule missing")
else:
    ok("hub core rule preserved")
if any(p.name.startswith("common-") for p in hub_rules):
    fail("ECC always-on common-* rules leaked into .cursor/rules")
else:
    ok("no ECC common-* rules in hub .cursor/rules")

mcp = json.loads((root / ".cursor" / "mcp.json").read_text())
servers = (mcp.get("mcpServers") or {})
if "chrome-devtools" in servers:
    fail("ECC chrome-devtools MCP should not be merged into hub mcp.json")
else:
    ok("hub mcp.json has no ECC chrome-devtools")
if "notion" not in servers or "linear" not in servers:
    fail("hub Notion/Linear MCP missing")
else:
    ok("hub Notion/Linear MCP preserved")

if source.get("marketplace") != "https://github.com/affaan-m/ECC.git":
    fail("SOURCE.json marketplace URL mismatch")
else:
    ok("SOURCE.json pins affaan-m/ECC")

sys.exit(1 if errors else 0)
PY

if [[ -L "${HOME}/.cursor/plugins/local/ecc" ]]; then
  target="$(readlink -f "${HOME}/.cursor/plugins/local/ecc")"
  if [[ "${target}" == "$(readlink -f "${PLUGIN}")" ]]; then
    pass "local plugin symlink points at plugins/ecc"
  else
    fail "local plugin symlink target is ${target}"
  fi
else
  echo "WARN: ~/.cursor/plugins/local/ecc is not linked (run install script without --no-local)"
fi

if [[ "${errors}" -ne 0 ]]; then
  echo "ECC plugin verification failed (${errors} errors)" >&2
  exit 1
fi

echo "ECC plugin verification passed"
