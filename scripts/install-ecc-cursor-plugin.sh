#!/usr/bin/env bash
# Install ECC (https://github.com/affaan-m/ECC) as a Cursor-native plugin.
#
# ECC ships a Claude Code marketplace (.claude-plugin/), not a Cursor
# marketplace (.cursor-plugin/). This adapter copies on-demand catalog
# (skills, agents, commands, scripts) into plugins/ecc/ and optionally
# registers it under ~/.cursor/plugins/local/ecc.
#
# It does NOT install ECC always-on Cursor rules or hooks into .cursor/,
# so this repo's career-hub rules stay in control.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PLUGIN_DIR="${ROOT}/plugins/ecc"
ECC_SOURCE="${ECC_SOURCE:-}"
ECC_GIT_URL="${ECC_GIT_URL:-https://github.com/affaan-m/ECC.git}"
INSTALL_LOCAL="${INSTALL_LOCAL:-1}"
SKIP_FETCH="${SKIP_FETCH:-0}"

usage() {
  cat <<'EOF'
Usage: scripts/install-ecc-cursor-plugin.sh [options]

Options:
  --source <path>   Existing ECC checkout to copy from (default: clone)
  --git-url <url>   ECC git URL (default: https://github.com/affaan-m/ECC.git)
  --no-local        Do not symlink into ~/.cursor/plugins/local/ecc
  --help            Show this help

Environment:
  ECC_SOURCE, ECC_GIT_URL, INSTALL_LOCAL=0, SKIP_FETCH=1
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --source)
      ECC_SOURCE="$2"
      SKIP_FETCH=1
      shift 2
      ;;
    --git-url)
      ECC_GIT_URL="$2"
      shift 2
      ;;
    --no-local)
      INSTALL_LOCAL=0
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

workdir="$(mktemp -d "${TMPDIR:-/tmp}/ecc-cursor-plugin.XXXXXX")"
cleanup() { rm -rf "${workdir}"; }
trap cleanup EXIT

if [[ -n "${ECC_SOURCE}" ]]; then
  if [[ ! -d "${ECC_SOURCE}/skills" || ! -f "${ECC_SOURCE}/.claude-plugin/marketplace.json" ]]; then
    echo "ECC source is missing skills/ or .claude-plugin/marketplace.json: ${ECC_SOURCE}" >&2
    exit 1
  fi
  src="${ECC_SOURCE}"
else
  echo "[ecc] Cloning ${ECC_GIT_URL}"
  git clone --depth 1 "${ECC_GIT_URL}" "${workdir}/ECC"
  src="${workdir}/ECC"
fi

sha="$(git -C "${src}" rev-parse HEAD 2>/dev/null || true)"
short_sha="$(git -C "${src}" rev-parse --short HEAD 2>/dev/null || echo unknown)"
plugin_version="$(python3 - <<PY
import json
from pathlib import Path
p = Path("${src}") / ".claude-plugin" / "plugin.json"
print(json.loads(p.read_text()).get("version", "0.0.0"))
PY
)"

echo "[ecc] Building Cursor plugin from ${src} (${short_sha}, v${plugin_version})"

python3 - "${src}" "${PLUGIN_DIR}" "${ECC_GIT_URL}" "${sha}" "${plugin_version}" <<'PY'
import json
import re
import shutil
import sys
from pathlib import Path

src = Path(sys.argv[1])
dst = Path(sys.argv[2])
git_url = sys.argv[3]
sha = sys.argv[4]
version = sys.argv[5]

if dst.exists():
    shutil.rmtree(dst)
dst.mkdir(parents=True)

def copytree(name: str) -> None:
    source = src / name
    if not source.is_dir():
        raise SystemExit(f"Missing ECC directory: {source}")
    shutil.copytree(
        source,
        dst / name,
        ignore=shutil.ignore_patterns("node_modules", ".git", "__pycache__", "*.pyc"),
        dirs_exist_ok=True,
    )

copytree("skills")
copytree("commands")
copytree("scripts")
shutil.copy2(src / "LICENSE", dst / "LICENSE")

agents_src = src / "agents"
agents_dst = dst / "agents"
agents_dst.mkdir(parents=True)
name_re = re.compile(r"^(name:\s*)([^\n]+)$", re.M)
copied_agents = 0
for path in sorted(agents_src.glob("*.md")):
    text = path.read_text(encoding="utf-8")
    stem = path.stem
    plugin_stem = stem if stem.startswith("ecc-") else f"ecc-{stem}"
    def repl(match: re.Match[str]) -> str:
        current = match.group(2).strip().strip('"').strip("'")
        if current.startswith("ecc-"):
            return f"{match.group(1)}{current}"
        return f"{match.group(1)}ecc-{current}"
    text = name_re.sub(repl, text, count=1)
    (agents_dst / f"{plugin_stem}.md").write_text(text, encoding="utf-8")
    copied_agents += 1

skill_count = len(list((dst / "skills").glob("*/SKILL.md")))
command_count = len(list((dst / "commands").glob("*.md")))

plugin_manifest = {
    "name": "ecc",
    "displayName": "ECC",
    "version": version,
    "description": "Cursor adapter for ECC (affaan-m/ECC): on-demand engineering skills, agents, and commands. Does not inject always-on rules or hooks.",
    "author": {
        "name": "Affaan Mustafa",
        "email": "me@affaanmustafa.com",
    },
    "homepage": "https://ecc.tools",
    "repository": git_url.rstrip(".git"),
    "license": "MIT",
    "keywords": [
        "ecc",
        "agents",
        "skills",
        "tdd",
        "code-review",
        "security",
        "workflow",
    ],
    "category": "developer-tools",
    "skills": "./skills/",
    "agents": "./agents/",
    "commands": "./commands/",
}

manifest_dir = dst / ".cursor-plugin"
manifest_dir.mkdir(parents=True)
(manifest_dir / "plugin.json").write_text(
    json.dumps(plugin_manifest, indent=2) + "\n", encoding="utf-8"
)

source_meta = {
    "marketplace": "https://github.com/affaan-m/ECC.git",
    "upstreamFormat": "claude-plugin",
    "adapter": "cursor-plugin",
    "version": version,
    "resolvedCommitSha": sha or None,
    "copied": {
        "skills": skill_count,
        "agents": copied_agents,
        "commands": command_count,
    },
    "excluded": [
        "always-on .cursor/rules",
        "hooks runtime",
        "mcp.json chrome-devtools default",
    ],
    "notes": [
        "ECC is a Claude Code / Codex marketplace, not a Cursor marketplace.",
        "This adapter exposes the catalog as a Cursor plugin without overriding hub rules.",
    ],
}
(dst / "SOURCE.json").write_text(json.dumps(source_meta, indent=2) + "\n", encoding="utf-8")

readme = f"""# ECC Cursor adapter

Cursor-native plugin built from [affaan-m/ECC]({git_url.rstrip('.git')}) `{sha or 'local'}` (v{version}).

## What is included

- {skill_count} skills
- {copied_agents} agents (prefixed `ecc-`)
- {command_count} commands
- ECC `scripts/` for skills that call `node scripts/ecc.js`

## What is not included

- ECC always-on Cursor rules (`alwaysApply: true`)
- ECC hook runtime
- ECC default MCP (`chrome-devtools`)

Those stay out so this repo's career-hub rules in `.cursor/rules/` remain in control.

## Refresh

```bash
bash scripts/install-ecc-cursor-plugin.sh
```
"""
(dst / "README.md").write_text(readme, encoding="utf-8")

print(f"skills={skill_count} agents={copied_agents} commands={command_count}")
PY

mkdir -p "${ROOT}/.cursor-plugin"
python3 - "${ROOT}/.cursor-plugin/marketplace.json" "${plugin_version}" <<'PY'
import json
from pathlib import Path
import sys

path = Path(sys.argv[1])
version = sys.argv[2]
manifest = {
    "name": "codex-ai-hub",
    "owner": {
        "name": "Joonha Lee",
    },
    "metadata": {
        "description": "Personal Cursor marketplace for Joonha Lee. Includes an ECC adapter because upstream ECC is a Claude Code marketplace, not a Cursor marketplace.",
        "version": "1.0.0",
        "pluginRoot": "plugins",
    },
    "plugins": [
        {
            "name": "ecc",
            "source": "ecc",
            "description": f"ECC {version} Cursor adapter: on-demand skills, agents, and commands from https://github.com/affaan-m/ECC",
            "version": version,
            "homepage": "https://ecc.tools",
            "repository": "https://github.com/affaan-m/ECC",
            "license": "MIT",
        }
    ],
}
path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print(f"wrote {path}")
PY

if [[ "${INSTALL_LOCAL}" == "1" ]]; then
  local_root="${HOME}/.cursor/plugins/local"
  mkdir -p "${local_root}"
  ln -sfn "${PLUGIN_DIR}" "${local_root}/ecc"
  echo "[ecc] Linked ${local_root}/ecc -> ${PLUGIN_DIR}"
fi

echo "[ecc] Cursor plugin ready at ${PLUGIN_DIR}"
