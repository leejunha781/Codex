#!/usr/bin/env bash
# Configuration integrity check for the Codex AI Hub repository.
#
# This repository has no application runtime or package manager — its "product"
# is the Cursor configuration itself (rules, skills, MCP/plugin JSON, and the
# context hub markdown). This script is the repository's end-to-end check: it
# verifies that every config artifact is present and well-formed so future
# agents boot against a coherent workspace.
#
# It depends only on tools present in the default Cloud Agent image (bash,
# python3 stdlib). No third-party packages are required.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

exec python3 scripts/validate_config.py
