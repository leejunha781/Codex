#!/usr/bin/env bash
set -euo pipefail

export NODE_USE_SYSTEM_CA="${NODE_USE_SYSTEM_CA:-1}"

if [[ -z "${NODE_OPTIONS:-}" ]]; then
  major="$(node -p "process.versions.node.split('.')[0]")"
  minor="$(node -p "process.versions.node.split('.')[1]")"

  if [[ "$major" -ge 24 ]] || [[ "$major" -eq 23 && "$minor" -ge 8 ]] || [[ "$major" -eq 22 && "$minor" -ge 15 ]]; then
    export NODE_OPTIONS="--use-system-ca"
  else
    export NODE_OPTIONS="--use-openssl-ca"
  fi
fi

echo "Node $(node --version) | NODE_USE_SYSTEM_CA=${NODE_USE_SYSTEM_CA} | NODE_OPTIONS=${NODE_OPTIONS}"

node <<'NODE'
const targets = [
  "https://registry.npmjs.org",
  "https://api.github.com",
];

(async () => {
  for (const url of targets) {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`${url} returned HTTP ${response.status}`);
    }
    console.log(`TLS OK: ${url} (${response.status})`);
  }
})();
NODE
