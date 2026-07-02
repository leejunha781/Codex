#!/usr/bin/env bash
set -euo pipefail

export NODE_USE_SYSTEM_CA="${NODE_USE_SYSTEM_CA:-1}"
export NODE_OPTIONS="${NODE_OPTIONS:---use-openssl-ca}"

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
