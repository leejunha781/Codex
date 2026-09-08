# ECC Cursor adapter

Cursor-native plugin built from [affaan-m/ECC](https://github.com/affaan-m/ECC) `5064474d4d762dc9640234a41617cccb79185cec` (v2.2.1).

## What is included

- 286 skills
- 68 agents (prefixed `ecc-`)
- 94 commands
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
