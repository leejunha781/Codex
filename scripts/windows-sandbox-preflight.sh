#!/usr/bin/env bash
# Preflight for Cursor's Linux sandbox (used on Linux desktops and inside WSL2 on Windows).
# Cursor Windows sandbox = this Linux Landlock/seccomp path, not a native Win32 sandbox.
set -u

fail=0
warn=0

pass() { printf '[PASS] %s\n' "$1"; }
fail_msg() { printf '[FAIL] %s\n' "$1"; fail=$((fail + 1)); }
warn_msg() { printf '[WARN] %s\n' "$1"; warn=$((warn + 1)); }
info() { printf '[INFO] %s\n' "$1"; }

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd "$script_dir/.." && pwd)"

printf 'Cursor Linux/WSL sandbox preflight\n'
printf '==================================\n'
info "Host: $(uname -s) $(uname -r) $(uname -m)"

if grep -qi microsoft /proc/version 2>/dev/null; then
  info "Detected WSL. This is the backend Cursor uses for Windows sandbox."
else
  info "Not WSL. This check still validates the Linux Landlock sandbox used on Linux desktops."
fi

kernel_raw="$(uname -r)"
kernel_base="${kernel_raw%%-*}"
kernel_major="${kernel_base%%.*}"
kernel_rest="${kernel_base#*.}"
kernel_minor="${kernel_rest%%.*}"
kernel_major="${kernel_major:-0}"
kernel_minor="${kernel_minor:-0}"

if [ "$kernel_major" -gt 6 ] || { [ "$kernel_major" -eq 6 ] && [ "$kernel_minor" -ge 2 ]; }; then
  pass "Kernel $kernel_raw meets 6.2+ Landlock v3 baseline."
else
  fail_msg "Kernel $kernel_raw is older than 6.2. On Windows run: wsl --update"
fi

userns_file="/proc/sys/kernel/unprivileged_userns_clone"
if [ -r "$userns_file" ]; then
  userns_val="$(tr -d '[:space:]' < "$userns_file")"
  if [ "$userns_val" = "1" ]; then
    pass "Unprivileged user namespaces enabled ($userns_file=1)."
  else
    fail_msg "Unprivileged user namespaces disabled ($userns_file=$userns_val)."
  fi
else
  warn_msg "$userns_file not readable; cannot confirm user namespaces."
fi

max_userns="/proc/sys/user/max_user_namespaces"
if [ -r "$max_userns" ]; then
  max_val="$(tr -d '[:space:]' < "$max_userns")"
  if [ "$max_val" = "0" ]; then
    fail_msg "user.max_user_namespaces=0; sandbox user namespaces cannot be created."
  else
    pass "user.max_user_namespaces=$max_val"
  fi
fi

if [ -r /sys/kernel/security/lsm ]; then
  lsm="$(tr -d '\n' < /sys/kernel/security/lsm)"
  info "LSM list: $lsm"
  if printf '%s' "$lsm" | grep -q landlock; then
    pass "Landlock appears in the LSM list."
  else
    warn_msg "Landlock not listed in LSM. Cursor may still work if CONFIG_SECURITY_LANDLOCK=y; confirm with dmesg or Cursor logs."
  fi
else
  warn_msg "/sys/kernel/security/lsm not readable."
fi

config_file="/boot/config-$(uname -r)"
if [ -r "$config_file" ]; then
  if grep -q 'CONFIG_SECURITY_LANDLOCK=y' "$config_file"; then
    pass "CONFIG_SECURITY_LANDLOCK=y in $config_file"
  else
    warn_msg "CONFIG_SECURITY_LANDLOCK=y not found in $config_file"
  fi
else
  info "Kernel config file not present at $config_file (common in WSL)."
fi

if [ -n "${CURSOR_SANDBOX:-}" ]; then
  info "Currently inside a Cursor sandbox (CURSOR_SANDBOX=$CURSOR_SANDBOX)."
  info "CURSOR_ORIG_UID=${CURSOR_ORIG_UID:-unset} CURSOR_ORIG_GID=${CURSOR_ORIG_GID:-unset}"
  info "CURSOR_SANDBOX_LANDLOCK_STATUS=${CURSOR_SANDBOX_LANDLOCK_STATUS:-unset}"
else
  info "Not currently inside a Cursor-sandboxed child process (CURSOR_SANDBOX unset). Expected for a manual preflight."
fi

if command -v unshare >/dev/null 2>&1; then
  if unshare --user --map-root-user true >/dev/null 2>&1; then
    pass "unshare --user succeeded (user namespace usable)."
  else
    fail_msg "unshare --user failed. AppArmor or distro policy may block user namespaces."
    info "Remote/CLI only: install https://downloads.cursor.com/lab/enterprise/cursor-sandbox-apparmor_0.6.0_all.deb"
    info "Local Cursor desktop usually ships the AppArmor profile and does not need this package."
  fi
else
  warn_msg "unshare not installed; skip user-namespace smoke test."
fi

helper=""
if [ -n "${HOME:-}" ]; then
  helper="$(find "$HOME/.cursor-server/bin" -path '*/resources/helpers/cursorsandbox' 2>/dev/null | sort | tail -n 1 || true)"
fi
if [ -z "$helper" ]; then
  helper="$(find /usr /opt "$HOME/.local" -name cursorsandbox -type f 2>/dev/null | head -n 1 || true)"
fi
if [ -n "$helper" ]; then
  pass "Found sandbox helper: $helper"
  if [ -x "$helper" ]; then
    if "$helper" --help >/dev/null 2>&1 || "$helper" -h >/dev/null 2>&1; then
      pass "Sandbox helper is executable."
    else
      info "Sandbox helper exists; it may only support Cursor's internal preflight flags."
    fi
  else
    warn_msg "Sandbox helper is not executable: $helper"
  fi
else
  info "cursorsandbox helper not found yet. Cursor copies it after the desktop/WSL remote server starts."
fi

sandbox_json="$repo_root/.cursor/sandbox.json"
permissions_json="$repo_root/.cursor/permissions.json"
if [ -f "$sandbox_json" ]; then
  if python3 -c "import json,sys; json.load(open(sys.argv[1]))" "$sandbox_json"; then
    pass "Valid JSON: .cursor/sandbox.json"
  else
    fail_msg "Invalid JSON: .cursor/sandbox.json"
  fi
else
  fail_msg "Missing $sandbox_json"
fi
if [ -f "$permissions_json" ]; then
  if python3 -c "import json,sys; json.load(open(sys.argv[1]))" "$permissions_json"; then
    pass "Valid JSON: .cursor/permissions.json"
  else
    fail_msg "Invalid JSON: .cursor/permissions.json"
  fi
else
  warn_msg "Missing $permissions_json"
fi

printf '\nWindows desktop settings (confirm in Cursor UI)\n'
printf '----------------------------------------------\n'
printf '1. Settings > Agents > Approvals & Execution = Auto-review\n'
printf '2. Settings > Agents > Inline Editing & Terminal > Legacy Terminal Tool = Off\n'
printf '3. Network mode = sandbox.json + Defaults\n'
printf '4. Do not use Run Everything if you need a sandbox\n'
printf '\n'

if [ "$fail" -gt 0 ]; then
  fail_msg "Preflight finished with $fail failure(s) and $warn warning(s)."
  exit 1
fi
if [ "$warn" -gt 0 ]; then
  warn_msg "Preflight finished with $warn warning(s) and no hard failures."
  exit 0
fi
pass "Linux/WSL sandbox prerequisites look ready."
exit 0
