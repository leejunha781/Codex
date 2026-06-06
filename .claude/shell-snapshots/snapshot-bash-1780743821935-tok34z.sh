# Snapshot file
# Unset all aliases to avoid conflicts with functions
unalias -a 2>/dev/null || true
shopt -s expand_aliases
# Check for rg availability
if ! (unalias rg 2>/dev/null; command -v rg) >/dev/null 2>&1; then
  function rg {
  local _cc_bin="${CLAUDE_CODE_EXECPATH:-}"
  [[ -x $_cc_bin ]] || _cc_bin=/c/Users/namma/.local/bin/claude.exe
  if [[ ! -x $_cc_bin ]]; then command rg "$@"; return; fi
  if [[ -n $ZSH_VERSION ]]; then
    ARGV0=rg "$_cc_bin" "$@"
  elif [[ "$OSTYPE" == "msys" ]] || [[ "$OSTYPE" == "cygwin" ]] || [[ "$OSTYPE" == "win32" ]]; then
    ARGV0=rg "$_cc_bin" "$@"
  elif [[ $BASHPID != $$ ]]; then
    exec -a rg "$_cc_bin" "$@"
  else
    (exec -a rg "$_cc_bin" "$@")
  fi
}
fi
export PATH='/c/Users/namma/bin:/mingw64/bin:/usr/local/bin:/usr/bin:/bin:/mingw64/bin:/usr/bin:/c/Users/namma/bin:/c/windows/system32:/c/windows:/c/windows/System32/Wbem:/c/windows/System32/WindowsPowerShell/v1.0:/c/windows/System32/OpenSSH:/c/Program Files/dotnet:/c/Program Files/Bandizip:/c/Program Files/Microsoft SQL Server/170/Tools/Binn:/c/Program Files/Microsoft SQL Server/Client SDK/ODBC/170/Tools/Binn:/cmd:/c/Program Files (x86)/Windows Kits/10/Windows Performance Toolkit:/c/Program Files/PowerShell/7:/c/Users/namma/AppData/Local/Microsoft/WindowsApps:/c/Users/namma/AppData/Local/Programs/Microsoft VS Code/bin:/c/Users/namma/.dotnet/tools:/c/Users/namma/.local/bin:/mingw64/bin:/usr/bin/vendor_perl:/usr/bin/core_perl:/c/Users/namma/AppData/Roaming/Claude/local-agent-mode-sessions/skills-plugin/d3453b4f-497b-4e7e-bbff-aa1b470046d5/dc934002-beb5-48c7-84bf-566d0d5aefca/bin'
