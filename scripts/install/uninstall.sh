#!/usr/bin/env bash
# Takes back what install.sh set up, in the reverse order:
#
#   bash scripts/install/uninstall.sh
#
#   1. the logon task that starts WSL            (autostart.ps1 -Remove)
#   2. the registration of the Word add-in       (register-addin.ps1 -Unregister)
#   3. the systemd user service                  (daemon-service.sh uninstall)
#
# What is not there is passed over. Left behind on purpose: the daemon log
# and the other files of ~/.claude/office-addins (keys, the bridge token),
# and everything next to your manuscripts.
set -euo pipefail

HERE=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
REPO=$(cd "$HERE/../.." && pwd)

win() {
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(wslpath -w "$HERE/$1")" "${@:2}" |
    tr -d '\r'
}

if [ -n "${WSL_DISTRO_NAME:-}" ]; then
  echo "== 1/3 WSL at Windows logon"
  win autostart.ps1 -Remove
  echo "== 2/3 Word add-in"
  win register-addin.ps1 -Manifest "$(wslpath -w "$REPO/manifests/word.xml")" -Unregister
else
  echo "== 1/3, 2/3 skipped: this is not WSL"
fi

echo "== 3/3 daemon service"
bash "$HERE/daemon-service.sh" uninstall
