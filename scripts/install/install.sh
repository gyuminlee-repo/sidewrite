#!/usr/bin/env bash
# One command for a Windows PC with WSL2: after it, the daemon runs without a
# terminal and Word shows the add-in.
#
#   bash scripts/install/install.sh
#
#   1. the daemon as a systemd user service      (daemon-service.sh install)
#   2. the Word add-in registered for the user   (register-addin.ps1)
#   3. WSL started at Windows logon              (autostart.ps1)
#   4. the doctor: what is in place, what is not (doctor.mjs)
#
# Every step leaves alone what is already in place, so the command can be run
# again at any time (after moving the checkout, after a Node upgrade). It
# sets no environment variable and installs no package. Undo: uninstall.sh.
#
# The two PowerShell scripts are started with -ExecutionPolicy Bypass for
# that one process (Windows refuses unsigned script files by default); the
# policy of the machine is not changed.
#
# Exit: the number of problems the doctor found (0: everything in place); a
# step that fails stops the install with its own code.
set -euo pipefail

HERE=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
REPO=$(cd "$HERE/../.." && pwd)

# win <script.ps1> <args...>: run one of our PowerShell scripts in Windows.
win() {
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(wslpath -w "$HERE/$1")" "${@:2}" |
    tr -d '\r'
}

echo "== 1/4 daemon service"
bash "$HERE/daemon-service.sh" install

if [ -n "${WSL_DISTRO_NAME:-}" ]; then
  echo "== 2/4 Word add-in"
  win register-addin.ps1 -Manifest "$(wslpath -w "$REPO/manifests/word.xml")"
  echo "== 3/4 WSL at Windows logon"
  win autostart.ps1 -Distro "$WSL_DISTRO_NAME"
else
  echo "== 2/4, 3/4 skipped: this is not WSL, there is no Windows side to set up"
fi

echo "== 4/4 doctor"
node "$HERE/doctor.mjs"
