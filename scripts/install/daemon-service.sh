#!/usr/bin/env bash
# The daemon as a systemd user service (WSL with systemd, or Linux).
#
#   daemon-service.sh install     write the unit, enable and start it, keep the
#                                 user manager running without a login
#                                 (linger), wait until /health answers
#   daemon-service.sh uninstall   stop and disable it, remove the unit; linger
#                                 is turned off only if install turned it on
#   daemon-service.sh status      the unit and /health
#   daemon-service.sh logs        follow the daemon log
#
# install can be run again at any time: what is in place is left alone, a
# unit that differs from the one this checkout renders is replaced and the
# daemon restarted. The unit starts THIS checkout (the folder the script is
# in) with the node found on PATH now.
#
# Exit: 0 done, 1 the daemon did not come up, 2 no systemd, 64 usage.
set -euo pipefail

HERE=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
NAME=$(node "$HERE/unit.mjs" name)
SERVICE="$NAME.service"
UNIT_DIR="$HOME/.config/systemd/user"
UNIT="$UNIT_DIR/$SERVICE"
STATE_DIR="$HOME/.claude/office-addins"
LOG="$STATE_DIR/daemon.log"
# Present while linger is on because this script turned it on.
LINGER_MARK="$STATE_DIR/linger-set-by-install"
HEALTH="http://127.0.0.1:47834/health"
ME=${USER:-$(id -un)}

need_systemd() {
  if [ "$(ps -p 1 -o comm=)" != systemd ]; then
    {
      echo "This system does not run systemd, so there is no user service to install."
      echo "In WSL: add the two lines below to /etc/wsl.conf, then run 'wsl --shutdown'"
      echo "in Windows and open WSL again."
      echo "  [boot]"
      echo "  systemd=true"
    } >&2
    exit 2
  fi
}

wait_health() {
  for _ in $(seq 20); do
    if curl -fsS "$HEALTH" >/dev/null 2>&1; then return 0; fi
    sleep 0.5
  done
  return 1
}

install() {
  need_systemd
  mkdir -p "$UNIT_DIR" "$STATE_DIR"
  # Directories of the tools the daemon starts, as found now.
  local dirs=() tool path
  for tool in claude codex python3; do
    path=$(command -v "$tool" || true)
    if [ -n "$path" ]; then dirs+=("$(dirname "$path")"); fi
  done
  local want
  want=$(node "$HERE/unit.mjs" render "$LOG" "${dirs[@]}")
  if [ -f "$UNIT" ] && [ "$want" = "$(cat "$UNIT")" ]; then
    echo "unit: in place ($UNIT)"
    systemctl --user enable --now "$SERVICE"
  else
    printf '%s\n' "$want" >"$UNIT"
    echo "unit: written ($UNIT)"
    systemctl --user daemon-reload
    systemctl --user enable "$SERVICE"
    systemctl --user restart "$SERVICE"
  fi
  if [ "$(loginctl show-user "$ME" -p Linger --value)" = yes ]; then
    echo "linger: already on"
  else
    loginctl enable-linger "$ME"
    : >"$LINGER_MARK"
    echo "linger: turned on"
  fi
  if wait_health; then
    echo "daemon is up: $HEALTH"
  else
    {
      echo "The daemon did not answer $HEALTH within 10 s."
      echo "See why: $HERE/daemon-service.sh logs"
    } >&2
    exit 1
  fi
}

uninstall() {
  systemctl --user disable --now "$SERVICE" 2>/dev/null || true
  rm -f "$UNIT"
  systemctl --user daemon-reload || true
  if [ -f "$LINGER_MARK" ]; then
    loginctl disable-linger "$ME" || true
    rm -f "$LINGER_MARK"
  fi
  echo "removed: $UNIT (the log stays at $LOG)"
}

case "${1:-}" in
  install) install ;;
  uninstall) uninstall ;;
  status)
    systemctl --user --no-pager status "$SERVICE" || true
    curl -fsS "$HEALTH" || echo "no answer from $HEALTH"
    echo
    ;;
  logs) exec tail -n 100 -f "$LOG" ;;
  *)
    echo "usage: daemon-service.sh install|uninstall|status|logs" >&2
    exit 64
    ;;
esac
