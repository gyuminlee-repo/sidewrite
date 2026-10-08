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
LEGACY_SERVICE="$(node "$HERE/unit.mjs" legacy-name).service"
LEGACY_UNIT="$UNIT_DIR/$LEGACY_SERVICE"
STATE_DIR="$HOME/.claude/office-addins"
LOG="$STATE_DIR/daemon.log"
# Present while linger is on because this script turned it on.
LINGER_MARK="$STATE_DIR/linger-set-by-install"
HEALTH="http://127.0.0.1:47834/health"
ME=${USER:-$(id -un)}
LEGACY_STAGED=0
SIDEWRITE_TOUCHED=0

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

legacy_owned() {
  if [ ! -e "$LEGACY_UNIT" ] && [ ! -L "$LEGACY_UNIT" ]; then return 1; fi
  if [ -L "$LEGACY_UNIT" ] || [ ! -f "$LEGACY_UNIT" ] || [ ! -O "$LEGACY_UNIT" ] ||
     [ -e "$LEGACY_UNIT.d" ] || [ -L "$LEGACY_UNIT.d" ] ||
     ! node "$HERE/unit.mjs" is-legacy "$LEGACY_UNIT"; then
    echo "legacy unit: preserved custom service ($LEGACY_UNIT); review it manually" >&2
    return 1
  fi
  # Drop-ins may also live in runtime/system user-unit directories, not only
  # beside this file. Ask the manager rather than ignoring those overrides.
  local dropins fragment effective
  if ! dropins=$(systemctl --user show "$LEGACY_SERVICE" -p DropInPaths --value) || [ -n "$dropins" ]; then
    echo "legacy unit: preserved custom service ($LEGACY_UNIT); cannot rule out drop-ins" >&2
    return 1
  fi
  if ! fragment=$(systemctl --user show "$LEGACY_SERVICE" -p FragmentPath --value) ||
     [ "$fragment" != "$LEGACY_UNIT" ] ||
     ! effective=$(systemctl --user show "$LEGACY_SERVICE" -p ExecStart --value) ||
     ! node "$HERE/unit.mjs" is-legacy-exec "$LEGACY_UNIT" "$effective"; then
    echo "legacy unit: preserved custom service ($LEGACY_UNIT); effective unit does not match" >&2
    return 1
  fi
}

cleanup_legacy() {
  if ! legacy_owned; then return; fi
  systemctl --user disable --now "$LEGACY_SERVICE"
  rm -- "$LEGACY_UNIT"
  systemctl --user daemon-reload
  echo "legacy unit: removed ($LEGACY_UNIT)"
}

rollback_legacy() {
  if [ "$LEGACY_STAGED" != 1 ]; then return; fi
  echo "Sidewrite install failed; restoring the previous legacy service." >&2
  local failed=0
  local safe_to_restart=1 state
  if [ "$SIDEWRITE_TOUCHED" = 1 ]; then
    if ! systemctl --user disable --now "$SERVICE"; then failed=1; safe_to_restart=0; fi
    state=$(systemctl --user is-active "$SERVICE" || true)
    case "$state" in inactive|failed) ;; *) failed=1; safe_to_restart=0 ;; esac
    state=$(systemctl --user is-enabled "$SERVICE" || true)
    if [ "$state" != disabled ]; then failed=1; safe_to_restart=0; fi
  fi
  printf '%s\n' "$LEGACY_CONTENT" >"$LEGACY_UNIT" || failed=1
  systemctl --user daemon-reload || failed=1
  if [ "$LEGACY_ENABLED" = enabled ]; then
    systemctl --user enable "$LEGACY_SERVICE" || failed=1
  else
    systemctl --user disable "$LEGACY_SERVICE" || failed=1
  fi
  state=$(systemctl --user is-enabled "$LEGACY_SERVICE" || true)
  if [ "$state" != "$LEGACY_ENABLED" ]; then failed=1; fi
  if [ "$LEGACY_ACTIVE" = active ] && [ "$safe_to_restart" = 1 ]; then
    systemctl --user start "$LEGACY_SERVICE" || failed=1
    state=$(systemctl --user is-active "$LEGACY_SERVICE" || true)
    if [ "$state" != active ]; then failed=1; fi
  fi
  if [ "$safe_to_restart" != 1 ]; then
    echo "Sidewrite could not be confirmed stopped/disabled; legacy restart was withheld to avoid a port conflict." >&2
  fi
  if [ "$failed" = 1 ]; then
    echo "Legacy rollback could not finish; inspect both services before retrying." >&2
  fi
}

stage_legacy() {
  if ! legacy_owned; then return; fi
  LEGACY_CONTENT=$(cat "$LEGACY_UNIT")
  LEGACY_ENABLED=$(systemctl --user is-enabled "$LEGACY_SERVICE" || true)
  LEGACY_ACTIVE=$(systemctl --user is-active "$LEGACY_SERVICE" || true)
  # The old installer creates a normal enabled/disabled user unit. Unknown
  # states are left alone rather than guessing what to restore on failure.
  case "$LEGACY_ENABLED/$LEGACY_ACTIVE" in
    enabled/active|enabled/inactive|enabled/failed|disabled/active|disabled/inactive|disabled/failed) ;;
    *) echo "Cannot safely migrate legacy service state: $LEGACY_ENABLED/$LEGACY_ACTIVE" >&2; return 1 ;;
  esac
  LEGACY_STAGED=1
  trap 'rc=$?; if [ "$rc" != 0 ]; then rollback_legacy; fi; exit "$rc"' EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  # Keep its unit and enablement until the replacement is healthy. Only stop
  # it for now, since both daemons bind the same ports.
  systemctl --user stop "$LEGACY_SERVICE"
  local state
  state=$(systemctl --user is-active "$LEGACY_SERVICE" || true)
  case "$state" in
    inactive|failed) ;;
    *) echo "Legacy service did not stop; refusing to start a second daemon." >&2; return 1 ;;
  esac
}

install() {
  need_systemd
  mkdir -p "$UNIT_DIR" "$STATE_DIR"
  # Directories of the tools the daemon starts, as found now.
  local dirs=() tool path
  for tool in claude codex agy python3; do
    path=$(command -v "$tool" || true)
    if [ -n "$path" ]; then dirs+=("$(dirname "$path")"); fi
  done
  local want
  want=$(node "$HERE/unit.mjs" render "$LOG" "${dirs[@]}")
  stage_legacy
  SIDEWRITE_TOUCHED=1
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
    if [ "$LEGACY_STAGED" = 1 ]; then
      if [ "$(systemctl --user is-active "$SERVICE" || true)" != active ]; then
        echo "Sidewrite service is not active; a healthy port alone does not verify the handover." >&2
        return 1
      fi
      cleanup_legacy
      LEGACY_STAGED=0
      trap - EXIT INT TERM
    fi
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
  cleanup_legacy
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
