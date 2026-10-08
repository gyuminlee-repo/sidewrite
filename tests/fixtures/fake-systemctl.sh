# State-only systemctl stand-in, executed by fake-bin.mjs after logging args.
# Everything is confined to the fake HOME. No systemctl is invoked here.
if [ "$*" = "${FAKE_SYSTEMCTL_FAIL:-}" ]; then exit 1; fi
if [ "$*" = "${FAKE_SYSTEMCTL_NOOP:-}" ]; then exit 0; fi

case "$*" in
  *draftspect-daemon.service*)
    state_file="$HOME/legacy-service-state"
    active=${FAKE_LEGACY_ACTIVE:-active}
    enabled=${FAKE_LEGACY_ENABLED:-enabled}
    ;;
  *sidewrite-daemon.service*)
    state_file="$HOME/sidewrite-service-state"
    active=inactive
    enabled=disabled
    ;;
  *) exit 0 ;;
esac
if [ -f "$state_file" ]; then read -r active enabled <"$state_file"; fi

case "$*" in
  *DropInPaths*) echo "${FAKE_LEGACY_DROPINS:-}"; exit 0 ;;
  *FragmentPath*) echo "${FAKE_LEGACY_FRAGMENT:-$HOME/.config/systemd/user/draftspect-daemon.service}"; exit 0 ;;
  *ExecStart*)
    if [ -n "${FAKE_LEGACY_EXEC:-}" ]; then echo "$FAKE_LEGACY_EXEC";
    else echo '{ path=/old/bin/node ; argv[]=/old/bin/node /old/checkout/daemon/index.mjs ; ignore_errors=no ; start_time=[n/a] ; }'; fi
    exit 0
    ;;
  *is-enabled*) echo "$enabled"; [ "$enabled" = enabled ]; exit $? ;;
  *is-active*) echo "$active"; [ "$active" = active ]; exit $? ;;
  *" disable "*) enabled=disabled; if [[ "$*" == *--now* ]]; then active=inactive; fi ;;
  *" enable "*) enabled=enabled; if [[ "$*" == *--now* ]]; then active=active; fi ;;
  *" stop "*) active=inactive ;;
  *" start "*|*" restart "*) active=active ;;
esac
printf '%s %s\n' "$active" "$enabled" >"$state_file"
