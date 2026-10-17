# scripts/install/

## OVERVIEW
Install, uninstall and health check for the resident daemon (systemd user unit on Linux/WSL) and the Word add-in registration plus autostart on Windows. Earned its file: distinct domain (OS service lifecycle, registry, scheduled task; separate from the Electron sideload in `app/`).

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| One-shot install / uninstall | `install.sh`, `uninstall.sh` | |
| systemd unit lifecycle | `daemon-service.sh` | `install|uninstall|status|logs`; migration rollback, health check |
| Unit text and naming | `unit.mjs` | `renderUnit`, `unitName`, `legacyUnitName`, `HEALTH_URL`, `isLegacyUnit`, `isLegacyExecStart` |
| Windows registry + task constants | `win-args.mjs` | `DEVELOPER_KEY`, `APP_FOLDER`, `MANIFEST`, `TASK_NAME`, `registerPlan`, `autostartArgs` |
| Windows registration | `register-addin.ps1` | `HKCU\...\WEF\Developer` |
| Windows autostart task | `autostart.ps1` | `-Distro <name>` or `-Remove` |
| Doctor | `doctor.mjs` (facts, `windowsProbe`), `doctor-judge.mjs` (pure `judgeDoctor`, `failures`, `report`) | |

## CONVENTIONS
- Pure/OS split: `unit.mjs`, `win-args.mjs`, `doctor-judge.mjs` are pure and unit-tested; shell/PowerShell scripts consume them.
- Constants in `win-args.mjs` are mirrored as literals in the `.ps1` files; `tests/install-win.test.mjs` asserts they match. Change both.
- Every operation is idempotent; rerunning install converges.
- Unit `PATH` = node directory + directories passed at install + `/usr/local/bin:/usr/bin:/bin` (`unit.mjs` l.19-32); reinstall after moving tools. systemd reads a blank as word end and `:` as PATH separator (l.21); handle such directories there, never by hand-editing the unit.
- Imported elsewhere: `scripts/e2e/resident-unit.mjs` and `run-acceptance.mjs` use `unit.mjs`.

## ANTI-PATTERNS
- Never remove a legacy systemd service or scheduled task by name alone; full content, effective command, ownership and overrides must match the known old installer. Unknown custom state is kept.
- Never register through a local-path Trusted Catalog or UNC share; Office ignores it.
- Never require admin rights; HKCU only.

## TESTS
```bash
node --test tests/install-unit.test.mjs tests/install-win.test.mjs tests/install-all.test.mjs tests/doctor.test.mjs
```
`install-unit` runs `install.sh` via `spawnSync("bash", ...)` with a fake PATH (`tests/fixtures/fake-bin.mjs`, `fake-systemctl.sh`); no real systemd is touched.
