// The names and arguments of the Windows side of the install, in one place:
// what register-addin.ps1 writes to the registry and what autostart.ps1 asks
// Task Scheduler to run. The PowerShell scripts repeat these as literals (a
// .ps1 cannot import this file); tests/install-win.test.mjs holds the two
// together, and the doctor reads them from here.

// The same key, folder and file as app/sideload.mjs (winInstall): Office
// reads each value NAME under the key as the path of a manifest to sideload.
export const DEVELOPER_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";
export const APP_FOLDER = "Sidewrite";
export const MANIFEST = "word.xml";

// appData: %APPDATA% of the Windows user.
export function registerPlan({ appData, manifest = MANIFEST, name = APP_FOLDER }) {
  const dir = `${appData}\\${name}\\manifests`;
  const path = `${dir}\\${manifest}`;
  return { dir, path, key: DEVELOPER_KEY, valueName: path, valueData: path, type: "REG_SZ" };
}

export const TASK_NAME = "Sidewrite WSL keep-alive";

// What keeps the WSL distro running after logon: a process of it that never
// ends. The systemd user service in the distro then keeps the daemon up.
export function autostartArgs({ distro }) {
  if (!/^[\w.-]+$/.test(String(distro ?? "")))
    throw new Error(`distro must be the name of a WSL distro: ${JSON.stringify(distro)}`);
  return ["-d", distro, "--exec", "sleep", "infinity"];
}

// The action of the scheduled task. wsl.exe is a console program; started
// through conhost --headless it shows no window.
export function autostartCommand({ distro }) {
  return {
    execute: "conhost.exe",
    argument: `--headless wsl.exe ${autostartArgs({ distro }).join(" ")}`,
  };
}
