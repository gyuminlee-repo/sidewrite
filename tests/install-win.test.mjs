import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  APP_FOLDER,
  DEVELOPER_KEY,
  MANIFEST,
  TASK_NAME,
  autostartArgs,
  autostartCommand,
  registerPlan,
} from "../scripts/install/win-args.mjs";

const read = (rel) => readFile(new URL(`../${rel}`, import.meta.url), "utf8");

test("registerPlan: the key, folder and value the tray app's sideload writes", async () => {
  const plan = registerPlan({ appData: "C:\\Users\\u\\AppData\\Roaming" });
  assert.deepEqual(plan, {
    dir: "C:\\Users\\u\\AppData\\Roaming\\Draftspect\\manifests",
    path: "C:\\Users\\u\\AppData\\Roaming\\Draftspect\\manifests\\word.xml",
    key: "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer",
    valueName: "C:\\Users\\u\\AppData\\Roaming\\Draftspect\\manifests\\word.xml",
    valueData: "C:\\Users\\u\\AppData\\Roaming\\Draftspect\\manifests\\word.xml",
    type: "REG_SZ",
  });
  // The same constants as app/sideload.mjs (winInstall): a manifest the tray
  // app registered and one this script registered are one entry, not two.
  const sideload = await read("app/sideload.mjs");
  assert.ok(sideload.includes(JSON.stringify(DEVELOPER_KEY)), "the Developer key");
  assert.ok(sideload.includes(`"AppData", "Roaming", "${APP_FOLDER}", "manifests"`));
  assert.ok(sideload.includes(`file: "${MANIFEST}"`));
  assert.match(sideload, /regAddString\(WIN_DEVELOPER_KEY, dst, dst\)/);
  assert.match(sideload, /"\/t", "REG_SZ"/);
});

test("autostartArgs: wsl.exe keeps the distro running with a sleep that never ends", () => {
  assert.deepEqual(autostartArgs({ distro: "Ubuntu-24.04" }), [
    "-d",
    "Ubuntu-24.04",
    "--exec",
    "sleep",
    "infinity",
  ]);
  assert.throws(() => autostartArgs({ distro: "" }), /distro/);
  assert.throws(() => autostartArgs({ distro: "a b" }), /distro/);
  // Started through conhost so no console window is shown.
  assert.deepEqual(autostartCommand({ distro: "Ubuntu" }), {
    execute: "conhost.exe",
    argument: "--headless wsl.exe -d Ubuntu --exec sleep infinity",
  });
});

// The two PowerShell scripts cannot run under test. What is checked here is
// that they are plain ASCII, that their brackets close, and that they hold
// the same names and arguments as the functions above.
const closes = (text) => {
  const stack = [];
  const pair = { ")": "(", "}": "{", "]": "[" };
  // Strings and comments are taken out first: a bracket in them is text.
  const code = text
    .replace(/#.*$/gm, "")
    .replace(/"(?:[^"`]|`.)*"/g, '""')
    .replace(/'[^']*'/g, "''");
  for (const ch of code) {
    if ("({[".includes(ch)) stack.push(ch);
    else if (pair[ch] && stack.pop() !== pair[ch]) return false;
  }
  return stack.length === 0;
};

test("register-addin.ps1 holds the key and folder of registerPlan", async () => {
  const ps = await read("scripts/install/register-addin.ps1");
  assert.doesNotMatch(ps, /[^\x09\x0a\x0d\x20-\x7e]/, "ASCII only");
  assert.ok(closes(ps), "brackets close");
  assert.ok(ps.includes(DEVELOPER_KEY.replace(/^HKCU\\/, "HKCU:\\")), "the Developer key");
  assert.ok(ps.includes(`"${APP_FOLDER}\\manifests"`));
  assert.match(ps, /^param\(/m);
  assert.match(ps, /\[switch\]\$Unregister/);
  assert.match(ps, /\$ErrorActionPreference = "Stop"/);
  // The value: its name and its data are the full path of the copied manifest.
  assert.match(ps, /-Name \$dst -Value \$dst -Type String/);
  // The current user's hive only; nothing that asks for elevation.
  assert.doesNotMatch(ps, /HKLM|RunAs|Start-Process/);
});

test("autostart.ps1 registers the task of autostartCommand, for this user, without elevation", async () => {
  const ps = await read("scripts/install/autostart.ps1");
  assert.doesNotMatch(ps, /[^\x09\x0a\x0d\x20-\x7e]/, "ASCII only");
  assert.ok(closes(ps), "brackets close");
  assert.ok(ps.includes(`$name = "${TASK_NAME}"`));
  const cmd = autostartCommand({ distro: "DISTRO" });
  assert.ok(ps.includes(`$execute = "${cmd.execute}"`));
  assert.ok(ps.includes(`$argument = "${cmd.argument.replace("DISTRO", "$Distro")}"`));
  assert.match(ps, /\[switch\]\$Remove/);
  assert.match(ps, /New-ScheduledTaskTrigger -AtLogOn -User \$me/);
  assert.match(ps, /-LogonType Interactive -RunLevel Limited/);
  assert.doesNotMatch(ps, /RunLevel Highest|RunAs|-Password/);
  // Task Scheduler's defaults would stop the sleep after three days and
  // would not start it on battery.
  assert.match(ps, /-ExecutionTimeLimit \(New-TimeSpan -Seconds 0\)/);
  assert.match(ps, /-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries/);
  assert.match(ps, /-Hidden/);
  assert.match(ps, /Unregister-ScheduledTask -TaskName \$name -Confirm:\$false/);
});
