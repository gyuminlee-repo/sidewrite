import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { LEGACY_MANIFEST_HASHES } from "../app/legacy-manifests.mjs";
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
    dir: "C:\\Users\\u\\AppData\\Roaming\\Sidewrite\\manifests",
    path: "C:\\Users\\u\\AppData\\Roaming\\Sidewrite\\manifests\\word.xml",
    key: "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer",
    valueName: "C:\\Users\\u\\AppData\\Roaming\\Sidewrite\\manifests\\word.xml",
    valueData: "C:\\Users\\u\\AppData\\Roaming\\Sidewrite\\manifests\\word.xml",
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

test("legacy PowerShell manifest source contract: exact hashes, paths and value ownership", async () => {
  // Source-contract checks only: PowerShell/Office are not available in this
  // test environment. JS behavior is covered with stubbed reg.exe separately.
  const ps = await read("scripts/install/register-addin.ps1");
  for (const [file, hash] of Object.entries(LEGACY_MANIFEST_HASHES)) {
    assert.ok(ps.includes(`"${file}" = "${hash}"`));
  }
  assert.match(ps, /if \(-not \$hashes\.ContainsKey\(\$file\)\) \{ return \}/);
  assert.match(ps, /\$legacyRoot = Join-Path \$env:APPDATA "Draftspect"/);
  assert.match(ps, /foreach \(\$path in @\(\$legacyRoot, \$legacyDir, \$legacy\)\)/);
  assert.match(ps, /FileAttributes\]::ReparsePoint/);
  assert.match(ps, /if \(\$hash -ne \$hashes\[\$file\]\)/);
  assert.match(ps, /GetValueKind\(\$legacy\) -ne \[Microsoft\.Win32\.RegistryValueKind\]::String/);
  assert.match(ps, /\$value -ne \$legacy/);
  assert.match(ps, /Remove-ItemProperty -LiteralPath \$keyPath -Name \$legacy/);
  assert.match(ps, /if \(\$exists\) \{ Remove-Item -LiteralPath \$legacy -Force \}/);
  assert.doesNotMatch(ps, /Remove-Item[^\n]*-Recurse|Remove-ItemProperty[^\n]*\*/);
  assert.match(ps, /\.Replace\("`r`n", "`n"\)/);
  assert.equal(ps.match(/^\s*Remove-LegacyManifest$/gm).length, 2, "install and uninstall");
  assert.ok(ps.lastIndexOf("Remove-LegacyManifest") > ps.indexOf("Set-ItemProperty"));
});

test("legacy PowerShell task source contract: root task, current user, exact single action", async () => {
  const ps = await read("scripts/install/autostart.ps1");
  const cleanup = ps.slice(ps.indexOf("function Remove-LegacyTask"), ps.indexOf("if ($Remove)"));
  assert.ok(cleanup.includes('$legacyName = "Draftspect WSL keep-alive"'));
  assert.ok(cleanup.includes('Get-ScheduledTask -TaskName $legacyName -TaskPath "\\"'));
  for (const condition of [
    '$legacy.TaskPath -eq "\\"',
    "$actions.Count -eq 1",
    "$triggers.Count -eq 1",
    '$actions[0].Execute -eq "conhost.exe"',
    "-not $actions[0].WorkingDirectory",
    "'^--headless wsl\\.exe -d [\\w.-]+ --exec sleep infinity$'",
    "$users -contains $legacy.Principal.UserId",
    '$legacy.Principal.LogonType -eq "Interactive"',
    '$legacy.Principal.RunLevel -eq "Limited"',
    '$triggers[0].CimClass.CimClassName -eq "MSFT_TaskLogonTrigger"',
    "$users -contains $triggers[0].UserId",
    '$actions[0].Arguments -ceq "--headless wsl.exe -d $Distro --exec sleep infinity"',
  ])
    assert.ok(cleanup.includes(condition), condition);
  assert.match(cleanup, /if \(-not \$owned\) \{[\s\S]*?return\s*\}/);
  assert.ok(cleanup.indexOf("return\n  }") < cleanup.indexOf("Stop-ScheduledTask"));
  assert.match(cleanup, /Stop-ScheduledTask[^\n]*\n\s*Unregister-ScheduledTask/);
  assert.doesNotMatch(cleanup, /Stop-ScheduledTask[^\n]*SilentlyContinue|TaskName[^\n]*\*/);
  assert.equal(ps.match(/^\s*Remove-LegacyTask$/gm).length, 2, "install and remove");
  assert.ok(ps.lastIndexOf("Remove-LegacyTask") > ps.indexOf("Start-ScheduledTask"));
});

test("legacy PowerShell migration requires the exact replacement registration before cleanup", async () => {
  // Source-contract regression only; this does not execute PowerShell or
  // touch a Windows registry. A non-null stale/wrong-type value must not
  // let install retire the known-good Draftspect registration.
  const ps = await read("scripts/install/register-addin.ps1");
  const guard = ps.slice(ps.indexOf("$replacementKey ="));
  assert.match(guard, /^\$replacementKey = Get-Item -LiteralPath \$keyPath/);
  assert.match(guard, /GetValueKind\(\$dst\) -ne \[Microsoft\.Win32\.RegistryValueKind\]::String/);
  assert.match(guard, /-or\s*\(\$replacementKey\.GetValue\(\$dst, \$null\) -ne \$dst\)/);
  assert.match(guard, /throw "Preserved an unexpected Sidewrite registration/);
  assert.ok(guard.indexOf("throw ") < guard.indexOf("Remove-LegacyManifest"));
  assert.ok(ps.indexOf("$replacementKey =") > ps.indexOf("Set-ItemProperty"));
  assert.doesNotMatch(guard, /Set-ItemProperty|Remove-ItemProperty/);
  assert.match(ps, /\$ErrorActionPreference = "Stop"/);
});
