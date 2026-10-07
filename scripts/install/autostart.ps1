# Keeps the WSL distro running from Windows logon on, so the daemon (a
# systemd user service inside the distro) is up when Word opens. It
# registers one scheduled task for the current user that runs
#
#   conhost.exe --headless wsl.exe -d <distro> --exec sleep infinity
#
# at logon: a process of the distro that never ends, started without a
# console window. Current user only; no administrator rights and no stored
# password (the task runs only while the user is logged on).
#
#   autostart.ps1 -Distro <name>     register the task and start it now
#   autostart.ps1 -Remove            stop and remove it
#
# Run again, it leaves a task that is in place alone. The names here are the
# ones of scripts/install/win-args.mjs (TASK_NAME, autostartCommand); a test
# holds them together.
param(
  [string]$Distro,
  [switch]$Remove
)
$ErrorActionPreference = "Stop"

$name = "Sidewrite WSL keep-alive"
$task = Get-ScheduledTask -TaskName $name -TaskPath "\" -ErrorAction SilentlyContinue
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$me = $identity.Name

function Remove-LegacyTask {
  $legacyName = "Draftspect WSL keep-alive"
  $legacy = Get-ScheduledTask -TaskName $legacyName -TaskPath "\" -ErrorAction SilentlyContinue
  if (-not $legacy) { return }
  $actions = @($legacy.Actions)
  $triggers = @($legacy.Triggers)
  $users = @($me, $identity.User.Value)
  # Match the old installer's complete action and current-user logon scope.
  # Never remove a same-named task with a custom action or another principal.
  $owned = ($legacy.TaskPath -eq "\") -and
    ($actions.Count -eq 1) -and ($triggers.Count -eq 1) -and
    ($actions[0].Execute -eq "conhost.exe") -and
    (-not $actions[0].WorkingDirectory) -and
    ($actions[0].Arguments -cmatch '^--headless wsl\.exe -d [\w.-]+ --exec sleep infinity$') -and
    ($users -contains $legacy.Principal.UserId) -and
    ($legacy.Principal.LogonType -eq "Interactive") -and
    ($legacy.Principal.RunLevel -eq "Limited") -and
    ($triggers[0].CimClass.CimClassName -eq "MSFT_TaskLogonTrigger") -and
    ($users -contains $triggers[0].UserId)
  if ($Distro) {
    $owned = $owned -and ($actions[0].Arguments -ceq "--headless wsl.exe -d $Distro --exec sleep infinity")
  }
  if (-not $owned) {
    Write-Warning "legacy task: preserved custom task ($legacyName); review it manually"
    return
  }
  Stop-ScheduledTask -TaskName $legacyName -TaskPath "\"
  Unregister-ScheduledTask -TaskName $legacyName -TaskPath "\" -Confirm:$false
  Write-Output "legacy task: removed"
}

if ($Remove) {
  if ($task) {
    Stop-ScheduledTask -TaskName $name -TaskPath "\" -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $name -Confirm:$false -TaskPath "\"
    Write-Output "task: removed"
  } else {
    Write-Output "task: not there"
  }
  Remove-LegacyTask
  exit 0
}

if ($Distro -notmatch '^[\w.-]+$') {
  throw "Give the WSL distro: autostart.ps1 -Distro <name>"
}
$execute = "conhost.exe"
$argument = "--headless wsl.exe -d $Distro --exec sleep infinity"

$inPlace = $false
if ($task) {
  $a = @($task.Actions)[0]
  $inPlace = ($a.Execute -eq $execute) -and ($a.Arguments -eq $argument)
}
if ($inPlace) {
  Write-Output "task: in place"
} else {
  $action = New-ScheduledTaskAction -Execute $execute -Argument $argument
  $trigger = New-ScheduledTaskTrigger -AtLogOn -User $me
  $principal = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
  # Without these Task Scheduler ends the task after three days and does not
  # start it on battery.
  $settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -MultipleInstances IgnoreNew
  Register-ScheduledTask -TaskName $name -TaskPath "\" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Output "task: registered"
}
if ((Get-ScheduledTask -TaskName $name -TaskPath "\").State -ne "Running") {
  Start-ScheduledTask -TaskName $name -TaskPath "\"
  Write-Output "task: started"
}
# Keep WSL alive through the handover, then retire the exact old task.
Remove-LegacyTask
