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

$name = "Draftspect WSL keep-alive"
$task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue

if ($Remove) {
  if ($task) {
    Stop-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $name -Confirm:$false
    Write-Output "task: removed"
  } else {
    Write-Output "task: not there"
  }
  exit 0
}

if (-not $Distro) {
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
  $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $action = New-ScheduledTaskAction -Execute $execute -Argument $argument
  $trigger = New-ScheduledTaskTrigger -AtLogOn -User $me
  $principal = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
  # Without these Task Scheduler ends the task after three days and does not
  # start it on battery.
  $settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -MultipleInstances IgnoreNew
  Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Output "task: registered"
}
if ((Get-ScheduledTask -TaskName $name).State -ne "Running") {
  Start-ScheduledTask -TaskName $name
  Write-Output "task: started"
}
