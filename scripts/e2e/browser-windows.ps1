# The browser windows around the "Figures and other files" button of the
# acceptance run. Prints one JSON object.
#
# Actions
#   list    the processes of the usual browsers (msedge, chrome, firefox,
#           brave, opera, vivaldi): pid, name, title of the main window and
#           whether there is one. Read only.
#   close   -ProcessId <n> -Title <s>: closes the main window of that
#           process with CloseMainWindow (as a click on its X would), and
#           only when, read again now, it is a browser of the list, its main
#           window shows the title, and UI Automation finds exactly one tab
#           in that window. Otherwise nothing is closed and the answer says
#           why. No process is ever stopped by force. Which process to name is decided by the
#           runner (judge.mjs browserCloseDecision): one that started after
#           the click, of a browser no process of which ran before it.
param(
  [Parameter(Mandatory = $true)][ValidateSet("list", "close")][string]$Action,
  [int]$ProcessId,
  [string]$Title
)
$ErrorActionPreference = "Stop"
$Names = @("msedge", "chrome", "firefox", "brave", "opera", "vivaldi")

function Out($o) { $o | ConvertTo-Json -Compress -Depth 5 }

function Browsers {
  @(Get-Process -Name $Names -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
      pid = $_.Id
      name = $_.ProcessName
      title = [string]$_.MainWindowTitle
      has_window = ($_.MainWindowHandle -ne [IntPtr]::Zero)
    }
  })
}

switch ($Action) {
  "list" {
    Out ([ordered]@{ ok = $true; processes = @(Browsers) })
  }
  "close" {
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $p) { Out ([ordered]@{ ok = $false; closed = $false; reason = "no such process" }); break }
    if ($Names -notcontains $p.ProcessName) {
      Out ([ordered]@{ ok = $false; closed = $false; reason = "not a browser of the list: $($p.ProcessName)" }); break
    }
    if ($p.MainWindowHandle -eq [IntPtr]::Zero -or -not $Title -or -not ([string]$p.MainWindowTitle).Contains($Title)) {
      Out ([ordered]@{ ok = $false; closed = $false; reason = "its window does not show the title"; title = [string]$p.MainWindowTitle }); break
    }
    $tabs = $null
    try {
      Add-Type -AssemblyName UIAutomationClient
      Add-Type -AssemblyName UIAutomationTypes
      $root = [System.Windows.Automation.AutomationElement]::FromHandle($p.MainWindowHandle)
      $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::TabItem)
      $tabs = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond).Count
    } catch {
      Out ([ordered]@{ ok = $false; closed = $false; reason = "the tabs could not be counted: $($_.Exception.Message)" }); break
    }
    if ($tabs -ne 1) {
      Out ([ordered]@{ ok = $false; closed = $false; reason = "the window has $tabs tabs"; tabs = $tabs }); break
    }
    [void]$p.CloseMainWindow()
    $gone = $p.WaitForExit(10000)
    Out ([ordered]@{ ok = $true; closed = $true; exited = $gone; tabs = $tabs; title = [string]$p.MainWindowTitle })
  }
}
