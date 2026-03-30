<#
.SYNOPSIS
    Checks GitHub for new GeekMagic SmallTV-Ultra firmware versions.
.DESCRIPTION
    Queries the GitHub API for the latest version directory (Ultra-V*),
    compares it against the last known version, shows a Windows Toast
    notification if a newer version is available, or installs the script
    into Windows Task Scheduler when requested.
.PARAMETER InstallScheduledTask
    Installs this script as a recurring Windows Task Scheduler task and exits
    without performing a version check.
.PARAMETER TaskName
    The Task Scheduler entry name to create when -InstallScheduledTask is used.
.PARAMETER CheckIntervalMinutes
    The number of minutes between scheduled runs.
.PARAMETER StartTime
    The first daily start time for the recurring task in 24-hour HH:mm format.
.PARAMETER Force
    Replaces an existing scheduled task with the same name when installing.
#>

[CmdletBinding()]
param(
    [switch]$InstallScheduledTask,
    [string]$TaskName = "GeekMagicClock Version Checker",
    [ValidateRange(5, 1439)]
    [int]$CheckIntervalMinutes = 60,
    [ValidatePattern('^(?:[01]\d|2[0-3]):[0-5]\d$')]
    [string]$StartTime = "09:00",
    [switch]$Force
)

$StateFile   = "$PSScriptRoot\last_known_version.txt"
$LogFile     = "$PSScriptRoot\version_check.log"
$ApiUrl      = "https://api.github.com/repos/GeekMagicClock/smalltv-ultra/contents"
$RepoUrl     = "https://github.com/GeekMagicClock/smalltv-ultra"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$timestamp] $Message"
}

function Get-HiddenLauncherPath {
    return Join-Path -Path $PSScriptRoot -ChildPath "Run-GeekMagicVersionChecker.vbs"
}

function Ensure-HiddenLauncher {
    $LauncherPath = Get-HiddenLauncherPath
    $LauncherContent = @"
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
quote = Chr(34)
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\Check-GeekMagicVersion.ps1"
powershell = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
command = quote & powershell & quote & " -NoProfile -ExecutionPolicy Bypass -File " & quote & scriptPath & quote
shell.Run command, 0, True
"@

    Set-Content -Path $LauncherPath -Value $LauncherContent -NoNewline -Encoding ASCII
    return $LauncherPath
}

function Show-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Url
    )
    $WindowsPowerShellAppId = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"

    $ToastScript = @'
param(
    [string]$Title,
    [string]$Message,
    [string]$Url,
    [string]$AppId
)

[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

$EscapedTitle = [System.Security.SecurityElement]::Escape($Title)
$EscapedMessage = [System.Security.SecurityElement]::Escape($Message)
$EscapedUrl = [System.Security.SecurityElement]::Escape($Url)

$ToastXml = @"
<toast activationType="protocol" launch="$EscapedUrl">
  <visual>
    <binding template="ToastGeneric">
      <text>$EscapedTitle</text>
      <text>$EscapedMessage</text>
      <text>Click to open the GitHub repository</text>
    </binding>
  </visual>
  <actions>
    <action content="Open GitHub" activationType="protocol" arguments="$EscapedUrl" />
  </actions>
</toast>
"@

$Xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$Xml.LoadXml($ToastXml)
$Toast = [Windows.UI.Notifications.ToastNotification]::new($Xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($Toast)
'@

    if ($PSVersionTable.PSEdition -eq "Desktop") {
        & ([scriptblock]::Create($ToastScript)) -Title $Title -Message $Message -Url $Url -AppId $WindowsPowerShellAppId
        return
    }

    $WindowsPowerShell = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
    if (-not $WindowsPowerShell) {
        throw "Toast notifications require Windows PowerShell 5.1, but powershell.exe was not found."
    }

    $EncodedToastScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ToastScript))
    $EncodedTitle = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Title))
    $EncodedMessage = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Message))
    $EncodedUrl = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Url))
    $EncodedAppId = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($WindowsPowerShellAppId))

    $BootstrapScript = @"
`$ToastScript = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$EncodedToastScript'))
`$Title = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$EncodedTitle'))
`$Message = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$EncodedMessage'))
`$Url = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$EncodedUrl'))
`$AppId = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$EncodedAppId'))
& ([scriptblock]::Create(`$ToastScript)) -Title `$Title -Message `$Message -Url `$Url -AppId `$AppId
"@

    $EncodedBootstrapScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($BootstrapScript))
    $ToastProcess = Start-Process -FilePath $WindowsPowerShell `
        -ArgumentList "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", $EncodedBootstrapScript `
        -Wait -PassThru

    if ($ToastProcess.ExitCode -ne 0) {
        throw "Windows PowerShell toast process failed with exit code $($ToastProcess.ExitCode)."
    }
}

function Install-ScheduledTask {
    param(
        [string]$TaskName,
        [int]$CheckIntervalMinutes,
        [string]$StartTime,
        [switch]$Force
    )

    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if ($ExistingTask) {
        if (-not $Force) {
            throw "Scheduled task '$TaskName' already exists. Re-run with -Force to replace it."
        }
    }

    $TaskCommand = (Get-Command wscript.exe -ErrorAction Stop).Source
    $LauncherPath = Ensure-HiddenLauncher
    $TaskRunCommand = "`"$TaskCommand`" `"$LauncherPath`""
    $SchtasksArgs = @(
        "/Create",
        "/TN", $TaskName,
        "/TR", $TaskRunCommand,
        "/SC", "MINUTE",
        "/MO", $CheckIntervalMinutes,
        "/ST", $StartTime
    )

    if ($Force) {
        $SchtasksArgs += "/F"
    }

    $SchtasksOutput = & schtasks.exe @SchtasksArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "schtasks.exe failed to install '$TaskName': $($SchtasksOutput -join ' ')"
    }

    $TaskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction Stop

    Write-Log "Scheduled task '$TaskName' installed. Next run at $($TaskInfo.NextRunTime), repeating every $CheckIntervalMinutes minutes from $StartTime."
    Write-Host "Scheduled task installed: $TaskName"
    Write-Host "Hidden launcher: $LauncherPath"
    Write-Host "Next run: $($TaskInfo.NextRunTime.ToString('yyyy-MM-dd HH:mm'))"
    Write-Host "Repeat interval: $CheckIntervalMinutes minute(s)"
}

function Get-VersionFromDirName {
    param([string]$Name)
    # Extracts version tuple from directory name like "Ultra-V9.0.46"
    if ($Name -match 'Ultra-V(\d+)\.(\d+)\.(\d+)') {
        return [version]"$($Matches[1]).$($Matches[2]).$($Matches[3])"
    }
    return $null
}

# --- Main logic ---
try {
    if ($InstallScheduledTask) {
        Install-ScheduledTask -TaskName $TaskName -CheckIntervalMinutes $CheckIntervalMinutes -StartTime $StartTime -Force:$Force
        return
    }

    Write-Log "Starting version check..."

    $Headers = @{ "User-Agent" = "GeekMagicClock-VersionChecker/1.0" }
    $Response = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers -ErrorAction Stop

    # Find all version directories and pick the highest
    $VersionDirs = $Response | Where-Object { $_.type -eq "dir" -and $_.name -match "^Ultra-V\d+\.\d+\.\d+$" }

    if (-not $VersionDirs) {
        Write-Log "ERROR: No version directories found in repository."
        exit 1
    }

    $Latest = $VersionDirs |
        Sort-Object { Get-VersionFromDirName $_.name } -Descending |
        Select-Object -First 1

    $LatestVersion = Get-VersionFromDirName $Latest.name
    Write-Log "Latest version on GitHub: $($Latest.name)"

    # Read previously known version
    $KnownVersionStr = if (Test-Path $StateFile) { (Get-Content $StateFile -Raw).Trim() } else { "" }
    $KnownVersion    = if ($KnownVersionStr) { try { [version]$KnownVersionStr } catch { $null } } else { $null }

    if ($null -eq $KnownVersion) {
        # First run — save baseline, no notification
        $LatestVersion.ToString() | Set-Content -Path $StateFile -NoNewline
        Write-Log "First run. Baseline version saved: $LatestVersion"
        Write-Host "Baseline set to $($Latest.name). Future checks will compare against this."
    }
    elseif ($LatestVersion -gt $KnownVersion) {
        Write-Log "NEW VERSION DETECTED: $LatestVersion (was $KnownVersion)"

        # Update state file FIRST so a failed toast doesn't cause repeated re-notifications
        $LatestVersion.ToString() | Set-Content -Path $StateFile -NoNewline

        $Title   = "GeekMagic SmallTV-Ultra Update Available!"
        $Msg     = "New version: $($Latest.name)  (you had $KnownVersionStr)"
        try {
            Show-ToastNotification -Title $Title -Message $Msg -Url $RepoUrl
        }
        catch {
            Write-Log "WARNING: Toast notification failed (no interactive session?): $_"
        }

        Write-Host "New version detected: $($Latest.name)"
    }
    else {
        Write-Log "No update. Current: $LatestVersion"
        Write-Host "Already up to date: $($Latest.name)"
    }
}
catch {
    Write-Log "ERROR: $_"
    Write-Error $_
    exit 1
}
