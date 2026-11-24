[CmdletBinding()]
param([switch]$RemoveStartup)

$ErrorActionPreference = 'Stop'
if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is unavailable.' }
$bridgeHome = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'ClinicACTPFBridge'))
$localRoot = [IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\') + '\'
if (-not $bridgeHome.StartsWith($localRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid private directory.' }
$manifestPath = Join-Path $bridgeHome 'install.json'
$installedPath = Join-Path $bridgeHome 'pf_bridge.cjs'
$pidPath = Join-Path $bridgeHome 'worker.pid'
$expectedShortcut = Join-Path ([Environment]::GetFolderPath('Startup')) 'Clinic ACT PF Bridge.lnk'
$stopped = $false

if (Test-Path -LiteralPath $pidPath) {
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw 'Cannot verify bridge ownership without install.json. No process was stopped.' }
    try { $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json }
    catch { throw 'Installation manifest is invalid. No process was stopped.' }
    if ([IO.Path]::GetFullPath([string]$manifest.scriptPath) -ne $installedPath) { throw 'Unexpected bridge script path. No process was stopped.' }
    $pidText = (Get-Content -LiteralPath $pidPath -Raw).Trim()
    if ($pidText -notmatch '^\d+$' -or [long]$pidText -gt [int]::MaxValue) { throw 'Invalid worker PID. No process was stopped.' }
    $workerId = [int]$pidText
    $worker = Get-CimInstance Win32_Process -Filter ('ProcessId={0}' -f $workerId) -ErrorAction SilentlyContinue
    if ($worker) {
        $quotedScriptPattern = '(?i)(?:^|\s)"' + [regex]::Escape($installedPath) + '"(?:\s|$)'
        if (-not $worker.ExecutablePath -or
            [IO.Path]::GetFullPath([string]$worker.ExecutablePath) -ne [IO.Path]::GetFullPath([string]$manifest.nodePath) -or
            $worker.CommandLine -notmatch $quotedScriptPattern) {
            throw 'PID belongs to a different process or cannot be verified. No process was stopped.'
        }
        Set-Content -LiteralPath (Join-Path $bridgeHome 'stop.requested') -Value 'User requested bridge stop.' -Encoding utf8
        Stop-Process -Id $workerId -ErrorAction Stop
        $stopped = $true
    }
    Remove-Item -LiteralPath $pidPath -ErrorAction SilentlyContinue
}

if ($RemoveStartup -and (Test-Path -LiteralPath $expectedShortcut)) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($expectedShortcut)
    $expectedLauncher = Join-Path $bridgeHome 'start-hidden.vbs'
    $expectedWscript = Join-Path $env:SystemRoot 'System32\wscript.exe'
    if ($shortcut.TargetPath -ne $expectedWscript -or $shortcut.Arguments -ne ('//B //Nologo "' + $expectedLauncher + '"')) {
        throw 'Startup shortcut belongs to another target. It was not removed.'
    }
    Remove-Item -LiteralPath $expectedShortcut
}
[pscustomobject]@{ Stopped = $stopped; StartupRemoved = [bool]$RemoveStartup; ConfigurationPreserved = $true; StatePreserved = $true }
