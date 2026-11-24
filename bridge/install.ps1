[CmdletBinding()]
param(
    [switch]$PrepareOnly,
    [switch]$StartIfConfigured,
    [string]$NodePath = 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
)

$ErrorActionPreference = 'Stop'
if (-not $IsWindows -and $PSVersionTable.PSEdition -eq 'Core') { throw 'Windows is required.' }
if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is unavailable.' }
$bridgeHome = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'ClinicACTPFBridge'))
$localRoot = [IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\') + '\'
if (-not $bridgeHome.StartsWith($localRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid private directory.' }
$userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User
$systemSid = [Security.Principal.SecurityIdentifier]::new('S-1-5-18')

function Test-BridgePrivateAcl {
    param([string]$LiteralPath, [switch]$Directory)
    $item = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Private paths cannot be reparse points.' }
    $acl = Get-Acl -LiteralPath $LiteralPath -ErrorAction Stop
    if ($acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $userSid.Value) { return $false }
    $rules = @($acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
    if ($rules.Count -ne 2) { return $false }
    if ($Directory -and -not $acl.AreAccessRulesProtected) { return $false }
    $seen = @{}
    foreach ($rule in $rules) {
        $sidValue = $rule.IdentityReference.Value
        if ($sidValue -ne $userSid.Value -and $sidValue -ne $systemSid.Value) { return $false }
        if ($seen.ContainsKey($sidValue)) { return $false }
        $seen[$sidValue] = $true
        if ($rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow -or
            $rule.FileSystemRights -ne [Security.AccessControl.FileSystemRights]::FullControl -or
            $rule.PropagationFlags -ne [Security.AccessControl.PropagationFlags]::None) { return $false }
        if ($Directory) {
            if ($rule.IsInherited -or $rule.InheritanceFlags -ne
                ([Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [Security.AccessControl.InheritanceFlags]::ObjectInherit)) { return $false }
        } elseif (-not $acl.AreAccessRulesProtected -and -not $rule.IsInherited) {
            # Unprotected files are acceptable only when both private rules were
            # inherited from the already verified private parent directory.
            return $false
        }
    }
    return $seen.ContainsKey($userSid.Value) -and $seen.ContainsKey($systemSid.Value)
}

$newPrivateDirectory = $false
if (Test-Path -LiteralPath $bridgeHome) {
    if ((Get-Item -LiteralPath $bridgeHome -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Private directory cannot be a reparse point.' }
} else {
    New-Item -ItemType Directory -Path $bridgeHome | Out-Null
    $newPrivateDirectory = $true
}
if ($newPrivateDirectory) {
    # Update only the DACL. Set-Acl with a newly constructed complete security
    # descriptor can request SeSecurityPrivilege even for an owner-run install.
    $icaclsPath = Join-Path $env:SystemRoot 'System32\icacls.exe'
    & $icaclsPath $bridgeHome '/inheritance:r' '/grant:r' ("*{0}:(OI)(CI)F" -f $userSid.Value) '*S-1-5-18:(OI)(CI)F' | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not create the private directory DACL. Nothing was installed or started.' }
}
if (-not (Test-BridgePrivateAcl -LiteralPath $bridgeHome -Directory)) {
    throw 'Private directory must be owned by the current user with protected user/SYSTEM-only full control. Nothing was installed or started.'
}
foreach ($file in @(Get-ChildItem -LiteralPath $bridgeHome -File -Force)) {
    # No file contents or credentials are read. Strict explicit private ACLs or
    # the two inherited rules from this verified parent are already sufficient.
    if (-not (Test-BridgePrivateAcl -LiteralPath $file.FullName)) {
        throw 'A private file has unexpected permissions. Nothing was installed or started.'
    }
}

$configPath = Join-Path $bridgeHome 'config.json'
if ($PrepareOnly) {
    [pscustomobject]@{ Prepared = $true; PrivateDirectory = $bridgeHome; ConfigurationPath = $configPath; Installed = $false; Started = $false }
    return
}
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
    throw 'Private config.json is missing. Only the private directory was prepared; no startup or worker was installed.'
}
try { $configuration = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json -ErrorAction Stop }
catch { throw 'Private config.json is invalid. Configuration content is intentionally not displayed.' }
try {
    $overlayUrl = [uri]$configuration.overlayBase
    $configValid = $overlayUrl.Scheme -eq 'http' -and $overlayUrl.Host -eq '127.0.0.1' -and $overlayUrl.Port -eq 8080 -and $overlayUrl.AbsolutePath -eq '/' -and
        [string]$configuration.scriptUrl -match '^https://script\.google\.com/macros/s/[A-Za-z0-9_-]+/exec$' -and
        $configuration.syncKey -is [string] -and $configuration.syncKey.Length -ge 32
} catch { $configValid = $false }
if (-not $configValid) { throw 'Private configuration failed validation. Nothing was installed or started.' }
$configuration = $null
if (-not [IO.Path]::IsPathRooted($NodePath)) { throw 'NodePath must be an absolute node.exe path.' }
$NodePath = [IO.Path]::GetFullPath($NodePath)
$nodeFile = Get-Item -LiteralPath $NodePath -ErrorAction Stop
if ($nodeFile.PSIsContainer -or $nodeFile.Name -ne 'node.exe') { throw 'NodePath must identify node.exe.' }
$nodeFullPath = $nodeFile.FullName

$sourcePath = Join-Path $PSScriptRoot 'pf_bridge.cjs'
$installedPath = Join-Path $bridgeHome 'pf_bridge.cjs'
$pidPath = Join-Path $bridgeHome 'worker.pid'
if (Test-Path -LiteralPath $pidPath) {
    $pidText = (Get-Content -LiteralPath $pidPath -Raw).Trim()
    if ($pidText -match '^\d+$') {
        $existing = Get-CimInstance Win32_Process -Filter ('ProcessId={0}' -f [int]$pidText) -ErrorAction SilentlyContinue
        if ($existing -and $existing.CommandLine -like ('*' + $installedPath + '*')) { throw 'The bridge is running. Run stop.ps1 before updating its installed copy.' }
    }
}
Copy-Item -LiteralPath $sourcePath -Destination $installedPath -Force
$launcherPath = Join-Path $bridgeHome 'start-hidden.vbs'
$stopMarker = Join-Path $bridgeHome 'stop.requested'
$escapedNode = $nodeFullPath.Replace('"', '""')
$escapedScript = $installedPath.Replace('"', '""')
$escapedHome = $bridgeHome.Replace('"', '""')
$launcher = @"
Option Explicit
Dim shell, files, code, marker, report
Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
shell.CurrentDirectory = "$escapedHome"
shell.Environment("PROCESS")("ACT_PF_BRIDGE_HOME") = "$escapedHome"
marker = "$escapedHome\stop.requested"
If files.FileExists(marker) Then files.DeleteFile marker, True
code = shell.Run(Chr(34) & "$escapedNode" & Chr(34) & " " & Chr(34) & "$escapedScript" & Chr(34), 0, True)
If code <> 0 And Not files.FileExists(marker) Then
  Set report = files.CreateTextFile("$escapedHome\startup-error.txt", True)
  report.WriteLine "ACT PF bridge stopped; check configuration and health.json. No successful synchronization is claimed."
  report.Close
  shell.Popup "ACT PF bridge could not stay running. Check health.json or startup-error.txt in ClinicACTPFBridge.", 15, "ACT PF Bridge", 48
End If
"@
Set-Content -LiteralPath $launcherPath -Value $launcher -Encoding Unicode
$startupPath = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupPath 'Clinic ACT PF Bridge.lnk'
$wscriptPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
if ((Test-Path -LiteralPath $shortcutPath) -and
    ($shortcut.TargetPath -ne $wscriptPath -or $shortcut.Arguments -ne ('//B //Nologo "' + $launcherPath + '"'))) {
    throw 'The startup shortcut name is already used by a different target. It was not overwritten.'
}
$shortcut.TargetPath = $wscriptPath
$shortcut.Arguments = '//B //Nologo "' + $launcherPath + '"'
$shortcut.WorkingDirectory = $bridgeHome
$shortcut.WindowStyle = 7
$shortcut.Description = 'Clinic ACT PF bridge; independent worker, no Overlay restart'
$manifest = @{ schemaVersion = 1; nodePath = $nodeFullPath; scriptPath = $installedPath; launcherPath = $launcherPath; shortcutPath = $shortcutPath; installedAt = [DateTimeOffset]::UtcNow.ToString('o') }
$manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $bridgeHome 'install.json') -Encoding utf8
$shortcut.Save()

$healthy = $false
$processVerified = $false
if ($StartIfConfigured) {
    if (Test-Path -LiteralPath $stopMarker) { Remove-Item -LiteralPath $stopMarker }
    $startedAt = [DateTimeOffset]::UtcNow
    Start-Process -FilePath $wscriptPath -ArgumentList @('//B', '//Nologo', ('"{0}"' -f $launcherPath)) -WindowStyle Hidden | Out-Null
    $healthPath = Join-Path $bridgeHome 'health.json'
    $quotedScriptPattern = '(?i)(?:^|\s)"' + [regex]::Escape($installedPath) + '"(?:\s|$)'
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        Start-Sleep -Milliseconds 500
        $pidText = if (Test-Path -LiteralPath $pidPath) { (Get-Content -LiteralPath $pidPath -Raw).Trim() } else { '' }
        if ($pidText -match '^\d+$' -and [long]$pidText -le [int]::MaxValue) {
            $worker = Get-CimInstance Win32_Process -Filter ('ProcessId={0}' -f [int]$pidText) -ErrorAction SilentlyContinue
            if ($worker -and $worker.ExecutablePath -eq $nodeFullPath -and $worker.CommandLine -match $quotedScriptPattern) {
                $processVerified = $true; break
            }
        }
    }
    if (-not $processVerified) { throw 'Bridge worker identity could not be verified. Inspect startup-error.txt; no clinical synchronization is claimed.' }
    if (Test-Path -LiteralPath $healthPath) {
        try {
            $health = Get-Content -LiteralPath $healthPath -Raw | ConvertFrom-Json
            if ([DateTimeOffset]::Parse($health.startedAt) -ge $startedAt.AddSeconds(-1)) {
                if ($health.errorCode) {
                    $safeCode = if ([string]$health.errorCode -match '^[A-Z0-9_]+$') { [string]$health.errorCode } else { 'BRIDGE_CYCLE_FAILED' }
                    throw "Bridge health check failed: $safeCode"
                }
                $healthy = $true
            }
        } catch {
            if ($_.Exception.Message -like 'Bridge health check failed:*') { throw }
        }
    }
}
[pscustomobject]@{ Installed = $true; Started = [bool]$StartIfConfigured; ProcessVerified = $processVerified;
    HealthPending = [bool]($StartIfConfigured -and -not $healthy); HealthyCycleVerified = $healthy;
    PrivateDirectory = $bridgeHome; StartupShortcut = $shortcutPath; WorkerSha256 = (Get-FileHash -LiteralPath $installedPath -Algorithm SHA256).Hash }
