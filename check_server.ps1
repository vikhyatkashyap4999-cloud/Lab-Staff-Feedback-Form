<#
  check_server.ps1  -  READ-ONLY check of the Windows VM for the Lab Feedback Form.
  It changes nothing. It only reports what is installed and what is missing.

  HOW TO RUN (on the VM):
    1. Open PowerShell as Administrator
    2. cd to the app folder (for example: cd E:\labfeedback)
    3. powershell -ExecutionPolicy Bypass -File .\check_server.ps1

  Optional:  -SiteFolder "D:\sites\labfeedback"  -AppPool "LabFeedback"  -TestUrl "https://labfeedback.upes.ac.in/login"
#>
param(
    [string]$SiteFolder = $PSScriptRoot,
    [string]$AppPool = "LabFeedback",
    [string]$TestUrl = ""
)

$script:fails = 0
function Show-Ok($msg)   { Write-Host "  [ OK ]  $msg" -ForegroundColor Green }
function Show-Bad($msg, $fix) {
    $script:fails++
    Write-Host "  [FAIL]  $msg" -ForegroundColor Red
    if ($fix) { Write-Host "          Fix: $fix" -ForegroundColor Yellow }
}
function Show-Warn($msg) { Write-Host "  [WARN]  $msg" -ForegroundColor Yellow }
function Show-Head($msg) { Write-Host ""; Write-Host "== $msg ==" -ForegroundColor Cyan }

if (-not $SiteFolder) { $SiteFolder = "E:\labfeedback" }
$appcmd = Join-Path $env:windir "System32\inetsrv\appcmd.exe"

Write-Host "Lab Feedback Form - server check (read-only)"
Write-Host "Site folder: $SiteFolder"

# ------------------------------------------------------------------ admin
Show-Head "PowerShell"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) { Show-Ok "Running as Administrator" }
else { Show-Warn "Not running as Administrator - some IIS checks may fail. Right-click PowerShell > Run as administrator." }

# ------------------------------------------------------------------ IIS
Show-Head "IIS (web server)"
$w3 = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
if ($w3) {
    if ($w3.Status -eq "Running") { Show-Ok "IIS is installed and running" }
    else { Show-Bad "IIS is installed but not running (status: $($w3.Status))" "Run: iisreset /start" }
} else {
    Show-Bad "IIS is not installed" "Server Manager > Add Roles and Features > Web Server (IIS)"
}
$hasAppcmd = Test-Path $appcmd

# ------------------------------------------------------------------ HttpPlatformHandler
Show-Head "HttpPlatformHandler (lets IIS run Python)"
$dll = Join-Path $env:windir "System32\inetsrv\httpPlatformHandler.dll"
$hostCfg = Join-Path $env:windir "System32\inetsrv\config\applicationHost.config"
$registered = $null   # $null = could not read IIS config (needs Administrator)
try {
    if (Test-Path $hostCfg -ErrorAction Stop) {
        $registered = [bool](Select-String -Path $hostCfg -Pattern "httpPlatformHandler" -SimpleMatch -Quiet -ErrorAction Stop)
    }
} catch { $registered = $null }
$hasDll = $false
try { $hasDll = [bool](Test-Path $dll -ErrorAction Stop) } catch {}
if ($hasDll -and $registered -eq $true) { Show-Ok "HttpPlatformHandler is installed and registered in IIS" }
elseif ($hasDll -and $null -eq $registered) { Show-Warn "HttpPlatformHandler file is present, but IIS settings could not be read (run as Administrator to confirm)" }
elseif ($hasDll) { Show-Bad "HttpPlatformHandler file exists but is not registered in IIS" "Reinstall HttpPlatformHandler v1.2 (x64), then run: iisreset" }
else { Show-Bad "HttpPlatformHandler is NOT installed" "Install v1.2 (x64) from https://www.iis.net/downloads/microsoft/httpplatformhandler then run: iisreset (at a quiet time - it briefly restarts all sites)" }

# ------------------------------------------------------------------ Python
Show-Head "Python"
$pyExe = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $ver = ((& py --version 2>&1) | Out-String).Trim()
    $pyExe = ((& py -c "import sys; print(sys.base_prefix)" 2>&1) | Out-String).Trim()
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $ver = ((& python --version 2>&1) | Out-String).Trim()
    $pyExe = ((& python -c "import sys; print(sys.base_prefix)" 2>&1) | Out-String).Trim()
}
if (-not $pyExe) {
    Show-Bad "Python is not installed" "Install Python 3.12 or newer (64-bit) for ALL users from python.org"
} elseif ($ver -match "Python 3\.(\d+)" -and [int]$Matches[1] -ge 10) {
    Show-Ok "$ver is installed at $pyExe"
    if ($pyExe -match "\\Users\\") {
        Show-Bad "Python is installed for ONE user only ($pyExe). IIS cannot use it." "Reinstall Python: Customize installation > tick 'Install Python for all users' (goes to C:\Program Files\Python3xx)"
    }
} else {
    Show-Bad "Python version too old: $ver" "Install Python 3.12 or newer (64-bit) for all users"
}

# ------------------------------------------------------------------ app folder
Show-Head "App folder"
if (-not (Test-Path $SiteFolder)) {
    Show-Bad "Folder not found: $SiteFolder" "Copy the app folder there, or pass -SiteFolder <path>"
} else {
    foreach ($f in @("app.py", "icons.py", "workbook_utils.py", "requirements.txt", "web.config", "static", "templates")) {
        if (Test-Path (Join-Path $SiteFolder $f)) { Show-Ok "$f present" }
        else { Show-Bad "$f is missing" "Copy it from the project folder" }
    }
    if (Test-Path (Join-Path $SiteFolder "logs")) { Show-Ok "logs folder present" }
    else { Show-Bad "logs folder is missing" "mkdir `"$SiteFolder\logs`"" }

    # web.config paths
    $wc = Join-Path $SiteFolder "web.config"
    if (Test-Path $wc) {
        try {
            [xml]$cfg = Get-Content $wc -Raw
            $hp = $cfg.configuration.'system.webServer'.httpPlatform
            $procPath = $hp.processPath
            $logPath = $hp.stdoutLogFile
            if (Test-Path $procPath) { Show-Ok "web.config Python path exists: $procPath" }
            else { Show-Bad "web.config points to a Python that does not exist: $procPath" "Do Step 4 of the guide (create venv), or fix processPath in web.config" }
            $logDir = Split-Path $logPath -Parent
            if (Test-Path $logDir) { Show-Ok "web.config log folder exists: $logDir" }
            else { Show-Bad "web.config log folder does not exist: $logDir" "Create it, or fix stdoutLogFile in web.config" }
        } catch { Show-Bad "web.config could not be read as XML: $_" "Re-copy web.config from the project" }
    }

    # venv packages
    $venvPy = Join-Path $SiteFolder "venv\Scripts\python.exe"
    if (Test-Path $venvPy) {
        $out = (& $venvPy -c "import flask, openpyxl, waitress; print('ok')" 2>&1) | Out-String
        if ($out -match "ok") { Show-Ok "venv has flask, openpyxl and waitress" }
        else { Show-Bad "venv is missing packages" "Run: venv\Scripts\pip install -r requirements.txt" }
        $imp = (& $venvPy -c "import sys; sys.path.insert(0, r'$SiteFolder'); import app; print('ok')" 2>&1) | Out-String
        if ($imp -match "ok") { Show-Ok "app.py loads without errors" }
        else { Show-Bad "app.py failed to load: $($imp.Trim())" "Read the error above" }
    } else {
        Show-Bad "venv not found at $venvPy" "Run: py -m venv venv  then  venv\Scripts\pip install -r requirements.txt"
    }

    # admin password set?
    $db = Join-Path $SiteFolder "instance\feedback.db"
    if (Test-Path $db) { Show-Ok "Database exists (instance\feedback.db)" }
    else { Show-Warn "No database yet - it is created on first start. Run 'venv\Scripts\python app.py reset-admin' to set the admin password." }

    # data files that should not be on the server
    foreach ($bad in @("data received", "Lab_Feedback_Upload_Template.xlsx")) {
        if (Test-Path (Join-Path $SiteFolder $bad)) { Show-Warn "'$bad' is in the site folder - it holds real data and is not needed. Remove it." }
    }
}

# ------------------------------------------------------------------ IIS site + pool
Show-Head "IIS site and application pool"
if (-not $hasAppcmd) {
    Show-Bad "Cannot check IIS site (appcmd.exe not found)" "Install IIS first"
} else {
    # application pool
    $poolXml = (& $appcmd list apppool "$AppPool" /config:* 2>&1) | Out-String
    if ($poolXml -match "<add ") {
        Show-Ok "Application pool '$AppPool' exists"
        try {
            [xml]$px = $poolXml
            $add = $px.add
            if ([string]::IsNullOrEmpty($add.managedRuntimeVersion)) { Show-Ok ".NET CLR version = No Managed Code" }
            else { Show-Warn ".NET CLR version is '$($add.managedRuntimeVersion)' - 'No Managed Code' is recommended" }
            $maxp = $add.processModel.maxProcesses
            if ($maxp -eq "1") { Show-Ok "Maximum Worker Processes = 1" }
            else { Show-Bad "Maximum Worker Processes = $maxp" "Set it to 1 (Application Pools > $AppPool > Advanced Settings)" }
            $idle = $add.processModel.idleTimeout
            if ($idle -eq "00:00:00") { Show-Ok "Idle Time-out = 0 (site will not sleep)" }
            else { Show-Warn "Idle Time-out = $idle - set it to 0 so the first visitor after a quiet period does not wait" }
        } catch { Show-Warn "Could not read pool settings: $_" }
    } else {
        Show-Bad "Application pool '$AppPool' not found" "Do Step 6 of the guide, or pass -AppPool <name>"
    }

    # site / application pointing at the folder
    $vdirs = (& $appcmd list vdir /xml 2>&1) | Out-String
    $found = $null
    try {
        [xml]$vx = $vdirs
        $target = [IO.Path]::GetFullPath($SiteFolder).TrimEnd('\')
        foreach ($v in $vx.appcmd.VDIR) {
            $pp = [Environment]::ExpandEnvironmentVariables($v.physicalPath).TrimEnd('\')
            if ($pp -ieq $target) { $found = $v; break }
        }
    } catch {}
    if ($found) {
        Show-Ok "IIS site/app points to this folder: $($found.'APP.NAME')"
        $apps = (& $appcmd list app "$($found.'APP.NAME')" /xml 2>&1) | Out-String
        try {
            [xml]$ax = $apps
            $usedPool = $ax.appcmd.APP.'APPPOOL.NAME'
            if ($usedPool -eq $AppPool) { Show-Ok "It uses the '$AppPool' application pool" }
            else { Show-Bad "It uses pool '$usedPool', not '$AppPool'" "Change the site's application pool to $AppPool" }
            $siteName = $ax.appcmd.APP.'SITE.NAME'
            $sx = (& $appcmd list site "$siteName" /xml 2>&1) | Out-String
            [xml]$sxx = $sx
            $b = $sxx.appcmd.SITE.bindings
            Write-Host "          Bindings: $b"
            if ($b -match "https/") { Show-Ok "Site has an https binding" }
            else { Show-Warn "No https binding - add one with the SSL certificate (or set LAB_HTTPS=0 in web.config for an http-only test)" }
        } catch { Show-Warn "Could not read site details: $_" }
    } else {
        Show-Bad "No IIS site or application points to $SiteFolder" "Do Step 6 of the guide (Add Website, physical path = this folder)"
    }
}

# ------------------------------------------------------------------ permissions
Show-Head "Folder permissions"
foreach ($sub in @("instance", "logs")) {
    $p = Join-Path $SiteFolder $sub
    if (Test-Path $p) {
        $acl = (& icacls $p 2>&1) | Out-String
        if ($acl -match ([regex]::Escape("IIS AppPool\$AppPool") + ":.*\((M|F)\)")) {
            Show-Ok "$sub is writable by IIS AppPool\$AppPool"
        } else {
            Show-Bad "$sub is not writable by IIS AppPool\$AppPool" "icacls `"$p`" /grant `"IIS AppPool\$($AppPool)`":(OI)(CI)M"
        }
    } elseif ($sub -eq "instance") {
        Show-Warn "instance folder does not exist yet - create it and grant write access (Step 7)"
    }
}

# ------------------------------------------------------------------ live test
if ($TestUrl) {
    Show-Head "Live test: $TestUrl"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $r = Invoke-WebRequest -Uri $TestUrl -UseBasicParsing -TimeoutSec 60
        if ($r.StatusCode -eq 200 -and $r.Content -match "captcha|Login|login") { Show-Ok "Site answered with the login page (HTTP 200)" }
        else { Show-Warn "Site answered HTTP $($r.StatusCode) but it did not look like the login page" }
    } catch {
        Show-Bad "Site did not answer: $($_.Exception.Message)" "Read $SiteFolder\logs\python.log"
    }
}

# ------------------------------------------------------------------ summary
Write-Host ""
if ($script:fails -eq 0) { Write-Host "RESULT: everything checked is OK." -ForegroundColor Green }
else { Write-Host "RESULT: $script:fails problem(s) found. Fix the [FAIL] lines above, then run this again." -ForegroundColor Red }
