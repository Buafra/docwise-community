# DocWise Community - first-run setup
# Installs Python 3.13 and Tesseract OCR if missing, creates the .venv,
# and installs app dependencies. Safe to re-run anytime: it skips
# anything that is already installed. Normally launched by start.bat.

# NOTE: keep 'Continue'. Under Windows PowerShell 5.1 with 'Stop', any line a
# child process writes to stderr (pip/winget/ensurepip warnings) becomes a
# terminating error when stderr is redirected. Failures are detected via
# $LASTEXITCODE checks instead.
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$PythonWingetId = 'Python.Python.3.13'
$PythonUrl      = 'https://www.python.org/ftp/python/3.13.15/python-3.13.15-amd64.exe'
$TessWingetId   = 'UB-Mannheim.TesseractOCR'
$TessUrl        = 'https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe'

function Refresh-Path {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Get-WingetExe {
    $w = Get-Command winget -ErrorAction SilentlyContinue
    if ($w) { return $w.Source }
    return $null
}

function Download-File($Url, $Dest) {
    foreach ($try in 1..3) {
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing -TimeoutSec 600 -ErrorAction Stop
            return
        } catch {
            if ($try -eq 3) { throw }
            Start-Sleep -Seconds 3
        }
    }
}

# Returns @{Exe=...; Pre=@(...)} for a working Python 3.10+, or $null.
function Find-Python {
    $candidates = @()
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { $candidates += , @{ Exe = $py.Source; Pre = @('-3') } }
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { $candidates += , @{ Exe = $python.Source; Pre = @() } }
    $globs = @("$env:LocalAppData\Programs\Python\Python3*\python.exe",
               "$env:ProgramFiles\Python3*\python.exe")
    foreach ($g in $globs) {
        $found = Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($f in $found) { $candidates += , @{ Exe = $f.FullName; Pre = @() } }
    }
    foreach ($c in $candidates) {
        try {
            $out = (& $c.Exe ($c.Pre + '--version') 2>&1 | Out-String).Trim()
            if ($LASTEXITCODE -eq 0 -and $out -match 'Python (\d+)\.(\d+)') {
                $maj = [int]$Matches[1]; $min = [int]$Matches[2]
                if (($maj -eq 3 -and $min -ge 10) -or $maj -gt 3) {
                    $c.Version = $out
                    return $c
                }
            }
        } catch {}
    }
    return $null
}

function Find-Tesseract {
    if ($env:TESSERACT_CMD -and (Test-Path $env:TESSERACT_CMD)) { return $env:TESSERACT_CMD }
    $t = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($t) { return $t.Source }
    $paths = @("$env:ProgramFiles\Tesseract-OCR\tesseract.exe",
               "${env:ProgramFiles(x86)}\Tesseract-OCR\tesseract.exe",
               "$env:LocalAppData\Programs\Tesseract-OCR\tesseract.exe")
    foreach ($p in $paths) { if ($p -and (Test-Path $p)) { return $p } }
    return $null
}

try {
    # ---------- 1. Python ----------
    Write-Host ""
    Write-Host "[1/3] Checking Python..."
    $pyCmd = Find-Python
    if ($pyCmd) {
        Write-Host "  OK: $($pyCmd.Version) at $($pyCmd.Exe)"
    } else {
        Write-Host "  Python 3.10+ not found. Installing Python 3.13 (no admin needed)..."
        $installed = $false
        $winget = Get-WingetExe
        if ($winget) {
            try {
                & $winget install -e --id $PythonWingetId --silent --accept-package-agreements --accept-source-agreements
                if ($LASTEXITCODE -eq 0) { $installed = $true }
            } catch {}
        }
        if (-not $installed) {
            Write-Host "  Downloading Python from python.org (28 MB)..."
            $tmp = Join-Path $env:TEMP 'docwise-python-setup.exe'
            Download-File $PythonUrl $tmp
            Write-Host "  Running Python installer..."
            $p = Start-Process -Wait -PassThru -FilePath $tmp -ArgumentList '/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_launcher=1'
            if ($p.ExitCode -ne 0) { throw "Python installer failed with code $($p.ExitCode)" }
        }
        Refresh-Path
        $pyCmd = Find-Python
        if (-not $pyCmd) {
            throw "Python was installed but could not be located. Close this window and run start.bat again."
        }
        Write-Host "  OK: installed $($pyCmd.Version)"
    }

    # ---------- 2. Virtual environment + dependencies ----------
    Write-Host ""
    Write-Host "[2/3] Checking app dependencies..."
    $venvPy = Join-Path $Root '.venv\Scripts\python.exe'
    if (-not (Test-Path $venvPy)) {
        Write-Host "  Creating virtual environment..."
        $venvOut = & $pyCmd.Exe ($pyCmd.Pre + @('-m', 'venv', '.venv')) 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $venvPy)) {
            Write-Host $venvOut
            throw "Could not create the virtual environment."
        }
    }
    & $venvPy -c 'import fastapi, uvicorn, onnxruntime' 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Installing dependencies (a few minutes on first run)..."
        & $venvPy -m pip install --upgrade pip 2>&1 | ForEach-Object { "$_" }
        & $venvPy -m pip install -r requirements.txt 2>&1 | ForEach-Object { "$_" }
        if ($LASTEXITCODE -ne 0) { throw "pip install failed. Check your internet connection and run start.bat again." }
    }
    Write-Host "  OK: dependencies ready"

    # PyMuPDF (the PDF engine) ships a native DLL that needs the Microsoft
    # Visual C++ runtime. Fresh Windows machines often lack it, and PDFs then
    # silently extract nothing while the rest of the app works.
    & $venvPy -c 'import fitz' 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  PDF engine cannot load - installing Microsoft Visual C++ runtime..."
        Write-Host "  NOTE: Windows may show an admin (UAC) prompt - click Yes to allow it."
        $vcOk = $false
        $winget = Get-WingetExe
        if ($winget) {
            try {
                & $winget install -e --id 'Microsoft.VCRedist.2015+.x64' --silent --accept-package-agreements --accept-source-agreements
                if ($LASTEXITCODE -eq 0) { $vcOk = $true }
            } catch {}
        }
        if (-not $vcOk) {
            try {
                $tmp = Join-Path $env:TEMP 'docwise-vcredist.exe'
                Download-File 'https://aka.ms/vs/17/release/vc_redist.x64.exe' $tmp
                Start-Process -Wait -FilePath $tmp -ArgumentList '/install', '/quiet', '/norestart'
            } catch {}
        }
        & $venvPy -c 'import fitz' 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "The PDF engine still cannot load. PDF files will not be processed until the Microsoft Visual C++ runtime is installed: https://aka.ms/vs/17/release/vc_redist.x64.exe"
        } else {
            Write-Host "  OK: PDF engine ready"
        }
    }

    # ---------- 3. Tesseract OCR (optional but recommended) ----------
    Write-Host ""
    Write-Host "[3/3] Checking Tesseract OCR (for scanned documents and images)..."
    $tess = Find-Tesseract
    if (-not $tess) {
        Write-Host "  Tesseract not found. Installing it now."
        Write-Host "  NOTE: Windows may show an admin (UAC) prompt - click Yes to allow it."
        $winget = Get-WingetExe
        $ok = $false
        if ($winget) {
            try {
                & $winget install -e --id $TessWingetId --silent --accept-package-agreements --accept-source-agreements
                if ($LASTEXITCODE -eq 0) { $ok = $true }
            } catch {}
        }
        if (-not $ok) {
            try {
                Write-Host "  Downloading Tesseract installer (48 MB)..."
                $tmp = Join-Path $env:TEMP 'docwise-tesseract-setup.exe'
                Download-File $TessUrl $tmp
                Write-Host "  Running Tesseract installer..."
                Start-Process -Wait -FilePath $tmp -ArgumentList '/S'
            } catch {}
        }
        Refresh-Path
        $tess = Find-Tesseract
    }
    if ($tess) {
        Write-Host "  OK: Tesseract at $tess"
    } else {
        Write-Host ""
        Write-Warning "Tesseract could not be installed. The app still works, but OCR of scanned PDFs and images is disabled."
        Write-Host "  To add OCR later: install https://github.com/UB-Mannheim/tesseract/wiki then restart the app,"
        Write-Host "  or run setup.ps1 again."
    }

    Write-Host ""
    Write-Host "Setup complete."
    exit 0
} catch {
    Write-Host ""
    Write-Host "SETUP FAILED: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
