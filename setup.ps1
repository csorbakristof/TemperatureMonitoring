# TemperatureMonitoring Application Environment Setup Script
param([switch]$Help, [switch]$SkipTests)

if ($Help) {
    Write-Host "Temperature Monitoring Environment Setup Script" -ForegroundColor Cyan
    Write-Host "USAGE: .\setup.ps1 [-SkipTests] [-Help]" -ForegroundColor Yellow
    exit 0
}

Write-Host "Temperature Monitoring Application Setup" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# Check for Python
$pythonExe = $null
$pythonCommands = @("python", "python3", "py")
foreach ($cmd in $pythonCommands) {
    try {
        $version = & $cmd --version 2>&1
        if ($LASTEXITCODE -eq 0 -and $version -match "Python (\d+)\.(\d+)") {
            $major = [int]$matches[1]
            $minor = [int]$matches[2]
            if ($major -ge 3 -and ($major -gt 3 -or $minor -ge 8)) {
                $pythonExe = $cmd
                Write-Host "Found Python: $version" -ForegroundColor Green
                break
            }
        }
    } catch { continue }
}

if (-not $pythonExe) {
    Write-Host "ERROR: Python 3.8+ not found!" -ForegroundColor Red
    exit 1
}

# Create virtual environment if not exists
if (-not (Test-Path "venv") -and -not $env:VIRTUAL_ENV) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    & $pythonExe -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to create virtual environment!" -ForegroundColor Red
        exit 1
    }
}

# Activate if not already active
if (-not $env:VIRTUAL_ENV -and (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "Activating virtual environment..." -ForegroundColor Yellow
    & "venv\Scripts\Activate.ps1"
}

# Update pip
Write-Host "Updating pip..." -ForegroundColor Yellow
& $pythonExe -m pip install --upgrade pip

# Install requirements
if (Test-Path "requirements.txt") {
    Write-Host "Installing from requirements.txt..." -ForegroundColor Yellow
    & $pythonExe -m pip install -r requirements.txt
} else {
    Write-Host "Installing basic packages..." -ForegroundColor Yellow
    $packages = @("pandas", "matplotlib", "seaborn", "openpyxl", "pytest", "numpy")
    foreach ($pkg in $packages) {
        & $pythonExe -m pip install $pkg
    }
}

# Verify installation
if (-not $SkipTests) {
    Write-Host "Verifying packages..." -ForegroundColor Yellow
    $packages = @("pandas", "matplotlib", "seaborn", "openpyxl", "pytest", "numpy")
    foreach ($pkg in $packages) {
        & $pythonExe -c "import $pkg; print('OK: $pkg')" 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ $pkg" -ForegroundColor Green
        } else {
            Write-Host "✗ $pkg failed" -ForegroundColor Red
        }
    }
}

# Create directories
@("data", "output") | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Host "Created: $_" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "SETUP COMPLETED!" -ForegroundColor Green
Write-Host "Run: $pythonExe src\main.py data\sample_data.zip" -ForegroundColor Cyan
