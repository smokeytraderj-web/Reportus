[CmdletBinding()]
param(
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$BuildVenv = Join-Path $ProjectRoot ".build-venv"
$BuildPython = Join-Path $BuildVenv "Scripts\python.exe"

if ($env:OS -ne "Windows_NT") {
    throw "Reporticles must be packaged on Windows."
}

Set-Location $ProjectRoot

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$CommandArguments = @()
    )
    & $FilePath @CommandArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE: $FilePath"
    }
}

if (-not (Test-Path $BuildPython)) {
    $PyLauncher = Get-Command "py" -ErrorAction SilentlyContinue
    if ($PyLauncher) {
        Invoke-Checked -FilePath $PyLauncher.Source -CommandArguments @("-3", "-m", "venv", $BuildVenv)
    } else {
        $Python = (Get-Command "python" -ErrorAction Stop).Source
        Invoke-Checked -FilePath $Python -CommandArguments @("-m", "venv", $BuildVenv)
    }
}

Invoke-Checked -FilePath $BuildPython -CommandArguments @("-m", "pip", "install", "--upgrade", "pip")
Invoke-Checked -FilePath $BuildPython -CommandArguments @("-m", "pip", "install", "-r", "requirements.txt", "-r", "requirements-build.txt")
Invoke-Checked -FilePath $BuildPython -CommandArguments @("-m", "unittest", "discover", "-s", "tests", "-v")
$env:PLAYWRIGHT_BROWSERS_PATH = "0"
Invoke-Checked -FilePath $BuildPython -CommandArguments @("-m", "playwright", "install", "chromium")
Invoke-Checked -FilePath $BuildPython -CommandArguments @("-m", "PyInstaller", "--noconfirm", "--clean", "packaging\windows\Reporticles.spec")

$PackagedApp = Join-Path $ProjectRoot "dist\Reporticles\Reporticles.exe"
if (-not (Test-Path $PackagedApp)) {
    throw "PyInstaller did not create Reporticles.exe."
}
Invoke-Checked -FilePath $PackagedApp -CommandArguments @("--smoke-test")

if (-not $SkipInstaller) {
    $CompilerCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
    ) | Where-Object { $_ -and (Test-Path $_) }
    if (-not $CompilerCandidates) {
        throw "Inno Setup 6 is required. Install it, then run this script again."
    }
    Invoke-Checked -FilePath $CompilerCandidates[0] -CommandArguments @("packaging\windows\Reporticles.iss")
    $Installer = Get-ChildItem "dist\installer\Reporticles-Setup-*.exe" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $Installer) {
        throw "Inno Setup did not create an installer."
    }
    $Hash = Get-FileHash $Installer.FullName -Algorithm SHA256
    $HashLine = "$($Hash.Hash.ToLowerInvariant())  $($Installer.Name)"
    $HashPath = "$($Installer.FullName).sha256"
    Set-Content -Path $HashPath -Value $HashLine -Encoding ascii
    Write-Host $HashLine
}
