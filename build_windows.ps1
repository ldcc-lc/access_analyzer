$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$BuildEnvironment = Join-Path $ProjectRoot ".venv-build-windows"
$BuildPython = Join-Path $BuildEnvironment "Scripts\python.exe"

if (-not (Test-Path $BuildPython)) {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -3 -m venv $BuildEnvironment
    } else {
        python -m venv $BuildEnvironment
    }
    if ($LASTEXITCODE -ne 0) {
        throw "No se pudo crear el entorno virtual. Instala Python 3 para Windows y vuelve a intentarlo."
    }
}

& $BuildPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "No se pudo actualizar pip." }

& $BuildPython -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw "No se pudieron instalar las dependencias de construcción." }

New-Item -ItemType Directory -Force -Path (Join-Path $ProjectRoot "build") | Out-Null
$Arguments = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", "AnalizadorDeAccesos",
    "--icon", (Join-Path $ProjectRoot "assets\icono.ico"),
    "--workpath", (Join-Path $ProjectRoot "build\work"),
    "--distpath", (Join-Path $ProjectRoot "dist"),
    "--specpath", (Join-Path $ProjectRoot "build"),
    "--add-data", "$(Join-Path $ProjectRoot 'documentos base\plantilla informe.xlsx');documentos base",
    "--add-data", "$(Join-Path $ProjectRoot 'assets\logo_bdp.jpeg');assets",
    "--add-data", "$(Join-Path $ProjectRoot 'assets\icono.ico');assets",
    "interfaz.py"
)
& $BuildPython @Arguments
if ($LASTEXITCODE -ne 0) { throw "PyInstaller no pudo generar el ejecutable." }

Write-Host "Ejecutable generado: $(Join-Path $ProjectRoot 'dist\AnalizadorDeAccesos.exe')"
