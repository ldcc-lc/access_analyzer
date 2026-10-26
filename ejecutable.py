"""Genera el ejecutable de Windows de Analizador de Accesos."""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def main() -> int:
    # Valida el entorno de Windows antes de invocar el script de empaquetado.
    if sys.platform != "win32":
        print("Este generador debe ejecutarse en Windows.", file=sys.stderr)
        return 1

    raiz = Path(__file__).resolve().parent
    script_compilacion = raiz / "build_windows.ps1"
    powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe")

    if not script_compilacion.is_file():
        print(f"No se encontró el script de compilación: {script_compilacion}", file=sys.stderr)
        return 1
    if not powershell:
        print("No se encontró PowerShell. Instala o habilita Windows PowerShell.", file=sys.stderr)
        return 1

    # Ejecuta la compilación dentro de la raíz del proyecto y propaga su código de salida.
    comando = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_compilacion),
    ]
    resultado = subprocess.run(comando, cwd=raiz, check=False)
    if resultado.returncode != 0:
        return resultado.returncode

    # Confirma que PyInstaller produjo el archivo esperado antes de informar éxito.
    ejecutable = raiz / "dist" / "AnalizadorDeAccesos.exe"
    if not ejecutable.is_file():
        print(f"La compilación terminó, pero no se encontró: {ejecutable}", file=sys.stderr)
        return 1

    print(f"\nListo: {ejecutable}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
