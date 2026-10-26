"""Configuración y punto de entrada del análisis de ASFI."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import ClassVar

from src.sistemas_internos.lectores import (
    DOMINIO_CONFIG, PersonasConfig, TemporalesConfig,
)
from .comun import SistemaExternoConfig, analizar_sistema_externo


@dataclass(frozen=True)
class ASFIConfig(SistemaExternoConfig):
    # Compatibilidad con el código mal digitado en la matriz ASFI.
    equivalencias_roles: ClassVar[dict[str, str]] = {"rmi-ii": "rmi-rii"}
    nombres_roles: ClassVar[dict[str, str]] = {"rmi-ii": "RMI-RII"}


def analizar_asfi(matriz: str | Path, salar: str | Path,
                  asfi_txt: str | Path, dominio: str | Path | None = None,
                  archivo_temporales: str | Path | None = None,
                  config_temporales: TemporalesConfig = TemporalesConfig(),
                  config_dominio: PersonasConfig = DOMINIO_CONFIG,
                  config_asfi: ASFIConfig = ASFIConfig()) -> dict:
    """Aplica el análisis compartido con los campos y reglas propios de ASFI."""
    resultado = analizar_sistema_externo(
        matriz, salar, asfi_txt, dominio, archivo_temporales,
        config_temporales, config_dominio, config_asfi)
    resultado["usuarios_temporales_asfi"] = len(resultado["observaciones_temporales"])
    return resultado
