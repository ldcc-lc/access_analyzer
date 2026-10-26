"""Configuración y punto de entrada del análisis de INFOCRED."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import ClassVar

from src.sistemas_internos.lectores import (
    DOMINIO_CONFIG, PersonasConfig, TemporalesConfig,
)
from .comun import SistemaExternoConfig, analizar_sistema_externo


@dataclass(frozen=True)
class INFOCREDConfig(SistemaExternoConfig):
    nombre_sistema: str = "INFOCRED"
    campos_txt: int = 17
    campo_correo: int = 7
    campo_nombre: int | None = None
    campo_rol: int = 9
    campo_estado: int | None = 13
    estado_inactivo: str = "BAJA"
    seccion_codigos: str = "Infocred"
    columna_rol_matriz: str = "C"
    columna_codigo_txt: str = "B"
    columna_matriz: str = "P"
    # En la matriz vigente el nombre empieza con ele minúscula: "lnfocred".
    aliases_seccion_codigos: ClassVar[tuple[str, ...]] = ("lnfocred",)


def analizar_infocred(matriz: str | Path, salar: str | Path,
                      infocred_txt: str | Path, dominio: str | Path | None = None,
                      archivo_temporales: str | Path | None = None,
                      config_temporales: TemporalesConfig = TemporalesConfig(),
                      config_dominio: PersonasConfig = DOMINIO_CONFIG,
                      config_infocred: INFOCREDConfig = INFOCREDConfig()) -> dict:
    """Aplica el análisis compartido con los campos y reglas de INFOCRED."""
    return analizar_sistema_externo(
        matriz, salar, infocred_txt, dominio, archivo_temporales,
        config_temporales, config_dominio, config_infocred)
