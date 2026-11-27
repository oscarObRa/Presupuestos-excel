"""
factory.py
==========
Fábrica de ADNs visuales por empresa.

El resto del sistema NUNCA instancia un ADN concreto: siempre pide
`get_dna("icel")` (o el nombre de la empresa que venga en el proyecto).
Así, agregar una nueva empresa es:
1. Crear `src/visual_dna/<empresa>.py` con una clase que herede de BaseVisualDNA.
2. Registrarla en `VisualDNAFactory.REGISTRO` (o pasar la clase a `registrar`).
"""

from src.visual_dna.base import BaseVisualDNA
from src.visual_dna.icel import ICELVisualDNA


class VisualDNAFactory:
    """Registro de identidades visuales disponibles, por clave de empresa."""

    REGISTRO: dict[str, type] = {
        "icel": ICELVisualDNA,
    }

    DEFAULT: str = "icel"  # empresa por defecto mientras exista solo una

    @classmethod
    def registrar(cls, clave: str, clase_dna: type):
        """Registra (o sustituye) el ADN de una empresa."""
        if not issubclass(clase_dna, BaseVisualDNA):
            raise TypeError(f"{clase_dna.__name__} debe heredar de BaseVisualDNA")
        cls.REGISTRO[clave.lower()] = clase_dna

    @classmethod
    def empresas_disponibles(cls) -> list[str]:
        return sorted(cls.REGISTRO.keys())

    @classmethod
    def get_dna(cls, empresa: str | None = None) -> BaseVisualDNA:
        """Devuelve una instancia del ADN visual de la empresa indicada.

        - `None` o vacío → ADN por defecto (ICEL, mientras sea la única).
        - Clave desconocida → error explícito (no silenciar con ICEL: la
          identidad visual es contractual, mejor fallar que maquillar).
        """
        clave = (empresa or cls.DEFAULT or "").strip().lower()

        # Tolerancia: el nombre de la empresa puede venir largo
        # ("ICEL Contratistas S.A. de C.V."), se mapea por prefijo conocido.
        if clave not in cls.REGISTRO:
            for registrada in cls.REGISTRO:
                if registrada in clave:
                    clave = registrada
                    break

        if clave not in cls.REGISTRO:
            raise ValueError(
                f"Empresa sin ADN visual registrado: '{empresa}'. "
                f"Disponibles: {', '.join(cls.empresas_disponibles())}"
            )
        return cls.REGISTRO[clave]()


def get_dna(empresa: str | None = None) -> BaseVisualDNA:
    """Atajo funcional: get_dna('icel') -> ICELVisualDNA()."""
    return VisualDNAFactory.get_dna(empresa)
