"""Registro de fuentes. Cada módulo registra sus tipos con @register("nombre")."""
from .base import REGISTRY, SourceResult, apply_transform  # noqa: F401
from . import apis  # noqa: F401  (bcb_sgs, ibge_sidra, bcrp, bis, fred, banxico, inegi, bcch_siete, bcch_api)
from . import web_files  # noqa: F401  (archivo_web)
from . import text_values  # noqa: F401  (texto_web: boletines PDF/HTML)
from .. import news  # noqa: F401,E402  (noticia_dato)
