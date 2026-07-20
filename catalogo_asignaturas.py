# -*- coding: utf-8 -*-
"""Catálogo único y reconocimiento seguro de asignaturas oficiales."""

from __future__ import annotations

from difflib import SequenceMatcher
import json
from pathlib import Path
import re
import unicodedata


_CATALOGO_PATH = Path(__file__).with_name("catalogo_asignaturas.json")
with _CATALOGO_PATH.open("r", encoding="utf-8") as archivo_catalogo:
    CATALOGO_ASIGNATURAS = json.load(archivo_catalogo)


def normalizar_texto_asignatura(texto) -> str:
    """Normaliza sólo para comparar; nunca reemplaza el nombre canónico visible."""
    if texto is None:
        return ""
    valor = str(texto).strip().upper()
    valor = valor.replace("–", "-").replace("—", "-")
    valor = unicodedata.normalize("NFD", valor)
    valor = "".join(c for c in valor if unicodedata.category(c) != "Mn")
    valor = re.sub(r"[-_/|;:,]+", " ", valor)
    valor = re.sub(r"[^A-Z0-9\s]", " ", valor)
    return " ".join(valor.split())


def normalizar_grado(grado) -> str | None:
    """Convierte las variantes de grado usadas por la app/Excel a un código común."""
    texto = normalizar_texto_asignatura(grado)
    if not texto:
        return None

    if "INICIAL" in texto:
        if re.search(r"(?:^|\s)(?:2|II)(?:\s|$)", texto):
            return "INICIAL_2"
        if re.search(r"(?:^|\s)(?:1|I)(?:\s|$)", texto):
            return "INICIAL_1"

    es_bgu = "BACHILLERATO" in texto or re.search(r"(?:^|\s)BGU(?:\s|$)", texto)
    if es_bgu:
        if re.search(r"(?:^|\s)(?:3RO|3ER|3|TERCERO|TERCER)(?:\s|$)", texto):
            return "BGU_3"
        if re.search(r"(?:^|\s)(?:2DO|2|SEGUNDO)(?:\s|$)", texto):
            return "BGU_2"
        if re.search(r"(?:^|\s)(?:1RO|1ER|1|PRIMERO|PRIMER)(?:\s|$)", texto):
            return "BGU_1"
        return None

    es_egb = "EGB" in texto or "EDUCACION GENERAL BASICA" in texto
    if "PREPARATORIA" in texto or (es_egb and re.search(r"(?:^|\s)(?:1RO|1ER|1|PRIMERO|PRIMER)(?:\s|$)", texto)):
        return "EGB_1"
    if "ELEMENTAL" in texto or (es_egb and re.search(r"(?:^|\s)(?:2DO|3RO|4TO|2|3|4|SEGUNDO|TERCERO|CUARTO)(?:\s|$)", texto)):
        return "EGB_ELEMENTAL"
    if "MEDIA" in texto or (es_egb and re.search(r"(?:^|\s)(?:5TO|6TO|7MO|5|6|7|QUINTO|SEXTO|SEPTIMO)(?:\s|$)", texto)):
        return "EGB_MEDIA"
    if "SUPERIOR" in texto or (es_egb and re.search(r"(?:^|\s)(?:8VO|9NO|10MO|8|9|10|OCTAVO|NOVENO|DECIMO)(?:\s|$)", texto)):
        return "EGB_SUPERIOR"
    return None


def grados_equivalentes(grado_a, grado_b) -> bool:
    codigo_a = normalizar_grado(grado_a)
    codigo_b = normalizar_grado(grado_b)
    return bool(codigo_a and codigo_b and codigo_a == codigo_b)


def _variantes(entrada: dict) -> list[tuple[str, str]]:
    valores = [(entrada["nombre"], "coincidencia_exacta")]
    valores.extend((alias, "alias") for alias in entrada.get("aliases", []))
    return valores


def _umbral_conservador(texto_normalizado: str) -> float:
    longitud = len(texto_normalizado.replace(" ", ""))
    if longitud <= 5:
        return 0.96
    if longitud <= 8:
        return 0.88
    if longitud <= 14:
        return 0.86
    return 0.84


def _es_extension_de_variante(texto: str, variante: str) -> bool:
    palabras_texto = texto.split()
    palabras_variante = variante.split()
    return len(palabras_texto) > len(palabras_variante) and set(palabras_variante) < set(palabras_texto)


def reconocer_asignatura_global(texto) -> dict | None:
    """Reconoce contra todo el catálogo antes de aplicar el filtro del curso."""
    original = str(texto).strip() if texto is not None else ""
    normalizado = normalizar_texto_asignatura(original)
    if not normalizado:
        return None

    for entrada in CATALOGO_ASIGNATURAS:
        for variante, metodo in _variantes(entrada):
            if normalizado == normalizar_texto_asignatura(variante):
                return {"original": original, "canonica": entrada["nombre"], "metodo": metodo, "entrada": entrada, "similitud": 1.0}

    puntuaciones = []
    for entrada in CATALOGO_ASIGNATURAS:
        mejor = 0.0
        for variante, _ in _variantes(entrada):
            variante_normalizada = normalizar_texto_asignatura(variante)
            if _es_extension_de_variante(normalizado, variante_normalizada):
                continue
            if abs(len(normalizado.split()) - len(variante_normalizada.split())) > 1:
                continue
            mejor = max(mejor, SequenceMatcher(None, normalizado, variante_normalizada).ratio())
        puntuaciones.append((mejor, entrada))

    puntuaciones.sort(key=lambda item: item[0], reverse=True)
    mejor_puntaje, mejor_entrada = puntuaciones[0]
    segundo_puntaje = puntuaciones[1][0] if len(puntuaciones) > 1 else 0.0
    if mejor_puntaje < _umbral_conservador(normalizado):
        return None
    if mejor_puntaje - segundo_puntaje < 0.06:
        return None

    return {
        "original": original,
        "canonica": mejor_entrada["nombre"],
        "metodo": "coincidencia_aproximada",
        "entrada": mejor_entrada,
        "similitud": round(mejor_puntaje, 4),
    }


def clasificar_asignatura(texto, grado_activo) -> dict:
    reconocimiento = reconocer_asignatura_global(texto)
    if reconocimiento is None:
        return {"estado": "no_reconocida", "original": str(texto).strip()}

    codigo_grado = normalizar_grado(grado_activo)
    if codigo_grado not in reconocimiento["entrada"].get("grados", []):
        return {
            "estado": "ignorada_por_curso",
            "original": reconocimiento["original"],
            "canonica": reconocimiento["canonica"],
            "motivo": f"No pertenece a {grado_activo}",
            "metodo": reconocimiento["metodo"],
        }

    return {
        "estado": "reconocida",
        "original": reconocimiento["original"],
        "canonica": reconocimiento["canonica"],
        "metodo": reconocimiento["metodo"],
        "tipo": reconocimiento["entrada"].get("tipo", "cuantitativa"),
        "orden": reconocimiento["entrada"].get("orden", 999),
    }


def orden_asignatura(nombre: str) -> tuple[int, str]:
    normalizado = normalizar_texto_asignatura(nombre)
    for entrada in CATALOGO_ASIGNATURAS:
        if normalizar_texto_asignatura(entrada["nombre"]) == normalizado:
            return int(entrada.get("orden", 999)), entrada["nombre"]
    return 999, str(nombre)
