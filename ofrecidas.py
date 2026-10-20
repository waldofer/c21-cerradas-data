"""Propiedades ofrecidas (a la venta) exportadas de 21online.

Modulo puro: solo openpyxl y la libreria estandar. No importa
`actualizar_cerradas` (ese arrastra playwright).

Privacidad: el export trae datos del propietario y columnas comerciales. Nada de
eso puede salir de aca; la salida se arma SIEMPRE desde la lista blanca `CAMPOS`.
"""
import hashlib
import json
import math
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

# Orden y lista blanca de claves de salida.
CAMPOS = ("id", "mu", "ti", "su", "dir", "en", "of", "pp", "mon", "mt", "mc", "fa", "la", "lo")

COLUMNAS_REQUERIDAS = {"ID", "PRECIO", "M2T", "FECHA ALTA", "MUNICIPIO"}

# Columna del export -> clave de salida (texto sin conversion).
_TEXTOS = {
    "mu": "MUNICIPIO",
    "ti": "TIPO",
    "su": "SUBTIPO",
    "dir": "DIRECCION",
    "en": "ENCABEZADO",
    "of": "OFICINA",
}

# Moneda: "USD", "PYG" o "Gs"/"Gs." (guaranies), antes o despues del numero.
_RE_MONEDA = re.compile(r"\b(USD|PYG|GS)\b\.?", re.IGNORECASE)


def parse_precio(s):
    """'420.000.000,00 PYG' -> (420000000.0, 'PYG'). None si no hay numero o moneda."""
    if s is None:
        return None
    s = str(s).strip()
    m = _RE_MONEDA.search(s)
    if not m:
        return None
    moneda = "USD" if m.group(1).upper() == "USD" else "PYG"
    numero = (s[:m.start()] + " " + s[m.end():]).strip()
    # Formato paraguayo: punto de miles, coma decimal.
    numero = numero.replace(".", "").replace(",", ".")
    try:
        valor = float(numero)
    except ValueError:
        return None
    return (valor, moneda) if math.isfinite(valor) else None  # sin NaN / inf


def validar_columnas(header):
    """Motivo (str) si falta alguna columna requerida; None si esta todo."""
    presentes = {str(h).strip() for h in header if h}
    faltantes = COLUMNAS_REQUERIDAS - presentes
    if faltantes:
        return f"Faltan columnas esperadas en el Excel de ofrecidas: {sorted(faltantes)}"
    return None


# Separador entre grupos de digitos.
_S = r"[\s.\-]"
# Cuerpo de un telefono: 6 digitos en 3+3 o en 2+2+2, que no sean cola de otro numero.
_CUERPO = rf"(?:{_S}?\d{{3}}{_S}?\d{{3}}|{_S}?\d{{2}}{_S}?\d{{2}}{_S}?\d{{2}})(?!\d)"
# Igual pero sin puntos (para formas sin prefijo, que se confunden con precios "Gs. 1.981.123.456").
_CUERPO_SIN_PUNTOS = r"(?:[\s-]?\d{3}[\s-]?\d{3}|[\s-]?\d{2}[\s-]?\d{2}[\s-]?\d{2})(?!\d)"
# Ni pegado a otro numero ni como tramo de un numero agrupado ("1 050 000 000", "1.025.000.000").
_SUELTO = r"(?<!\d)(?<!\d[\s.,\-])"

# Patrones que se borran del texto libre, en orden. Cada uno arranca comiendo los separadores
# previos para no dejar guiones colgando ("Casa 12-0981-123-456" -> "Casa 12").
_PATRONES_TELEFONO = tuple(re.compile(p) for p in (
    # Emails.
    r"[\s-]*\S+@\S+\.\w+",
    # Celulares con prefijo (+595 / 595 / 0, con o sin parentesis): basta que no sea cola de otro
    # numero pegado ("Mz 5 0981 123 456" -> "Mz 5").
    rf"[\s-]*(?<!\d)\(?(?:\+?595{_S}?|0)9\d{{2}}\)?{_CUERPO}",
    # Celulares sin prefijo: solo 96x-99x, sueltos y sin puntos.
    rf"[\s-]*{_SUELTO}9[6-9]\d{_CUERPO_SIN_PUNTOS}",
    # Fijos con codigo de area y 0 (021, 061, 0521...), sueltos.
    rf"[\s-]*{_SUELTO}\(?0[2-8]\d{{1,2}}\)?{_CUERPO}",
    # Fijos con +595 / 595 y codigo de area sin 0 (21, 61, 521...).
    rf"[\s-]*(?<![\d.,])\(?\+?595\)?{_S}?\(?0?[2-8]\d{{1,2}}\)?{_CUERPO}",
    # Fijo de Asuncion sin prefijo: 21 + 6 digitos, suelto y sin puntos ni comas que lo
    # pegen a otro numero ("Gs.21612345", "21612345,50" quedan; un punto final no molesta).
    rf"[\s-]*{_SUELTO}(?<![.,])21[\s-]?\d{{3}}[\s-]?\d{{3}}(?!\d)(?![.,]\d)",
    # Cualquier corrida de 9 o mas digitos seguidos.
    r"[\s-]*\d{9,}",
))


def quitar_telefonos(v):
    """Saca telefonos del texto libre (DIRECCION, ENCABEZADO). None si el campo queda vacio.

    Saca celulares, fijos y emails. No toca precios con puntos de miles ni numeros cortos.
    """
    if v is None:
        return None
    for patron in _PATRONES_TELEFONO:
        v = patron.sub(" ", v)
    v = re.sub(r"\s+", " ", v).strip()
    return v or None


# --- Barrera antes de publicar ---
# El filtro de arriba borra lo que reconoce. Esta barrera es mas desconfiada: si en dir/en queda
# algo con pinta de telefono (aunque sea una forma que el filtro no conoce), no se publica nada.
# Preferimos frenar una semana a tener que reescribir el historial del repo publico.

# Solo una letra corta un numero: espacios, guiones, barras, parentesis, puntos o lo que sea
# entre digitos se toma como parte del mismo numero ("(021) - 612345", "0 9 8 1 1 2 3 4 5 6").
# Antes se sacan las cosas conocidas que tienen digitos y no son telefonos.
_RE_LETRAS = re.compile(r"[^\W\d_]+")
# Monto con miles ("1.050.000.000", "Gs 900 000 000", "1,050,000,000"): nunca empieza con 0.
_RE_MONTO = re.compile(r"(?<!\d)(?<!\d[.,])[1-9]\d{0,2}([., ])\d{3}(?:\1\d{3})*(?!\d)")
_RE_MONEDA_ANTES = re.compile(r"(?:\bgs?|guaran\w*|usd|us\$|u\$s|pyg|\$|₲)[\s.:\-]*$", re.IGNORECASE)
_RE_MONEDA_DESPUES = re.compile(
    r"\s*(?:gs\b|usd\b|pyg\b|u\$s|us\$|guaran|millones|mill\b|d[oó]lares)", re.IGNORECASE
)
# Coordenada de Paraguay ("-25.2674758", "-57.6123456"): con signo y en rango de lat/lon.
_RE_COORDENADA = re.compile(r"(?<![\d.])-(?:19|2[0-8]|5[4-9]|6[0-3])\.\d{4,}")
# Fecha real con anio de 4 digitos ("12/05/2021", "1-2-2026"); un "21-61-23-45" no lo es.
_RE_FECHA = re.compile(r"(?<!\d)(?:0?[1-9]|[12]\d|3[01])([/-])(?:0?[1-9]|1[0-2])\1(?:19|20)\d{2}(?!\d)")
# Hora real ("08:00"), sin mas ":NN" detras ("12:34:56" no es hora).
_RE_HORA = re.compile(r"(?<![\d:])(?:[01]?\d|2[0-3]):[0-5]\d(?![\d:])")
# Codigo postal paraguayo (6 digitos que empiezan con 0), que el geocodificador agrega a la
# direccion: "Asuncion 001204 388", "SAN JUAN BAUTISTA 080106 080106".
_RE_CODIGO_POSTAL = re.compile(r"0\d{5}")
# Un telefono tiene a lo sumo 13 digitos (+595 981 123 456).
_MAX_DIGITOS = 13
_CORTE = "|"


def _sin_numeros_conocidos(texto: str) -> str:
    def monto(m):
        grupos = m.group().split(m.group(1))
        # El monto termina en el ultimo grupo "000" ("Gs 900 000 000 120 m2" -> deja "120").
        ultimo = max((i for i, g in enumerate(grupos) if g == "000"), default=0)
        if ultimo == 0:
            return m.group()
        con_moneda = _RE_MONEDA_ANTES.search(texto[max(0, m.start() - 12): m.start()]) or _RE_MONEDA_DESPUES.match(
            texto, m.start() + len(m.group(1).join(grupos[: ultimo + 1]))
        )
        # Sin moneda al lado solo se cree en montos de mil millones o mas escritos con puntos o
        # comas ("1.050.000.000"): "21.500.000", "500 000" o "1 021 600 000" (un "1" pegado a
        # un fijo) podrian ser un telefono.
        # Con puntos o comas y terminado en "000.000" tampoco hay duda: ningun telefono termina en
        # seis ceros ("Precio 980.000.000", "Precio 21.000.000").
        seis_ceros = m.group(1) != " " and ultimo >= 2 and grupos[ultimo - 1] == "000"
        if not con_moneda and not seis_ceros and (ultimo < 3 or m.group(1) == " "):
            return m.group()
        resto = m.group(1).join(grupos[ultimo + 1:])
        return _CORTE + (m.group(1) + resto if resto else "")

    texto = _RE_MONTO.sub(monto, texto)
    for patron in (_RE_COORDENADA, _RE_FECHA, _RE_HORA):
        texto = patron.sub(_CORTE, texto)
    return _RE_LETRAS.sub(_CORTE, texto)


def _digitos_parecen_telefono(d: str) -> bool:
    return (
        (9 <= len(d) <= _MAX_DIGITOS and (d.startswith("0") or d.startswith("595")))  # con prefijo
        or (len(d) == 9 and d[0] == "9" and d[1] in "6789")  # celular sin prefijo
        or (len(d) == 8 and d.startswith("21"))  # fijo de Asuncion sin el 0
    )


def parece_telefono(texto) -> bool:
    """True si el texto tiene un email o algo que podria ser un telefono."""
    if not texto:
        return False
    # Digitos de ancho completo ("０９８１") pasan a digitos comunes.
    texto = unicodedata.normalize("NFKC", texto)
    if "@" in texto or re.search(r"\d{9,}", texto):
        return True
    for tramo in _sin_numeros_conocidos(texto).split(_CORTE):
        grupos = re.findall(r"\d+", tramo)
        # El telefono puede empezar en cualquier grupo ("Lote 3 981 123 456"); un codigo
        # postal corta la ventana.
        for i in range(len(grupos)):
            d = ""
            for g in grupos[i:]:
                if _RE_CODIGO_POSTAL.fullmatch(g):
                    break
                d += g
                if len(d) > _MAX_DIGITOS:
                    break
                if _digitos_parecen_telefono(d):
                    return True
    return False


# IDs revisados a mano que la barrera frena sin ser telefonos. Cada uno guarda la huella del texto
# revisado: si la direccion o el titulo cambian, la excepcion deja de valer y se vuelve a frenar.
EXCEPCIONES_PATH = Path(__file__).parent / "excepciones_barrera.json"


def huella_texto(registro: dict) -> str:
    texto = f"{registro.get('dir') or ''}\n{registro.get('en') or ''}"
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()[:16]


def cargar_excepciones(path: Path = EXCEPCIONES_PATH) -> dict[str, str]:
    path = Path(path)
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def telefonos_residuales(registros: list[dict], excepciones: dict[str, str] | None = None) -> list[int]:
    """IDs de los registros cuya direccion o titulo tiene algo con pinta de telefono.

    `excepciones` ({id: huella}) saltea los revisados a mano, solo si el texto no cambio."""
    if excepciones is None:
        excepciones = cargar_excepciones()
    return [
        r["id"] for r in registros
        if (parece_telefono(r.get("dir")) or parece_telefono(r.get("en")))
        and excepciones.get(str(r["id"])) != huella_texto(r)
    ]


def _lista_ids(ids: list[int], tope: int = 20) -> str:
    return ", ".join(str(i) for i in ids[:tope]) + (" ..." if len(ids) > tope else "")


def _texto(v):
    """Celda -> str sin espacios, o None si esta vacia."""
    if v is None:
        return None
    v = str(v).strip()
    return v or None


def _numero(v):
    """Celda de texto/numero -> float, o None si esta vacia, no es un numero o no es finito."""
    try:
        valor = float(_texto(v))
    except (TypeError, ValueError):
        return None
    return valor if math.isfinite(valor) else None


def parsear_ofrecidas(path: Path) -> list[dict]:
    """Un dict por ID (gana la ultima fila), solo con las claves de CAMPOS.

    Se descartan las filas sin ID, con precio ilegible o con M2T faltante o <= 0.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active  # el nombre de la hoja varia segun la exportacion
        filas = ws.iter_rows(values_only=True)
        header = next(filas, None) or ()
        motivo = validar_columnas(header)
        if motivo:
            raise ValueError(motivo)
        idx = {str(h).strip(): i for i, h in enumerate(header) if h}

        def celda(fila, col):
            i = idx.get(col)
            return fila[i] if i is not None and i < len(fila) else None

        por_id = {}
        for fila in filas:
            pid = _numero(celda(fila, "ID"))
            precio = parse_precio(celda(fila, "PRECIO"))
            m2t = _numero(celda(fila, "M2T"))
            if pid is None or precio is None or m2t is None or m2t <= 0:
                continue
            fa = _texto(celda(fila, "FECHA ALTA"))
            valores = {
                "id": int(pid),
                "pp": precio[0],
                "mon": precio[1],
                "mt": m2t,
                "mc": _numero(celda(fila, "M2C")),
                "fa": fa[:10] if fa else None,
                "la": _numero(celda(fila, "LAT")),
                "lo": _numero(celda(fila, "LON")),
            }
            for clave, col in _TEXTOS.items():
                valores[clave] = _texto(celda(fila, col))
            # Texto libre: el agente a veces escribe telefonos adentro.
            valores["dir"] = quitar_telefonos(valores["dir"])
            valores["en"] = quitar_telefonos(valores["en"])
            # Lista blanca: solo salen las claves de CAMPOS.
            por_id[valores["id"]] = {k: valores[k] for k in CAMPOS}
        return list(por_id.values())
    finally:
        wb.close()


# Si la descarga trae menos del 80% de las ofrecidas de la semana anterior,
# se la trata como rota y no se publica.
UMBRAL_DESCARGA = 0.8


def descarga_sospechosa(n_nuevas: int, n_anteriores: int, umbral: float = UMBRAL_DESCARGA):
    """Mensaje (str) si la descarga parece cortada; None si es normal o es la primera corrida."""
    if n_anteriores == 0 or n_nuevas >= umbral * n_anteriores:
        return None
    return (
        f"La descarga de ofrecidas trajo {n_nuevas} propiedades y la semana anterior "
        f"habia {n_anteriores} (menos del {umbral:.0%}); se la considera incompleta."
    )


def comparar(anteriores: list[dict], descargadas: list[dict], vendidas_por_id: dict[int, dict]) -> dict:
    """Compara la descarga de esta semana con la anterior y con las vendidas.

    Una ofrecida que ya figura en `vendidas_por_id` no se publica como ofrecida.
    """
    ant = {r["id"]: r for r in anteriores}
    desc = {r["id"]: r for r in descargadas}

    ofrecidas = sorted(
        (r for i, r in desc.items() if i not in vendidas_por_id), key=lambda r: r["id"]
    )
    nuevas = [r for r in ofrecidas if r["id"] not in ant]
    pasaron_a_vendidas = [
        {
            "id": i,
            "dir": r["dir"],
            "mu": r["mu"],
            "pp": r["pp"],
            "mon": r["mon"],
            "pr_cierre": vendidas_por_id[i]["pr"],
        }
        for i, r in ant.items()
        if i in vendidas_por_id
    ]
    dadas_de_baja = [
        r for i, r in ant.items() if i not in desc and i not in vendidas_por_id
    ]
    cambios_precio = [
        {
            "id": i,
            "antes": (ant[i]["pp"], ant[i]["mon"]),
            "despues": (r["pp"], r["mon"]),
        }
        for i, r in desc.items()
        if i in ant and (ant[i]["pp"], ant[i]["mon"]) != (r["pp"], r["mon"])
    ]
    return {
        "ofrecidas": ofrecidas,
        "nuevas": nuevas,
        "pasaron_a_vendidas": pasaron_a_vendidas,
        "dadas_de_baja": dadas_de_baja,
        "cambios_precio": cambios_precio,
        "en_venta_antes": len(anteriores),
    }


OFRECIDAS_PATH = Path(__file__).parent / "data" / "ofrecidas.json"
CERRADAS_PATH = Path(__file__).parent / "data" / "cerradas.json"


def guardar_ofrecidas(registros: list[dict], path: Path = OFRECIDAS_PATH) -> None:
    """JSON compacto ordenado por id. Escribe a .tmp y reemplaza, para no dejar un archivo roto."""
    path = Path(path)
    # Ultimo cerrojo: aunque alguien llame a esto sin pasar por procesar_ofrecidas (ej. `inicial`).
    sospechosos = telefonos_residuales(registros)
    if sospechosos:
        raise ValueError(
            f"No se guarda ofrecidas.json: {len(sospechosos)} propiedades tienen algo con pinta de "
            f"telefono en la direccion o el titulo. IDs: {_lista_ids(sospechosos)}"
        )
    ordenados = sorted(registros, key=lambda r: r["id"])
    tmp = path.with_name(path.name + ".tmp")
    # allow_nan=False: un NaN/inf no es JSON valido; mejor fallar que publicar un archivo roto.
    texto = json.dumps(ordenados, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    tmp.write_text(texto, encoding="utf-8")
    tmp.replace(path)


def _cerradas_por_id(path: Path = CERRADAS_PATH) -> dict[int, dict]:
    with open(path, encoding="utf-8") as f:
        return {r["id"]: r for r in json.load(f)}


def _filas_de_datos(ruta: Path) -> int:
    """Cantidad de filas con algun valor debajo del encabezado."""
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        filas = wb.active.iter_rows(values_only=True)
        next(filas, None)
        return sum(1 for f in filas if any(c is not None and str(c).strip() for c in f))
    finally:
        wb.close()


def _header(ruta: Path) -> tuple:
    """Fila de encabezado del Excel (tupla vacia si no hay)."""
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        return next(wb.active.iter_rows(values_only=True), None) or ()
    finally:
        wb.close()


# Estatus que tiene que traer cada fila del export de ofrecidas.
ESTATUS_ESPERADO = "enPromocion"
# Mas nuevas que esto en una semana (con semana anterior) es una descarga inflada.
LIMITE_NUEVAS_OFRECIDAS = 1500


def estatus_inesperados(ruta: Path) -> int:
    """Cantidad de filas con ID cuyo ESTATUS no es enPromocion (o falta la columna).

    El ESTATUS solo se usa para validar; no sale del modulo.
    """
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        filas = wb.active.iter_rows(values_only=True)
        header = next(filas, None) or ()
        idx = {str(h).strip(): i for i, h in enumerate(header) if h}
        i_id, i_est = idx.get("ID"), idx.get("ESTATUS")
        n = 0
        for fila in filas:
            if i_id is None or i_id >= len(fila) or _texto(fila[i_id]) is None:
                continue
            estatus = _texto(fila[i_est]) if i_est is not None and i_est < len(fila) else None
            if estatus != ESTATUS_ESPERADO:
                n += 1
        return n
    finally:
        wb.close()


def procesar_ofrecidas(
    ruta_export: Path, anteriores: list[dict], vendidas_por_id: dict[int, dict]
) -> tuple[list[dict] | None, dict | None, str | None]:
    """Parsea el export y lo cruza con la semana anterior y con las vendidas.

    Devuelve (final, resultado, alerta). Si hay alerta, final es None y no se debe escribir nada.
    """
    motivo = validar_columnas(_header(ruta_export))
    if motivo:
        return None, None, motivo
    try:
        descargadas = parsear_ofrecidas(ruta_export)
    except ValueError as e:
        return None, None, str(e)
    if not descargadas:
        return None, None, "El export de ofrecidas no trajo ninguna propiedad valida."
    otros = estatus_inesperados(ruta_export)
    if otros:
        return None, None, (
            f"El export de ofrecidas trae {otros} filas con estatus distinto de "
            f"{ESTATUS_ESPERADO}; no se publica."
        )
    alerta = descarga_sospechosa(len(descargadas), len(anteriores))
    if alerta:
        return None, None, alerta
    resultado = comparar(anteriores, descargadas, vendidas_por_id)
    n_nuevas = len(resultado["nuevas"])
    if anteriores and n_nuevas > LIMITE_NUEVAS_OFRECIDAS:
        return None, resultado, (
            f"Las ofrecidas nuevas ({n_nuevas}) superan el limite de seguridad "
            f"({LIMITE_NUEVAS_OFRECIDAS}); se la considera una descarga inflada."
        )
    sospechosos = set(telefonos_residuales(resultado["ofrecidas"]))
    if sospechosos:
        # IDs y huellas, nunca el texto. Si se revisan en 21online y no son telefonos, se copian
        # "id": "huella" a excepciones_barrera.json y se corre de nuevo con --solo-ofrecidas.
        detalle = [f"{r['id']} ({huella_texto(r)})" for r in resultado["ofrecidas"] if r["id"] in sospechosos]
        return None, resultado, (
            f"{len(sospechosos)} propiedades tienen en la direccion o el titulo algo con pinta de "
            f"telefono que el filtro no limpio; esta semana no se publica. ID (huella): "
            f"{', '.join(detalle[:20])}{' ...' if len(detalle) > 20 else ''}"
        )
    return resultado["ofrecidas"], resultado, None


def escribir_hoja_ofrecidas(ruta_resumen: Path, resultado: dict | None, alerta: str | None) -> None:
    """Crea o reemplaza la hoja "Ofrecidas" de resumen_actualizacion.xlsx (crea el libro si falta)."""
    ruta_resumen = Path(ruta_resumen)
    if ruta_resumen.exists():
        wb = openpyxl.load_workbook(ruta_resumen)
        if "Ofrecidas" in wb.sheetnames:
            del wb["Ofrecidas"]
        ws = wb.create_sheet("Ofrecidas")
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Ofrecidas"

    ws.append(["Estado", f"ALERTA - {alerta}" if alerta else "OK"])
    if resultado:
        ws.append(["En venta antes", resultado["en_venta_antes"]])
        ws.append(["Nuevas", len(resultado["nuevas"])])
        ws.append(["Pasaron a vendidas", len(resultado["pasaron_a_vendidas"])])
        ws.append(["Dadas de baja", len(resultado["dadas_de_baja"])])
        ws.append(["En venta ahora", len(resultado["ofrecidas"])])
        ws.append([])
        ws.append(["Pasaron a vendidas"])
        ws.append(["ID", "Dirección", "Barrio", "Precio pedido", "Moneda", "Precio de cierre USD"])
        vendidas = set()
        for r in resultado["pasaron_a_vendidas"]:
            vendidas.add(r["id"])
            ws.append([r["id"], r["dir"], r["mu"], r["pp"], r["mon"], r["pr_cierre"]])
        ws.append([])
        ws.append(["Cambios de precio"])
        ws.append(["ID", "Antes", "Moneda", "Después", "Moneda"])
        for c in resultado["cambios_precio"]:
            if c["id"] in vendidas:  # ya figura como vendida: su cambio de precio es ruido
                continue
            ws.append([c["id"], c["antes"][0], c["antes"][1], c["despues"][0], c["despues"][1]])
    wb.save(ruta_resumen)


def _inicial(ruta: Path) -> None:
    leidas = _filas_de_datos(ruta)
    parseadas = parsear_ofrecidas(ruta)  # valida columnas; levanta ValueError si faltan
    publicables = comparar([], parseadas, _cerradas_por_id())["ofrecidas"]
    guardar_ofrecidas(publicables)
    print(
        f"Filas leidas: {leidas} | descartadas (sin M2T, precio ilegible, sin ID o repetidas): "
        f"{leidas - len(parseadas)} | ya vendidas (excluidas): {len(parseadas) - len(publicables)} "
        f"| guardadas: {len(publicables)}"
    )


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "inicial":
        _inicial(Path(sys.argv[2]))
    else:
        sys.exit("Uso: python ofrecidas.py inicial <ruta.xlsx>")
