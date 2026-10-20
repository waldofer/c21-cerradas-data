import json

import openpyxl

from ofrecidas import (
    CAMPOS,
    COLUMNAS_REQUERIDAS,
    comparar,
    descarga_sospechosa,
    quitar_telefonos,
    guardar_ofrecidas,
    parse_precio,
    parsear_ofrecidas,
    validar_columnas,
)

# Encabezado real del export de 21online (el orden es el del sistema).
HEADER = [
    "LICENCIA", "OFICINA", "FHE", "COMERCIAL", "ID", "CLAVE", "ESTATUS", "DIRECCION",
    "COLONIA", "MUNICIPIO", "ESTADO", "PAIS", "DESARROLLO", "FECHA ALTA",
    "FECHA PROMOCIÓN", "TIPO", "SUBTIPO", "OPERACIÓN", "PRECIO", "TIPO CONTRATO",
    "DIAS CONTRATO", "PRECIO CONTRATO", "TIPO DE COMISION", "COMISION CONTRATO",
    "COMISION A COMPARTIR", "ANTIGUEDAD", "PISO", "HABITACIONES", "M2T", "M2C",
    "LAT", "LON", "INTERNET", "ASESOR EXCLUSIVA", "PROPIETARIO", "TELEFONO PROPIETARIO",
    "EMAIL PROPIETARIO", "FECHA SUSPENDIDA", "FECHA INTENCIÓN", "FECHA CANCELACIÓN",
    "FECHA CIERRE", "PRECIO CIERRE", "COMISIÓN CIERRE", "COMISIÓN CIERRE VENDEDOR",
    "COMISIÓN CIERRE COMPRADOR", "ASESOR CIERRE", "COMPARTIDA", "COMPARTIDA CON",
    "Contraparte", "COMISION COMPARTIDA", "REFERIDO", "REFERIDO CON", "CANTIDAD REFERIDA",
    "MES CIERRE", "ZONA", "ENCABEZADO", "NOTAS",
]

PROPIETARIO = "Juan Perez"
TELEFONO = "0981123456"
EMAIL = "juan@example.com"


def fila_valida(**cambios):
    """Una fila valida con todas las celdas como texto, igual que el export real."""
    fila = {
        "OFICINA": "CENTURY 21 Demo",
        "ID": "51505",
        "DIRECCION": "Calle Falsa 123",
        "MUNICIPIO": "Asuncion",
        "FECHA ALTA": "2026-07-17",
        "TIPO": "Casa",
        "SUBTIPO": "Casa",
        "PRECIO": "420.000.000,00 PYG",
        "M2T": "360",
        "M2C": "200",
        "LAT": "-25.26",
        "LON": "-57.58",
        "PROPIETARIO": PROPIETARIO,
        "TELEFONO PROPIETARIO": TELEFONO,
        "EMAIL PROPIETARIO": EMAIL,
        "COMISION CONTRATO": "3",
        "NOTAS": "Llamar al propietario",
        "ENCABEZADO": "Casa en Asuncion",
    }
    fila.update(cambios)
    return fila


def armar_excel(tmp_path, filas, columnas_extra=()):
    """Escribe un xlsx con el encabezado real (+ columnas extra); celdas ausentes quedan vacias."""
    header = HEADER + list(columnas_extra)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for fila in filas:
        ws.append([fila.get(h, "") for h in header])
    path = tmp_path / "ofrecidas.xlsx"
    wb.save(path)
    return path


def test_parse_precio_pyg():
    assert parse_precio("420.000.000,00 PYG") == (420000000.0, "PYG")


def test_parse_precio_usd_sufijo_y_prefijo():
    assert parse_precio("150.000,00 USD") == (150000.0, "USD")
    assert parse_precio("USD 150.000") == (150000.0, "USD")


def test_parse_precio_chico_se_acepta():
    assert parse_precio("50,00 USD") == (50.0, "USD")


def test_parse_precio_ilegible():
    for x in (None, "", "Consultar", "150.000"):
        assert parse_precio(x) is None


def test_parse_precio_gs_es_pyg():
    assert parse_precio("Gs. 1.500.000") == (1500000.0, "PYG")
    assert parse_precio("350.000.000 Gs") == (350000000.0, "PYG")


def test_parsear_solo_campos_permitidos(tmp_path):
    filas = [
        fila_valida(**{"CELULAR PROPIETARIO": "0971999888"}),
        fila_valida(ID="51506"),
    ]
    path = armar_excel(tmp_path, filas, columnas_extra=("CELULAR PROPIETARIO",))
    resultado = parsear_ofrecidas(path)
    assert len(resultado) == 2
    for r in resultado:
        assert set(r) == set(CAMPOS)
    volcado = json.dumps(resultado)
    for dato in (PROPIETARIO, TELEFONO, EMAIL, "0971999888", "Llamar al propietario"):
        assert dato not in volcado


def test_parsear_descarta_incompletas(tmp_path):
    filas = [
        fila_valida(ID=""),
        fila_valida(ID="1", PRECIO="Consultar"),
        fila_valida(ID="2", M2T=""),
        fila_valida(ID="3", M2T="0"),
    ]
    assert parsear_ofrecidas(armar_excel(tmp_path, filas)) == []

    r = parsear_ofrecidas(armar_excel(tmp_path, [fila_valida()]))
    assert len(r) == 1
    assert r[0]["id"] == 51505
    assert r[0]["pp"] == 420000000.0
    assert r[0]["mon"] == "PYG"
    assert r[0]["mt"] == 360.0
    assert r[0]["mc"] == 200.0
    assert r[0]["fa"] == "2026-07-17"
    assert r[0]["la"] == -25.26
    assert r[0]["lo"] == -57.58
    assert r[0]["mu"] == "Asuncion"
    assert r[0]["of"] == "CENTURY 21 Demo"
    assert r[0]["en"] == "Casa en Asuncion"


def test_parsear_celdas_vacias_son_none(tmp_path):
    fila = fila_valida(M2C="", LAT="", LON="", **{"FECHA ALTA": "", "ENCABEZADO": ""})
    (r,) = parsear_ofrecidas(armar_excel(tmp_path, [fila]))
    assert r["mc"] is None and r["la"] is None and r["lo"] is None
    assert r["fa"] is None and r["en"] is None


def test_parsear_id_duplicado(tmp_path):
    filas = [
        fila_valida(ID="51505", PRECIO="100.000,00 USD"),
        fila_valida(ID="51505", PRECIO="120.000,00 USD"),
    ]
    r = parsear_ofrecidas(armar_excel(tmp_path, filas))
    assert len(r) == 1
    assert r[0]["pp"] == 120000.0


def test_validar_columnas():
    sin_precio = [h for h in HEADER if h != "PRECIO"]
    assert "PRECIO" in validar_columnas(sin_precio)
    assert validar_columnas(HEADER) is None
    assert COLUMNAS_REQUERIDAS <= set(HEADER)


def rec(id, pp=100000.0, mon="USD", **cambios):
    """Registro de ofrecida ya parseado (forma de salida de parsear_ofrecidas)."""
    r = {k: None for k in CAMPOS}
    r.update(id=id, mu="Asuncion", dir=f"Calle {id}", pp=pp, mon=mon, mt=100.0)
    r.update(cambios)
    return r


def ids(lista):
    return [r["id"] for r in lista]


def test_comparar_nueva():
    res = comparar([rec(1)], [rec(1), rec(2)], {})
    assert ids(res["ofrecidas"]) == [1, 2]
    assert ids(res["nuevas"]) == [2]
    assert res["en_venta_antes"] == 1


def test_comparar_paso_a_vendidas():
    res = comparar([rec(1), rec(2)], [rec(1)], {2: {"pr": 80000.0}})
    assert ids(res["ofrecidas"]) == [1]
    assert res["pasaron_a_vendidas"][0]["id"] == 2
    assert res["pasaron_a_vendidas"][0]["pr_cierre"] == 80000.0
    assert res["dadas_de_baja"] == []


def test_comparar_dada_de_baja():
    res = comparar([rec(1), rec(2)], [rec(1)], {})
    assert ids(res["dadas_de_baja"]) == [2]
    assert res["pasaron_a_vendidas"] == []


def test_comparar_esta_en_las_dos():
    res = comparar([rec(1)], [rec(1)], {1: {"pr": 90000.0}})
    assert res["ofrecidas"] == []
    assert ids(res["pasaron_a_vendidas"]) == [1]
    assert res["dadas_de_baja"] == []
    assert res["nuevas"] == []


def test_comparar_cambio_precio():
    antes = rec(1, pp=420000000.0, mon="PYG")
    despues = rec(1, pp=400000000.0, mon="PYG")
    res = comparar([antes], [despues], {})
    assert res["cambios_precio"] == [
        {"id": 1, "antes": (420000000.0, "PYG"), "despues": (400000000.0, "PYG")}
    ]
    assert comparar([antes], [dict(antes)], {})["cambios_precio"] == []


def test_comparar_pasaron_a_vendidas_trae_datos_de_la_ofrecida():
    ant = rec(7, pp=150000.0, mon="USD", dir="Av. Mcal. Lopez 100", mu="Luque")
    (v,) = comparar([ant], [], {7: {"pr": 140000.0}})["pasaron_a_vendidas"]
    assert v == {"id": 7, "dir": "Av. Mcal. Lopez 100", "mu": "Luque",
                 "pp": 150000.0, "mon": "USD", "pr_cierre": 140000.0}


def test_comparar_nueva_vendida_no_cuenta_como_nueva():
    res = comparar([], [rec(5)], {5: {"pr": 1.0}})
    assert res["ofrecidas"] == [] and res["nuevas"] == []
    assert res["pasaron_a_vendidas"] == []


def test_descarga_sospechosa_corta():
    msg = descarga_sospechosa(3000, 8440)
    assert isinstance(msg, str)
    assert "3000" in msg and "8440" in msg


def test_descarga_sospechosa_normal():
    assert descarga_sospechosa(8300, 8440) is None
    assert descarga_sospechosa(6752, 8440) is None  # exactamente 80%
    assert descarga_sospechosa(5, 0) is None  # primera corrida


def test_guardar_ofrecidas_ida_y_vuelta(tmp_path):
    destino = tmp_path / "ofrecidas.json"
    registros = [
        {"id": 20, "dir": "Calle Ñandú", "pp": 1.5},
        {"id": 3, "dir": "Av. España", "pp": 2.0},
    ]
    guardar_ofrecidas(registros, destino)
    assert json.loads(destino.read_text(encoding="utf-8")) == sorted(registros, key=lambda r: r["id"])
    assert "Ñandú" in destino.read_text(encoding="utf-8")  # sin escapes ASCII
    assert list(tmp_path.glob("*.tmp")) == []


TELEFONOS = [
    "0981123456", "0981 123 456", "0981-123-456", "+595 981 123456",
    "595981123456", "+595981123456", "0971-234 567",
]


def test_quitar_telefonos_formas():
    for tel in TELEFONOS:
        assert quitar_telefonos(f"Villa Gioconda {tel}") == "Villa Gioconda", tel
        assert quitar_telefonos(f"Llamar {tel} ahora") == "Llamar ahora", tel
    assert quitar_telefonos("1234567890123") is None


def test_quitar_telefonos_respeta_precios_y_numeros_cortos():
    for txt in [
        "Gs. 420.000.000", "USD 150.000", "Casa 3", "lote 12", "1440",
        "Manzana 12 lote 3 - 1440 m2", "Gs. 900 000 000",
    ]:
        assert quitar_telefonos(txt) == txt
    assert quitar_telefonos(None) is None


def test_quitar_telefonos_campo_solo_telefono_queda_none():
    assert quitar_telefonos("0981 123 456") is None
    assert quitar_telefonos("   ") is None


def _excel_con(tmp_path, dir_, en):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADER)
    fila = {h: None for h in HEADER}
    fila.update({"ID": 1, "PRECIO": "100.000,00 USD", "M2T": 500, "FECHA ALTA": "2026-01-01",
                 "MUNICIPIO": "Luque", "DIRECCION": dir_, "ENCABEZADO": en})
    ws.append([fila[h] for h in HEADER])
    ruta = tmp_path / "o.xlsx"
    wb.save(ruta)
    return ruta


def test_parsear_ofrecidas_quita_telefonos_de_dir_y_en(tmp_path):
    ruta = _excel_con(tmp_path, "0981-123-456", "Terreno +595 981 123456 en Luque")
    (r,) = parsear_ofrecidas(ruta)
    assert r["dir"] is None
    assert r["en"] == "Terreno en Luque"
