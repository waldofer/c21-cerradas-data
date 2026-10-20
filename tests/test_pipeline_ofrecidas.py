import sys

import pytest

import actualizar_cerradas as ac


def test_url_vendidas_no_cambia():
    assert ac.construir_url_busqueda("2026-10-01", "2026-10-19") == (
        "https://21online.century21.com.py/inventarios?oficinaStatus=todas&oficina="
        "&estatus=cerrada&tipoOperacion=venta&mesCierreDesde=2026-10-01"
        "&mesCierreHasta=2026-10-19&cantidad=30000&conFoto=1"
    )


def test_url_ofrecidas():
    url = ac.construir_url_busqueda("2015-01-01", "2026-10-19", estatus="enPromocion")
    assert "estatus=enPromocion" in url
    assert "fechaAltaDesde=2015-01-01" in url
    assert "fechaAltaHasta=2026-10-19" in url
    assert "conFoto=1" in url
    assert "tipoOperacion=venta" in url
    assert "mesCierre" not in url


def test_main_corre_ofrecidas_aunque_vendidas_aborte(monkeypatch):
    llamadas = []

    def vendidas_rota():
        raise Exception("fallo vendidas")

    monkeypatch.setattr(sys, "argv", ["actualizar_cerradas.py"])
    monkeypatch.setattr(ac, "actualizar_vendidas", vendidas_rota)
    monkeypatch.setattr(ac, "actualizar_ofrecidas", lambda: llamadas.append("ofrecidas"))
    ac.main()
    assert llamadas == ["ofrecidas"]


def test_main_solo_ofrecidas(monkeypatch):
    llamadas = []
    monkeypatch.setattr(sys, "argv", ["actualizar_cerradas.py", "--solo-ofrecidas"])
    monkeypatch.setattr(ac, "actualizar_vendidas", lambda: llamadas.append("vendidas"))
    monkeypatch.setattr(ac, "actualizar_ofrecidas", lambda: llamadas.append("ofrecidas"))
    ac.main()
    assert llamadas == ["ofrecidas"]


def test_vendidas_publicadas_lee_head(monkeypatch):
    comandos = []

    class R:
        stdout = '[{"id": 7, "pr": 90000.0, "dir": "Peña 1"}]'.encode("utf-8")

    def falso_run(cmd, **kw):
        comandos.append((cmd, kw))
        return R()

    monkeypatch.setattr(ac.subprocess, "run", falso_run)
    assert ac.vendidas_publicadas() == {7: {"id": 7, "pr": 90000.0, "dir": "Peña 1"}}
    cmd, kw = comandos[0]
    assert cmd == ["git", "show", "HEAD:data/cerradas.json"]
    assert kw["cwd"] == ac.REPO_DIR and kw["check"] is True and kw["capture_output"] is True


def test_git_ofrecidas_stagea_solo_su_archivo(monkeypatch):
    comandos = []

    class R:
        returncode = 1  # hay diff

    monkeypatch.setattr(ac.subprocess, "run", lambda cmd, **kw: comandos.append(cmd) or R())
    ac.git_commit_push_ofrecidas("msg")
    assert ["git", "add", "data/ofrecidas.json"] in comandos
    assert ["git", "push"] in comandos
    for cmd in comandos:
        assert "-A" not in cmd and "." not in cmd and "--all" not in cmd


def test_git_ofrecidas_sin_diff_no_commitea(monkeypatch):
    comandos = []

    class R:
        returncode = 0

    monkeypatch.setattr(ac.subprocess, "run", lambda cmd, **kw: comandos.append(cmd) or R())
    ac.git_commit_push_ofrecidas("msg")
    assert not any(c[:2] in (["git", "commit"], ["git", "push"]) for c in comandos)


def test_actualizar_ofrecidas_excepcion_queda_como_alerta(monkeypatch, tmp_path):
    import openpyxl

    resumen = tmp_path / "resumen.xlsx"
    tmp_export = tmp_path / "export.xlsx"
    tmp_export.write_bytes(b"x")

    def sin_navegador():
        raise RuntimeError("sin red")

    monkeypatch.setattr(ac, "RESUMEN_PATH", resumen)
    monkeypatch.setattr(ac, "DESCARGA_OFRECIDAS_TMP", tmp_export)
    monkeypatch.setattr(ac, "sync_playwright", sin_navegador)
    ac.actualizar_ofrecidas()  # no debe levantar
    celdas = [c for f in openpyxl.load_workbook(resumen)["Ofrecidas"].iter_rows(values_only=True) for c in f]
    assert any(isinstance(c, str) and c.startswith("ALERTA - ") and "sin red" in c for c in celdas)
    assert not tmp_export.exists()
