"""Testa a agregação — o único trecho não trivial de gerar.py.

Roda sem xlsx, sem rede, sem credenciais.
"""
from datetime import datetime, timedelta

from gerar import agregar, calcular_perdido, fmt_horas


def test_agregar():
    # valida a agregação — simultaneidade e contagem por código.
    linhas = [
        ("EQP-01", "20-07-2026 10:00:00", "sim"),
        ("EQP-02", "20-07-2026 10:00:00", "nao"),   # 2 juntos neste ciclo
        ("EQP-01", "20-07-2026 10:01:00", "nao"),    # só EQP-01
    ]
    ciclos = {"20-07-2026 10:00:00", "20-07-2026 10:01:00", "20-07-2026 10:02:00"}
    d = agregar(linhas, ciclos)
    assert d["totais"]["equipamentos"] == 2, d["totais"]
    assert d["totais"]["ciclosOffline"] == 3
    assert d["totais"]["resets"] == 1
    assert d["totais"]["pico"] == 2                  # 2 offline ao mesmo tempo
    assert [p["n"] for p in d["timeline"]] == [2, 1, 0]  # inclui o ciclo vazio
    assert d["ranking"][0]["codigo"] == "EQP-01" and d["ranking"][0]["ciclos"] == 2


def test_agregar_cruza_mes():
    # cruza mês pra pegar regressão de comparar "registrada" como string
    # ("01-08" < "31-07" como texto, mesmo Ago vindo depois de Jul) - bug real, corrigido.
    linhas_mes = [
        ("EQP-09", "31-07-2026 23:57:52", "nao"),
        ("EQP-09", "01-08-2026 00:01:38", "nao"),
    ]
    d = agregar(linhas_mes, set())
    r = d["ranking"][0]
    assert r["primeiro"] == "31-07-2026 23:57:52", r
    assert r["ultimo"] == "01-08-2026 00:01:38", r


def test_fmt_horas():
    assert fmt_horas(208.2) == "8d 16h 12m", fmt_horas(208.2)
    assert fmt_horas(88.0) == "3d 16h", fmt_horas(88.0)
    assert fmt_horas(16.2) == "16h 12m", fmt_horas(16.2)


def test_calcular_perdido():
    base = datetime(2026, 1, 1)
    ts = [base, base + timedelta(minutes=1), base + timedelta(minutes=2),
          base + timedelta(hours=2), base + timedelta(hours=2, minutes=1)]
    perdido, total, inicio, fim = calcular_perdido(ts, limite_min=10)
    assert perdido == timedelta(hours=1, minutes=58), perdido
    assert total == timedelta(hours=2, minutes=1)


if __name__ == "__main__":
    test_agregar()
    test_agregar_cruza_mes()
    test_fmt_horas()
    test_calcular_perdido()
    print("ok")
