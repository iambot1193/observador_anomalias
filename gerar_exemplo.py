"""Gera exemplo/index.html com dados 100% fictícios, sem tocar em API ou arquivos reais.

Uso: python gerar_exemplo.py
Serve só pra visualizar o layout do dashboard sem precisar de credenciais.
"""
import json
import os
from datetime import datetime, timedelta

from gerar import TEMPLATE

AQUI = os.path.dirname(__file__)
SAIDA = os.path.join(AQUI, "exemplo", "index.html")
FMT = "%d-%m-%Y %H:%M:%S"

CODIGOS = ["EQP-01", "EQP-02", "EQP-03", "EQP-04", "EQP-05", "EQP-06"]
CLIENTES = ["Cliente Demo A", "Cliente Demo B", "Cliente Demo C"]
OPERADORAS = ["Claro", "Vivo", "Tim"]


def equipamentos_ficticios():
    agora = datetime(2026, 8, 3, 9, 0)
    ranking = [
        {"codigo": c, "ciclos": n, "resets": max(n // 4, 0),
         "primeiro": (agora - timedelta(hours=n)).strftime(FMT), "ultimo": agora.strftime(FMT)}
        for c, n in zip(CODIGOS, [42, 31, 20, 14, 9, 3])
    ]
    timeline = [
        {"t": (agora - timedelta(minutes=i)).strftime(FMT), "n": n, "codigos": CODIGOS[:n]}
        for i, n in enumerate([0, 1, 1, 2, 0, 0, 3, 1, 0, 1])
    ]
    timeline.reverse()
    dist = [{"n": 1, "ciclos": 5}, {"n": 2, "ciclos": 2}, {"n": 3, "ciclos": 1}]
    return {
        "gerado_em": agora.strftime(FMT),
        "ranking": ranking,
        "timeline": timeline,
        "dist": dist,
        "totais": {
            "equipamentos": len(ranking),
            "ciclosOffline": sum(r["ciclos"] for r in ranking),
            "resets": sum(r["resets"] for r in ranking),
            "pico": max(p["n"] for p in timeline),
        },
    }


def chips_ficticios():
    lista = []
    for i in range(1, 13):
        pct = round([8.5, 12.0, 55.3, 61.0, 92.4, 30.1, 45.0, 105.2, 3.4, 70.8, 15.6, 98.1][i - 1], 1)
        status = ["atualizado", "atualizado", "atrasado", "sem_comunicacao"][i % 4]
        lista.append({
            "codigo": f"DEMO-{i:03d}",
            "cliente": CLIENTES[i % len(CLIENTES)],
            "operadora": OPERADORAS[i % len(OPERADORAS)],
            "iccid": f"8955000000000{i:06d}",
            "telefone": f"55119999{i:05d}",
            "cidade": ["Cidade A", "Cidade B", "Cidade C"][i % 3],
            "conectado": status == "atualizado",
            "ultimaComunicacao": "2026-08-03T08:00:00.000-03:00",
            "status": status,
            "critico": status == "sem_comunicacao",
            "franquiaMb": 100.0,
            "consumidoMb": round(pct, 1),
            "restanteMb": round(max(100 - pct, 0), 1),
            "pct": pct,
        })
    conectados = sum(1 for x in lista if x["conectado"])
    return {
        "lista": lista,
        "totais": {
            "chips": len(lista),
            "conectados": conectados,
            "desconectados": len(lista) - conectados,
            "atualizados": sum(1 for x in lista if x["status"] == "atualizado"),
            "atrasados": sum(1 for x in lista if x["status"] == "atrasado"),
            "semComunicacao": sum(1 for x in lista if x["status"] == "sem_comunicacao"),
            "criticos": sum(1 for x in lista if x["critico"]),
            "consumidoMb": round(sum(x["consumidoMb"] for x in lista), 1),
            "franquiaMb": round(sum(x["franquiaMb"] for x in lista), 1),
        },
    }


def main():
    completo = equipamentos_ficticios()
    payload = {
        "observacao": {"desde": "27-07-2026 08:00:00", "observado": "6d 22.0h", "perdido": "0.5h", "perdidoHoras": 0.5},
        "completo": completo,
        "filtrado": completo,
        "chips": chips_ficticios(),
    }
    html = TEMPLATE.replace("__DADOS__", json.dumps(payload, ensure_ascii=False))
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Gerado: {SAIDA}")


if __name__ == "__main__":
    main()
