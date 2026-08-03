"""Lê registro_offline_view.xlsx (monitoramento de equipamentos) e a API DATATEM
(status/consumo de SIM cards) e gera um dashboard estático (index.html) com dois painéis.

Uso:
  1. cp .env.example .env  (preencha com suas credenciais DATATEM)
  2. Coloque seu registro_offline_view.xlsx e monitor_fundo.log na raiz do projeto
     (ou ajuste XLSX/LOG abaixo)
  3. python gerar.py  ->  abre index.html no navegador

Rode de novo pra atualizar depois que o monitor gravar mais dados.
"""
import base64
import json
import os
import re
import time
from collections import Counter, defaultdict
from datetime import datetime, timedelta

import requests
from dotenv import dotenv_values

AQUI = os.path.dirname(__file__)
XLSX = os.path.join(AQUI, "registro_offline_view.xlsx")
LOG = os.path.join(AQUI, "monitor_fundo.log")
SAIDA = os.path.join(AQUI, "index.html")
ENV_CHIPS = os.path.join(AQUI, ".env")
TOKEN_CACHE = os.path.join(AQUI, ".token_cache.json")
HIST_CHIPS = os.path.join(AQUI, "chips_historico.jsonl")
FMT = "%d-%m-%Y %H:%M:%S"
DIVISOR_RE = re.compile(r"-{2,}\s*(\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}:\d{2})\s*-{2,}")
TS_RE = re.compile(r"^\[(\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2})\]", re.MULTILINE)
GAP_LIMITE_MIN = 10  # ponytail: acima disso considera que o script de monitoramento parou


def carregar():
    from openpyxl import load_workbook
    ws = load_workbook(XLSX, read_only=True).active
    linhas, ciclos = [], set()  # ciclos = todos os timestamps de ciclo (divisores)
    for codigo, data_desat, campo, registrada, reset in ws.iter_rows(min_row=2, values_only=True):
        if codigo is None:
            continue
        m = DIVISOR_RE.match(str(codigo))
        if m:
            ciclos.add(m.group(1))
            continue
        if registrada:
            linhas.append((codigo, registrada, str(reset or "")))
    return linhas, ciclos


def agregar(linhas, ciclos):
    por_codigo = defaultdict(lambda: {"ciclos": 0, "resets": 0, "primeiro": None, "ultimo": None})
    por_ciclo = defaultdict(set)  # registrada -> set de codigos offline naquele ciclo
    for codigo, registrada, reset in linhas:
        d = por_codigo[codigo]
        d["ciclos"] += 1
        if reset.strip().lower() == "sim":
            d["resets"] += 1
        if d["primeiro"] is None or registrada < d["primeiro"]:
            d["primeiro"] = registrada
        if d["ultimo"] is None or registrada > d["ultimo"]:
            d["ultimo"] = registrada
        por_ciclo[registrada].add(codigo)

    ranking = sorted(
        ({"codigo": c, **v} for c, v in por_codigo.items()),
        key=lambda x: x["ciclos"], reverse=True,
    )

    # eixo do tempo = união de divisores + ciclos com offline, ordenado cronologicamente
    todos = sorted(ciclos | set(por_ciclo), key=lambda s: datetime.strptime(s, FMT))
    timeline = [
        {"t": ts, "n": len(por_ciclo.get(ts, ())), "codigos": sorted(por_ciclo.get(ts, ()))}
        for ts in todos
    ]
    pico = max((p["n"] for p in timeline), default=0)
    # quantos ciclos tiveram N simultâneos (distribuição da simultaneidade)
    dist = Counter(p["n"] for p in timeline if p["n"] > 0)

    return {
        "gerado_em": datetime.now().strftime(FMT),
        "ranking": ranking,
        "timeline": timeline,
        "dist": [{"n": n, "ciclos": dist[n]} for n in sorted(dist)],
        "totais": {
            "equipamentos": len(por_codigo),
            "ciclosOffline": sum(v["ciclos"] for v in por_codigo.values()),
            "resets": sum(v["resets"] for v in por_codigo.values()),
            "pico": pico,
        },
    }


def fmt_horas(h):
    if h < 24:
        return f"{h:.1f}h"
    d = int(h // 24)
    return f"{d}d {h - d*24:.1f}h"


def calcular_perdido(timestamps, limite_min=GAP_LIMITE_MIN):
    """Soma os intervalos entre ciclos maiores que limite_min (= o script parou de rodar)."""
    ts = sorted(set(timestamps))
    if len(ts) < 2:
        return timedelta(0), timedelta(0), None, None
    limite = timedelta(minutes=limite_min)
    perdido = sum((b - a for a, b in zip(ts, ts[1:]) if b - a > limite), timedelta(0))
    return perdido, ts[-1] - ts[0], ts[0], ts[-1]


def tempo_observacao():
    try:
        with open(LOG, encoding="utf-8", errors="replace") as f:
            texto = f.read()
    except FileNotFoundError:
        return {"desde": None, "observado": "0.0h", "perdido": "0.0h", "perdidoHoras": 0}
    timestamps = (datetime.strptime(t, FMT) for t in TS_RE.findall(texto))
    perdido, total, inicio, fim = calcular_perdido(timestamps)
    observado = total - perdido
    return {
        "desde": inicio.strftime(FMT) if inicio else None,
        "observado": fmt_horas(observado.total_seconds() / 3600),
        "perdido": fmt_horas(perdido.total_seconds() / 3600),
        "perdidoHoras": round(perdido.total_seconds() / 3600, 2),
    }


EXCLUIVEIS = set()  # ponytail: códigos a excluir da visão "filtrada" (ex: {"EQP-16"})


def _login_datatem(env, base_url):
    senha_b64 = base64.b64encode(env["DATATEM_PASSWORD"].encode()).decode()
    resp = requests.post(
        f"{base_url}/authorization/api/v1/auth/login",
        json={
            "username": env["DATATEM_USER"], "password": senha_b64,
            "grantType": "password", "clientId": env["DATATEM_CLIENT_ID"],
            "clientSecret": env["DATATEM_CLIENT_SECRET"],
        }, headers={"User-Agent": "Mozilla/5.0"}, timeout=15,
    ).json()
    return resp["authenticationToken"]


def _refresh_datatem(base_url, refresh_token):
    resp = requests.post(
        f"{base_url}/authorization/api/v1/auth/refresh_token",
        json={"refreshToken": refresh_token, "grantType": "refresh_token"},
        headers={"User-Agent": "Mozilla/5.0"}, timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def _obter_token(env, base_url):
    """Reaproveita o token salvo em cache; só loga de novo quando expira (ou refresh falha)."""
    cache = {}
    try:
        with open(TOKEN_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass

    agora = time.time()
    if cache.get("accessToken") and cache.get("expiresAt", 0) > agora + 60:
        return cache["accessToken"]

    if cache.get("refreshToken"):
        try:
            novo = _refresh_datatem(base_url, cache["refreshToken"])
            cache = {
                "accessToken": novo["accessToken"], "refreshToken": novo["refreshToken"],
                "expiresAt": agora + novo.get("expiresIn", 3600),
            }
            with open(TOKEN_CACHE, "w", encoding="utf-8") as f:
                json.dump(cache, f)
            return cache["accessToken"]
        except Exception:
            pass  # refresh token também expirou/inválido -> loga do zero abaixo

    auth = _login_datatem(env, base_url)
    cache = {
        "accessToken": auth["accessToken"], "refreshToken": auth["refreshToken"],
        "expiresAt": agora + auth.get("expiresIn", 3600),
    }
    with open(TOKEN_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    return cache["accessToken"]


def _gravar_historico(lista, gerado_em):
    # ponytail: cresce ~1 linha/chip por execução; se virar gigabyte, rotacionar por mês.
    with open(HIST_CHIPS, "a", encoding="utf-8") as f:
        for x in lista:
            f.write(json.dumps({
                "ts": gerado_em, "iccid": x["iccid"], "codigo": x["codigo"], "cliente": x["cliente"],
                "status": x["status"], "conectado": x["conectado"], "pct": x["pct"],
            }, ensure_ascii=False) + "\n")


def carregar_chips():
    """Busca status de conexão + consumo de todos os simcards na API DATATEM."""
    env = dotenv_values(ENV_CHIPS)
    base_url = env.get("DATATEM_API_URL", "https://app-gateway.brcaptura.com.br")
    token = _obter_token(env, base_url)
    headers = {"Authorization": f"Bearer {token}", "User-Agent": "Mozilla/5.0"}

    chips, page = [], 1
    while True:
        resp = requests.get(
            f"{base_url}/simcard/connection/v1/connections",
            params={"pageSize": 100, "pageNumber": page}, headers=headers, timeout=15,
        )
        if resp.status_code in (401, 403) and page == 1:
            # ponytail: token em cache pode ter sido revogado no servidor mesmo sem ter expirado -> força login novo
            auth = _login_datatem(env, base_url)
            with open(TOKEN_CACHE, "w", encoding="utf-8") as f:
                json.dump({"accessToken": auth["accessToken"], "refreshToken": auth["refreshToken"],
                           "expiresAt": time.time() + auth.get("expiresIn", 3600)}, f)
            headers["Authorization"] = f"Bearer {auth['accessToken']}"
            continue
        r = resp.json()
        content = r.get("content", [])
        if not content:
            break
        chips.extend(content)
        if len(content) < 100:
            break
        page += 1

    def custom_field(c, nome):
        for v in (c.get("customField") or {}).get("values") or []:
            if v.get("name") == nome and v.get("value"):
                return v["value"]
        return None

    def horas_desde(iso_str):
        if not iso_str:
            return None
        dt = datetime.fromisoformat(iso_str)
        agora = datetime.now(dt.tzinfo) if dt.tzinfo else datetime.now()
        return (agora - dt).total_seconds() / 3600

    def status_bucket(horas):
        if horas is None:
            return "sem_comunicacao"
        if horas <= 2:
            return "atualizado"
        if horas <= 24:
            return "atrasado"
        return "sem_comunicacao"

    lista = []
    for c in chips:
        franquia = int(c.get("franchise") or 0)
        consumido = int(c.get("cycletotalbytes") or 0)
        horas = horas_desde(c.get("lastAccounting"))
        lista.append({
            "codigo": custom_field(c, "device") or c.get("simcard"),
            "cliente": (custom_field(c, "description2") or c.get("clientName") or "").strip().upper() or None,
            "operadora": c.get("operatorName"),
            "iccid": c.get("iccid"),
            "telefone": c.get("phonenumber"),
            "cidade": c.get("city"),
            "conectado": c.get("connectivitystatus") == "CONNECTED",
            "ultimaComunicacao": c.get("lastAccounting"),
            "status": status_bucket(horas),
            "critico": horas is None or horas > 48,
            "franquiaMb": round(franquia / 1_000_000, 1),
            "consumidoMb": round(consumido / 1_000_000, 1),
            "restanteMb": round(max(franquia - consumido, 0) / 1_000_000, 1),
            "pct": round(c.get("cycleTotalBytesPct") or 0, 1),
        })
    lista.sort(key=lambda x: x["pct"], reverse=True)
    conectados = sum(1 for x in lista if x["conectado"])
    _gravar_historico(lista, datetime.now().strftime(FMT))
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
    linhas, ciclos = carregar()
    completo = agregar(linhas, ciclos)
    filtrado = agregar([l for l in linhas if l[0] not in EXCLUIVEIS], ciclos)
    obs = tempo_observacao()
    try:
        chips = carregar_chips()
    except Exception as e:
        chips = {"erro": str(e)}
    payload = {"observacao": obs, "completo": completo, "filtrado": filtrado, "chips": chips}
    html = TEMPLATE.replace("__DADOS__", json.dumps(payload, ensure_ascii=False))
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Gerado: {SAIDA}")
    print(f"  {completo['totais']['equipamentos']} equipamentos, "
          f"{completo['totais']['ciclosOffline']} ciclos offline, "
          f"pico de {completo['totais']['pico']} simultâneos. "
          f"Observado {obs['observado']}, perdido {obs['perdido']}.")


# ponytail: Chart.js via CDN. Se precisar 100% offline, baixar chart.umd.min.js
# pra pasta e trocar o src abaixo.
TEMPLATE = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Observador de Anomalias</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{ --border:rgba(255,255,255,.10) }
  *{box-sizing:border-box}
  body{margin:0;background:#05060A;color:#F5F6FA;
    font-family:system-ui,-apple-system,"Segoe UI",sans-serif;line-height:1.45}
  .wrap{max-width:1100px;margin:0 auto;padding:28px 20px 60px}
  h1{font-size:22px;margin:0 0 2px}
  .sub{font-size:13px;margin:0 0 20px;opacity:.65}
  .tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:26px}
  .tile{border-radius:12px;padding:14px 16px}
  .tile .v{font-size:28px;font-weight:700;letter-spacing:-.5px}
  .tile .k{font-size:12px;opacity:.65;margin-top:2px}
  .card{border-radius:14px;padding:18px 18px 8px;margin-bottom:22px}
  .card h2{font-size:15px;margin:0 0 2px}
  .card p{font-size:12.5px;opacity:.65;margin:0 0 14px}
  .chart{position:relative;height:340px}
  .chart.short{height:250px}
  table{width:100%;border-collapse:collapse;font-size:13px;font-variant-numeric:tabular-nums}
  th,td{text-align:right;padding:8px 10px;border-bottom:1px solid var(--border)}
  th:first-child,td:first-child{text-align:left;font-variant-numeric:normal}
  th{font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;opacity:.55}
  details summary{cursor:pointer;font-size:13px;margin-top:4px;opacity:.7}
  .chk{display:inline-flex;align-items:center;gap:8px;font-size:14px;opacity:.75;
    margin:-8px 0 22px;cursor:pointer}
  .chk input{width:20px;height:20px;cursor:pointer}
  .tabs{display:flex;gap:8px;margin-bottom:22px;border-bottom:1px solid var(--border)}
  .tabBtn{padding:9px 16px;border:none;background:transparent;color:inherit;opacity:.55;
    font-size:13.5px;font-weight:600;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
  .tabBtn.active{opacity:1;border-bottom-color:#2563EB}

  /* Dark Glassmorphism — mesmo tema nas duas abas (cobre a área toda) */
  body.theme-equip,body.theme-chips{background:linear-gradient(160deg, #171233 0%, #0D0F18 45%, #05060A 100%) fixed}
  #tabEquip .card,#tabEquip .tile,#tabChips .card,#tabChips .tile{
    background:rgba(255,255,255,.04);backdrop-filter:blur(16px);
    -webkit-backdrop-filter:blur(16px);border:1px solid rgba(255,255,255,.09)}
  #tabEquip .tile.alerta .v,#tabChips .tile.alerta .v{color:#FF4D4D}
  #tabChips .tile.ok .v{color:#00E676}

  .biRow4{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
  @media (max-width:980px){.biRow4{grid-template-columns:1fr 1fr}}
  @media (max-width:560px){.biRow4{grid-template-columns:1fr}}

  .filtros{display:flex;gap:10px;flex-wrap:wrap;margin:22px 0 14px}
  .filtros input,.filtros select{padding:9px 12px;border-radius:8px;border:1px solid var(--border);
    background:rgba(255,255,255,.05);color:inherit;font-size:13px}
  .filtros input{flex:1;min-width:220px}
  .statusChips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
  .chipBtn{padding:6px 14px;border-radius:20px;border:1px solid var(--border);background:rgba(255,255,255,.05);
    color:inherit;opacity:.65;font-size:12.5px;cursor:pointer}
  .chipBtn.active{opacity:1;border-color:#2563EB}
  .chipBtn b{margin-left:4px}

  .filtrosAtivos{display:none;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:14px}
  .filtroBadge{background:rgba(37,99,235,.18);border:1px solid rgba(37,99,235,.4);border-radius:20px;
    padding:5px 10px;font-size:12px;cursor:pointer}
  .filtroBadge b{margin-left:5px;opacity:.7}
  #btnLimparFiltros{border:1px solid var(--border);background:transparent;color:inherit;opacity:.75;
    border-radius:20px;padding:5px 12px;font-size:12px;cursor:pointer}

  .badge{padding:2px 9px;border-radius:10px;font-size:11.5px;font-weight:600;white-space:nowrap;
    display:inline-flex;align-items:center;gap:5px}
  .badge .dot{width:7px;height:7px;border-radius:50%;display:inline-block}
  .badge.atualizado{background:rgba(0,230,118,.14);color:#00E676}
  .badge.atualizado .dot{background:#00E676}
  .badge.atrasado{background:rgba(255,193,7,.16);color:#FFC107}
  .badge.atrasado .dot{background:#FFC107}
  .badge.sem_comunicacao{background:rgba(255,77,77,.14);color:#FF4D4D}
  .badge.sem_comunicacao .dot{background:#FF4D4D}
  tr.linhaCritica{background:rgba(255,77,77,.07)}

  .tableToolbar{display:flex;gap:10px;align-items:center;margin:0 0 10px;position:relative}
  .tbBtn{padding:8px 14px;border-radius:8px;border:1px solid var(--border);background:rgba(255,255,255,.05);
    color:inherit;font-size:12.5px;cursor:pointer}
  .colunasMenu{display:none;position:absolute;top:38px;left:0;z-index:20;background:#12141c;
    border:1px solid var(--border);border-radius:10px;padding:10px 14px;min-width:210px;
    box-shadow:0 10px 30px rgba(0,0,0,.4)}
  .colunasMenu label{display:flex;gap:8px;align-items:center;font-size:12.5px;padding:4px 0;cursor:pointer}

  .tableWrap{margin-left:20px;background:#12141c;border:1px solid rgba(255,255,255,.09);
    border-radius:14px;max-height:640px;overflow:auto;padding:0}
  .tableWrap table{margin:0}
  .tableWrap td,.tableWrap th{padding:9px 10px 9px 14px}
  .tableWrap th{position:sticky;top:0;background:#12141c;cursor:pointer;user-select:none;z-index:5}
  .tableWrap th[data-col]:hover{opacity:.85}
  .tableWrap th[data-col]::after{content:'⇅';margin-left:5px;opacity:.3;font-size:10px}
  .tableWrap th.sortAsc::after{content:'▲';opacity:1;color:#2563EB}
  .tableWrap th.sortDesc::after{content:'▼';opacity:1;color:#2563EB}

  .rowMenu{position:relative}
  .rowMenu summary{list-style:none;cursor:pointer;text-align:center;opacity:.6;font-size:16px}
  .rowMenu summary::-webkit-details-marker{display:none}
  .rowMenu[open] summary{opacity:1}
  .rowMenu .menu{position:absolute;right:0;background:#1a1c26;border:1px solid var(--border);
    border-radius:8px;padding:6px;min-width:160px;z-index:15;box-shadow:0 10px 30px rgba(0,0,0,.4)}
  .rowMenu .menu button{display:block;width:100%;text-align:left;background:none;border:none;color:inherit;
    padding:7px 10px;font-size:12.5px;border-radius:6px;cursor:pointer}
  .rowMenu .menu button:hover{background:rgba(255,255,255,.06)}

  #modalDetalhes{display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:100;
    align-items:center;justify-content:center;padding:20px}
  #modalDetalhes .box{background:#12141c;border:1px solid var(--border);border-radius:14px;
    padding:22px;max-width:480px;width:100%;max-height:80vh;overflow:auto}
  #modalDetalhes h3{margin:0 0 14px;font-size:16px}
  #modalDetalhes .linha{display:flex;justify-content:space-between;gap:12px;padding:6px 0;
    border-bottom:1px solid var(--border);font-size:13px}
  #modalDetalhes .fechar{margin-top:16px;padding:8px 14px;border-radius:8px;border:1px solid var(--border);
    background:transparent;color:inherit;cursor:pointer}
</style>
</head>
<body>
<div class="wrap">
  <div class="tabs" id="tabs">
    <button class="tabBtn active" data-tab="equip">Comunicação dos equipamentos</button>
    <button class="tabBtn" data-tab="chips">Chips e consumo</button>
  </div>

  <div id="tabEquip" class="tabPanel">
    <h1>Comunicação dos equipamentos</h1>
    <p class="sub">A planilha registra os equipamentos <b>fora do ar</b> a cada ciclo (~1&nbsp;min).
       Mais ciclos = mais tempo sem comunicar. Observando desde <span id="desde"></span>.
       Gerado em <span id="ger"></span>.</p>

    <label class="chk"><input type="checkbox" id="chkSem16"> Remover exceções das métricas visuais</label>

    <div class="tiles" id="tiles"></div>

    <div class="card">
      <h2>Quanto cada equipamento ficou fora do ar</h2>
      <p>Total de ciclos registrados por código (cada ciclo ≈ 1 minuto offline).</p>
      <div class="chart" id="wrapRank"><canvas id="rank"></canvas></div>
    </div>

    <div class="card">
      <h2>Quantos ficaram sem comunicar ao mesmo tempo</h2>
      <p>Cada ciclo grava <b>uma linha por equipamento</b> offline — todos entram, não só um.
         Aqui: nº de equipamentos simultaneamente fora do ar ao longo do tempo.</p>
      <div class="chart" id="wrapTl"><canvas id="tl"></canvas></div>
    </div>

    <div class="card">
      <h2>Distribuição da simultaneidade</h2>
      <p>Em quantos ciclos houve 1, 2, 3… equipamentos offline ao mesmo tempo.</p>
      <div class="chart short" id="wrapDist"><canvas id="dist"></canvas></div>
    </div>

    <details open>
      <summary>Ver tabela</summary>
      <table><thead><tr><th>Código</th><th>Ciclos offline</th><th>Resets enviados</th>
        <th>Primeiro</th><th>Último</th></tr></thead><tbody id="tbody"></tbody></table>
    </details>
  </div>

  <div id="tabChips" class="tabPanel" style="display:none">
    <h1>Chips e consumo</h1>
    <p class="sub" id="chipsErro" style="display:none">Não foi possível carregar os dados da API.</p>
    <p class="sub" id="chipsSync"></p>

    <div class="tiles" id="kpiRow"></div>

    <div class="biRow4">
      <div class="card">
        <h2>Status de conexão</h2>
        <p>Clique numa fatia pra filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsStatusChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Quantidade por cliente</h2>
        <p>Clique numa barra pra filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsClienteChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Quantidade por operadora</h2>
        <p>Clique numa fatia pra filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsOperadoraChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Faixa de consumo</h2>
        <p>Clique numa fatia pra filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsConsumoChart"></canvas></div>
      </div>
    </div>

    <div class="filtros">
      <input id="fBusca" type="text" placeholder="Buscar por código, ICCID, telefone, cliente ou cidade...">
      <select id="fCliente"><option value="">Todos os clientes</option></select>
      <select id="fOperadora"><option value="">Todas as operadoras</option></select>
    </div>
    <div class="statusChips" id="statusChips">
      <button class="chipBtn active" data-status="">Todos <b id="cntTodos">0</b></button>
      <button class="chipBtn" data-status="atualizado">Atualizados <b id="cntAtualizado">0</b></button>
      <button class="chipBtn" data-status="atrasado">Atrasados <b id="cntAtrasado">0</b></button>
      <button class="chipBtn" data-status="sem_comunicacao">Sem comunicação <b id="cntSemComunicacao">0</b></button>
    </div>
    <div class="filtrosAtivos" id="filtrosAtivos"></div>

    <div class="tableToolbar">
      <button class="tbBtn" id="btnColunas">Colunas ▾</button>
      <div class="colunasMenu" id="colunasMenu"></div>
      <button class="tbBtn" id="btnExportar">Exportar CSV</button>
    </div>

    <div class="tableWrap" id="tableWrap">
      <table id="tabelaChips"><thead><tr>
        <th data-col="cliente" data-type="text">Cliente</th>
        <th data-col="codigo" data-type="text">Código</th>
        <th data-col="iccid" data-type="text">ICCID</th>
        <th data-col="telefone" data-type="text">Telefone</th>
        <th data-col="cidade" data-type="text">Cidade</th>
        <th data-col="operadora" data-type="text">Operadora</th>
        <th data-col="ultimaComunicacao" data-type="date">Última comunicação</th>
        <th data-col="status" data-type="text">Status</th>
        <th data-col="franquiaMb" data-type="num">Franquia</th>
        <th data-col="consumidoMb" data-type="num">Consumido</th>
        <th data-col="restanteMb" data-type="num">Restante</th>
        <th data-col="pct" data-type="num">%</th>
        <th></th>
      </tr></thead>
      <tbody id="tbodyChips"></tbody></table>
    </div>
  </div>
</div>

<div id="modalDetalhes">
  <div class="box">
    <h3>Detalhes do chip</h3>
    <div id="modalCorpo"></div>
    <button class="fechar" id="modalFechar">Fechar</button>
  </div>
</div>

<script>
const PAYLOAD = __DADOS__;
// ponytail: paletas fixas por tema (não dependem de prefers-color-scheme — os dois temas são sempre escuros)
const EQUIP = {blue:'#2563EB', red:'#FF4D4D', amber:'#FFC107', green:'#00E676',
  ink:'#F5F6FA', ink2:'#9CA3AF', muted:'#6B7280', grid:'#20232F'};
const PAL = {blue:'#2563EB', red:'#FF4D4D', amber:'#FFC107', green:'#00E676',
  ink:'#F5F6FA', ink2:'rgba(245,246,250,.62)', muted:'rgba(245,246,250,.42)', grid:'rgba(255,255,255,.08)'};
const INK=EQUIP.ink, INK2=EQUIP.ink2, MUTED=EQUIP.muted, GRID=EQUIP.grid, BLUE=EQUIP.blue, RED=EQUIP.red;
Chart.defaults.font.family = 'system-ui,-apple-system,"Segoe UI",sans-serif';
Chart.defaults.color = '#9CA3AF';
const grid = {color:GRID, drawTicks:false, drawBorder:false};
function hexAlpha(hex, a){
  const n = parseInt(hex.replace('#',''),16);
  return `rgba(${(n>>16)&255},${(n>>8)&255},${n&255},${a})`;
}

document.getElementById('ger').textContent = PAYLOAD.completo.gerado_em;
document.getElementById('desde').textContent = PAYLOAD.observacao.desde || '—';

let charts = {};
function render(D){
  const T = D.totais;
  document.getElementById('tiles').innerHTML = [
    ['Tempo observado', PAYLOAD.observacao.observado, ''],
    ['Tempo perdido (falha do script)', PAYLOAD.observacao.perdido, PAYLOAD.observacao.perdidoHoras>0?'alerta':''],
    ['Equipamentos com falha', T.equipamentos, ''],
    ['Ciclos offline (~min)', T.ciclosOffline, ''],
    ['Resets enviados', T.resets, ''],
    ['Pico simultâneo', T.pico, T.pico>1?'alerta':''],
  ].map(([k,v,c])=>`<div class="tile ${c}"><div class="v">${v}</div><div class="k">${k}</div></div>`).join('');

  Object.values(charts).forEach(c=>c.destroy());

  // Ranking — barras horizontais, uma cor (magnitude), rótulo direto no fim
  charts.rank = new Chart(document.getElementById('rank'), {
    type:'bar',
    data:{labels:D.ranking.map(r=>r.codigo),
      datasets:[{data:D.ranking.map(r=>r.ciclos), backgroundColor:BLUE,
        borderRadius:4, borderSkipped:false, barThickness:'flex', maxBarThickness:26}]},
    options:{indexAxis:'y', maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>` ${c.raw} ciclos (~${c.raw} min)`}}},
      scales:{x:{beginAtZero:true, grid, ticks:{color:MUTED}},
        y:{grid:{display:false}, ticks:{color:INK}}}}
  });
  document.getElementById('wrapRank').style.height = Math.max(260, D.ranking.length*34+40)+'px';

  // Timeline — área, uma série (título já nomeia), com stepped pra ler ciclo a ciclo
  charts.tl = new Chart(document.getElementById('tl'), {
    type:'line',
    data:{labels:D.timeline.map(p=>p.t.slice(11,16)),
      datasets:[{data:D.timeline.map(p=>p.n), borderColor:BLUE,
        backgroundColor:'rgba(42,120,214,.14)', fill:true, stepped:true,
        borderWidth:2, pointRadius:0, pointHoverRadius:5}]},
    options:{maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins:{legend:{display:false},
        tooltip:{callbacks:{
          title:it=>D.timeline[it[0].dataIndex].t,
          label:c=>{const p=D.timeline[c.dataIndex];
            return p.n?` ${p.n} offline: ${p.codigos.join(', ')}`:' tudo comunicando'}}}},
      scales:{y:{beginAtZero:true, ticks:{precision:0, color:MUTED, stepSize:1}, grid},
        x:{grid:{display:false}, ticks:{color:MUTED, maxTicksLimit:12, autoSkip:true}}}}
  });

  // Distribuição da simultaneidade — barras verticais
  charts.dist = new Chart(document.getElementById('dist'), {
    type:'bar',
    data:{labels:D.dist.map(d=>d.n+(d.n>1?' juntos':' sozinho')),
      datasets:[{data:D.dist.map(d=>d.ciclos), backgroundColor:D.dist.map(d=>d.n>1?RED:BLUE),
        borderRadius:4, borderSkipped:false, maxBarThickness:60}]},
    options:{maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>` ${c.raw} ciclos`}}},
      scales:{y:{beginAtZero:true, ticks:{precision:0, color:MUTED}, grid},
        x:{grid:{display:false}, ticks:{color:INK}}}}
  });

  document.getElementById('tbody').innerHTML = D.ranking.map(r=>
    `<tr><td>${r.codigo}</td><td>${r.ciclos}</td><td>${r.resets}</td>
     <td>${r.primeiro||''}</td><td>${r.ultimo||''}</td></tr>`).join('');
}

document.getElementById('chkSem16').addEventListener('change', e=>{
  render(e.target.checked ? PAYLOAD.filtrado : PAYLOAD.completo);
});
render(PAYLOAD.completo);

function ativarTema(tab){
  document.body.classList.toggle('theme-equip', tab==='equip');
  document.body.classList.toggle('theme-chips', tab==='chips');
}
document.querySelectorAll('.tabBtn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tabBtn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tabEquip').style.display = btn.dataset.tab==='equip' ? '' : 'none';
    document.getElementById('tabChips').style.display = btn.dataset.tab==='chips' ? '' : 'none';
    ativarTema(btn.dataset.tab);
  });
});
ativarTema('equip');

if (PAYLOAD.chips && !PAYLOAD.chips.erro) {
  const C = PAYLOAD.chips;
  const STATUS_LABEL = {atualizado:'Atualizado', atrasado:'Atrasado', sem_comunicacao:'Sem comunicação'};
  const COLS = [
    {key:'cliente', label:'Cliente', type:'text'}, {key:'codigo', label:'Código', type:'text'},
    {key:'iccid', label:'ICCID', type:'text'}, {key:'telefone', label:'Telefone', type:'text'},
    {key:'cidade', label:'Cidade', type:'text'}, {key:'operadora', label:'Operadora', type:'text'},
    {key:'ultimaComunicacao', label:'Última comunicação', type:'date'}, {key:'status', label:'Status', type:'text'},
    {key:'franquiaMb', label:'Franquia', type:'num'}, {key:'consumidoMb', label:'Consumido', type:'num'},
    {key:'restanteMb', label:'Restante', type:'num'}, {key:'pct', label:'%', type:'num'},
  ];

  document.getElementById('chipsSync').textContent = 'Dados atualizados em ' + PAYLOAD.completo.gerado_em;

  // --- KPIs ---
  const pctFranquia = C.totais.franquiaMb ? Math.round(C.totais.consumidoMb / C.totais.franquiaMb * 100) : 0;
  document.getElementById('kpiRow').innerHTML = [
    ['Chips ativos', C.totais.chips, ''],
    ['Consumo do mês', (C.totais.consumidoMb/1024).toFixed(2)+' GB', ''],
    ['Crítico (sem comunicação > 48h)', C.totais.criticos, C.totais.criticos>0?'alerta':'ok'],
    ['Franquia consumida', pctFranquia+'%', pctFranquia>=80?'alerta':''],
  ].map(([k,v,c])=>`<div class="tile ${c}"><div class="v">${v}</div><div class="k">${k}</div></div>`).join('');

  // --- BIs (clicáveis) ---
  const porCliente = {}; C.lista.forEach(x=>{ const k=x.cliente||'—'; porCliente[k]=(porCliente[k]||0)+1; });
  const clientesOrd = Object.entries(porCliente).sort((a,b)=>b[1]-a[1]);
  const porOperadora = {}; C.lista.forEach(x=>{ const k=x.operadora||'—'; porOperadora[k]=(porOperadora[k]||0)+1; });
  const operadorasOrd = Object.entries(porOperadora).sort((a,b)=>b[1]-a[1]);
  const faixas = [
    {label:'Abaixo de 50%', min:0, max:50, cor:PAL.green},
    {label:'50% a 100%', min:50, max:100, cor:PAL.amber},
    {label:'Acima de 100%', min:100, max:Infinity, cor:PAL.red},
  ];
  const operadoraCores = [PAL.blue,'#9ec5f4',PAL.green,PAL.amber,PAL.red,'#8a63d2'];

  let filtroStatus = '', filtroCliente = '', filtroOperadora = '', filtroFaixa = null;
  let sortCol = null, sortDir = 0;
  let ultimosFiltrados = [];

  const statusChart = new Chart(document.getElementById('chipsStatusChart'), {
    type:'doughnut',
    data:{labels:['Atualizado','Atrasado','Sem comunicação'],
      datasets:[{data:[C.totais.atualizados, C.totais.atrasados, C.totais.semComunicacao],
        backgroundColor:[PAL.green, PAL.amber, PAL.red], borderWidth:0}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{legend:{position:'bottom', labels:{color:PAL.ink2, boxWidth:12}}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = ['atualizado','atrasado','sem_comunicacao'][els[0].index];
        filtroStatus = filtroStatus===val ? '' : val;
        document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.toggle('active', b.dataset.status===filtroStatus));
        atualizarSelecaoGraficos(); aplicarFiltrosChips();
      }}
  });

  const clienteChart = new Chart(document.getElementById('chipsClienteChart'), {
    type:'bar',
    data:{labels:clientesOrd.map(c=>c[0]),
      datasets:[{data:clientesOrd.map(c=>c[1]), backgroundColor:BLUE, borderRadius:4, borderSkipped:false, maxBarThickness:26}]},
    options:{indexAxis:'y', maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true, ticks:{precision:0, color:PAL.muted}, grid:{color:PAL.grid}},
        y:{grid:{display:false}, ticks:{color:PAL.ink}}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = clientesOrd[els[0].index][0];
        filtroCliente = filtroCliente===val ? '' : val;
        document.getElementById('fCliente').value = filtroCliente;
        atualizarSelecaoGraficos(); aplicarFiltrosChips();
      }}
  });

  const operadoraChart = new Chart(document.getElementById('chipsOperadoraChart'), {
    type:'doughnut',
    data:{labels:operadorasOrd.map(o=>o[0]),
      datasets:[{data:operadorasOrd.map(o=>o[1]), backgroundColor:operadorasOrd.map((o,i)=>operadoraCores[i%operadoraCores.length]), borderWidth:0}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{legend:{position:'bottom', labels:{color:PAL.ink2, boxWidth:12}}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = operadorasOrd[els[0].index][0];
        filtroOperadora = filtroOperadora===val ? '' : val;
        document.getElementById('fOperadora').value = filtroOperadora;
        atualizarSelecaoGraficos(); aplicarFiltrosChips();
      }}
  });

  const consumoChart = new Chart(document.getElementById('chipsConsumoChart'), {
    type:'doughnut',
    data:{labels:faixas.map(f=>f.label), datasets:[{data:faixas.map(f=>C.lista.filter(x=>x.pct>=f.min && x.pct<f.max).length),
      backgroundColor:faixas.map(f=>f.cor), borderWidth:0}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{legend:{position:'bottom', labels:{color:PAL.ink2, boxWidth:12}}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const f = faixas[els[0].index];
        filtroFaixa = (filtroFaixa && filtroFaixa.label===f.label) ? null : f;
        atualizarSelecaoGraficos(); aplicarFiltrosChips();
      }}
  });

  function atualizarSelecaoGraficos(){
    statusChart.data.datasets[0].backgroundColor = [
      ['atualizado',PAL.green],['atrasado',PAL.amber],['sem_comunicacao',PAL.red],
    ].map(([v,cor])=> (!filtroStatus||filtroStatus===v) ? cor : hexAlpha(cor,.25));
    statusChart.update();
    clienteChart.data.datasets[0].backgroundColor = clientesOrd.map(c=>
      (!filtroCliente||filtroCliente===c[0]) ? BLUE : hexAlpha(BLUE,.25));
    clienteChart.update();
    operadoraChart.data.datasets[0].backgroundColor = operadorasOrd.map((o,i)=>{
      const cor = operadoraCores[i%operadoraCores.length];
      return (!filtroOperadora||filtroOperadora===o[0]) ? cor : hexAlpha(cor,.25);
    });
    operadoraChart.update();
    consumoChart.data.datasets[0].backgroundColor = faixas.map(f=>
      (!filtroFaixa||filtroFaixa.label===f.label) ? f.cor : hexAlpha(f.cor,.25));
    consumoChart.update();
  }

  // --- Filtros de texto/select ---
  const clientesUnicos = [...new Set(C.lista.map(x=>x.cliente).filter(Boolean))].sort();
  document.getElementById('fCliente').innerHTML += clientesUnicos.map(c=>`<option value="${c}">${c}</option>`).join('');
  const operadorasUnicas = [...new Set(C.lista.map(x=>x.operadora).filter(Boolean))].sort();
  document.getElementById('fOperadora').innerHTML += operadorasUnicas.map(o=>`<option value="${o}">${o}</option>`).join('');
  document.getElementById('cntTodos').textContent = C.lista.length;
  document.getElementById('cntAtualizado').textContent = C.totais.atualizados;
  document.getElementById('cntAtrasado').textContent = C.totais.atrasados;
  document.getElementById('cntSemComunicacao').textContent = C.totais.semComunicacao;

  // --- Colunas exibir/ocultar ---
  document.getElementById('colunasMenu').innerHTML = COLS.map((c,i)=>
    `<label><input type="checkbox" checked data-idx="${i+1}"> ${c.label}</label>`).join('');
  document.querySelectorAll('#colunasMenu input').forEach(chk=>{
    chk.addEventListener('change', ()=>{
      const idx = chk.dataset.idx, hidden = !chk.checked;
      document.querySelectorAll(`#tabelaChips th:nth-child(${idx}), #tabelaChips td:nth-child(${idx})`)
        .forEach(el=>el.style.display = hidden ? 'none' : '');
    });
  });
  document.getElementById('btnColunas').addEventListener('click', ()=>{
    const m = document.getElementById('colunasMenu');
    m.style.display = m.style.display==='block' ? 'none' : 'block';
  });
  document.addEventListener('click', e=>{
    const menu = document.getElementById('colunasMenu');
    if (menu.style.display==='block' && !menu.contains(e.target) && e.target.id!=='btnColunas') menu.style.display='none';
  });

  // --- Exportar CSV (dos dados filtrados/ordenados na tela) ---
  document.getElementById('btnExportar').addEventListener('click', ()=>{
    const linhas = [COLS.map(c=>c.label).join(';')].concat(ultimosFiltrados.map(x=>
      COLS.map(c=>{
        let v = c.key==='status' ? STATUS_LABEL[x.status] : x[c.key];
        return `"${(v??'').toString().replace(/"/g,'""')}"`;
      }).join(';')));
    const blob = new Blob(['﻿'+linhas.join('\n')], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'chips_consumo.csv'; a.click();
    URL.revokeObjectURL(url);
  });

  // --- Ordenação tri-estado nos cabeçalhos ---
  function ordenar(lista){
    if (!sortCol || sortDir===0) return lista;
    const tipo = COLS.find(c=>c.key===sortCol).type;
    return [...lista].sort((a,b)=>{
      let va=a[sortCol], vb=b[sortCol];
      const aNull = va===null||va===undefined||va==='', bNull = vb===null||vb===undefined||vb==='';
      if (aNull && bNull) return 0;
      if (aNull) return 1;
      if (bNull) return -1;
      if (tipo==='num'){ va=Number(va); vb=Number(vb); }
      else if (tipo==='date'){ va=new Date(va).getTime(); vb=new Date(vb).getTime(); }
      else { va=String(va).toLowerCase(); vb=String(vb).toLowerCase(); }
      if (va<vb) return sortDir===1?-1:1;
      if (va>vb) return sortDir===1?1:-1;
      return 0;
    });
  }
  document.querySelectorAll('#tabelaChips th[data-col]').forEach(th=>{
    th.addEventListener('click', ()=>{
      const col = th.dataset.col;
      if (sortCol !== col) { sortCol = col; sortDir = 1; }
      else { sortDir = sortDir===1?2:(sortDir===2?0:1); if (sortDir===0) sortCol=null; }
      document.querySelectorAll('#tabelaChips th[data-col]').forEach(h=>h.classList.remove('sortAsc','sortDesc'));
      if (sortCol) th.classList.add(sortDir===1?'sortAsc':'sortDesc');
      aplicarFiltrosChips();
    });
  });

  // --- Linha da tabela + ações ---
  function linhaHtml(x){
    return `<tr class="${x.critico?'linhaCritica':''}">
      <td>${x.cliente||''}</td><td>${x.codigo}</td><td>${x.iccid||''}</td><td>${x.telefone||''}</td>
      <td>${x.cidade||''}</td><td>${x.operadora||''}</td>
      <td>${x.ultimaComunicacao ? x.ultimaComunicacao.slice(0,16).replace('T',' ') : ''}</td>
      <td><span class="badge ${x.status}"><span class="dot"></span>${STATUS_LABEL[x.status]}</span></td>
      <td>${x.franquiaMb} MB</td><td>${x.consumidoMb} MB</td><td>${x.restanteMb} MB</td><td>${x.pct}%</td>
      <td><details class="rowMenu"><summary>⋯</summary><div class="menu">
        <button data-acao="copiar" data-iccid="${x.iccid}">Copiar ICCID</button>
        <button data-acao="detalhes" data-iccid="${x.iccid}">Ver detalhes</button>
      </div></details></td></tr>`;
  }
  document.getElementById('tbodyChips').addEventListener('click', e=>{
    const btn = e.target.closest('button[data-acao]');
    if (!btn) return;
    if (btn.dataset.acao==='copiar'){
      navigator.clipboard?.writeText(btn.dataset.iccid);
      const original = btn.textContent;
      btn.textContent = 'Copiado!';
      setTimeout(()=>{ btn.textContent = original; }, 1200);
    } else if (btn.dataset.acao==='detalhes'){
      const item = C.lista.find(x=>x.iccid===btn.dataset.iccid);
      if (item) abrirDetalhes(item);
    }
    btn.closest('details')?.removeAttribute('open');
  });
  function abrirDetalhes(x){
    const campos = [
      ['Cliente', x.cliente], ['Código', x.codigo], ['ICCID', x.iccid], ['Telefone', x.telefone],
      ['Cidade', x.cidade], ['Operadora', x.operadora], ['Última comunicação', x.ultimaComunicacao],
      ['Status', STATUS_LABEL[x.status]], ['Franquia', x.franquiaMb+' MB'], ['Consumido', x.consumidoMb+' MB'],
      ['Restante', x.restanteMb+' MB'], ['% consumido', x.pct+'%'],
    ];
    document.getElementById('modalCorpo').innerHTML = campos.map(([k,v])=>
      `<div class="linha"><span>${k}</span><b>${v??'—'}</b></div>`).join('');
    document.getElementById('modalDetalhes').style.display = 'flex';
  }
  document.getElementById('modalFechar').addEventListener('click', ()=>{
    document.getElementById('modalDetalhes').style.display = 'none';
  });
  document.getElementById('modalDetalhes').addEventListener('click', e=>{
    if (e.target.id === 'modalDetalhes') e.target.style.display = 'none';
  });

  // --- Badges de filtros ativos + limpar tudo ---
  function renderBadges(){
    const badges = [];
    if (filtroCliente) badges.push(['Cliente', filtroCliente, ()=>{filtroCliente='';document.getElementById('fCliente').value='';}]);
    if (filtroOperadora) badges.push(['Operadora', filtroOperadora, ()=>{filtroOperadora='';document.getElementById('fOperadora').value='';}]);
    if (filtroStatus) badges.push(['Status', STATUS_LABEL[filtroStatus], ()=>{
      filtroStatus=''; document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.toggle('active', b.dataset.status===''));
    }]);
    if (filtroFaixa) badges.push(['Consumo', filtroFaixa.label, ()=>{filtroFaixa=null;}]);
    const cont = document.getElementById('filtrosAtivos');
    if (!badges.length){ cont.style.display='none'; cont.innerHTML=''; return; }
    cont.style.display='flex';
    cont.innerHTML = badges.map(([k,v],i)=>`<span class="filtroBadge" data-i="${i}">${k}: ${v} <b>×</b></span>`).join('')
      + `<button id="btnLimparFiltros">Limpar todos os filtros</button>`;
    badges.forEach(([k,v,fn],i)=>{
      cont.querySelector(`.filtroBadge[data-i="${i}"]`).addEventListener('click', ()=>{
        fn(); atualizarSelecaoGraficos(); aplicarFiltrosChips();
      });
    });
    document.getElementById('btnLimparFiltros').addEventListener('click', ()=>{
      filtroStatus=''; filtroCliente=''; filtroOperadora=''; filtroFaixa=null;
      document.getElementById('fCliente').value=''; document.getElementById('fOperadora').value=''; document.getElementById('fBusca').value='';
      document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.toggle('active', b.dataset.status===''));
      atualizarSelecaoGraficos(); aplicarFiltrosChips();
    });
  }

  // --- Aplicação combinada: filtros + ordenação + render ---
  function aplicarFiltrosChips(){
    const busca = document.getElementById('fBusca').value.trim().toLowerCase();
    let filtrados = C.lista.filter(x=>{
      if (filtroStatus && x.status !== filtroStatus) return false;
      if (filtroCliente && x.cliente !== filtroCliente) return false;
      if (filtroOperadora && x.operadora !== filtroOperadora) return false;
      if (filtroFaixa && !(x.pct>=filtroFaixa.min && x.pct<filtroFaixa.max)) return false;
      if (busca && !`${x.codigo} ${x.iccid} ${x.telefone} ${x.cliente} ${x.cidade}`.toLowerCase().includes(busca)) return false;
      return true;
    });
    filtrados = ordenar(filtrados);
    ultimosFiltrados = filtrados;
    document.getElementById('tbodyChips').innerHTML = filtrados.length
      ? filtrados.map(linhaHtml).join('')
      : `<tr><td colspan="13" style="text-align:center;padding:24px;opacity:.6">Nenhum chip encontrado.</td></tr>`;
    renderBadges();
  }
  document.getElementById('fBusca').addEventListener('input', aplicarFiltrosChips);
  document.getElementById('fCliente').addEventListener('change', e=>{ filtroCliente=e.target.value; atualizarSelecaoGraficos(); aplicarFiltrosChips(); });
  document.getElementById('fOperadora').addEventListener('change', e=>{ filtroOperadora=e.target.value; atualizarSelecaoGraficos(); aplicarFiltrosChips(); });
  document.querySelectorAll('#statusChips .chipBtn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      filtroStatus = btn.dataset.status;
      atualizarSelecaoGraficos();
      aplicarFiltrosChips();
    });
  });
  aplicarFiltrosChips();
} else if (PAYLOAD.chips && PAYLOAD.chips.erro) {
  document.getElementById('chipsErro').style.display = '';
}
</script>
</body>
</html>"""


def _check():
    # ponytail: valida a agregação — simultaneidade e contagem por código.
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

    base = datetime(2026, 1, 1)
    ts = [base, base + timedelta(minutes=1), base + timedelta(minutes=2),
          base + timedelta(hours=2), base + timedelta(hours=2, minutes=1)]
    perdido, total, inicio, fim = calcular_perdido(ts, limite_min=10)
    assert perdido == timedelta(hours=1, minutes=58), perdido
    assert total == timedelta(hours=2, minutes=1)
    print("check ok")


if __name__ == "__main__":
    import sys
    _check() if "check" in sys.argv else main()
