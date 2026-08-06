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
    # ponytail: comparar os dt parseados, não a string "registrada" direto - "DD-MM-YYYY" como
    # string ordena errado entre meses/dias diferentes (bug real encontrado 2026-08-03: "01-08"
    # perdia pra "31-07" porque '0' < '3' no primeiro caractere, mesmo Ago vindo depois de Jul).
    primeiro_dt, ultimo_dt = {}, {}
    por_ciclo = defaultdict(set)  # registrada -> set de codigos offline naquele ciclo
    for codigo, registrada, reset in linhas:
        d = por_codigo[codigo]
        d["ciclos"] += 1
        if reset.strip().lower() == "sim":
            d["resets"] += 1
        dt = datetime.strptime(registrada, FMT)
        if codigo not in primeiro_dt or dt < primeiro_dt[codigo]:
            primeiro_dt[codigo] = dt
            d["primeiro"] = registrada
        if codigo not in ultimo_dt or dt > ultimo_dt[codigo]:
            ultimo_dt[codigo] = dt
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
    """8d 16h 12m / 3d 16h / 16h 12m / 16h - sem decimais confusas (16.2h != "16h e 2min")."""
    total_min = round(h * 60)
    d, resto_min = divmod(total_min, 24 * 60)
    hh, mm = divmod(resto_min, 60)
    partes = ([f"{d}d"] if d else []) + [f"{hh}h"] + ([f"{mm}m"] if mm else [])
    return " ".join(partes)


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


# ponytail: Chart.js via CDN (a máquina do monitor está online). Se precisar 100%
# offline, baixar chart.umd.min.js pra pasta e trocar o src.
TEMPLATE = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Observador Raptor — comunicação dos equipamentos</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root{ --border:rgba(255,255,255,.10) }
  *{box-sizing:border-box}
  button{font:inherit;appearance:none;-webkit-appearance:none;-moz-appearance:none}
  html{scrollbar-width:thin;scrollbar-color:#2A2E3F #13151E}
  ::-webkit-scrollbar{width:8px;height:8px}
  ::-webkit-scrollbar-track{background:#13151E}
  ::-webkit-scrollbar-thumb{background:#2A2E3F;border-radius:9999px;border:2px solid transparent;background-clip:padding-box}
  ::-webkit-scrollbar-thumb:hover{background:#3B82F6}
  ::-webkit-scrollbar-corner{background:#13151E}
  body{margin:0;background:#05060A;color:#F5F6FA;
    font-family:system-ui,-apple-system,"Segoe UI",sans-serif;line-height:1.45}
  .wrap{display:flex;flex-direction:column;align-items:stretch;
    max-width:1280px;width:100%;margin:0 auto;padding:28px 20px 60px}
  .tabPanel{display:flex;flex-direction:column;align-items:stretch;gap:24px}
  .panelHeader{display:flex;flex-direction:column}
  h1{font-size:22px;margin:0 0 2px}
  .sub{font-size:13px;margin:0 0 20px;opacity:.65}
  .tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:0}
  #tiles{grid-template-columns:repeat(6,1fr)}
  #kpiRow{grid-template-columns:repeat(4,1fr)}
  @media (max-width:860px){#tiles,#kpiRow{grid-template-columns:repeat(auto-fit,minmax(150px,1fr))}}
  .tile{padding:8px 20px 6px;cursor:default}
  .tile:not(:first-child){border-left:1px solid rgba(255,255,255,.10)}
  .tile .v{font-size:28px;font-weight:700;letter-spacing:-.5px;line-height:1.25}
  .tile .k{font-size:12px;opacity:.65;margin-top:4px;line-height:1.3}
  .card{border-radius:14px;padding:18px 18px 8px;width:100%}
  .card h2{font-size:15px;margin:0 0 2px}
  .card p{font-size:12.5px;opacity:.65;margin:0 0 14px}
  .chart{position:relative;height:340px}
  .chart.short{height:250px}
  table{width:100%;border-collapse:collapse;font-size:13px;font-variant-numeric:tabular-nums}
  th,td{text-align:right;padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
  th:first-child,td:first-child{text-align:left;font-variant-numeric:normal}
  .mono{font-family:'SFMono-Regular',Consolas,'Liberation Mono',monospace}
  .detailsToggle{cursor:pointer;font-size:13px;opacity:.85;list-style:none;user-select:none;
    display:inline-flex;align-items:center;gap:6px;margin-top:4px;padding:8px 14px;
    border-radius:8px;border:1px solid var(--border);background:rgba(255,255,255,.05)}
  .detailsToggle::-webkit-details-marker{display:none}
  .detailsToggle::before{content:'▶';font-size:9px;transition:transform .15s}
  details[open]>.detailsToggle::before{transform:rotate(90deg)}
  .detailsToggle:hover{background:rgba(255,255,255,.09)}
  th{font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:rgba(245,246,250,.55)}
  details summary{cursor:pointer;font-size:13px;margin-top:4px;opacity:.7}
  .chk{display:inline-flex;align-items:center;gap:8px;font-size:14px;opacity:.75;
    margin:-8px 0 0;cursor:pointer}
  .chk input{width:20px;height:20px;cursor:pointer}
  .tabs{display:flex;gap:8px;margin-bottom:22px;border-bottom:1px solid var(--border);
    position:sticky;top:0;z-index:50;background:rgba(5,6,10,.85);backdrop-filter:blur(12px);
    -webkit-backdrop-filter:blur(12px);box-shadow:0 2px 12px rgba(0,0,0,.35)}
  .tabBtn{padding:9px 16px;border:none;background:transparent;color:inherit;opacity:.55;
    font-size:13.5px;font-weight:600;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
  .tabBtn.active{opacity:1;border-bottom-color:#2563EB}

  /* Dark Glassmorphism — mesmo tema nas duas abas (cobre a área toda) */
  body.theme-equip,body.theme-chips{background:linear-gradient(160deg, #171233 0%, #0D0F18 45%, #05060A 100%) fixed}
  #tabEquip .card,#tabChips .card{
    background:rgba(255,255,255,.04);backdrop-filter:blur(16px);
    -webkit-backdrop-filter:blur(16px);border:1px solid rgba(255,255,255,.09)}
  #tabEquip .tile.alerta .v,#tabChips .tile.alerta .v{color:#EF4444}
  #tabChips .tile.ok .v{color:#10B981}

  .biRow4{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
  .biRow4 .chart{border-radius:10px}
  @media (max-width:980px){.biRow4{grid-template-columns:1fr 1fr}}
  @media (max-width:560px){.biRow4{grid-template-columns:1fr}}

  .tlMiniCards{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}
  .tlMiniCard{background:rgba(255,255,255,.04);border:1px solid var(--border);border-radius:10px;padding:10px 14px}
  .tlMiniCard .v{font-size:20px;font-weight:700}
  .tlMiniCard .k{font-size:11.5px;opacity:.65;margin-top:2px}
  @media (max-width:560px){.tlMiniCards{grid-template-columns:1fr}}
  .periodoBtns{display:inline-flex;gap:2px;padding:4px;margin:0 0 14px;
    background:rgba(255,255,255,.04);border-radius:10px}

  .tableCard{width:100%;
    background:#12141c;border:1px solid rgba(255,255,255,.09);border-radius:14px;overflow:hidden}
  .tableCard .filtros,.tableCard .statusChips,.tableCard .filtrosAtivos,.tableCard .tableToolbar{
    padding-left:20px;padding-right:20px}
  .tableCard .tableToolbar{padding-top:14px;padding-bottom:14px;margin:0;border-bottom:1px solid rgba(255,255,255,.09)}

  .filtros{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:0;padding-top:14px;padding-bottom:10px}
  .filtros input,.filtros select{padding:9px 12px;border-radius:8px;border:1px solid var(--border);
    background:rgba(255,255,255,.05);color:inherit;font-size:13px}
  .filtros input{flex:1;min-width:220px}
  .filtros select option{background:#13151E;color:#F1F5F9}
  .filtros select option:hover,.filtros select option:checked,.filtros select option:focus{background:#2563EB;color:#fff}
  .filtrosAcoes{display:flex;gap:10px;margin-left:auto;position:relative}
  .statusChips{display:inline-flex;align-items:center;gap:2px;margin:0 0 14px;padding:4px;
    background:rgba(255,255,255,.04);border-radius:10px}
  .chipBtn{padding:7px 14px;border:none;border-radius:7px;background:none;
    color:inherit;opacity:.6;font-size:13px;cursor:pointer;display:inline-flex;align-items:center;gap:6px;
    transition:opacity .15s,background-color .15s}
  .chipBtn:hover{opacity:.9;background:rgba(255,255,255,.06)}
  .chipBtn.active{opacity:1;background:rgba(37,99,235,.18);color:#60A5FA;font-weight:600}
  .chipBtn b{margin-left:0;background:rgba(255,255,255,.08);padding:1px 7px;border-radius:10px;font-size:11px;font-weight:600}
  .chipBtn.active b{background:rgba(37,99,235,.3);color:#93C5FD}

  .filtrosAtivos{display:none;align-items:center;gap:8px;flex-wrap:wrap;margin:0;padding-top:14px}
  .filtroBadge{background:rgba(37,99,235,.18);border:1px solid rgba(37,99,235,.4);border-radius:20px;
    padding:5px 10px;font-size:12px;cursor:pointer}
  .filtroBadge b{margin-left:5px;opacity:.7}
  #btnLimparFiltros{border:1px solid var(--border);background:transparent;color:inherit;opacity:.75;
    border-radius:20px;padding:5px 12px;font-size:12px;cursor:pointer}

  .badge{padding:2px 9px;border-radius:10px;font-size:11.5px;font-weight:600;white-space:nowrap;
    display:inline-flex;align-items:center;gap:5px}
  .badge .dot{width:7px;height:7px;border-radius:50%;display:inline-block}
  .badge.atualizado{background:rgba(16,185,129,.14);color:#10B981}
  .badge.atualizado .dot{background:#10B981}
  .badge.atrasado{background:rgba(245,158,11,.16);color:#F59E0B}
  .badge.atrasado .dot{background:#F59E0B}
  .badge.sem_comunicacao{background:rgba(239,68,68,.14);color:#EF4444}
  .badge.sem_comunicacao .dot{background:#EF4444}
  tr.linhaCritica{background:rgba(239,68,68,.07)}

  .tableToolbar{display:flex;gap:10px;align-items:center;position:relative}
  .tbBtn{padding:8px 14px;border-radius:8px;border:1px solid var(--border);background:rgba(255,255,255,.05);
    color:inherit;font-size:12.5px;cursor:pointer}
  .colunasMenu{display:none;position:absolute;top:38px;left:0;z-index:20;background:#12141c;
    border:1px solid var(--border);border-radius:10px;padding:10px 14px;min-width:210px;
    box-shadow:0 10px 30px rgba(0,0,0,.4)}
  .colunasMenu label{display:flex;gap:8px;align-items:center;font-size:12.5px;padding:4px 0;cursor:pointer}
  .colunasAcoes{display:flex;gap:8px;margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid var(--border)}
  .colunasAcoes button{flex:1;background:rgba(255,255,255,.05);border:1px solid var(--border);color:inherit;
    border-radius:6px;padding:5px 0;font-size:11.5px;cursor:pointer}
  .colunasAcoes button:hover{background:rgba(37,99,235,.18)}

  .tableWrap{position:relative;max-height:640px;overflow-y:auto;overflow-x:auto;padding:0}
  .tableWrap table{margin:0;border-collapse:separate;border-spacing:0}
  .tableWrap td,.tableWrap th{height:48px;padding:0 16px;white-space:nowrap;box-sizing:border-box}
  .tableWrap th{position:sticky !important;top:0 !important;background-color:#13151E !important;
    background-clip:padding-box;cursor:pointer;user-select:none;z-index:100 !important;
    box-shadow:0 2px 5px rgba(0,0,0,.5)}
  .tableWrap tbody tr:hover{background:rgba(255,255,255,.04)}
  .tableWrap th[data-col]:hover{color:rgba(245,246,250,.85)}
  .tableWrap th[data-col]::after{content:'⇅';margin-left:5px;opacity:.3;font-size:10px}
  .tableWrap th.sortAsc::after{content:'▲';opacity:1;color:#2563EB}
  .tableWrap th.sortDesc::after{content:'▼';opacity:1;color:#2563EB}

  /* Tabela de chips: coluna Cliente fixa na rolagem horizontal + alinhamento por tipo de coluna */
  #tabelaChips td:first-child,#tabelaChips th:first-child{position:sticky;left:0;background-color:#12141c;z-index:2}
  #tabelaChips thead th:first-child{z-index:101 !important}
  #tabelaChips tbody tr:hover td:first-child{background-color:#181b26}
  #tabelaChips td:nth-child(2),#tabelaChips td:nth-child(3),#tabelaChips td:nth-child(4),
  #tabelaChips td:nth-child(5),#tabelaChips td:nth-child(6),#tabelaChips td:nth-child(8),
  #tabelaChips th:nth-child(2),#tabelaChips th:nth-child(3),#tabelaChips th:nth-child(4),
  #tabelaChips th:nth-child(5),#tabelaChips th:nth-child(6),#tabelaChips th:nth-child(8){text-align:left}
  #tabelaChips td:last-child,#tabelaChips th:last-child{text-align:center}
  .vazio{opacity:.35}

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
    <div class="panelHeader">
      <h1>Comunicação dos equipamentos</h1>
      <p class="sub">A planilha registra os equipamentos <b>fora do ar</b> a cada ciclo (~1&nbsp;min).
         Mais ciclos = mais tempo sem comunicar. Observando desde <span id="desde"></span>.
         Gerado em <span id="ger"></span>.</p>
      <label class="chk"><input type="checkbox" id="chkSem16"> Remover exceções das métricas visuais</label>
    </div>

    <div class="tiles" id="tiles"></div>

    <div class="card">
      <h2>Quanto cada equipamento ficou fora do ar</h2>
      <p>Tempo total offline por código (cada ciclo ≈ 1 minuto). <span style="color:#EF4444">Vermelho</span> ≥24h,
         <span style="color:#F59E0B">amarelo</span> 5h–24h, <span style="color:#2563EB">azul</span> abaixo de 5h.</p>
      <label class="chk" style="margin:0 0 10px"><input type="checkbox" id="chkOcultarMaior"> Ocultar o maior valor (ver proporção dos demais)</label>
      <div class="chart" id="wrapRank"><canvas id="rank"></canvas></div>
    </div>

    <div class="card">
      <h2>Equipamentos Simultaneamente Offline</h2>
      <p>Histórico de indisponibilidade ao longo do tempo.</p>
      <div class="periodoBtns" id="periodoTl">
        <button class="chipBtn active" data-periodo="0">Tudo</button>
        <button class="chipBtn" data-periodo="1">24h</button>
        <button class="chipBtn" data-periodo="7">7 dias</button>
        <button class="chipBtn" data-periodo="30">30 dias</button>
      </div>
      <div class="tlMiniCards" id="tlMiniCards"></div>
      <div class="chart" id="wrapTl"><canvas id="tl"></canvas></div>
    </div>

    <div class="card">
      <h2>Distribuição da simultaneidade</h2>
      <p>Em quantos ciclos houve 1, 2, 3… equipamentos offline ao mesmo tempo.</p>
      <div class="chart short" id="wrapDist"><canvas id="dist"></canvas></div>
    </div>

    <details open>
      <summary class="detailsToggle">Ver tabela</summary>
      <div class="tableWrap" style="max-height:480px">
        <table id="tabelaRanking"><thead><tr>
          <th data-col="codigo" data-type="text">Código</th>
          <th data-col="ciclos" data-type="num">Ciclos offline</th>
          <th data-col="resets" data-type="num">Resets enviados</th>
          <th data-col="primeiro" data-type="date">Primeiro</th>
          <th data-col="ultimo" data-type="date">Último</th>
        </tr></thead><tbody id="tbody"></tbody></table>
      </div>
    </details>
  </div>

  <div id="tabChips" class="tabPanel" style="display:none">
    <div class="panelHeader">
      <h1>Chips e consumo</h1>
      <p class="sub" id="chipsErro" style="display:none">Não foi possível carregar os dados da API.</p>
      <p class="sub" id="chipsSync"></p>
    </div>

    <div class="tiles" id="kpiRow"></div>

    <div class="biRow4">
      <div class="card">
        <h2>Status de conexão</h2>
        <p>Clique em uma fatia para filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsStatusChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Quantidade por cliente</h2>
        <p>Clique em uma barra para filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsClienteChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Quantidade por operadora</h2>
        <p>Clique em uma fatia para filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsOperadoraChart"></canvas></div>
      </div>
      <div class="card">
        <h2>Faixa de consumo</h2>
        <p>Clique em uma fatia para filtrar a tabela.</p>
        <div class="chart short"><canvas id="chipsConsumoChart"></canvas></div>
      </div>
    </div>

    <div class="tableCard">
      <div class="filtros">
        <input id="fBusca" type="text" placeholder="Buscar por código, ICCID, telefone, cliente ou cidade...">
        <select id="fCliente"><option value="">Todos os clientes</option></select>
        <select id="fOperadora"><option value="">Todas as operadoras</option></select>
        <select id="fCidade"><option value="">Todas as cidades</option></select>
        <div class="filtrosAcoes">
          <button class="tbBtn" id="btnColunas">Colunas ▾</button>
          <div class="colunasMenu" id="colunasMenu"></div>
          <button class="tbBtn" id="btnExportarCsv">Exportar CSV</button>
          <button class="tbBtn" id="btnExportarXlsx">Exportar XLSX</button>
        </div>
      </div>
      <div class="statusChips" id="statusChips">
        <button class="chipBtn active" data-status="">Todos <b id="cntTodos">0</b></button>
        <button class="chipBtn" data-status="atualizado">Atualizados <b id="cntAtualizado">0</b></button>
        <button class="chipBtn" data-status="atrasado">Atrasados <b id="cntAtrasado">0</b></button>
        <button class="chipBtn" data-status="sem_comunicacao">Sem comunicação <b id="cntSemComunicacao">0</b></button>
      </div>
      <div class="filtrosAtivos" id="filtrosAtivos"></div>

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
// ponytail: paleta suavizada (2026-08-03) - neon (#FF4D4D/#FFC107/#00E676) trocado por
// tons mais elegantes de dark mode; badges em CSS (.badge.*) usam os mesmos hex, manter em sincronia
const EQUIP = {blue:'#2563EB', red:'#EF4444', amber:'#F59E0B', green:'#10B981',
  ink:'#F5F6FA', ink2:'#9CA3AF', muted:'#6B7280', grid:'#20232F'};
const PAL = {blue:'#2563EB', red:'#EF4444', amber:'#F59E0B', green:'#10B981',
  ink:'#F5F6FA', ink2:'rgba(245,246,250,.62)', muted:'rgba(245,246,250,.42)', grid:'rgba(255,255,255,.08)'};
const INK=EQUIP.ink, INK2=EQUIP.ink2, MUTED=EQUIP.muted, GRID=EQUIP.grid, BLUE=EQUIP.blue, RED=EQUIP.red;
Chart.defaults.font.family = 'system-ui,-apple-system,"Segoe UI",sans-serif';
Chart.defaults.color = '#9CA3AF';
Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15,17,24,.95)';
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.boxPadding = 6;
Chart.defaults.plugins.tooltip.titleColor = '#F5F6FA';
Chart.defaults.plugins.tooltip.bodyColor = '#E2E8F0';
Chart.defaults.plugins.tooltip.cornerRadius = 8;
const grid = {color:GRID, drawTicks:false, drawBorder:false};
function hexAlpha(hex, a){
  const n = parseInt(hex.replace('#',''),16);
  return `rgba(${(n>>16)&255},${(n>>8)&255},${n&255},${a})`;
}
function fmtNum(n){ return Number(n).toLocaleString('pt-BR'); }
// ponytail: ciclo ~ 1min - converte pra "Xd Yh Zm" em vez de forçar cálculo mental no usuário
function fmtMin(min){
  min = Math.round(min);
  const d = Math.floor(min/1440), h = Math.floor((min%1440)/60), m = min%60;
  return ([d?`${d}d`:'', `${h}h`, m?`${m}m`:''].filter(Boolean)).join(' ');
}
function fmtDataCurta(ts){
  if (!ts) return '';
  const [data, hora] = ts.split(' ');
  const [d,m] = data.split('-');
  return `${d}/${m} ${(hora||'').slice(0,5)}`;
}
const MESES_ABREV = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function fmtEixoTl(ts){
  const [data, hora] = ts.split(' ');
  const [d,m] = data.split('-');
  return `${d}/${MESES_ABREV[+m-1]} ${hora.slice(0,5)}`;
}
function parseTs(ts){
  const [data, hora] = ts.split(' ');
  const [d,m,a] = data.split('-').map(Number);
  const [hh,mi,ss] = hora.split(':').map(Number);
  return new Date(a, m-1, d, hh, mi, ss);
}
function filtrarTimelinePorPeriodo(timeline, dias){
  if (!dias || !timeline.length) return timeline;
  const corte = parseTs(timeline[timeline.length-1].t).getTime() - dias*86400000;
  return timeline.filter(p => parseTs(p.t).getTime() >= corte);
}
// --- Callout permanente no ponto de pico (data/hora do evento, não só o tooltip) ---
const picoCalloutPlugin = {
  id: 'picoCallout',
  afterDatasetsDraw(chart){
    const cfg = chart.options.plugins.picoCallout;
    if (!cfg || cfg.indice==null) return;
    const pt = chart.getDatasetMeta(0).data[cfg.indice];
    if (!pt) return;
    const {ctx} = chart;
    ctx.save();
    ctx.font = '600 11px system-ui,-apple-system,"Segoe UI",sans-serif';
    const largura = ctx.measureText(cfg.texto).width + 16;
    const x = Math.min(Math.max(pt.x - largura/2, chart.chartArea.left+4), chart.chartArea.right - largura - 4);
    const y = Math.max(pt.y - 32, chart.chartArea.top + 4);
    ctx.fillStyle = 'rgba(239,68,68,.95)';
    ctx.beginPath();
    ctx.roundRect(x, y, largura, 22, 6);
    ctx.fill();
    ctx.fillStyle = '#fff';
    ctx.textBaseline = 'middle';
    ctx.fillText(cfg.texto, x+8, y+11);
    ctx.restore();
  }
};
// ponytail: mesmos limiares do texto (>=24h crítico, >=5h atenção) - ajustar aqui se mudar o critério
function severidade(min){
  if (min >= 1440) return 'critico';
  if (min >= 300) return 'atencao';
  return 'baixo';
}
function corSeveridade(min){
  const s = severidade(min);
  return s==='critico' ? RED : s==='atencao' ? PAL.amber : BLUE;
}

// --- Rosca com total no centro (aproveita o furo em vez de deixar vazio) ---
const centroRoscaPlugin = {
  id: 'centroRosca',
  afterDraw(chart){
    const cfg = chart.options.plugins.centroRosca;
    if (!cfg) return;
    const {ctx, chartArea:{left,right,top,bottom}} = chart;
    const cx = (left+right)/2, cy = (top+bottom)/2;
    ctx.save();
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillStyle = PAL.ink;
    ctx.font = '700 26px system-ui,-apple-system,"Segoe UI",sans-serif';
    ctx.fillText(cfg.valor, cx, cy - 8);
    ctx.fillStyle = PAL.ink2;
    ctx.font = '12px system-ui,-apple-system,"Segoe UI",sans-serif';
    ctx.fillText(cfg.rotulo, cx, cy + 14);
    ctx.restore();
  }
};

// Tooltip na borda externa da fatia, nunca sobre o número central da rosca.
Chart.Tooltip.positioners.roscaFora = function(itens){
  if (!itens.length) return false;
  const el = itens[0].element, {left,right,top,bottom} = this.chart.chartArea;
  const cx=(left+right)/2, cy=(top+bottom)/2, ang=(el.startAngle+el.endAngle)/2, raio=el.outerRadius+14;
  return {x: cx+Math.cos(ang)*raio, y: cy+Math.sin(ang)*raio};
};

// --- Legenda com contagem + porcentagem (em vez de só a cor+nome) ---
function legendComContagem(){
  return {
    generateLabels(chart){
      const ds = chart.data.datasets[0];
      const total = ds.data.reduce((a,b)=>a+b,0) || 1;
      return chart.data.labels.map((l,i)=>({
        text: `${l}: ${fmtNum(ds.data[i])} (${(ds.data[i]/total*100).toFixed(1)}%)`,
        fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], index: i,
      }));
    }
  };
}

// --- Ordenação tri-estado reutilizável (cabeçalhos data-col) ---
function criarOrdenacao(tabelaId, cols, aoOrdenar){
  let sortCol = null, sortDir = 0;
  document.querySelectorAll(`#${tabelaId} th[data-col]`).forEach(th=>{
    th.addEventListener('click', ()=>{
      const col = th.dataset.col;
      if (sortCol !== col) { sortCol = col; sortDir = 1; }
      else { sortDir = sortDir===1?2:(sortDir===2?0:1); if (sortDir===0) sortCol=null; }
      document.querySelectorAll(`#${tabelaId} th[data-col]`).forEach(h=>h.classList.remove('sortAsc','sortDesc'));
      if (sortCol) th.classList.add(sortDir===1?'sortAsc':'sortDesc');
      aoOrdenar();
    });
  });
  return function ordenar(lista){
    if (!sortCol || sortDir===0) return lista;
    const tipo = cols.find(c=>c.key===sortCol).type;
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
  };
}

// --- Exportação CSV/XLSX reutilizável, respeitando filtro/ordenação atuais ---
function exportarTabela(nomeBase, cols, linhas, formato){
  if (formato === 'xlsx') {
    const dados = linhas.map(x => { const o={}; cols.forEach(c=>{ o[c.label] = c.fmt ? c.fmt(x) : (x[c.key] ?? ''); }); return o; });
    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados');
    XLSX.writeFile(wb, `${nomeBase}.xlsx`);
  } else {
    const linhasCsv = [cols.map(c=>c.label).join(';')].concat(linhas.map(x=>
      cols.map(c=>`"${((c.fmt ? c.fmt(x) : x[c.key]) ?? '').toString().replace(/"/g,'""')}"`).join(';')));
    const blob = new Blob(['﻿'+linhasCsv.join('\n')], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${nomeBase}.csv`; a.click();
    URL.revokeObjectURL(url);
  }
}

document.getElementById('ger').textContent = PAYLOAD.completo.gerado_em;
document.getElementById('desde').textContent = PAYLOAD.observacao.desde || '—';

// --- Rótulo (tempo + valor bruto) desenhado no fim de cada barra horizontal ---
const rotuloBarraPlugin = {
  id: 'rotuloBarra',
  afterDatasetsDraw(chart){
    const {ctx} = chart;
    const meta = chart.getDatasetMeta(0);
    ctx.save();
    ctx.font = '11px system-ui,-apple-system,"Segoe UI",sans-serif';
    ctx.fillStyle = INK2;
    ctx.textBaseline = 'middle';
    meta.data.forEach((bar, i)=>{
      const v = chart.data.datasets[0].data[i];
      ctx.fillText(`${fmtMin(v)} (${fmtNum(v)})`, bar.x + 6, bar.y);
    });
    ctx.restore();
  }
};

const COLS_RANKING = [
  {key:'codigo', label:'Código', type:'text'}, {key:'ciclos', label:'Ciclos offline', type:'num'},
  {key:'resets', label:'Resets enviados', type:'num'},
  {key:'primeiro', label:'Primeiro', type:'date'}, {key:'ultimo', label:'Último', type:'date'},
];
let dadosAtual = PAYLOAD.completo;
let ocultarMaior = false;
let periodoTl = 0;
const ordenarRanking = criarOrdenacao('tabelaRanking', COLS_RANKING, ()=>renderTbody(dadosAtual));

function renderTbody(D){
  const linhas = ordenarRanking(D.ranking);
  document.getElementById('tbody').innerHTML = linhas.map(r=>{
    const critico = severidade(r.ciclos)==='critico';
    return `<tr class="${critico?'linhaCritica':''}">
      <td class="mono">${r.codigo} ${critico?'<span class="badge sem_comunicacao"><span class="dot"></span>CRÍTICO</span>':''}</td>
      <td class="mono">${fmtNum(r.ciclos)}</td>
      <td class="mono">${fmtNum(r.resets)}</td>
      <td title="${r.primeiro||''}">${fmtDataCurta(r.primeiro)}</td>
      <td title="${r.ultimo||''}">${fmtDataCurta(r.ultimo)}</td>
    </tr>`;
  }).join('');
}

let charts = {};
function render(D){
  dadosAtual = D;
  const T = D.totais;
  document.getElementById('tiles').innerHTML = [
    ['Tempo observado', PAYLOAD.observacao.observado, ''],
    ['Tempo perdido (falha do script)', PAYLOAD.observacao.perdido, PAYLOAD.observacao.perdidoHoras>0?'alerta':''],
    ['Equipamentos com falha', fmtNum(T.equipamentos), ''],
    ['Ciclos offline (~min)', fmtNum(T.ciclosOffline), ''],
    ['Resets enviados', fmtNum(T.resets), ''],
    ['Pico simultâneo', fmtNum(T.pico), T.pico>1?'alerta':''],
  ].map(([k,v,c])=>`<div class="tile ${c}"><div class="v">${v}</div><div class="k">${k}</div></div>`).join('');

  Object.values(charts).forEach(c=>c.destroy());

  // Ranking — barras horizontais coloridas por severidade (tempo real, não ciclo cru),
  // com opção de ocultar o maior valor pra não esmagar a escala dos demais
  const rankData = ocultarMaior ? D.ranking.slice(1) : D.ranking;
  charts.rank = new Chart(document.getElementById('rank'), {
    type:'bar', plugins:[rotuloBarraPlugin],
    data:{labels:rankData.map(r=>r.codigo),
      datasets:[{data:rankData.map(r=>r.ciclos), backgroundColor:rankData.map(r=>corSeveridade(r.ciclos)),
        borderRadius:4, borderSkipped:false, barThickness:'flex', maxBarThickness:26}]},
    options:{indexAxis:'y', maintainAspectRatio:false, layout:{padding:{right:70}},
      plugins:{legend:{display:false},
        tooltip:{callbacks:{label:c=>` ${fmtMin(c.raw)} (${fmtNum(c.raw)} ciclos)`}}},
      scales:{x:{beginAtZero:true, grid, ticks:{color:MUTED}},
        y:{grid:{display:false}, ticks:{color:INK}}}}
  });
  document.getElementById('wrapRank').style.height = Math.max(260, rankData.length*34+40)+'px';

  // Timeline filtrada pelo período escolhido (Tudo/24h/7 dias/30 dias)
  const tlD = filtrarTimelinePorPeriodo(D.timeline, periodoTl);
  const picoLocal = tlD.reduce((m,p)=>Math.max(m,p.n), 0);
  const picoIdx = picoLocal>0 ? tlD.findIndex(p=>p.n===picoLocal) : -1;

  // Mini-cards de resumo (atual / pico / média), com indicador de cor e unidade
  const tlAtual = tlD.length ? tlD[tlD.length-1].n : 0;
  const tlMedia = tlD.length ? tlD.reduce((a,p)=>a+p.n,0)/tlD.length : 0;
  const dot = cor=>`<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${cor};margin-right:6px"></span>`;
  document.getElementById('tlMiniCards').innerHTML = [
    ['Atual', fmtNum(tlAtual), tlAtual>0?RED:PAL.green],
    ['Pico máximo', fmtNum(picoLocal), RED],
    ['Média', `${tlMedia.toFixed(1)} equip.`, BLUE],
  ].map(([k,v,cor])=>`<div class="tlMiniCard"><div class="v">${dot(cor)}${v}</div><div class="k">${k}</div></div>`).join('');

  // área com gradiente (em vez de bloco sólido), stepped pra ler ciclo a ciclo,
  // ponto de pico destacado em vermelho com callout permanente de data/hora
  const tlEl = document.getElementById('tl');
  const gradTl = tlEl.getContext('2d').createLinearGradient(0, 0, 0, tlEl.clientHeight || 340);
  gradTl.addColorStop(0, hexAlpha(BLUE,.30));
  gradTl.addColorStop(1, hexAlpha(BLUE,0));
  const tlEhPico = tlD.map((p,i)=>i===picoIdx);
  charts.tl = new Chart(tlEl, {
    type:'line', plugins:[picoCalloutPlugin],
    data:{labels:tlD.map(p=>fmtEixoTl(p.t)),
      datasets:[{data:tlD.map(p=>p.n), borderColor:BLUE,
        backgroundColor:gradTl, fill:true, stepped:true, borderWidth:2,
        pointRadius:tlEhPico.map(pico=>pico?5:0), pointHoverRadius:6,
        pointBackgroundColor:tlEhPico.map(pico=>pico?RED:BLUE),
        pointBorderColor:tlEhPico.map(pico=>pico?RED:BLUE)}]},
    options:{maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins:{legend:{display:false},
        picoCallout: picoIdx>=0 ? {indice:picoIdx, texto:`Pico: ${picoLocal} equip. em ${fmtEixoTl(tlD[picoIdx].t)}`} : null,
        tooltip:{callbacks:{
          title:it=>tlD[it[0].dataIndex].t,
          label:c=>{const p=tlD[c.dataIndex];
            return p.n?` ${p.n} offline: ${p.codigos.join(', ')}`:' tudo comunicando'}}}},
      scales:{y:{min:0, beginAtZero:true, max:picoLocal||undefined, ticks:{precision:0, color:MUTED, stepSize:1}, grid},
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
        tooltip:{callbacks:{label:c=>` ${fmtNum(c.raw)} ciclos`}}},
      scales:{y:{beginAtZero:true, ticks:{precision:0, color:MUTED}, grid},
        x:{grid:{display:false}, ticks:{color:INK}}}}
  });

  renderTbody(D);
}

document.getElementById('chkSem16').addEventListener('change', e=>{
  render(e.target.checked ? PAYLOAD.filtrado : PAYLOAD.completo);
});
document.getElementById('chkOcultarMaior').addEventListener('change', e=>{
  ocultarMaior = e.target.checked;
  render(dadosAtual);
});
document.querySelectorAll('#periodoTl .chipBtn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#periodoTl .chipBtn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    periodoTl = Number(btn.dataset.periodo);
    render(dadosAtual);
  });
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
    // ponytail: gráficos criados com a aba escondida (display:none) ficam com canvas 0x0 e
    // o Chart.js não redimensiona sozinho quando ela aparece - força um resize ao trocar de aba.
    document.querySelectorAll('canvas').forEach(cv => Chart.getChart(cv)?.resize());
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
    {key:'ultimaComunicacao', label:'Última comunicação', type:'date'},
    {key:'status', label:'Status', type:'text', fmt:x=>STATUS_LABEL[x.status]},
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

  let filtroStatus = '', filtroCliente = '', filtroOperadora = '', filtroCidade = '', filtroFaixa = null;
  let ultimosFiltrados = [];

  const statusChart = new Chart(document.getElementById('chipsStatusChart'), {
    type:'doughnut', plugins:[centroRoscaPlugin],
    data:{labels:['Atualizado','Atrasado','Sem comunicação'],
      datasets:[{data:[C.totais.atualizados, C.totais.atrasados, C.totais.semComunicacao],
        backgroundColor:[PAL.green, PAL.amber, PAL.red], borderWidth:0, hoverOffset:8}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{tooltip:{position:'roscaFora'}, legend:{position:'bottom', labels:{color:'#E2E8F0', boxWidth:12, ...legendComContagem()}},
        centroRosca:{valor:fmtNum(C.totais.chips), rotulo:'Chips Ativos'}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = ['atualizado','atrasado','sem_comunicacao'][els[0].index];
        filtroStatus = filtroStatus===val ? '' : val;
        document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.toggle('active', b.dataset.status===filtroStatus));
        recalcularGraficos(); aplicarFiltrosChips();
      }}
  });

  const rotuloClientePlugin = {
    id: 'rotuloCliente',
    afterDatasetsDraw(chart){
      const {ctx} = chart;
      ctx.save();
      ctx.font = '11px system-ui,-apple-system,"Segoe UI",sans-serif';
      ctx.fillStyle = '#E2E8F0';
      ctx.textBaseline = 'middle';
      chart.getDatasetMeta(0).data.forEach((bar,i)=>{
        ctx.fillText(fmtNum(chart.data.datasets[0].data[i]), bar.x+6, bar.y);
      });
      ctx.restore();
    }
  };
  const clienteChart = new Chart(document.getElementById('chipsClienteChart'), {
    type:'bar', plugins:[rotuloClientePlugin],
    data:{labels:clientesOrd.map(c=>c[0]),
      datasets:[{data:clientesOrd.map(c=>c[1]), backgroundColor:clientesOrd.map(()=>BLUE),
        borderRadius:4, borderSkipped:false, maxBarThickness:26}]},
    options:{indexAxis:'y', maintainAspectRatio:false, layout:{padding:{right:40}},
      onHover:(e,els,chart)=>{
        e.native.target.style.cursor = els.length?'pointer':'default';
        chart.data.datasets[0].backgroundColor = clientesOrd.map((c,i)=>
          !els.length || i===els[0].index ? BLUE : hexAlpha(BLUE,.35));
        chart.update('none');
      },
      plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true, ticks:{precision:0, color:PAL.muted}, grid:{color:PAL.grid}},
        y:{grid:{display:false}, afterFit:scale=>{ scale.width = 110; },
          ticks:{color:PAL.ink,
            callback:function(val){ const l=this.getLabelForValue(val); return l.length>14 ? l.slice(0,13)+'…' : l; }}}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = clientesOrd[els[0].index][0];
        filtroCliente = filtroCliente===val ? '' : val;
        document.getElementById('fCliente').value = filtroCliente;
        recalcularGraficos(); aplicarFiltrosChips();
      }}
  });

  const operadoraChart = new Chart(document.getElementById('chipsOperadoraChart'), {
    type:'doughnut', plugins:[centroRoscaPlugin],
    data:{labels:operadorasOrd.map(o=>o[0]),
      datasets:[{data:operadorasOrd.map(o=>o[1]), backgroundColor:operadorasOrd.map((o,i)=>operadoraCores[i%operadoraCores.length]), borderWidth:0, hoverOffset:8}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{tooltip:{position:'roscaFora'}, legend:{position:'bottom', labels:{color:'#E2E8F0', boxWidth:12, ...legendComContagem()}},
        centroRosca:{valor:fmtNum(C.lista.length), rotulo:'Chips'}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const val = operadorasOrd[els[0].index][0];
        filtroOperadora = filtroOperadora===val ? '' : val;
        document.getElementById('fOperadora').value = filtroOperadora;
        recalcularGraficos(); aplicarFiltrosChips();
      }}
  });

  const consumoChart = new Chart(document.getElementById('chipsConsumoChart'), {
    type:'doughnut', plugins:[centroRoscaPlugin],
    data:{labels:faixas.map(f=>f.label), datasets:[{data:faixas.map(f=>C.lista.filter(x=>x.pct>=f.min && x.pct<f.max).length),
      backgroundColor:faixas.map(f=>f.cor), borderWidth:0, hoverOffset:8}]},
    options:{maintainAspectRatio:false, onHover:(e,els)=>{e.native.target.style.cursor=els.length?'pointer':'default'},
      plugins:{tooltip:{position:'roscaFora'}, legend:{position:'bottom', labels:{color:'#E2E8F0', boxWidth:12, ...legendComContagem()}},
        centroRosca:{valor:fmtNum(C.lista.length), rotulo:'Chips'}},
      onClick:(e,els)=>{
        if(!els.length) return;
        const f = faixas[els[0].index];
        filtroFaixa = (filtroFaixa && filtroFaixa.label===f.label) ? null : f;
        recalcularGraficos(); aplicarFiltrosChips();
      }}
  });

  // --- Cross-filtering: cada gráfico recalcula sua própria dimensão ignorando o filtro
  // que ele mesmo controla, mas respeitando os demais (padrão dc.js/crossfilter). ---
  function filtrarExceto(excluir){
    const busca = document.getElementById('fBusca').value.trim().toLowerCase();
    return C.lista.filter(x=>{
      if (excluir!=='status' && filtroStatus && x.status !== filtroStatus) return false;
      if (excluir!=='cliente' && filtroCliente && x.cliente !== filtroCliente) return false;
      if (excluir!=='operadora' && filtroOperadora && x.operadora !== filtroOperadora) return false;
      if (excluir!=='cidade' && filtroCidade && x.cidade !== filtroCidade) return false;
      if (excluir!=='faixa' && filtroFaixa && !(x.pct>=filtroFaixa.min && x.pct<filtroFaixa.max)) return false;
      if (busca && !`${x.codigo} ${x.iccid} ${x.telefone} ${x.cliente} ${x.cidade}`.toLowerCase().includes(busca)) return false;
      return true;
    });
  }
  function contarPor(lista, campo){
    const m = {}; lista.forEach(x=>{ const k=x[campo]||'—'; m[k]=(m[k]||0)+1; }); return m;
  }
  function recalcularGraficos(){
    const porStatusF = contarPor(filtrarExceto('status'), 'status');
    statusChart.data.datasets[0].data = ['atualizado','atrasado','sem_comunicacao'].map(k=>porStatusF[k]||0);
    statusChart.data.datasets[0].backgroundColor = [
      ['atualizado',PAL.green],['atrasado',PAL.amber],['sem_comunicacao',PAL.red],
    ].map(([v,cor])=> (!filtroStatus||filtroStatus===v) ? cor : hexAlpha(cor,.25));
    // fatia selecionada -> contagem dela no centro; sem seleção -> total geral.
    statusChart.options.plugins.centroRosca.valor = fmtNum(filtroStatus ? (porStatusF[filtroStatus]||0) : Object.values(porStatusF).reduce((a,b)=>a+b,0));
    statusChart.update();

    const porClienteF = contarPor(filtrarExceto('cliente'), 'cliente');
    clienteChart.data.datasets[0].data = clientesOrd.map(c=>porClienteF[c[0]]||0);
    clienteChart.data.datasets[0].backgroundColor = clientesOrd.map(c=>
      (!filtroCliente||filtroCliente===c[0]) ? BLUE : hexAlpha(BLUE,.25));
    clienteChart.update();

    const porOperadoraF = contarPor(filtrarExceto('operadora'), 'operadora');
    operadoraChart.data.datasets[0].data = operadorasOrd.map(o=>porOperadoraF[o[0]]||0);
    operadoraChart.data.datasets[0].backgroundColor = operadorasOrd.map((o,i)=>{
      const cor = operadoraCores[i%operadoraCores.length];
      return (!filtroOperadora||filtroOperadora===o[0]) ? cor : hexAlpha(cor,.25);
    });
    operadoraChart.options.plugins.centroRosca.valor = fmtNum(filtroOperadora ? (porOperadoraF[filtroOperadora]||0) : Object.values(porOperadoraF).reduce((a,b)=>a+b,0));
    operadoraChart.update();

    const listaFaixaF = filtrarExceto('faixa');
    consumoChart.data.datasets[0].data = faixas.map(f=>listaFaixaF.filter(x=>x.pct>=f.min && x.pct<f.max).length);
    consumoChart.data.datasets[0].backgroundColor = faixas.map(f=>
      (!filtroFaixa||filtroFaixa.label===f.label) ? f.cor : hexAlpha(f.cor,.25));
    const nFaixaSel = filtroFaixa ? listaFaixaF.filter(x=>x.pct>=filtroFaixa.min && x.pct<filtroFaixa.max).length : listaFaixaF.length;
    consumoChart.options.plugins.centroRosca.valor = fmtNum(nFaixaSel);
    consumoChart.update();

    document.getElementById('cntTodos').textContent = filtrarExceto('status').length;
    document.getElementById('cntAtualizado').textContent = porStatusF.atualizado||0;
    document.getElementById('cntAtrasado').textContent = porStatusF.atrasado||0;
    document.getElementById('cntSemComunicacao').textContent = porStatusF.sem_comunicacao||0;
  }

  // --- Filtros de texto/select ---
  const clientesUnicos = [...new Set(C.lista.map(x=>x.cliente).filter(Boolean))].sort();
  document.getElementById('fCliente').innerHTML += clientesUnicos.map(c=>`<option value="${c}">${c}</option>`).join('');
  const operadorasUnicas = [...new Set(C.lista.map(x=>x.operadora).filter(Boolean))].sort();
  document.getElementById('fOperadora').innerHTML += operadorasUnicas.map(o=>`<option value="${o}">${o}</option>`).join('');
  const cidadesUnicas = [...new Set(C.lista.map(x=>x.cidade).filter(Boolean))].sort();
  document.getElementById('fCidade').innerHTML += cidadesUnicas.map(c=>`<option value="${c}">${c}</option>`).join('');

  // --- Colunas exibir/ocultar ---
  document.getElementById('colunasMenu').innerHTML =
    `<div class="colunasAcoes">
      <button type="button" data-acao="marcar">Marcar todos</button>
      <button type="button" data-acao="desmarcar">Desmarcar todos</button>
    </div>` +
    COLS.map((c,i)=>`<label><input type="checkbox" checked data-idx="${i+1}"> ${c.label}</label>`).join('');
  function aplicarColuna(chk){
    const idx = chk.dataset.idx, hidden = !chk.checked;
    document.querySelectorAll(`#tabelaChips th:nth-child(${idx}), #tabelaChips td:nth-child(${idx})`)
      .forEach(el=>el.style.display = hidden ? 'none' : '');
  }
  document.querySelectorAll('#colunasMenu input').forEach(chk=>{
    chk.addEventListener('change', ()=>aplicarColuna(chk));
  });
  document.querySelectorAll('#colunasMenu .colunasAcoes button').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const marcar = btn.dataset.acao === 'marcar';
      document.querySelectorAll('#colunasMenu input[type=checkbox]').forEach(chk=>{
        chk.checked = marcar; aplicarColuna(chk);
      });
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

  // --- Exportar CSV/XLSX (dos dados filtrados/ordenados na tela) ---
  document.getElementById('btnExportarCsv').addEventListener('click', ()=> exportarTabela('chips_consumo', COLS, ultimosFiltrados, 'csv'));
  document.getElementById('btnExportarXlsx').addEventListener('click', ()=> exportarTabela('chips_consumo', COLS, ultimosFiltrados, 'xlsx'));

  // --- Ordenação tri-estado nos cabeçalhos ---
  const ordenar = criarOrdenacao('tabelaChips', COLS, aplicarFiltrosChips);

  // --- Linha da tabela + ações ---
  const vazio = '<span class="vazio">—</span>';
  function linhaHtml(x){
    return `<tr>
      <td>${x.cliente||vazio}</td><td>${x.codigo}</td><td class="mono" style="font-size:11.5px">${x.iccid||vazio}</td><td>${x.telefone||vazio}</td>
      <td>${x.cidade||vazio}</td><td>${x.operadora||vazio}</td>
      <td>${x.ultimaComunicacao ? x.ultimaComunicacao.slice(0,16).replace('T',' ') : vazio}</td>
      <td><span class="badge ${x.status}"><span class="dot"></span>${STATUS_LABEL[x.status]}</span>${x.critico?' <span class="badge sem_comunicacao"><span class="dot"></span>CRÍTICO</span>':''}</td>
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
    if (filtroCidade) badges.push(['Cidade', filtroCidade, ()=>{filtroCidade='';document.getElementById('fCidade').value='';}]);
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
        fn(); recalcularGraficos(); aplicarFiltrosChips();
      });
    });
    document.getElementById('btnLimparFiltros').addEventListener('click', ()=>{
      filtroStatus=''; filtroCliente=''; filtroOperadora=''; filtroCidade=''; filtroFaixa=null;
      document.getElementById('fCliente').value=''; document.getElementById('fOperadora').value=''; document.getElementById('fCidade').value=''; document.getElementById('fBusca').value='';
      document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.toggle('active', b.dataset.status===''));
      recalcularGraficos(); aplicarFiltrosChips();
    });
  }

  // --- Aplicação combinada: filtros + ordenação + render ---
  function aplicarFiltrosChips(){
    const busca = document.getElementById('fBusca').value.trim().toLowerCase();
    let filtrados = C.lista.filter(x=>{
      if (filtroStatus && x.status !== filtroStatus) return false;
      if (filtroCliente && x.cliente !== filtroCliente) return false;
      if (filtroOperadora && x.operadora !== filtroOperadora) return false;
      if (filtroCidade && x.cidade !== filtroCidade) return false;
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
  document.getElementById('fBusca').addEventListener('input', ()=>{ recalcularGraficos(); aplicarFiltrosChips(); });
  document.getElementById('fCliente').addEventListener('change', e=>{ filtroCliente=e.target.value; recalcularGraficos(); aplicarFiltrosChips(); });
  document.getElementById('fOperadora').addEventListener('change', e=>{ filtroOperadora=e.target.value; recalcularGraficos(); aplicarFiltrosChips(); });
  document.getElementById('fCidade').addEventListener('change', e=>{ filtroCidade=e.target.value; recalcularGraficos(); aplicarFiltrosChips(); });
  document.querySelectorAll('#statusChips .chipBtn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      document.querySelectorAll('#statusChips .chipBtn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      filtroStatus = btn.dataset.status;
      recalcularGraficos();
      aplicarFiltrosChips();
    });
  });
  recalcularGraficos();
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

    # ponytail: cruza mês pra pegar regressão de comparar "registrada" como string
    # ("01-08" < "31-07" como texto, mesmo Ago vindo depois de Jul) - bug real, corrigido.
    linhas_mes = [
        ("EQP-09", "31-07-2026 23:57:52", "nao"),
        ("EQP-09", "01-08-2026 00:01:38", "nao"),
    ]
    d2 = agregar(linhas_mes, set())
    r = d2["ranking"][0]
    assert r["primeiro"] == "31-07-2026 23:57:52", r
    assert r["ultimo"] == "01-08-2026 00:01:38", r

    assert fmt_horas(208.2) == "8d 16h 12m", fmt_horas(208.2)
    assert fmt_horas(88.0) == "3d 16h", fmt_horas(88.0)
    assert fmt_horas(16.2) == "16h 12m", fmt_horas(16.2)

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
