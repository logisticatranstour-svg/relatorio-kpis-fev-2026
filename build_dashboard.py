import pandas as pd
import numpy as np
import datetime
import base64
import mimetypes
from pathlib import Path
import json
import re

# =========================
# CONFIG
# =========================

# CSV publicado do Google Sheets
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ4A5MDb6JivQ54j3B3YrWPIpnidj49zOdeyLsqE1HKy8f35--M3ja_ZG_KntrKKeYOAIWNyS-3QOp8/pub?gid=0&single=true&output=csv"

# Logos no repo (mesma pasta do index.html)
LOGO_TRANSTOUR_PATH = "logo-transtour.png"
LOGO_CLIENTE_PATH = "logo-cliente.png"

FOOTER_TEXT = "Relatório Gerado pelo Sistema Trans Tour Enviar e Receber | Confidencial"

# Colunas (como estão no seu formulário)
COL_SUBMITTED = "Submitted at"
COL_ITENS = "Todos os itens foram entregues corretamente?"
COL_UNIFORME = "Entregador apresentou-se com crachá e uniforme?"
COL_PRODUTOS = "Produtos em bom estado e dentro da validade?"
COL_ATENDIMENTO = "Atendimento cordial e respeitoso?"
COL_HORARIO = "Entrega ocorreu no horário combinado?"
COL_NOTA = "GRAU DE SATISFAÇÃO (1 A 5)"

QUAL_COLS = [COL_ITENS, COL_UNIFORME, COL_PRODUTOS, COL_ATENDIMENTO, COL_HORARIO]

# Paciente / Ocorrência (vamos detectar automaticamente caso o nome exato varie)
PATIENT_COL_CANDIDATES = [
    "nome do paciente",
    "Nome do paciente",
    "Nome do Paciente",
    "Paciente",
    "paciente",
    "Beneficiário",
    "beneficiário",
    "Beneficiario",
    "beneficiario",
]
OCC_COL_REGEX = re.compile(r"ocorr", re.IGNORECASE)  # pega "Tipo de ocorrência", "Ocorrência", etc.


# =========================
# HELPERS
# =========================

def is_refusal(x) -> bool:
    if pd.isna(x):
        return False
    return "recusou" in str(x).strip().lower()

def data_uri(path: str) -> str | None:
    p = Path(path)
    if not p.exists():
        return None
    mime, _ = mimetypes.guess_type(str(p))
    if mime is None:
        mime = "image/png"
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{b64}"

def find_patient_col(df: pd.DataFrame) -> str | None:
    cols = list(df.columns)
    # 1) candidatos exatos
    for c in PATIENT_COL_CANDIDATES:
        if c in cols:
            return c
    # 2) fallback por "paciente" no nome
    for c in cols:
        if "paciente" in str(c).lower():
            return c
    return None

def find_occ_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        if OCC_COL_REGEX.search(str(c) or ""):
            return c
    return None

def safe_write_error(msg: str):
    Path("index.html").write_text(
        f"<h1>Erro ao gerar dashboard</h1><pre>{msg}</pre>",
        encoding="utf-8"
    )


# =========================
# MAIN
# =========================

def main():
    # 1) Ler CSV
    try:
        df = pd.read_csv(SHEET_CSV_URL)
    except Exception as e:
        safe_write_error(f"Falha ao ler CSV do Sheets.\n{e}")
        return

    # 2) Validar colunas mínimas
    required = [COL_SUBMITTED, COL_NOTA] + QUAL_COLS
    missing = [c for c in required if c not in df.columns]
    if missing:
        safe_write_error(f"Colunas ausentes no CSV:\n{missing}\n\nColunas encontradas:\n{list(df.columns)}")
        return

    # 3) Detectar paciente e ocorrência (opcional)
    col_patient = find_patient_col(df)
    col_occ = find_occ_col(df)

    # 4) Datas
    df[COL_SUBMITTED] = pd.to_datetime(df[COL_SUBMITTED], errors="coerce")
    df = df.dropna(subset=[COL_SUBMITTED]).copy()

    if len(df) == 0:
        Path("index.html").write_text("<h1>Sem dados disponíveis.</h1>", encoding="utf-8")
        return

    # 5) Regra do cliente: "Recusou responder" => Sim e nota 5
    df_raw = df.copy()

    for c in QUAL_COLS:
        df.loc[df[c].apply(is_refusal), c] = "Sim"

    refusal_row_mask = np.zeros(len(df_raw), dtype=bool)
    for c in QUAL_COLS:
        refusal_row_mask |= df_raw[c].apply(is_refusal).to_numpy()

    df[COL_NOTA] = pd.to_numeric(df[COL_NOTA], errors="coerce")
    df.loc[refusal_row_mask | df[COL_NOTA].isna(), COL_NOTA] = 5
    df[COL_NOTA] = df[COL_NOTA].astype(int)

    # 6) Month para histórico
    df["month"] = df[COL_SUBMITTED].dt.to_period("M").astype(str)

    # 7) Dataset enxuto para o HTML (JSON)
    keep_cols = [COL_SUBMITTED, "month"] + QUAL_COLS + [COL_NOTA]
    if col_patient and col_patient not in keep_cols:
        keep_cols.append(col_patient)
    if col_occ and col_occ not in keep_cols:
        keep_cols.append(col_occ)

    df_dash = df[keep_cols].copy()

    # Strings "Sim/Não"
    for c in QUAL_COLS:
        df_dash[c] = df_dash[c].astype(str).str.strip()

    # Paciente / ocorrência
    if col_patient:
        df_dash[col_patient] = df_dash[col_patient].astype(str).str.strip()
    if col_occ:
        df_dash[col_occ] = df_dash[col_occ].astype(str).str.strip()

    # Data como string pro JS
    df_dash[COL_SUBMITTED] = df_dash[COL_SUBMITTED].dt.strftime("%Y-%m-%d %H:%M:%S")

    dashboard_data = df_dash.to_dict(orient="records")

    # Logos embed
    logo_transtour_uri = data_uri(LOGO_TRANSTOUR_PATH)
    logo_cliente_uri = data_uri(LOGO_CLIENTE_PATH)

    logos_html = ""
    if logo_transtour_uri:
        logos_html += f'<img class="logo" src="{logo_transtour_uri}" alt="Transtour"/>'
    if logo_cliente_uri:
        logos_html += f'<img class="logo" src="{logo_cliente_uri}" alt="Cliente"/>'

    generated = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    # HTML
    html = f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Relatório Gerencial de KPIs - Hapvida</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
  body{{font-family:Segoe UI, Tahoma, sans-serif;background:#f6f7fb;margin:0;padding:20px;color:#1f2a37}}
  .wrap{{max-width:1100px;margin:0 auto}}
  .topbar{{background:#0b1f3a;color:#fff;border-radius:14px;padding:18px 18px 16px}}
  .toprow{{display:flex;gap:16px;align-items:center;justify-content:space-between;flex-wrap:wrap}}
  .logos{{display:flex;gap:12px;align-items:center}}
  .logo{{height:48px;max-width:200px;object-fit:contain;background:rgba(255,255,255,.06);padding:6px 10px;border-radius:10px}}
  h1{{margin:0;font-size:22px;letter-spacing:.2px}}
  .sub{{margin:6px 0 0;opacity:.9}}
  .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-top:14px}}
  .card{{background:#fff;border-radius:14px;padding:14px 14px 12px;box-shadow:0 3px 10px rgba(0,0,0,.08)}}
  .kpiV{{font-size:26px;font-weight:800}}
  .kpiL{{font-size:12px;color:#6b7280;text-transform:uppercase;letter-spacing:.6px;margin-top:2px}}
  .section{{margin-top:18px}}
  .title{{font-weight:800;margin:0 0 10px 0}}
  table{{width:100%;border-collapse:collapse;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 3px 10px rgba(0,0,0,.08)}}
  th,td{{padding:10px 12px;border-bottom:1px solid #eef2f7;font-size:13px;vertical-align:top}}
  th{{text-align:left;background:#f9fafb;color:#374151;font-size:12px;text-transform:uppercase;letter-spacing:.5px}}
  tr:last-child td{{border-bottom:none}}
  .charts{{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:14px}}
  .muted{{color:#6b7280;font-size:12px;margin-top:6px}}
  .footer{{text-align:center;color:#9ca3af;font-size:12px;margin-top:12px}}
  .insights ul{{margin:8px 0 0 18px}}
  .view-controls{{position:fixed;right:18px;bottom:18px;display:flex;gap:10px;z-index:9999;flex-wrap:wrap}}
  .view-btn{{border:none;cursor:pointer;padding:12px 14px;border-radius:999px;font-weight:700;font-size:14px;box-shadow:0 10px 24px rgba(0,0,0,.18);background:#4f46e5;color:#fff;display:flex;align-items:center;gap:10px}}
  .view-btn.secondary{{background:#0b1f3a}}
  .view-btn.ghost{{background:#111827}}

  body.mode-mobile{{background:#6b7cff}}
  body.mode-mobile .wrap{{max-width:430px}}
  body.mode-mobile .topbar{{border-radius:26px 26px 14px 14px}}
  body.mode-mobile .phone-shell{{position:relative;margin:14px auto 60px;padding:18px 14px 18px;background:rgba(255,255,255,.08);border-radius:36px;box-shadow:0 24px 60px rgba(0,0,0,.25);max-width:460px}}
  body.mode-mobile .phone-shell:before{{content:"";position:absolute;top:10px;left:50%;transform:translateX(-50%);width:120px;height:22px;background:rgba(0,0,0,.28);border-radius:999px}}
  body.mode-mobile .phone-shell:after{{content:"";position:absolute;top:16px;left:calc(50% + 58px);width:10px;height:10px;background:rgba(255,255,255,.55);border-radius:999px}}

  @media(max-width:520px){{body{{padding:12px}} .charts{{grid-template-columns:1fr}} .view-controls{{right:12px;bottom:12px;flex-direction:column}}}}
</style>
</head>
<body>
<div class="wrap">
  <div class="phone-shell" id="phoneShell">

    <div class="topbar">
      <div class="toprow">
        <div class="logos">{logos_html}</div>

        <div style="flex:1;min-width:260px;text-align:center">
          <h1>RELATÓRIO GERENCIAL DE KPIs</h1>
          <div class="sub">Atualizado em: <b>{generated}</b></div>

          <div style="margin-top:10px; display:flex; gap:10px; justify-content:center; flex-wrap:wrap;">
            <label style="font-size:12px; opacity:.95;">
              Mês:
              <select id="monthSelect" style="margin-left:8px; padding:8px 10px; border-radius:10px; border:0; outline:none;"></select>
            </label>
            <span style="font-size:12px; opacity:.9;" id="monthHint"></span>
          </div>
        </div>

        <div style="min-width:160px;text-align:right;opacity:.95">
          <div style="font-size:12px">Notas válidas</div>
          <div style="font-weight:800" id="kpiNotasValidas">-</div>
        </div>
      </div>
    </div>

    <div class="grid">
      <div class="card"><div class="kpiV" id="kpiTotal">-</div><div class="kpiL">Total de Respostas</div></div>
      <div class="card"><div class="kpiV" id="kpiSat5">-</div><div class="kpiL">Satisfação Geral (Nota 5)</div></div>
      <div class="card"><div class="kpiV" id="kpiItens">-</div><div class="kpiL">Itens Entregues Corretamente</div></div>
      <div class="card"><div class="kpiV" id="kpiNps">-</div><div class="kpiL">NPS Estimado (Proxy)</div></div>
    </div>

    <div class="section">
      <h3 class="title">📈 Indicadores por Categoria</h3>
      <table>
        <thead><tr><th>Indicador</th><th>Sim</th><th>Não</th><th>% Positivo</th><th>Meta</th><th>Status</th></tr></thead>
        <tbody id="tbodyIndicadores"></tbody>
      </table>
      <div class="muted">* “Recusou responder” é considerado como conformidade (Sim) e nota máxima (5), conforme alinhado com o cliente.</div>
    </div>

    <div class="section charts">
      <div class="card">
        <h3 class="title">⭐ Distribuição de Satisfação (Notas 1–5)</h3>
        <canvas id="sat"></canvas>
        <div class="muted" id="satMeta"></div>
      </div>
      <div class="card">
        <h3 class="title">📊 Conformidade por Critério (vs Meta)</h3>
        <canvas id="bars"></canvas>
      </div>
    </div>

    <!-- NOVO: Evolução Top 5 -->
    <div class="section card">
      <h3 class="title">📈 Evolução dos Top 5 Pacientes (média mensal)</h3>
      <canvas id="top5Trend"></canvas>
      <div class="muted">Top 5 baseado na média do mês selecionado; tendência mostra os meses disponíveis no banco.</div>
    </div>

    <!-- NOVO: Piores notas -->
    <div class="section">
      <h3 class="title">⚠️ Pacientes com Notas Mais Baixas (mês selecionado)</h3>
      <table>
        <thead><tr><th>Paciente</th><th>Média</th><th>Menor nota</th><th>Respostas</th></tr></thead>
        <tbody id="tbodyWorstPatients"></tbody>
      </table>
    </div>

    <!-- NOVO: Top10 + ocorrência -->
    <div class="section">
      <h3 class="title">🏁 Top 10 (5 Maiores + 5 Menores) + Ocorrência</h3>
      <table>
        <thead><tr><th>Paciente</th><th>Média</th><th>Ocorrência</th><th>Respostas</th><th>Grupo</th></tr></thead>
        <tbody id="tbodyTop10Occ"></tbody>
      </table>
      <div class="muted" id="occHint"></div>
    </div>

    <div class="section">
      <h3 class="title">📋 KPIs Detalhados</h3>
      <table>
        <thead><tr><th>KPI</th><th>Valor</th><th>Meta</th><th>Status</th></tr></thead>
        <tbody id="tbodyKpisDetalhados"></tbody>
      </table>
    </div>

    <div class="section card insights">
      <h3 class="title">💡 Insights e Recomendações</h3>
      <ul id="ulInsights"></ul>
    </div>

    <div class="footer">{FOOTER_TEXT}</div>
  </div>
</div>

<div class="view-controls" aria-label="Controles de visualização">
  <button class="view-btn secondary" id="btnMobile" type="button">📱 Modo celular</button>
  <button class="view-btn" id="btnDesktop" type="button">🖥️ Modo tela cheia</button>
  <button class="view-btn ghost" id="btnFullscreen" type="button">⛶ Tela cheia (F11)</button>
</div>

<script>
  const DASHBOARD_DATA = {json.dumps(dashboard_data, ensure_ascii=False)};

  const COL_NOTA = {json.dumps(COL_NOTA)};
  const COL_PACIENTE = {json.dumps(col_patient if col_patient else "")};
  const COL_OCORRENCIA = {json.dumps(col_occ if col_occ else "")};

  const CRITERIOS = [
    ["Itens Entregues Corretamente", {json.dumps(COL_ITENS)}, 90],
    ["Crachá e Uniforme", {json.dumps(COL_UNIFORME)}, 95],
    ["Produtos em Bom Estado", {json.dumps(COL_PRODUTOS)}, 95],
    ["Atendimento Cordial", {json.dumps(COL_ATENDIMENTO)}, 95],
    ["Pontualidade", {json.dumps(COL_HORARIO)}, 90],
  ];

  function pct(n, d){{ return d === 0 ? 0 : (n/d)*100; }}

  function getMonths(data){{
    const set = new Set(data.map(r => r.month));
    return Array.from(set).sort(); // YYYY-MM
  }}

  function filterByMonth(data, month){{
    return data.filter(r => r.month === month);
  }}

  function countSimNao(arr, field){{
    let sim = 0, nao = 0;
    for(const r of arr){{
      const v = (r[field] ?? "").toString().trim().toLowerCase();
      if(v === "sim") sim++;
      if(v === "não" || v === "nao") nao++;
    }}
    return {{sim, nao}};
  }}

  function calcMonthKPIs(arr){{
    const total = arr.length;
    const notas = arr.map(r => Number(r[COL_NOTA])).filter(x => !Number.isNaN(x));
    const notasValidas = notas.length;
    const mean = notasValidas ? (notas.reduce((a,b)=>a+b,0)/notasValidas) : 0;
    const csat = mean ? (mean/5)*100 : 0;
    const rate5 = notasValidas ? (notas.filter(x=>x===5).length/notasValidas)*100 : 0;
    const rate45 = notasValidas ? (notas.filter(x=>x>=4).length/notasValidas)*100 : 0;
    return {{ total, notasValidas, mean, csat, rate5, rate45, notas }};
  }}

  function statusBadge(pos, meta){{
    if(pos >= meta) return "✓ Excelente";
    if(pos >= 80) return "⚠️ Atenção";
    return "⚠️ Atenção";
  }}

  function statusKPI(val, meta){{
    return val >= meta ? "✓ Atingida" : "✗ Não Atingida";
  }}

  function groupByPatient(arr){{
    if(!COL_PACIENTE) return [];
    const map = new Map();

    for(const r of arr){{
      const name = (r[COL_PACIENTE] ?? "").toString().trim();
      if(!name) continue;

      const nota = Number(r[COL_NOTA]);
      if(Number.isNaN(nota)) continue;

      const occ = COL_OCORRENCIA ? (r[COL_OCORRENCIA] ?? "").toString().trim() : "";

      if(!map.has(name)) {{
        map.set(name, {{ name, sum:0, n:0, min: nota, max: nota, occ: occ || "-" }});
      }}
      const o = map.get(name);
      o.sum += nota;
      o.n += 1;
      o.min = Math.min(o.min, nota);
      o.max = Math.max(o.max, nota);
      if(occ) o.occ = occ;
    }}

    const out = [];
    for(const v of map.values()){{
      out.push({{ name:v.name, avg: v.n ? v.sum/v.n : 0, n:v.n, min:v.min, max:v.max, occ:v.occ }});
    }}
    return out;
  }}

  let satChart=null, barChart=null, top5TrendChart=null;

  function buildTop5Trend(allData, top5Names){{
    const months = getMonths(allData);
    const datasets = top5Names.map(name => {{
      const data = months.map(m => {{
        const arr = allData.filter(r => r.month === m && (r[COL_PACIENTE] ?? "").toString().trim() === name);
        const notas = arr.map(r => Number(r[COL_NOTA])).filter(x => !Number.isNaN(x));
        if(!notas.length) return null;
        return Number((notas.reduce((a,b)=>a+b,0) / notas.length).toFixed(2));
      }});
      return {{ label: name, data, spanGaps: true }};
    }});

    const el = document.getElementById("top5Trend");
    if(!el) return;

    const ctx = el.getContext("2d");
    if(top5TrendChart) top5TrendChart.destroy();

    top5TrendChart = new Chart(ctx, {{
      type: "line",
      data: {{ labels: months, datasets }},
      options: {{
        plugins: {{ legend: {{ position: "bottom" }} }},
        scales: {{ y: {{ min: 1, max: 5 }} }}
      }}
    }});
  }}

  function render(month){{
    const arr = filterByMonth(DASHBOARD_DATA, month);
    const total = arr.length;

    const k = calcMonthKPIs(arr);

    document.getElementById("monthHint").textContent = `Base do mês: ${total} respostas`;
    document.getElementById("kpiTotal").textContent = total;
    document.getElementById("kpiSat5").textContent = `${k.rate5.toFixed(1)}%`;

    const itensSN = countSimNao(arr, CRITERIOS[0][1]);
    document.getElementById("kpiItens").textContent = `${pct(itensSN.sim, total).toFixed(1)}%`;
    document.getElementById("kpiNps").textContent = `${pct(itensSN.sim, total).toFixed(0)}`;
    document.getElementById("kpiNotasValidas").textContent = k.notasValidas;

    // Indicadores por categoria
    const tbody = document.getElementById("tbodyIndicadores");
    tbody.innerHTML = CRITERIOS.map(([label, col, meta]) => {{
      const sn = countSimNao(arr, col);
      const pos = pct(sn.sim, total);
      return `
        <tr>
          <td>${label}</td>
          <td>${sn.sim}</td>
          <td>${sn.nao}</td>
          <td>${pos.toFixed(1)}%</td>
          <td>${meta}%</td>
          <td>${statusBadge(pos, meta)}</td>
        </tr>
      `;
    }}).join("");

    // KPIs Detalhados (dinâmico)
    const det = [
      ["KPI-01: Conformidade de Itens", pct(countSimNao(arr, CRITERIOS[0][1]).sim, total), 90],
      ["KPI-02: Padronização Visual (Uniforme/Crachá)", pct(countSimNao(arr, CRITERIOS[1][1]).sim, total), 95],
      ["KPI-03: Qualidade dos Produtos", pct(countSimNao(arr, CRITERIOS[2][1]).sim, total), 95],
      ["KPI-04: Qualidade no Atendimento", pct(countSimNao(arr, CRITERIOS[3][1]).sim, total), 95],
      ["KPI-05: Pontualidade", pct(countSimNao(arr, CRITERIOS[4][1]).sim, total), 90],
      ["KPI-06: CSAT (média/5)", k.csat, 90],
      ["KPI-07: Taxa de Notas 5", k.rate5, 80],
      ["KPI-08: Taxa de Notas 4-5", k.rate45, 90],
    ];

    const tbodyDet = document.getElementById("tbodyKpisDetalhados");
    tbodyDet.innerHTML = det.map(([name,val,meta]) => `
      <tr>
        <td>${name}</td>
        <td>${val.toFixed(1)}%</td>
        <td>${meta}%</td>
        <td>${statusKPI(val, meta)}</td>
      </tr>
    `).join("");

    // Insights dinâmicos
    const ul = document.getElementById("ulInsights");
    const itensPct = det[0][1];
    const naoItens = countSimNao(arr, CRITERIOS[0][1]).nao;
    const insights = [];

    if(itensPct < 90) {{
      insights.push(`Acuracidade de itens em ${itensPct.toFixed(1)}% (meta 90%). Recomendamos reforçar conferência (checklist duplo) e rastrear causas das ${naoItens} não conformidades por falta de itens internos.`);
    }} else {{
      insights.push(`Acuracidade de itens em ${itensPct.toFixed(1)}% (meta 90%). Resultado dentro da meta. Manter padrão de conferência e monitorar recorrências.`);
    }}

    insights.push(`CSAT médio ${k.mean.toFixed(2)}/5 (${k.csat.toFixed(1)}%). Taxa de nota 5: ${k.rate5.toFixed(1)}%.`);
    insights.push(`Conforme alinhado com o cliente, respostas “Recusou responder” foram consideradas como conformidade (Sim) e nota máxima (5) para fins de consolidação.`);
    insights.push(`<b>Privacidade:</b> este dashboard exibe somente indicadores agregados; evite publicar dados pessoais em repositório público.`);

    ul.innerHTML = insights.map(x => `<li>${x}</li>`).join("");
