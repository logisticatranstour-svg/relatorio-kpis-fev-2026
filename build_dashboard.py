from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd

# =========================
# CONFIGURAÇÕES PRINCIPAIS
# =========================

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ4A5MDb6JivQ54j3B3YrWPIpnidj49zOdeyLsqE1HKy8f35--M3ja_ZG_KntrKKeYOAIWNyS-3QOp8/pub?gid=0&single=true&output=csv"
OUTPUT_HTML = "index.html"

LOGO_TRANSTOUR = "logo-transtour.png"
LOGO_CLIENTE = "logo-cliente.png"

PAGE_TITLE = "RELATÓRIO GERENCIAL DE KPIs"

METAS = {
    "Itens Entregues Corretamente": 98.0,
    "Crachá e Uniforme": 97.0,
    "Produtos em Bom Estado": 97.0,
    "Atendimento Cordial": 98.0,
    "Pontualidade": 95.0,
    "CSAT (média/5)": 90.0,
    "Taxa de Notas 5": 80.0,
    "Taxa de Notas 4-5": 90.0,
}

COLUNAS = {
    "data": "Submitted at",
    "paciente": "Nome do Paciente",
    "nota": "GRAU DE SATISFAÇÃO (1 A 5)",
    "itens": "Todos os itens foram entregues corretamente?",
    "uniforme": "Entregador apresentou-se com crachá e uniforme?",
    "produtos": "Produtos em bom estado e dentro da validade?",
    "atendimento": "Atendimento cordial e respeitoso?",
    "pontualidade": "Entrega ocorreu no horário combinado?",
}

# =========================
# FUNÇÕES AUXILIARES
# =========================


def norm_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def norm_text_lower(value) -> str:
    return norm_text(value).lower()


def canonical_yes_no(value: str) -> str:
    txt = norm_text_lower(value)
    if not txt:
        return ""
    if txt in {"sim", "s", "yes", "y", "true", "1"}:
        return "Sim"
    if txt in {"não", "nao", "n", "no", "false", "0"}:
        return "Não"
    if "recusou responder" in txt:
        return "Sim"
    return norm_text(value)


def parse_score(value):
    txt = norm_text(value)
    if not txt:
        return None
    if "recusou responder" in txt.lower():
        return 5
    try:
        score = float(str(txt).replace(",", "."))
        if 1 <= score <= 5:
            return score
    except Exception:
        return None
    return None


def status_text(valor: float, meta: float) -> str:
    if valor >= meta:
        return "✓ Excelente"
    if valor >= meta - 10:
        return "⚠ Atenção"
    return "X Não Atingida"


def build_occurrence(row: pd.Series) -> str:
    falhas = []

    if row["itens_resp"] == "Não":
        falhas.append("Itens entregues incorretamente")
    if row["uniforme_resp"] == "Não":
        falhas.append("Crachá/Uniforme")
    if row["produtos_resp"] == "Não":
        falhas.append("Produtos/Validade")
    if row["atendimento_resp"] == "Não":
        falhas.append("Atendimento")
    if row["pontualidade_resp"] == "Não":
        falhas.append("Pontualidade")

    if not falhas:
        return "Sem ocorrência"
    return " / ".join(falhas)


def detect_refused_response(row: pd.Series) -> bool:
    # Como não existe coluna específica, qualquer "Recusou responder"
    # nas respostas principais da pesquisa será tratado como satisfatório.
    campos_verificacao = [
        COLUNAS["itens"],
        COLUNAS["uniforme"],
        COLUNAS["produtos"],
        COLUNAS["atendimento"],
        COLUNAS["pontualidade"],
        COLUNAS["nota"],
    ]

    for col in campos_verificacao:
        if col in row.index and "recusou responder" in norm_text_lower(row[col]):
            return True
    return False


def month_sort_key(month_str: str):
    return datetime.strptime(month_str, "%Y-%m")


def pct(value: float) -> float:
    return round(float(value), 1)


# =========================
# PROCESSAMENTO PRINCIPAL
# =========================


def load_data() -> pd.DataFrame:
    df = pd.read_csv(SHEET_CSV_URL)

    missing = [v for v in COLUNAS.values() if v not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias não encontradas no CSV: {missing}")

    df = df.copy()

    df["refused_response"] = df.apply(detect_refused_response, axis=1)

    df["date"] = pd.to_datetime(df[COLUNAS["data"]], errors="coerce")
    df["month"] = df["date"].dt.strftime("%Y-%m")

    df["patient"] = (
        df[COLUNAS["paciente"]]
        .fillna("NÃO INFORMADO")
        .astype(str)
        .str.strip()
        .replace("", "NÃO INFORMADO")
        .str.upper()
    )

    df["score"] = df.apply(
        lambda r: 5 if r["refused_response"] else parse_score(r[COLUNAS["nota"]]),
        axis=1,
    )

    df["itens_resp"] = df.apply(
        lambda r: "Sim" if r["refused_response"] else canonical_yes_no(r[COLUNAS["itens"]]),
        axis=1,
    )
    df["uniforme_resp"] = df.apply(
        lambda r: "Sim" if r["refused_response"] else canonical_yes_no(r[COLUNAS["uniforme"]]),
        axis=1,
    )
    df["produtos_resp"] = df.apply(
        lambda r: "Sim" if r["refused_response"] else canonical_yes_no(r[COLUNAS["produtos"]]),
        axis=1,
    )
    df["atendimento_resp"] = df.apply(
        lambda r: "Sim" if r["refused_response"] else canonical_yes_no(r[COLUNAS["atendimento"]]),
        axis=1,
    )
    df["pontualidade_resp"] = df.apply(
        lambda r: "Sim" if r["refused_response"] else canonical_yes_no(r[COLUNAS["pontualidade"]]),
        axis=1,
    )

    df["occurrence_type"] = df.apply(build_occurrence, axis=1)

    df = df[df["date"].notna()].copy()
    df = df[df["month"].notna()].copy()
    df = df[df["score"].notna()].copy()

    return df


def build_monthly_kpis(df: pd.DataFrame) -> List[Dict]:
    output = []

    for month, g in df.groupby("month", dropna=True):
        total = len(g)
        media = g["score"].mean()
        taxa_5 = (g["score"].eq(5).mean() * 100) if total else 0
        taxa_45 = (g["score"].ge(4).mean() * 100) if total else 0
        taxa_13 = (g["score"].le(3).mean() * 100) if total else 0

        output.append(
            {
                "month": month,
                "total_respostas": int(total),
                "notas_validas": int(total),
                "media_nota": round(float(media), 2),
                "csat_45": pct(taxa_45),
                "taxa_nota_5": pct(taxa_5),
                "taxa_notas_45": pct(taxa_45),
                "taxa_notas_13": pct(taxa_13),
            }
        )

    return sorted(output, key=lambda x: month_sort_key(x["month"]))


def build_category_indicators(df: pd.DataFrame) -> List[Dict]:
    specs = [
        ("Itens Entregues Corretamente", "itens_resp"),
        ("Crachá e Uniforme", "uniforme_resp"),
        ("Produtos em Bom Estado", "produtos_resp"),
        ("Atendimento Cordial", "atendimento_resp"),
        ("Pontualidade", "pontualidade_resp"),
    ]

    output = []

    for month, g in df.groupby("month", dropna=True):
        for label, col in specs:
            sim = int(g[col].eq("Sim").sum())
            nao = int(g[col].eq("Não").sum())
            total = sim + nao
            positivo = (sim / total * 100) if total else 0

            output.append(
                {
                    "month": month,
                    "indicador": label,
                    "sim": sim,
                    "nao": nao,
                    "positivo": pct(positivo),
                    "meta": METAS[label],
                }
            )

    return sorted(output, key=lambda x: (month_sort_key(x["month"]), x["indicador"]))


def build_detailed_kpis(df: pd.DataFrame, categories: List[Dict]) -> List[Dict]:
    output = []

    category_map = {}
    for item in categories:
        category_map.setdefault(item["month"], []).append(item)

    monthly_lookup = {item["month"]: item for item in build_monthly_kpis(df)}

    for month in sorted(df["month"].dropna().unique().tolist(), key=month_sort_key):
        items = {x["indicador"]: x["positivo"] for x in category_map.get(month, [])}
        mk = monthly_lookup[month]

        rows = [
            ("KPI-01: Conformidade de Itens", items.get("Itens Entregues Corretamente", 0), METAS["Itens Entregues Corretamente"]),
            ("KPI-02: Padronização Visual (Uniforme/Crachá)", items.get("Crachá e Uniforme", 0), METAS["Crachá e Uniforme"]),
            ("KPI-03: Qualidade dos Produtos", items.get("Produtos em Bom Estado", 0), METAS["Produtos em Bom Estado"]),
            ("KPI-04: Qualidade no Atendimento", items.get("Atendimento Cordial", 0), METAS["Atendimento Cordial"]),
            ("KPI-05: Pontualidade", items.get("Pontualidade", 0), METAS["Pontualidade"]),
            ("KPI-06: CSAT (média/5)", mk["csat_45"], METAS["CSAT (média/5)"]),
            ("KPI-07: Taxa de Notas 5", mk["taxa_nota_5"], METAS["Taxa de Notas 5"]),
            ("KPI-08: Taxa de Notas 4-5", mk["taxa_notas_45"], METAS["Taxa de Notas 4-5"]),
        ]

        for nome, valor, meta in rows:
            output.append(
                {
                    "month": month,
                    "kpi": nome,
                    "valor": pct(valor),
                    "meta": float(meta),
                    "status": status_text(valor, meta),
                }
            )

    return output


def build_records(df: pd.DataFrame) -> List[Dict]:
    records = df[["month", "date", "patient", "score", "occurrence_type"]].copy()
    records["date"] = records["date"].dt.strftime("%Y-%m-%d")

    return [
        {
            "month": row["month"],
            "date": row["date"],
            "patient": row["patient"],
            "score": round(float(row["score"]), 2),
            "occurrence_type": row["occurrence_type"],
        }
        for _, row in records.iterrows()
    ]


def build_dashboard_data(df: pd.DataFrame) -> Dict:
    monthly_kpis = build_monthly_kpis(df)
    category_indicators = build_category_indicators(df)
    detailed_kpis = build_detailed_kpis(df, category_indicators)
    records = build_records(df)

    months = [item["month"] for item in monthly_kpis]

    return {
        "generated_at": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),
        "months": months,
        "monthly_kpis": monthly_kpis,
        "category_indicators": category_indicators,
        "detailed_kpis": detailed_kpis,
        "records": records,
    }


# =========================
# TEMPLATE HTML
# =========================

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>__PAGE_TITLE__</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root{
      --bg:#eef2f7;
      --card:#ffffff;
      --navy:#0b2347;
      --navy-2:#102e5c;
      --line:#e5e7eb;
      --text:#1f2937;
      --muted:#6b7280;
      --success:#15803d;
      --warning:#d97706;
      --danger:#b91c1c;
      --shadow:0 10px 24px rgba(16,24,40,.08);
      --radius:18px;
    }
    *{box-sizing:border-box}
    body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--text)}
    .container{max-width:1400px;margin:0 auto;padding:0 16px 28px}
    .header{margin-top:12px;background:linear-gradient(180deg,var(--navy),var(--navy-2));border-radius:0 0 18px 18px;color:#fff;padding:18px 16px;box-shadow:var(--shadow)}
    .header-grid{display:grid;grid-template-columns:200px 1fr 200px;gap:16px;align-items:center}
    .logo-slot{display:flex;align-items:center;justify-content:center;min-height:92px}
    .logo-box{width:180px;height:78px;border-radius:14px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);display:flex;align-items:center;justify-content:center;padding:10px;overflow:hidden;backdrop-filter:blur(3px)}
    .logo-box img{max-width:100%;max-height:58px;object-fit:contain;display:block}
    .logo-fallback{text-align:center;font-weight:700;line-height:1.2;font-size:14px;color:#dbeafe;letter-spacing:.4px}
    .title-wrap{text-align:center}
    .title-wrap h1{margin:0 0 4px;font-size:22px;font-weight:800;text-transform:uppercase}
    .title-wrap .updated{font-size:14px;font-weight:700;margin-bottom:8px}
    .header-controls{display:flex;flex-wrap:wrap;gap:10px 14px;align-items:center;justify-content:center;font-size:14px}
    .header-controls label{font-weight:700}
    .header-controls select{border:none;border-radius:10px;padding:10px 12px;font-size:14px;min-width:120px;color:#111827}
    .toolbar{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;margin:16px 0 18px}
    .pill-btn{background:#0f1f3a;color:#fff;border:none;border-radius:999px;padding:14px 18px;font-weight:700;cursor:pointer;box-shadow:var(--shadow);transition:.2s ease;font-size:14px}
    .pill-btn:hover{transform:translateY(-1px)}
    .pill-btn.active{background:#4f46e5}
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:16px}
    .card{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:18px 18px 16px}
    .kpi small{display:block;color:var(--muted);font-weight:700;margin-bottom:8px;font-size:13px}
    .kpi .value{font-size:32px;font-weight:800;color:var(--navy)}
    .kpi .sub{margin-top:6px;color:var(--muted);font-size:13px}
    .section-title{font-size:17px;font-weight:800;margin:0 0 12px;color:var(--navy);display:flex;align-items:center;gap:8px}
    .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
    .grid-1{display:grid;grid-template-columns:1fr;gap:16px;margin-bottom:16px}
    .table-wrap{overflow-x:auto}
    table{width:100%;min-width:620px;border-collapse:collapse}
    thead th{text-align:left;font-size:12px;letter-spacing:.4px;text-transform:uppercase;color:#334155;padding:12px;border-bottom:1px solid var(--line);background:#f8fafc}
    tbody td{padding:11px 12px;border-bottom:1px solid var(--line);font-size:14px;vertical-align:top}
    tbody tr:hover{background:#fafcff}
    .status{font-weight:700;white-space:nowrap}
    .status.success{color:var(--success)}
    .status.warning{color:var(--warning)}
    .status.danger{color:var(--danger)}
    .meta-note{margin-top:8px;color:var(--muted);font-size:12px}
    .chart-box{height:310px;position:relative}
    .chart-note{margin-top:8px;color:var(--muted);font-size:13px;font-weight:700}
    .badge-score{display:inline-block;padding:4px 10px;border-radius:999px;font-weight:800;min-width:54px;text-align:center;color:#fff}
    .score-low{background:var(--danger)}
    .score-mid{background:var(--warning)}
    .score-high{background:var(--success)}
    .insights{margin:0;padding-left:22px}
    .insights li{margin:6px 0;line-height:1.45;font-size:15px}
    .footer{text-align:left;color:#64748b;font-size:13px;padding:4px 2px 0}
    body.mobile-preview .container{max-width:430px;padding:0 10px 20px}
    body.mobile-preview .header-grid{grid-template-columns:1fr}
    body.mobile-preview .logo-slot{min-height:auto}
    body.mobile-preview .title-wrap{order:2}
    body.mobile-preview .kpis{grid-template-columns:1fr}
    body.mobile-preview .grid-2{grid-template-columns:1fr}
    body.mobile-preview .logo-box{width:150px;height:70px}
    @media (max-width:1100px){
      .header-grid{grid-template-columns:150px 1fr 150px}
      .kpis{grid-template-columns:repeat(2,1fr)}
      .grid-2{grid-template-columns:1fr}
      .logo-box{width:140px}
    }
    @media (max-width:760px){
      .container{padding:0 10px 20px}
      .header-grid{grid-template-columns:1fr}
      .logo-slot{min-height:auto}
      .kpis{grid-template-columns:1fr}
      .title-wrap h1{font-size:19px}
      .kpi .value{font-size:28px}
    }
  </style>
</head>
<body>
  <div class="container">
    <header class="header">
      <div class="header-grid">
        <div class="logo-slot">
          <div class="logo-box" id="logoTransTourBox">
            <img src="__LOGO_TRANSTOUR__" alt="Logo Trans Tour" onerror="this.parentElement.innerHTML='<div class=&quot;logo-fallback&quot;>LOGO<br>TRANS TOUR</div>';">
          </div>
        </div>

        <div class="title-wrap">
          <h1>__PAGE_TITLE__</h1>
          <div class="updated" id="updatedAt">Atualizado em: --</div>
          <div class="header-controls">
            <label for="monthSelect">Mês:</label>
            <select id="monthSelect"></select>
            <span>Base do mês: <strong id="baseMes">0 respostas</strong></span>
          </div>
        </div>

        <div class="logo-slot" style="justify-content:flex-end">
          <div class="logo-box" id="logoClientBox">
            <img src="__LOGO_CLIENTE__" alt="Logo Cliente" onerror="this.parentElement.innerHTML='<div class=&quot;logo-fallback&quot;>LOGO<br>CLIENTE</div>';">
          </div>
        </div>
      </div>
    </header>

    <div class="toolbar">
      <button class="pill-btn" id="btnMobile">📱 Modo celular</button>
      <button class="pill-btn active" id="btnNormal">🖥️ Modo tela cheia</button>
      <button class="pill-btn" id="btnFullscreen">⛶ Tela cheia (F11)</button>
    </div>

    <section class="kpis">
      <div class="card kpi">
        <small>Total de respostas</small>
        <div class="value" id="kpiTotal">0</div>
        <div class="sub">Respostas válidas no mês selecionado</div>
      </div>
      <div class="card kpi">
        <small>Média da nota</small>
        <div class="value" id="kpiMedia">0.00</div>
        <div class="sub" id="kpiMediaSub">CSAT: 0%</div>
      </div>
      <div class="card kpi">
        <small>Taxa de Notas 5</small>
        <div class="value" id="kpiNota5">0%</div>
        <div class="sub">Meta oficial: 80%</div>
      </div>
      <div class="card kpi">
        <small>Taxa de Notas 4–5</small>
        <div class="value" id="kpiNota45">0%</div>
        <div class="sub">Meta oficial: 90%</div>
      </div>
    </section>

    <section class="grid-1">
      <div class="card">
        <h2 class="section-title">📈 Indicadores por Categoria</h2>
        <div class="table-wrap" id="categoryTable"></div>
        <div class="meta-note">* No tratamento da Trans Tour, qualquer resposta contendo “Recusou responder” é considerada satisfatória, com nota 5 e critérios favoráveis.</div>
      </div>
    </section>

    <section class="grid-2">
      <div class="card">
        <h2 class="section-title">⭐ Distribuição de Satisfação (Notas 1–5)</h2>
        <div class="chart-box"><canvas id="distributionChart"></canvas></div>
        <div class="chart-note" id="distributionNote">Média: 0/5 • CSAT: 0%</div>
      </div>
      <div class="card">
        <h2 class="section-title">📊 Conformidade por Critério (vs Meta)</h2>
        <div class="chart-box"><canvas id="criteriaChart"></canvas></div>
      </div>
    </section>

    <section class="grid-2">
      <div class="card">
        <h2 class="section-title">🔻 Pacientes com Notas Mais Baixas</h2>
        <div class="table-wrap" id="lowestPatientsTable"></div>
      </div>
      <div class="card">
        <h2 class="section-title">🏅 Top 10 Pacientes (5 Maiores + 5 Menores)</h2>
        <div class="table-wrap" id="topBottomTable"></div>
      </div>
    </section>

    <section class="grid-2">
      <div class="card">
        <h2 class="section-title">📉 Evolução dos Top 5 Pacientes</h2>
        <div class="chart-box"><canvas id="trendChart"></canvas></div>
      </div>
      <div class="card">
        <h2 class="section-title">📋 KPIs Detalhados</h2>
        <div class="table-wrap" id="detailedKpisTable"></div>
      </div>
    </section>

    <section class="grid-1">
      <div class="card">
        <h2 class="section-title">💡 Insights e Recomendações</h2>
        <ul class="insights" id="insightsList"></ul>
      </div>
    </section>

    <div class="footer">
      Relatório Gerado pelo Sistema Trans Tour Enviar e Receber | Confidencial
    </div>
  </div>

<script>
const dashboard_data = __DASHBOARD_JSON__;

const monthSelect = document.getElementById("monthSelect");
const bodyEl = document.body;
let distributionChart = null;
let criteriaChart = null;
let trendChart = null;

function formatMonth(month) {
  const [year, m] = month.split("-");
  const names = {"01":"Jan","02":"Fev","03":"Mar","04":"Abr","05":"Mai","06":"Jun","07":"Jul","08":"Ago","09":"Set","10":"Out","11":"Nov","12":"Dez"};
  return `${names[m]}/${year}`;
}

function pct(v) {
  return `${Number(v || 0).toFixed(1)}%`;
}

function statusClass(result, meta) {
  if (result >= meta) return "success";
  if (result >= meta - 10) return "warning";
  return "danger";
}

function statusText(result, meta) {
  if (result >= meta) return "✓ Excelente";
  if (result >= meta - 10) return "⚠ Atenção";
  return "X Não Atingida";
}

function scoreClass(score) {
  if (score <= 2) return "score-low";
  if (score === 3) return "score-mid";
  return "score-high";
}

function getMonthKpi(month) {
  return dashboard_data.monthly_kpis.find(x => x.month === month) || {
    total_respostas:0, media_nota:0, csat_45:0, taxa_nota_5:0, taxa_notas_45:0
  };
}

function getMonthRecords(month) {
  return dashboard_data.records.filter(x => x.month === month);
}

function getCategoryIndicators(month) {
  return dashboard_data.category_indicators.filter(x => x.month === month);
}

function getDetailedKpis(month) {
  return dashboard_data.detailed_kpis.filter(x => x.month === month);
}

function patientAverages(month) {
  const map = new Map();
  getMonthRecords(month).forEach(r => {
    const patient = (r.patient || "NÃO INFORMADO").trim();
    const score = Number(r.score || 0);
    const occ = (r.occurrence_type || "Sem ocorrência").trim();
    if (!map.has(patient)) {
      map.set(patient, { patient, total:0, count:0, occurrences:{} });
    }
    const item = map.get(patient);
    item.total += score;
    item.count += 1;
    item.occurrences[occ] = (item.occurrences[occ] || 0) + 1;
  });

  return Array.from(map.values()).map(item => {
    const topOcc = Object.entries(item.occurrences).sort((a,b) => b[1]-a[1])[0];
    return {
      patient:item.patient,
      avgScore:Number((item.total / item.count).toFixed(2)),
      responses:item.count,
      occurrenceType:topOcc ? topOcc[0] : "Sem ocorrência"
    };
  });
}

function renderHeaderInfo(month) {
  document.getElementById("updatedAt").textContent = `Atualizado em: ${dashboard_data.generated_at}`;
  document.getElementById("baseMes").textContent = `${getMonthKpi(month).total_respostas} respostas`;
}

function renderKpis(month) {
  const k = getMonthKpi(month);
  document.getElementById("kpiTotal").textContent = k.total_respostas;
  document.getElementById("kpiMedia").textContent = Number(k.media_nota).toFixed(2);
  document.getElementById("kpiMediaSub").textContent = `CSAT: ${pct(k.csat_45)}`;
  document.getElementById("kpiNota5").textContent = pct(k.taxa_nota_5);
  document.getElementById("kpiNota45").textContent = pct(k.taxa_notas_45);
}

function renderCategoryTable(month) {
  const rows = getCategoryIndicators(month);
  const target = document.getElementById("categoryTable");
  target.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Indicador</th>
          <th>Sim</th>
          <th>Não</th>
          <th>% Positivo</th>
          <th>Meta</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td>${r.indicador}</td>
            <td>${r.sim}</td>
            <td>${r.nao}</td>
            <td>${pct(r.positivo)}</td>
            <td>${pct(r.meta)}</td>
            <td class="status ${statusClass(r.positivo, r.meta)}">${statusText(r.positivo, r.meta)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderDistributionChart(month) {
  const monthRecords = getMonthRecords(month);
  const counts = {1:0,2:0,3:0,4:0,5:0};
  monthRecords.forEach(r => {
    const s = Number(r.score || 0);
    if (counts[s] !== undefined) counts[s] += 1;
  });

  const k = getMonthKpi(month);
  document.getElementById("distributionNote").textContent = `Média: ${Number(k.media_nota).toFixed(2)}/5 • CSAT: ${pct(k.csat_45)}`;

  if (distributionChart) distributionChart.destroy();
  distributionChart = new Chart(document.getElementById("distributionChart"), {
    type:"bar",
    data:{
      labels:["1","2","3","4","5"],
      datasets:[{
        label:"Qtd. respostas",
        data:[counts[1],counts[2],counts[3],counts[4],counts[5]],
        backgroundColor:"rgba(96,165,250,.55)",
        borderColor:"rgba(96,165,250,1)",
        borderWidth:1
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{y:{beginAtZero:true}}
    }
  });
}

function renderCriteriaChart(month) {
  const rows = getCategoryIndicators(month);

  if (criteriaChart) criteriaChart.destroy();
  criteriaChart = new Chart(document.getElementById("criteriaChart"), {
    type:"bar",
    data:{
      labels: rows.map(r => r.indicador),
      datasets:[
        {label:"Resultado (%)", data: rows.map(r => r.positivo), backgroundColor:"rgba(125, 191, 236, .8)"},
        {label:"Meta (%)", data: rows.map(r => r.meta), backgroundColor:"rgba(240, 168, 188, .85)"}
      ]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      plugins:{legend:{position:"bottom"}},
      scales:{y:{beginAtZero:true, max:100}, x:{ticks:{maxRotation:18, minRotation:18}}}
    }
  });
}

function renderLowestPatients(month) {
  const rows = [...patientAverages(month)].sort((a,b) => a.avgScore - b.avgScore || b.responses - a.responses).slice(0,10);
  const target = document.getElementById("lowestPatientsTable");

  target.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Paciente</th>
          <th>Média</th>
          <th>Respostas</th>
          <th>Ocorrência</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td>${r.patient}</td>
            <td><span class="badge-score ${scoreClass(Math.round(r.avgScore))}">${r.avgScore.toFixed(2)}</span></td>
            <td>${r.responses}</td>
            <td>${r.occurrenceType}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderTopBottomTable(month) {
  const all = patientAverages(month);
  const top5 = [...all].sort((a,b) => b.avgScore - a.avgScore || b.responses - a.responses).slice(0,5);
  const bottom5 = [...all].sort((a,b) => a.avgScore - b.avgScore || b.responses - a.responses).slice(0,5);
  const target = document.getElementById("topBottomTable");

  target.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Grupo</th>
          <th>Paciente</th>
          <th>Média</th>
          <th>Ocorrência</th>
        </tr>
      </thead>
      <tbody>
        ${top5.map(r => `
          <tr>
            <td>Top 5</td>
            <td>${r.patient}</td>
            <td><span class="badge-score ${scoreClass(Math.round(r.avgScore))}">${r.avgScore.toFixed(2)}</span></td>
            <td>${r.occurrenceType}</td>
          </tr>
        `).join("")}
        ${bottom5.map(r => `
          <tr>
            <td>Bottom 5</td>
            <td>${r.patient}</td>
            <td><span class="badge-score ${scoreClass(Math.round(r.avgScore))}">${r.avgScore.toFixed(2)}</span></td>
            <td>${r.occurrenceType}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderTrendChart(month) {
  const selectedPatients = [...patientAverages(month)]
    .sort((a,b) => a.avgScore - b.avgScore || b.responses - a.responses)
    .slice(0,5)
    .map(x => x.patient);

  const labels = dashboard_data.months.map(formatMonth);
  const datasets = selectedPatients.map((patient) => {
    const values = dashboard_data.months.map(m => {
      const recs = dashboard_data.records.filter(r => r.month === m && r.patient === patient);
      if (!recs.length) return null;
      const avg = recs.reduce((s,r) => s + Number(r.score || 0), 0) / recs.length;
      return Number(avg.toFixed(2));
    });

    return {label: patient, data: values, borderWidth:1};
  });

  if (trendChart) trendChart.destroy();
  trendChart = new Chart(document.getElementById("trendChart"), {
    type:"bar",
    data:{labels, datasets},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:"top"}}, scales:{y:{beginAtZero:true, max:5}}}
  });
}

function renderDetailedKpis(month) {
  const rows = getDetailedKpis(month);
  const target = document.getElementById("detailedKpisTable");
  target.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>KPI</th>
          <th>Valor</th>
          <th>Meta</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td>${r.kpi}</td>
            <td>${pct(r.valor)}</td>
            <td>${pct(r.meta)}</td>
            <td class="status ${statusClass(r.valor, r.meta)}">${statusText(r.valor, r.meta)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderInsights(month) {
  const categories = getCategoryIndicators(month);
  const detailed = getDetailedKpis(month);
  const k = getMonthKpi(month);

  const belowTarget = categories.filter(x => x.positivo < x.meta)
    .sort((a,b) => (a.positivo - a.meta) - (b.positivo - b.meta));

  const worstPatients = [...patientAverages(month)]
    .sort((a,b) => a.avgScore - b.avgScore)
    .slice(0,3);

  const note5 = detailed.find(x => x.kpi.includes("Taxa de Notas 5"));
  const insights = [];

  if (belowTarget.length) {
    const first = belowTarget[0];
    insights.push(`${first.indicador} está em ${pct(first.positivo)} para meta de ${pct(first.meta)}. Recomenda-se reforçar conferência operacional e tratar a principal causa da não conformidade.`);
  }

  insights.push(`CSAT médio em ${Number(k.media_nota).toFixed(2)}/5 (${pct(k.csat_45)}). A taxa de notas 4–5 do mês está em ${pct(k.taxa_notas_45)}.`);

  if (note5) {
    insights.push(`A taxa de notas 5 está em ${pct(note5.valor)}. Vale trabalhar experiência final da entrega, comunicação com paciente e previsibilidade da operação.`);
  }

  if (worstPatients.length) {
    insights.push(`Pacientes mais críticos do mês: ${worstPatients.map(x => `${x.patient} (${x.avgScore.toFixed(2)})`).join(", ")}. Sugerido acompanhar ocorrência associada e repetir contato de qualidade.`);
  }

  insights.push(`Privacidade: este preview mostra estrutura gerencial e lógica de análise. No painel oficial, os dados serão atualizados automaticamente via pipeline do sistema Trans Tour.`);

  const target = document.getElementById("insightsList");
  target.innerHTML = insights.map(text => `<li>${text}</li>`).join("");
}

function populateMonths() {
  monthSelect.innerHTML = "";
  dashboard_data.months.forEach(month => {
    const opt = document.createElement("option");
    opt.value = month;
    opt.textContent = month;
    monthSelect.appendChild(opt);
  });
  monthSelect.value = dashboard_data.months[dashboard_data.months.length - 1];
}

function updateDashboard(month) {
  renderHeaderInfo(month);
  renderKpis(month);
  renderCategoryTable(month);
  renderDistributionChart(month);
  renderCriteriaChart(month);
  renderLowestPatients(month);
  renderTopBottomTable(month);
  renderTrendChart(month);
  renderDetailedKpis(month);
  renderInsights(month);
}

document.getElementById("btnMobile").addEventListener("click", () => {
  bodyEl.classList.add("mobile-preview");
  document.getElementById("btnMobile").classList.add("active");
  document.getElementById("btnNormal").classList.remove("active");
});

document.getElementById("btnNormal").addEventListener("click", () => {
  bodyEl.classList.remove("mobile-preview");
  document.getElementById("btnNormal").classList.add("active");
  document.getElementById("btnMobile").classList.remove("active");
});

document.getElementById("btnFullscreen").addEventListener("click", async () => {
  if (!document.fullscreenElement) {
    await document.documentElement.requestFullscreen?.();
  } else {
    await document.exitFullscreen?.();
  }
});

monthSelect.addEventListener("change", e => updateDashboard(e.target.value));

populateMonths();
updateDashboard(monthSelect.value);
</script>
</body>
</html>
"""


def render_html(dashboard_data: Dict) -> str:
    html = HTML_TEMPLATE
    html = html.replace("__PAGE_TITLE__", PAGE_TITLE)
    html = html.replace("__LOGO_TRANSTOUR__", LOGO_TRANSTOUR)
    html = html.replace("__LOGO_CLIENTE__", LOGO_CLIENTE)
    html = html.replace(
        "__DASHBOARD_JSON__",
        json.dumps(dashboard_data, ensure_ascii=False),
    )
    return html


def main() -> None:
    df = load_data()
    dashboard_data = build_dashboard_data(df)

    html = render_html(dashboard_data)
    Path(OUTPUT_HTML).write_text(html, encoding="utf-8")

    print(f"Dashboard gerado com sucesso: {OUTPUT_HTML}")
    print(f"Meses processados: {dashboard_data['months']}")
    print(f"Total de registros válidos: {len(dashboard_data['records'])}")


if __name__ == "__main__":
    main()
