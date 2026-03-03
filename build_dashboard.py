import pandas as pd
import numpy as np
import datetime
import base64
import mimetypes
from pathlib import Path
import json

# Link CSV do Google Sheets publicado
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ4A5MDb6JivQ54j3B3YrWPIpnidj49zOdeyLsqE1HKy8f35--M3ja_ZG_KntrKKeYOAIWNyS-3QOp8/pub?gid=0&single=true&output=csv"

# Logos no repo (opcional)
LOGO_TRANSTOUR_PATH = "logo-transtour.png"
LOGO_CLIENTE_PATH = "logo-cliente.png"

FOOTER_TEXT = "Relatório Gerado pelo Sistema Trans Tour Enviar e Receber | Confidencial"

COL_SUBMITTED = "Submitted at"
COL_ITENS = "Todos os itens foram entregues corretamente?"
COL_UNIFORME = "Entregador apresentou-se com crachá e uniforme?"
COL_PRODUTOS = "Produtos em bom estado e dentro da validade?"
COL_ATENDIMENTO = "Atendimento cordial e respeitoso?"
COL_HORARIO = "Entrega ocorreu no horário combinado?"
COL_NOTA = "GRAU DE SATISFAÇÃO (1 A 5)"

QUAL_COLS = [COL_ITENS, COL_UNIFORME, COL_PRODUTOS, COL_ATENDIMENTO, COL_HORARIO]


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


def main():
    df = pd.read_csv(SHEET_CSV_URL)

    # valida colunas
    required = [COL_SUBMITTED, COL_NOTA] + QUAL_COLS
    missing = [c for c in required if c not in df.columns]
    if missing:
        Path("index.html").write_text(
            f"<h1>Erro: colunas ausentes no CSV</h1><pre>{missing}</pre>",
            encoding="utf-8"
        )
        return

    # datas
    df[COL_SUBMITTED] = pd.to_datetime(df[COL_SUBMITTED], errors="coerce")
    df = df.dropna(subset=[COL_SUBMITTED]).copy()

    # regra do cliente: recusou => Sim e nota 5
    df_raw = df.copy()
    for c in QUAL_COLS:
        df.loc[df[c].apply(is_refusal), c] = "Sim"

    refusal_row_mask = np.zeros(len(df_raw), dtype=bool)
    for c in QUAL_COLS:
        refusal_row_mask |= df_raw[c].apply(is_refusal).to_numpy()

    df[COL_NOTA] = pd.to_numeric(df[COL_NOTA], errors="coerce")
    df.loc[refusal_row_mask | df[COL_NOTA].isna(), COL_NOTA] = 5

    # month YYYY-MM
    df["month"] = df[COL_SUBMITTED].dt.to_period("M").astype(str)

    # dataset enxuto pro dashboard
    df_dash = df[[COL_SUBMITTED, "month"] + QUAL_COLS + [COL_NOTA]].copy()
    for c in QUAL_COLS:
        df_dash[c] = df_dash[c].astype(str).str.strip()
    df_dash[COL_NOTA] = pd.to_numeric(df_dash[COL_NOTA], errors="coerce").fillna(5).astype(int)

    # data como string pro JS
    df_dash[COL_SUBMITTED] = df_dash[COL_SUBMITTED].dt.strftime("%Y-%m-%d %H:%M:%S")

    dashboard_data = df_dash.to_dict(orient="records")

    # logos
    logo_transtour_uri = data_uri(LOGO_TRANSTOUR_PATH)
    logo_cliente_uri = data_uri(LOGO_CLIENTE_PATH)

    logos_html = ""
    if logo_transtour_uri:
        logos_html += f'<img class="logo" src="{logo_transtour_uri}" alt="Transtour"/>'
    if logo_cliente_uri:
        logos_html += f'<img class="logo" src="{logo_cliente_uri}" alt="Cliente"/>'

    generated = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

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
  .badge{{display:inline-block;margin-top:8px;padding:4px 10px;border-radius:999px;font-size:12px}}
  .ok{{background:#eafaf1;color:#166534}}
  .warn{{background:#fdecea;color:#991b1b}}
  .section{{margin-top:18px}}
  .title{{font-weight:800;margin:0 0 10px 0}}
  table{{width:100%;border-collapse:collapse;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 3px 10px rgba(0,0,0,.08)}}
  th,td{{padding:10px 12px;border-bottom:1px solid #eef2f7;font-size:13px;vertical-align:top}}
  th{{text-align:left;background:#f9fafb;color:#374151;font-size:12px;text-transform:uppercase;letter-spacing:.5px}}
  tr:last-child td{{border-bottom:none}}
  .charts{{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:14px}}
  .muted{{color:#6b7280;font-size:12px;margin-top:6px}}
  .footer{{text-align:center;color:#9ca3af;font-size:12px;margin-top:12px}}
  .view-controls{{position:fixed;right:18px;bottom:18px;display:flex;gap:10px;z-index:9999;flex-wrap:wrap}}
  .view-btn{{border:none;cursor:pointer;padding:12px 14px;border-radius:999px;font-weight:700;font-size:14px;box-shadow:0 10px 24px rgba(0,0,0,.18);background:#4f46e5;color:#fff;display:flex;align-items:center;gap:10px}}
  .view-btn.secondary{{background:#0b1f3a}}
  .view-btn.ghost{{background:#111827}}
  body.mode-mobile{{background:#6b7cff}}
  body.mode-mobile .wrap{{max-width:430px}}
  body.mode-mobile .topbar{{border-radius:26px 26px 14px 14px}}
  body.mode-mobile .phone-shell{{position:relative;margin:14px auto 60px;padding:18px 14px 18px;background:rgba(255,255,255,.08);border-radius:36px;box-shadow:0 24px 60px rgba(0,0,0,.25);max-width:460px}}
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
      </div>
      <div class="card">
        <h3 class="title">📊 Conformidade por Critério (vs Meta)</h3>
        <canvas id="bars"></canvas>
      </div>
    </div>

    <div class="footer">{FOOTER_TEXT}</div>
  </div>
</div>

<div class="view-controls">
  <button class="view-btn secondary" id="btnMobile" type="button">📱 Modo celular</button>
  <button class="view-btn" id="btnDesktop" type="button">🖥️ Modo tela cheia</button>
  <button class="view-btn ghost" id="btnFullscreen" type="button">⛶ Tela cheia (F11)</button>
</div>

<script>
  const DASHBOARD_DATA = {json.dumps(dashboard_data)};

  function pct(n, d){{ return d === 0 ? 0 : (n/d)*100; }}

  function getMonths(data){{
    const set = new Set(data.map(r => r.month));
    return Array.from(set).sort();
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

  function calc(arr){{
    const total = arr.length;
    const itens = countSimNao(arr, "{COL_ITENS}");
    const uniforme = countSimNao(arr, "{COL_UNIFORME}");
    const produtos = countSimNao(arr, "{COL_PRODUTOS}");
    const atendimento = countSimNao(arr, "{COL_ATENDIMENTO}");
    const horario = countSimNao(arr, "{COL_HORARIO}");

    const notas = arr.map(r => Number(r["{COL_NOTA}"])).filter(x => !Number.isNaN(x));
    const rate5 = notas.length ? (notas.filter(x=>x===5).length/notas.length)*100 : 0;

    return {{ total, itens, uniforme, produtos, atendimento, horario, notasValidas: notas.length, rate5 }};
  }}

  function status(pos, meta){{
    if(pos >= meta) return "✓ Excelente";
    if(pos >= 80) return "⚠️ Atenção";
    return "⚠️ Atenção";
  }}

  let satChart = null;
  let barChart = null;

  function render(month){{
    const arr = filterByMonth(DASHBOARD_DATA, month);
    const k = calc(arr);

    document.getElementById("monthHint").textContent = `Base do mês: ${{k.total}} respostas`;
    document.getElementById("kpiTotal").textContent = k.total;
    document.getElementById("kpiSat5").textContent = `${{k.rate5.toFixed(1)}}%`;
    document.getElementById("kpiItens").textContent = `${{pct(k.itens.sim, k.total).toFixed(1)}}%`;
    document.getElementById("kpiNps").textContent = `${{pct(k.itens.sim, k.total).toFixed(0)}}`;
    document.getElementById("kpiNotasValidas").textContent = k.notasValidas;

    const tbody = document.getElementById("tbodyIndicadores");
    const rows = [
      ["Itens Entregues Corretamente", k.itens, 90],
      ["Crachá e Uniforme", k.uniforme, 95],
      ["Produtos em Bom Estado", k.produtos, 95],
      ["Atendimento Cordial", k.atendimento, 95],
      ["Pontualidade", k.horario, 90],
    ];

    tbody.innerHTML = rows.map(([name,obj,meta]) => {{
      const pos = pct(obj.sim, k.total);
      return `<tr>
        <td>${{name}}</td>
        <td>${{obj.sim}}</td>
        <td>${{obj.nao}}</td>
        <td>${{pos.toFixed(1)}}%</td>
        <td>${{meta}}%</td>
        <td>${{status(pos, meta)}}</td>
      </tr>`;
    }}).join("");

    // charts
    const counts = [1,2,3,4,5].map(n => arr.filter(r => Number(r["{COL_NOTA}"]) === n).length);

    const satCtx = document.getElementById("sat").getContext("2d");
    if(satChart) satChart.destroy();
    satChart = new Chart(satCtx, {{
      type: "bar",
      data: {{
        labels: ["1","2","3","4","5"],
        datasets: [{{ label: "Qtd de respostas", data: counts }}]
      }},
      options: {{
        plugins: {{ legend: {{ display:false }} }},
        scales: {{ y: {{ beginAtZero:true }} }}
      }}
    }});

    const barCtx = document.getElementById("bars").getContext("2d");
    if(barChart) barChart.destroy();

    const results = [
      pct(k.itens.sim, k.total),
      pct(k.uniforme.sim, k.total),
      pct(k.produtos.sim, k.total),
      pct(k.atendimento.sim, k.total),
      pct(k.horario.sim, k.total),
    ].map(x => Number(x.toFixed(1)));

    barChart = new Chart(barCtx, {{
      type: "bar",
      data: {{
        labels: ["Itens", "Uniforme", "Produtos", "Atendimento", "Pontualidade"],
        datasets: [
          {{ label: "Resultado (%)", data: results }},
          {{ label: "Meta (%)", data: [90,95,95,95,90] }}
        ]
      }},
      options: {{
        plugins: {{ legend: {{ position:"bottom" }} }},
        scales: {{ y: {{ min:0, max:100 }} }}
      }}
    }});
  }}

  // init selector
  const months = getMonths(DASHBOARD_DATA);
  const select = document.getElementById("monthSelect");
  select.innerHTML = months.map(m => `<option value="${{m}}">${{m}}</option>`).join("");
  const last = months[months.length-1];
  select.value = last;
  render(last);

  select.addEventListener("change", () => render(select.value));

  // mobile/fullscreen toggles
  (function(){{
    const body = document.body;
    const btnMobile = document.getElementById('btnMobile');
    const btnDesktop = document.getElementById('btnDesktop');
    const btnFullscreen = document.getElementById('btnFullscreen');

    function setMode(mode){{
      if(mode === 'mobile') body.classList.add('mode-mobile');
      else body.classList.remove('mode-mobile');
      localStorage.setItem('dashboard_view_mode', mode);
    }}

    const saved = localStorage.getItem('dashboard_view_mode');
    if(saved) setMode(saved);

    btnMobile.addEventListener('click', () => setMode('mobile'));
    btnDesktop.addEventListener('click', () => setMode('desktop'));

    btnFullscreen.addEventListener('click', async () => {{
      try {{
        if(!document.fullscreenElement) {{
          await document.documentElement.requestFullscreen();
        }} else {{
          await document.exitFullscreen();
        }}
      }} catch(e) {{
        alert('Use F11 para alternar tela cheia.');
      }}
    }});
  }})();
</script>

</body>
</html>
"""

    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
