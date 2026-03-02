import pandas as pd
import numpy as np
import datetime
import base64
import mimetypes
from pathlib import Path
import json

# 1) Link CSV (Sheets publicado)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ4A5MDb6JivQ54j3B3YrWPIpnidj49zOdeyLsqE1HKy8f35--M3ja_ZG_KntrKKeYOAIWNyS-3QOp8/pub?gid=0&single=true&output=csv"

# 2) Período: mês atual (ajuste se quiser “últimos 30 dias”)
USE_CURRENT_MONTH = False

# 3) Logos (opcional)
# Se você não quiser mexer com arquivo de logo, deixe vazio e o dashboard não mostra as imagens.
LOGO_TRANSTOUR_PATH = "logo-transtour.png"
LOGO_CLIENTE_PATH = "logo-cliente.png"

FOOTER_TEXT = "Relatório Gerado pelo Sistema Trans Tour Enviar e Receber | Confidencial"


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
    if "Submitted at" not in df.columns:
        raise ValueError("Coluna 'Submitted at' não encontrada no CSV. Verifique a aba publicada no Sheets.")

    df["Submitted at"] = pd.to_datetime(df["Submitted at"], errors="coerce")

    # Colunas esperadas (baseadas no seu formulário)
    col_itens = "Todos os itens foram entregues corretamente?"
    col_uniforme = "Entregador apresentou-se com crachá e uniforme?"
    col_produtos = "Produtos em bom estado e dentro da validade?"
    col_atendimento = "Atendimento cordial e respeitoso?"
    col_horario = "Entrega ocorreu no horário combinado?"
    col_nota = "GRAU DE SATISFAÇÃO (1 A 5)"

    qual_cols = [col_itens, col_uniforme, col_produtos, col_atendimento, col_horario]

    # Garante que as colunas existam
    missing = [c for c in qual_cols + [col_nota] if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas ausentes no CSV: {missing}")

    # Regra do cliente: "Recusou responder" => Sim e nota 5
    df_raw = df.copy()

    for c in qual_cols:
        df.loc[df[c].apply(is_refusal), c] = "Sim"

    refusal_row_mask = np.zeros(len(df_raw), dtype=bool)
    for c in qual_cols:
        refusal_row_mask |= df_raw[c].apply(is_refusal).to_numpy()

    df[col_nota] = pd.to_numeric(df[col_nota], errors="coerce")
    df.loc[refusal_row_mask | df[col_nota].isna(), col_nota] = 5
    df["month"] = df["Submitted at"].dt.to_period("M").astype(str)

    dataset_cols = [
    "Submitted at",
    "month",
    col_itens,
    col_uniforme,
    col_produtos,
    col_atendimento,
    col_horario,
    col_nota
]

df_dash = df[dataset_cols].copy()

for c in [col_itens, col_uniforme, col_produtos, col_atendimento, col_horario]:
    df_dash[c] = df_dash[c].astype(str).str.strip()

df_dash[col_nota] = pd.to_numeric(df_dash[col_nota], errors="coerce").fillna(5).astype(int)

# === Parte 2A: gerar JSON para o seletor de mês ===
df_dash["Submitted at"] = pd.to_datetime(df_dash["Submitted at"], errors="coerce")
df_dash = df_dash.dropna(subset=["Submitted at"])

# Data como string (pra não quebrar no JS)
df_dash["Submitted at"] = df_dash["Submitted at"].dt.strftime("%Y-%m-%d %H:%M:%S")

dashboard_data = df_dash.to_dict(orient="records")
    # Recorte do mês atual (opcional)
    if USE_CURRENT_MONTH:
        now = datetime.datetime.now()
        start = datetime.datetime(now.year, now.month, 1)
        end = (start + datetime.timedelta(days=32)).replace(day=1)
        df = df[(df["Submitted at"] >= start) & (df["Submitted at"] < end)].copy()

    df = df.dropna(subset=["Submitted at"])

    # Se não houver dados, gera um HTML informativo
    if len(df) == 0:
        Path("index.html").write_text(
            "<h1>Sem dados disponíveis para o período.</h1>",
            encoding="utf-8"
        )
        return

    n = len(df)

    def norm(s):
        return s.astype(str).str.strip().str.lower()

    def sim_nao(series):
        s = norm(series)
        sim = (s == "sim").sum()
        nao = ((s == "não") | (s == "nao")).sum()
        return int(sim), int(nao)

    cats = [
        ("Itens Entregues Corretamente", col_itens, 0.90),
        ("Crachá e Uniforme", col_uniforme, 0.95),
        ("Produtos em Bom Estado", col_produtos, 0.95),
        ("Atendimento Cordial", col_atendimento, 0.95),
        ("Pontualidade", col_horario, 0.90),
    ]

    rows = []
    for name, col, meta in cats:
        sim, nao = sim_nao(df[col])
        pct_pos = sim / n if n else 0
        rows.append((name, sim, nao, pct_pos, meta))

    nota = df[col_nota].astype(float)
    mean = float(nota.mean())
    csat = mean / 5 * 100
    rate5 = float((nota == 5).mean() * 100)
    rate45 = float((nota >= 4).mean() * 100)

    period_start = df["Submitted at"].min()
    period_end = df["Submitted at"].max()
    period_label = f"{period_start:%d/%m/%Y} a {period_end:%d/%m/%Y}"
    period_title = f"{period_start:%B/%Y}".capitalize()

    itens_pct = rows[0][3] * 100
    nao_itens = rows[0][2]

    def status_icon(pct):
        if pct >= 95:
            return "✓ Excelente"
        if pct >= 80:
            return "✓ Bom"
        return "⚠️ Atenção"

    def status_text(val, meta):
        return "✓ Atingida" if val >= meta else "✗ Não Atingida"

    cat_rows_html = ""
    for name, sim, nao, pct_pos, meta in rows:
        cat_rows_html += (
            f"<tr><td>{name}</td><td>{sim}</td><td>{nao}</td>"
            f"<td>{pct_pos*100:.1f}%</td><td>{meta*100:.0f}%</td><td>{status_icon(pct_pos*100)}</td></tr>"
        )

    kpi_det = [
        ("KPI-01: Conformidade de Itens", itens_pct, 90),
        ("KPI-02: Padronização Visual (Uniforme/Crachá)", rows[1][3] * 100, 95),
        ("KPI-03: Qualidade dos Produtos", rows[2][3] * 100, 95),
        ("KPI-04: Qualidade no Atendimento", rows[3][3] * 100, 95),
        ("KPI-05: Pontualidade", rows[4][3] * 100, 90),
        ("KPI-06: CSAT (média/5)", csat, 90),
        ("KPI-07: Taxa de Notas 5", rate5, 80),
        ("KPI-08: Taxa de Notas 4-5", rate45, 90),
    ]

    kpi_rows_html = ""
    for name, val, meta in kpi_det:
        kpi_rows_html += f"<tr><td>{name}</td><td>{val:.1f}%</td><td>{meta:.0f}%</td><td>{status_text(val, meta)}</td></tr>"

    labels = [r[0] for r in rows]
    values = [round(r[3] * 100, 1) for r in rows]
    metas = [round(r[4] * 100, 1) for r in rows]
    sat_counts = [int((nota == i).sum()) for i in range(1, 6)]

    # Insights (sem mencionar falta de mês, como você pediu antes)
    ins = [
        f"Acuracidade de itens em {itens_pct:.1f}% (meta 90%). Recomendamos reforçar conferência (checklist duplo) e rastrear causas das {nao_itens} não conformidades por falta de itens internos.",
        f"CSAT médio {mean:.2f}/5 ({csat:.1f}%). Taxa de nota 5: {rate5:.1f}%.",
        "Conforme alinhado com o cliente, respostas 'Recusou responder' foram consideradas como conformidade (Sim) e nota máxima (5) para fins de consolidação."
    ]

    generated = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    logo_transtour_uri = data_uri(LOGO_TRANSTOUR_PATH)
    logo_cliente_uri = data_uri(LOGO_CLIENTE_PATH)

    logos_html = ""
    if logo_transtour_uri:
        logos_html += f'<img class="logo" src="{logo_transtour_uri}" alt="Transtour"/>'
    if logo_cliente_uri:
        logos_html += f'<img class="logo" src="{logo_cliente_uri}" alt="Cliente"/>'

    html = f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Relatório Gerencial de KPIs - Hapvida | {period_title}</title>
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
  .insights ul{{margin:8px 0 0 18px}}
  .footer{{text-align:center;color:#9ca3af;font-size:12px;margin-top:12px}}

  .view-controls{{position:fixed;right:18px;bottom:18px;display:flex;gap:10px;z-index:9999;flex-wrap:wrap}}
  .view-btn{{border:none;cursor:pointer;padding:12px 14px;border-radius:999px;font-weight:700;font-size:14px;box-shadow:0 10px 24px rgba(0,0,0,.18);background:#4f46e5;color:#fff;display:flex;align-items:center;gap:10px}}
  .view-btn.secondary{{background:#0b1f3a}}
  .view-btn.ghost{{background:#111827}}
  .view-btn:active{{transform:translateY(1px)}}

  body.mode-mobile{{background:#6b7cff}}
  body.mode-mobile .wrap{{max-width:430px}}
  body.mode-mobile .topbar{{border-radius:26px 26px 14px 14px}}
  body.mode-mobile .phone-shell{{position:relative;margin:14px auto 60px;padding:18px 14px 18px;background:rgba(255,255,255,.08);border-radius:36px;box-shadow:0 24px 60px rgba(0,0,0,.25);max-width:460px}}
  body.mode-mobile .phone-shell:before{{content:"";position:absolute;top:10px;left:50%;transform:translateX(-50%);width:120px;height:22px;background:rgba(0,0,0,.28);border-radius:999px}}
  body.mode-mobile .phone-shell:after{{content:"";position:absolute;top:10px;left:calc(50% + 58px);width:10px;height:10px;background:rgba(255,255,255,.55);border-radius:999px}}

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
          <div class="sub">Satisfação de Entregas • Período: <b>{period_title}</b> • Base: <b>{n}</b> respostas</div>
          <div class="sub" style="font-size:12px">Recorte detectado no arquivo: <b>{period_label}</b> • Atualizado: <b>{generated}</b></div>
        </div>
        <div style="min-width:160px;text-align:right;opacity:.95">
          <div style="font-size:12px">Notas válidas</div>
          <div style="font-weight:800">{len(nota)}</div>
        </div>
      </div>
    </div>

    <div class="grid">
      <div class="card"><div class="kpiV">{n}</div><div class="kpiL">Total de Respostas</div><span class="badge ok">Amostra do período</span></div>
      <div class="card"><div class="kpiV">{rate5:.1f}%</div><div class="kpiL">Taxa de Satisfação Geral (Nota 5)</div><span class="badge ok">Meta sugerida: ≥ 80%</span></div>
      <div class="card"><div class="kpiV">{itens_pct:.1f}%</div><div class="kpiL">Itens Entregues Corretamente</div><span class="badge {"warn" if itens_pct<90 else "ok"}">Meta: 90%</span></div>
      <div class="card"><div class="kpiV">{itens_pct:.0f}</div><div class="kpiL">NPS Estimado (proxy)</div><span class="badge ok">Proxy = acuracidade</span></div>
    </div>

    <div class="section">
      <h3 class="title">📈 Indicadores por Categoria</h3>
      <table>
        <thead><tr><th>Indicador</th><th>Sim</th><th>Não</th><th>% Positivo</th><th>Meta</th><th>Status</th></tr></thead>
        <tbody>{cat_rows_html}</tbody>
      </table>
      <div class="muted">* % Positivo calculado sobre o total de respostas do período.</div>
    </div>

    <div class="section charts">
      <div class="card">
        <h3 class="title">⭐ Distribuição de Satisfação (Notas 1–5)</h3>
        <canvas id="sat"></canvas>
        <div class="muted">Média: <b>{mean:.2f}/5</b> • CSAT: <b>{csat:.1f}%</b></div>
      </div>
      <div class="card">
        <h3 class="title">📊 Conformidade por Critério (vs Meta)</h3>
        <canvas id="bars"></canvas>
      </div>
    </div>

    <div class="section">
      <h3 class="title">📋 KPIs Detalhados</h3>
      <table>
        <thead><tr><th>KPI</th><th>Valor</th><th>Meta</th><th>Status</th></tr></thead>
        <tbody>{kpi_rows_html}</tbody>
      </table>
    </div>

    <div class="section card insights">
      <h3 class="title">💡 Insights e Recomendações</h3>
      <ul>
        {''.join(f"<li>{x}</li>" for x in ins)}
        <li><b>Privacidade:</b> publique somente indicadores agregados.</li>
      </ul>
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
  const satCtx = document.getElementById('sat').getContext('2d');
  new Chart(satCtx, {{
    type: 'bar',
    data: {{
      labels: ['1','2','3','4','5'],
      datasets: [{{ label: 'Qtd de respostas', data: {json.dumps(sat_counts)} }}]
    }},
    options: {{
      plugins: {{ legend: {{ display:false }} }},
      scales: {{ y: {{ beginAtZero:true, ticks: {{ stepSize:1 }} }} }}
    }}
  }});

  const barCtx = document.getElementById('bars').getContext('2d');
  new Chart(barCtx, {{
    type: 'bar',
    data: {{
      labels: {json.dumps(labels)},
      datasets: [
        {{ label: 'Resultado (%)', data: {json.dumps(values)} }},
        {{ label: 'Meta (%)', data: {json.dumps(metas)} }}
      ]
    }},
    options: {{
      plugins: {{ legend: {{ position:'bottom' }} }},
      scales: {{ y: {{ min:0, max:100 }} }}
    }}
  }});

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

    btnMobile?.addEventListener('click', () => setMode('mobile'));
    btnDesktop?.addEventListener('click', () => setMode('desktop'));

    btnFullscreen?.addEventListener('click', async () => {{
      try{{
        if(!document.fullscreenElement){{
          await document.documentElement.requestFullscreen();
          btnFullscreen.textContent = '⤢ Sair da tela cheia';
        }} else {{
          await document.exitFullscreen();
          btnFullscreen.textContent = '⛶ Tela cheia (F11)';
        }}
      }}catch(e){{
        alert('Seu navegador bloqueou a tela cheia via botão. Use F11 para alternar.');
      }}
    }});

    document.addEventListener('fullscreenchange', () => {{
      if(!document.fullscreenElement) btnFullscreen.textContent = '⛶ Tela cheia (F11)';
    }});
  }})();
</script>

</body>
</html>
"""
    Path("index.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
