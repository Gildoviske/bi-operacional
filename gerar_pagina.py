# -*- coding: utf-8 -*-
"""Gera a página HTML do BI Operacional a partir das planilhas de controle.

Uso: python gerar_pagina.py
Lê os 3 arquivos em "I:\\Meu Drive\\CONTROLE COMPRAS" e grava index.html
nesta mesma pasta (bi-web), pronta para publicar no GitHub Pages.
"""
import json
import openpyxl
from datetime import datetime
from pathlib import Path

BASE = Path(r"I:\Meu Drive\CONTROLE COMPRAS")
OUT = Path(__file__).resolve().parent / "index.html"


def load(arquivo, aba):
    wb = openpyxl.load_workbook(BASE / arquivo, data_only=True, read_only=True)
    return list(wb[aba].iter_rows(values_only=True))


def pct(v):
    return f"{v * 100:.1f}%".replace(".", ",")


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_str(v):
    if v is None:
        return ""
    return v.strftime("%d/%m/%Y")


# ---------------------------------------------------------------- PEDIDOS
rows = load("CONTROLE DE PEDIDOS.xlsx", "RESUMO")


def find(label):
    for r in rows:
        if r[1] == label:
            return r
    return None


data_atualizacao = next(
    r[1].replace("Atualizado em ", "").strip() for r in rows if r[1] and str(r[1]).startswith("Atualizado em")
)
pedidos = {
    "total": find("Total de pedidos")[2],
    "entregue": (find("Entregue")[2], find("Entregue")[3]),
    "pendente": (find("Pendente de entrada")[2], find("Pendente de entrada")[3]),
    "atrasado": (find("Atrasado")[2], find("Atrasado")[3]),
    "entrada_ok": (find("Entrada no sistema realizada")[2], find("Entrada no sistema realizada")[3]),
    "entrada_pendente": (find("Entrada pendente no sistema")[2], find("Entrada pendente no sistema")[3]),
}

# pedidos por filial (Filial, Total, Entregue, Pendente, Atrasado)
pedidos_filiais = []
capturando = False
for r in rows:
    if r[1] == "Filial" and r[2] == "Total":
        capturando = True
        continue
    if capturando:
        if r[1] is None:
            continue
        if r[1] == "TOTAL":
            break
        nome = r[1].replace("ROCHA TELECOM - ", "")
        pedidos_filiais.append({"filial": nome, "total": r[2], "entregue": r[3], "pendente": r[4], "atrasado": r[5]})
pedidos_filiais.sort(key=lambda f: (f["pendente"] + f["atrasado"]), reverse=True)

# --------------------------------------------------------- NOTAS FISCAIS
rows = load("CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx", "NOTAS FISCAIS PENDENTES")
data_rows = rows[1:]

notas_resumo = {}
for r in data_rows:
    label = r[13]
    if label:
        notas_resumo[label] = r[14]

notas_atrasadas = [r for r in data_rows if r[12] == "ATRASADO"]
notas_atrasadas.sort(key=lambda r: (r[11] or 0), reverse=True)

# ------------------------------------------------------- TRANSFERÊNCIAS
rows = load("TRANSFERÊNCIAS PENDENTES.xlsx", "TRANSFERENCIAS PENDENTES")
data_rows = rows[1:]

transf_resumo = {}
for r in data_rows:
    label = r[14]
    if label and label != "RESUMO GERAL":
        transf_resumo[label] = r[15]

transf_criticas = [r for r in data_rows if r[11] == "CRÍTICO"]
transf_criticas.sort(key=lambda r: (r[2] or 0), reverse=True)

gerado_em = datetime.now().strftime("%d/%m/%Y %H:%M")

chart_data = {
    "pedidosStatus": {
        "labels": ["Entregue", "Pendente de entrada", "Atrasado"],
        "values": [pedidos["entregue"][0], pedidos["pendente"][0], pedidos["atrasado"][0]],
    },
    "pedidosEntrada": {
        "labels": ["Entrada realizada", "Entrada pendente"],
        "values": [pedidos["entrada_ok"][0], pedidos["entrada_pendente"][0]],
    },
    "pedidosFilial": {
        "labels": [f["filial"] for f in pedidos_filiais],
        "entregue": [f["entregue"] for f in pedidos_filiais],
        "pendente": [f["pendente"] for f in pedidos_filiais],
        "atrasado": [f["atrasado"] for f in pedidos_filiais],
    },
    "notasStatus": {
        "labels": ["Atrasadas", "Aguardando (no prazo)"],
        "values": [notas_resumo.get("Atrasadas", 0), notas_resumo.get("Aguardando (no prazo)", 0)],
    },
    "transfStatus": {
        "labels": ["No prazo", "Atenção", "Crítico"],
        "values": [transf_resumo.get("No prazo", 0), transf_resumo.get("ATENÇÃO", 0), transf_resumo.get("CRÍTICO", 0)],
    },
    "transfPrazo": {
        "labels": ["0-3 dias", "4-7 dias", "8-15 dias", "16-30 dias", "30+ dias"],
        "values": [
            transf_resumo.get("0-3 dias", 0), transf_resumo.get("4-7 dias", 0),
            transf_resumo.get("8-15 dias", 0), transf_resumo.get("16-30 dias", 0),
            transf_resumo.get("30+ dias", 0),
        ],
    },
}
chart_data_json = json.dumps(chart_data, ensure_ascii=False)

# ----------------------------------------------------------------- HTML
def linhas_notas():
    out = []
    for r in notas_atrasadas:
        out.append(
            "<tr><td>{fil}</td><td>{nf}</td><td>{desc}</td><td class='num'>{qtd}</td>"
            "<td>{dia}</td><td class='num'>{atr}</td></tr>".format(
                fil=r[0], nf=r[1], desc=r[5], qtd=r[6], dia=data_str(r[10]), atr=round(r[11] or 0)
            )
        )
    return "\n".join(out)


def linhas_transf():
    out = []
    for r in transf_criticas:
        out.append(
            "<tr><td>{orig}</td><td class='num'>{dias}</td><td>{dest}</td><td>{user}</td>"
            "<td>{nf}</td><td>{prod}</td><td>{desc}</td><td class='num'>{qtd}</td></tr>".format(
                orig=r[0], dias=round(r[2] or 0), dest=r[3], user=r[4], nf=r[5],
                prod=r[7], desc=r[8], qtd=r[9]
            )
        )
    return "\n".join(out)


html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>BI Operacional — Rocha Telecom</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  :root {{
    --bg: #0f172a; --card: #1e293b; --card2: #16213a; --text: #e2e8f0; --muted: #94a3b8;
    --ok: #22c55e; --warn: #f59e0b; --bad: #ef4444; --accent: #38bdf8; --border: #334155;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); }}
  header {{ padding: 24px 32px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; flex-wrap: wrap; }}
  header h1 {{ margin: 0 0 4px; font-size: 22px; }}
  header p {{ margin: 0; color: var(--muted); font-size: 14px; }}
  .badge-atualizacao {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 10px 16px; text-align: right; white-space: nowrap; }}
  .badge-atualizacao .data {{ font-size: 15px; font-weight: 700; color: var(--accent); }}
  .badge-atualizacao .gerado {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}
  main {{ padding: 24px 32px 64px; max-width: 1200px; margin: 0 auto; }}
  section {{ margin-bottom: 40px; }}
  section > h2 {{ font-size: 18px; border-left: 4px solid var(--accent); padding-left: 10px; margin-bottom: 16px; }}
  .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 20px; }}
  .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }}
  .card .label {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }}
  .card .value {{ font-size: 26px; font-weight: 700; margin-top: 6px; }}
  .card .sub {{ font-size: 13px; color: var(--muted); margin-top: 2px; }}
  .card.ok .value {{ color: var(--ok); }}
  .card.warn .value {{ color: var(--warn); }}
  .card.bad .value {{ color: var(--bad); }}
  table {{ width: 100%; border-collapse: collapse; background: var(--card2); border-radius: 10px; overflow: hidden; font-size: 13px; }}
  th, td {{ padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: left; }}
  th {{ background: #0b1120; color: var(--muted); font-size: 11px; text-transform: uppercase; position: sticky; top: 0; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  tr:hover td {{ background: #22314f; }}
  .table-wrap {{ max-height: 480px; overflow: auto; border: 1px solid var(--border); border-radius: 10px; }}
  .charts {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; margin-bottom: 20px; }}
  .chart-box {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; position: relative; height: 320px; }}
  .chart-box h4 {{ margin: 0 0 12px; font-size: 13px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }}
  .chart-box .canvas-wrap {{ position: relative; height: calc(100% - 28px); }}
  .chart-box.wide {{ grid-column: 1 / -1; height: 720px; }}
  input.filtro {{ width: 100%; max-width: 320px; padding: 8px 10px; margin-bottom: 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); }}
  footer {{ text-align: center; color: var(--muted); font-size: 12px; padding: 24px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>📊 BI Operacional — Pedidos, Notas Fiscais e Transferências</h1>
    <p>Giovanni Brochini · Rocha Telecom</p>
  </div>
  <div class="badge-atualizacao">
    <div class="data">🕒 Atualizado em {data_atualizacao}</div>
    <div class="gerado">Página gerada em {gerado_em}</div>
  </div>
</header>
<main>

<section>
  <h2>📦 Pedidos GN — Status Geral</h2>
  <div class="cards">
    <div class="card"><div class="label">Total de pedidos</div><div class="value">{pedidos['total']}</div></div>
    <div class="card ok"><div class="label">Entregue</div><div class="value">{pedidos['entregue'][0]}</div><div class="sub">{pct(pedidos['entregue'][1])}</div></div>
    <div class="card warn"><div class="label">Pendente de entrada</div><div class="value">{pedidos['pendente'][0]}</div><div class="sub">{pct(pedidos['pendente'][1])}</div></div>
    <div class="card bad"><div class="label">Atrasado</div><div class="value">{pedidos['atrasado'][0]}</div><div class="sub">{pct(pedidos['atrasado'][1])}</div></div>
    <div class="card ok"><div class="label">Entrada no sistema realizada</div><div class="value">{pedidos['entrada_ok'][0]}</div><div class="sub">{pct(pedidos['entrada_ok'][1])}</div></div>
    <div class="card warn"><div class="label">Entrada pendente no sistema</div><div class="value">{pedidos['entrada_pendente'][0]}</div><div class="sub">{pct(pedidos['entrada_pendente'][1])}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status geral dos pedidos</h4><div class="canvas-wrap"><canvas id="chart-pedidos-status"></canvas></div></div>
    <div class="chart-box"><h4>Entrada no sistema</h4><div class="canvas-wrap"><canvas id="chart-pedidos-entrada"></canvas></div></div>
    <div class="chart-box wide"><h4>Pedidos por filial (entregue / pendente / atrasado)</h4><div class="canvas-wrap"><canvas id="chart-pedidos-filial"></canvas></div></div>
  </div>
</section>

<section>
  <h2>🧾 Notas Fiscais Pendentes de Entrada</h2>
  <div class="cards">
    <div class="card"><div class="label">Total de notas</div><div class="value">{notas_resumo.get('Total de notas (linhas)', 0)}</div></div>
    <div class="card warn"><div class="label">Pendentes</div><div class="value">{notas_resumo.get('Pendentes (NÃO)', 0)}</div></div>
    <div class="card bad"><div class="label">Atrasadas</div><div class="value">{notas_resumo.get('Atrasadas', 0)}</div></div>
    <div class="card ok"><div class="label">Aguardando (no prazo)</div><div class="value">{notas_resumo.get('Aguardando (no prazo)', 0)}</div></div>
    <div class="card"><div class="label">Valor total pendente</div><div class="value">{brl(notas_resumo.get('Valor total pendente', 0))}</div></div>
    <div class="card bad"><div class="label">Maior atraso</div><div class="value">{int(notas_resumo.get('Maior atraso (dias)', 0))} dias</div></div>
    <div class="card"><div class="label">Média de atraso</div><div class="value">{notas_resumo.get('Média de atraso (dias)', 0):.1f} dias</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Notas pendentes: atrasadas x no prazo</h4><div class="canvas-wrap"><canvas id="chart-notas-status"></canvas></div></div>
  </div>
  <h3>⚠️ Notas com criticidade elevada (atrasadas)</h3>
  <input class="filtro" data-target="tbl-notas" placeholder="Filtrar por filial, NF ou produto...">
  <div class="table-wrap">
  <table id="tbl-notas">
    <thead><tr><th>Filial</th><th>Nota Fiscal</th><th>Descrição</th><th>Qtde</th><th>Dia da Entrega</th><th>Dias em Atraso</th></tr></thead>
    <tbody>
    {linhas_notas()}
    </tbody>
  </table>
  </div>
</section>

<section>
  <h2>🔄 Transferências Pendentes</h2>
  <div class="cards">
    <div class="card"><div class="label">Qtde total (unidades)</div><div class="value">{transf_resumo.get('Qtde total (unidades)', 0)}</div></div>
    <div class="card ok"><div class="label">No prazo</div><div class="value">{transf_resumo.get('No prazo', 0)}</div></div>
    <div class="card bad"><div class="label">Atrasados</div><div class="value">{transf_resumo.get('Atrasados', 0)}</div></div>
    <div class="card warn"><div class="label">Atenção</div><div class="value">{transf_resumo.get('ATENÇÃO', 0)}</div></div>
    <div class="card bad"><div class="label">Crítico</div><div class="value">{transf_resumo.get('CRÍTICO', 0)}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Transferências por criticidade</h4><div class="canvas-wrap"><canvas id="chart-transf-status"></canvas></div></div>
    <div class="chart-box"><h4>Tempo fora do estoque</h4><div class="canvas-wrap"><canvas id="chart-transf-prazo"></canvas></div></div>
  </div>
  <h3>⚠️ Transferências com criticidade elevada</h3>
  <input class="filtro" data-target="tbl-transf" placeholder="Filtrar por filial, NF ou produto...">
  <div class="table-wrap">
  <table id="tbl-transf">
    <thead><tr><th>Filial Origem</th><th>Dias fora do estoque</th><th>Filial Destino</th><th>Usuário Solicitante</th><th>NF</th><th>Produto</th><th>Descrição</th><th>Qtde</th></tr></thead>
    <tbody>
    {linhas_transf()}
    </tbody>
  </table>
  </div>
</section>

</main>
<footer>Gerado automaticamente a partir das planilhas de controle · Rocha Telecom</footer>
<script>
document.querySelectorAll('input.filtro').forEach(function(inp) {{
  inp.addEventListener('input', function() {{
    var table = document.getElementById(inp.dataset.target);
    var q = inp.value.toLowerCase();
    table.querySelectorAll('tbody tr').forEach(function(tr) {{
      tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
    }});
  }});
}});

const CHART_DATA = {chart_data_json};
Chart.defaults.color = '#94a3b8';
Chart.defaults.borderColor = '#334155';
const COR_OK = '#22c55e', COR_WARN = '#f59e0b', COR_BAD = '#ef4444', COR_ACCENT = '#38bdf8';

new Chart(document.getElementById('chart-pedidos-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.pedidosStatus.labels,
    datasets: [{{ data: CHART_DATA.pedidosStatus.values, backgroundColor: [COR_OK, COR_WARN, COR_BAD] }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-pedidos-entrada'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.pedidosEntrada.labels,
    datasets: [{{ data: CHART_DATA.pedidosEntrada.values, backgroundColor: [COR_OK, COR_WARN] }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-pedidos-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.pedidosFilial.labels,
    datasets: [
      {{ label: 'Entregue', data: CHART_DATA.pedidosFilial.entregue, backgroundColor: COR_OK }},
      {{ label: 'Pendente', data: CHART_DATA.pedidosFilial.pendente, backgroundColor: COR_WARN }},
      {{ label: 'Atrasado', data: CHART_DATA.pedidosFilial.atrasado, backgroundColor: COR_BAD }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true }} }},
    plugins: {{ legend: {{ position: 'bottom' }} }}
  }}
}});

new Chart(document.getElementById('chart-notas-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.notasStatus.labels,
    datasets: [{{ data: CHART_DATA.notasStatus.values, backgroundColor: [COR_BAD, COR_OK] }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-transf-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.transfStatus.labels,
    datasets: [{{ data: CHART_DATA.transfStatus.values, backgroundColor: [COR_OK, COR_WARN, COR_BAD] }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-transf-prazo'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.transfPrazo.labels,
    datasets: [{{ label: 'Transferências', data: CHART_DATA.transfPrazo.values, backgroundColor: COR_ACCENT }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
}});
</script>
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(f"OK: {OUT} gerado com {len(notas_atrasadas)} notas atrasadas e {len(transf_criticas)} transferências críticas.")
