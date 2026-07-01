# -*- coding: utf-8 -*-
"""Gera a página HTML do BI Operacional a partir das planilhas de controle.

Uso: python gerar_pagina.py
Lê os 3 arquivos em "I:\\Meu Drive\\CONTROLE COMPRAS" e grava index.html
nesta mesma pasta (bi-web), pronta para publicar no GitHub Pages.
"""
import openpyxl
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
<style>
  :root {{
    --bg: #0f172a; --card: #1e293b; --card2: #16213a; --text: #e2e8f0; --muted: #94a3b8;
    --ok: #22c55e; --warn: #f59e0b; --bad: #ef4444; --accent: #38bdf8; --border: #334155;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); }}
  header {{ padding: 24px 32px; border-bottom: 1px solid var(--border); }}
  header h1 {{ margin: 0 0 4px; font-size: 22px; }}
  header p {{ margin: 0; color: var(--muted); font-size: 14px; }}
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
  input.filtro {{ width: 100%; max-width: 320px; padding: 8px 10px; margin-bottom: 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); }}
  footer {{ text-align: center; color: var(--muted); font-size: 12px; padding: 24px; }}
</style>
</head>
<body>
<header>
  <h1>📊 BI Operacional — Pedidos, Notas Fiscais e Transferências</h1>
  <p>Atualizado em {data_atualizacao} · Giovanni Brochini</p>
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
</script>
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(f"OK: {OUT} gerado com {len(notas_atrasadas)} notas atrasadas e {len(transf_criticas)} transferências críticas.")
