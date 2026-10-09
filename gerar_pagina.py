# -*- coding: utf-8 -*-
"""Gera a página HTML do Controle Operacional a partir das planilhas de controle.

Uso: python gerar_pagina.py
Lê os 3 arquivos em "C:\\Users\\recru\\Desktop\\CONTROLE COMPRAS" e grava index.html
nesta mesma pasta (bi-web), pronta para publicar no GitHub Pages.
"""
import collections
import json
import re
import unicodedata
import openpyxl
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

FAVICON_SVG = (
    "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
    "<rect width='64' height='64' rx='14' fill='#0f172a'/>"
    "<rect x='13' y='34' width='10' height='18' rx='2' fill='#38bdf8'/>"
    "<rect x='27' y='22' width='10' height='30' rx='2' fill='#22c55e'/>"
    "<rect x='41' y='12' width='10' height='40' rx='2' fill='#f59e0b'/>"
    "</svg>"
)
FAVICON_HREF = "data:image/svg+xml," + quote(FAVICON_SVG)

BASE = Path(r"C:\Users\recru\Desktop\CONTROLE COMPRAS")
# malotes e campanhas continuam no Google Drive, porque outras pessoas alimentam esses dois
BASE_DRIVE = Path(r"I:\Meu Drive\CONTROLE COMPRAS")
NO_DRIVE = {"CONTROLE MALOTES.xlsm", "CONTROLE DE CAMPANHAS.xlsx"}


def caminho(arquivo):
    return (BASE_DRIVE if arquivo in NO_DRIVE else BASE) / arquivo
OUT = Path(__file__).resolve().parent / "index.html"


def load(arquivo, aba):
    wb = openpyxl.load_workbook(caminho(arquivo), data_only=True, read_only=True)
    return list(wb[aba].iter_rows(values_only=True))


def load_dicts(arquivo, aba):
    """Lê a aba usando a 1ª linha como cabeçalho e devolve dicts (nome da coluna -> valor).
    Mais robusto que acessar por índice fixo: continua funcionando mesmo se colunas
    forem inseridas/removidas/reordenadas na planilha."""
    linhas = load(arquivo, aba)
    header = [(h or f"_col{i}").strip() if isinstance(h, str) else (h or f"_col{i}") for i, h in enumerate(linhas[0])]
    return [dict(zip(header, r)) for r in linhas[1:]]


def load_dicts_skip(arquivo, aba, marcador_cabecalho):
    """Como load_dicts, mas pula linhas de título até achar a linha cujo 1º valor bate com
    marcador_cabecalho — usa essa como cabeçalho. Para abas que têm um título antes da tabela."""
    linhas = load(arquivo, aba)
    idx = next(i for i, r in enumerate(linhas) if r and r[0] == marcador_cabecalho)
    header = [(h or f"_col{i}").strip() if isinstance(h, str) else (h or f"_col{i}") for i, h in enumerate(linhas[idx])]
    return [dict(zip(header, r)) for r in linhas[idx + 1:]]


def mtime_str(arquivo):
    ts = caminho(arquivo).stat().st_mtime
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y às %H:%M")


def pct(v):
    return f"{v * 100:.1f}%".replace(".", ",")


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_str(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v)


# ---------------------------------------------------------------- PEDIDOS
pedidos_mtime = mtime_str("CONTROLE DE PEDIDOS.xlsx")
rows = load("CONTROLE DE PEDIDOS.xlsx", "RESUMO")


data_atualizacao = next(
    r[1].replace("Atualizado em ", "").strip() for r in rows if r[1] and str(r[1]).startswith("Atualizado em")
)
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

# tabela detalhada de pedidos (igual a "TABELA DE PEDIDOS REALIZADOS NO GN" do BI)
# por nome de coluna: a aba ja ganhou uma coluna nova no meio (REFERENCIA) e, com
# indice fixo, isso desloca silenciosamente todo o resto da tabela sem dar erro
pedidos_detalhe = load_dicts("CONTROLE DE PEDIDOS.xlsx", "CONTROLE PEDIDOS")

COLUNAS_PEDIDOS = [
    "FILIAL DESTINO", "N° DO PEDIDO NO GN", "DATA DE REALIZAÇÃO DO PEDIDO",
    "DATA DE PROCESSAMENTO", "REFERENCIA", "DESCRIÇÃO DO PRODUTO", "QUANTIDADE",
    "STATUS DO PRODUTO", "DATA PREVISTA DE ENTRADA", "STATUS", "DIAS EM ABERTO",
    "REALIZADO ENTRADA NO SISTEMA ?",
]
_faltando = [c for c in COLUNAS_PEDIDOS if c not in pedidos_detalhe[0]]
if _faltando:
    raise SystemExit(
        "CONTROLE DE PEDIDOS.xlsx / aba CONTROLE PEDIDOS: colunas nao encontradas: "
        + ", ".join(_faltando)
        + "\nColunas disponiveis: "
        + ", ".join(str(c) for c in pedidos_detalhe[0])
    )

# a aba tem formula arrastada muito abaixo da ultima linha preenchida, o que gera
# milhares de linhas so com STATUS='ATRASADO' e todo o resto vazio
pedidos_detalhe = [r for r in pedidos_detalhe if r.get("FILIAL DESTINO")]
pedidos_detalhe.sort(key=lambda r: r.get("DATA DE REALIZAÇÃO DO PEDIDO") or datetime.min, reverse=True)

# --------------------------------------------------------- NOTAS FISCAIS
notas_mtime = mtime_str("CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx")
# o resumo ficava nas colunas N/O da propria aba e passou a ter aba propria
notas_resumo = {}
for r in load("CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx", "RESUMO"):
    label, valor = (list(r) + [None, None])[:2]
    label = str(label).strip() if label else ""
    if label == "FILIAL":
        break          # daqui pra baixo comeca a tabela por filial
    if label and label != "INDICADOR" and valor is not None:
        notas_resumo[label] = valor

# por nome de coluna: a aba ja mudou de largura e o indice fixo quebrava sem avisar
COLUNAS_NOTAS = ["FILIAIS", "NOTA FISCAL", "DESCRIÇÃO", "QUANTIDADE UNITÁRIA",
                 "DIA DA ENTREGA", "DIAS EM ATRASO", "STATUS"]
notas_linhas = load_dicts("CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx", "NOTAS FISCAIS PENDENTES")
_faltando = [c for c in COLUNAS_NOTAS if notas_linhas and c not in notas_linhas[0]]
if _faltando:
    raise SystemExit(
        "CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx / aba NOTAS FISCAIS PENDENTES: "
        "colunas nao encontradas: " + ", ".join(_faltando)
        + "\nColunas disponiveis: " + ", ".join(str(c) for c in (notas_linhas[0] if notas_linhas else []))
    )

notas_atrasadas = [r for r in notas_linhas if str(r.get("STATUS") or "").strip().upper() == "ATRASADO"]
notas_atrasadas.sort(key=lambda r: (r.get("DIAS EM ATRASO") or 0), reverse=True)

# a aba tem uma linha por produto, entao uma NF com 5 itens contava 5 vezes nos cards.
# Os indicadores de quantidade passam a contar notas fiscais distintas; o valor pendente
# continua somando as linhas, que e onde cada item tem seu valor.
def _nfs(linhas):
    return {str(r.get("NOTA FISCAL")).strip() for r in linhas if r.get("NOTA FISCAL")}


_notas_pendentes = [r for r in notas_linhas
                    if str(r.get("REALIZADO ENTRADA NO SISTEMA ?") or "").strip().upper() in ("NÃO", "NAO")]
_status_nota = lambda r: str(r.get("STATUS") or "").strip().upper()

# maior atraso de cada NF, para a media nao pesar mais quem tem mais itens
_atraso_por_nf = {}
for r in notas_atrasadas:
    nf = str(r.get("NOTA FISCAL")).strip()
    _atraso_por_nf[nf] = max(_atraso_por_nf.get(nf, 0), r.get("DIAS EM ATRASO") or 0)

notas_resumo.update({
    "Total de notas (linhas)": len(_nfs(notas_linhas)),
    "Pendentes (NÃO)": len(_nfs(_notas_pendentes)),
    "Atrasadas": len(_nfs([r for r in _notas_pendentes if _status_nota(r) == "ATRASADO"])),
    "Aguardando (no prazo)": len(_nfs([r for r in _notas_pendentes if _status_nota(r) == "AGUARDANDO"])),
    "Média de atraso (dias)": (sum(_atraso_por_nf.values()) / len(_atraso_por_nf)) if _atraso_por_nf else 0,
})
# produtos fora do fluxo normal (voltaram para a TIM, extraviados, recusa de nota...). Entram pelo
# status, mesmo que a entrada no sistema ja tenha sido feita, porque o status e que pede acompanhamento.
STATUS_NOTA_NORMAIS = {"ENTRADA OK", "ATRASADO", "AGUARDANDO"}
notas_excecoes = [r for r in notas_linhas if _status_nota(r) and _status_nota(r) not in STATUS_NOTA_NORMAIS]
notas_excecoes.sort(key=lambda r: (_status_nota(r), str(r.get("FILIAIS") or ""), str(r.get("NOTA FISCAL") or "")))
notas_excecoes_resumo = {
    "produtos": len(notas_excecoes),
    "notas": len(_nfs(notas_excecoes)),
    "valor": round(sum(r.get("VALOR TOTAL") or 0 for r in notas_excecoes), 2),
    "por_status": sorted(collections.Counter(str(r.get("STATUS")).strip() for r in notas_excecoes).items(),
                         key=lambda kv: -kv[1]),
}

_ordem_status_nota = {"ATRASADO": 0, "AGUARDANDO": 1}
# todas as pendentes: atrasadas primeiro (maior atraso), depois as no prazo e por fim os outros status
notas_pendentes_tabela = sorted(
    _notas_pendentes,
    key=lambda r: (_ordem_status_nota.get(_status_nota(r), 2), -(r.get("DIAS EM ATRASO") or 0),
                   str(r.get("FILIAIS") or ""), str(r.get("NOTA FISCAL") or "")),
)

notas_linhas_resumo = {
    "total": len(notas_linhas),
    "pendentes": len(_notas_pendentes),
    "excecoes": len(_nfs([r for r in _notas_pendentes
                          if _status_nota(r) not in ("ATRASADO", "AGUARDANDO")])),
}

# ------------------------------------------------------- TRANSFERÊNCIAS
transf_mtime = mtime_str("TRANSFERÊNCIAS PENDENTES.xlsx")
rows = load("TRANSFERÊNCIAS PENDENTES.xlsx", "TRANSFERENCIAS PENDENTES")
data_rows = rows[1:]

# o resumo ficava nas colunas O/P da propria aba e passou a ter aba propria
transf_resumo = {}
for r in load("TRANSFERÊNCIAS PENDENTES.xlsx", "RESUMO"):
    label, valor = (list(r) + [None, None])[1:3]
    if label and label != "RESUMO GERAL" and valor is not None:
        transf_resumo[str(label).strip()] = valor

transf_criticas = [r for r in data_rows if r[11] == "CRÍTICO"]

transf_todas = [r for r in data_rows if r[0] is not None]
transf_todas.sort(key=lambda r: (r[2] or 0), reverse=True)

# a aba tem uma linha por produto, entao a mesma NF contava varias vezes nos cards e
# graficos. Por NF, criticidade, prazo e faixa de atraso sao sempre iguais entre as
# linhas, entao contar a NF uma vez so nao perde informacao. Unidades continuam somadas.
_transf_por_nf = {}
for r in transf_todas:
    _transf_por_nf.setdefault(str(r[5] or "").strip(), r)
_transf_nfs = [r for nf, r in _transf_por_nf.items() if nf]


def _conta_nf(coluna, valor):
    return sum(1 for r in _transf_nfs if str(r[coluna] or "").strip().upper() == valor)


for _faixa in ("0-3 dias", "4-7 dias", "8-15 dias", "16-30 dias", "30+ dias"):
    transf_resumo[_faixa] = _conta_nf(10, _faixa.upper())
transf_resumo.update({
    "Total de linhas": len(_transf_nfs),
    "No prazo": _conta_nf(12, "NO PRAZO"),
    "Atrasados": _conta_nf(12, "ATRASADO"),
    "ATENÇÃO": _conta_nf(11, "ATENÇÃO"),
    "CRÍTICO": _conta_nf(11, "CRÍTICO"),
})
transf_linhas_total = len(transf_todas)

# -------------------------------------------------- PELÍCULAS (AMET / DEVIA / UPMASTER)
# lê pelas posições do cabeçalho (linha com "Filial"/"Total Geral") em vez de índice fixo de
# coluna, pois a planilha já mudou de estrutura (colunas de produto adicionadas/removidas).
def localizar_colunas_estoque_vendas(rows):
    header = next(r for r in rows if r[0] == "Filial")
    col_filial_1, col_total_1 = 0, header.index("Total Geral")
    col_filial_2 = header.index("Filial", col_total_1 + 1)
    col_total_2 = header.index("Total Geral", col_filial_2 + 1)
    return col_filial_1, col_total_1, col_filial_2, col_total_2


def capturar_totais(rows, col_filial, col_total):
    out = {}
    capturando = False
    for r in rows:
        if len(r) <= max(col_filial, col_total):
            continue
        if r[col_filial] == "Filial":
            capturando = True
            continue
        if capturando:
            if r[col_filial] is None:
                continue
            if r[col_filial] == "Total Geral":
                break
            out[r[col_filial]] = r[col_total] or 0
    return out


def limpar_nome_produto(nome, marca):
    n = nome
    if "(" in n and "un por caixa)" in n:
        n = n.split("(")[0]
    n = n.strip().replace("PELICULA ", "").replace(f"{marca} ", "").strip()
    return n.title()


def capturar_produtos_e_totais(rows, col_filial, col_total):
    """Como capturar_totais, mas também devolve a quantidade de cada produto individual
    (as colunas entre 'Filial' e 'Total Geral'), não só o total."""
    header = next(r for r in rows if r[col_filial] == "Filial")
    produtos_cols = list(range(col_filial + 1, col_total))
    produtos_labels_brutos = [header[c] for c in produtos_cols]
    out = {}
    capturando = False
    for r in rows:
        if len(r) <= col_total:
            continue
        if r[col_filial] == "Filial":
            capturando = True
            continue
        if capturando:
            if r[col_filial] is None:
                continue
            if r[col_filial] == "Total Geral":
                break
            produtos = {produtos_labels_brutos[i]: (r[c] or 0) for i, c in enumerate(produtos_cols)}
            out[r[col_filial]] = {"produtos": produtos, "total": r[col_total] or 0}
    return produtos_labels_brutos, out


def carregar_pelicula(arquivo, aba, marca, filiais_mestre=None):
    mtime = mtime_str(arquivo)
    rows = load(arquivo, aba)
    data_estoque = (rows[0][0] or "").split("ATUALIZADO DIA ")[-1].rstrip(")")

    col_filial_1, col_total_1, col_filial_2, col_total_2 = localizar_colunas_estoque_vendas(rows)
    periodo_vendas = (rows[0][col_filial_2] or "").split("(DO DIA ")[-1].rstrip(")")

    produtos_brutos, estoque_info = capturar_produtos_e_totais(rows, col_filial_1, col_total_1)
    vendido_por_filial = capturar_totais(rows, col_filial_2, col_total_2)
    produtos_labels = [limpar_nome_produto(p, marca) for p in produtos_brutos]

    nomes = set(estoque_info) | set(vendido_por_filial)
    if filiais_mestre:
        nomes |= set(filiais_mestre)

    filiais = []
    for nome in sorted(nomes):
        info = estoque_info.get(nome, {"produtos": {}, "total": 0})
        produtos_limpos = {limpo: info["produtos"].get(bruto, 0) for bruto, limpo in zip(produtos_brutos, produtos_labels)}
        filiais.append({
            "filial": nome, "produtos": produtos_limpos,
            "estoque": info["total"], "vendido": vendido_por_filial.get(nome, 0),
        })
    filiais.sort(key=lambda f: f["vendido"], reverse=True)

    return {
        "mtime": mtime, "data_estoque": data_estoque, "periodo_vendas": periodo_vendas,
        "filiais": filiais, "produtos_labels": produtos_labels,
        "estoque_total": sum(f["estoque"] for f in filiais),
        "vendido_total": sum(f["vendido"] for f in filiais),
    }


# ------------------------------------------------------------- AMET
amet = carregar_pelicula("CONTROLE QUANTIDADE DE AMET NAS FILIAIS.xlsx", "AMET NAS FILIAIS", "AMET")

# ------------------------------------------------------- ACESSÓRIOS
acessorios_mtime = mtime_str("CONTROLE CONFIGURAÇÕES PRODUTOS.xlsx")
acessorios_rows = load_dicts("CONTROLE CONFIGURAÇÕES PRODUTOS.xlsx", "SALDO PRODUTOS NAS FILIAIS")
acessorios_rows = [r for r in acessorios_rows if r.get("Filial") is not None]

FILIAIS_MESTRE = sorted(set(r.get("Filial") for r in acessorios_rows if r.get("Filial")))

# mostra as 35 filiais da lista mestra, mesmo as que não têm nenhum registro na planilha
devia = carregar_pelicula("CONTROLE QUANTIDADE DE DEVIA NAS FILIAIS.xlsx", "DEVIA NAS FILIAIS", "DEVIA", FILIAIS_MESTRE)
upmaster = carregar_pelicula("CONTROLE QUANTIDADE DE UPMASTER NAS FILIAIS.xlsx", "Planilha1", "UPMASTER", FILIAIS_MESTRE)


def _nome_col(c):
    """Nome da coluna sem acento e em maiusculas: a exportacao troca "Valor de venda" por
    "VALOR DE VENDA" de um mes para o outro, e o valor acabava zerado."""
    return unicodedata.normalize("NFKD", str(c)).encode("ascii", "ignore").decode("ascii").upper().strip()


def _coluna(r, *nomes):
    alvo = {_nome_col(n) for n in nomes}
    for k, v in r.items():
        if isinstance(k, str) and _nome_col(k) in alvo:
            return v
    return None


def _numero(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def montar_acessorios(rows):
    itens = []
    saldo_por_filial = {}
    for r in rows:
        filial = r.get("Filial") or ""
        saldo = _numero(_coluna(r, "Saldo"))
        valor = _numero(_coluna(r, "Valor de venda", "Valor Venda", "Preco de venda"))
        # a planilha ja traz o total; so recalculo quando ela nao tiver a coluna
        total = _numero(_coluna(r, "Valor total", "Valor Total Venda")) or saldo * valor
        itens.append({
            "filial": filial, "ref": r.get("Produto") or "", "desc": r.get("Descrição") or "",
            "subgrupo": r.get("Sub Grupo Estoque") or "-", "fabricante": r.get("Fabricante") or "-",
            "saldo": saldo, "disponivel": _numero(_coluna(r, "Disponível")),
            "valor": valor, "valor_total": round(total, 2),
        })
        saldo_por_filial[filial] = saldo_por_filial.get(filial, 0) + saldo
    itens.sort(key=lambda x: x["saldo"], reverse=True)
    filiais_ordenadas = sorted(saldo_por_filial.items(), key=lambda kv: kv[1], reverse=True)
    resumo = {
        "itens": len(itens),
        "saldo_total": sum(i["saldo"] for i in itens),
        "valor_total": sum(i["valor_total"] for i in itens),
        "filiais": len(saldo_por_filial),
    }
    return itens, filiais_ordenadas, resumo


acessorios_diversos_rows = [r for r in acessorios_rows if r.get("Grupo Estoque") == "ACESSORIOS DIVERSOS"]
acessorios_diversos_itens, acessorios_diversos_filiais, acessorios_diversos_resumo = montar_acessorios(acessorios_diversos_rows)
acessorios_diversos_json = json.dumps(acessorios_diversos_itens, ensure_ascii=False)

# ------------------------------------- ACESSÓRIOS FIDELIZADOS TIM (com número de série)
seriais_rows = load_dicts("CONTROLE CONFIGURAÇÕES PRODUTOS.xlsx", "SERIAIS ACESSÓRIOS FIDELIZADOS")
seriais_rows = [r for r in seriais_rows if r.get("Filial Atual") is not None]


def faixa_dias_estoque(d):
    if d <= 30:
        return "0-30 dias"
    if d <= 90:
        return "31-90 dias"
    if d <= 180:
        return "91-180 dias"
    if d <= 365:
        return "181-365 dias"
    return "365+ dias"


seriais_itens = []
seriais_por_filial = {}
for r in seriais_rows:
    filial = r.get("Filial Atual") or ""
    valor = r.get("Valor Venda") or 0
    dias_val = r.get("Dias em Estoque")
    dias = dias_val if isinstance(dias_val, (int, float)) else 0
    produto = r.get("Produto") or ""
    seriais_itens.append({
        "filial": filial, "serial": r.get("Serial") or "", "produto": produto.lstrip("'") if isinstance(produto, str) else produto,
        "desc": r.get("Descricao") or "", "fabricante": r.get("Fabricante") or "-", "data_compra": data_str(r.get("Data Compra")),
        "dias": dias, "dias_faixa": faixa_dias_estoque(dias), "valor": valor,
    })
    seriais_por_filial[filial] = seriais_por_filial.get(filial, 0) + 1
seriais_itens.sort(key=lambda x: x["dias"], reverse=True)
seriais_filiais_ordenadas = sorted(seriais_por_filial.items(), key=lambda kv: kv[1], reverse=True)

seriais_resumo = {
    "itens": len(seriais_itens),
    "valor_total": sum(i["valor"] for i in seriais_itens),
    "filiais": len(seriais_por_filial),
    "dias_medio": round(sum(i["dias"] for i in seriais_itens) / len(seriais_itens), 1) if seriais_itens else 0,
}
seriais_json = json.dumps(seriais_itens, ensure_ascii=False)

# ---------------------------------------------- DEVOLVIDOS E DEFEITOS
devolvidos_mtime = mtime_str("CONTROLE DEVOLVIDOS E DEFEITOS.xlsx")
devolvidos_rows = load_dicts("CONTROLE DEVOLVIDOS E DEFEITOS.xlsx", "DEVOLVIDOS E DEFEITOS")
devolvidos_rows = [r for r in devolvidos_rows if r.get("Filial") is not None]

devolvidos_itens = []
devolvidos_por_filial = {}
for r in devolvidos_rows:
    filial = r.get("Filial") or ""
    saldo = r.get("Saldo") or 0
    custo = r.get("Custo Movimento") or r.get("Custo Padrão") or 0
    devolvidos_itens.append({
        "filial": filial, "desc": r.get("Descrição") or "", "grupo": r.get("Grupo Estoque") or "-",
        "fabricante": r.get("Fabricante") or "-", "saldo": saldo, "custo": custo,
        "custo_total": round(saldo * custo, 2), "data_mov": data_str(r.get("Data Movimento")),
    })
    devolvidos_por_filial[filial] = devolvidos_por_filial.get(filial, 0) + saldo
devolvidos_itens.sort(key=lambda x: x["saldo"], reverse=True)
devolvidos_filiais_ordenadas = sorted(devolvidos_por_filial.items(), key=lambda kv: kv[1], reverse=True)

devolvidos_resumo = {
    "itens": len(devolvidos_itens),
    "saldo_total": sum(i["saldo"] for i in devolvidos_itens),
    "custo_total": sum(i["custo_total"] for i in devolvidos_itens),
    "filiais": len(devolvidos_por_filial),
}
devolvidos_json = json.dumps(devolvidos_itens, ensure_ascii=False)

# ---------------------------------------------------------------- MALOTES
malotes_mtime = mtime_str("CONTROLE MALOTES.xlsm")

MALOTE_STATUS_ADM_CLASSE = {"CRÍTICO": "bad", "BAIXO": "warn", "NORMAL": "ok", "EXCEDENTE": "ok"}

malotes_capacidade = {r.get("FILIAIS"): r.get("QTD MALOTE") for r in load_dicts("CONTROLE MALOTES.xlsm", "TB_CAPACIDADE")}

malotes_filiais = []
for r in load_dicts("CONTROLE MALOTES.xlsm", "BASE_DADOS"):
    filial = r.get("FILIAIS")
    if not filial:
        continue
    status_adm = r.get("STATUS ADM") or "-"
    malotes_filiais.append({
        "filial": filial, "na_filial": r.get("MALOTES NA FILIAL") or 0,
        "no_adm": r.get("MALOTES NO ADM") or 0, "capacidade": malotes_capacidade.get(filial, "-"),
        "status_adm": status_adm, "status_cls": MALOTE_STATUS_ADM_CLASSE.get(status_adm, ""),
        "acao": r.get("AÇÃO") or "-",
    })
malotes_filiais.sort(key=lambda f: f["no_adm"])

malotes_resumo = {
    "total_parque": sum(f["na_filial"] + f["no_adm"] for f in malotes_filiais),
    "no_adm": sum(f["no_adm"] for f in malotes_filiais),
    "nas_filiais": sum(f["na_filial"] for f in malotes_filiais),
    "filiais": len(malotes_filiais),
    "sem_malote_adm": sum(1 for f in malotes_filiais if f["no_adm"] == 0),
}

malotes_status_adm_buckets = {}
for f in malotes_filiais:
    malotes_status_adm_buckets[f["status_adm"]] = malotes_status_adm_buckets.get(f["status_adm"], 0) + 1
MALOTE_ORDEM_STATUS_ADM = ["CRÍTICO", "BAIXO", "NORMAL", "EXCEDENTE"]

MALOTE_STATUS_LOG_CLASSE = {"POSTADO": "ok", "PENDENTE": "warn", "CANCELADO": "bad"}


def malote_epoch(dt):
    return int(dt.timestamp() * 1000) if hasattr(dt, "timestamp") else 0


malotes_log_itens = []
for r in load_dicts_skip("CONTROLE MALOTES.xlsm", "GERAL", "ID"):
    if r.get("ID") is None:
        continue
    status = r.get("STATUS") or "-"
    dt_sol = r.get("DATA SOLICITAÇÃO")
    dt_post = r.get("DATA POSTAGEM")
    horas = None
    if hasattr(dt_sol, "timestamp") and hasattr(dt_post, "timestamp"):
        horas = round((dt_post - dt_sol).total_seconds() / 3600, 1)
    malotes_log_itens.append({
        "id": r.get("ID"), "solicitante": r.get("SOLICITANTE") or "-",
        "filial": r.get("FILIAL DESTINO") or "-", "qtd": r.get("QUANTIDADE") or 0,
        "conteudo": r.get("CONTEÚDO") or "-",
        "data_sol": data_str(dt_sol), "data_sol_ts": malote_epoch(dt_sol),
        "status": status, "status_cls": MALOTE_STATUS_LOG_CLASSE.get(status, ""),
        "data_post": data_str(dt_post), "quem_postou": r.get("QUEM POSTOU?") or "-",
        "baixado_adm": r.get("BAIXADO ADM") or "-", "obs": r.get("OBSERVAÇÃO") or "-",
        "horas_postagem": horas if horas is not None else "-",
    })
malotes_log_itens.sort(key=lambda x: x["data_sol_ts"], reverse=True)

horas_validas = [x["horas_postagem"] for x in malotes_log_itens if x["horas_postagem"] != "-"]
malotes_log_resumo = {
    "total": len(malotes_log_itens),
    "pendentes": sum(1 for x in malotes_log_itens if x["status"] == "PENDENTE"),
    "postados": sum(1 for x in malotes_log_itens if x["status"] == "POSTADO"),
    "tempo_medio_h": round(sum(horas_validas) / len(horas_validas), 1) if horas_validas else 0,
}
malotes_log_json = json.dumps(malotes_log_itens, ensure_ascii=False)

# ---------------------------------------------------------- MANUTENÇÕES
manutencoes_mtime = mtime_str("CONTROLE DE MANUTENÇÕES.xlsx")

def classificar_status_manutencao(status):
    """Devolve (grupo, classe css) para um status escrito à mão na planilha.

    A mesma conclusão aparece como "CONCLUÍDA" ou "MANUTENÇÃO REALIZADA", então
    casar o texto exato deixava chamados concluídos fora da contagem. Aqui o
    reconhecimento é por radical, sem acento e sem depender da frase inteira."""
    s = unicodedata.normalize("NFKD", str(status).upper()).encode("ascii", "ignore").decode("ascii")
    negado = "NAO " in s or "NAO-" in s
    if not negado and any(p in s for p in ("CONCLU", "REALIZAD", "FINALIZAD", "RESOLVID")):
        return "concluida", "ok"
    if "CANCELAD" in s:
        return "cancelada", "bad"
    if any(p in s for p in ("ANDAMENTO", "EXECUCAO")):
        return "andamento", "warn"
    if "PENDENTE" in s or negado:
        return "pendente", "warn"
    return "outro", ""

# a mesma aba guarda os chamados de manutenção e as obras pendentes das lojas,
# separados pela coluna TIPO. obras não têm chamado nem contagem de dias.
manutencoes_itens = []
obras_itens = []
for r in load_dicts("CONTROLE DE MANUTENÇÕES.xlsx", "CONTROLE MANUTENÇÕES"):
    status = str(r.get("STATUS") or "-").strip()
    if str(r.get("TIPO") or "").strip().upper() == "OBRA":
        obras_itens.append({
            "filial": r.get("FILIAL") or "-",
            "obra": r.get("MANUTENÇÕES SOLICITADAS") or "-",
            "responsavel": str(r.get("RESPONSÁVEL") or "-").strip() or "-",
            "prazo": str(r.get("PRAZO") or "-").strip() or "-",
            "obs": r.get("OBSERVAÇÃO") or "-",
            "status": status, "status_cls": classificar_status_manutencao(status)[1],
        })
        continue
    chamado = r.get("N° DO CHAMADO")
    if not chamado:
        continue
    dias = r.get("DIAS SEM CONCLUSÃO")
    grupo, classe = classificar_status_manutencao(status)
    manutencoes_itens.append({
        "chamado": chamado, "filial": r.get("FILIAL") or "-",
        "solicitado": r.get("MANUTENÇÕES SOLICITADAS") or "-",
        "status": status, "status_cls": classe, "grupo": grupo,
        "dias": dias if isinstance(dias, (int, float)) else "-",
        "obs": r.get("OBSERVAÇÃO") or "-",
    })
manutencoes_itens.sort(key=lambda x: x["dias"] if isinstance(x["dias"], (int, float)) else -1, reverse=True)

manutencoes_por_filial = {}
for it in manutencoes_itens:
    manutencoes_por_filial[it["filial"]] = manutencoes_por_filial.get(it["filial"], 0) + 1
manutencoes_filiais_ordenadas = sorted(manutencoes_por_filial.items(), key=lambda kv: kv[1], reverse=True)

manutencoes_status_buckets = {}
manutencoes_status_classes = {}
for it in manutencoes_itens:
    manutencoes_status_buckets[it["status"]] = manutencoes_status_buckets.get(it["status"], 0) + 1
    manutencoes_status_classes[it["status"]] = it["status_cls"]

dias_manutencoes_validos = [x["dias"] for x in manutencoes_itens if isinstance(x["dias"], (int, float))]
manutencoes_resumo = {
    "total": len(manutencoes_itens),
    "pendentes": sum(1 for x in manutencoes_itens if x["grupo"] == "pendente"),
    "concluidas": sum(1 for x in manutencoes_itens if x["grupo"] == "concluida"),
    "filiais": len(manutencoes_por_filial),
    "dias_medio": round(sum(dias_manutencoes_validos) / len(dias_manutencoes_validos), 1) if dias_manutencoes_validos else 0,
}

obras_por_filial = {}
obras_por_responsavel = {}
for it in obras_itens:
    obras_por_filial[it["filial"]] = obras_por_filial.get(it["filial"], 0) + 1
    obras_por_responsavel[it["responsavel"]] = obras_por_responsavel.get(it["responsavel"], 0) + 1
obras_filiais_ordenadas = sorted(obras_por_filial.items(), key=lambda kv: kv[1], reverse=True)
obras_responsaveis_ordenados = sorted(
    ((k, v) for k, v in obras_por_responsavel.items() if k != "-"), key=lambda kv: kv[1], reverse=True
)

_prazos = [it["prazo"] for it in obras_itens if it["prazo"] != "-"]
obras_resumo = {
    "total": len(obras_itens),
    "filiais": len(obras_por_filial),
    "sem_responsavel": obras_por_responsavel.get("-", 0),
    "prazo": max(set(_prazos), key=_prazos.count) if _prazos else "-",
}

# -------------------------------------------------------------- DESPESAS
despesas_mtime = mtime_str("CONTROLE DE DESPESAS.xlsx")

DATA_INICIO_DESPESAS = datetime(2026, 1, 1)
DATA_FIM_DESPESAS = datetime.now()


def _despesa_no_periodo(dt):
    return hasattr(dt, "timestamp") and DATA_INICIO_DESPESAS <= dt <= DATA_FIM_DESPESAS


despesas_itens = []

for r in load_dicts("CONTROLE DE DESPESAS.xlsx", "MATERIAL DE LIMPEZA"):
    if not r.get("Filial"):
        continue
    dt = r.get("Vencimento")
    if not _despesa_no_periodo(dt):
        continue
    despesas_itens.append({
        "filial": r.get("Filial"), "categoria": "Material de Limpeza",
        "fornecedor": (r.get("Fornecedor") or "-").strip(), "documento": str(r.get("Título") or "-"),
        "data": data_str(dt), "data_ts": malote_epoch(dt),
        "valor": round(r.get("Valor Pago") or 0, 2), "historico": r.get("Histórico") or "-",
    })

for r in load_dicts("CONTROLE DE DESPESAS.xlsx", "MATERIAL DE ESCRITÓRIO"):
    if not r.get("Filial"):
        continue
    dt = r.get("Vencimento")
    if not _despesa_no_periodo(dt):
        continue
    despesas_itens.append({
        "filial": r.get("Filial"), "categoria": "Material de Escritório",
        "fornecedor": (r.get("Fornecedor") or "-").strip(), "documento": str(r.get("Título") or "-"),
        "data": data_str(dt), "data_ts": malote_epoch(dt),
        "valor": round(r.get("Valor Pago") or 0, 2), "historico": r.get("Histórico") or "-",
    })

for r in load_dicts("CONTROLE DE DESPESAS.xlsx", "REGISTRO MANUAL TOTAL"):
    if not r.get("Loja"):
        continue
    dt = r.get("Data")
    if not _despesa_no_periodo(dt):
        continue
    despesas_itens.append({
        "filial": r.get("Loja"), "categoria": "Registro Manual",
        "fornecedor": (r.get("Operação") or "-").strip(), "documento": str(r.get("Documento") or "-"),
        "data": data_str(dt), "data_ts": malote_epoch(dt),
        "valor": round(r.get("Valor") or 0, 2), "historico": r.get("Histórico") or "-",
    })

despesas_itens.sort(key=lambda x: x["data_ts"], reverse=True)

despesas_por_categoria = {}
for it in despesas_itens:
    despesas_por_categoria[it["categoria"]] = despesas_por_categoria.get(it["categoria"], 0) + it["valor"]

despesas_por_filial = {}
for it in despesas_itens:
    despesas_por_filial[it["filial"]] = despesas_por_filial.get(it["filial"], 0) + it["valor"]
despesas_filiais_ordenadas = sorted(despesas_por_filial.items(), key=lambda kv: kv[1], reverse=True)

despesas_por_mes = {}
for it in despesas_itens:
    chave = datetime.fromtimestamp(it["data_ts"] / 1000).strftime("%Y-%m")
    bucket = despesas_por_mes.setdefault(chave, {"Material de Limpeza": 0, "Material de Escritório": 0, "Registro Manual": 0})
    bucket[it["categoria"]] += it["valor"]
MESES_PT = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
despesas_meses_ordenados = sorted(despesas_por_mes.keys())
despesas_meses_labels = [f"{MESES_PT[int(m.split('-')[1])]}/{m.split('-')[0]}" for m in despesas_meses_ordenados]

despesas_por_filial_mes = {}
for it in despesas_itens:
    chave = datetime.fromtimestamp(it["data_ts"] / 1000).strftime("%Y-%m")
    bucket = despesas_por_filial_mes.setdefault(it["filial"], {})
    bucket[chave] = bucket.get(chave, 0) + it["valor"]

TOP_N_FILIAIS_DESPESAS = 10
despesas_top_filiais = [f[0] for f in despesas_filiais_ordenadas[:TOP_N_FILIAIS_DESPESAS]]
despesas_filial_mensal_series = [
    {
        "filial": filial,
        "valores": [round(despesas_por_filial_mes.get(filial, {}).get(m, 0), 2) for m in despesas_meses_ordenados],
    }
    for filial in despesas_top_filiais
]

despesas_resumo = {
    "total_geral": round(sum(it["valor"] for it in despesas_itens), 2),
    "total_limpeza": round(despesas_por_categoria.get("Material de Limpeza", 0), 2),
    "total_escritorio": round(despesas_por_categoria.get("Material de Escritório", 0), 2),
    "total_manual": round(despesas_por_categoria.get("Registro Manual", 0), 2),
    "lancamentos": len(despesas_itens),
    "filiais": len(despesas_por_filial),
}
despesas_json = json.dumps(despesas_itens, ensure_ascii=False)

# -------------------------------------------------------------- CAMPANHAS
campanhas_mtime = mtime_str("CONTROLE DE CAMPANHAS.xlsx")

# nomes que aparecem diferentes entre a aba de orçamento e a de lançamentos
ALIAS_FILIAL_CAMPANHAS = {
    "SANTA BARBARA TIVOLI": "SHOPPING TIVOLI",
    "VALINHOS SHOPPING": "VALINHOS",
    "LIMEIRA SHOPPING": "PATIO LIMEIRA",
    # nomes que so apareciam nos lancamentos e nao batiam com a aba de orcamento,
    # entao o gasto dessas 8 filiais sumia da tabela e do total
    "CAMPINAS BANDEIRAS": "SHOPPING BANDEIRAS",
    "GUARATINGUETA SHOPPING": "GUARATINGUETA",
    "HORTOLANDIA SHOPPING": "SHOPPING HORTOLANDIA",
    "INDAIATUBA POLO SHOPPING": "INDAIATUBA SHOPPING",
    "MARILIA CENTRO": "MARILIA",
    "NOVA ODESSA CENTRO": "NOVA ODESSA",
    "RIO CLARO SHOPPING": "SHOPPING RIO CLARO",
    "SUMARE CENTRO": "SUMARE",
    "TAUBATE CENTRO": "TAUBATE",
    "PINDAMONHANGABA CENTRO": "PINDAMONHANGABA",
}
SUFIXOS_FILIAL_CAMPANHAS = (" CENTRO", " SHOPPING")


def normalizar_filial_campanhas(nome):
    if not nome:
        return ""
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode("ascii")
    s = s.strip().upper()
    return ALIAS_FILIAL_CAMPANHAS.get(s, s)


def extrair_area(area_str):
    m = re.match(r"[ÁA]REA\s*-?\s*(\d+)\s*(?:\(([^)]+)\))?", (area_str or "").strip(), re.IGNORECASE)
    if m:
        return m.group(1), (m.group(2) or "").strip().title()
    return "-", ""


campanhas_orcamento = []
_area_num, _area_coord = "-", ""
_vistos_orcamento = set()
for r in load_dicts("CONTROLE DE CAMPANHAS.xlsx", "Planilha2"):
    filial_raw = r.get("FILIAL")
    if not filial_raw:
        continue
    if r.get("ÁREA"):
        _area_num, _area_coord = extrair_area(r.get("ÁREA"))
    filial = normalizar_filial_campanhas(filial_raw)
    if filial in _vistos_orcamento:
        continue
    _vistos_orcamento.add(filial)
    campanhas_orcamento.append({
        "filial": filial, "area_num": _area_num, "coordenador": _area_coord,
        "verba": r.get("DINHEIRO PARA CAMPANHA") or 0,
    })

AREA_COORDENADOR = {o["area_num"]: o["coordenador"] for o in campanhas_orcamento if o["coordenador"]}
_filiais_orcamento = {o["filial"] for o in campanhas_orcamento}


def filial_lancamento_campanhas(nome):
    """Casa o nome do lançamento com o da aba de orçamento.

    Sem isso a filial aparecia como "sem valor aprovado" (ex.: TAUBATE CENTRO x TAUBATE).
    Se não houver alias, tenta sem o sufixo CENTRO/SHOPPING e depois com ele.
    """
    filial = normalizar_filial_campanhas(nome)
    if filial in _filiais_orcamento:
        return filial
    for suf in SUFIXOS_FILIAL_CAMPANHAS:
        if filial.endswith(suf) and filial[:-len(suf)] in _filiais_orcamento:
            return filial[:-len(suf)]
        if filial + suf in _filiais_orcamento:
            return filial + suf
    return filial

def classificar_tipo_campanha(operacao):
    """Separa campanha de verba de área pela coluna Operação.

    Até 20/08/2026 a planilha usava um rótulo único com os dois ("CAMPANHAS / VERBA DE
    ÁREA"), que não dá para desmembrar retroativamente e por isso vira grupo próprio.
    O reconhecimento é por radical para não depender do texto exato dos lançamentos novos."""
    s = unicodedata.normalize("NFKD", str(operacao or "").upper()).encode("ascii", "ignore").decode("ascii")
    tem_campanha, tem_verba = "CAMPANHA" in s, "VERBA" in s
    if tem_campanha and tem_verba:
        return "Não separado"
    if tem_verba:
        return "Verba de Área"
    if tem_campanha:
        return "Campanha"
    return "Outros"


campanhas_lancamentos = []
for r in load_dicts("CONTROLE DE CAMPANHAS.xlsx", "Planilha1"):
    loja = r.get("Loja")
    if not loja:
        continue
    dt = r.get("Data")
    area_num, _ = extrair_area(r.get("ÁREA"))
    campanhas_lancamentos.append({
        "filial": filial_lancamento_campanhas(loja), "area_num": area_num,
        "data": data_str(dt), "data_ts": malote_epoch(dt),
        "documento": str(r.get("Documento") or "-"),
        "tipo": classificar_tipo_campanha(r.get("Operação")),
        "valor": round(r.get("Valor") or 0, 2), "historico": r.get("Histórico") or "-",
    })
campanhas_lancamentos.sort(key=lambda x: x["data_ts"], reverse=True)

campanhas_gasto_por_filial = {}
for it in campanhas_lancamentos:
    campanhas_gasto_por_filial[it["filial"]] = campanhas_gasto_por_filial.get(it["filial"], 0) + it["valor"]

campanhas_filiais = []
for o in campanhas_orcamento:
    verba = o["verba"] or 0
    gasto = round(campanhas_gasto_por_filial.get(o["filial"], 0), 2)
    saldo = round(verba - gasto, 2)
    pct_usado = round((gasto / verba * 100), 1) if verba else 0
    if pct_usado >= 90:
        pct_cls = "bad"
    elif pct_usado >= 60:
        pct_cls = "warn"
    else:
        pct_cls = "ok"
    coord = AREA_COORDENADOR.get(o["area_num"], "")
    campanhas_filiais.append({
        "filial": o["filial"], "area": f"Área {o['area_num']}" + (f" ({coord})" if coord else ""),
        "area_num": o["area_num"], "verba": verba, "gasto": gasto, "saldo": saldo,
        "pct_usado": pct_usado, "pct_cls": pct_cls,
    })
campanhas_filiais.sort(key=lambda f: f["pct_usado"], reverse=True)

campanhas_por_area = {}
for f in campanhas_filiais:
    b = campanhas_por_area.setdefault(f["area_num"], {"verba": 0, "gasto": 0, "coordenador": AREA_COORDENADOR.get(f["area_num"], "")})
    b["verba"] += f["verba"]
    b["gasto"] += f["gasto"]
campanhas_areas_ordenadas = sorted(campanhas_por_area.items(), key=lambda kv: kv[0])

campanhas_lancamentos_json = json.dumps(campanhas_lancamentos, ensure_ascii=False)
campanhas_verba_json = json.dumps(
    [{"filial": f["filial"], "area_num": f["area_num"], "area": f["area"], "verba": f["verba"]}
     for f in campanhas_filiais],
    ensure_ascii=False,
)

# a verba da planilha e MENSAL, mas os lancamentos cobrem o ano todo. sem recorte a
# pagina comparava gasto de 8 meses contra verba de 1 mes e mostrava 596% de uso,
# entao o filtro ja abre no mes corrente.
_hoje_campanhas = datetime.now()
CAMPANHAS_DE_PADRAO = _hoje_campanhas.replace(day=1).strftime("%Y-%m-%d")
CAMPANHAS_ATE_PADRAO = _hoje_campanhas.strftime("%Y-%m-%d")

gerado_em = datetime.now().strftime("%d/%m/%Y %H:%M")


# ----------------------------------------------------------------- HTML
STATUS_NAO_FATURA = "NÃO SERÁ FATURADO"
# a planilha escreve "PENDENTE DE ENTRADA", mas o oposto de ENTREGUE é a entrega. "Entrada"
# aqui é a entrada da nota no sistema, que é outra coluna e tem cards próprios.
STATUS_PENDENTE = "PENDENTE DE ENTREGA"
STATUS_CLASSE = {"ENTREGUE": "ok", STATUS_PENDENTE: "warn", "ATRASADO": "bad"}
CRITICIDADE_CLASSE = {"OK": "ok", "ATENÇÃO": "warn", "CRÍTICO": "bad"}


def epoch(dt):
    return int(dt.timestamp() * 1000) if hasattr(dt, "timestamp") else 0


pedidos_detalhe_data = []
for r in pedidos_detalhe:
    status = r.get("STATUS") or ""
    if status.strip().upper() == "PENDENTE DE ENTRADA":
        status = STATUS_PENDENTE
    dias = r.get("DIAS EM ABERTO")
    dentrega = r.get("DATA DA ENTREGA")
    # "DIAS EM ABERTO" na planilha é HOJE()-DATA DA ENTREGA+1, ou seja, dias desde que o
    # produto chegou na filial. Quando a data de entrega está no futuro a conta fica
    # negativa — a data é de chegada efetiva, então isso é erro de preenchimento e vira
    # uma categoria própria em vez de um número negativo solto na tabela.
    if isinstance(dias, (int, float)):
        dias_num = dias
        if dias < 0:
            dias_cat = "futura"
            dias_txt = "ENTREGA FUTURA" + (f" ({data_str(dentrega)})" if dentrega else "")
        else:
            dias_cat = ""
            dias_txt = dias
    else:
        dias_num = -1
        dias_txt = dias or "-"
        dias_cat = "fin" if "FINALIZADO" in str(dias_txt).upper() else "np"
    # "Em backlog" e "Expirado" não geram OV e não serão faturados, então não são atraso
    # nem pendência — só "OV Criada" pode virar ATRASADO. A planilha ainda decide pela data
    # prevista de entrada, então a reclassificação acontece aqui.
    stprod = str(r.get("STATUS DO PRODUTO") or "").strip()
    if status != "ENTREGUE" and stprod.upper() != "OV CRIADA":
        status = STATUS_NAO_FATURA
        if dias_cat == "np":
            dias_cat, dias_txt = "nfat", STATUS_NAO_FATURA

    dreal, dproc = r.get("DATA DE REALIZAÇÃO DO PEDIDO"), r.get("DATA DE PROCESSAMENTO")
    dprev = r.get("DATA PREVISTA DE ENTRADA")
    pedidos_detalhe_data.append({
        "dreal": data_str(dreal), "dreal_ts": epoch(dreal),
        "dproc": data_str(dproc), "dproc_ts": epoch(dproc),
        "dprev": data_str(dprev), "dprev_ts": epoch(dprev),
        "fil": (r.get("FILIAL DESTINO") or "").replace("ROCHA TELECOM - ", ""),
        "ped": r.get("N° DO PEDIDO NO GN") or "-", "ref": str(r.get("REFERENCIA") or "-"),
        "qtd": r.get("QUANTIDADE"), "desc": r.get("DESCRIÇÃO DO PRODUTO") or "-",
        "stprod": stprod,
        "dias": dias_txt, "dias_num": dias_num, "dias_cat": dias_cat,
        "entrada": r.get("REALIZADO ENTRADA NO SISTEMA ?") or "", "status": status,
        "cls": STATUS_CLASSE.get(status, ""),
    })
pedidos_detalhe_json = json.dumps(pedidos_detalhe_data, ensure_ascii=False)

# os cards e a tabela por filial passam a sair das linhas reais em vez da aba RESUMO,
# que conta pela coluna STATUS da planilha e por isso ainda somaria backlog e expirado
# como atrasados. Também resolve a diferença do RESUMO, que só cobre as filiais da aba SAP.
def _conta_pedidos(pred):
    n = sum(1 for x in pedidos_detalhe_data if pred(x))
    return n, (n / len(pedidos_detalhe_data) if pedidos_detalhe_data else 0)


# backlog e expirado nunca vão gerar nota, então contá-los como "entrada pendente"
# inflava o número da mesma forma que inflava o de atrasados
_faturaveis = [x for x in pedidos_detalhe_data if x["status"] != STATUS_NAO_FATURA]


def _conta_faturaveis(pred):
    n = sum(1 for x in _faturaveis if pred(x))
    return n, (n / len(_faturaveis) if _faturaveis else 0)


pedidos = {
    "total": len(pedidos_detalhe_data),
    "entregue": _conta_pedidos(lambda x: x["status"] == "ENTREGUE"),
    "pendente": _conta_pedidos(lambda x: x["status"] == STATUS_PENDENTE),
    "atrasado": _conta_pedidos(lambda x: x["status"] == "ATRASADO"),
    "nao_fatura": _conta_pedidos(lambda x: x["status"] == STATUS_NAO_FATURA),
    "entrada_ok": _conta_faturaveis(lambda x: x["entrada"] == "Entrada"),
    "entrada_pendente": _conta_faturaveis(lambda x: x["entrada"] != "Entrada"),
    # produto já na filial com a nota não lançada: é o que exige providência imediata
    "entregue_sem_entrada": _conta_pedidos(
        lambda x: x["status"] == "ENTREGUE" and x["entrada"] != "Entrada"),
    "faturaveis": len(_faturaveis),
}

_por_filial = {}
for x in pedidos_detalhe_data:
    b = _por_filial.setdefault(x["fil"], {"total": 0, "entregue": 0, "pendente": 0, "atrasado": 0, "nao_fatura": 0})
    b["total"] += 1
    if x["status"] == "ENTREGUE":
        b["entregue"] += 1
    elif x["status"] == STATUS_PENDENTE:
        b["pendente"] += 1
    elif x["status"] == "ATRASADO":
        b["atrasado"] += 1
    elif x["status"] == STATUS_NAO_FATURA:
        b["nao_fatura"] += 1
pedidos_filiais = [dict(filial=k, **v) for k, v in _por_filial.items()]
pedidos_filiais.sort(key=lambda f: (f["pendente"] + f["atrasado"]), reverse=True)



def linhas_pedidos_filiais():
    out = []
    for f in pedidos_filiais:
        pct_entrega = f["entregue"] / f["total"] if f["total"] else 0
        out.append(
            "<tr><td>{fil}</td><td class='num'>{tot}</td><td class='num'>{ent}</td>"
            "<td class='num'>{pen}</td><td class='num'>{atr}</td><td class='num'>{nfat}</td>"
            "<td class='num'>{pct}</td></tr>".format(
                fil=f["filial"], tot=f["total"], ent=f["entregue"], pen=f["pendente"],
                atr=f["atrasado"], nfat=f["nao_fatura"], pct=pct(pct_entrega)
            )
        )
    return "\n".join(out)


def _classe_status_excecao(status):
    s = unicodedata.normalize("NFKD", status.upper()).encode("ascii", "ignore").decode("ascii")
    # extravio, recusa e nota sem produto exigem acao; os demais estao em andamento
    return "bad" if any(p in s for p in ("EXTRAVI", "RECUSA", "NAO CHEGARAM")) else "warn"


def linhas_notas_excecoes():
    out = []
    for r in notas_excecoes:
        status = str(r.get("STATUS") or "").strip()
        entrada_feita = str(r.get("REALIZADO ENTRADA NO SISTEMA ?") or "").strip().upper() == "ENTRADA"
        out.append(
            "<tr><td>{fil}</td><td>{nf}</td><td>{desc}</td><td class='num'>{qtd}</td>"
            "<td class='num'>{valor}</td><td>{dia}</td><td><span class='pill {ent_cls}'>{ent}</span></td>"
            "<td><span class='pill {st_cls}'>{status}</span></td></tr>".format(
                fil=r.get("FILIAIS") or "-", nf=r.get("NOTA FISCAL") or "-",
                desc=r.get("DESCRIÇÃO") or "-", qtd=r.get("QUANTIDADE UNITÁRIA") or 0,
                valor=brl(r.get("VALOR TOTAL") or 0), dia=data_str(r.get("DIA DA ENTREGA")),
                ent="Feita" if entrada_feita else "Pendente", ent_cls="ok" if entrada_feita else "warn",
                status=status, st_cls=_classe_status_excecao(status),
            )
        )
    return "\n".join(out)


def _classe_status_nota(status):
    s = status.upper()
    if s == "ATRASADO":
        return "bad"
    if s == "AGUARDANDO":
        return "ok"
    return _classe_status_excecao(status)


def linhas_notas():
    out = []
    for r in notas_pendentes_tabela:
        status = str(r.get("STATUS") or "").strip() or "-"
        atraso = round(r.get("DIAS EM ATRASO") or 0)
        out.append(
            "<tr><td>{fil}</td><td>{nf}</td><td>{desc}</td><td class='num'>{qtd}</td>"
            "<td class='num'>{valor}</td><td>{dia}</td><td class='num'>{atr}</td>"
            "<td><span class='pill {st_cls}'>{status}</span></td></tr>".format(
                fil=r.get("FILIAIS") or "-", nf=r.get("NOTA FISCAL") or "-",
                desc=r.get("DESCRIÇÃO") or "-", qtd=r.get("QUANTIDADE UNITÁRIA") or 0,
                valor=brl(r.get("VALOR TOTAL") or 0), dia=data_str(r.get("DIA DA ENTREGA")),
                atr=atraso if atraso > 0 else "-",
                status=status, st_cls=_classe_status_nota(status),
            )
        )
    return "\n".join(out)


transf_todas_data = []
for r in transf_todas:
    crit = r[11] or ""
    prazo = r[12] or ""
    dias_num = round(r[2] or 0)
    transf_todas_data.append({
        "orig": r[0] or "", "dias": dias_num, "dias_num": dias_num, "dest": r[3] or "",
        "user": r[4] or "", "nf": r[5] or "", "prod": r[7] or "", "desc": r[8] or "",
        "qtd": r[9], "crit": crit, "cls": CRITICIDADE_CLASSE.get(crit, ""), "prazo": prazo,
        "rastreio": (r[13] or "-") if len(r) > 13 else "-",
    })

# ------------------------------------------------------------- RASTREIOS
# a aba é alimentada por atualizar_rastreios.py, que lê os relatórios dos Correios
rastreios_itens = []
for r in load_dicts("TRANSFERÊNCIAS PENDENTES.xlsx", "RASTREIO DOS CORREIOS"):
    codigo = r.get("RASTREIO")
    if not codigo:
        continue
    dt = r.get("DATA")
    tipo = str(r.get("TIPO") or "-").strip()
    destino = r.get("FILIAL") if tipo == "FILIAL" else r.get("DESTINATÁRIO")
    rastreios_itens.append({
        "data": data_str(dt), "data_ts": malote_epoch(dt),
        "rastreio": str(codigo).strip(),
        "tipo": tipo,
        "destino": str(destino or "-").strip() or "-",
        "filial": str(r.get("FILIAL") or "-").strip(),
        "cidade": str(r.get("CIDADE/UF") or "-").strip(),
        "cep": str(r.get("CEP") or "-").strip(),
        "produto": str(r.get("PRODUTO") or "-").strip(),
        "valor": round(r.get("VALOR") or 0, 2),
        "entrega": str(r.get("STATUS ENTREGA") or "-").strip() or "-",
        "local": str(r.get("LOCAL ATUAL") or "-").strip() or "-",
        "rastreado_em": str(r.get("ATUALIZADO EM") or "-").strip() or "-",
        "previsao_dt": r.get("PREVISÃO ENTREGA"),
    })
rastreios_itens.sort(key=lambda x: x["data_ts"], reverse=True)


def classificar_entrega(texto):
    """Agrupa a descrição do evento dos Correios em situações comparáveis."""
    s = unicodedata.normalize("NFKD", str(texto or "").upper()).encode("ascii", "ignore").decode("ascii")
    if not s or s == "-":
        return "Sem rastreamento", ""
    # postagem de balcao (varejo/PAC) nao esta sob o contrato da chave de acesso
    # postada no balcao, fora do contrato: a API restrita ao contrato nao a enxerga
    if "SRO-0" in s or "NAO PERTENCE AO CONTRATO" in s:
        return "Postagem avulsa", ""
    # devolucao: contem "entregue", mas ao remetente, nao ao destino
    if "REMETENTE" in s:
        return "Devolvido ao remetente", "bad"
    if "ENTREGUE" in s:
        return "Entregue", "ok"
    if "SAIU PARA ENTREGA" in s:
        return "Saiu para entrega", "warn"
    if "AGUARDANDO RETIRADA" in s or "DISPONIVEL PARA RETIRADA" in s:
        return "Aguardando retirada", "warn"
    if any(p in s for p in ("DEVOL", "EXTRAVI", "ROUBO", "AVARIA", "NAO ENTREGUE")):
        return "Problema", "bad"
    if "POSTADO" in s:
        return "Postado", ""
    return "Em trânsito", ""


_hoje_rastreio = datetime.now().date()

rastreios_entrega_buckets = {}
rastreios_entrega_classes = {}
for it in rastreios_itens:
    it["entrega_grupo"], it["entrega_cls"] = classificar_entrega(it["entrega"])
    prev = it.pop("previsao_dt", None)
    prev = prev.date() if hasattr(prev, "date") else None
    it["previsao"] = prev.strftime("%d/%m/%Y") if prev else "-"
    it["previsao_ts"] = malote_epoch(prev) if prev else 0
    # atrasado = prazo vencido e ainda sem entrega. "Postagem avulsa" fica de fora
    # porque nao temos rastreamento dela para afirmar nada.
    it["atrasado"] = bool(
        prev and prev < _hoje_rastreio
        and it["entrega_grupo"] not in ("Entregue", "Postagem avulsa", "Sem rastreamento"))
    if it["atrasado"]:
        it["previsao_cls"] = "bad"
    elif prev:
        it["previsao_cls"] = "ok" if it["entrega_grupo"] == "Entregue" else ""
    else:
        it["previsao_cls"] = ""
    rastreios_entrega_buckets[it["entrega_grupo"]] = rastreios_entrega_buckets.get(it["entrega_grupo"], 0) + 1
    rastreios_entrega_classes[it["entrega_grupo"]] = it["entrega_cls"]

# vínculo com as transferências: preenchido conforme você põe o código na coluna RASTREIO
def normalizar_rastreio(codigo):
    """Casa o codigo digitado na transferencia com o dos Correios, com ou sem o BR."""
    s = re.sub(r"[^A-Z0-9]", "", str(codigo or "").upper())
    return s + "BR" if re.fullmatch(r"[A-Z]{2}\d{9}", s) else s


transferencias_por_rastreio = {}
for t in transf_todas_data:
    codigo = normalizar_rastreio(t.get("rastreio"))
    if codigo and codigo != "-":
        transferencias_por_rastreio.setdefault(codigo, []).append(t)
entrega_por_rastreio = {normalizar_rastreio(it["rastreio"]): it for it in rastreios_itens}
for t in transf_todas_data:
    info = entrega_por_rastreio.get(normalizar_rastreio(t.get("rastreio")))
    t["rastreio"] = str(t.get("rastreio") or "-").strip() or "-"
    t["entrega"] = info["entrega_grupo"] if info else "-"
    t["entrega_cls"] = info["entrega_cls"] if info else ""

for it in rastreios_itens:
    ligadas = transferencias_por_rastreio.get(normalizar_rastreio(it["rastreio"]), [])
    it["transferencias"] = len(ligadas)
    it["nfs"] = ", ".join(sorted({str(t["nf"]) for t in ligadas if t["nf"]})) or "-"
    prazos = sorted({t["prazo"] for t in ligadas if t["prazo"]})
    it["situacao"] = ", ".join(prazos) or "-"
    # basta uma transferência atrasada para o envio merecer atenção
    it["situacao_cls"] = "bad" if any("ATRAS" in p.upper() for p in prazos) else ("ok" if prazos else "")

rastreios_por_filial = {}
for it in rastreios_itens:
    if it["tipo"] == "FILIAL":
        rastreios_por_filial[it["filial"]] = rastreios_por_filial.get(it["filial"], 0) + 1
rastreios_filiais_ordenadas = sorted(rastreios_por_filial.items(), key=lambda kv: kv[1], reverse=True)

rastreios_resumo = {
    "total": len(rastreios_itens),
    "filial": sum(1 for i in rastreios_itens if i["tipo"] == "FILIAL"),
    "cliente": sum(1 for i in rastreios_itens if i["tipo"] == "CLIENTE"),
    "frete": round(sum(i["valor"] for i in rastreios_itens), 2),
    "vinculadas": sum(1 for i in rastreios_itens if i["transferencias"]),
    "entregues": sum(1 for i in rastreios_itens if i["entrega_grupo"] == "Entregue"),
    "rastreados": sum(1 for i in rastreios_itens if i["entrega_grupo"] != "Sem rastreamento"),
    "atrasados": sum(1 for i in rastreios_itens if i["atrasado"]),
}
rastreios_json = json.dumps(rastreios_itens, ensure_ascii=False)
transf_todas_json = json.dumps(transf_todas_data, ensure_ascii=False)

chart_data = {
    "pedidosStatus": {
        "labels": ["Entregue", "Pendente de entrega", "Atrasado", "Não será faturado"],
        "values": [pedidos["entregue"][0], pedidos["pendente"][0], pedidos["atrasado"][0],
                   pedidos["nao_fatura"][0]],
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
        "nao_fatura": [f["nao_fatura"] for f in pedidos_filiais],
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
    "amet": {
        "labels": [f["filial"] for f in amet["filiais"]],
        "estoque": [f["estoque"] for f in amet["filiais"]],
        "vendido": [f["vendido"] for f in amet["filiais"]],
    },
    "devia": {
        "labels": [f["filial"] for f in devia["filiais"]],
        "estoque": [f["estoque"] for f in devia["filiais"]],
        "vendido": [f["vendido"] for f in devia["filiais"]],
    },
    "upmaster": {
        "labels": [f["filial"] for f in upmaster["filiais"]],
        "estoque": [f["estoque"] for f in upmaster["filiais"]],
        "vendido": [f["vendido"] for f in upmaster["filiais"]],
    },
    "acessoriosDiversos": {
        "labels": [f[0] for f in acessorios_diversos_filiais],
        "saldo": [f[1] for f in acessorios_diversos_filiais],
    },
    "acessoriosTim": {
        "labels": [f[0] for f in seriais_filiais_ordenadas],
        "saldo": [f[1] for f in seriais_filiais_ordenadas],
    },
    "devolvidos": {
        "labels": [f[0] for f in devolvidos_filiais_ordenadas],
        "saldo": [f[1] for f in devolvidos_filiais_ordenadas],
    },
    "malotesStatusAdm": {
        "labels": [b for b in MALOTE_ORDEM_STATUS_ADM if malotes_status_adm_buckets.get(b)],
        "values": [malotes_status_adm_buckets.get(b, 0) for b in MALOTE_ORDEM_STATUS_ADM if malotes_status_adm_buckets.get(b)],
    },
    "malotesFilial": {
        "labels": [f["filial"] for f in malotes_filiais],
        "na_filial": [f["na_filial"] for f in malotes_filiais],
        "no_adm": [f["no_adm"] for f in malotes_filiais],
    },
    "manutencoesStatus": {
        "labels": list(manutencoes_status_buckets.keys()),
        "values": list(manutencoes_status_buckets.values()),
        "classes": [manutencoes_status_classes[s] for s in manutencoes_status_buckets],
    },
    "rastreiosDestino": {
        "labels": ["Para filiais", "Para clientes"],
        "values": [rastreios_resumo["filial"], rastreios_resumo["cliente"]],
    },
    "rastreiosEntrega": {
        "labels": list(rastreios_entrega_buckets.keys()),
        "values": list(rastreios_entrega_buckets.values()),
        "classes": [rastreios_entrega_classes[k] for k in rastreios_entrega_buckets],
    },
    "rastreiosFilial": {
        "labels": [f[0] for f in rastreios_filiais_ordenadas],
        "values": [f[1] for f in rastreios_filiais_ordenadas],
    },
    "obrasFilial": {
        "labels": [f[0] for f in obras_filiais_ordenadas],
        "values": [f[1] for f in obras_filiais_ordenadas],
    },
    "obrasResponsavel": {
        "labels": [r[0] for r in obras_responsaveis_ordenados],
        "values": [r[1] for r in obras_responsaveis_ordenados],
    },
    "manutencoesFilial": {
        "labels": [f[0] for f in manutencoes_filiais_ordenadas],
        "values": [f[1] for f in manutencoes_filiais_ordenadas],
    },
    "despesasCategoria": {
        "labels": ["Material de Limpeza", "Material de Escritório", "Registro Manual"],
        "values": [despesas_resumo["total_limpeza"], despesas_resumo["total_escritorio"], despesas_resumo["total_manual"]],
    },
    "despesasFilial": {
        "labels": [f[0] for f in despesas_filiais_ordenadas],
        "values": [round(f[1], 2) for f in despesas_filiais_ordenadas],
    },
    "despesasMensal": {
        "labels": despesas_meses_labels,
        "limpeza": [round(despesas_por_mes[m]["Material de Limpeza"], 2) for m in despesas_meses_ordenados],
        "escritorio": [round(despesas_por_mes[m]["Material de Escritório"], 2) for m in despesas_meses_ordenados],
        "manual": [round(despesas_por_mes[m]["Registro Manual"], 2) for m in despesas_meses_ordenados],
    },
    "despesasFilialMensal": {
        "labels": despesas_meses_labels,
        "series": despesas_filial_mensal_series,
    },
    "campanhasArea": {
        "labels": [f"Área {num}" + (f" ({b['coordenador']})" if b["coordenador"] else "") for num, b in campanhas_areas_ordenadas],
        "verba": [round(b["verba"], 2) for _, b in campanhas_areas_ordenadas],
        "gasto": [round(b["gasto"], 2) for _, b in campanhas_areas_ordenadas],
    },
    "campanhasFilial": {
        "labels": [f["filial"] for f in campanhas_filiais],
        "verba": [f["verba"] for f in campanhas_filiais],
        "gasto": [f["gasto"] for f in campanhas_filiais],
    },
}
chart_data_json = json.dumps(chart_data, ensure_ascii=False)


def secao_pelicula(id_, titulo, dados, prefixo):
    produtos_labels = dados["produtos_labels"]
    colunas_produtos = "".join(f"<th class='num'>{p}</th>" for p in produtos_labels)
    linhas = []
    for f in dados["filiais"]:
        giro = pct(f["vendido"] / f["estoque"]) if f["estoque"] else "-"
        cels_produtos = "".join(
            f"<td class='num'>{f['produtos'].get(p, 0)}</td>" for p in produtos_labels
        )
        linhas.append(
            f"<tr><td>{f['filial']}</td>{cels_produtos}"
            f"<td class='num'>{f['estoque']}</td><td class='num'>{f['vendido']}</td>"
            f"<td class='num'>{giro}</td></tr>"
        )
    linhas_html = "\n".join(linhas)
    return f"""
<section id="{id_}">
  <h2>{titulo}</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {dados['mtime']} · Estoque referente a {dados['data_estoque']} · Vendas de {dados['periodo_vendas']}</p>
  <div class="cards">
    <div class="card"><div class="label">Estoque total (peças)</div><div class="value">{dados['estoque_total']}</div></div>
    <div class="card ok"><div class="label">Vendido no período (peças)</div><div class="value">{dados['vendido_total']}</div></div>
    <div class="card"><div class="label">Filiais monitoradas</div><div class="value">{len(dados['filiais'])}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Estoque x Vendido por filial <button class="chart-export-btn" data-chart-export="chart-{prefixo}-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-{prefixo}-filial"></canvas></div></div>
  </div>
  <h3>📋 Estoque por tipo e vendas por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-{prefixo}" placeholder="Filtrar por filial...">
    <button class="table-export-btn" data-export="tbl-{prefixo}">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-{prefixo}" class="sortable">
    <thead><tr><th>Filial</th>{colunas_produtos}<th class="num">Total Estoque</th><th class="num">Vendido (período)</th><th class="num">Giro (vendido/estoque)</th></tr></thead>
    <tbody>
    {linhas_html}
    </tbody>
  </table>
  </div>
</section>
"""


def secao_acessorios(id_, titulo, mtime, resumo, prefixo):
    return f"""
<section id="{id_}">
  <h2>{titulo}</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Itens (linhas de estoque)</div><div class="value">{resumo['itens']}</div></div>
    <div class="card ok"><div class="label">Saldo total (unidades)</div><div class="value">{resumo['saldo_total']}</div></div>
    <div class="card"><div class="label">Valor total em estoque</div><div class="value">{brl(resumo['valor_total'])}</div></div>
    <div class="card"><div class="label">Filiais com estoque</div><div class="value">{resumo['filiais']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Saldo por filial <button class="chart-export-btn" data-chart-export="chart-{prefixo}-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-{prefixo}-filial"></canvas></div></div>
  </div>
  <h3>📋 Itens em estoque por filial</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-{prefixo}" placeholder="Buscar por filial, referência, descrição...">
    <div class="msel" id="msel-{prefixo}-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-{prefixo}-subgrupo"><button type="button" class="msel-btn" data-default="Todos os tipos">Todos os tipos</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-{prefixo}-fabricante"><button type="button" class="msel-btn" data-default="Todos os fabricantes">Todos os fabricantes</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-{prefixo}" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-{prefixo}">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-{prefixo}">
    <thead><tr>
      <th data-col="filial">Filial</th><th data-col="ref">Referência</th><th data-col="desc">Descrição</th><th data-col="subgrupo">Tipo</th>
      <th data-col="fabricante">Fabricante</th><th data-col="saldo" class="num">Saldo</th><th data-col="disponivel" class="num">Disponível</th>
      <th data-col="valor" class="num">Valor Unit.</th><th data-col="valor_total" class="num">Valor Total</th>
    </tr></thead>
    <tbody id="tbody-{prefixo}"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-{prefixo}"></div>
</section>
"""


def secao_seriais_tim(mtime, resumo):
    return f"""
<section id="acessorios-tim">
  <h2>📶 Acessórios Fidelizados TIM nas Filiais</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Peças (com serial)</div><div class="value">{resumo['itens']}</div></div>
    <div class="card ok"><div class="label">Valor total em estoque</div><div class="value">{brl(resumo['valor_total'])}</div></div>
    <div class="card"><div class="label">Filiais com estoque</div><div class="value">{resumo['filiais']}</div></div>
    <div class="card warn"><div class="label">Dias médio em estoque</div><div class="value">{resumo['dias_medio']:.0f}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Peças por filial <button class="chart-export-btn" data-chart-export="chart-acessorios-tim-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-acessorios-tim-filial"></canvas></div></div>
  </div>
  <h3>📋 Peças com número de série</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-acessorios-tim" placeholder="Buscar por filial, serial, descrição...">
    <div class="msel" id="msel-acessorios-tim-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-acessorios-tim-fabricante"><button type="button" class="msel-btn" data-default="Todos os fabricantes">Todos os fabricantes</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-acessorios-tim-dias"><button type="button" class="msel-btn" data-default="Dias em estoque (todos)">Dias em estoque (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-acessorios-tim" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-acessorios-tim">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-acessorios-tim">
    <thead><tr>
      <th data-col="filial">Filial</th><th data-col="serial">Serial</th><th data-col="desc">Descrição</th>
      <th data-col="fabricante">Fabricante</th><th data-col="data_compra">Data Compra</th>
      <th data-col="dias" class="num">Dias em Estoque</th><th data-col="valor" class="num">Valor</th>
    </tr></thead>
    <tbody id="tbody-acessorios-tim"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-acessorios-tim"></div>
</section>
"""


def secao_devolvidos(mtime, resumo):
    return f"""
<section id="devolvidos">
  <h2>♻️ Devolvidos e Defeitos nas Filiais</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Itens (linhas de estoque)</div><div class="value">{resumo['itens']}</div></div>
    <div class="card bad"><div class="label">Saldo total (unidades)</div><div class="value">{resumo['saldo_total']}</div></div>
    <div class="card"><div class="label">Custo total imobilizado</div><div class="value">{brl(resumo['custo_total'])}</div></div>
    <div class="card"><div class="label">Filiais com devolução/defeito</div><div class="value">{resumo['filiais']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Saldo por filial <button class="chart-export-btn" data-chart-export="chart-devolvidos-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-devolvidos-filial"></canvas></div></div>
  </div>
  <h3>📋 Itens devolvidos / com defeito</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-devolvidos" placeholder="Buscar por filial, descrição...">
    <div class="msel" id="msel-devolvidos-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-devolvidos-grupo"><button type="button" class="msel-btn" data-default="Todas as categorias">Todas as categorias</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-devolvidos-fabricante"><button type="button" class="msel-btn" data-default="Todos os fabricantes">Todos os fabricantes</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-devolvidos" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-devolvidos">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-devolvidos">
    <thead><tr>
      <th data-col="filial">Filial</th><th data-col="desc">Descrição</th><th data-col="grupo">Categoria</th>
      <th data-col="fabricante">Fabricante</th><th data-col="saldo" class="num">Saldo</th><th data-col="custo" class="num">Custo Unit.</th>
      <th data-col="custo_total" class="num">Custo Total</th><th data-col="data_mov">Última Movimentação</th>
    </tr></thead>
    <tbody id="tbody-devolvidos"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-devolvidos"></div>
</section>
"""


def secao_despesas(mtime, resumo):
    return f"""
<section id="despesas">
  <h2>💰 Despesas (Material de Limpeza, Escritório e Registros Manuais)</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime} · Período de 01/01/2026 até {datetime.now().strftime('%d/%m/%Y')}</p>
  <div class="cards">
    <div class="card ok"><div class="label">Total Geral Pago</div><div class="value">{brl(resumo['total_geral'])}</div></div>
    <div class="card"><div class="label">Material de Limpeza</div><div class="value">{brl(resumo['total_limpeza'])}</div></div>
    <div class="card"><div class="label">Material de Escritório</div><div class="value">{brl(resumo['total_escritorio'])}</div></div>
    <div class="card"><div class="label">Registro Manual</div><div class="value">{brl(resumo['total_manual'])}</div></div>
    <div class="card"><div class="label">Lançamentos</div><div class="value">{resumo['lancamentos']}</div></div>
    <div class="card"><div class="label">Filiais com despesas</div><div class="value">{resumo['filiais']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Total por categoria <button class="chart-export-btn" data-chart-export="chart-despesas-categoria" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-despesas-categoria"></canvas></div></div>
    <div class="chart-box wide"><h4>Total por filial <button class="chart-export-btn" data-chart-export="chart-despesas-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-despesas-filial"></canvas></div></div>
    <div class="chart-box wide"><h4>Total por mês <button class="chart-export-btn" data-chart-export="chart-despesas-mensal" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-despesas-mensal"></canvas></div></div>
    <div class="chart-box wide"><h4>Comparativo das top 10 filiais por mês <button class="chart-export-btn" data-chart-export="chart-despesas-filial-mensal" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-despesas-filial-mensal"></canvas></div></div>
  </div>
  <h3>📋 Lançamentos de despesas</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-despesas" placeholder="Buscar por filial, fornecedor, documento, histórico...">
    <label class="filtro-data">De <input type="date" id="despesas-data-de"></label>
    <label class="filtro-data">Até <input type="date" id="despesas-data-ate"></label>
    <div class="msel" id="msel-despesas-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-despesas-categoria"><button type="button" class="msel-btn" data-default="Todas as categorias">Todas as categorias</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-despesas" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-despesas">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-despesas">
    <thead><tr>
      <th data-col="filial">Filial</th><th data-col="categoria">Categoria</th><th data-col="fornecedor">Fornecedor / Operação</th>
      <th data-col="documento">Documento</th><th data-col="data">Data</th><th data-col="valor" class="num">Valor</th>
      <th data-col="historico">Histórico</th>
    </tr></thead>
    <tbody id="tbody-despesas"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-despesas"></div>
</section>
"""


def secao_campanhas(mtime):
    """As linhas das duas tabelas, os cards e os graficos sao montados no navegador
    a partir do filtro de datas (ver renderCampanhas no JS), porque a verba e mensal
    e o total depende do intervalo escolhido."""
    return f"""
<section id="campanhas">
  <h2>📣 Campanhas / Valor Aprovado por Área</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime} · <span id="campanhas-periodo-label">—</span></p>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-campanhas" placeholder="Buscar por filial, documento ou histórico...">
    <label class="filtro-data">De <input type="date" id="campanhas-data-de" value="{CAMPANHAS_DE_PADRAO}"></label>
    <label class="filtro-data">Até <input type="date" id="campanhas-data-ate" value="{CAMPANHAS_ATE_PADRAO}"></label>
    <div class="msel" id="msel-campanhas-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-campanhas-area"><button type="button" class="msel-btn" data-default="Todas as áreas">Todas as áreas</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-campanhas-tipo"><button type="button" class="msel-btn" data-default="Campanha e verba (todos)">Campanha e verba (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-campanhas" type="button">Limpar filtros</button>
  </div>
  <div class="cards">
    <div class="card"><div class="label" id="card-camp-verba-label">Aprovado no período</div><div class="value" id="card-camp-verba">—</div></div>
    <div class="card warn"><div class="label">Total Utilizado</div><div class="value" id="card-camp-gasto">—</div></div>
    <div class="card ok"><div class="label">Saldo Disponível</div><div class="value" id="card-camp-saldo">—</div></div>
    <div class="card"><div class="label">% Utilizado</div><div class="value" id="card-camp-pct">—</div></div>
    <div class="card"><div class="label">Lançamentos</div><div class="value" id="card-camp-lanc">—</div></div>
    <div class="card"><div class="label">Filiais que já utilizaram</div><div class="value" id="card-camp-filiais">—</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Utilizado: campanha x verba de área <button class="chart-export-btn" data-chart-export="chart-campanhas-tipo" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-campanhas-tipo"></canvas></div></div>
    <div class="chart-box wide"><h4>Aprovado x Utilizado por área <button class="chart-export-btn" data-chart-export="chart-campanhas-area" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-campanhas-area"></canvas></div></div>
    <div class="chart-box wide"><h4>Aprovado x Utilizado por filial <button class="chart-export-btn" data-chart-export="chart-campanhas-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-campanhas-filial"></canvas></div></div>
  </div>
  <h3>📋 Aprovado por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-campanhas-filiais" placeholder="Filtrar por filial ou área...">
    <button class="table-export-btn" data-export="tbl-campanhas-filiais">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-campanhas-filiais" class="sortable">
    <thead><tr><th>Filial</th><th>Área</th><th class="num">Aprovado no Período</th><th class="num">Utilizado</th><th class="num">Saldo</th><th class="num">% Utilizado</th></tr></thead>
    <tbody id="tbody-campanhas-filiais"></tbody>
  </table>
  </div>

  <h3 style="margin-top:32px;">📋 Lançamentos de campanhas</h3>
  <div class="table-toolbar">
    <button class="table-export-btn" data-export="tbl-campanhas-log">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-campanhas-log">
    <thead><tr>
      <th data-col="filial">Filial</th><th data-col="area">Área</th><th data-col="tipo">Tipo</th><th data-col="data">Data</th>
      <th data-col="documento">Documento</th><th data-col="valor" class="num">Valor</th><th data-col="historico">Histórico</th>
    </tr></thead>
    <tbody id="tbody-campanhas-log"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-campanhas-log"></div>
</section>
"""


def linhas_malotes():
    out = []
    for f in malotes_filiais:
        out.append(
            "<tr><td>{fil}</td><td class='num'>{na_filial}</td><td class='num'>{no_adm}</td>"
            "<td class='num'>{cap}</td><td><span class='pill {s_cls}'>{s}</span></td><td>{acao}</td></tr>".format(
                fil=f["filial"], na_filial=f["na_filial"], no_adm=f["no_adm"], cap=f["capacidade"],
                s=f["status_adm"], s_cls=f["status_cls"], acao=f["acao"]
            )
        )
    return "\n".join(out)


def linhas_manutencoes():
    out = []
    for m in manutencoes_itens:
        out.append(
            "<tr><td>{chamado}</td><td>{fil}</td><td>{sol}</td>"
            "<td><span class='pill {s_cls}'>{s}</span></td><td class='num'>{dias}</td><td>{obs}</td></tr>".format(
                chamado=m["chamado"], fil=m["filial"], sol=m["solicitado"],
                s=m["status"], s_cls=m["status_cls"], dias=m["dias"], obs=m["obs"]
            )
        )
    return "\n".join(out)


def linhas_obras():
    out = []
    for o in obras_itens:
        out.append(
            "<tr><td>{fil}</td><td>{obra}</td><td>{resp}</td><td>{obs}</td><td>{prazo}</td>"
            "<td><span class='pill {s_cls}'>{s}</span></td></tr>".format(
                fil=o["filial"], obra=o["obra"], resp=o["responsavel"], obs=o["obs"],
                prazo=o["prazo"], s=o["status"], s_cls=o["status_cls"]
            )
        )
    return "\n".join(out)


def secao_rastreios(mtime, resumo):
    return f"""
<section id="rastreios">
  <h2>📍 Rastreios dos Correios</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime} · CEP que não é de filial foi envio para cliente</p>
  <div class="cards">
    <div class="card"><div class="label">Postagens</div><div class="value">{resumo['total']}</div></div>
    <div class="card ok"><div class="label">Para filiais</div><div class="value">{resumo['filial']}</div></div>
    <div class="card warn"><div class="label">Para clientes</div><div class="value">{resumo['cliente']}</div></div>
    <div class="card"><div class="label">Frete total</div><div class="value">{brl(resumo['frete'])}</div></div>
    <div class="card"><div class="label">Ligadas a transferência</div><div class="value">{resumo['vinculadas']} de {resumo['total']}</div></div>
    <div class="card ok"><div class="label">Entregues</div><div class="value">{resumo['entregues']}</div><div class="sub">de {resumo['rastreados']} com rastreamento</div></div>
    <div class="card bad"><div class="label">Entrega atrasada</div><div class="value">{resumo['atrasados']}</div><div class="sub">passou da previsão dos Correios</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status das entregas <button class="chart-export-btn" data-chart-export="chart-rastreios-entrega" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-rastreios-entrega"></canvas></div></div>
    <div class="chart-box"><h4>Destino das postagens <button class="chart-export-btn" data-chart-export="chart-rastreios-destino" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-rastreios-destino"></canvas></div></div>
    <div class="chart-box wide"><h4>Postagens por filial <button class="chart-export-btn" data-chart-export="chart-rastreios-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-rastreios-filial"></canvas></div></div>
  </div>
  <h3>📋 Postagens</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-rastreios" placeholder="Buscar por código, filial, destinatário ou cidade...">
    <label class="filtro-data">De <input type="date" id="rastreios-data-de"></label>
    <label class="filtro-data">Até <input type="date" id="rastreios-data-ate"></label>
    <div class="msel" id="msel-rastreios-tipo"><button type="button" class="msel-btn" data-default="Filial e cliente (todos)">Filial e cliente (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-rastreios-destino"><button type="button" class="msel-btn" data-default="Todos os destinos">Todos os destinos</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-rastreios-entrega"><button type="button" class="msel-btn" data-default="Status da entrega (todos)">Status da entrega (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-rastreios" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-rastreios">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-rastreios">
    <thead><tr>
      <th data-col="data">Data</th><th data-col="rastreio">Código</th><th data-col="tipo">Tipo</th>
      <th data-col="destino">Destino</th><th data-col="cidade">Cidade/UF</th><th data-col="produto">Serviço</th>
      <th data-col="valor" class="num">Frete</th><th data-col="entrega">Status da entrega</th>
      <th data-col="local">Onde está</th><th data-col="previsao">Previsão</th>
      <th data-col="nfs">NF da transferência</th>
      <th data-col="situacao">Situação da transferência</th>
    </tr></thead>
    <tbody id="tbody-rastreios"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-rastreios"></div>
</section>
"""


def secao_obras(mtime, resumo):
    return f"""
<section id="obras">
  <h2>🏗️ Obras Pendentes</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime} · Prazo: {resumo['prazo']}</p>
  <div class="cards">
    <div class="card warn"><div class="label">Total de Obras</div><div class="value">{resumo['total']}</div></div>
    <div class="card"><div class="label">Filiais Envolvidas</div><div class="value">{resumo['filiais']}</div></div>
    <div class="card bad"><div class="label">Responsável a Definir</div><div class="value">{resumo['sem_responsavel']}</div></div>
    <div class="card"><div class="label">Prazo</div><div class="value">{resumo['prazo']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Obras por filial <button class="chart-export-btn" data-chart-export="chart-obras-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-obras-filial"></canvas></div></div>
    <div class="chart-box"><h4>Obras por responsável <button class="chart-export-btn" data-chart-export="chart-obras-responsavel" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-obras-responsavel"></canvas></div></div>
  </div>
  <h3>📋 Obras pendentes por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-obras" placeholder="Filtrar por filial, obra ou responsável...">
    <button class="table-export-btn" data-export="tbl-obras">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-obras" class="sortable">
    <thead><tr><th>Filial</th><th>Obra / Serviço</th><th>Responsável</th><th>Observação</th><th>Prazo</th><th>Status</th></tr></thead>
    <tbody>
    {linhas_obras()}
    </tbody>
  </table>
  </div>
</section>
"""


def secao_manutencoes(mtime, resumo):
    return f"""
<section id="manutencoes">
  <h2>🔧 Solicitações de Manutenções</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de Chamados</div><div class="value">{resumo['total']}</div></div>
    <div class="card warn"><div class="label">Pendentes</div><div class="value">{resumo['pendentes']}</div></div>
    <div class="card ok"><div class="label">Concluídas</div><div class="value">{resumo['concluidas']}</div></div>
    <div class="card"><div class="label">Filiais com Chamados</div><div class="value">{resumo['filiais']}</div></div>
    <div class="card"><div class="label">Dias Médio sem Conclusão</div><div class="value">{resumo['dias_medio']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status dos chamados <button class="chart-export-btn" data-chart-export="chart-manutencoes-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-manutencoes-status"></canvas></div></div>
    <div class="chart-box wide"><h4>Chamados por filial <button class="chart-export-btn" data-chart-export="chart-manutencoes-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-manutencoes-filial"></canvas></div></div>
  </div>
  <h3>📋 Chamados de manutenção</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-manutencoes" placeholder="Filtrar por filial, chamado ou status...">
    <button class="table-export-btn" data-export="tbl-manutencoes">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-manutencoes" class="sortable">
    <thead><tr><th>Nº Chamado</th><th>Filial</th><th>Manutenção Solicitada</th><th>Status</th><th class="num">Dias sem Conclusão</th><th>Observação</th></tr></thead>
    <tbody>
    {linhas_manutencoes()}
    </tbody>
  </table>
  </div>
</section>
"""


def secao_malotes(mtime, resumo_filiais, resumo_log):
    return f"""
<section id="malotes">
  <h2>👜 Controle de Malotes</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de Malotes (Parque)</div><div class="value">{resumo_filiais['total_parque']}</div></div>
    <div class="card ok"><div class="label">No ADM</div><div class="value">{resumo_filiais['no_adm']}</div></div>
    <div class="card ok"><div class="label">Nas Filiais</div><div class="value">{resumo_filiais['nas_filiais']}</div></div>
    <div class="card bad"><div class="label">Filiais aguardando coleta</div><div class="value">{resumo_filiais['sem_malote_adm']}</div><div class="sub">todos os malotes estão na loja</div></div>
    <div class="card warn"><div class="label">Solicitações Pendentes</div><div class="value">{resumo_log['pendentes']}</div></div>
    <div class="card"><div class="label">Tempo Médio até Postagem</div><div class="value">{resumo_log['tempo_medio_h']}h</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status ADM por filial <button class="chart-export-btn" data-chart-export="chart-malotes-status-adm" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-malotes-status-adm"></canvas></div></div>
    <div class="chart-box wide"><h4>Malotes por filial (na filial x no ADM) <button class="chart-export-btn" data-chart-export="chart-malotes-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-malotes-filial"></canvas></div></div>
  </div>
  <h3>📋 Malotes por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-malotes-filial" placeholder="Filtrar por filial ou ação...">
    <button class="table-export-btn" data-export="tbl-malotes-filial">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-malotes-filial" class="sortable">
    <thead><tr><th>Filial</th><th class="num">Malotes na Filial</th><th class="num">Malotes no ADM</th><th class="num">Capacidade</th><th>Status ADM</th><th>Ação</th></tr></thead>
    <tbody>
    {linhas_malotes()}
    </tbody>
  </table>
  </div>

  <h3 style="margin-top:32px;">📋 Solicitações de postagem</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-malotes-log" placeholder="Buscar por filial, solicitante, ID...">
    <div class="msel" id="msel-malotes-log-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-malotes-log-status"><button type="button" class="msel-btn" data-default="Todos os status">Todos os status</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-malotes-log-conteudo"><button type="button" class="msel-btn" data-default="Todo conteúdo">Todo conteúdo</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-malotes-log" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-malotes-log">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-malotes-log">
    <thead><tr>
      <th data-col="id" class="num">ID</th><th data-col="solicitante">Solicitante</th><th data-col="filial">Filial Destino</th>
      <th data-col="qtd" class="num">Qtde</th><th data-col="conteudo">Conteúdo</th><th data-col="data_sol">Data Solicitação</th>
      <th data-col="status">Status</th><th data-col="data_post">Data Postagem</th><th data-col="quem_postou">Quem Postou</th>
    </tr></thead>
    <tbody id="tbody-malotes-log"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-malotes-log"></div>
</section>
"""


html = rf"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Controle Operacional — Rocha Telecom</title>
<link rel="icon" type="image/svg+xml" href="{FAVICON_HREF}">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script>
(function() {{
  try {{
    var t = localStorage.getItem('bi-tema');
    if (t === 'light' || t === 'dark') document.documentElement.setAttribute('data-tema', t);
  }} catch (e) {{}}
}})();
</script>
<style>
  :root {{
    --bg: #0f172a; --card: #1e293b; --card2: #16213a; --text: #e2e8f0; --muted: #94a3b8;
    --ok: #22c55e; --warn: #f59e0b; --bad: #ef4444; --accent: #38bdf8; --border: #334155;
    --hover: #22314f; --th-bg: #0b1120; --sombra: rgba(0,0,0,.5); --esquema: dark;
    --chart-texto: #94a3b8; --chart-grade: #334155; --chart-rotulo: #e2e8f0; --chart-ponto: #0f172a;
  }}
  /* o tema é aplicado no <head>, antes da primeira pintura, para não piscar escuro */
  :root[data-tema="light"] {{
    --bg: #f1f5f9; --card: #ffffff; --card2: #f8fafc; --text: #0f172a; --muted: #64748b;
    --ok: #15803d; --warn: #b45309; --bad: #b91c1c; --accent: #0369a1; --border: #cbd5e1;
    --hover: #e2e8f0; --th-bg: #eef2f7; --sombra: rgba(15,23,42,.18); --esquema: light;
    --chart-texto: #475569; --chart-grade: #cbd5e1; --chart-rotulo: #0f172a; --chart-ponto: #ffffff;
  }}
  * {{ box-sizing: border-box; }}
  html {{ scroll-behavior: smooth; }}
  body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); }}
  .layout {{ display: flex; align-items: flex-start; }}
  .sidebar {{ width: 220px; flex-shrink: 0; position: sticky; top: 0; height: 100vh; overflow-y: auto; background: var(--card); border-right: 1px solid var(--border); padding: 20px 0; }}
  .tema-btn {{ display: flex; align-items: center; gap: 8px; width: calc(100% - 40px); margin: 0 20px 10px; padding: 8px 10px; background: var(--card2); border: 1px solid var(--border); border-radius: 6px; color: var(--muted); cursor: pointer; font-size: 13px; font-family: inherit; }}
  .tema-btn:hover {{ color: var(--text); border-color: var(--accent); }}
  .sidebar .brand {{ padding: 0 20px 16px; font-size: 14px; font-weight: 700; color: var(--text); border-bottom: 1px solid var(--border); margin-bottom: 8px; }}
  .sidebar a {{ display: flex; align-items: flex-start; gap: 8px; padding: 12px 20px; color: var(--muted); text-decoration: none; font-size: 14px; line-height: 1.35; border-left: 3px solid transparent; }}
  .sidebar a .nav-icon {{ flex-shrink: 0; }}
  .sidebar a .nav-label {{ flex: 1; }}
  .sidebar a:hover {{ background: var(--hover); color: var(--text); }}
  .sidebar a.active {{ color: var(--accent); border-left-color: var(--accent); background: var(--card2); font-weight: 600; }}
  .content {{ flex: 1; min-width: 0; }}
  header {{ padding: 24px 32px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; flex-wrap: wrap; }}
  header h1 {{ margin: 0 0 4px; font-size: 22px; }}
  header p {{ margin: 0; color: var(--muted); font-size: 14px; }}
  .badge-atualizacao {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 10px 16px; text-align: right; white-space: nowrap; }}
  .badge-atualizacao .data {{ font-size: 15px; font-weight: 700; color: var(--accent); }}
  .badge-atualizacao .gerado {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}
  main {{ padding: 24px 32px 64px; max-width: 1200px; margin: 0 auto; }}
  section {{ margin-bottom: 40px; scroll-margin-top: 16px; }}
  section > h2 {{ font-size: 18px; border-left: 4px solid var(--accent); padding-left: 10px; margin-bottom: 4px; }}
  .secao-mtime {{ color: var(--muted); font-size: 13px; margin: 0 0 16px 14px; }}
  @media (max-width: 860px) {{
    .layout {{ display: block; }}
    .sidebar {{ position: sticky; width: 100%; height: auto; display: flex; overflow-x: auto; border-right: none; border-bottom: 1px solid var(--border); padding: 0; z-index: 10; }}
    .sidebar .brand {{ display: none; }}
    .tema-btn {{ width: auto; margin: 0 10px; align-self: center; flex-shrink: 0; }}
    .sidebar a {{ white-space: nowrap; border-left: none; border-bottom: 3px solid transparent; padding: 14px 16px; }}
    .sidebar a.active {{ border-left: none; border-bottom-color: var(--accent); }}
    section {{ scroll-margin-top: 56px; }}
  }}
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
  th {{ background: var(--th-bg); color: var(--muted); font-size: 11px; text-transform: uppercase; position: sticky; top: 0; }}
  table.sortable th {{ cursor: pointer; user-select: none; }}
  table.sortable th:hover {{ color: var(--text); }}
  table.sortable th .sort-ind {{ color: var(--accent); }}
  td.num, th.num {{ text-align: center; font-variant-numeric: tabular-nums; }}
  .pill {{ display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 600; }}
  .pill.ok {{ background: rgba(34,197,94,.15); color: var(--ok); }}
  .pill.warn {{ background: rgba(245,158,11,.15); color: var(--warn); }}
  .pill.bad {{ background: rgba(239,68,68,.15); color: var(--bad); }}
  .pager {{ display: flex; align-items: center; gap: 10px; margin-top: 10px; font-size: 13px; color: var(--muted); }}
  .pager button {{ background: var(--card); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 12px; cursor: pointer; }}
  .pager button:disabled {{ opacity: .4; cursor: default; }}
  .chart-box h4 {{ display: flex; align-items: center; justify-content: space-between; gap: 8px; }}
  .chart-export-btn {{ background: none; border: 1px solid var(--border); color: var(--muted); border-radius: 6px; padding: 2px 7px; font-size: 12px; cursor: pointer; line-height: 1.6; }}
  .chart-export-btn:hover {{ color: var(--text); border-color: var(--accent); }}
  .table-export-btn {{ background: var(--ok); border: 1px solid var(--ok); color: #06210f; border-radius: 6px; padding: 8px 14px; font-weight: 600; cursor: pointer; font-size: 13px; }}
  .table-export-btn:hover {{ filter: brightness(1.1); }}
  .table-toolbar {{ display: flex; align-items: center; gap: 10px; margin-bottom: 10px; flex-wrap: wrap; }}
  tr:hover td {{ background: var(--hover); }}
  .table-wrap {{ max-height: 480px; overflow: auto; border: 1px solid var(--border); border-radius: 10px; }}
  .charts {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; margin-bottom: 20px; }}
  .chart-box {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; position: relative; height: 320px; }}
  .chart-box h4 {{ margin: 0 0 12px; font-size: 13px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }}
  .chart-box .canvas-wrap {{ position: relative; height: calc(100% - 28px); }}
  .chart-box.wide {{ grid-column: 1 / -1; height: 720px; }}
  input.filtro {{ width: 100%; max-width: 320px; padding: 8px 10px; margin-bottom: 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); }}
  .filtros-pedidos {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }}
  .filtros-pedidos input.filtro {{ margin-bottom: 0; flex: 1 1 240px; }}
  .filtros-pedidos .filtro-data {{ display: flex; align-items: center; gap: 6px; color: var(--muted); font-size: 13px; }}
  .filtros-pedidos .filtro-data input[type="date"] {{ padding: 7px 8px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); color-scheme: var(--esquema); }}
  .filtros-pedidos select {{ padding: 8px 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); flex: 1 1 170px; }}
  .filtros-pedidos button {{ padding: 8px 14px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--muted); cursor: pointer; }}
  .filtros-pedidos button:hover {{ color: var(--text); border-color: var(--accent); }}
  .msel {{ position: relative; flex: 1 1 170px; }}
  .msel-btn {{ width: 100%; text-align: left; padding: 8px 26px 8px 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); cursor: pointer; position: relative; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .msel-btn:after {{ content: '▾'; position: absolute; right: 10px; top: 50%; transform: translateY(-50%); color: var(--muted); }}
  .msel.open .msel-btn {{ border-color: var(--accent); }}
  .msel-panel {{ display: none; position: absolute; top: calc(100% + 4px); left: 0; min-width: 240px; max-height: 280px; overflow-y: auto; background: var(--card2); border: 1px solid var(--border); border-radius: 8px; padding: 8px; z-index: 30; box-shadow: 0 8px 24px var(--sombra); }}
  .msel.open .msel-panel {{ display: block; }}
  .msel-actions {{ display: flex; gap: 8px; margin-bottom: 6px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }}
  .msel-actions button {{ font-size: 11px; background: none; border: 1px solid var(--border); color: var(--muted); border-radius: 4px; padding: 3px 8px; cursor: pointer; }}
  .msel-actions button:hover {{ color: var(--text); }}
  .msel-options label {{ display: flex; align-items: center; gap: 7px; padding: 5px 4px; font-size: 13px; cursor: pointer; border-radius: 4px; }}
  .msel-options label:hover {{ background: var(--hover); }}
  footer {{ text-align: center; color: var(--muted); font-size: 12px; padding: 24px; }}
</style>
</head>
<body>
<div class="layout">
<nav class="sidebar">
  <div class="brand">📊 Controle Operacional</div>
  <button type="button" id="tema-toggle" class="tema-btn" title="Alternar entre tema claro e escuro"><span id="tema-icone">🌙</span><span id="tema-texto">Modo escuro</span></button>
  <a href="#pedidos" class="nav-link"><span class="nav-icon">📦</span><span class="nav-label">Pedidos</span></a>
  <a href="#notas" class="nav-link"><span class="nav-icon">🧾</span><span class="nav-label">Notas Fiscais</span></a>
  <a href="#transferencias" class="nav-link"><span class="nav-icon">🔄</span><span class="nav-label">Transferências</span></a>
  <a href="#rastreios" class="nav-link"><span class="nav-icon">📍</span><span class="nav-label">Rastreios</span></a>
  <a href="#malotes" class="nav-link"><span class="nav-icon">👜</span><span class="nav-label">Controle de Malotes</span></a>
  <a href="#amet" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas AMET</span></a>
  <a href="#devia" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas DEVIA</span></a>
  <a href="#upmaster" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas UPMASTER</span></a>
  <a href="#manutencoes" class="nav-link"><span class="nav-icon">🔧</span><span class="nav-label">Manutenções</span></a>
  <a href="#obras" class="nav-link"><span class="nav-icon">🏗️</span><span class="nav-label">Obras</span></a>
  <a href="#despesas" class="nav-link"><span class="nav-icon">💰</span><span class="nav-label">Despesas</span></a>
  <a href="#campanhas" class="nav-link"><span class="nav-icon">📣</span><span class="nav-label">Campanhas</span></a>
  <a href="#acessorios" class="nav-link"><span class="nav-icon">🎧</span><span class="nav-label">Acessórios</span></a>
  <a href="#acessorios-tim" class="nav-link"><span class="nav-icon">📶</span><span class="nav-label">Fidelizados TIM</span></a>
  <a href="#devolvidos" class="nav-link"><span class="nav-icon">♻️</span><span class="nav-label">Devolvidos e Defeitos</span></a>
</nav>
<div class="content">
<header>
  <div>
    <h1>📊 Controle Operacional</h1>
    <p>Giovanni Brochini · Rocha Telecom</p>
  </div>
  <div class="badge-atualizacao">
    <div class="data">🕒 Atualizado em {data_atualizacao}</div>
    <div class="gerado">Página gerada em {gerado_em}</div>
  </div>
</header>
<main>

<section id="pedidos">
  <h2>📦 Pedidos GN — Status Geral</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {pedidos_mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de pedidos</div><div class="value">{pedidos['total']}</div></div>
    <div class="card ok"><div class="label">Entregue</div><div class="value">{pedidos['entregue'][0]}</div><div class="sub">{pct(pedidos['entregue'][1])}</div></div>
    <div class="card warn"><div class="label">Pendente de entrega</div><div class="value">{pedidos['pendente'][0]}</div><div class="sub">{pct(pedidos['pendente'][1])}</div></div>
    <div class="card bad"><div class="label">Atrasado</div><div class="value">{pedidos['atrasado'][0]}</div><div class="sub">{pct(pedidos['atrasado'][1])}</div></div>
    <div class="card"><div class="label">Não será faturado</div><div class="value">{pedidos['nao_fatura'][0]}</div><div class="sub">{pct(pedidos['nao_fatura'][1])} · backlog/expirado</div></div>
    <div class="card ok"><div class="label">Entrada no sistema realizada</div><div class="value">{pedidos['entrada_ok'][0]}</div><div class="sub">{pct(pedidos['entrada_ok'][1])} dos {pedidos['faturaveis']} faturáveis</div></div>
    <div class="card warn"><div class="label">Entrada pendente no sistema</div><div class="value">{pedidos['entrada_pendente'][0]}</div><div class="sub">{pct(pedidos['entrada_pendente'][1])} dos {pedidos['faturaveis']} faturáveis</div></div>
    <div class="card bad"><div class="label">Entregue sem entrada lançada</div><div class="value">{pedidos['entregue_sem_entrada'][0]}</div><div class="sub">produto na filial, nota não lançada</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status geral dos pedidos <button class="chart-export-btn" data-chart-export="chart-pedidos-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-status"></canvas></div></div>
    <div class="chart-box"><h4>Entrada no sistema <button class="chart-export-btn" data-chart-export="chart-pedidos-entrada" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-entrada"></canvas></div></div>
    <div class="chart-box wide"><h4>Pedidos por filial (entregue / pendente / atrasado / não faturado) <button class="chart-export-btn" data-chart-export="chart-pedidos-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-filial"></canvas></div></div>
  </div>
  <h3>📋 Pedidos por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-pedidos-filial" placeholder="Filtrar por filial...">
    <button class="table-export-btn" data-export="tbl-pedidos-filial">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-pedidos-filial" class="sortable">
    <thead><tr><th>Filial</th><th class="num">Total</th><th class="num">Entregue</th><th class="num">Pendente</th><th class="num">Atrasado</th><th class="num">Não Faturado</th><th class="num">% Entregue</th></tr></thead>
    <tbody>
    {linhas_pedidos_filiais()}
    </tbody>
  </table>
  </div>

  <h3>📋 Tabela de pedidos realizados no GN</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-pedidos-detalhe" placeholder="Buscar por filial, pedido, referência ou produto...">
    <div class="msel" id="msel-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-status"><button type="button" class="msel-btn" data-default="Todos os status">Todos os status</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-stprod"><button type="button" class="msel-btn" data-default="Status do produto (todos)">Status do produto (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-entrada"><button type="button" class="msel-btn" data-default="Entrada no sistema (todos)">Entrada no sistema (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-dias"><button type="button" class="msel-btn" data-default="Dias em aberto (todos)">Dias em aberto (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-filtros-pedidos" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-pedidos-detalhe">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-pedidos-detalhe">
    <thead><tr>
      <th data-col="dreal">Data Realização</th><th data-col="dproc">Data Processamento</th><th data-col="dprev">Data Prevista Entrega</th>
      <th data-col="fil">Filial Destino</th><th data-col="ped" class="num">Nº Pedido GN</th><th data-col="qtd" class="num">Qtde</th>
      <th data-col="ref">Referência</th><th data-col="desc">Descrição do Produto</th>
      <th data-col="stprod">Status do Produto</th><th data-col="dias" class="num">Dias em Aberto</th><th data-col="entrada">Entrada no Sistema</th><th data-col="status">Status</th>
    </tr></thead>
    <tbody id="tbody-pedidos-detalhe"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-pedidos-detalhe"></div>
</section>

<section id="notas">
  <h2>🧾 Notas Fiscais Pendentes de Entrada</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {notas_mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de notas</div><div class="value">{notas_resumo.get('Total de notas (linhas)', 0)}</div><div class="sub">{notas_linhas_resumo['total']} linhas de produto</div></div>
    <div class="card warn"><div class="label">Pendentes</div><div class="value">{notas_resumo.get('Pendentes (NÃO)', 0)}</div><div class="sub">{notas_linhas_resumo['pendentes']} linhas · {notas_linhas_resumo['excecoes']} em exceção</div></div>
    <div class="card bad"><div class="label">Atrasadas</div><div class="value">{notas_resumo.get('Atrasadas', 0)}</div></div>
    <div class="card ok"><div class="label">Aguardando (no prazo)</div><div class="value">{notas_resumo.get('Aguardando (no prazo)', 0)}</div></div>
    <div class="card"><div class="label">Valor total pendente</div><div class="value">{brl(notas_resumo.get('Valor total pendente', 0))}</div></div>
    <div class="card bad"><div class="label">Maior atraso</div><div class="value">{int(notas_resumo.get('Maior atraso (dias)', 0))} dias</div></div>
    <div class="card"><div class="label">Média de atraso</div><div class="value">{notas_resumo.get('Média de atraso (dias)', 0):.1f} dias</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Notas pendentes: atrasadas x no prazo <button class="chart-export-btn" data-chart-export="chart-notas-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-notas-status"></canvas></div></div>
  </div>
  <h3>📋 Todas as notas pendentes de entrada</h3>
  <p class="secao-mtime">{notas_resumo.get('Pendentes (NÃO)', 0)} notas · {notas_linhas_resumo['pendentes']} linhas de produto · atrasadas primeiro</p>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-notas" placeholder="Filtrar por filial, NF ou produto...">
    <button class="table-export-btn" data-export="tbl-notas">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-notas" class="sortable">
    <thead><tr><th>Filial</th><th>Nota Fiscal</th><th>Descrição</th><th class="num">Qtde</th><th class="num">Valor</th><th>Dia da Entrega</th><th class="num">Dias em Atraso</th><th>Status</th></tr></thead>
    <tbody>
    {linhas_notas()}
    </tbody>
  </table>
  </div>

  <h3 style="margin-top:32px;">🔎 Produtos com status diferente</h3>
  <p class="secao-mtime">{notas_excecoes_resumo['produtos']} produtos em {notas_excecoes_resumo['notas']} notas · {brl(notas_excecoes_resumo['valor'])} · {' · '.join(f"{s}: {n}" for s, n in notas_excecoes_resumo['por_status'])}</p>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-notas-excecoes" placeholder="Filtrar por filial, NF, produto ou status...">
    <button class="table-export-btn" data-export="tbl-notas-excecoes">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-notas-excecoes" class="sortable">
    <thead><tr><th>Filial</th><th>Nota Fiscal</th><th>Descrição</th><th class="num">Qtde</th><th class="num">Valor</th><th>Dia da Entrega</th><th>Entrada</th><th>Status</th></tr></thead>
    <tbody>
    {linhas_notas_excecoes()}
    </tbody>
  </table>
  </div>
</section>

<section id="transferencias">
  <h2>🔄 Transferências Pendentes</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {transf_mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Notas fiscais</div><div class="value">{transf_resumo.get('Total de linhas', 0)}</div><div class="sub">{transf_linhas_total} linhas · {transf_resumo.get('Qtde total (unidades)', 0)} unidades</div></div>
    <div class="card ok"><div class="label">No prazo</div><div class="value">{transf_resumo.get('No prazo', 0)}</div></div>
    <div class="card bad"><div class="label">Atrasados</div><div class="value">{transf_resumo.get('Atrasados', 0)}</div></div>
    <div class="card warn"><div class="label">Atenção</div><div class="value">{transf_resumo.get('ATENÇÃO', 0)}</div></div>
    <div class="card bad"><div class="label">Crítico</div><div class="value">{transf_resumo.get('CRÍTICO', 0)}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Transferências por criticidade <button class="chart-export-btn" data-chart-export="chart-transf-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-transf-status"></canvas></div></div>
    <div class="chart-box"><h4>Tempo fora do estoque <button class="chart-export-btn" data-chart-export="chart-transf-prazo" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-transf-prazo"></canvas></div></div>
  </div>
  <h3>📋 Todas as transferências pendentes</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-transf" placeholder="Buscar por filial, NF ou produto...">
    <div class="msel" id="msel-torig"><button type="button" class="msel-btn" data-default="Todas as filiais de origem">Todas as filiais de origem</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-tdest"><button type="button" class="msel-btn" data-default="Todas as filiais de destino">Todas as filiais de destino</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-tcrit"><button type="button" class="msel-btn" data-default="Criticidade (todas)">Criticidade (todas)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-tprazo"><button type="button" class="msel-btn" data-default="Status de prazo (todos)">Status de prazo (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-tdias"><button type="button" class="msel-btn" data-default="Dias fora do estoque (todos)">Dias fora do estoque (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-filtros-transf" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-transf">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-transf">
    <thead><tr>
      <th data-col="orig">Filial Origem</th><th data-col="dias" class="num">Dias fora do estoque</th><th data-col="dest">Filial Destino</th>
      <th data-col="user">Usuário Solicitante</th><th data-col="nf">NF</th><th data-col="prod">Produto</th>
      <th data-col="desc">Descrição</th><th data-col="qtd" class="num">Qtde</th><th data-col="crit">Criticidade</th>
      <th data-col="rastreio">Rastreio</th><th data-col="entrega">Status da entrega</th>
    </tr></thead>
    <tbody id="tbody-transf"></tbody>
  </table>
  </div>
  <div class="pager" id="contador-transf"></div>
</section>

{secao_rastreios(transf_mtime, rastreios_resumo)}
{secao_malotes(malotes_mtime, malotes_resumo, malotes_log_resumo)}

{secao_pelicula("amet", "🛡️ Películas AMET por Filial", amet, "amet")}

{secao_pelicula("devia", "🛡️ Películas DEVIA por Filial", devia, "devia")}

{secao_pelicula("upmaster", "🛡️ Películas UPMASTER por Filial", upmaster, "upmaster")}

{secao_manutencoes(manutencoes_mtime, manutencoes_resumo)}
{secao_obras(manutencoes_mtime, obras_resumo)}

{secao_despesas(despesas_mtime, despesas_resumo)}

{secao_campanhas(campanhas_mtime)}

{secao_acessorios("acessorios", "🎧 Acessórios nas Filiais", acessorios_mtime, acessorios_diversos_resumo, "acessorios")}
{secao_seriais_tim(acessorios_mtime, seriais_resumo)}
{secao_devolvidos(devolvidos_mtime, devolvidos_resumo)}

</main>
<footer>Gerado automaticamente a partir das planilhas de controle · Rocha Telecom</footer>
</div>
</div>
<script>
document.querySelectorAll('input.filtro[data-target]').forEach(function(inp) {{
  inp.addEventListener('input', function() {{
    var table = document.getElementById(inp.dataset.target);
    var q = inp.value.toLowerCase();
    table.querySelectorAll('tbody tr').forEach(function(tr) {{
      tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
    }});
  }});
}});

// ---- ordenacao das tabelas por coluna (clique no cabecalho) ----
function parseCelula(texto) {{
  texto = texto.trim();
  var dataMatch = texto.match(/^(\d{{2}})\/(\d{{2}})\/(\d{{4}})$/);
  if (dataMatch) return new Date(dataMatch[3], dataMatch[2] - 1, dataMatch[1]).getTime();
  var limpo = texto.replace(/[^\d,.\-]/g, '');
  if (limpo && /^-?\d{{1,3}}(\.\d{{3}})*(,\d+)?$|^-?\d+(,\d+)?$/.test(limpo)) {{
    var n = parseFloat(limpo.replace(/\./g, '').replace(',', '.'));
    if (!isNaN(n)) return n;
  }}
  if (/^-?\d+(\.\d+)?$/.test(texto)) return parseFloat(texto);
  return texto.toLowerCase();
}}

document.querySelectorAll('table.sortable').forEach(function(table) {{
  var headers = table.querySelectorAll('thead th');
  headers.forEach(function(th, idx) {{
    th.addEventListener('click', function() {{
      var dir = th.dataset.dir === 'asc' ? 'desc' : 'asc';
      headers.forEach(function(h) {{
        h.dataset.dir = '';
        var ind = h.querySelector('.sort-ind');
        if (ind) ind.remove();
      }});
      th.dataset.dir = dir;
      var ind = document.createElement('span');
      ind.className = 'sort-ind';
      ind.textContent = dir === 'asc' ? ' ▲' : ' ▼';
      th.appendChild(ind);

      var tbody = table.querySelector('tbody');
      var linhas = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
      linhas.sort(function(a, b) {{
        var av = parseCelula(a.children[idx].textContent);
        var bv = parseCelula(b.children[idx].textContent);
        if (av < bv) return dir === 'asc' ? -1 : 1;
        if (av > bv) return dir === 'asc' ? 1 : -1;
        return 0;
      }});
      linhas.forEach(function(tr) {{ tbody.appendChild(tr); }});
    }});
  }});
}});

// ---- alternancia de tema claro/escuro ----
function aplicarTema(tema) {{
  document.documentElement.setAttribute('data-tema', tema);
  try {{ localStorage.setItem('bi-tema', tema); }} catch (e) {{}}
  rotularBotaoTema(tema);
  repintarGraficos();
}}

// o Chart.js guarda as opcoes de escala ja resolvidas, entao mudar Chart.defaults nao
// repinta os eixos dos graficos que ja existem — a cor precisa ser escrita em cada escala
function repintarGraficos() {{
  // corDoTema é declaração de função, então já está disponível mesmo definida mais abaixo
  var texto = corDoTema('--chart-texto');
  var grade = corDoTema('--chart-grade');
  Chart.defaults.color = texto;
  Chart.defaults.borderColor = grade;
  document.querySelectorAll('canvas').forEach(function(cv) {{
    var grafico = Chart.getChart(cv);
    if (!grafico) return;
    // só escreve nas folhas: reatribuir options.scales[id] cria referência circular no
    // proxy de opções do Chart.js e estoura a pilha
    var escalas = grafico.options.scales || {{}};
    Object.keys(grafico.scales).forEach(function(id) {{
      var op = escalas[id];
      if (!op) return;
      if (op.ticks) op.ticks.color = texto;
      if (op.grid) op.grid.color = grade;
      if (op.title) op.title.color = texto;
    }});
    var legenda = grafico.options.plugins && grafico.options.plugins.legend;
    if (legenda && legenda.labels) legenda.labels.color = texto;
    var subtitulo = grafico.options.plugins && grafico.options.plugins.subtitle;
    if (subtitulo) subtitulo.color = texto;
    grafico.update('none');
  }});
}}

function rotularBotaoTema(tema) {{
  var claro = tema === 'light';
  document.getElementById('tema-icone').textContent = claro ? '☀️' : '🌙';
  document.getElementById('tema-texto').textContent = claro ? 'Modo claro' : 'Modo escuro';
}}

rotularBotaoTema(document.documentElement.getAttribute('data-tema') || 'dark');
document.getElementById('tema-toggle').addEventListener('click', function() {{
  aplicarTema(document.documentElement.getAttribute('data-tema') === 'light' ? 'dark' : 'light');
}});

// ---- destaque do item ativo no menu lateral conforme a rolagem ----
var navLinks = document.querySelectorAll('.sidebar .nav-link');
var secoes = Array.prototype.slice.call(document.querySelectorAll('main section[id]'));
function atualizarMenuAtivo() {{
  var pos = window.scrollY + 100;
  var atual = secoes[0];
  secoes.forEach(function(s) {{ if (s.offsetTop <= pos) atual = s; }});
  navLinks.forEach(function(a) {{
    a.classList.toggle('active', atual && a.getAttribute('href') === '#' + atual.id);
  }});
}}
window.addEventListener('scroll', atualizarMenuAtivo);
atualizarMenuAtivo();

// ---- helpers compartilhados de filtro (multi-select, faixas de dias) ----
function valoresUnicos(dados, campo) {{
  var vistos = {{}};
  var out = [];
  dados.forEach(function(r) {{
    var v = r[campo];
    if (v && !vistos[v]) {{ vistos[v] = true; out.push(v); }}
  }});
  out.sort();
  return out;
}}

var FAIXAS_DIAS = [
  ['np', 'Ainda não processado'], ['fin', 'Finalizado (produtos na filial)'],
  ['futura', 'Data de entrega futura'], ['nfat', 'Não será faturado'],
  ['0-3', '0-3 dias'], ['4-7', '4-7 dias'],
  ['8-15', '8-15 dias'], ['16-30', '16-30 dias'], ['30+', '30+ dias']
];

// `categoria` separa as linhas que não têm contagem de dias ("não processado",
// "finalizado", "entrega futura"). Sem ela as três caíam todas em "não processado",
// porque o código usava número negativo como marcador. A tabela de transferências
// chama sem categoria e continua com o comportamento antigo.
function diasNaFaixa(dias, faixa, categoria) {{
  if (categoria) return faixa === categoria;
  if (dias < 0) return faixa === 'np';
  if (faixa === 'np' || faixa === 'fin' || faixa === 'futura' || faixa === 'nfat') return false;
  if (faixa === '0-3') return dias <= 3;
  if (faixa === '4-7') return dias >= 4 && dias <= 7;
  if (faixa === '8-15') return dias >= 8 && dias <= 15;
  if (faixa === '16-30') return dias >= 16 && dias <= 30;
  if (faixa === '30+') return dias > 30;
  return true;
}}

// ---- componente de multi-selecao (checkbox dropdown), reutilizavel em qualquer tabela ----
function criarMultiSelect(id, pares, set, onChange) {{
  var raiz = document.getElementById('msel-' + id);
  var btn = raiz.querySelector('.msel-btn');
  var painel = raiz.querySelector('.msel-panel');
  var opcoes = raiz.querySelector('.msel-options');
  painel.addEventListener('click', function(e) {{ e.stopPropagation(); }});

  pares.forEach(function(par) {{
    var label = document.createElement('label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = par[0];
    cb.addEventListener('change', function() {{
      if (cb.checked) set.add(par[0]); else set.delete(par[0]);
      atualizarLabel();
      onChange();
    }});
    label.appendChild(cb);
    label.appendChild(document.createTextNode(par[1]));
    opcoes.appendChild(label);
  }});

  function atualizarLabel() {{
    var padrao = btn.dataset.default;
    if (set.size === 0) btn.textContent = padrao;
    else if (set.size === 1) btn.textContent = pares.filter(function(p) {{ return set.has(p[0]); }})[0][1];
    else btn.textContent = set.size + ' selecionadas';
  }}

  btn.addEventListener('click', function(e) {{
    e.stopPropagation();
    var jaAberto = raiz.classList.contains('open');
    document.querySelectorAll('.msel.open').forEach(function(m) {{ m.classList.remove('open'); }});
    if (!jaAberto) raiz.classList.add('open');
  }});

  raiz.querySelector('[data-act="all"]').addEventListener('click', function() {{
    opcoes.querySelectorAll('input').forEach(function(cb) {{ cb.checked = true; set.add(cb.value); }});
    atualizarLabel();
    onChange();
  }});
  raiz.querySelector('[data-act="none"]').addEventListener('click', function() {{
    opcoes.querySelectorAll('input').forEach(function(cb) {{ cb.checked = false; }});
    set.clear();
    atualizarLabel();
    onChange();
  }});

  return {{
    limpar: function() {{
      set.clear();
      opcoes.querySelectorAll('input').forEach(function(cb) {{ cb.checked = false; }});
      atualizarLabel();
    }}
  }};
}}

document.addEventListener('click', function() {{
  document.querySelectorAll('.msel.open').forEach(function(m) {{ m.classList.remove('open'); }});
}});

// ---- exportacao: graficos como PNG, tabelas como Excel (.xlsx) ----
var EXPORTADORES = {{}};

function registrarExportDOM(tableId) {{
  EXPORTADORES[tableId] = function() {{
    var table = document.getElementById(tableId);
    var headers = Array.prototype.map.call(table.querySelectorAll('thead th'), function(th) {{
      return (th.childNodes[0] ? th.childNodes[0].textContent : th.textContent).trim();
    }});
    var linhas = Array.prototype.filter.call(table.querySelectorAll('tbody tr'), function(tr) {{
      return tr.style.display !== 'none';
    }}).map(function(tr) {{
      return Array.prototype.map.call(tr.children, function(td) {{ return td.textContent.trim(); }});
    }});
    return [headers].concat(linhas);
  }};
}}

function exportarAOA(aoa, nomeArquivo) {{
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dados');
  XLSX.writeFile(wb, nomeArquivo.replace(/^tbl-/, '') + '.xlsx');
}}

document.addEventListener('click', function(e) {{
  var expBtn = e.target.closest('[data-export]');
  if (expBtn) {{
    var tableId = expBtn.dataset.export;
    var fn = EXPORTADORES[tableId];
    if (fn) exportarAOA(fn(), tableId);
  }}
  var chartBtn = e.target.closest('[data-chart-export]');
  if (chartBtn) {{
    var canvas = document.getElementById(chartBtn.dataset.chartExport);
    var link = document.createElement('a');
    link.download = chartBtn.dataset.chartExport.replace(/^chart-/, '') + '.png';
    link.href = canvas.toDataURL('image/png', 1.0);
    link.click();
  }}
}});

['tbl-pedidos-filial', 'tbl-notas', 'tbl-notas-excecoes', 'tbl-amet', 'tbl-devia', 'tbl-upmaster', 'tbl-transf', 'tbl-malotes-filial', 'tbl-manutencoes', 'tbl-obras', 'tbl-campanhas-filiais'].forEach(registrarExportDOM);

// ---- fabrica generica de tabela paginada com busca + multi-selecao, usada pelas tabelas de acessorios ----
function criarTabelaPaginada(opts) {{
  var estado = {{ filtro: '', sortCol: opts.sortInicial || null, sortDir: 'desc', pagina: 1, sets: {{}} }};
  opts.filtros.forEach(function(f) {{ estado.sets[f.campo] = new Set(); }});

  function dadosFiltrados() {{
    var q = estado.filtro.toLowerCase();
    var deTs = null, ateTs = null;
    if (opts.filtroData) {{
      var deVal = document.getElementById(opts.filtroData.deId).value;
      var ateVal = document.getElementById(opts.filtroData.ateId).value;
      if (deVal) deTs = new Date(deVal + 'T00:00:00').getTime();
      if (ateVal) ateTs = new Date(ateVal + 'T23:59:59').getTime();
    }}
    estado.deTs = deTs;
    estado.ateTs = ateTs;
    var out = opts.dados.filter(function(r) {{
      if (q && opts.busca(r).toLowerCase().indexOf(q) === -1) return false;
      for (var campo in estado.sets) {{
        if (estado.sets[campo].size && !estado.sets[campo].has(r[campo])) return false;
      }}
      if (opts.filtroData) {{
        var ts = r[opts.filtroData.campo];
        if (deTs !== null && ts < deTs) return false;
        if (ateTs !== null && ts > ateTs) return false;
      }}
      return true;
    }});
    if (estado.sortCol) {{
      var getVal = opts.colunas[estado.sortCol];
      var dir = estado.sortDir === 'asc' ? 1 : -1;
      out.sort(function(a, b) {{
        var av = getVal(a), bv = getVal(b);
        if (av < bv) return -1 * dir;
        if (av > bv) return 1 * dir;
        return 0;
      }});
    }}
    return out;
  }}

  function render() {{
    var filtrados = dadosFiltrados();
    // deixa a secao recalcular cards/graficos com o mesmo recorte da tabela
    if (opts.aoFiltrar) opts.aoFiltrar(filtrados, estado.deTs, estado.ateTs);
    var pager = document.getElementById(opts.pagerId);
    if (opts.pageSize) {{
      var totalPaginas = Math.max(1, Math.ceil(filtrados.length / opts.pageSize));
      if (estado.pagina > totalPaginas) estado.pagina = totalPaginas;
      var inicio = (estado.pagina - 1) * opts.pageSize;
      document.getElementById(opts.tbodyId).innerHTML = filtrados.slice(inicio, inicio + opts.pageSize).map(opts.linhaHtml).join('');
      pager.innerHTML =
        '<button class="pp-prev" ' + (estado.pagina <= 1 ? 'disabled' : '') + '>‹ Anterior</button>' +
        '<span>Página ' + estado.pagina + ' de ' + totalPaginas + ' (' + filtrados.length + ' itens)</span>' +
        '<button class="pp-next" ' + (estado.pagina >= totalPaginas ? 'disabled' : '') + '>Próxima ›</button>';
      pager.querySelector('.pp-prev').addEventListener('click', function() {{ estado.pagina--; render(); }});
      pager.querySelector('.pp-next').addEventListener('click', function() {{ estado.pagina++; render(); }});
    }} else {{
      document.getElementById(opts.tbodyId).innerHTML = filtrados.map(opts.linhaHtml).join('');
      pager.innerHTML = '<span>Mostrando ' + filtrados.length + ' de ' + opts.dados.length + ' itens</span>';
    }}
    document.querySelectorAll('#' + opts.tableId + ' thead th').forEach(function(th) {{
      var ind = th.querySelector('.sort-ind');
      if (ind) ind.remove();
      if (th.dataset.col === estado.sortCol) {{
        var span = document.createElement('span');
        span.className = 'sort-ind';
        span.textContent = estado.sortDir === 'asc' ? ' ▲' : ' ▼';
        th.appendChild(span);
      }}
    }});
  }}

  document.getElementById(opts.buscaId).addEventListener('input', function(e) {{
    estado.filtro = e.target.value;
    estado.pagina = 1;
    render();
  }});

  if (opts.filtroData) {{
    [opts.filtroData.deId, opts.filtroData.ateId].forEach(function(id) {{
      document.getElementById(id).addEventListener('change', function() {{
        estado.pagina = 1;
        render();
      }});
    }});
  }}

  var msels = opts.filtros.map(function(f) {{
    var pares = f.pares || valoresUnicos(opts.dados, f.campo).map(function(v) {{ return [v, v]; }});
    return criarMultiSelect(f.id, pares, estado.sets[f.campo], function() {{
      estado.pagina = 1;
      render();
    }});
  }});

  document.getElementById(opts.limparId).addEventListener('click', function() {{
    estado.filtro = '';
    document.getElementById(opts.buscaId).value = '';
    if (opts.filtroData) {{
      document.getElementById(opts.filtroData.deId).value = '';
      document.getElementById(opts.filtroData.ateId).value = '';
    }}
    msels.forEach(function(m) {{ m.limpar(); }});
    estado.pagina = 1;
    render();
  }});

  document.querySelectorAll('#' + opts.tableId + ' thead th').forEach(function(th) {{
    th.style.cursor = 'pointer';
    th.addEventListener('click', function() {{
      var col = th.dataset.col;
      estado.sortDir = (estado.sortCol === col && estado.sortDir === 'asc') ? 'desc' : 'asc';
      estado.sortCol = col;
      estado.pagina = 1;
      render();
    }});
  }});

  if (opts.colunasExport) {{
    EXPORTADORES[opts.tableId] = function() {{
      var filtrados = dadosFiltrados();
      var headers = opts.colunasExport.map(function(c) {{ return c.label; }});
      var linhas = filtrados.map(function(r) {{ return opts.colunasExport.map(function(c) {{ return c.get(r); }}); }});
      return [headers].concat(linhas);
    }};
  }}

  render();
}}

// ---- tabela de pedidos (paginada, com muitas linhas nao da pra desenhar tudo de uma vez) ----
(function() {{
  var DADOS = {pedidos_detalhe_json};
  var PAGE_SIZE = 50;
  var estado = {{
    filtro: '', sortCol: null, sortDir: 'asc', pagina: 1,
    filial: new Set(), status: new Set(), stprod: new Set(), entrada: new Set(), dias: new Set()
  }};

  var COLS = {{
    dreal: {{ sort: function(r) {{ return r.dreal_ts; }} }},
    dproc: {{ sort: function(r) {{ return r.dproc_ts; }} }},
    dprev: {{ sort: function(r) {{ return r.dprev_ts; }} }},
    fil: {{ sort: function(r) {{ return r.fil.toLowerCase(); }} }},
    ped: {{ sort: function(r) {{ return r.ped; }} }},
    qtd: {{ sort: function(r) {{ return r.qtd; }} }},
    ref: {{ sort: function(r) {{ return r.ref.toLowerCase(); }} }},
    desc: {{ sort: function(r) {{ return r.desc.toLowerCase(); }} }},
    stprod: {{ sort: function(r) {{ return r.stprod.toLowerCase(); }} }},
    dias: {{ sort: function(r) {{ return r.dias_num; }} }},
    entrada: {{ sort: function(r) {{ return r.entrada.toLowerCase(); }} }},
    status: {{ sort: function(r) {{ return r.status.toLowerCase(); }} }}
  }};

  function linhaHtml(r) {{
    return '<tr><td>' + r.dreal + '</td><td>' + r.dproc + '</td><td>' + r.dprev + '</td>' +
      '<td>' + r.fil + '</td><td class="num">' + r.ped + '</td><td class="num">' + r.qtd + '</td>' +
      '<td>' + r.ref + '</td><td>' + r.desc + '</td><td>' + r.stprod + '</td><td class="num">' + r.dias + '</td>' +
      '<td>' + r.entrada + '</td><td><span class="pill ' + r.cls + '">' + r.status + '</span></td></tr>';
  }}

  function dadosFiltrados() {{
    var q = estado.filtro.toLowerCase();
    var out = DADOS.filter(function(r) {{
      if (q && (r.fil + ' ' + r.ped + ' ' + r.ref + ' ' + r.desc + ' ' + r.status).toLowerCase().indexOf(q) === -1) return false;
      if (estado.filial.size && !estado.filial.has(r.fil)) return false;
      if (estado.status.size && !estado.status.has(r.status)) return false;
      if (estado.stprod.size && !estado.stprod.has(r.stprod)) return false;
      if (estado.entrada.size && !estado.entrada.has(r.entrada)) return false;
      if (estado.dias.size) {{
        var algumaFaixaBate = false;
        estado.dias.forEach(function(faixa) {{ if (diasNaFaixa(r.dias_num, faixa, r.dias_cat)) algumaFaixaBate = true; }});
        if (!algumaFaixaBate) return false;
      }}
      return true;
    }});
    if (estado.sortCol) {{
      var getVal = COLS[estado.sortCol].sort;
      var dir = estado.sortDir === 'asc' ? 1 : -1;
      out.sort(function(a, b) {{
        var av = getVal(a), bv = getVal(b);
        if (av < bv) return -1 * dir;
        if (av > bv) return 1 * dir;
        return 0;
      }});
    }}
    return out;
  }}

  function render() {{
    var filtrados = dadosFiltrados();
    var totalPaginas = Math.max(1, Math.ceil(filtrados.length / PAGE_SIZE));
    if (estado.pagina > totalPaginas) estado.pagina = totalPaginas;
    var inicio = (estado.pagina - 1) * PAGE_SIZE;
    var pagina = filtrados.slice(inicio, inicio + PAGE_SIZE);

    document.getElementById('tbody-pedidos-detalhe').innerHTML = pagina.map(linhaHtml).join('');

    var pager = document.getElementById('pager-pedidos-detalhe');
    pager.innerHTML =
      '<button id="pp-prev" ' + (estado.pagina <= 1 ? 'disabled' : '') + '>‹ Anterior</button>' +
      '<span>Página ' + estado.pagina + ' de ' + totalPaginas + ' (' + filtrados.length + ' pedidos)</span>' +
      '<button id="pp-next" ' + (estado.pagina >= totalPaginas ? 'disabled' : '') + '>Próxima ›</button>';
    document.getElementById('pp-prev').addEventListener('click', function() {{ estado.pagina--; render(); }});
    document.getElementById('pp-next').addEventListener('click', function() {{ estado.pagina++; render(); }});

    document.querySelectorAll('#tbl-pedidos-detalhe thead th').forEach(function(th) {{
      var ind = th.querySelector('.sort-ind');
      if (ind) ind.remove();
      if (th.dataset.col === estado.sortCol) {{
        var span = document.createElement('span');
        span.className = 'sort-ind';
        span.textContent = estado.sortDir === 'asc' ? ' ▲' : ' ▼';
        th.appendChild(span);
      }}
    }});
  }}

  document.getElementById('filtro-pedidos-detalhe').addEventListener('input', function(e) {{
    estado.filtro = e.target.value;
    estado.pagina = 1;
    render();
  }});

  function aoMudarFiltro() {{
    estado.pagina = 1;
    render();
  }}

  var mselFilial = criarMultiSelect('filial', valoresUnicos(DADOS, 'fil').map(function(v) {{ return [v, v]; }}), estado.filial, aoMudarFiltro);
  var mselStatus = criarMultiSelect('status', valoresUnicos(DADOS, 'status').map(function(v) {{ return [v, v]; }}), estado.status, aoMudarFiltro);
  var mselStprod = criarMultiSelect('stprod', valoresUnicos(DADOS, 'stprod').map(function(v) {{ return [v, v]; }}), estado.stprod, aoMudarFiltro);
  var mselEntrada = criarMultiSelect('entrada', valoresUnicos(DADOS, 'entrada').map(function(v) {{ return [v, v]; }}), estado.entrada, aoMudarFiltro);
  var mselDias = criarMultiSelect('dias', FAIXAS_DIAS, estado.dias, aoMudarFiltro);

  document.getElementById('limpar-filtros-pedidos').addEventListener('click', function() {{
    estado.filtro = '';
    estado.pagina = 1;
    document.getElementById('filtro-pedidos-detalhe').value = '';
    [mselFilial, mselStatus, mselStprod, mselEntrada, mselDias].forEach(function(m) {{ m.limpar(); }});
    render();
  }});

  document.querySelectorAll('#tbl-pedidos-detalhe thead th').forEach(function(th) {{
    th.style.cursor = 'pointer';
    th.addEventListener('click', function() {{
      var col = th.dataset.col;
      estado.sortDir = (estado.sortCol === col && estado.sortDir === 'asc') ? 'desc' : 'asc';
      estado.sortCol = col;
      estado.pagina = 1;
      render();
    }});
  }});

  EXPORTADORES['tbl-pedidos-detalhe'] = function() {{
    var headers = ['Data Realização', 'Data Processamento', 'Data Prevista Entrega', 'Filial Destino', 'Nº Pedido GN', 'Qtde', 'Referência', 'Descrição do Produto', 'Status do Produto', 'Dias em Aberto', 'Entrada no Sistema', 'Status'];
    var linhas = dadosFiltrados().map(function(r) {{
      return [r.dreal, r.dproc, r.dprev, r.fil, r.ped, r.qtd, r.ref, r.desc, r.stprod, r.dias, r.entrada, r.status];
    }});
    return [headers].concat(linhas);
  }};

  render();
}})();

// ---- tabela de transferencias pendentes (com filtros combinaveis, sem paginacao pois sao poucas centenas) ----
(function() {{
  var DADOS = {transf_todas_json};
  var estado = {{
    filtro: '', sortCol: 'dias', sortDir: 'desc',
    orig: new Set(), dest: new Set(), crit: new Set(), prazo: new Set(), dias: new Set()
  }};

  var FAIXAS_DIAS_TRANSF = [
    ['0-3', '0-3 dias'], ['4-7', '4-7 dias'], ['8-15', '8-15 dias'], ['16-30', '16-30 dias'], ['30+', '30+ dias']
  ];

  var COLS = {{
    orig: {{ sort: function(r) {{ return r.orig.toLowerCase(); }} }},
    dias: {{ sort: function(r) {{ return r.dias_num; }} }},
    dest: {{ sort: function(r) {{ return r.dest.toLowerCase(); }} }},
    user: {{ sort: function(r) {{ return r.user.toLowerCase(); }} }},
    nf: {{ sort: function(r) {{ return r.nf.toLowerCase(); }} }},
    prod: {{ sort: function(r) {{ return r.prod.toLowerCase(); }} }},
    desc: {{ sort: function(r) {{ return r.desc.toLowerCase(); }} }},
    qtd: {{ sort: function(r) {{ return r.qtd; }} }},
    crit: {{ sort: function(r) {{ return r.crit.toLowerCase(); }} }}
  }};

  function linhaHtml(r) {{
    return '<tr><td>' + r.orig + '</td><td class="num">' + r.dias + '</td><td>' + r.dest + '</td>' +
      '<td>' + r.user + '</td><td>' + r.nf + '</td><td>' + r.prod + '</td><td>' + r.desc + '</td>' +
      '<td class="num">' + r.qtd + '</td><td><span class="pill ' + r.cls + '">' + r.crit + '</span></td>' +
      '<td>' + r.rastreio + '</td>' +
      '<td>' + (r.entrega_cls ? '<span class="pill ' + r.entrega_cls + '">' + r.entrega + '</span>' : r.entrega) + '</td></tr>';
  }}

  function dadosFiltrados() {{
    var q = estado.filtro.toLowerCase();
    var out = DADOS.filter(function(r) {{
      if (q && (r.orig + ' ' + r.dest + ' ' + r.nf + ' ' + r.prod + ' ' + r.desc).toLowerCase().indexOf(q) === -1) return false;
      if (estado.orig.size && !estado.orig.has(r.orig)) return false;
      if (estado.dest.size && !estado.dest.has(r.dest)) return false;
      if (estado.crit.size && !estado.crit.has(r.crit)) return false;
      if (estado.prazo.size && !estado.prazo.has(r.prazo)) return false;
      if (estado.dias.size) {{
        var algumaFaixaBate = false;
        estado.dias.forEach(function(faixa) {{ if (diasNaFaixa(r.dias_num, faixa)) algumaFaixaBate = true; }});
        if (!algumaFaixaBate) return false;
      }}
      return true;
    }});
    if (estado.sortCol) {{
      var getVal = COLS[estado.sortCol].sort;
      var dir = estado.sortDir === 'asc' ? 1 : -1;
      out.sort(function(a, b) {{
        var av = getVal(a), bv = getVal(b);
        if (av < bv) return -1 * dir;
        if (av > bv) return 1 * dir;
        return 0;
      }});
    }}
    return out;
  }}

  function render() {{
    var filtrados = dadosFiltrados();
    document.getElementById('tbody-transf').innerHTML = filtrados.map(linhaHtml).join('');
    document.getElementById('contador-transf').innerHTML =
      '<span>Mostrando ' + filtrados.length + ' de ' + DADOS.length + ' transferências</span>';

    document.querySelectorAll('#tbl-transf thead th').forEach(function(th) {{
      var ind = th.querySelector('.sort-ind');
      if (ind) ind.remove();
      if (th.dataset.col === estado.sortCol) {{
        var span = document.createElement('span');
        span.className = 'sort-ind';
        span.textContent = estado.sortDir === 'asc' ? ' ▲' : ' ▼';
        th.appendChild(span);
      }}
    }});
  }}

  document.getElementById('filtro-transf').addEventListener('input', function(e) {{
    estado.filtro = e.target.value;
    render();
  }});

  var mselOrig = criarMultiSelect('torig', valoresUnicos(DADOS, 'orig').map(function(v) {{ return [v, v]; }}), estado.orig, render);
  var mselDest = criarMultiSelect('tdest', valoresUnicos(DADOS, 'dest').map(function(v) {{ return [v, v]; }}), estado.dest, render);
  var mselCrit = criarMultiSelect('tcrit', valoresUnicos(DADOS, 'crit').map(function(v) {{ return [v, v]; }}), estado.crit, render);
  var mselPrazo = criarMultiSelect('tprazo', valoresUnicos(DADOS, 'prazo').map(function(v) {{ return [v, v]; }}), estado.prazo, render);
  var mselDias = criarMultiSelect('tdias', FAIXAS_DIAS_TRANSF, estado.dias, render);

  document.getElementById('limpar-filtros-transf').addEventListener('click', function() {{
    estado.filtro = '';
    document.getElementById('filtro-transf').value = '';
    [mselOrig, mselDest, mselCrit, mselPrazo, mselDias].forEach(function(m) {{ m.limpar(); }});
    render();
  }});

  document.querySelectorAll('#tbl-transf thead th').forEach(function(th) {{
    th.style.cursor = 'pointer';
    th.addEventListener('click', function() {{
      var col = th.dataset.col;
      estado.sortDir = (estado.sortCol === col && estado.sortDir === 'asc') ? 'desc' : 'asc';
      estado.sortCol = col;
      render();
    }});
  }});

  render();
}})();

// ---- tabelas de acessorios (diversos e fidelizados TIM) ----
function linhaAcessorioHtml(r) {{
  return '<tr><td>' + r.filial + '</td><td>' + r.ref + '</td><td>' + r.desc + '</td><td>' + r.subgrupo + '</td>' +
    '<td>' + r.fabricante + '</td><td class="num">' + r.saldo + '</td><td class="num">' + r.disponivel + '</td>' +
    '<td class="num">' + brlJs(r.valor) + '</td><td class="num">' + brlJs(r.valor_total) + '</td></tr>';
}}

function brlJs(v) {{
  return 'R$ ' + v.toLocaleString('pt-BR', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
}}

var ACESSORIOS_COLS = {{
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  ref: function(r) {{ return r.ref.toLowerCase(); }},
  desc: function(r) {{ return r.desc.toLowerCase(); }},
  subgrupo: function(r) {{ return r.subgrupo.toLowerCase(); }},
  fabricante: function(r) {{ return r.fabricante.toLowerCase(); }},
  saldo: function(r) {{ return r.saldo; }},
  disponivel: function(r) {{ return r.disponivel; }},
  valor: function(r) {{ return r.valor; }},
  valor_total: function(r) {{ return r.valor_total; }}
}};

function buscaAcessorio(r) {{ return r.filial + ' ' + r.ref + ' ' + r.desc + ' ' + r.fabricante; }}

criarTabelaPaginada({{
  dados: {acessorios_diversos_json},
  pageSize: 50, sortInicial: 'saldo',
  tbodyId: 'tbody-acessorios', pagerId: 'pager-acessorios', tableId: 'tbl-acessorios',
  buscaId: 'filtro-acessorios', limparId: 'limpar-acessorios',
  colunas: ACESSORIOS_COLS, linhaHtml: linhaAcessorioHtml, busca: buscaAcessorio,
  filtros: [
    {{ id: 'acessorios-filial', campo: 'filial' }},
    {{ id: 'acessorios-subgrupo', campo: 'subgrupo' }},
    {{ id: 'acessorios-fabricante', campo: 'fabricante' }}
  ],
  colunasExport: [
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Referência', get: function(r) {{ return r.ref; }} }},
    {{ label: 'Descrição', get: function(r) {{ return r.desc; }} }},
    {{ label: 'Tipo', get: function(r) {{ return r.subgrupo; }} }},
    {{ label: 'Fabricante', get: function(r) {{ return r.fabricante; }} }},
    {{ label: 'Saldo', get: function(r) {{ return r.saldo; }} }},
    {{ label: 'Disponível', get: function(r) {{ return r.disponivel; }} }},
    {{ label: 'Valor Unit.', get: function(r) {{ return r.valor; }} }},
    {{ label: 'Valor Total', get: function(r) {{ return r.valor_total; }} }}
  ]
}});

function linhaSerialHtml(r) {{
  return '<tr><td>' + r.filial + '</td><td>' + r.serial + '</td><td>' + r.desc + '</td>' +
    '<td>' + r.fabricante + '</td><td>' + r.data_compra + '</td>' +
    '<td class="num">' + r.dias + '</td><td class="num">' + brlJs(r.valor) + '</td></tr>';
}}

var SERIAIS_COLS = {{
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  serial: function(r) {{ return r.serial.toLowerCase(); }},
  desc: function(r) {{ return r.desc.toLowerCase(); }},
  fabricante: function(r) {{ return r.fabricante.toLowerCase(); }},
  data_compra: function(r) {{ return r.dias; }},
  dias: function(r) {{ return r.dias; }},
  valor: function(r) {{ return r.valor; }}
}};

function buscaSerial(r) {{ return r.filial + ' ' + r.serial + ' ' + r.desc + ' ' + r.fabricante; }}

var FAIXAS_DIAS_ESTOQUE = [
  ['0-30 dias', '0-30 dias'], ['31-90 dias', '31-90 dias'], ['91-180 dias', '91-180 dias'],
  ['181-365 dias', '181-365 dias'], ['365+ dias', '365+ dias']
];

criarTabelaPaginada({{
  dados: {seriais_json},
  pageSize: 50, sortInicial: 'dias',
  tbodyId: 'tbody-acessorios-tim', pagerId: 'pager-acessorios-tim', tableId: 'tbl-acessorios-tim',
  buscaId: 'filtro-acessorios-tim', limparId: 'limpar-acessorios-tim',
  colunas: SERIAIS_COLS, linhaHtml: linhaSerialHtml, busca: buscaSerial,
  filtros: [
    {{ id: 'acessorios-tim-filial', campo: 'filial' }},
    {{ id: 'acessorios-tim-fabricante', campo: 'fabricante' }},
    {{ id: 'acessorios-tim-dias', campo: 'dias_faixa', pares: FAIXAS_DIAS_ESTOQUE }}
  ],
  colunasExport: [
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Serial', get: function(r) {{ return r.serial; }} }},
    {{ label: 'Descrição', get: function(r) {{ return r.desc; }} }},
    {{ label: 'Fabricante', get: function(r) {{ return r.fabricante; }} }},
    {{ label: 'Data Compra', get: function(r) {{ return r.data_compra; }} }},
    {{ label: 'Dias em Estoque', get: function(r) {{ return r.dias; }} }},
    {{ label: 'Valor', get: function(r) {{ return r.valor; }} }}
  ]
}});

// ---- tabela de devolvidos e defeitos ----
function linhaDevolvidoHtml(r) {{
  return '<tr><td>' + r.filial + '</td><td>' + r.desc + '</td><td>' + r.grupo + '</td>' +
    '<td>' + r.fabricante + '</td><td class="num">' + r.saldo + '</td><td class="num">' + brlJs(r.custo) + '</td>' +
    '<td class="num">' + brlJs(r.custo_total) + '</td><td>' + r.data_mov + '</td></tr>';
}}

var DEVOLVIDOS_COLS = {{
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  desc: function(r) {{ return r.desc.toLowerCase(); }},
  grupo: function(r) {{ return r.grupo.toLowerCase(); }},
  fabricante: function(r) {{ return r.fabricante.toLowerCase(); }},
  saldo: function(r) {{ return r.saldo; }},
  custo: function(r) {{ return r.custo; }},
  custo_total: function(r) {{ return r.custo_total; }},
  data_mov: function(r) {{ return r.data_mov; }}
}};

function buscaDevolvido(r) {{ return r.filial + ' ' + r.desc + ' ' + r.fabricante; }}

criarTabelaPaginada({{
  dados: {devolvidos_json},
  pageSize: null, sortInicial: 'saldo',
  tbodyId: 'tbody-devolvidos', pagerId: 'pager-devolvidos', tableId: 'tbl-devolvidos',
  buscaId: 'filtro-devolvidos', limparId: 'limpar-devolvidos',
  colunas: DEVOLVIDOS_COLS, linhaHtml: linhaDevolvidoHtml, busca: buscaDevolvido,
  filtros: [
    {{ id: 'devolvidos-filial', campo: 'filial' }},
    {{ id: 'devolvidos-grupo', campo: 'grupo' }},
    {{ id: 'devolvidos-fabricante', campo: 'fabricante' }}
  ],
  colunasExport: [
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Descrição', get: function(r) {{ return r.desc; }} }},
    {{ label: 'Categoria', get: function(r) {{ return r.grupo; }} }},
    {{ label: 'Fabricante', get: function(r) {{ return r.fabricante; }} }},
    {{ label: 'Saldo', get: function(r) {{ return r.saldo; }} }},
    {{ label: 'Custo Unit.', get: function(r) {{ return r.custo; }} }},
    {{ label: 'Custo Total', get: function(r) {{ return r.custo_total; }} }},
    {{ label: 'Última Movimentação', get: function(r) {{ return r.data_mov; }} }}
  ]
}});

// ---- tabela de solicitacoes de postagem de malotes ----
function linhaMaloteLogHtml(r) {{
  return '<tr><td class="num">' + r.id + '</td><td>' + r.solicitante + '</td><td>' + r.filial + '</td>' +
    '<td class="num">' + r.qtd + '</td><td>' + r.conteudo + '</td><td>' + r.data_sol + '</td>' +
    '<td><span class="pill ' + r.status_cls + '">' + r.status + '</span></td>' +
    '<td>' + r.data_post + '</td><td>' + r.quem_postou + '</td></tr>';
}}

var MALOTES_LOG_COLS = {{
  id: function(r) {{ return r.id; }},
  solicitante: function(r) {{ return r.solicitante.toLowerCase(); }},
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  qtd: function(r) {{ return r.qtd; }},
  conteudo: function(r) {{ return r.conteudo.toLowerCase(); }},
  data_sol: function(r) {{ return r.data_sol_ts; }},
  status: function(r) {{ return r.status.toLowerCase(); }},
  data_post: function(r) {{ return r.data_post; }},
  quem_postou: function(r) {{ return r.quem_postou.toLowerCase(); }}
}};

function buscaMaloteLog(r) {{ return r.id + ' ' + r.solicitante + ' ' + r.filial + ' ' + r.conteudo; }}

criarTabelaPaginada({{
  dados: {malotes_log_json},
  pageSize: 50, sortInicial: 'data_sol',
  tbodyId: 'tbody-malotes-log', pagerId: 'pager-malotes-log', tableId: 'tbl-malotes-log',
  buscaId: 'filtro-malotes-log', limparId: 'limpar-malotes-log',
  colunas: MALOTES_LOG_COLS, linhaHtml: linhaMaloteLogHtml, busca: buscaMaloteLog,
  filtros: [
    {{ id: 'malotes-log-filial', campo: 'filial' }},
    {{ id: 'malotes-log-status', campo: 'status' }},
    {{ id: 'malotes-log-conteudo', campo: 'conteudo' }}
  ],
  colunasExport: [
    {{ label: 'ID', get: function(r) {{ return r.id; }} }},
    {{ label: 'Solicitante', get: function(r) {{ return r.solicitante; }} }},
    {{ label: 'Filial Destino', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Quantidade', get: function(r) {{ return r.qtd; }} }},
    {{ label: 'Conteúdo', get: function(r) {{ return r.conteudo; }} }},
    {{ label: 'Data Solicitação', get: function(r) {{ return r.data_sol; }} }},
    {{ label: 'Status', get: function(r) {{ return r.status; }} }},
    {{ label: 'Data Postagem', get: function(r) {{ return r.data_post; }} }},
    {{ label: 'Quem Postou', get: function(r) {{ return r.quem_postou; }} }},
    {{ label: 'Baixado ADM', get: function(r) {{ return r.baixado_adm; }} }},
    {{ label: 'Observação', get: function(r) {{ return r.obs; }} }},
    {{ label: 'Horas até Postagem', get: function(r) {{ return r.horas_postagem; }} }}
  ]
}});

// ---- tabela de rastreios dos correios ----
function linhaRastreioHtml(r) {{
  var cls = r.tipo === 'FILIAL' ? 'ok' : 'warn';
  return '<tr><td>' + r.data + '</td><td>' + r.rastreio + '</td>' +
    '<td><span class="pill ' + cls + '">' + r.tipo + '</span></td>' +
    '<td>' + r.destino + '</td><td>' + r.cidade + '</td><td>' + r.produto + '</td>' +
    '<td class="num">' + brlJs(r.valor) + '</td>' +
    '<td>' + (r.entrega_cls ? '<span class="pill ' + r.entrega_cls + '">' + r.entrega_grupo + '</span>' : r.entrega_grupo) + '</td>' +
    '<td>' + r.local + '</td>' +
    '<td>' + (r.previsao_cls ? '<span class="pill ' + r.previsao_cls + '">' + r.previsao + '</span>' : r.previsao) + '</td>' +
    '<td>' + r.nfs + '</td>' +
    '<td>' + (r.situacao_cls ? '<span class="pill ' + r.situacao_cls + '">' + r.situacao + '</span>' : r.situacao) + '</td></tr>';
}}

var RASTREIOS_COLS = {{
  data: function(r) {{ return r.data_ts; }},
  rastreio: function(r) {{ return r.rastreio; }},
  tipo: function(r) {{ return r.tipo; }},
  destino: function(r) {{ return r.destino.toLowerCase(); }},
  cidade: function(r) {{ return r.cidade.toLowerCase(); }},
  produto: function(r) {{ return r.produto.toLowerCase(); }},
  valor: function(r) {{ return r.valor; }},
  nfs: function(r) {{ return r.nfs.toLowerCase(); }},
  entrega: function(r) {{ return r.entrega_grupo.toLowerCase(); }},
  local: function(r) {{ return r.local.toLowerCase(); }},
  previsao: function(r) {{ return r.previsao_ts; }},
  situacao: function(r) {{ return r.situacao.toLowerCase(); }}
}};

criarTabelaPaginada({{
  dados: {rastreios_json},
  pageSize: 50, sortInicial: 'data',
  tbodyId: 'tbody-rastreios', pagerId: 'pager-rastreios', tableId: 'tbl-rastreios',
  buscaId: 'filtro-rastreios', limparId: 'limpar-rastreios',
  colunas: RASTREIOS_COLS, linhaHtml: linhaRastreioHtml,
  busca: function(r) {{ return r.rastreio + ' ' + r.destino + ' ' + r.cidade + ' ' + r.cep + ' ' + r.nfs + ' ' + r.entrega_grupo + ' ' + r.local; }},
  filtroData: {{ campo: 'data_ts', deId: 'rastreios-data-de', ateId: 'rastreios-data-ate' }},
  filtros: [
    {{ id: 'rastreios-tipo', campo: 'tipo' }},
    {{ id: 'rastreios-destino', campo: 'destino' }},
    {{ id: 'rastreios-entrega', campo: 'entrega_grupo' }}
  ],
  colunasExport: [
    {{ label: 'Data', get: function(r) {{ return r.data; }} }},
    {{ label: 'Código', get: function(r) {{ return r.rastreio; }} }},
    {{ label: 'Tipo', get: function(r) {{ return r.tipo; }} }},
    {{ label: 'Destino', get: function(r) {{ return r.destino; }} }},
    {{ label: 'Cidade/UF', get: function(r) {{ return r.cidade; }} }},
    {{ label: 'CEP', get: function(r) {{ return r.cep; }} }},
    {{ label: 'Serviço', get: function(r) {{ return r.produto; }} }},
    {{ label: 'Frete', get: function(r) {{ return r.valor; }} }},
    {{ label: 'Status da entrega', get: function(r) {{ return r.entrega_grupo; }} }},
    {{ label: 'Onde está', get: function(r) {{ return r.local; }} }},
    {{ label: 'Previsão de entrega', get: function(r) {{ return r.previsao; }} }},
    {{ label: 'Evento dos Correios', get: function(r) {{ return r.entrega; }} }},
    {{ label: 'NF da transferência', get: function(r) {{ return r.nfs; }} }},
    {{ label: 'Situação da transferência', get: function(r) {{ return r.situacao; }} }}
  ]
}});

// ---- tabela de despesas ----
function linhaDespesaHtml(r) {{
  return '<tr><td>' + r.filial + '</td><td>' + r.categoria + '</td><td>' + r.fornecedor + '</td>' +
    '<td>' + r.documento + '</td><td>' + r.data + '</td><td class="num">' + brlJs(r.valor) + '</td>' +
    '<td>' + r.historico + '</td></tr>';
}}

var DESPESAS_COLS = {{
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  categoria: function(r) {{ return r.categoria.toLowerCase(); }},
  fornecedor: function(r) {{ return r.fornecedor.toLowerCase(); }},
  documento: function(r) {{ return r.documento.toLowerCase(); }},
  data: function(r) {{ return r.data_ts; }},
  valor: function(r) {{ return r.valor; }},
  historico: function(r) {{ return r.historico.toLowerCase(); }}
}};

function buscaDespesa(r) {{ return r.filial + ' ' + r.fornecedor + ' ' + r.documento + ' ' + r.historico; }}

criarTabelaPaginada({{
  dados: {despesas_json},
  pageSize: 50, sortInicial: 'data',
  tbodyId: 'tbody-despesas', pagerId: 'pager-despesas', tableId: 'tbl-despesas',
  buscaId: 'filtro-despesas', limparId: 'limpar-despesas',
  colunas: DESPESAS_COLS, linhaHtml: linhaDespesaHtml, busca: buscaDespesa,
  filtroData: {{ campo: 'data_ts', deId: 'despesas-data-de', ateId: 'despesas-data-ate' }},
  filtros: [
    {{ id: 'despesas-filial', campo: 'filial' }},
    {{ id: 'despesas-categoria', campo: 'categoria' }}
  ],
  colunasExport: [
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Categoria', get: function(r) {{ return r.categoria; }} }},
    {{ label: 'Fornecedor / Operação', get: function(r) {{ return r.fornecedor; }} }},
    {{ label: 'Documento', get: function(r) {{ return r.documento; }} }},
    {{ label: 'Data', get: function(r) {{ return r.data; }} }},
    {{ label: 'Valor', get: function(r) {{ return r.valor; }} }},
    {{ label: 'Histórico', get: function(r) {{ return r.historico; }} }}
  ]
}});

const CHART_DATA = {chart_data_json};
function corDoTema(nome) {{
  return getComputedStyle(document.documentElement).getPropertyValue(nome).trim();
}}
Chart.defaults.color = corDoTema('--chart-texto');
Chart.defaults.borderColor = corDoTema('--chart-grade');

// escreve o valor na ponta de cada barra. `totalNoTopoPlugin`, usado no gráfico de
// despesas, soma as séries empilhadas; este rotula cada barra individualmente, que é o
// que faz sentido em gráficos agrupados como Aprovado x Utilizado.
function rotuloPorBarra(formatar, tamanho) {{
  formatar = formatar || function(v) {{ return v; }};
  return {{
    id: 'rotuloPorBarra',
    afterDatasetsDraw: function(chart) {{
      var ctx = chart.ctx;
      var horizontal = chart.options.indexAxis === 'y';
      ctx.save();
      ctx.font = 'bold ' + (tamanho || 11) + 'px Segoe UI, Arial, sans-serif';
      ctx.fillStyle = corDoTema('--chart-rotulo');
      chart.data.datasets.forEach(function(ds, di) {{
        var meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach(function(barra, i) {{
          var valor = ds.data[i];
          if (!valor) return;
          if (horizontal) {{
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            ctx.fillText(formatar(valor), barra.x + 5, barra.y);
          }} else {{
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(formatar(valor), barra.x, barra.y - 4);
          }}
        }});
      }});
      ctx.restore();
    }}
  }};
}}
const COR_OK = '#22c55e', COR_WARN = '#f59e0b', COR_BAD = '#ef4444', COR_ACCENT = '#38bdf8';
// cinza para o que não é bom nem ruim, só não entra no fluxo de faturamento
const COR_NEUTRO = '#64748b';

// gráficos de barra horizontal com uma categoria por filial escondiam rótulos quando
// não cabiam todos na altura fixa (só apareciam no hover) — aqui a altura vira proporcional
// ao número de filiais e o eixo de categorias nunca pula rótulo (autoSkip: false)
function prepararGraficoPorCategoria(canvasId, n, porItem) {{
  var canvas = document.getElementById(canvasId);
  var box = canvas.closest('.chart-box');
  box.style.height = Math.max(320, n * (porItem || 24) + 60) + 'px';
}}

new Chart(document.getElementById('chart-pedidos-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.pedidosStatus.labels,
    datasets: [{{ data: CHART_DATA.pedidosStatus.values, backgroundColor: [COR_OK, COR_WARN, COR_BAD, COR_NEUTRO] }}]
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

prepararGraficoPorCategoria('chart-pedidos-filial', CHART_DATA.pedidosFilial.labels.length);
new Chart(document.getElementById('chart-pedidos-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.pedidosFilial.labels,
    datasets: [
      {{ label: 'Entregue', data: CHART_DATA.pedidosFilial.entregue, backgroundColor: COR_OK }},
      {{ label: 'Pendente', data: CHART_DATA.pedidosFilial.pendente, backgroundColor: COR_WARN }},
      {{ label: 'Atrasado', data: CHART_DATA.pedidosFilial.atrasado, backgroundColor: COR_BAD }},
      {{ label: 'Não faturado', data: CHART_DATA.pedidosFilial.nao_fatura, backgroundColor: COR_NEUTRO }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true, ticks: {{ autoSkip: false }} }} }},
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

prepararGraficoPorCategoria('chart-amet-filial', CHART_DATA.amet.labels.length);
new Chart(document.getElementById('chart-amet-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.amet.labels,
    datasets: [
      {{ label: 'Estoque', data: CHART_DATA.amet.estoque, backgroundColor: COR_ACCENT }},
      {{ label: 'Vendido', data: CHART_DATA.amet.vendido, backgroundColor: COR_OK }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{ legend: {{ position: 'bottom' }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

prepararGraficoPorCategoria('chart-devia-filial', CHART_DATA.devia.labels.length);
new Chart(document.getElementById('chart-devia-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.devia.labels,
    datasets: [
      {{ label: 'Estoque', data: CHART_DATA.devia.estoque, backgroundColor: COR_ACCENT }},
      {{ label: 'Vendido', data: CHART_DATA.devia.vendido, backgroundColor: COR_OK }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{ legend: {{ position: 'bottom' }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

prepararGraficoPorCategoria('chart-upmaster-filial', CHART_DATA.upmaster.labels.length);
new Chart(document.getElementById('chart-upmaster-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.upmaster.labels,
    datasets: [
      {{ label: 'Estoque', data: CHART_DATA.upmaster.estoque, backgroundColor: COR_ACCENT }},
      {{ label: 'Vendido', data: CHART_DATA.upmaster.vendido, backgroundColor: COR_OK }}
    ]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {{ legend: {{ position: 'bottom' }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

prepararGraficoPorCategoria('chart-acessorios-filial', CHART_DATA.acessoriosDiversos.labels.length);
new Chart(document.getElementById('chart-acessorios-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.acessoriosDiversos.labels,
    datasets: [{{ label: 'Saldo', data: CHART_DATA.acessoriosDiversos.saldo, backgroundColor: COR_ACCENT }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

prepararGraficoPorCategoria('chart-acessorios-tim-filial', CHART_DATA.acessoriosTim.labels.length);
new Chart(document.getElementById('chart-acessorios-tim-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.acessoriosTim.labels,
    datasets: [{{ label: 'Saldo', data: CHART_DATA.acessoriosTim.saldo, backgroundColor: COR_OK }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

prepararGraficoPorCategoria('chart-devolvidos-filial', CHART_DATA.devolvidos.labels.length);
new Chart(document.getElementById('chart-devolvidos-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.devolvidos.labels,
    datasets: [{{ label: 'Saldo (unidades)', data: CHART_DATA.devolvidos.saldo, backgroundColor: COR_BAD }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }},
    scales: {{
      x: {{ type: 'logarithmic', title: {{ display: true, text: 'Saldo (escala logarítmica — um chip com milhares de unidades em Estoque Matriz dominaria a escala linear)', font: {{ size: 10 }} }} }},
      y: {{ ticks: {{ autoSkip: false }} }}
    }}
  }}
}});

var CORES_STATUS_ADM_MALOTES = {{ 'CRÍTICO': COR_BAD, 'BAIXO': COR_WARN, 'NORMAL': COR_OK, 'EXCEDENTE': COR_ACCENT }};
new Chart(document.getElementById('chart-malotes-status-adm'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.malotesStatusAdm.labels,
    datasets: [{{
      data: CHART_DATA.malotesStatusAdm.values,
      backgroundColor: CHART_DATA.malotesStatusAdm.labels.map(function(l) {{ return CORES_STATUS_ADM_MALOTES[l] || COR_ACCENT; }})
    }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

prepararGraficoPorCategoria('chart-malotes-filial', CHART_DATA.malotesFilial.labels.length);
new Chart(document.getElementById('chart-malotes-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.malotesFilial.labels,
    datasets: [
      {{ label: 'Na Filial', data: CHART_DATA.malotesFilial.na_filial, backgroundColor: COR_ACCENT }},
      {{ label: 'No ADM', data: CHART_DATA.malotesFilial.no_adm, backgroundColor: COR_OK }}
    ]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: {{ x: {{ stacked: true }}, y: {{ stacked: true, ticks: {{ autoSkip: false }} }} }}, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

// a cor vem classificada do Python, para nao existir uma segunda lista de status aqui
var COR_POR_CLASSE = {{ ok: COR_OK, warn: COR_WARN, bad: COR_BAD }};
new Chart(document.getElementById('chart-rastreios-entrega'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.rastreiosEntrega.labels,
    datasets: [{{
      data: CHART_DATA.rastreiosEntrega.values,
      backgroundColor: CHART_DATA.rastreiosEntrega.classes.map(function(c) {{ return COR_POR_CLASSE[c] || COR_NEUTRO; }})
    }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-rastreios-destino'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.rastreiosDestino.labels,
    datasets: [{{ data: CHART_DATA.rastreiosDestino.values, backgroundColor: [COR_OK, COR_WARN] }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

prepararGraficoPorCategoria('chart-rastreios-filial', CHART_DATA.rastreiosFilial.labels.length);
new Chart(document.getElementById('chart-rastreios-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.rastreiosFilial.labels,
    datasets: [{{ label: 'Postagens', data: CHART_DATA.rastreiosFilial.values, backgroundColor: COR_ACCENT }}]
  }},
  plugins: [rotuloPorBarra()],
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }} }},
    scales: {{ x: {{ beginAtZero: true, grace: '12%', ticks: {{ precision: 0 }} }}, y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

new Chart(document.getElementById('chart-obras-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.obrasFilial.labels,
    datasets: [{{ label: 'Obras', data: CHART_DATA.obrasFilial.values, backgroundColor: COR_WARN }}]
  }},
  plugins: [rotuloPorBarra()],
  options: {{
    responsive: true, maintainAspectRatio: false,
    layout: {{ padding: {{ top: 18 }} }},
    plugins: {{ legend: {{ display: false }} }},
    scales: {{ y: {{ beginAtZero: true, grace: '12%', ticks: {{ precision: 0 }} }} }}
  }}
}});

new Chart(document.getElementById('chart-obras-responsavel'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.obrasResponsavel.labels,
    datasets: [{{ label: 'Obras', data: CHART_DATA.obrasResponsavel.values, backgroundColor: COR_ACCENT }}]
  }},
  plugins: [rotuloPorBarra()],
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }} }},
    scales: {{ x: {{ beginAtZero: true, grace: '12%', ticks: {{ precision: 0 }} }}, y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

new Chart(document.getElementById('chart-manutencoes-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.manutencoesStatus.labels,
    datasets: [{{
      data: CHART_DATA.manutencoesStatus.values,
      backgroundColor: CHART_DATA.manutencoesStatus.classes.map(function(c) {{ return COR_POR_CLASSE[c] || COR_ACCENT; }})
    }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

prepararGraficoPorCategoria('chart-manutencoes-filial', CHART_DATA.manutencoesFilial.labels.length);
new Chart(document.getElementById('chart-manutencoes-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.manutencoesFilial.labels,
    datasets: [{{ label: 'Chamados', data: CHART_DATA.manutencoesFilial.values, backgroundColor: COR_ACCENT }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

new Chart(document.getElementById('chart-despesas-categoria'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.despesasCategoria.labels,
    datasets: [{{ data: CHART_DATA.despesasCategoria.values, backgroundColor: [COR_ACCENT, COR_WARN, COR_OK] }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'bottom' }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.label + ': ' + brlJs(ctx.parsed); }} }} }}
    }}
  }}
}});

prepararGraficoPorCategoria('chart-despesas-filial', CHART_DATA.despesasFilial.labels.length);
new Chart(document.getElementById('chart-despesas-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.despesasFilial.labels,
    datasets: [{{ label: 'Total pago', data: CHART_DATA.despesasFilial.values, backgroundColor: COR_ACCENT }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return brlJs(ctx.parsed.x); }} }} }}
    }},
    scales: {{ y: {{ ticks: {{ autoSkip: false }} }} }}
  }}
}});

var totalNoTopoPlugin = {{
  id: 'totalNoTopo',
  afterDatasetsDraw: function(chart) {{
    var ctx = chart.ctx;
    var meta0 = chart.getDatasetMeta(0);
    ctx.save();
    ctx.font = 'bold 12px Segoe UI, Arial, sans-serif';
    ctx.fillStyle = corDoTema('--chart-rotulo');
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    chart.data.labels.forEach(function(label, i) {{
      var total = 0;
      chart.data.datasets.forEach(function(ds) {{ total += ds.data[i] || 0; }});
      var bar = meta0.data[i];
      var yScale = chart.scales.y;
      var xPos = bar.x;
      var yPos = yScale.getPixelForValue(total);
      ctx.fillText(brlJs(total), xPos, yPos - 6);
    }});
    ctx.restore();
  }}
}};

new Chart(document.getElementById('chart-despesas-mensal'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.despesasMensal.labels,
    datasets: [
      {{ label: 'Material de Limpeza', data: CHART_DATA.despesasMensal.limpeza, backgroundColor: COR_ACCENT }},
      {{ label: 'Material de Escritório', data: CHART_DATA.despesasMensal.escritorio, backgroundColor: COR_WARN }},
      {{ label: 'Registro Manual', data: CHART_DATA.despesasMensal.manual, backgroundColor: COR_OK }}
    ]
  }},
  plugins: [totalNoTopoPlugin],
  options: {{
    responsive: true, maintainAspectRatio: false,
    layout: {{ padding: {{ top: 24 }} }},
    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true, grace: '10%' }} }},
    plugins: {{
      legend: {{ position: 'bottom' }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.dataset.label + ': ' + brlJs(ctx.parsed.y); }} }} }}
    }}
  }}
}});

var PALETA_FILIAIS = ['#38bdf8', '#fb7185', '#4ade80', '#fbbf24', '#a78bfa', '#22d3ee', '#f472b6', '#fb923c', '#818cf8', '#a3e635'];
new Chart(document.getElementById('chart-despesas-filial-mensal'), {{
  type: 'line',
  data: {{
    labels: CHART_DATA.despesasFilialMensal.labels,
    datasets: CHART_DATA.despesasFilialMensal.series.map(function(s, i) {{
      return {{
        label: s.filial, data: s.valores, borderColor: PALETA_FILIAIS[i % PALETA_FILIAIS.length],
        backgroundColor: PALETA_FILIAIS[i % PALETA_FILIAIS.length], borderWidth: 2.5,
        tension: 0, pointRadius: 3, pointHoverRadius: 6, pointBackgroundColor: corDoTema('--chart-ponto'), pointBorderWidth: 2
      }};
    }})
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    interaction: {{ mode: 'index', intersect: false }},
    plugins: {{
      legend: {{ position: 'bottom' }},
      tooltip: {{
        itemSort: function(a, b) {{ return b.parsed.y - a.parsed.y; }},
        callbacks: {{ label: function(ctx) {{ return ctx.dataset.label + ': ' + brlJs(ctx.parsed.y); }} }}
      }}
    }},
    scales: {{ y: {{ ticks: {{ callback: function(v) {{ return brlJs(v); }} }} }} }}
  }}
}});

// Aprovado fica com o azul, então os tipos usam as outras cores
var CORES_TIPO_CAMPANHA = {{
  'Campanha': COR_WARN, 'Verba de Área': COR_OK, 'Não separado': COR_NEUTRO, 'Outros': '#a78bfa'
}};
var TIPOS_CAMPANHA_ORDEM = ['Campanha', 'Verba de Área', 'Não separado', 'Outros'];

// escreve o total de cada pilha no topo dela. o totalNoTopoPlugin das despesas soma
// todos os datasets; aqui há duas pilhas por área (Aprovado e Utilizado), então o total
// precisa ser por grupo de empilhamento.
var totalPorPilhaPlugin = {{
  id: 'totalPorPilha',
  afterDatasetsDraw: function(chart) {{
    var ctx = chart.ctx;
    var pilhas = {{}};
    chart.data.datasets.forEach(function(ds, di) {{
      var meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      var nome = ds.stack || ('ds' + di);
      var p = pilhas[nome] || (pilhas[nome] = {{ totais: [], x: [] }});
      meta.data.forEach(function(barra, i) {{
        p.totais[i] = (p.totais[i] || 0) + (ds.data[i] || 0);
        p.x[i] = barra.x;
      }});
    }});
    ctx.save();
    ctx.font = 'bold 10px Segoe UI, Arial, sans-serif';
    ctx.fillStyle = corDoTema('--chart-rotulo');
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    var yScale = chart.scales.y;
    Object.keys(pilhas).forEach(function(nome) {{
      var p = pilhas[nome];
      p.totais.forEach(function(total, i) {{
        if (!total) return;
        ctx.fillText(brlJs(total), p.x[i], yScale.getPixelForValue(total) - 4);
      }});
    }});
    ctx.restore();
  }}
}};

// os dados entram em renderCampanhas, para acompanhar o filtro de datas
var chartCampanhasTipo = new Chart(document.getElementById('chart-campanhas-tipo'), {{
  type: 'doughnut',
  data: {{ labels: [], datasets: [{{ data: [], backgroundColor: [] }}] }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'bottom' }},
      subtitle: {{ display: true, text: '', color: corDoTema('--chart-texto'), font: {{ size: 11 }}, padding: {{ bottom: 10 }} }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.label + ': ' + brlJs(ctx.parsed); }} }} }}
    }}
  }}
}});

var chartCampanhasArea = new Chart(document.getElementById('chart-campanhas-area'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.campanhasArea.labels,
    datasets: [
      {{ label: 'Aprovado', data: CHART_DATA.campanhasArea.verba, backgroundColor: COR_ACCENT, stack: 'aprovado' }},
      {{ label: 'Utilizado', data: CHART_DATA.campanhasArea.gasto, backgroundColor: COR_WARN, stack: 'utilizado' }}
    ]
  }},
  plugins: [totalPorPilhaPlugin],
  options: {{
    responsive: true, maintainAspectRatio: false,
    layout: {{ padding: {{ top: 20 }} }},
    plugins: {{
      legend: {{ position: 'bottom' }},
      // preenchido por renderCampanhas: assim o PNG exportado sai com o período junto
      subtitle: {{ display: true, text: '', color: corDoTema('--chart-texto'), font: {{ size: 11 }}, padding: {{ bottom: 10 }} }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.dataset.label + ': ' + brlJs(ctx.parsed.y); }} }} }}
    }},
    scales: {{
      x: {{ stacked: true }},
      y: {{ stacked: true, grace: '12%', ticks: {{ callback: function(v) {{ return brlJs(v); }} }} }}
    }}
  }}
}});

// 34 filiais com duas barras cada precisam de mais folga vertical para o rótulo caber
prepararGraficoPorCategoria('chart-campanhas-filial', CHART_DATA.campanhasFilial.labels.length, 34);
var chartCampanhasFilial = new Chart(document.getElementById('chart-campanhas-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.campanhasFilial.labels,
    datasets: [
      {{ label: 'Aprovado', data: CHART_DATA.campanhasFilial.verba, backgroundColor: COR_ACCENT }},
      {{ label: 'Utilizado', data: CHART_DATA.campanhasFilial.gasto, backgroundColor: COR_WARN }}
    ]
  }},
  plugins: [rotuloPorBarra(brlJs, 10)],
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ position: 'bottom' }},
      subtitle: {{ display: true, text: '', color: corDoTema('--chart-texto'), font: {{ size: 11 }}, padding: {{ bottom: 10 }} }},
      tooltip: {{ callbacks: {{ label: function(ctx) {{ return ctx.dataset.label + ': ' + brlJs(ctx.parsed.x); }} }} }}
    }},
    scales: {{
      x: {{ grace: '18%', ticks: {{ callback: function(v) {{ return brlJs(v); }} }} }},
      y: {{ ticks: {{ autoSkip: false }} }}
    }}
  }}
}});
// ---- campanhas: o filtro de datas recalcula cards, tabelas e graficos ----
// a verba da planilha e MENSAL e os lancamentos cobrem o ano todo, entao a verba do
// periodo e verba_mensal * numero de meses do intervalo. sem isso o "% utilizado"
// comparava gasto de 8 meses contra verba de 1 mes.
var CAMPANHAS_LANC = {campanhas_lancamentos_json};
var CAMPANHAS_VERBA = {campanhas_verba_json};
var CAMPANHAS_VERBA_MENSAL = CAMPANHAS_VERBA.reduce(function(s, f) {{ return s + f.verba; }}, 0);
var CAMPANHAS_TEM_VERBA = {{}};
CAMPANHAS_VERBA.forEach(function(f) {{ CAMPANHAS_TEM_VERBA[f.filial] = true; }});

var CAMP_TS_MIN = null, CAMP_TS_MAX = null;
CAMPANHAS_LANC.forEach(function(r) {{
  if (!r.data_ts) return;
  if (CAMP_TS_MIN === null || r.data_ts < CAMP_TS_MIN) CAMP_TS_MIN = r.data_ts;
  if (CAMP_TS_MAX === null || r.data_ts > CAMP_TS_MAX) CAMP_TS_MAX = r.data_ts;
}});

function round2(v) {{ return Math.round(v * 100) / 100; }}

function limitesPeriodo(deTs, ateTs) {{
  return [
    (deTs === null || deTs === undefined) ? CAMP_TS_MIN : deTs,
    (ateTs === null || ateTs === undefined) ? CAMP_TS_MAX : ateTs
  ];
}}

// conta meses de calendario tocados pelo intervalo: jan a jan = 1, jan a mar = 3
function mesesNoPeriodo(deTs, ateTs) {{
  var lim = limitesPeriodo(deTs, ateTs);
  if (lim[0] === null || lim[1] === null) return 1;
  var ini = new Date(lim[0]), fim = new Date(lim[1]);
  var n = (fim.getFullYear() - ini.getFullYear()) * 12 + (fim.getMonth() - ini.getMonth()) + 1;
  return n > 0 ? n : 1;
}}

function rotuloPeriodo(deTs, ateTs, nMeses) {{
  var lim = limitesPeriodo(deTs, ateTs);
  if (lim[0] === null || lim[1] === null) return 'sem lançamentos no período';
  function fmt(ts) {{
    var d = new Date(ts);
    return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
  }}
  return 'Período de ' + fmt(lim[0]) + ' até ' + fmt(lim[1]) +
    ' (' + nMeses + (nMeses === 1 ? ' mês' : ' meses') + ')';
}}

function renderCampanhas(filtrados, deTs, ateTs) {{
  var nMeses = mesesNoPeriodo(deTs, ateTs);

  var gastoPorFilial = {{}};
  var gastoPorFilialTipo = {{}};
  filtrados.forEach(function(r) {{
    gastoPorFilial[r.filial] = (gastoPorFilial[r.filial] || 0) + r.valor;
    var porTipoDaFilial = gastoPorFilialTipo[r.filial] || (gastoPorFilialTipo[r.filial] = {{}});
    porTipoDaFilial[r.tipo] = (porTipoDaFilial[r.tipo] || 0) + r.valor;
  }});

  var linhas = CAMPANHAS_VERBA.map(function(f) {{
    return {{
      filial: f.filial, area: f.area, area_num: f.area_num,
      verba: round2(f.verba * nMeses), gasto: round2(gastoPorFilial[f.filial] || 0)
    }};
  }});
  // filial que aparece so nos lancamentos entra com verba 0 em vez de sumir do total
  Object.keys(gastoPorFilial).forEach(function(fil) {{
    if (!CAMPANHAS_TEM_VERBA[fil]) {{
      linhas.push({{ filial: fil, area: 'sem valor aprovado', area_num: '-', verba: 0, gasto: round2(gastoPorFilial[fil]) }});
    }}
  }});

  linhas.forEach(function(l) {{
    l.saldo = round2(l.verba - l.gasto);
    l.pct = l.verba ? Math.round(l.gasto / l.verba * 1000) / 10 : null;
    if (l.pct === null) l.cls = l.gasto ? 'bad' : 'ok';
    else if (l.pct >= 90) l.cls = 'bad';
    else if (l.pct >= 60) l.cls = 'warn';
    else l.cls = 'ok';
  }});
  linhas.sort(function(a, b) {{
    var av = a.pct === null ? Infinity : a.pct;
    var bv = b.pct === null ? Infinity : b.pct;
    return bv - av;
  }});

  var verbaTotal = round2(CAMPANHAS_VERBA_MENSAL * nMeses);
  var gastoTotal = round2(linhas.reduce(function(s, l) {{ return s + l.gasto; }}, 0));
  var comGasto = Object.keys(gastoPorFilial).filter(function(f) {{ return gastoPorFilial[f]; }}).length;

  document.getElementById('card-camp-verba-label').textContent =
    'Aprovado no período (' + nMeses + (nMeses === 1 ? ' mês' : ' meses') + ')';
  document.getElementById('card-camp-verba').textContent = brlJs(verbaTotal);
  document.getElementById('card-camp-gasto').textContent = brlJs(gastoTotal);
  document.getElementById('card-camp-saldo').textContent = brlJs(round2(verbaTotal - gastoTotal));
  document.getElementById('card-camp-pct').textContent =
    (verbaTotal ? Math.round(gastoTotal / verbaTotal * 1000) / 10 : 0) + '%';
  document.getElementById('card-camp-lanc').textContent = filtrados.length;
  document.getElementById('card-camp-filiais').textContent = comGasto + ' de ' + CAMPANHAS_VERBA.length;
  var periodo = rotuloPeriodo(deTs, ateTs, nMeses);
  document.getElementById('campanhas-periodo-label').textContent = periodo;

  document.getElementById('tbody-campanhas-filiais').innerHTML = linhas.map(function(l) {{
    return '<tr><td>' + l.filial + '</td><td>' + l.area + '</td>' +
      '<td class="num">' + brlJs(l.verba) + '</td><td class="num">' + brlJs(l.gasto) + '</td>' +
      '<td class="num">' + brlJs(l.saldo) + '</td>' +
      '<td><span class="pill ' + l.cls + '">' + (l.pct === null ? '—' : l.pct + '%') + '</span></td></tr>';
  }}).join('');
  // o filtro de texto dessa tabela esconde linhas no DOM, entao precisa ser reaplicado
  var inpFiliais = document.querySelector('input.filtro[data-target="tbl-campanhas-filiais"]');
  if (inpFiliais && inpFiliais.value) inpFiliais.dispatchEvent(new Event('input'));

  var porTipo = {{}};
  filtrados.forEach(function(r) {{ porTipo[r.tipo] = (porTipo[r.tipo] || 0) + r.valor; }});
  var tipos = Object.keys(porTipo).sort();
  chartCampanhasTipo.data.labels = tipos;
  chartCampanhasTipo.data.datasets[0].data = tipos.map(function(t) {{ return round2(porTipo[t]); }});
  chartCampanhasTipo.data.datasets[0].backgroundColor = tipos.map(function(t) {{ return CORES_TIPO_CAMPANHA[t] || COR_NEUTRO; }});
  chartCampanhasTipo.options.plugins.subtitle.text = periodo;
  chartCampanhasTipo.update();

  // a área de cada filial vem do orçamento (linhas), não do lançamento, para o
  // rateio bater com a coluna Aprovado
  var porArea = {{}};
  linhas.forEach(function(l) {{
    var b = porArea[l.area_num];
    if (!b) b = porArea[l.area_num] = {{ verba: 0, label: l.area, tipos: {{}} }};
    b.verba += l.verba;
    var porTipoDaFilial = gastoPorFilialTipo[l.filial] || {{}};
    Object.keys(porTipoDaFilial).forEach(function(t) {{
      b.tipos[t] = (b.tipos[t] || 0) + porTipoDaFilial[t];
    }});
  }});
  var areas = Object.keys(porArea).sort();
  var tiposPresentes = TIPOS_CAMPANHA_ORDEM.filter(function(t) {{
    return areas.some(function(a) {{ return porArea[a].tipos[t]; }});
  }});
  chartCampanhasArea.data.labels = areas.map(function(a) {{ return porArea[a].label; }});
  chartCampanhasArea.data.datasets = [
    {{ label: 'Aprovado', stack: 'aprovado', backgroundColor: COR_ACCENT,
      data: areas.map(function(a) {{ return round2(porArea[a].verba); }}) }}
  ].concat(tiposPresentes.map(function(t) {{
    return {{
      label: t, stack: 'utilizado', backgroundColor: CORES_TIPO_CAMPANHA[t] || COR_NEUTRO,
      data: areas.map(function(a) {{ return round2(porArea[a].tipos[t] || 0); }})
    }};
  }}));
  chartCampanhasArea.options.plugins.subtitle.text = periodo;
  chartCampanhasArea.update();

  prepararGraficoPorCategoria('chart-campanhas-filial', linhas.length, 34);
  chartCampanhasFilial.data.labels = linhas.map(function(l) {{ return l.filial; }});
  chartCampanhasFilial.data.datasets[0].data = linhas.map(function(l) {{ return l.verba; }});
  chartCampanhasFilial.data.datasets[1].data = linhas.map(function(l) {{ return l.gasto; }});
  chartCampanhasFilial.options.plugins.subtitle.text = periodo;
  chartCampanhasFilial.update();
}}

function linhaCampanhaLogHtml(r) {{
  return '<tr><td>' + r.filial + '</td><td>Área ' + r.area_num + '</td><td>' + r.tipo + '</td><td>' + r.data + '</td>' +
    '<td>' + r.documento + '</td><td class="num">' + brlJs(r.valor) + '</td><td>' + r.historico + '</td></tr>';
}}

var CAMPANHAS_COLS = {{
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  area: function(r) {{ return r.area_num; }},
  tipo: function(r) {{ return r.tipo.toLowerCase(); }},
  data: function(r) {{ return r.data_ts; }},
  documento: function(r) {{ return r.documento.toLowerCase(); }},
  valor: function(r) {{ return r.valor; }},
  historico: function(r) {{ return r.historico.toLowerCase(); }}
}};

criarTabelaPaginada({{
  dados: CAMPANHAS_LANC,
  pageSize: 50, sortInicial: 'data',
  tbodyId: 'tbody-campanhas-log', pagerId: 'pager-campanhas-log', tableId: 'tbl-campanhas-log',
  buscaId: 'filtro-campanhas', limparId: 'limpar-campanhas',
  colunas: CAMPANHAS_COLS, linhaHtml: linhaCampanhaLogHtml,
  busca: function(r) {{ return r.filial + ' ' + r.documento + ' ' + r.historico + ' ' + r.tipo; }},
  filtroData: {{ campo: 'data_ts', deId: 'campanhas-data-de', ateId: 'campanhas-data-ate' }},
  filtros: [
    {{ id: 'campanhas-filial', campo: 'filial' }},
    {{ id: 'campanhas-area', campo: 'area_num',
      pares: valoresUnicos(CAMPANHAS_LANC, 'area_num').map(function(a) {{ return [a, 'Área ' + a]; }}) }},
    {{ id: 'campanhas-tipo', campo: 'tipo' }}
  ],
  aoFiltrar: renderCampanhas,
  colunasExport: [
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Área', get: function(r) {{ return 'Área ' + r.area_num; }} }},
    {{ label: 'Tipo', get: function(r) {{ return r.tipo; }} }},
    {{ label: 'Data', get: function(r) {{ return r.data; }} }},
    {{ label: 'Documento', get: function(r) {{ return r.documento; }} }},
    {{ label: 'Valor', get: function(r) {{ return r.valor; }} }},
    {{ label: 'Histórico', get: function(r) {{ return r.historico; }} }}
  ]
}});
</script>
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(
    f"OK: {OUT} gerado com {len(notas_atrasadas)} notas atrasadas, {len(transf_todas)} transferências pendentes "
    f"({len(transf_criticas)} críticas), {len(acessorios_diversos_itens)} itens de acessórios, "
    f"{len(seriais_itens)} peças com serial de acessórios fidelizados TIM, "
    f"{len(devolvidos_itens)} itens devolvidos/com defeito e {malotes_log_resumo['total']} solicitações de malotes "
    f"({malotes_log_resumo['pendentes']} pendentes)."
)
