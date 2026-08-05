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

BASE = Path(r"I:\Meu Drive\CONTROLE COMPRAS")
OUT = Path(__file__).resolve().parent / "index.html"


def load(arquivo, aba):
    wb = openpyxl.load_workbook(BASE / arquivo, data_only=True, read_only=True)
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
    ts = (BASE / arquivo).stat().st_mtime
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

# tabela detalhada de pedidos (igual a "TABELA DE PEDIDOS REALIZADOS NO GN" do BI)
pedidos_detalhe = load("CONTROLE DE PEDIDOS.xlsx", "CONTROLE PEDIDOS")[1:]
pedidos_detalhe.sort(key=lambda r: r[3] or datetime.min, reverse=True)

# --------------------------------------------------------- NOTAS FISCAIS
notas_mtime = mtime_str("CONTROLE NOTAS FISCAIS PENDENTES DE ENTRADA.xlsx")
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
transf_mtime = mtime_str("TRANSFERÊNCIAS PENDENTES.xlsx")
rows = load("TRANSFERÊNCIAS PENDENTES.xlsx", "TRANSFERENCIAS PENDENTES")
data_rows = rows[1:]

transf_resumo = {}
for r in data_rows:
    label = r[14]
    if label and label != "RESUMO GERAL":
        transf_resumo[label] = r[15]

transf_criticas = [r for r in data_rows if r[11] == "CRÍTICO"]

transf_todas = [r for r in data_rows if r[0] is not None]
transf_todas.sort(key=lambda r: (r[2] or 0), reverse=True)

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


def montar_acessorios(rows):
    itens = []
    saldo_por_filial = {}
    for r in rows:
        filial = r.get("Filial") or ""
        saldo = r.get("Saldo") or 0
        valor = r.get("VALOR DE VENDA") or 0
        itens.append({
            "filial": filial, "ref": r.get("Produto") or "", "desc": r.get("Descrição") or "",
            "subgrupo": r.get("Sub Grupo Estoque") or "-", "fabricante": r.get("Fabricante") or "-",
            "saldo": saldo, "disponivel": r.get("Disponível") or 0,
            "valor": valor, "valor_total": round(saldo * valor, 2),
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

MANUTENCAO_STATUS_CLASSE = {
    "PENDENTE": "warn", "EM ANDAMENTO": "warn",
    "CONCLUÍDO": "ok", "CONCLUIDO": "ok", "CONCLUÍDA": "ok", "CONCLUIDA": "ok",
    "CANCELADO": "bad", "CANCELADA": "bad",
}

manutencoes_itens = []
for r in load_dicts("CONTROLE DE MANUTENÇÕES.xlsx", "CONTROLE MANUTENÇÕES"):
    chamado = r.get("N° DO CHAMADO")
    if not chamado:
        continue
    status = str(r.get("STATUS") or "-").strip()
    dias = r.get("DIAS SEM CONCLUSÃO")
    manutencoes_itens.append({
        "chamado": chamado, "filial": r.get("FILIAL") or "-",
        "solicitado": r.get("MANUTENÇÕES SOLICITADAS") or "-",
        "status": status, "status_cls": MANUTENCAO_STATUS_CLASSE.get(status.upper(), ""),
        "dias": dias if isinstance(dias, (int, float)) else "-",
        "obs": r.get("OBSERVAÇÃO") or "-",
    })
manutencoes_itens.sort(key=lambda x: x["dias"] if isinstance(x["dias"], (int, float)) else -1, reverse=True)

manutencoes_por_filial = {}
for it in manutencoes_itens:
    manutencoes_por_filial[it["filial"]] = manutencoes_por_filial.get(it["filial"], 0) + 1
manutencoes_filiais_ordenadas = sorted(manutencoes_por_filial.items(), key=lambda kv: kv[1], reverse=True)

manutencoes_status_buckets = {}
for it in manutencoes_itens:
    manutencoes_status_buckets[it["status"]] = manutencoes_status_buckets.get(it["status"], 0) + 1

dias_manutencoes_validos = [x["dias"] for x in manutencoes_itens if isinstance(x["dias"], (int, float))]
manutencoes_resumo = {
    "total": len(manutencoes_itens),
    "pendentes": sum(1 for x in manutencoes_itens if x["status"].upper() == "PENDENTE"),
    "concluidas": sum(1 for x in manutencoes_itens if x["status"].upper() in ("CONCLUÍDO", "CONCLUIDO", "CONCLUÍDA", "CONCLUIDA")),
    "filiais": len(manutencoes_por_filial),
    "dias_medio": round(sum(dias_manutencoes_validos) / len(dias_manutencoes_validos), 1) if dias_manutencoes_validos else 0,
}

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
    },
    "manutencoesFilial": {
        "labels": [f[0] for f in manutencoes_filiais_ordenadas],
        "values": [f[1] for f in manutencoes_filiais_ordenadas],
    },
}
chart_data_json = json.dumps(chart_data, ensure_ascii=False)

# ----------------------------------------------------------------- HTML
STATUS_CLASSE = {"ENTREGUE": "ok", "PENDENTE DE ENTRADA": "warn", "ATRASADO": "bad"}
CRITICIDADE_CLASSE = {"OK": "ok", "ATENÇÃO": "warn", "CRÍTICO": "bad"}


def epoch(dt):
    return int(dt.timestamp() * 1000) if dt else 0


pedidos_detalhe_data = []
for r in pedidos_detalhe:
    status = r[11] or ""
    dias_num = r[12] if isinstance(r[12], (int, float)) else -1
    pedidos_detalhe_data.append({
        "dreal": data_str(r[3]), "dreal_ts": epoch(r[3]),
        "dproc": data_str(r[4]), "dproc_ts": epoch(r[4]),
        "dprev": data_str(r[8]), "dprev_ts": epoch(r[8]),
        "fil": (r[1] or "").replace("ROCHA TELECOM - ", ""),
        "ped": r[2], "qtd": r[6], "desc": r[5], "stprod": r[7] or "",
        "dias": r[12] if isinstance(r[12], (int, float)) else (r[12] or "-"), "dias_num": dias_num,
        "entrada": r[14] or "", "status": status, "cls": STATUS_CLASSE.get(status, ""),
    })
pedidos_detalhe_json = json.dumps(pedidos_detalhe_data, ensure_ascii=False)


def linhas_pedidos_filiais():
    out = []
    for f in pedidos_filiais:
        pct_entrega = f["entregue"] / f["total"] if f["total"] else 0
        out.append(
            "<tr><td>{fil}</td><td class='num'>{tot}</td><td class='num'>{ent}</td>"
            "<td class='num'>{pen}</td><td class='num'>{atr}</td><td class='num'>{pct}</td></tr>".format(
                fil=f["filial"], tot=f["total"], ent=f["entregue"], pen=f["pendente"],
                atr=f["atrasado"], pct=pct(pct_entrega)
            )
        )
    return "\n".join(out)


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


transf_todas_data = []
for r in transf_todas:
    crit = r[11] or ""
    prazo = r[12] or ""
    dias_num = round(r[2] or 0)
    transf_todas_data.append({
        "orig": r[0] or "", "dias": dias_num, "dias_num": dias_num, "dest": r[3] or "",
        "user": r[4] or "", "nf": r[5] or "", "prod": r[7] or "", "desc": r[8] or "",
        "qtd": r[9], "crit": crit, "cls": CRITICIDADE_CLASSE.get(crit, ""), "prazo": prazo,
    })
transf_todas_json = json.dumps(transf_todas_data, ensure_ascii=False)


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
  <h2>📮 Controle de Malotes</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de Malotes (Parque)</div><div class="value">{resumo_filiais['total_parque']}</div></div>
    <div class="card ok"><div class="label">No ADM</div><div class="value">{resumo_filiais['no_adm']}</div></div>
    <div class="card ok"><div class="label">Nas Filiais</div><div class="value">{resumo_filiais['nas_filiais']}</div></div>
    <div class="card bad"><div class="label">Filiais sem Malote no ADM</div><div class="value">{resumo_filiais['sem_malote_adm']}</div></div>
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
<title>BI Operacional — Rocha Telecom</title>
<link rel="icon" type="image/svg+xml" href="{FAVICON_HREF}">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {{
    --bg: #0f172a; --card: #1e293b; --card2: #16213a; --text: #e2e8f0; --muted: #94a3b8;
    --ok: #22c55e; --warn: #f59e0b; --bad: #ef4444; --accent: #38bdf8; --border: #334155;
  }}
  * {{ box-sizing: border-box; }}
  html {{ scroll-behavior: smooth; }}
  body {{ margin:0; font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); }}
  .layout {{ display: flex; align-items: flex-start; }}
  .sidebar {{ width: 220px; flex-shrink: 0; position: sticky; top: 0; height: 100vh; overflow-y: auto; background: var(--card); border-right: 1px solid var(--border); padding: 20px 0; }}
  .sidebar .brand {{ padding: 0 20px 16px; font-size: 14px; font-weight: 700; color: var(--text); border-bottom: 1px solid var(--border); margin-bottom: 8px; }}
  .sidebar a {{ display: flex; align-items: flex-start; gap: 8px; padding: 12px 20px; color: var(--muted); text-decoration: none; font-size: 14px; line-height: 1.35; border-left: 3px solid transparent; }}
  .sidebar a .nav-icon {{ flex-shrink: 0; }}
  .sidebar a .nav-label {{ flex: 1; }}
  .sidebar a:hover {{ background: #22314f; color: var(--text); }}
  .sidebar a.active {{ color: var(--accent); border-left-color: var(--accent); background: #16213a; font-weight: 600; }}
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
  th {{ background: #0b1120; color: var(--muted); font-size: 11px; text-transform: uppercase; position: sticky; top: 0; }}
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
  tr:hover td {{ background: #22314f; }}
  .table-wrap {{ max-height: 480px; overflow: auto; border: 1px solid var(--border); border-radius: 10px; }}
  .charts {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; margin-bottom: 20px; }}
  .chart-box {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; position: relative; height: 320px; }}
  .chart-box h4 {{ margin: 0 0 12px; font-size: 13px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }}
  .chart-box .canvas-wrap {{ position: relative; height: calc(100% - 28px); }}
  .chart-box.wide {{ grid-column: 1 / -1; height: 720px; }}
  input.filtro {{ width: 100%; max-width: 320px; padding: 8px 10px; margin-bottom: 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); }}
  .filtros-pedidos {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }}
  .filtros-pedidos input.filtro {{ margin-bottom: 0; flex: 1 1 240px; }}
  .filtros-pedidos select {{ padding: 8px 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); flex: 1 1 170px; }}
  .filtros-pedidos button {{ padding: 8px 14px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--muted); cursor: pointer; }}
  .filtros-pedidos button:hover {{ color: var(--text); border-color: var(--accent); }}
  .msel {{ position: relative; flex: 1 1 170px; }}
  .msel-btn {{ width: 100%; text-align: left; padding: 8px 26px 8px 10px; background: var(--card); border: 1px solid var(--border); border-radius: 6px; color: var(--text); cursor: pointer; position: relative; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .msel-btn:after {{ content: '▾'; position: absolute; right: 10px; top: 50%; transform: translateY(-50%); color: var(--muted); }}
  .msel.open .msel-btn {{ border-color: var(--accent); }}
  .msel-panel {{ display: none; position: absolute; top: calc(100% + 4px); left: 0; min-width: 240px; max-height: 280px; overflow-y: auto; background: var(--card2); border: 1px solid var(--border); border-radius: 8px; padding: 8px; z-index: 30; box-shadow: 0 8px 24px rgba(0,0,0,.5); }}
  .msel.open .msel-panel {{ display: block; }}
  .msel-actions {{ display: flex; gap: 8px; margin-bottom: 6px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }}
  .msel-actions button {{ font-size: 11px; background: none; border: 1px solid var(--border); color: var(--muted); border-radius: 4px; padding: 3px 8px; cursor: pointer; }}
  .msel-actions button:hover {{ color: var(--text); }}
  .msel-options label {{ display: flex; align-items: center; gap: 7px; padding: 5px 4px; font-size: 13px; cursor: pointer; border-radius: 4px; }}
  .msel-options label:hover {{ background: #22314f; }}
  footer {{ text-align: center; color: var(--muted); font-size: 12px; padding: 24px; }}
</style>
</head>
<body>
<div class="layout">
<nav class="sidebar">
  <div class="brand">📊 BI Operacional</div>
  <a href="#pedidos" class="nav-link"><span class="nav-icon">📦</span><span class="nav-label">Pedidos</span></a>
  <a href="#notas" class="nav-link"><span class="nav-icon">🧾</span><span class="nav-label">Notas Fiscais</span></a>
  <a href="#transferencias" class="nav-link"><span class="nav-icon">🔄</span><span class="nav-label">Transferências</span></a>
  <a href="#amet" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas AMET</span></a>
  <a href="#devia" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas DEVIA</span></a>
  <a href="#upmaster" class="nav-link"><span class="nav-icon">🛡️</span><span class="nav-label">Películas UPMASTER</span></a>
  <a href="#manutencoes" class="nav-link"><span class="nav-icon">🔧</span><span class="nav-label">Manutenções</span></a>
  <a href="#acessorios" class="nav-link"><span class="nav-icon">🎧</span><span class="nav-label">Acessórios</span></a>
  <a href="#acessorios-tim" class="nav-link"><span class="nav-icon">📶</span><span class="nav-label">Fidelizados TIM</span></a>
  <a href="#devolvidos" class="nav-link"><span class="nav-icon">♻️</span><span class="nav-label">Devolvidos e Defeitos</span></a>
  <a href="#malotes" class="nav-link"><span class="nav-icon">📮</span><span class="nav-label">Controle de Malotes</span></a>
</nav>
<div class="content">
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

<section id="pedidos">
  <h2>📦 Pedidos GN — Status Geral</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {pedidos_mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Total de pedidos</div><div class="value">{pedidos['total']}</div></div>
    <div class="card ok"><div class="label">Entregue</div><div class="value">{pedidos['entregue'][0]}</div><div class="sub">{pct(pedidos['entregue'][1])}</div></div>
    <div class="card warn"><div class="label">Pendente de entrada</div><div class="value">{pedidos['pendente'][0]}</div><div class="sub">{pct(pedidos['pendente'][1])}</div></div>
    <div class="card bad"><div class="label">Atrasado</div><div class="value">{pedidos['atrasado'][0]}</div><div class="sub">{pct(pedidos['atrasado'][1])}</div></div>
    <div class="card ok"><div class="label">Entrada no sistema realizada</div><div class="value">{pedidos['entrada_ok'][0]}</div><div class="sub">{pct(pedidos['entrada_ok'][1])}</div></div>
    <div class="card warn"><div class="label">Entrada pendente no sistema</div><div class="value">{pedidos['entrada_pendente'][0]}</div><div class="sub">{pct(pedidos['entrada_pendente'][1])}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status geral dos pedidos <button class="chart-export-btn" data-chart-export="chart-pedidos-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-status"></canvas></div></div>
    <div class="chart-box"><h4>Entrada no sistema <button class="chart-export-btn" data-chart-export="chart-pedidos-entrada" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-entrada"></canvas></div></div>
    <div class="chart-box wide"><h4>Pedidos por filial (entregue / pendente / atrasado) <button class="chart-export-btn" data-chart-export="chart-pedidos-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-pedidos-filial"></canvas></div></div>
  </div>
  <h3>📋 Pedidos por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-pedidos-filial" placeholder="Filtrar por filial...">
    <button class="table-export-btn" data-export="tbl-pedidos-filial">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-pedidos-filial" class="sortable">
    <thead><tr><th>Filial</th><th class="num">Total</th><th class="num">Entregue</th><th class="num">Pendente</th><th class="num">Atrasado</th><th class="num">% Entregue</th></tr></thead>
    <tbody>
    {linhas_pedidos_filiais()}
    </tbody>
  </table>
  </div>

  <h3>📋 Tabela de pedidos realizados no GN</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-pedidos-detalhe" placeholder="Buscar por filial, pedido ou produto...">
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
      <th data-col="fil">Filial Destino</th><th data-col="ped" class="num">Nº Pedido GN</th><th data-col="qtd" class="num">Qtde</th><th data-col="desc">Descrição do Produto</th>
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
    <div class="card"><div class="label">Total de notas</div><div class="value">{notas_resumo.get('Total de notas (linhas)', 0)}</div></div>
    <div class="card warn"><div class="label">Pendentes</div><div class="value">{notas_resumo.get('Pendentes (NÃO)', 0)}</div></div>
    <div class="card bad"><div class="label">Atrasadas</div><div class="value">{notas_resumo.get('Atrasadas', 0)}</div></div>
    <div class="card ok"><div class="label">Aguardando (no prazo)</div><div class="value">{notas_resumo.get('Aguardando (no prazo)', 0)}</div></div>
    <div class="card"><div class="label">Valor total pendente</div><div class="value">{brl(notas_resumo.get('Valor total pendente', 0))}</div></div>
    <div class="card bad"><div class="label">Maior atraso</div><div class="value">{int(notas_resumo.get('Maior atraso (dias)', 0))} dias</div></div>
    <div class="card"><div class="label">Média de atraso</div><div class="value">{notas_resumo.get('Média de atraso (dias)', 0):.1f} dias</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Notas pendentes: atrasadas x no prazo <button class="chart-export-btn" data-chart-export="chart-notas-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-notas-status"></canvas></div></div>
  </div>
  <h3>⚠️ Notas com criticidade elevada (atrasadas)</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-notas" placeholder="Filtrar por filial, NF ou produto...">
    <button class="table-export-btn" data-export="tbl-notas">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-notas" class="sortable">
    <thead><tr><th>Filial</th><th>Nota Fiscal</th><th>Descrição</th><th class="num">Qtde</th><th>Dia da Entrega</th><th class="num">Dias em Atraso</th></tr></thead>
    <tbody>
    {linhas_notas()}
    </tbody>
  </table>
  </div>
</section>

<section id="transferencias">
  <h2>🔄 Transferências Pendentes</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {transf_mtime}</p>
  <div class="cards">
    <div class="card"><div class="label">Qtde total (unidades)</div><div class="value">{transf_resumo.get('Qtde total (unidades)', 0)}</div></div>
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
    </tr></thead>
    <tbody id="tbody-transf"></tbody>
  </table>
  </div>
  <div class="pager" id="contador-transf"></div>
</section>

{secao_pelicula("amet", "🛡️ Películas AMET por Filial", amet, "amet")}

{secao_pelicula("devia", "🛡️ Películas DEVIA por Filial", devia, "devia")}

{secao_pelicula("upmaster", "🛡️ Películas UPMASTER por Filial", upmaster, "upmaster")}

{secao_manutencoes(manutencoes_mtime, manutencoes_resumo)}

{secao_acessorios("acessorios", "🎧 Acessórios nas Filiais", acessorios_mtime, acessorios_diversos_resumo, "acessorios")}
{secao_seriais_tim(acessorios_mtime, seriais_resumo)}
{secao_devolvidos(devolvidos_mtime, devolvidos_resumo)}
{secao_malotes(malotes_mtime, malotes_resumo, malotes_log_resumo)}

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
  ['np', 'Ainda não processado'], ['0-3', '0-3 dias'], ['4-7', '4-7 dias'],
  ['8-15', '8-15 dias'], ['16-30', '16-30 dias'], ['30+', '30+ dias']
];

function diasNaFaixa(dias, faixa) {{
  if (dias < 0) return faixa === 'np';
  if (faixa === 'np') return false;
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

['tbl-pedidos-filial', 'tbl-notas', 'tbl-amet', 'tbl-devia', 'tbl-upmaster', 'tbl-transf', 'tbl-malotes-filial', 'tbl-manutencoes'].forEach(registrarExportDOM);

// ---- fabrica generica de tabela paginada com busca + multi-selecao, usada pelas tabelas de acessorios ----
function criarTabelaPaginada(opts) {{
  var estado = {{ filtro: '', sortCol: opts.sortInicial || null, sortDir: 'desc', pagina: 1, sets: {{}} }};
  opts.filtros.forEach(function(f) {{ estado.sets[f.campo] = new Set(); }});

  function dadosFiltrados() {{
    var q = estado.filtro.toLowerCase();
    var out = opts.dados.filter(function(r) {{
      if (q && opts.busca(r).toLowerCase().indexOf(q) === -1) return false;
      for (var campo in estado.sets) {{
        if (estado.sets[campo].size && !estado.sets[campo].has(r[campo])) return false;
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
    desc: {{ sort: function(r) {{ return r.desc.toLowerCase(); }} }},
    stprod: {{ sort: function(r) {{ return r.stprod.toLowerCase(); }} }},
    dias: {{ sort: function(r) {{ return r.dias_num; }} }},
    entrada: {{ sort: function(r) {{ return r.entrada.toLowerCase(); }} }},
    status: {{ sort: function(r) {{ return r.status.toLowerCase(); }} }}
  }};

  function linhaHtml(r) {{
    return '<tr><td>' + r.dreal + '</td><td>' + r.dproc + '</td><td>' + r.dprev + '</td>' +
      '<td>' + r.fil + '</td><td class="num">' + r.ped + '</td><td class="num">' + r.qtd + '</td>' +
      '<td>' + r.desc + '</td><td>' + r.stprod + '</td><td class="num">' + r.dias + '</td>' +
      '<td>' + r.entrada + '</td><td><span class="pill ' + r.cls + '">' + r.status + '</span></td></tr>';
  }}

  function dadosFiltrados() {{
    var q = estado.filtro.toLowerCase();
    var out = DADOS.filter(function(r) {{
      if (q && (r.fil + ' ' + r.ped + ' ' + r.desc + ' ' + r.status).toLowerCase().indexOf(q) === -1) return false;
      if (estado.filial.size && !estado.filial.has(r.fil)) return false;
      if (estado.status.size && !estado.status.has(r.status)) return false;
      if (estado.stprod.size && !estado.stprod.has(r.stprod)) return false;
      if (estado.entrada.size && !estado.entrada.has(r.entrada)) return false;
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
    var headers = ['Data Realização', 'Data Processamento', 'Data Prevista Entrega', 'Filial Destino', 'Nº Pedido GN', 'Qtde', 'Descrição do Produto', 'Status do Produto', 'Dias em Aberto', 'Entrada no Sistema', 'Status'];
    var linhas = dadosFiltrados().map(function(r) {{
      return [r.dreal, r.dproc, r.dprev, r.fil, r.ped, r.qtd, r.desc, r.stprod, r.dias, r.entrada, r.status];
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
      '<td class="num">' + r.qtd + '</td><td><span class="pill ' + r.cls + '">' + r.crit + '</span></td></tr>';
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
    plugins: {{ legend: {{ position: 'bottom' }} }}
  }}
}});

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
    plugins: {{ legend: {{ position: 'bottom' }} }}
  }}
}});

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
    plugins: {{ legend: {{ position: 'bottom' }} }}
  }}
}});

new Chart(document.getElementById('chart-acessorios-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.acessoriosDiversos.labels,
    datasets: [{{ label: 'Saldo', data: CHART_DATA.acessoriosDiversos.saldo, backgroundColor: COR_ACCENT }}]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
}});

new Chart(document.getElementById('chart-acessorios-tim-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.acessoriosTim.labels,
    datasets: [{{ label: 'Saldo', data: CHART_DATA.acessoriosTim.saldo, backgroundColor: COR_OK }}]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
}});

new Chart(document.getElementById('chart-devolvidos-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.devolvidos.labels,
    datasets: [{{ label: 'Saldo (unidades)', data: CHART_DATA.devolvidos.saldo, backgroundColor: COR_BAD }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }},
    scales: {{ x: {{ type: 'logarithmic', title: {{ display: true, text: 'Saldo (escala logarítmica — um chip com milhares de unidades em Estoque Matriz dominaria a escala linear)', color: '#94a3b8', font: {{ size: 10 }} }} }} }}
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

new Chart(document.getElementById('chart-malotes-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.malotesFilial.labels,
    datasets: [
      {{ label: 'Na Filial', data: CHART_DATA.malotesFilial.na_filial, backgroundColor: COR_ACCENT }},
      {{ label: 'No ADM', data: CHART_DATA.malotesFilial.no_adm, backgroundColor: COR_OK }}
    ]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: {{ x: {{ stacked: true }}, y: {{ stacked: true }} }}, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

var CORES_STATUS_MANUTENCOES = {{ 'PENDENTE': COR_WARN, 'EM ANDAMENTO': COR_WARN, 'CONCLUÍDO': COR_OK, 'CONCLUIDO': COR_OK, 'CONCLUÍDA': COR_OK, 'CONCLUIDA': COR_OK, 'CANCELADO': COR_BAD, 'CANCELADA': COR_BAD }};
new Chart(document.getElementById('chart-manutencoes-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.manutencoesStatus.labels,
    datasets: [{{
      data: CHART_DATA.manutencoesStatus.values,
      backgroundColor: CHART_DATA.manutencoesStatus.labels.map(function(l) {{ return CORES_STATUS_MANUTENCOES[l] || COR_ACCENT; }})
    }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-manutencoes-filial'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.manutencoesFilial.labels,
    datasets: [{{ label: 'Chamados', data: CHART_DATA.manutencoesFilial.values, backgroundColor: COR_ACCENT }}]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
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
