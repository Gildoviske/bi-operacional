# -*- coding: utf-8 -*-
"""Gera a página HTML do BI Operacional a partir das planilhas de controle.

Uso: python gerar_pagina.py
Lê os 3 arquivos em "I:\\Meu Drive\\CONTROLE COMPRAS" e grava index.html
nesta mesma pasta (bi-web), pronta para publicar no GitHub Pages.
"""
import json
import re
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


def mtime_str(arquivo):
    ts = (BASE / arquivo).stat().st_mtime
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y às %H:%M")


def load_abs(caminho, aba):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True, keep_vba=caminho.suffix == ".xlsm")
    return list(wb[aba].iter_rows(values_only=True))


def mtime_str_abs(caminho):
    ts = caminho.stat().st_mtime
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y às %H:%M")


def pct(v):
    return f"{v * 100:.1f}%".replace(".", ",")


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_str(v):
    if v is None:
        return ""
    return v.strftime("%d/%m/%Y")


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

# ------------------------------------------------------------- AMET
amet_mtime = mtime_str("CONTROLE QUANTIDADE DE AMET NAS FILIAIS.xlsx")
rows = load("CONTROLE QUANTIDADE DE AMET NAS FILIAIS.xlsx", "AMET NAS FILIAIS")

amet_data_estoque = (rows[0][0] or "").split("ATUALIZADO DIA ")[-1].rstrip(")")

# lê pelas posições do cabeçalho (linha com "Filial"/"Total Geral") em vez de índice fixo de
# coluna, pois a planilha já mudou de estrutura (colunas de produto adicionadas/removidas).
header = next(r for r in rows if r[0] == "Filial")
col_filial_1, col_total_1 = 0, header.index("Total Geral")
col_filial_2 = header.index("Filial", col_total_1 + 1)
col_total_2 = header.index("Total Geral", col_filial_2 + 1)

amet_periodo_vendas = (rows[0][col_filial_2] or "").split("(DO DIA ")[-1].rstrip(")")


def capturar_totais(col_filial, col_total):
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


amet_estoque_por_filial = capturar_totais(col_filial_1, col_total_1)
amet_vendido_por_filial = capturar_totais(col_filial_2, col_total_2)

amet_filiais = []
for nome in sorted(set(amet_estoque_por_filial) | set(amet_vendido_por_filial)):
    estoque = amet_estoque_por_filial.get(nome, 0)
    vendido = amet_vendido_por_filial.get(nome, 0)
    amet_filiais.append({"filial": nome, "estoque": estoque, "vendido": vendido})
amet_filiais.sort(key=lambda f: f["vendido"], reverse=True)

amet_estoque_total = sum(amet_estoque_por_filial.values())
amet_vendido_total = sum(amet_vendido_por_filial.values())

# ------------------------------------------------------- ACESSÓRIOS
acessorios_mtime = mtime_str("CONTROLE CONFIGURAÇÕES PRODUTOS.xlsx")
acessorios_rows = load_dicts("CONTROLE CONFIGURAÇÕES PRODUTOS.xlsx", "SALDO PRODUTOS NAS FILIAIS")
acessorios_rows = [r for r in acessorios_rows if r.get("Filial") is not None]


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

# ---------------------------------------------------------- INVENTÁRIOS
INVENTARIO_FILE = Path(r"H:\Meu Drive\Controle de inventário\Controle de inventario automatizada.xlsm")
inventario_mtime = mtime_str_abs(INVENTARIO_FILE)

ap_chip_rows = load_abs(INVENTARIO_FILE, "Resumo (AP. e CHIP)")
data_contagem_raw = ap_chip_rows[0][1]
data_contagem_str = data_str(data_contagem_raw) if data_contagem_raw else "-"

STATUS_APARELHOS_CLASSE = {
    "OK": "ok", "OK - DIV": "warn", "OK - EX": "warn", "CONFIRMAR DIV": "warn",
    "NOK": "bad", "NOK - DIV": "bad",
}
# chips com excesso de divergência (OK - EX) é situação esperada/aceita, não conta como divergência
STATUS_CHIPS_CLASSE = {
    "OK": "ok", "OK - DIV": "warn", "OK - EX": "ok",
    "NOK": "bad", "NOK - DIV": "bad",
}
BUCKET_PRIORIDADE = {"Com divergência": 0, "Não iniciado": 1, "Parcial": 2, "OK com ajuste": 3, "Completo": 4}
BUCKET_CLASSE = {"Completo": "ok", "OK com ajuste": "ok", "Parcial": "warn", "Com divergência": "bad", "Não iniciado": ""}


def extrair_area(responsavel):
    if not responsavel:
        return "-"
    m = re.search(r"\(([^)]+)\)\s*$", responsavel.strip())
    return m.group(1) if m else "-"


def bucket_filial(aparelhos, chips):
    problematicos_aparelhos = {"NOK", "NOK - DIV", "OK - EX", "CONFIRMAR DIV"}
    problematicos_chips = {"NOK", "NOK - DIV"}
    if aparelhos is None and chips is None:
        return "Não iniciado"
    if aparelhos is None or chips is None:
        return "Parcial"
    if aparelhos in problematicos_aparelhos or chips in problematicos_chips:
        return "Com divergência"
    if aparelhos == "OK - DIV" or chips == "OK - DIV":
        return "OK com ajuste"
    return "Completo"


inventario_filiais = []
for r in ap_chip_rows[2:]:
    filial = r[0]
    if not filial:
        continue
    aparelhos, chips, qtd, responsavel = r[1], r[2], r[3], r[4]
    bucket = bucket_filial(aparelhos, chips)
    inventario_filiais.append({
        "filial": filial, "area": extrair_area(responsavel),
        "aparelhos": aparelhos or "PENDENTE",
        "aparelhos_cls": STATUS_APARELHOS_CLASSE.get(aparelhos, "warn" if aparelhos is None else ""),
        "chips": chips or "PENDENTE",
        "chips_cls": STATUS_CHIPS_CLASSE.get(chips, "warn" if chips is None else ""),
        "qtd": qtd if qtd is not None else "-",
        "responsavel": responsavel or "-",
        "bucket": bucket, "bucket_cls": BUCKET_CLASSE.get(bucket, ""),
    })
inventario_filiais.sort(key=lambda f: (BUCKET_PRIORIDADE.get(f["bucket"], 5), f["filial"]))

inventario_resumo = {
    "total": len(inventario_filiais),
    "completo": sum(1 for f in inventario_filiais if f["bucket"] in ("Completo", "OK com ajuste")),
    "pendente": sum(1 for f in inventario_filiais if f["bucket"] in ("Não iniciado", "Parcial")),
    "divergencia": sum(1 for f in inventario_filiais if f["bucket"] == "Com divergência"),
}

inventario_buckets = {}
for f in inventario_filiais:
    inventario_buckets[f["bucket"]] = inventario_buckets.get(f["bucket"], 0) + 1
INVENTARIO_ORDEM_BUCKET = ["Completo", "OK com ajuste", "Parcial", "Com divergência", "Não iniciado"]

# ---- histórico: a planilha só guarda o dia atual, então construímos o histórico
# nós mesmos, gravando o snapshot de cada dia num arquivo à parte a cada publicação.
HISTORICO_PATH = Path(__file__).resolve().parent / "historico_inventario.json"


def atualizar_historico_inventario(data_hoje, filiais):
    if not data_hoje or data_hoje == "-":
        historico_existente = json.loads(HISTORICO_PATH.read_text(encoding="utf-8")) if HISTORICO_PATH.exists() else []
        return historico_existente
    historico = json.loads(HISTORICO_PATH.read_text(encoding="utf-8")) if HISTORICO_PATH.exists() else []
    historico = [h for h in historico if h["data"] != data_hoje]
    data_ts = int(datetime.strptime(data_hoje, "%d/%m/%Y").timestamp() * 1000)
    for f in filiais:
        historico.append({
            "data": data_hoje, "data_ts": data_ts, "filial": f["filial"], "area": f["area"],
            "aparelhos": f["aparelhos"], "chips": f["chips"],
            "bucket": f["bucket"], "bucket_cls": f["bucket_cls"],
        })
    datas_unicas = sorted(set(h["data"] for h in historico), key=lambda d: datetime.strptime(d, "%d/%m/%Y"))
    datas_manter = set(datas_unicas[-60:])
    historico = [h for h in historico if h["data"] in datas_manter]
    HISTORICO_PATH.write_text(json.dumps(historico, ensure_ascii=False), encoding="utf-8")
    return historico


historico_inventario = atualizar_historico_inventario(data_contagem_str, inventario_filiais)
historico_dias = sorted(set(h["data"] for h in historico_inventario), key=lambda d: datetime.strptime(d, "%d/%m/%Y"))
historico_primeiro_dia = historico_dias[0] if historico_dias else "-"

historico_por_filial = {}
for h in historico_inventario:
    stats = historico_por_filial.setdefault(h["filial"], {"dias": 0, "completos": 0})
    stats["dias"] += 1
    if h["bucket"] in ("Completo", "OK com ajuste"):
        stats["completos"] += 1

historico_taxas = []
for filial, stats in historico_por_filial.items():
    taxa = round(100 * stats["completos"] / stats["dias"], 1) if stats["dias"] else 0
    historico_taxas.append({"filial": filial, "dias": stats["dias"], "completos": stats["completos"], "taxa": taxa})
historico_taxas.sort(key=lambda x: x["taxa"])

historico_resumo = {
    "dias_registrados": len(historico_dias),
    "registros": len(historico_inventario),
    "taxa_media": round(sum(t["taxa"] for t in historico_taxas) / len(historico_taxas), 1) if historico_taxas else 0,
}
historico_json = json.dumps(historico_inventario, ensure_ascii=False)

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
        "labels": [f["filial"] for f in amet_filiais],
        "estoque": [f["estoque"] for f in amet_filiais],
        "vendido": [f["vendido"] for f in amet_filiais],
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
    "inventarioBuckets": {
        "labels": [b for b in INVENTARIO_ORDEM_BUCKET if inventario_buckets.get(b)],
        "values": [inventario_buckets.get(b, 0) for b in INVENTARIO_ORDEM_BUCKET if inventario_buckets.get(b)],
    },
    "historicoTaxas": {
        "labels": [t["filial"] for t in historico_taxas],
        "taxa": [t["taxa"] for t in historico_taxas],
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


def linhas_amet():
    out = []
    for f in amet_filiais:
        giro = pct(f["vendido"] / f["estoque"]) if f["estoque"] else "-"
        out.append(
            "<tr><td>{fil}</td><td class='num'>{est}</td><td class='num'>{ven}</td>"
            "<td class='num'>{giro}</td></tr>".format(
                fil=f["filial"], est=f["estoque"], ven=f["vendido"], giro=giro
            )
        )
    return "\n".join(out)


def linhas_inventario():
    out = []
    for f in inventario_filiais:
        out.append(
            "<tr><td>{fil}</td><td>{area}</td>"
            "<td><span class='pill {ap_cls}'>{ap}</span></td>"
            "<td><span class='pill {ch_cls}'>{ch}</span></td>"
            "<td>{resp}</td>"
            "<td><span class='pill {b_cls}'>{b}</span></td></tr>".format(
                fil=f["filial"], area=f["area"], ap=f["aparelhos"], ap_cls=f["aparelhos_cls"],
                ch=f["chips"], ch_cls=f["chips_cls"], resp=f["responsavel"],
                b=f["bucket"], b_cls=f["bucket_cls"]
            )
        )
    return "\n".join(out)


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
      <th data-col="fabricante">Fabricante</th><th data-col="saldo">Saldo</th><th data-col="disponivel">Disponível</th>
      <th data-col="valor">Valor Unit.</th><th data-col="valor_total">Valor Total</th>
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
      <th data-col="dias">Dias em Estoque</th><th data-col="valor">Valor</th>
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
      <th data-col="fabricante">Fabricante</th><th data-col="saldo">Saldo</th><th data-col="custo">Custo Unit.</th>
      <th data-col="custo_total">Custo Total</th><th data-col="data_mov">Última Movimentação</th>
    </tr></thead>
    <tbody id="tbody-devolvidos"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-devolvidos"></div>
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
  .sidebar a {{ display: flex; align-items: center; gap: 8px; padding: 12px 20px; color: var(--muted); text-decoration: none; font-size: 14px; border-left: 3px solid transparent; }}
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
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
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
  <a href="#pedidos" class="nav-link">📦 Pedidos</a>
  <a href="#notas" class="nav-link">🧾 Notas Fiscais</a>
  <a href="#transferencias" class="nav-link">🔄 Transferências</a>
  <a href="#amet" class="nav-link">🛡️ Películas AMET</a>
  <a href="#acessorios" class="nav-link">🎧 Acessórios</a>
  <a href="#acessorios-tim" class="nav-link">📶 Fidelizados TIM</a>
  <a href="#devolvidos" class="nav-link">♻️ Devolvidos e Defeitos</a>
  <a href="#inventarios" class="nav-link">📋 Inventários</a>
  <a href="#inventarios-historico" class="nav-link">📆 Histórico de Contagem</a>
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
    <thead><tr><th>Filial</th><th>Total</th><th>Entregue</th><th>Pendente</th><th>Atrasado</th><th>% Entregue</th></tr></thead>
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
      <th data-col="fil">Filial Destino</th><th data-col="ped">Nº Pedido GN</th><th data-col="qtd">Qtde</th><th data-col="desc">Descrição do Produto</th>
      <th data-col="stprod">Status do Produto</th><th data-col="dias">Dias em Aberto</th><th data-col="entrada">Entrada no Sistema</th><th data-col="status">Status</th>
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
    <thead><tr><th>Filial</th><th>Nota Fiscal</th><th>Descrição</th><th>Qtde</th><th>Dia da Entrega</th><th>Dias em Atraso</th></tr></thead>
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
      <th data-col="orig">Filial Origem</th><th data-col="dias">Dias fora do estoque</th><th data-col="dest">Filial Destino</th>
      <th data-col="user">Usuário Solicitante</th><th data-col="nf">NF</th><th data-col="prod">Produto</th>
      <th data-col="desc">Descrição</th><th data-col="qtd">Qtde</th><th data-col="crit">Criticidade</th>
    </tr></thead>
    <tbody id="tbody-transf"></tbody>
  </table>
  </div>
  <div class="pager" id="contador-transf"></div>
</section>

<section id="amet">
  <h2>🛡️ Películas AMET por Filial</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {amet_mtime} · Estoque referente a {amet_data_estoque} · Vendas de {amet_periodo_vendas}</p>
  <div class="cards">
    <div class="card"><div class="label">Estoque total (peças)</div><div class="value">{amet_estoque_total}</div></div>
    <div class="card ok"><div class="label">Vendido no período (peças)</div><div class="value">{amet_vendido_total}</div></div>
    <div class="card"><div class="label">Filiais monitoradas</div><div class="value">{len(amet_filiais)}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Estoque x Vendido por filial <button class="chart-export-btn" data-chart-export="chart-amet-filial" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-amet-filial"></canvas></div></div>
  </div>
  <h3>📋 Estoque e vendas por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-amet" placeholder="Filtrar por filial...">
    <button class="table-export-btn" data-export="tbl-amet">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-amet" class="sortable">
    <thead><tr><th>Filial</th><th>Estoque</th><th>Vendido (período)</th><th>Giro (vendido/estoque)</th></tr></thead>
    <tbody>
    {linhas_amet()}
    </tbody>
  </table>
  </div>
</section>

{secao_acessorios("acessorios", "🎧 Acessórios nas Filiais", acessorios_mtime, acessorios_diversos_resumo, "acessorios")}
{secao_seriais_tim(acessorios_mtime, seriais_resumo)}
{secao_devolvidos(devolvidos_mtime, devolvidos_resumo)}

<section id="inventarios">
  <h2>📋 Inventários — Status Diário</h2>
  <p class="secao-mtime">🕒 Planilha atualizada em {inventario_mtime} · Contagem referente a {data_contagem_str}</p>
  <div class="cards">
    <div class="card"><div class="label">Filiais</div><div class="value">{inventario_resumo['total']}</div></div>
    <div class="card ok"><div class="label">Completo</div><div class="value">{inventario_resumo['completo']}</div></div>
    <div class="card warn"><div class="label">Pendente / Não iniciado</div><div class="value">{inventario_resumo['pendente']}</div></div>
    <div class="card bad"><div class="label">Com divergência</div><div class="value">{inventario_resumo['divergencia']}</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h4>Status geral das filiais <button class="chart-export-btn" data-chart-export="chart-inventario-status" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-inventario-status"></canvas></div></div>
  </div>
  <h3>📋 Status por filial</h3>
  <div class="table-toolbar">
    <input class="filtro" data-target="tbl-inventarios" placeholder="Filtrar por filial, área ou responsável...">
    <button class="table-export-btn" data-export="tbl-inventarios">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-inventarios" class="sortable">
    <thead><tr><th>Filial</th><th>Área</th><th>Aparelhos</th><th>Chips</th><th>Responsável</th><th>Status Geral</th></tr></thead>
    <tbody>
    {linhas_inventario()}
    </tbody>
  </table>
  </div>
</section>

<section id="inventarios-historico">
  <h2>📆 Histórico de Contagem por Filial</h2>
  <p class="secao-mtime">🕒 Histórico construído automaticamente a cada publicação, desde {historico_primeiro_dia} · {historico_resumo['dias_registrados']} dia(s) registrado(s) até agora</p>
  <div class="cards">
    <div class="card"><div class="label">Dias registrados</div><div class="value">{historico_resumo['dias_registrados']}</div></div>
    <div class="card"><div class="label">Registros no histórico</div><div class="value">{historico_resumo['registros']}</div></div>
    <div class="card ok"><div class="label">Taxa média de conclusão</div><div class="value">{historico_resumo['taxa_media']}%</div></div>
  </div>
  <div class="charts">
    <div class="chart-box wide"><h4>Taxa de conclusão por filial (histórico) <button class="chart-export-btn" data-chart-export="chart-historico-taxa" title="Baixar gráfico como imagem">📥 PNG</button></h4><div class="canvas-wrap"><canvas id="chart-historico-taxa"></canvas></div></div>
  </div>
  <h3>📋 Linha do tempo por filial</h3>
  <div class="filtros-pedidos">
    <input class="filtro" id="filtro-historico" placeholder="Buscar por filial, data...">
    <div class="msel" id="msel-historico-filial"><button type="button" class="msel-btn" data-default="Todas as filiais">Todas as filiais</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <div class="msel" id="msel-historico-bucket"><button type="button" class="msel-btn" data-default="Status geral (todos)">Status geral (todos)</button><div class="msel-panel"><div class="msel-actions"><button type="button" data-act="all">Marcar todos</button><button type="button" data-act="none">Limpar</button></div><div class="msel-options"></div></div></div>
    <button id="limpar-historico" type="button">Limpar filtros</button>
    <button class="table-export-btn" data-export="tbl-historico">📥 Exportar Excel</button>
  </div>
  <div class="table-wrap">
  <table id="tbl-historico">
    <thead><tr>
      <th data-col="data">Data</th><th data-col="filial">Filial</th><th data-col="area">Área</th>
      <th data-col="aparelhos">Aparelhos</th><th data-col="chips">Chips</th><th data-col="bucket">Status Geral</th>
    </tr></thead>
    <tbody id="tbody-historico"></tbody>
  </table>
  </div>
  <div class="pager" id="pager-historico"></div>
</section>

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

['tbl-pedidos-filial', 'tbl-notas', 'tbl-amet', 'tbl-transf', 'tbl-inventarios'].forEach(registrarExportDOM);

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

// ---- tabela de historico de inventarios ----
function linhaHistoricoHtml(r) {{
  return '<tr><td>' + r.data + '</td><td>' + r.filial + '</td><td>' + r.area + '</td>' +
    '<td>' + r.aparelhos + '</td><td>' + r.chips + '</td>' +
    '<td><span class="pill ' + r.bucket_cls + '">' + r.bucket + '</span></td></tr>';
}}

var HISTORICO_COLS = {{
  data: function(r) {{ return r.data_ts; }},
  filial: function(r) {{ return r.filial.toLowerCase(); }},
  area: function(r) {{ return r.area.toLowerCase(); }},
  aparelhos: function(r) {{ return r.aparelhos.toLowerCase(); }},
  chips: function(r) {{ return r.chips.toLowerCase(); }},
  bucket: function(r) {{ return r.bucket.toLowerCase(); }}
}};

function buscaHistorico(r) {{ return r.filial + ' ' + r.data + ' ' + r.area; }}

criarTabelaPaginada({{
  dados: {historico_json},
  pageSize: 50, sortInicial: 'data',
  tbodyId: 'tbody-historico', pagerId: 'pager-historico', tableId: 'tbl-historico',
  buscaId: 'filtro-historico', limparId: 'limpar-historico',
  colunas: HISTORICO_COLS, linhaHtml: linhaHistoricoHtml, busca: buscaHistorico,
  filtros: [
    {{ id: 'historico-filial', campo: 'filial' }},
    {{ id: 'historico-bucket', campo: 'bucket' }}
  ],
  colunasExport: [
    {{ label: 'Data', get: function(r) {{ return r.data; }} }},
    {{ label: 'Filial', get: function(r) {{ return r.filial; }} }},
    {{ label: 'Área', get: function(r) {{ return r.area; }} }},
    {{ label: 'Aparelhos', get: function(r) {{ return r.aparelhos; }} }},
    {{ label: 'Chips', get: function(r) {{ return r.chips; }} }},
    {{ label: 'Status Geral', get: function(r) {{ return r.bucket; }} }}
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
    datasets: [{{ label: 'Saldo', data: CHART_DATA.devolvidos.saldo, backgroundColor: COR_BAD }}]
  }},
  options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
}});

var CORES_BUCKET_INVENTARIO = {{
  'Completo': COR_OK, 'OK com ajuste': COR_ACCENT, 'Parcial': COR_WARN,
  'Com divergência': COR_BAD, 'Não iniciado': '#64748b'
}};
new Chart(document.getElementById('chart-inventario-status'), {{
  type: 'doughnut',
  data: {{
    labels: CHART_DATA.inventarioBuckets.labels,
    datasets: [{{
      data: CHART_DATA.inventarioBuckets.values,
      backgroundColor: CHART_DATA.inventarioBuckets.labels.map(function(l) {{ return CORES_BUCKET_INVENTARIO[l] || COR_ACCENT; }})
    }}]
  }},
  options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }} }} }}
}});

new Chart(document.getElementById('chart-historico-taxa'), {{
  type: 'bar',
  data: {{
    labels: CHART_DATA.historicoTaxas.labels,
    datasets: [{{ label: 'Taxa de conclusão (%)', data: CHART_DATA.historicoTaxas.taxa, backgroundColor: COR_OK }}]
  }},
  options: {{
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ display: false }} }}, scales: {{ x: {{ max: 100 }} }}
  }}
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
    f"{len(devolvidos_itens)} itens devolvidos/com defeito e inventário de {inventario_resumo['total']} filiais "
    f"({historico_resumo['dias_registrados']} dia(s) no histórico)."
)
