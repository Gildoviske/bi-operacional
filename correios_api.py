# -*- coding: utf-8 -*-
"""Consulta o rastreamento oficial dos Correios.

As credenciais ficam em %LOCALAPPDATA%\\rocha-bi\\correios.json — fora do Google Drive
e fora do repositório, para não sincronizarem nem irem parar no GitHub.

Sem credencial preenchida, ou se a API falhar, as funções devolvem o que já está em
cache e o restante da geração da página segue normalmente.
"""
import base64
import json
import os
import ssl
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

BASE_API = "https://api.correios.com.br"
# pasta visivel do perfil do usuario (AppData e oculta e dificultava editar o arquivo)
CREDENCIAIS = Path(os.environ.get("USERPROFILE", ".")) / "rocha-bi" / "correios.json"
# o token e um segredo de curta duracao: fica junto da credencial, fora do Drive
TOKEN_CACHE = CREDENCIAIS.with_name("token.json")
CACHE = Path(r"I:\Meu Drive\CONTROLE COMPRAS\rastreios") / "cache_rastreamento.json"

# um objeto entregue não muda mais; os demais são reconsultados depois desse intervalo
VALIDADE_EM_TRANSITO = timedelta(hours=6)
# a API de token aceita 3 req/s e devolve 429 se estourar, entao o token e reaproveitado
# ate perto de expirar, como a propria documentacao recomenda
MARGEM_TOKEN = timedelta(minutes=30)
LOTE = 50               # a API aceita vários códigos por chamada
TIMEOUT = 20


def _agora():
    return datetime.now()


def carregar_credenciais():
    """Aceita as duas formas de acesso dos Correios.

    - chave_acesso: gerada em CWS > Chaves de Acesso (subdelegacao). Vai direto no
      cabecalho como Bearer, sem passar pelo endpoint de token.
    - usuario + codigo_acesso: Basic no /token/v1/autentica, que devolve um token.
    """
    if not CREDENCIAIS.exists():
        return None
    try:
        dados = json.loads(CREDENCIAIS.read_text(encoding="utf-8"))
    except Exception:
        return None
    if dados.get("chaves") or (dados.get("chave_acesso") or "").strip():
        return dados
    if not dados.get("usuario") or not dados.get("codigo_acesso"):
        return None
    return dados


def carregar_cache():
    if CACHE.exists():
        try:
            return json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def gravar_json_seguro(caminho, dados):
    """Grava num temporario e troca no fim: se o disco encher no meio, o arquivo antigo fica intacto.
    Gravando direto, a falha truncava o cache e a leitura seguinte quebrava ou perdia tudo."""
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    tmp = caminho.with_name(caminho.name + ".tmp")
    try:
        tmp.write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(tmp, caminho)
    finally:
        if tmp.exists():
            tmp.unlink()


def ler_json_seguro(caminho, padrao):
    try:
        return json.loads(Path(caminho).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return padrao


def salvar_cache(cache):
    gravar_json_seguro(CACHE, cache)


def _pedir(url, dados=None, cabecalhos=None, metodo=None):
    req = urllib.request.Request(
        url,
        data=json.dumps(dados).encode("utf-8") if dados is not None else None,
        headers={"Accept": "application/json", "Content-Type": "application/json",
                 **(cabecalhos or {})},
        method=metodo or ("POST" if dados is not None else "GET"),
    )
    contexto = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=TIMEOUT, context=contexto) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _token_guardado():
    """Token ainda valido do cache, respeitando a margem antes de expirar."""
    if not TOKEN_CACHE.exists():
        return None
    try:
        dados = json.loads(TOKEN_CACHE.read_text(encoding="utf-8"))
        expira = datetime.fromisoformat(dados["expira_em"])
    except Exception:
        return None
    if _agora() + MARGEM_TOKEN < expira:
        return dados.get("token")
    return None


def _guardar_token(token, expira_em):
    TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
    gravar_json_seguro(TOKEN_CACHE, {"token": token, "expira_em": expira_em})
    try:
        os.chmod(TOKEN_CACHE, 0o600)
    except Exception:
        pass


def listar_chaves(cred):
    """Chaves de acesso disponiveis, uma por empresa/contrato.

    Cada chave so enxerga os objetos do contrato ou cartao a que foi vinculada, entao
    com duas empresas postando e preciso tentar as duas.
    """
    chaves = []
    for item in (cred.get("chaves") or []):
        valor = (item.get("chave") or "").strip()
        if valor:
            chaves.append((item.get("nome") or "?", valor))
    solta = (cred.get("chave_acesso") or "").strip()
    if solta and not chaves:
        chaves.append((cred.get("empresa") or "principal", solta))
    return chaves


def autenticar(cred):
    """Devolve o bearer token, reaproveitando o do cache enquanto valer."""
    # a chave de subdelegacao ja e o proprio Bearer: nao ha o que autenticar
    # vale tambem para a lista "chaves": sem isso a consulta de CEP caia no Basic sem codigo e dava 401
    chaves = listar_chaves(cred)
    if chaves:
        return chaves[0][1]

    guardado = _token_guardado()
    if guardado:
        return guardado

    basico = base64.b64encode(
        f"{cred['usuario']}:{cred['codigo_acesso']}".encode("utf-8")).decode("ascii")
    cabecalhos = {"Authorization": f"Basic {basico}"}
    cartao = (cred.get("cartao_postagem") or "").strip()
    contrato = (cred.get("contrato") or "").strip()
    dr = cred.get("dr")
    try:
        # o token com escopo de contrato carrega as APIs liberadas para ele; so funciona
        # quando o idCorreios e o titular do CNPJ do contrato
        if contrato and dr:
            resposta = _pedir(f"{BASE_API}/token/v1/autentica/contrato",
                              {"numero": contrato, "dr": int(dr)}, cabecalhos)
        elif cartao:
            resposta = _pedir(f"{BASE_API}/token/v1/autentica/cartaopostagem",
                              {"numero": cartao}, cabecalhos)
        else:
            resposta = _pedir(f"{BASE_API}/token/v1/autentica", {}, cabecalhos)
    except urllib.error.HTTPError as e:
        corpo = e.read().decode("utf-8", "replace")[:200]
        if e.code == 429:
            print("   Correios: limite de 3 requisicoes/segundo atingido na geracao do token")
            return None
        # TOK-015/016: o login nao e o titular do CNPJ do contrato. Ainda assim vale
        # autenticar so pelo usuario, que e o que funciona hoje.
        if "TOK-01" in corpo and (contrato or cartao):
            try:
                resposta = _pedir(f"{BASE_API}/token/v1/autentica", {}, cabecalhos)
            except Exception:
                print(f"   Correios: autenticacao falhou (HTTP {e.code}) {corpo}")
                return None
        else:
            print(f"   Correios: autenticacao falhou (HTTP {e.code}) {corpo}")
            return None
    except Exception as e:
        print(f"   Correios: autenticacao falhou ({type(e).__name__}: {e})")
        return None

    token = resposta.get("token")
    expira = resposta.get("expiraEm")
    if token and expira:
        try:
            # normaliza para ISO sem fuso, que e o que o cache compara
            _guardar_token(token, datetime.fromisoformat(
                str(expira).replace("Z", "")).isoformat(timespec="seconds"))
        except Exception:
            pass
    return token


def _previsao(objeto):
    """dtPrevista ja vem na resposta do rastreamento: nao precisa da API de Prazo."""
    bruto = objeto.get("dtPrevista") or ""
    if not bruto:
        return ""
    try:
        return datetime.fromisoformat(str(bruto).replace("Z", "")).date().isoformat()
    except Exception:
        return ""


_TOKEN_CEP = None


def buscar_cep(cep, token=None):
    """Endereço pelo CEP na API oficial dos Correios. None se não achar."""
    cep = "".join(c for c in str(cep) if c.isdigit())
    if len(cep) != 8:
        return None
    global _TOKEN_CEP
    if token is None:
        # autentica uma vez por execucao: com o cache de CEP vazio sao dezenas de consultas
        if _TOKEN_CEP is None:
            cred = carregar_credenciais()
            _TOKEN_CEP = (autenticar(cred) if cred else None) or ""
        token = _TOKEN_CEP
    if not token:
        return None
    try:
        dados = _pedir(f"{BASE_API}/cep/v2/enderecos/{cep}",
                       cabecalhos={"Authorization": f"Bearer {token}"})
    except Exception:
        return None
    return {
        "logradouro": dados.get("logradouro") or "",
        "bairro": dados.get("bairro") or "",
        "cidade": dados.get("localidade") or "",
        "uf": dados.get("uf") or "",
    }


def _resumir(objeto):
    """Extrai status legível do objeto devolvido pela API."""
    eventos = objeto.get("eventos") or []
    if not eventos:
        return {"status": objeto.get("mensagem") or "Sem movimentação",
                "local": "", "data": "", "entregue": False,
                "previsao": _previsao(objeto)}
    ev = eventos[0]                      # a API devolve do mais recente para o mais antigo
    unidade = ev.get("unidade") or {}
    endereco = unidade.get("endereco") or {}
    local = ", ".join(p for p in (endereco.get("cidade"), endereco.get("uf")) if p)
    data = ev.get("dtHrCriado") or ""
    if data:
        try:
            data = datetime.fromisoformat(data.replace("Z", "")).strftime("%d/%m/%Y %H:%M")
        except Exception:
            pass
    descricao = ev.get("descricao") or ""
    return {
        "status": descricao,
        "local": local or (unidade.get("nome") or ""),
        "data": data,
        # o código BDE/01 e a descrição "entregue" marcam a finalização
        "entregue": "entregue" in descricao.lower() and "destinat" in descricao.lower()
                    or ev.get("codigo") == "BDE",
        "previsao": _previsao(objeto),
    }


def _precisa_consultar(registro):
    if not registro:
        return True
    if registro.get("entregue"):
        return False
    consultado = registro.get("consultado_em")
    if not consultado:
        return True
    try:
        return _agora() - datetime.fromisoformat(consultado) > VALIDADE_EM_TRANSITO
    except Exception:
        return True


def atualizar(codigos, verboso=True):
    """Devolve {codigo: {status, local, data, entregue}} para os códigos pedidos.

    Consulta só o que está sem cache, vencido ou ainda não entregue.
    """
    cache = carregar_cache()
    pendentes = [c for c in dict.fromkeys(codigos) if _precisa_consultar(cache.get(c))]

    if not pendentes:
        if verboso:
            print(f"   Correios: {len(codigos)} objetos, todos em cache")
        return cache

    cred = carregar_credenciais()
    if not cred:
        if verboso:
            print(f"   Correios: {len(pendentes)} objetos sem rastreamento "
                  f"(credenciais nao preenchidas em {CREDENCIAIS})")
        return cache

    chaves = listar_chaves(cred)
    if not chaves:
        token = autenticar(cred)
        if not token:
            if verboso:
                print("   Correios: seguindo sem rastreamento nesta rodada")
            return cache
        chaves = [(cred.get("empresa") or "principal", token)]

    consultados = 0
    restantes = list(pendentes)
    for nome, token in chaves:
        if not restantes:
            break
        cabecalhos = {"Authorization": f"Bearer {token}"}
        nao_resolvidos = []
        for i in range(0, len(restantes), LOTE):
            lote = restantes[i:i + LOTE]
            url = (f"{BASE_API}/srorastro/v1/objetos?codigosObjetos={','.join(lote)}"
                   f"&resultado=U")
            try:
                resposta = _pedir(url, cabecalhos=cabecalhos)
            except urllib.error.HTTPError as e:
                corpo = e.read().decode("utf-8", "replace")[:300]
                if "GTW-012" in corpo:
                    print(f"   Correios [{nome}]: SRO Rastro nao liberado para esta chave")
                else:
                    print(f"   Correios [{nome}]: consulta falhou (HTTP {e.code}) {corpo}")
                nao_resolvidos.extend(lote)
                continue
            except Exception as e:
                print(f"   Correios [{nome}]: consulta falhou ({type(e).__name__}: {e})")
                nao_resolvidos.extend(lote)
                continue
            for objeto in resposta.get("objetos") or []:
                codigo = objeto.get("codObjeto")
                if not codigo:
                    continue
                registro = _resumir(objeto)
                # objeto de outra empresa: guarda para tentar na proxima chave
                if not (objeto.get("eventos") or []):
                    mensagem = str(objeto.get("mensagem") or "")
                    if "SRO-0" in mensagem or "nao pertence" in mensagem.lower() \
                            or "não pertence" in mensagem.lower():
                        nao_resolvidos.append(codigo)
                        cache.setdefault(codigo, dict(registro, empresa=""))
                        continue
                registro["empresa"] = nome
                registro["consultado_em"] = _agora().isoformat(timespec="seconds")
                cache[codigo] = registro
                consultados += 1
        restantes = nao_resolvidos

    salvar_cache(cache)
    if verboso:
        entregues = sum(1 for c in codigos if (cache.get(c) or {}).get("entregue"))
        print(f"   Correios: {consultados} objetos consultados, "
              f"{entregues} de {len(codigos)} entregues")
    return cache
