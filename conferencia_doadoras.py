"""
Conferência Doadoras
======================

Gera um Excel separado (conferencia_doadoras_AAAA-MM.xlsx) com as abas:

1. "SALDO pendente" - os MESMOS valores da aba "SALDO pendente" do relatório da
   reconciliação do mês (relatorio_reconciliacao_AAAA-MM.xlsx). A conferência
   NÃO recalcula nada: não lê lista de doadoras nem arquivos de custo, não faz
   matching e não mexe no controle_saldo.csv - só copia o que a reconciliação
   já contabilizou, com o rótulo da categoria por extenso (ex: "AJUDA DE
   CUSTO") e as linhas de resumo "MESES ANTERIORES"/"AJUSTE" (1ª vez só).
   => Por isso a RECONCILIAÇÃO DO MÊS PRECISA TER SIDO RODADA ANTES.

2. "TOTVS" - todos os lançamentos da conta contábil extraídos
   do TOTVS, com o valor líquido (CRÉDITO - DÉBITO) calculado e o "Saldo atual"
   (coluna L do extrato, SALDO ATUAL). Esse extrato
   é ACUMULADO: toda vez que você roda o script com um novo --totvs, os
   lançamentos desse mês são adicionados embaixo dos que já estavam salvos
   (não apaga nada). Na 1ª vez que --totvs é usado, entra também uma linha
   "SALDO INICIAL <CATEGORIA>" por categoria com o valor real que já estava
   na conta antes de começar a usar a automação (config.SALDO_INICIAL_
   CONTABIL_TOTVS) - só acontece 1 vez, depois fica salvo no acumulado.

3. "Resumo" - compara, por categoria, o SALDO pendente (soma da aba 1, vinda
   da reconciliação) com a CONTABILIDADE acumulada do TOTVS (aba 2), com uma
   fórmula de Excel que calcula a diferença.

4. "Controle Estoque" (opcional) - cópia do controle de estoque + linha nova
   do mês, com a coluna I (qte de lotes) e a coluna L (valor dos lançamentos
   contábeis) já preenchidas a partir do relatório da reconciliação (abas
   "Controle Doadoras" e "Lançamento Contábil" - nada é recalculado). J, N e
   W ficam em branco, pra preencher à mão.

COMO USAR
---------
    python conferencia_doadoras.py --mes 2026-09 \
        --totvs dados_entrada/extrato_totvs.xlsx \
        --controle-estoque "dados_entrada/Controle Estoque.xlsx"

Por padrão o relatório da reconciliação é procurado em
saida/relatorio_reconciliacao_AAAA-MM.xlsx (ou saida_teste/ com --modo-teste -
use o mesmo modo nas duas rodadas). Se ele estiver em outro lugar, informe
com --relatorio-reconciliacao.

Sem --totvs, saem só a aba "SALDO pendente" (e o estoque, se pedido).
"""

import argparse
import copy
import os
import re
import unicodedata

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

import config
from reconciliacao import (
    estilizar_cabecalho,
    escrever_dataframe,
    encontrar_conferencia_anterior,
    carregar_linhas_ajuste_anterior,
    PREFIXO_LINHA_AJUSTE_ANTERIOR,
)

PADRAO_TRF_EST_3 = re.compile(r"Trf\.\s*Est\.\s*3\s+([^\s_]+)\s*_\s*([A-Za-zÀ-ú ]+?)(?:\s*-|\s*$)")


# ---------------------------------------------------------------------------
# 1. ABA "Saldo Gerencial" - transpõe o SALDO PENDENTE acumulado, no
#    mesmo formato em blocos por mês da "Posição Consolidada" (reaproveita
#    gerar_posicao_consolidada()).
# ---------------------------------------------------------------------------

def gerar_linhas_resumo_legado(pasta_controle):
    """Replica, na aba 'SALDO pendente', as linhas de resumo (sem doadora
    individual) que existiam na planilha de referência antes da automação:
    'MESES ANTERIORES' (saldo agregado por categoria) e 'AJUSTE' (total já
    convertido em ajuste historicamente). Cada uma semeada só 1 vez, com
    seu próprio marcador - depois disso já ficam incorporadas pra sempre
    na aba, sem precisar semear de novo."""
    linhas = []

    caminho_saldo = os.path.join(pasta_controle, config.ARQUIVO_SALDO_LEGADO_CONFERENCIA_PROCESSADO)
    if not os.path.exists(caminho_saldo):
        for categoria, valor in config.SALDO_LEGADO_MESES_ANTERIORES.items():
            linhas.append({
                "nome_bruto": None,
                "categoria": config.CATEGORIA_ROTULO_CONTABIL.get(categoria, categoria),
                "valor": valor, "data": "MESES ANTERIORES", "observacao": None,
            })
        os.makedirs(pasta_controle, exist_ok=True)
        with open(caminho_saldo, "w", encoding="utf-8") as f:
            f.write("Saldo legado de meses anteriores semeado - ver config.SALDO_LEGADO_MESES_ANTERIORES\n")
        print(f"  [info] {len(config.SALDO_LEGADO_MESES_ANTERIORES)} linhas 'MESES ANTERIORES' adicionadas - só acontece 1 vez")

    caminho_ajuste = os.path.join(pasta_controle, config.ARQUIVO_AJUSTE_LEGADO_PROCESSADO)
    if not os.path.exists(caminho_ajuste):
        for categoria, valor in config.SALDO_LEGADO_AJUSTE_ACUMULADO.items():
            linhas.append({
                "nome_bruto": None,
                "categoria": config.CATEGORIA_ROTULO_CONTABIL.get(categoria, categoria),
                "valor": valor, "data": "AJUSTE", "observacao": None,
            })
        os.makedirs(pasta_controle, exist_ok=True)
        with open(caminho_ajuste, "w", encoding="utf-8") as f:
            f.write("Ajuste legado acumulado semeado - ver config.SALDO_LEGADO_AJUSTE_ACUMULADO\n")
        print(f"  [info] {len(config.SALDO_LEGADO_AJUSTE_ACUMULADO)} linhas 'AJUSTE' (acumulado legado) adicionadas - só acontece 1 vez")

    return pd.DataFrame(linhas, columns=["nome_bruto", "categoria", "valor", "data", "observacao"])


# prefixo das linhas agregadas "sem doadora" que a reconciliação acrescenta
# 1x no SALDO pendente dela (config.SALDO_LEGADO_MESES_ANTERIORES). Aqui elas
# são ignoradas na cópia porque a conferência já tem as próprias linhas
# "MESES ANTERIORES" (gerar_linhas_resumo_legado) - senão apareceriam 2x.
PREFIXO_LINHA_LEGADO_RECONCILIACAO = "(sem doadora - saldo agregado"


def carregar_saldo_pendente_da_reconciliacao(caminho_relatorio):
    """Lê a aba 'SALDO pendente' do relatório da reconciliação do mês, do jeito
    que ela já está contabilizada (nada é recalculado). Devolve dois DataFrames:
    - os itens de saldo (sem as linhas agregadas 'sem doadora' de 1ª vez);
    - as linhas de AJUSTE que a reconciliação manteve do mês anterior (as com
      PREFIXO_LINHA_AJUSTE_ANTERIOR), com 'categoria' interna e 'valor'."""
    if not os.path.exists(caminho_relatorio):
        raise FileNotFoundError(
            f"Não encontrei o relatório da reconciliação em '{caminho_relatorio}'. "
            f"A Conferência usa o SALDO pendente que a Reconciliação já calculou, então "
            f"rode primeiro a Reconciliação Mensal deste mês (no mesmo modo - teste ou "
            f"normal) ou informe o caminho do relatório em --relatorio-reconciliacao."
        )
    df = pd.read_excel(caminho_relatorio, sheet_name="SALDO pendente")
    colunas = {"nome_bruto", "categoria", "valor", "data", "observacao"}
    faltando = colunas - set(df.columns)
    if faltando:
        raise ValueError(
            f"A aba 'SALDO pendente' de '{caminho_relatorio}' não tem as colunas "
            f"esperadas {sorted(faltando)} - o layout do relatório da reconciliação mudou?"
        )
    df = df.dropna(how="all")
    nomes = df["nome_bruto"].astype(str)
    eh_ajuste = nomes.str.startswith(PREFIXO_LINHA_AJUSTE_ANTERIOR)
    eh_legado = nomes.str.startswith(PREFIXO_LINHA_LEGADO_RECONCILIACAO)
    itens = df[~eh_ajuste & ~eh_legado].reset_index(drop=True)
    ajustes = df[eh_ajuste][["categoria", "valor"]].reset_index(drop=True)
    return itens, ajustes


def gerar_aba_saldo_pendente(ws, df_saldo, df_ajuste, pasta_controle, usar_legado):
    """Copia o SALDO pendente que a reconciliação já contabilizou (df_saldo) com o
    rótulo da categoria por extenso (ex: 'AJUDA DE CUSTO' em vez de
    'AJUDA_DE_CUSTO'). Depois dos itens vêm as linhas de AJUSTE (sem nome, 'AJUSTE'
    na coluna data) mantidas do mês anterior (df_ajuste) - o formato é o mesmo
    que a usuária usa pra criá-las à mão. Os valores NÃO são recalculados aqui.

    usar_legado=True (só quando NÃO existe conferência de mês anterior, ou seja, na
    1ª vez): acrescenta também as linhas legado 'MESES ANTERIORES' e 'AJUSTE'
    semeadas do config. Com conferência anterior elas NÃO entram: a usuária já
    ajustou isso à mão e o que vale é o que ela deixou."""
    saldo = df_saldo[["nome_bruto", "categoria", "valor", "data", "observacao"]].copy()
    saldo["categoria"] = saldo["categoria"].apply(lambda c: config.CATEGORIA_ROTULO_CONTABIL.get(c, c))
    partes = [saldo]
    if usar_legado:
        partes.append(gerar_linhas_resumo_legado(pasta_controle))
    if df_ajuste is not None and not df_ajuste.empty:
        partes.append(pd.DataFrame({
            "nome_bruto": None,
            "categoria": df_ajuste["categoria"].apply(lambda c: config.CATEGORIA_ROTULO_CONTABIL.get(c, c)),
            "valor": df_ajuste["valor"], "data": "AJUSTE", "observacao": None,
        }))
    escrever_dataframe(ws, pd.concat(partes, ignore_index=True))
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# 2. ABA "TOTVS" (extrato acumulado)
# ---------------------------------------------------------------------------

def classificar_categoria_totvs(historico):
    texto = str(historico)
    m = PADRAO_TRF_EST_3.search(texto)
    if m:
        rotulo = m.group(2).strip().upper()
        categoria = config.TOTVS_ROTULO_PARA_CATEGORIA.get(rotulo)
        if categoria:
            return categoria, "Trf. Est. 3"
    texto_upper = texto.upper()
    for categoria, palavras in config.TOTVS_PALAVRAS_CHAVE_CATEGORIA.items():
        if any(p.upper() in texto_upper for p in palavras):
            return categoria, "Outro"
    return "NAO_CLASSIFICADO", "Outro"


def _detectar_linha_cabecalho_totvs(caminho_arquivo, aba, colunas_esperadas=("DATA", "HISTORICO"), max_linhas=10):
    """Acha em qual linha da aba está o cabeçalho de verdade (a que tem as
    colunas esperadas). O extrato do TOTVS já veio com uma linha de título
    ("Lançamentos Contábeis") acima do cabeçalho em setembro/2026, e fixar a
    posição quebra toda vez que o layout muda."""
    bruto = pd.read_excel(caminho_arquivo, sheet_name=aba, header=None, nrows=max_linhas)
    for i, linha in bruto.iterrows():
        valores = set(str(v).strip().upper() for v in linha.values)
        if set(colunas_esperadas).issubset(valores):
            return i
    raise ValueError(
        f"Não encontrei o cabeçalho (colunas {list(colunas_esperadas)}) nas primeiras "
        f"{max_linhas} linhas da aba '{aba}' de '{caminho_arquivo}'. O layout do extrato "
        f"do TOTVS pode ter mudado."
    )


def _sem_acento(texto):
    return "".join(ch for ch in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(ch))


def _resolver_aba_totvs(caminho_arquivo, aba_preferida, colunas_esperadas=("DATA", "HISTORICO", "DEBITO", "CREDITO")):
    """Acha a aba do extrato que tem os lançamentos. O nome já mudou entre extratos
    ('Lançamentos Contábeis' -> 'Lancamentos Contabeis', sem acento) e o arquivo novo traz
    outras abas (Parametros, Conta, Totalizadores...). Ordem: (1) o nome preferido, sem
    diferenciar maiúscula/acento; (2) qualquer aba cujo cabeçalho tenha as colunas de
    lançamento. Se nada servir, o erro lista as abas que o arquivo tem."""
    nomes = openpyxl.load_workbook(caminho_arquivo, read_only=True).sheetnames
    alvo = _sem_acento(aba_preferida).strip().lower()
    for nome in nomes:
        if _sem_acento(nome).strip().lower() == alvo:
            return nome
    for nome in nomes:
        try:
            _detectar_linha_cabecalho_totvs(caminho_arquivo, nome, colunas_esperadas)
            return nome
        except ValueError:
            continue
    raise ValueError(
        f"Não encontrei no extrato '{os.path.basename(caminho_arquivo)}' nenhuma aba com os lançamentos "
        f"(procurei '{aba_preferida}' ou uma aba com as colunas {list(colunas_esperadas)}). "
        f"Abas do arquivo: {nomes}"
    )


COLUNAS_TOTVS = {            # nome no extrato (em maiúsculas) -> nome interno
    "DATA": "data", "HISTORICO": "historico", "LOTE/SUB/DOC/LINHA": "lote_doc",
    "DEBITO": "debito", "CREDITO": "credito", "SALDO ATUAL": "saldo_atual",
}
COLUNAS_TOTVS_OBRIGATORIAS = ("data", "historico", "debito", "credito")


def carregar_totvs(caminho_arquivo, aba="Lançamentos Contábeis"):
    aba = _resolver_aba_totvs(caminho_arquivo, aba)
    linha_cabecalho = _detectar_linha_cabecalho_totvs(caminho_arquivo, aba)
    # o lote é um identificador, não um número: lido como TEXTO. Sem isso, quando todos os
    # lotes do extrato são só dígitos o pandas converte a coluna pra inteiro (e apaga zeros à
    # esquerda), e a chave que evita lançamento duplicado entre meses deixa de bater com o
    # texto que já está na aba do mês anterior. (O nome da coluna é achado sem diferenciar
    # maiúscula/minúscula.)
    colunas = pd.read_excel(caminho_arquivo, sheet_name=aba, header=linha_cabecalho, nrows=0).columns
    col_lote = next((c for c in colunas if str(c).strip().upper() == "LOTE/SUB/DOC/LINHA"), None)
    df = pd.read_excel(caminho_arquivo, sheet_name=aba, header=linha_cabecalho,
                       dtype={col_lote: str} if col_lote is not None else None)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={c: COLUNAS_TOTVS[c.upper()] for c in df.columns if c.upper() in COLUNAS_TOTVS})
    faltando = [c for c in COLUNAS_TOTVS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ValueError(f"O extrato '{os.path.basename(caminho_arquivo)}' (aba '{aba}') não tem as colunas "
                         f"{faltando}. Colunas do arquivo: {list(df.columns)}")
    for opcional, vazio in (("lote_doc", ""), ("saldo_atual", "")):   # sem elas a conferência segue, só sem esse dado
        if opcional not in df.columns:
            df[opcional] = vazio
    df = df.dropna(subset=["historico"])
    df["debito"] = pd.to_numeric(df["debito"], errors="coerce").fillna(0)
    df["credito"] = pd.to_numeric(df["credito"], errors="coerce").fillna(0)
    # "SALDO ATUAL" vem como texto formatado pelo TOTVS (ex: "183.207,35 D",
    # onde D/C = Devedor/Credor) - mantém como veio, não converte para
    # número, pra não perder essa informação nem arriscar interpretar errado
    df["saldo_atual"] = df["saldo_atual"].astype(str).where(df["saldo_atual"].notna(), "")
    df["valor"] = df["credito"] - df["debito"]
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    classificacao = df["historico"].apply(classificar_categoria_totvs)
    df["categoria"] = classificacao.apply(lambda t: t[0])
    df["origem"] = classificacao.apply(lambda t: t[1])
    return df[["lote_doc", "data", "historico", "debito", "credito", "saldo_atual", "valor", "categoria", "origem"]]


def gerar_linhas_saldo_inicial_contabil():
    """Uma linha 'SALDO INICIAL <CATEGORIA>' por categoria, com o valor real
    que já estava na conta contábil antes de começar a usar a automação
    (config.SALDO_INICIAL_CONTABIL_TOTVS). Usa um lote_doc textual próprio
    (não um número de lançamento de verdade) só para dar um identificador
    estável a cada linha e evitar que o drop_duplicates por lote_doc as
    confunda entre si (todas teriam lote_doc vazio senão)."""
    linhas = []
    for categoria, valor in config.SALDO_INICIAL_CONTABIL_TOTVS.items():
        rotulo = config.CATEGORIA_ROTULO_CONTABIL.get(categoria, categoria)
        linhas.append({
            "lote_doc": f"SALDO_INICIAL_{categoria}",
            "data": pd.NaT,
            "historico": f"SALDO INICIAL {rotulo}",
            "debito": 0,
            "credito": 0,
            "saldo_atual": "",
            "valor": valor,
            "categoria": categoria,
            "origem": "Saldo Inicial",
        })
    return pd.DataFrame(linhas)


def semear_saldo_inicial_totvs_se_necessario(pasta_controle):
    """Só na 1ª vez que --totvs é usado: gera as linhas de SALDO INICIAL
    contábil (valores reais confirmados com a usuária) para entrar no
    acumulado. Da 2ª vez em diante, o marcador já existe e essas linhas já
    estão salvas no totvs_acumulado.csv - não precisa (nem deve) gerar de
    novo, senão duplicaria o saldo."""
    caminho_marcador = os.path.join(pasta_controle, config.ARQUIVO_SALDO_INICIAL_TOTVS_PROCESSADO)
    if os.path.exists(caminho_marcador):
        return pd.DataFrame(columns=["lote_doc", "data", "historico", "debito", "credito", "saldo_atual", "valor", "categoria", "origem"])
    print("  [info] 1ª vez com --totvs - semeando o saldo inicial contábil real (config.SALDO_INICIAL_CONTABIL_TOTVS)")
    linhas = gerar_linhas_saldo_inicial_contabil()
    os.makedirs(pasta_controle, exist_ok=True)
    with open(caminho_marcador, "w", encoding="utf-8") as f:
        f.write("Saldo inicial contábil (TOTVS) semeado - ver config.SALDO_INICIAL_CONTABIL_TOTVS\n")
    return linhas


def atualizar_totvs_acumulado(df_totvs_mes, pasta_controle):
    caminho = os.path.join(pasta_controle, config.ARQUIVO_TOTVS_ACUMULADO)
    os.makedirs(pasta_controle, exist_ok=True)
    seed = semear_saldo_inicial_totvs_se_necessario(pasta_controle)
    df_totvs_mes = pd.concat([seed, df_totvs_mes], ignore_index=True)
    # lote_doc precisa ser texto de ponta a ponta: as linhas de SALDO
    # INICIAL usam um identificador textual (ex: "SALDO_INICIAL_EXAMES_
    # CERBA"), enquanto os lançamentos reais têm lote_doc numérico. Sem
    # forçar os dois pro mesmo tipo, o CSV às vezes guarda o lote_doc real
    # como texto e outras como número entre execuções, e o
    # drop_duplicates(subset=["lote_doc"]) para de reconhecer o mesmo
    # lançamento como duplicado (duplicava lançamento a cada rodada).
    df_totvs_mes["lote_doc"] = df_totvs_mes["lote_doc"].astype(str)
    # mesmo motivo para "data": o dataframe vazio do seed (quando já foi
    # semeado antes) tem dtype genérico, e concatenar isso com a coluna
    # "data" já convertida faz ela voltar a virar dtype genérico - sem essa
    # conversão explícita, o .dt usado mais adiante (aba "Resumo") quebra.
    df_totvs_mes["data"] = pd.to_datetime(df_totvs_mes["data"], errors="coerce")
    if os.path.exists(caminho):
        acumulado = pd.read_csv(caminho, dtype={"lote_doc": str})
        acumulado["data"] = pd.to_datetime(acumulado["data"], errors="coerce")
        # reaplica as regras de classificação do config nos meses já salvos
        # (menos as linhas de SALDO INICIAL, que têm categoria fixa): assim,
        # uma regra nova em TOTVS_PALAVRAS_CHAVE_CATEGORIA também resolve os
        # lançamentos antigos que tinham ficado como NAO_CLASSIFICADO
        linhas_reais = acumulado["origem"] != "Saldo Inicial"
        reclassificado = acumulado.loc[linhas_reais, "historico"].apply(classificar_categoria_totvs)
        acumulado.loc[linhas_reais, "categoria"] = reclassificado.apply(lambda t: t[0])
        acumulado.loc[linhas_reais, "origem"] = reclassificado.apply(lambda t: t[1])
        # evita duplicar o mesmo lançamento se o extrato for processado 2x:
        # o que já estava salvo e é idêntico a uma linha do extrato novo sai,
        # e a versão nova entra no lugar. A chave NÃO pode ser só o lote_doc:
        # no TOTVS duas notas fiscais diferentes podem vir com o mesmo
        # lote_doc (já aconteceu: NF 33 do táxi e NF 14655 do laboratório), e
        # remover por lote_doc descartava uma delas em silêncio.
        acumulado = acumulado[~_chaves_totvs(acumulado).isin(set(_chaves_totvs(df_totvs_mes)))]
        acumulado = pd.concat([acumulado, df_totvs_mes], ignore_index=True)
    else:
        acumulado = df_totvs_mes
    acumulado = acumulado.sort_values("data", na_position="first")
    acumulado.to_csv(caminho, index=False)
    return acumulado


ABA_TOTVS = "TOTVS"
COLUNAS_ABA_TOTVS = ["Data", "Histórico", "Debito", "Credito", "valor", "Categoria (inferida)", "Origem", "Lote/Doc", "Saldo atual"]


def _tabela_totvs_para_aba(df_totvs):
    """Linhas do TOTVS no formato das colunas da aba (A..I)."""
    tabela = df_totvs.copy()
    # mesmo rótulo usado na aba "Resumo" (ex: "MGR", "TAXI"), para
    # dar pra fazer fórmula/PROCV direto entre as duas abas depois de um
    # ajuste manual - "NAO_CLASSIFICADO" fica como está, pois não tem
    # categoria nossa equivalente
    tabela["categoria"] = tabela["categoria"].apply(lambda c: config.CATEGORIA_ROTULO_CONTABIL.get(c, c))
    tabela = tabela.rename(columns={
        "data": "Data",
        "historico": "Histórico",
        "debito": "Debito",
        "credito": "Credito",
        "valor": "valor",
        "categoria": "Categoria (inferida)",
        "origem": "Origem",
        "lote_doc": "Lote/Doc",
        # coluna L do extrato do TOTVS ("SALDO ATUAL"). Fica como TEXTO do jeito que
        # o TOTVS mostra (ex: "169.815,99 D" - D/C = devedor/credor), sem converter
        # pra número, pra não perder o D/C nem arriscar interpretar o sinal errado
        "saldo_atual": "Saldo atual",
    })
    return tabela[COLUNAS_ABA_TOTVS]


def gerar_aba_totvs(ws, df_totvs_acumulado):
    """Aba TOTVS montada a partir do acumulado em CSV (usado quando NÃO existe
    conferência de mês anterior pra partir dela - ver gerar_aba_totvs_com_historico)."""
    escrever_dataframe(ws, _tabela_totvs_para_aba(df_totvs_acumulado))
    ws.freeze_panes = "A2"


def _chaves_totvs(df):
    """Chave que identifica um lançamento: lote_doc + data + histórico + débito +
    crédito. NÃO dá pra usar só o lote_doc: no TOTVS duas notas fiscais diferentes
    podem vir com o mesmo lote_doc (já aconteceu: NF 33 do táxi e NF 14655 do laboratório)."""
    data = pd.to_datetime(df["data"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    debito = pd.to_numeric(df["debito"], errors="coerce").fillna(0).round(2).astype(str)
    credito = pd.to_numeric(df["credito"], errors="coerce").fillna(0).round(2).astype(str)
    return (df["lote_doc"].astype(str) + "|" + data + "|" + df["historico"].astype(str)
            + "|" + debito + "|" + credito)


def _ler_linhas_aba_totvs(ws):
    """Lê os lançamentos que já estão na aba TOTVS (inclusive os que a usuária
    acrescentou à mão, como as linhas 'AJUSTE'), com o nº da linha de cada um."""
    ultima = encontrar_ultima_linha_com_dado(ws, coluna_b=2)
    linhas = []
    for r in range(2, ultima + 1):
        historico = ws.cell(row=r, column=2).value
        if historico in (None, ""):
            continue
        lote = ws.cell(row=r, column=8).value
        if isinstance(lote, float) and lote.is_integer():
            lote = int(lote)
        linhas.append({"linha": r, "data": ws.cell(row=r, column=1).value, "historico": historico,
                       "debito": ws.cell(row=r, column=3).value, "credito": ws.cell(row=r, column=4).value,
                       "lote_doc": "" if lote is None else lote, "saldo": ws.cell(row=r, column=9).value})
    return pd.DataFrame(linhas, columns=["linha", "data", "historico", "debito", "credito", "lote_doc", "saldo"])


def _saldos_do_csv_acumulado(pasta_controle):
    """{chave do lançamento: 'Saldo atual'} do acumulado em CSV, se existir -
    usado só pra preencher o saldo das linhas antigas que ainda estão sem."""
    caminho = os.path.join(pasta_controle, config.ARQUIVO_TOTVS_ACUMULADO)
    if not os.path.exists(caminho):
        return {}
    csv = pd.read_csv(caminho, dtype={"lote_doc": str})
    if "saldo_atual" not in csv.columns:
        return {}
    csv = csv[csv["saldo_atual"].notna() & (csv["saldo_atual"].astype(str).str.strip() != "")]
    if csv.empty:
        return {}
    csv["data"] = pd.to_datetime(csv["data"], errors="coerce")
    return dict(zip(_chaves_totvs(csv), csv["saldo_atual"].astype(str)))


def gerar_aba_totvs_com_historico(wb_destino, caminho_conferencia_anterior, df_totvs_mes, pasta_controle):
    """Monta a aba TOTVS PARTINDO da aba do mês anterior, exatamente como a usuária
    deixou: saldos iniciais com as contas embutidas (fórmulas), lançamentos do
    histórico e linhas de AJUSTE feitas à mão ficam INTACTOS. Só são acrescentados,
    no fim, os lançamentos do extrato novo que ainda não estão na aba (o extrato
    pode repetir lançamentos de meses anteriores - esses não são duplicados).

    Devolve (ws, qtd_linhas_que_ja_existiam, qtd_linhas_novas, qtd_saldos_preenchidos),
    ou None se o arquivo anterior não tiver a aba TOTVS (aí vale o histórico em CSV)."""
    wb_ant = openpyxl.load_workbook(caminho_conferencia_anterior, data_only=False)
    if ABA_TOTVS not in wb_ant.sheetnames:
        return None
    ws = copiar_aba(wb_destino, wb_ant[ABA_TOTVS], ABA_TOTVS, freeze_panes="A2")

    # a coluna 'Saldo atual' é nova: se o arquivo anterior ainda não tinha, cria o cabeçalho
    if ws.cell(row=1, column=9).value in (None, ""):
        cab = ws.cell(row=1, column=9, value="Saldo atual")
        ref = ws.cell(row=1, column=8)
        if ref.has_style:
            cab.font, cab.fill = copy.copy(ref.font), copy.copy(ref.fill)
            cab.border, cab.alignment = copy.copy(ref.border), copy.copy(ref.alignment)
        ws.column_dimensions["I"].width = ws.column_dimensions["H"].width or 24

    existentes = _ler_linhas_aba_totvs(ws)
    chaves_existentes = set(_chaves_totvs(existentes)) if not existentes.empty else set()

    # preenche o saldo das linhas antigas (se o acumulado em CSV ainda tiver)
    preenchidos = 0
    saldos = _saldos_do_csv_acumulado(pasta_controle)
    if saldos and not existentes.empty:
        for chave, linha, saldo_atual in zip(_chaves_totvs(existentes), existentes["linha"], existentes["saldo"]):
            if saldo_atual in (None, "") and chave in saldos:
                ws.cell(row=int(linha), column=9, value=saldos[chave])
                preenchidos += 1

    novas = df_totvs_mes[~_chaves_totvs(df_totvs_mes).isin(chaves_existentes)]
    if not novas.empty:
        scratch = Workbook().active          # reaproveita as conversões do escrever_dataframe
        escrever_dataframe(scratch, _tabela_totvs_para_aba(novas))
        primeira = encontrar_ultima_linha_com_dado(ws, coluna_b=2) + 1
        for deslocamento, linha in enumerate(scratch.iter_rows(min_row=2)):
            for celula in linha:
                destino = ws.cell(row=primeira + deslocamento, column=celula.column,
                                  value=None if celula.value == "" else celula.value)
                destino.number_format = celula.number_format
    return ws, len(existentes), len(novas), preenchidos


# ---------------------------------------------------------------------------
# 3. ABA "Resumo"
# ---------------------------------------------------------------------------

def gerar_aba_resumo(ws, ws_totvs):
    """Compara, por categoria, o SALDO pendente (Saldo Gerencial) com o TOTVS
    (Contabilidade) usando FÓRMULAS - as mesmas SUMIF que a usuária já usa -, e
    não números calculados aqui. Assim, qualquer edição manual nas abas 'SALDO
    pendente' e 'TOTVS' (linhas de AJUSTE, saldo inicial com conta
    embutida...) entra no Resumo sozinha. Todas as categorias aparecem sempre."""
    categorias = [config.CATEGORIA_ROTULO_CONTABIL[c] for c in sorted(config.CATEGORIA_ROTULO_CONTABIL)]

    ws.append(["Categoria", "Saldo Gerencial", "Contabilidade", "Diferença (Saldo Gerencial - Contabilidade)"])
    for i, rotulo in enumerate(categorias, start=2):
        ws.append([
            rotulo,
            f"=SUMIF('SALDO pendente'!B:B,Resumo!A{i},'SALDO pendente'!C:C)",
            f"=SUMIF('{ABA_TOTVS}'!F:F,Resumo!A{i},'{ABA_TOTVS}'!E:E)",
            f"=B{i}-C{i}",
        ])
    linha_total = len(categorias) + 2
    ws.append([
        "TOTAL",
        f"=SUM(B2:B{linha_total - 1})",
        f"=SUM(C2:C{linha_total - 1})",
        f"=B{linha_total}-C{linha_total}",
    ])

    # lançamentos do TOTVS que não casaram com nenhuma categoria: ficam visíveis,
    # mas fora da comparação acima (lido da própria aba TOTVS)
    qtd_nao_class, valor_nao_class = 0, 0.0
    for r in range(2, ws_totvs.max_row + 1):
        if ws_totvs.cell(row=r, column=6).value == "NAO_CLASSIFICADO":
            qtd_nao_class += 1
            v = ws_totvs.cell(row=r, column=5).value
            valor_nao_class += float(v) if isinstance(v, (int, float)) else 0.0
    if qtd_nao_class:
        ws.append([])
        ws.append([f"{qtd_nao_class} lançamento(s) do TOTVS não foram classificados em nenhuma categoria "
                   f"(veja a aba '{ABA_TOTVS}', filtre por 'NAO_CLASSIFICADO') - total: {valor_nao_class:.2f}"])

    for col in range(1, 5):
        celula = ws.cell(row=1, column=col)
        celula.font = Font(bold=True, color="FFFFFF", name="Arial")
        celula.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        celula.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 22
    for col in "BCD":
        ws.column_dimensions[col].width = 26
    for row in range(2, linha_total + 1):
        for col in "BCD":
            ws[f"{col}{row}"].number_format = "#,##0.00"
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# 4. ABA "Controle Estoque"
# ---------------------------------------------------------------------------

MESES_ABREV_PT = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}

# Colunas da linha nova do mês que continuam sendo preenchidas à mão (sem
# fórmula e sem valor vindo do relatório): J (2ª Qtde de entrada), N (Qtde de
# saída) e W (valor final da conta). Ficam em branco na linha nova.
# I (qte de lotes) e L (valor dos lançamentos contábeis) NÃO são mais manuais:
# vêm do relatório da reconciliação (ver calcular_valores_estoque_do_relatorio).
COLUNAS_MANUAIS_CONTROLE_ESTOQUE = {"J", "N", "W"}

# Texto do comentário (nota) de cada coluna na linha nova do mês.
# J não estava entre as colunas que a usuária pediu pra trocar, então continua
# como era ("mapping").
COMENTARIOS_CONTROLE_ESTOQUE = {
    "I": "qte de lotes",
    "J": "mapping",
    "L": "valor dos lançamentos contábeis",
    "N": "mapping ovodoação",
    "W": "valor final da conta final 3",
}

# Textos ANTIGOS que a automação já gravou em linhas de meses anteriores e que
# são trocados pelos novos ao copiar a aba (só troca se o texto for exatamente
# um destes - comentário escrito à mão por outra pessoa não é mexido).
COMENTARIOS_ANTIGOS_CONTROLE_ESTOQUE = {
    "I": {"mapping"},
    "L": {"gasto doadoras"},
    "N": {"mapping"},
    "W": {"saldo final da conta de estoque"},
}


def rotulo_mes_controle_estoque(mes_referencia):
    """'2026-08' -> 'Ago.26' (mesmo padrão usado na coluna B da planilha)."""
    ano, mes = mes_referencia.split("-")
    return f"{MESES_ABREV_PT[int(mes)]}.{ano[-2:]}"


def copiar_aba(wb_destino, ws_origem, nome_nova_aba, freeze_panes=None):
    """Clona uma aba inteira (valores/fórmulas + estilo de cada célula,
    comentários/notas, larguras de coluna, altura de linha e freeze panes)
    para outro workbook. Para, propositalmente, na última linha que tem
    algum dado - não arrasta linhas totalmente vazias do final da planilha
    de origem. Não lida com células mescladas porque a planilha de origem
    não tem nenhuma - se um dia passar a ter, precisa tratar aqui também."""
    ultima_linha_com_dado = 0
    for linha in ws_origem.iter_rows():
        if any(celula.value is not None for celula in linha):
            ultima_linha_com_dado = linha[0].row

    ws_nova = wb_destino.create_sheet(nome_nova_aba)
    for linha in ws_origem.iter_rows(max_row=ultima_linha_com_dado):
        for celula in linha:
            nova = ws_nova.cell(row=celula.row, column=celula.column, value=celula.value)
            if celula.has_style:
                nova.font = copy.copy(celula.font)
                nova.fill = copy.copy(celula.fill)
                nova.border = copy.copy(celula.border)
                nova.alignment = copy.copy(celula.alignment)
                nova.number_format = celula.number_format
            if celula.comment:
                nova.comment = Comment(celula.comment.text, celula.comment.author)
    for letra, dim in ws_origem.column_dimensions.items():
        ws_nova.column_dimensions[letra].width = dim.width
    for idx, dim in ws_origem.row_dimensions.items():
        if dim.height:
            ws_nova.row_dimensions[idx].height = dim.height
    ws_nova.freeze_panes = freeze_panes
    return ws_nova


def encontrar_ultima_linha_com_dado(ws, coluna_b=2):
    """Última linha com algo preenchido na coluna B (mês/'Ajuste')."""
    ultima = 0
    for linha in range(1, ws.max_row + 1):
        if ws.cell(row=linha, column=coluna_b).value not in (None, ""):
            ultima = linha
    return ultima


def calcular_valores_estoque_do_relatorio(caminho_relatorio, mes_referencia):
    """Lê do relatório da reconciliação (sem recalcular nada) os dois números
    da linha nova do Controle Estoque:

    - qte de lotes (coluna I): soma de 'Lotes (Fresco)' + 'Lotes (Congelados/
      Banco)' + 'Mini lotes' (config.MINILOTES_ENTRAM_NA_QTE_DE_LOTES) das doadoras da
      aba 'Controle Doadoras' cuja 'Data (Mês)' é o mês sendo processado. Relatório
      antigo, sem a coluna 'Mini lotes', continua valendo (mini lotes = 0).
    - valor dos lançamentos contábeis (coluna L): soma do 'valor' da aba
      'Lançamento Contábil' - só das linhas que têm texto de lançamento, pra
      não somar uma eventual linha de total no rodapé (o relatório de agosto
      tinha uma, e somar ela dobrava o valor).

    Devolve (lotes_mes, valor_lancamentos); cada um é None se não achou nada
    daquele mês (aí a célula fica em branco, em vez de gravar um zero falso)."""
    cd = pd.read_excel(caminho_relatorio, sheet_name="Controle Doadoras")
    meses = pd.to_datetime(cd["Data (Mês)"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m")
    do_mes = cd[meses == mes_referencia]
    lotes_mes = None
    if do_mes.empty:
        print(f"  [aviso] nenhuma doadora com 'Data (Mês)' de {mes_referencia} na aba 'Controle Doadoras' - "
              f"a coluna I (qte de lotes) fica em branco")
    else:
        fresco = pd.to_numeric(do_mes["Lotes (Fresco)"], errors="coerce").fillna(0).sum()
        congelados = pd.to_numeric(do_mes["Lotes (Congelados/Banco)"], errors="coerce").fillna(0).sum()
        minilotes = 0
        if "Mini lotes" in do_mes.columns:
            minilotes = pd.to_numeric(do_mes["Mini lotes"], errors="coerce").fillna(0).sum()
        entra = config.MINILOTES_ENTRAM_NA_QTE_DE_LOTES
        lotes_mes = int(round(fresco + congelados + (minilotes if entra else 0)))
        detalhe = f"{int(round(fresco))} frescos + {int(round(congelados))} congelados + {int(round(minilotes))} mini lotes"
        if minilotes and not entra:
            detalhe += " (mini lotes NÃO contados - config.MINILOTES_ENTRAM_NA_QTE_DE_LOTES = False)"
        print(f"   lotes de {mes_referencia}: {lotes_mes} ({len(do_mes)} doadoras: {detalhe})")

    lc = pd.read_excel(caminho_relatorio, sheet_name="Lançamento Contábil")
    linhas = lc[lc["lancamento"].notna() & (lc["lancamento"].astype(str).str.strip() != "")]
    # o texto de cada lançamento termina com o mês (ex: "... - SETEMBRO/2026"):
    # só entram as linhas do mês pedido, pra um relatório de OUTRO mês (passado
    # por engano em --relatorio-reconciliacao) não virar o valor deste mês
    ano, mes = mes_referencia.split("-")
    sufixo_mes = f"{config.MESES_PT[int(mes)].upper()}/{ano}"
    do_mes_certo = linhas[linhas["lancamento"].astype(str).str.upper().str.contains(sufixo_mes, regex=False)]
    valor_lancamentos = None
    if do_mes_certo.empty:
        print(f"  [aviso] nenhum lançamento de {sufixo_mes} na aba 'Lançamento Contábil' - esse relatório é do "
              f"mês certo? A coluna L (valor dos lançamentos contábeis) fica em branco")
    else:
        if len(do_mes_certo) < len(linhas):
            print(f"  [aviso] {len(linhas) - len(do_mes_certo)} lançamento(s) da aba não são de {sufixo_mes} e "
                  f"ficaram fora da soma")
        valor_lancamentos = float(pd.to_numeric(do_mes_certo["valor"], errors="coerce").fillna(0).sum())
        print(f"   lançamentos contábeis de {mes_referencia}: R$ {valor_lancamentos:,.2f} ({len(do_mes_certo)} lançamentos)")
    return lotes_mes, valor_lancamentos


def atualizar_comentarios_existentes(ws):
    """Troca, nas linhas de meses anteriores, os textos antigos dos comentários
    das colunas I, L, N e W pelos novos (COMENTARIOS_CONTROLE_ESTOQUE)."""
    trocados = 0
    for letra, antigos in COMENTARIOS_ANTIGOS_CONTROLE_ESTOQUE.items():
        for linha in range(1, ws.max_row + 1):
            celula = ws[f"{letra}{linha}"]
            if celula.comment and celula.comment.text.strip().lower() in antigos:
                celula.comment = Comment(COMENTARIOS_CONTROLE_ESTOQUE[letra], celula.comment.author)
                trocados += 1
    return trocados


def adicionar_linha_do_mes(ws, mes_referencia, lotes_mes=None, valor_lancamentos=None):
    """Acrescenta, logo após a última linha com dado, uma linha nova pro mês
    sendo processado - com as MESMAS fórmulas da última linha real (ajustadas
    pra apontar pra linha nova/anterior). A coluna I recebe a qte de lotes do
    mês e a coluna L o valor dos lançamentos contábeis (ambos vindos do
    relatório da reconciliação); J, N e W ficam em branco, pra preencher à
    mão. Funciona também se a linha anterior for um 'Ajuste' (as fórmulas só
    olham pra 'linha-1', não pro mês em si, então continuam corretas nesse
    caso)."""
    linha_anterior = encontrar_ultima_linha_com_dado(ws)
    linha_nova = linha_anterior + 1

    formulas = {
        "C": f"=R{linha_anterior}",
        "D": f'=IF(J{linha_anterior}="",R{linha_anterior},S{linha_anterior})',
        "E": f"=IFERROR(F{linha_nova}/D{linha_nova},0)",
        "F": f"=U{linha_anterior}",
        "K": f"=L{linha_nova}/I{linha_nova}",
        "O": f"=-(F{linha_nova}+L{linha_nova})/(C{linha_nova}+J{linha_nova})",
        "P": f"=N{linha_nova}*-O{linha_nova}",
        "R": f"=C{linha_nova}+I{linha_nova}+N{linha_nova}",
        "S": f"=D{linha_nova}+J{linha_nova}+N{linha_nova}",
        "T": f"=IFERROR(U{linha_nova}/R{linha_nova},0)",
        "U": f"=F{linha_nova}+L{linha_nova}+P{linha_nova}",
        "X": f"=ROUND(U{linha_nova}-W{linha_nova},0)",
    }

    ws.cell(row=linha_nova, column=2, value=rotulo_mes_controle_estoque(mes_referencia))  # B
    ws.cell(row=linha_nova, column=22, value="razão:")  # V

    for letra, formula in formulas.items():
        ws[f"{letra}{linha_nova}"] = formula

    for letra in COLUNAS_MANUAIS_CONTROLE_ESTOQUE:
        ws[f"{letra}{linha_nova}"] = None
    ws[f"I{linha_nova}"] = lotes_mes            # qte de lotes (None = em branco)
    ws[f"L{linha_nova}"] = valor_lancamentos    # valor dos lançamentos contábeis (None = em branco)

    autor = "Automação Doadoras"
    for letra, texto in COMENTARIOS_CONTROLE_ESTOQUE.items():
        ws[f"{letra}{linha_nova}"].comment = Comment(texto, autor)

    # copia o estilo da linha anterior pra linha nova ficar visualmente igual
    for col in range(1, ws.max_column + 1):
        origem = ws.cell(row=linha_anterior, column=col)
        destino = ws.cell(row=linha_nova, column=col)
        if origem.has_style:
            destino.font = copy.copy(origem.font)
            destino.fill = copy.copy(origem.fill)
            destino.border = copy.copy(origem.border)
            destino.alignment = copy.copy(origem.alignment)
            destino.number_format = origem.number_format

    return linha_nova


def gerar_aba_controle_estoque(wb_destino, caminho_arquivo, mes_referencia, lotes_mes=None,
                               valor_lancamentos=None, nome_aba_origem="Controle de Estoque"):
    wb_origem = openpyxl.load_workbook(caminho_arquivo, data_only=False)
    ws_origem = wb_origem[nome_aba_origem]
    # congela as linhas 1-7 (cabeçalho) e as colunas A-B (mês), pra dar pra rolar
    # até os últimos meses sem perder de vista o que cada coluna significa. Ajuste
    # esse valor se a planilha de origem mudar de novo quantas linhas de
    # cabeçalho tem
    ws_nova = copiar_aba(wb_destino, ws_origem, "Controle Estoque", freeze_panes="C8")
    trocados = atualizar_comentarios_existentes(ws_nova)
    if trocados:
        print(f"   {trocados} comentário(s) de meses anteriores atualizados para o texto novo")
    linha_nova = adicionar_linha_do_mes(ws_nova, mes_referencia, lotes_mes, valor_lancamentos)
    return linha_nova


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Conferência de doadoras: SALDO pendente (da reconciliação) + TOTVS + resumo")
    parser.add_argument("--mes", required=True)
    parser.add_argument("--relatorio-reconciliacao", default=None,
                         help="Caminho do relatório da reconciliação do mês, de onde vem o SALDO pendente. "
                              "Padrão: saida/relatorio_reconciliacao_AAAA-MM.xlsx (saida_teste/ com --modo-teste).")
    # --lista-doadoras, --custos e --medicamentos-liberado não são mais usados (a
    # conferência não recalcula nada); continuam aceitos só pra não quebrar quem
    # ainda chama o script com eles
    parser.add_argument("--lista-doadoras", default=None, help=argparse.SUPPRESS)
    parser.add_argument("--custos", default=None, help=argparse.SUPPRESS)
    parser.add_argument("--medicamentos-liberado", choices=["sim", "nao"], default="nao", help=argparse.SUPPRESS)
    parser.add_argument("--totvs", default=None,
                         help="Caminho do extrato do TOTVS (conta contábil) para este mês. "
                              "Opcional - sem isso, só a aba 'SALDO pendente' é gerada.")
    parser.add_argument("--controle-estoque", default=None,
                         help="Caminho do arquivo 'Controle Estoque Único' (aba 'Controle de "
                              "Estoque'). Opcional - sem isso, a aba 'Controle Estoque' não é gerada.")
    parser.add_argument("--conferencia-anterior", default=None,
                         help="Caminho da conferência do mês anterior (conferencia_doadoras_AAAA-MM.xlsx, já com "
                              "as suas edições manuais). A aba TOTVS do mês novo PARTE da aba TOTVS dela (saldos "
                              "iniciais com contas, histórico, linhas de AJUSTE) e só acrescenta os lançamentos "
                              "novos. Padrão: o conferencia_doadoras_*.xlsx mais recente (mês anterior ao "
                              "processado) da pasta saida/ (saida_teste/ com --modo-teste).")
    parser.add_argument("--modo-teste", action="store_true",
                         help="Usa 'controle_teste/'/'saida_teste/' em vez de 'controle/'/'saida/'.")
    args = parser.parse_args()

    pasta_controle = "controle_teste" if args.modo_teste else "controle"
    pasta_saida = "saida_teste" if args.modo_teste else "saida"
    if args.modo_teste:
        print(f"[MODO TESTE] usando pasta '{pasta_controle}/' - o histórico real não será alterado.")

    print(f"== Conferência Doadoras {args.mes} ==")
    if args.lista_doadoras or args.custos:
        print("  [info] --lista-doadoras/--custos são ignorados: a conferência não recalcula o mês, "
              "usa o SALDO pendente que a reconciliação já gerou.")

    caminho_relatorio = args.relatorio_reconciliacao or f"{pasta_saida}/relatorio_reconciliacao_{args.mes}.xlsx"
    print(f"1. Lendo o SALDO pendente já contabilizado em '{caminho_relatorio}'...")
    df_saldo, df_ajuste = carregar_saldo_pendente_da_reconciliacao(caminho_relatorio)
    print(f"   {len(df_saldo)} itens de saldo pendente (R$ {df_saldo['valor'].sum():,.2f}) - copiados sem recalcular")

    caminho_conf_anterior = encontrar_conferencia_anterior(pasta_saida, args.mes, args.conferencia_anterior)
    if caminho_conf_anterior:
        print(f"   conferência do mês anterior: '{caminho_conf_anterior}' (a aba TOTVS parte dela)")
        n_ajuste_anterior = len(carregar_linhas_ajuste_anterior(caminho_conf_anterior))
        if len(df_ajuste):
            print(f"   {len(df_ajuste)} linha(s) de AJUSTE mantida(s) do mês anterior (R$ {df_ajuste['valor'].sum():,.2f})")
        elif n_ajuste_anterior:
            print(f"  [info] a conferência anterior tinha {n_ajuste_anterior} linha(s) de AJUSTE e o relatório da "
                  f"reconciliação deste mês não tem nenhuma - se você as retirou à mão, tudo certo; senão, rode a "
                  f"reconciliação de novo (com a versão mais nova) antes da conferência")
    else:
        print("   (nenhuma conferência de mês anterior encontrada - 1ª vez: valem os saldos legado do config e o histórico em CSV)")

    wb = Workbook()
    wb.remove(wb.active)

    print("2. Gerando aba 'SALDO pendente'...")
    gerar_aba_saldo_pendente(wb.create_sheet("SALDO pendente"), df_saldo, df_ajuste, pasta_controle,
                             usar_legado=caminho_conf_anterior is None)

    if args.totvs:
        print("3. Lendo extrato do TOTVS...")
        df_totvs_mes = carregar_totvs(args.totvs)
        resultado = None
        if caminho_conf_anterior:
            resultado = gerar_aba_totvs_com_historico(wb, caminho_conf_anterior, df_totvs_mes, pasta_controle)
            if resultado is None:
                print(f"  [aviso] '{caminho_conf_anterior}' não tem a aba '{ABA_TOTVS}' - usando o histórico em CSV")
        if resultado is not None:
            ws_totvs, n_existiam, n_novas, n_saldos = resultado
            print(f"   aba TOTVS partiu da do mês anterior ({n_existiam} linhas mantidas, inclusive saldos iniciais e AJUSTES "
                  f"suas) + {n_novas} lançamento(s) novo(s) deste extrato ({len(df_totvs_mes) - n_novas} já estavam)")
            if n_saldos:
                print(f"   'Saldo atual' preenchido em {n_saldos} linha(s) antiga(s) a partir do acumulado em CSV")
        else:
            df_totvs_acumulado = atualizar_totvs_acumulado(df_totvs_mes, pasta_controle)
            print(f"   {len(df_totvs_mes)} lançamentos neste extrato | {len(df_totvs_acumulado)} no acumulado total (CSV)")
            ws_totvs = wb.create_sheet(ABA_TOTVS)
            gerar_aba_totvs(ws_totvs, df_totvs_acumulado)

        print("4. Gerando aba 'Resumo' (fórmulas SUMIF: SALDO pendente x TOTVS)...")
        gerar_aba_resumo(wb.create_sheet("Resumo"), ws_totvs)
    else:
        print("3. Sem --totvs - pulando abas 'TOTVS' e 'Resumo'.")

    if args.controle_estoque:
        print("5. Gerando aba 'Controle Estoque' (cópia + linha nova do mês)...")
        print(f"   lendo lotes e lançamentos contábeis do relatório da reconciliação (sem recalcular)...")
        lotes_mes, valor_lancamentos = calcular_valores_estoque_do_relatorio(caminho_relatorio, args.mes)
        linha_nova = gerar_aba_controle_estoque(wb, args.controle_estoque, args.mes, lotes_mes, valor_lancamentos)
        print(f"   linha {linha_nova} adicionada (B{linha_nova} = '{rotulo_mes_controle_estoque(args.mes)}') - "
              f"colunas I e L preenchidas pelo relatório; J, N, W em branco pra preencher à mão")
    else:
        print("5. Sem --controle-estoque - pulando aba 'Controle Estoque'.")

    os.makedirs(pasta_saida, exist_ok=True)
    caminho_saida = f"{pasta_saida}/conferencia_doadoras_{args.mes}.xlsx"
    wb.save(caminho_saida)
    print(f"6. Salvo em: {caminho_saida}")


if __name__ == "__main__":
    main()
