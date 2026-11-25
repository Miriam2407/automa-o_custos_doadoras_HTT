"""
Reconciliação mensal de custos de doadoras de óvulos
======================================================

FLUXO (em ordem):

1. Lê a planilha mensal do estagiário (doadoras do mês) e acumula num
   cadastro mestre histórico, para que uma doadora continue "reconhecível"
   pelo matching nos meses seguintes.

2. Lê os arquivos de custo POR DOADORA (Uber, Táxi, Cerba, Ajuda
   de Custo e, quando liberado, Medicamentos/4BIO) e casa cada lançamento
   com uma doadora do cadastro mestre via nome aproximado.

3. Lê o arquivo de CUSTOS RATEADOS (CUSTOS RATEADOS: Médicos, Insumos, MGR),
   que vem como total fechado do mês sem nome — e divide cada total pela
   quantidade de doadoras do mês, lançando a fração para cada uma.

4. Quando um lançamento por doadora não bate com ninguém:
   - 1ª vez -> "SALDO", com a data salva em controle/controle_saldo.csv
     (esse arquivo persiste entre os meses - não apagar).
   - já pendente há 3 meses ou mais -> vira "AJUSTE DD/MM/AAAA" sozinho.

5. Gera um Excel com: uma aba por categoria, consolidado por doadora,
   aba de SALDO, aba de AJUSTE e uma aba de conferência contábil.

COMO USAR
---------
    python reconciliacao.py --mes 2026-07 \
        --lista-doadoras dados_entrada/lista_doadoras_2026-07.xlsx \
        --custos dados_entrada \
        --medicamentos-liberado nao

AJUSTES PARA O SEU CASO REAL: tudo em config.py (colunas, linha de
cabeçalho de cada arquivo, regra de rateio). Esse é o único arquivo que
normalmente precisa editar.
"""

import argparse
import os
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

import pandas as pd
from rapidfuzz import fuzz, process, utils
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config


# ---------------------------------------------------------------------------
# 1. CADASTRO MESTRE DE DOADORAS
# ---------------------------------------------------------------------------

def _detectar_linha_cabecalho(caminho_arquivo, colunas_esperadas, max_linhas_procuradas=10):
    """Acha automaticamente em qual linha está o cabeçalho de verdade,
    procurando pela linha que contém TODAS as colunas esperadas. Mais
    resistente do que fixar um número de linha (MASTER_HEADER_ROW) - a
    planilha da lista de doadoras já mudou de layout uma vez (a linha de
    título no topo foi removida em setembro/2026), e fixar a posição
    quebra toda vez que isso acontece de novo."""
    bruto = pd.read_excel(caminho_arquivo, header=None, nrows=max_linhas_procuradas)
    for i, linha in bruto.iterrows():
        valores = set(str(v).strip() for v in linha.values)
        if set(colunas_esperadas).issubset(valores):
            return i
    raise ValueError(
        f"Não encontrei uma linha de cabeçalho com as colunas {colunas_esperadas} "
        f"nas primeiras {max_linhas_procuradas} linhas de '{caminho_arquivo}'. "
        f"O layout da planilha pode ter mudado - confira se os nomes das colunas "
        f"ainda são esses (config.MASTER_COLUNA_*)."
    )


def _avisar_pin_repetido(df, mes_referencia):
    """O PIN deveria ser único por doadora. Repetido na lista quase sempre é erro de digitação
    (a automação NÃO corrige sozinha: não tem como saber qual dos dois PINs está certo)."""
    repetidos = df[df.duplicated("codigo", keep=False)]
    for pin, grupo in repetidos.groupby("codigo"):
        nomes = grupo["nome"].tolist()
        pin_txt = pin[:-2] if pin.endswith(".0") else pin
        if len(set(nomes)) > 1:
            print(f"  [aviso] o PIN {pin_txt} aparece para {len(set(nomes))} doadoras DIFERENTES em {mes_referencia}: "
                  f"{'; '.join(nomes)}. Provável erro de digitação do PIN na lista - corrija lá. "
                  f"(As duas foram mantidas.)")
        else:
            print(f"  [aviso] '{nomes[0]}' (PIN {pin_txt}) aparece {len(nomes)}x na lista de {mes_referencia} - "
                  f"se for repetição por engano, apague a linha; as {len(nomes)} contam no rateio e nos lotes")


def carregar_lista_doadoras_do_mes(caminho_arquivo, mes_referencia):
    """Lê a lista de doadoras e devolve SÓ as doadoras do mês processado.

    Dois layouts são aceitos:
    - lista ANTIGA: um mês por arquivo, sem a coluna 'Mini lote' (mini lotes = 0);
    - lista NOVA: com a coluna 'Mini lote' e podendo ACUMULAR vários meses na mesma aba.
    Como o número de doadoras do mês é o divisor do rateio (insumos, MGR, médicos), as
    linhas de OUTROS meses nunca podem entrar: é por isso que o filtro por 'Mês' existe."""
    colunas_esperadas = [
        config.MASTER_COLUNA_CODIGO, config.MASTER_COLUNA_NOME,
        config.MASTER_COLUNA_LOTES_DOADOS, config.MASTER_COLUNA_LOTES_CONGELADOS,
    ]   # 'Mini lote' fica de fora de propósito: a lista antiga não tem
    linha_cabecalho = _detectar_linha_cabecalho(caminho_arquivo, colunas_esperadas)
    df = pd.read_excel(caminho_arquivo, header=linha_cabecalho)
    df = df.rename(columns={
        config.MASTER_COLUNA_CODIGO: "codigo",
        config.MASTER_COLUNA_NOME: "nome",
        config.MASTER_COLUNA_LOTES_DOADOS: "lotes_doados",
        config.MASTER_COLUNA_LOTES_CONGELADOS: "lotes_congelados",
        config.MASTER_COLUNA_MINILOTES: "minilotes",
        config.MASTER_COLUNA_MES: "data_pertencimento",
    })
    if "minilotes" not in df.columns:
        df["minilotes"] = 0
    df["minilotes"] = pd.to_numeric(df["minilotes"], errors="coerce").fillna(0).astype(int)
    mes_da_linha = pd.to_datetime(df["data_pertencimento"], errors="coerce")

    # linhas com PIN e mês mas SEM nome: não dá pra casar custo com elas - ficam de fora
    # (como sempre), mas agora a usuária é avisada pra preencher o nome
    sem_nome = df[df["codigo"].notna() & df["nome"].isna() & mes_da_linha.notna()]

    # a planilha do estagiário costuma ter linhas de rodapé (totais, notas de
    # relatório tipo "Total Ciclo >>>") depois da última doadora - descarta
    # qualquer linha sem código, sem nome, ou cujo nome pareça rodapé/nota
    df = df.dropna(subset=["codigo", "nome"])
    df = df[~df["nome"].astype(str).str.contains(">", regex=False)]
    mes_da_linha = pd.to_datetime(df["data_pertencimento"], errors="coerce")
    meses_na_lista = sorted(mes_da_linha.dropna().dt.strftime("%Y-%m").unique())

    # filtro por mês (ver docstring). Só descarta linhas de OUTRO mês; linha com nome e
    # sem 'Mês' continua valendo como do mês (era assim na lista antiga).
    if meses_na_lista:
        do_mes = mes_da_linha.dt.strftime("%Y-%m") == mes_referencia
        if do_mes.any():
            fora = df[mes_da_linha.notna() & ~do_mes]
            if len(fora):
                outros = sorted(mes_da_linha[fora.index].dt.strftime("%Y-%m").unique())
                print(f"   [info] a lista tem {len(fora)} doadora(s) de outros meses ({', '.join(outros)}) - "
                      f"ignoradas; só as de {mes_referencia} foram usadas")
            sem_mes = int(mes_da_linha.isna().sum())
            if sem_mes:
                print(f"  [aviso] {sem_mes} doadora(s) com nome e SEM 'Mês' foram consideradas deste mês")
            df = df[do_mes | mes_da_linha.isna()]
            sem_nome = sem_nome[pd.to_datetime(sem_nome["data_pertencimento"], errors="coerce")
                                .dt.strftime("%Y-%m") == mes_referencia]
        elif len(meses_na_lista) == 1:
            # lista de um mês só, mas de outro mês que o processado: comportamento antigo (usa tudo)
            print(f"  [ATENÇÃO] a lista é de {meses_na_lista[0]}, mas o mês processado é {mes_referencia}. "
                  f"Usei TODAS as {len(df)} doadoras dela (comportamento da lista antiga): o rateio de "
                  f"insumos/MGR/médicos vai dividir por {len(df)} e o cadastro vai registrá-las como de "
                  f"{mes_referencia}. Se essa NÃO é a lista de {mes_referencia}, interrompa e use a certa")
        else:
            raise ValueError(
                f"A lista de doadoras tem doadoras dos meses {', '.join(meses_na_lista)}, mas nenhuma de "
                f"{mes_referencia}. Confira o mês informado (--mes) ou se a coluna 'Mês' foi preenchida."
            )
    if len(sem_nome):
        pins = ", ".join(str(int(p)) if float(p).is_integer() else str(p) for p in sem_nome["codigo"].tolist())
        print(f"  [aviso] {len(sem_nome)} linha(s) de {mes_referencia} têm PIN mas estão SEM NOME e foram ignoradas "
              f"(PIN: {pins}) - preencha o nome na lista pra elas entrarem no matching e no rateio")

    df["nome"] = df["nome"].astype(str).str.strip()
    df["codigo"] = df["codigo"].astype(str).str.strip()
    _avisar_pin_repetido(df, mes_referencia)
    # a coluna K (Mês) vem como data (ex: 2026-07-01 00:00:00) - formata pro
    # padrão DD/MM/AAAA que é usado no resto do relatório
    df["data_pertencimento"] = pd.to_datetime(df["data_pertencimento"], errors="coerce").dt.strftime("%d/%m/%Y")
    df["mes_referencia"] = mes_referencia
    return df[["codigo", "nome", "lotes_doados", "lotes_congelados", "minilotes", "data_pertencimento", "mes_referencia"]]


def atualizar_cadastro_mestre(caminho_historico, df_mes_atual):
    os.makedirs(os.path.dirname(caminho_historico), exist_ok=True)
    if os.path.exists(caminho_historico):
        historico = pd.read_csv(caminho_historico, dtype=str)
        historico = pd.concat([historico, df_mes_atual.astype(str)], ignore_index=True)
        # keep="last" é importante: se essa doadora/mês já existia no arquivo
        # antigo (de uma versão anterior do script, por ex. sem a coluna de
        # data), a entrada NOVA que acabamos de carregar tem que prevalecer
        historico = historico.drop_duplicates(subset=["codigo", "nome", "mes_referencia"], keep="last")
    else:
        historico = df_mes_atual.astype(str)
    historico.to_csv(caminho_historico, index=False)
    return historico


# ---------------------------------------------------------------------------
# 2. LEITURA DOS ARQUIVOS DE CUSTO POR DOADORA
# ---------------------------------------------------------------------------

def identificar_arquivo(nome_arquivo, mapa_padroes):
    nome_lower = nome_arquivo.lower()
    for categoria, cfg in mapa_padroes.items():
        if any(p in nome_lower for p in cfg["arquivo_padrao"]):
            return categoria
    return None


def resolver_aba_por_mes(caminho_arquivo, mes_referencia):
    """Para arquivos com 1 aba por mês (ex: 'Agosto2026'), tenta achar a
    aba certa mesmo com formato de ano inconsistente (26 vs 2026)."""
    ano, mes = mes_referencia.split("-")
    mes_nome = config.MESES_PT[int(mes)]
    candidatos = [f"{mes_nome}{ano}", f"{mes_nome}{ano[2:]}"]

    abas_disponiveis = load_workbook(caminho_arquivo, read_only=True).sheetnames
    for candidato in candidatos:
        if candidato in abas_disponiveis:
            return candidato
    raise ValueError(
        f"Não encontrei a aba do mês {mes_referencia} em '{caminho_arquivo}'. "
        f"Tentei: {candidatos}. Abas disponíveis: {abas_disponiveis}"
    )


def _achar_coluna(df, nome_ou_alternativas):
    """Acha a coluna do DataFrame ignorando maiúsculas/minúsculas e espaços
    nas pontas (ex: 'Tipo' acha 'TIPO' e 'Tipo '). Aceita um nome só ou uma
    lista de nomes alternativos (usa o primeiro que existir no arquivo).
    Devolve o nome REAL da coluna no arquivo, ou None se nenhum existir."""
    alternativas = (
        [nome_ou_alternativas]
        if isinstance(nome_ou_alternativas, str)
        else list(nome_ou_alternativas)
    )
    por_nome_limpo = {str(c).strip().lower(): c for c in df.columns}
    for alt in alternativas:
        achada = por_nome_limpo.get(str(alt).strip().lower())
        if achada is not None:
            return achada
    return None


def _exigir_coluna(df, nome_ou_alternativas, caminho_arquivo, papel):
    col = _achar_coluna(df, nome_ou_alternativas)
    if col is None:
        raise KeyError(
            f"Não encontrei a coluna de {papel} em '{os.path.basename(caminho_arquivo)}'. "
            f"Procurei por: {nome_ou_alternativas}. "
            f"Colunas do arquivo: {list(df.columns)}"
        )
    return col


def normalizar_arquivo_custo(caminho_arquivo, categoria, mes_referencia):
    cfg = config.FONTES_DE_CUSTO[categoria]

    sheet_name = cfg.get("sheet_name", 0)
    if cfg.get("aba_por_mes"):
        sheet_name = resolver_aba_por_mes(caminho_arquivo, mes_referencia)

    df = pd.read_excel(caminho_arquivo, header=cfg["header_row"], sheet_name=sheet_name)

    if "filtro" in cfg:
        f = cfg["filtro"]
        valores_aceitos = f.get("valores") or [f.get("valor")]
        col_filtro = _exigir_coluna(df, f["coluna"], caminho_arquivo, "filtro")
        df = df[df[col_filtro].isin(valores_aceitos)]

    # Acha as colunas pelo nome sem diferenciar maiúscula/minúscula. "coluna_valor"
    # pode ser uma lista de alternativas (o layout do arquivo já mudou antes).
    col_nome = _exigir_coluna(df, cfg["coluna_nome"], caminho_arquivo, "nome")
    col_valor = _exigir_coluna(df, cfg["coluna_valor"], caminho_arquivo, "valor")
    col_data = _achar_coluna(df, cfg["coluna_data"]) if cfg.get("coluna_data") else None

    renomear = {col_nome: "nome_bruto", col_valor: "valor"}
    if col_data is not None:
        renomear[col_data] = "data"
    df = df.rename(columns=renomear)
    if "data" not in df.columns:
        df["data"] = None
    else:
        df["data"] = df["data"].apply(converter_data_excel)

    df = df.dropna(subset=["nome_bruto"])
    df["nome_bruto"] = df["nome_bruto"].astype(str).str.strip()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df = df.dropna(subset=["valor"])
    df["categoria"] = categoria
    return df[["nome_bruto", "valor", "data", "categoria"]]


def carregar_todos_custos_por_doadora(pasta_custos, mes_referencia, rodar_medicamentos):
    partes = []
    for arquivo in os.listdir(pasta_custos):
        if arquivo.lower().startswith("lista_doadoras"):
            continue
        categoria = identificar_arquivo(arquivo, config.FONTES_DE_CUSTO)
        if categoria is None:
            continue  # pode ser o arquivo de custos rateados, tratado à parte
        if categoria == "MEDICAMENTOS_4BIO" and not rodar_medicamentos:
            print(f"  [info] pulando '{arquivo}' - medicamentos só entra após fechamento "
                  f"contábil (use --medicamentos-liberado sim quando estiver liberado)")
            continue
        caminho = os.path.join(pasta_custos, arquivo)
        print(f"  [ok] lendo '{arquivo}' como {categoria}")
        partes.append(normalizar_arquivo_custo(caminho, categoria, mes_referencia))
    if not partes:
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])
    return pd.concat(partes, ignore_index=True)


# ---------------------------------------------------------------------------
# 3. CUSTOS RATEADOS (Médicos / Insumos / MGR) - sem nome no arquivo de
#    origem, valor dividido igualmente entre as doadoras do mês
# ---------------------------------------------------------------------------

def avisar_arquivo_custos_rateados_nao_encontrado(pasta_custos, consequencia):
    """Avisa (em vez de ignorar em silêncio) quando o arquivo de custos rateados
    não foi achado na pasta. Se existir um arquivo com o NOME ANTIGO (ver config.py), o aviso diz pra renomear."""
    antigos = [a for a in os.listdir(pasta_custos)
               if any(p in a.lower() for p in config.ARQUIVO_CUSTOS_RATEADOS_NOMES_ANTIGOS)]
    if antigos:
        print(f"  [aviso] '{antigos[0]}' está com o NOME ANTIGO e foi ignorado - renomeie para "
              f"'{config.NOME_ARQUIVO_CUSTOS_RATEADOS}'. {consequencia} este mês")
    else:
        print(f"  [aviso] arquivo '{config.NOME_ARQUIVO_CUSTOS_RATEADOS}' não encontrado em '{pasta_custos}' - "
              f"{consequencia} este mês")


def carregar_custos_rateados(pasta_custos, df_mes, mes_referencia):
    linhas = []
    qtd_doadoras = len(df_mes)
    if qtd_doadoras == 0:
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])

    achou_arquivo = False
    for arquivo in os.listdir(pasta_custos):
        categoria_arquivo = identificar_arquivo(arquivo, config.ARQUIVO_CUSTOS_RATEADOS)
        if categoria_arquivo is None:
            continue
        achou_arquivo = True
        cfg = config.ARQUIVO_CUSTOS_RATEADOS[categoria_arquivo]
        caminho = os.path.join(pasta_custos, arquivo)
        print(f"  [ok] lendo '{arquivo}' como custo rateado ({qtd_doadoras} doadoras no mês)")
        df = pd.read_excel(caminho, header=cfg["header_row"], sheet_name=cfg["sheet_name"])
        df = df.rename(columns={
            cfg["coluna_categoria"]: "categoria_arquivo",
            cfg["coluna_valor"]: "valor_total",
            cfg["coluna_data"]: "data",
        })
        df["data"] = df["data"].apply(converter_data_excel)

        for _, linha in df.iterrows():
            mapeamento = cfg["mapa_categoria"].get(str(linha["categoria_arquivo"]).strip())
            if mapeamento is None:
                print(f"  [aviso] categoria '{linha['categoria_arquivo']}' de '{arquivo}' não "
                      f"está mapeada em config.py (ARQUIVO_CUSTOS_RATEADOS) - ignorada")
                continue
            categoria_interna, operacao = mapeamento
            if operacao == "dividir":
                valor_por_doadora = round(float(linha["valor_total"]) / qtd_doadoras, 2)
            else:  # "nenhuma" - cada doadora recebe o valor exatamente como está no arquivo
                valor_por_doadora = round(float(linha["valor_total"]), 2)
            for nome_doadora in df_mes["nome"]:
                linhas.append({
                    "nome_bruto": nome_doadora,
                    "valor": valor_por_doadora,
                    "data": linha["data"],
                    "categoria": categoria_interna,
                })
    if not achou_arquivo:
        avisar_arquivo_custos_rateados_nao_encontrado(pasta_custos, "INSUMOS e MGR não serão lançados")
    return pd.DataFrame(linhas)


def carregar_medicos_rateado(pasta_custos, df_mes, mes_referencia):
    """MÉDICOS: aba própria "MEDICOS" no arquivo CUSTOS RATEADOS.xlsx, uma linha
    por mês (CUSTO, VALOR, DATA, STATUS). Regra do mês defasado: o valor do
    mês ATUAL sempre fica em SALDO; o que é rateado este mês é o(s)
    valor(es) que ainda estiverem SALDO de mês(es) ANTERIOR(es) a
    mes_referencia (normalmente só o mês passado, mas soma tudo que
    estiver parado, caso tenha ficado algum mês esquecido pra trás).
    A automação só lê o STATUS - não escreve de volta no arquivo."""
    qtd_doadoras = len(df_mes)
    if qtd_doadoras == 0:
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])

    cfg = config.ARQUIVO_MEDICOS_RATEADO
    caminho = None
    for arquivo in os.listdir(pasta_custos):
        if identificar_arquivo(arquivo, {"CUSTOS_RATEADOS": {"arquivo_padrao": cfg["arquivo_padrao"]}}):
            caminho = os.path.join(pasta_custos, arquivo)
            break
    if caminho is None:
        avisar_arquivo_custos_rateados_nao_encontrado(pasta_custos, "MEDICOS não será lançado")
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])

    try:
        df = pd.read_excel(caminho, sheet_name=cfg["sheet_name"], header=cfg["header_row"])
    except ValueError:
        print(f"  [aviso] aba '{cfg['sheet_name']}' não encontrada em '{os.path.basename(caminho)}' - "
              f"MEDICOS não será lançado este mês")
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])

    df.columns = [str(c).strip() for c in df.columns]
    df[cfg["coluna_data"]] = pd.to_datetime(df[cfg["coluna_data"]], errors="coerce")
    df["mes_da_linha"] = df[cfg["coluna_data"]].dt.strftime("%Y-%m")
    df["status_norm"] = df[cfg["coluna_status"]].astype(str).str.strip().str.upper()

    pendentes = df[(df["status_norm"] == "SALDO") & (df["mes_da_linha"] < mes_referencia)]
    if pendentes.empty:
        print(f"  [aviso] nenhuma linha de MEDICOS com STATUS=SALDO de mês anterior a {mes_referencia} "
              f"- MEDICOS não será lançado este mês")
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])

    meses_consumidos = ", ".join(sorted(pendentes["mes_da_linha"]))
    if len(pendentes) > 1:
        print(f"  [atenção] mais de 1 mês de MEDICOS ainda em SALDO ({meses_consumidos}) - somando todos pra ratear este mês")

    valor_total = pendentes[cfg["coluna_valor"]].sum()
    valor_por_doadora = round(float(valor_total) / qtd_doadoras, 2)
    print(f"  [ok] MEDICOS: R$ {valor_total:.2f} (referente a {meses_consumidos}, estava em SALDO) "
          f"rateado entre {qtd_doadoras} doadoras = R$ {valor_por_doadora:.2f} cada")
    print(f"       lembrete: marque STATUS='CONSUMIDO' nessa(s) linha(s) na planilha, e adicione a "
          f"linha de {mes_referencia} com STATUS='SALDO'")

    linhas = [{
        "nome_bruto": nome, "valor": valor_por_doadora,
        "data": pendentes[cfg["coluna_data"]].max(), "categoria": cfg["categoria_interna"],
    } for nome in df_mes["nome"]]
    return pd.DataFrame(linhas)


# ---------------------------------------------------------------------------
# 3b. SALDO INICIAL (--saldo-inicial) - só na 1ª execução: custos que já
#     estavam em aberto ANTES de começar a usar a automação
# ---------------------------------------------------------------------------

_EPOCH_EXCEL = datetime(1899, 12, 30)


def converter_data_excel(valor):
    """Converte uma célula de data pro formato certo, mesmo quando ela veio
    como número serial do Excel (acontece quando a célula não está
    formatada como data na planilha de origem - o pandas/openpyxl então lê
    um float puro tipo 46199.3 em vez de um datetime). Sem isso, esse
    número bruto vaza pro relatório (já aconteceu na aba SALDO pendente)."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, (int, float)):
        return pd.Timestamp(_EPOCH_EXCEL + timedelta(days=float(valor)))
    return pd.to_datetime(valor, errors="coerce")


def _normalizar_nome_aba(nome):
    """Para comparar nomes de aba ignorando maiúsculas/minúsculas e a
    diferença entre espaço e underscore (ex: 'Taxi_Parceiro' == 'TAXI PARCEIRO'). Esse arquivo já apareceu com grafias diferentes entre meses."""
    return str(nome).strip().upper().replace("_", " ")


def carregar_saldo_inicial(caminho_arquivo):
    partes = []
    abas_reais = {_normalizar_nome_aba(a): a for a in pd.ExcelFile(caminho_arquivo).sheet_names}
    for aba, cfg in config.SALDO_INICIAL_ABAS.items():
        aba_real = abas_reais.get(_normalizar_nome_aba(aba))
        if aba_real is None:
            print(f"  [aviso] aba '{aba}' não encontrada em '{caminho_arquivo}' - pulando essa categoria")
            continue
        df = pd.read_excel(caminho_arquivo, sheet_name=aba_real)
        df.columns = [str(c).strip() for c in df.columns]

        # a coluna "UTILIZADO" NÃO é usada como filtro: toda linha com nome e
        # valor entra no matching. Quem bater com uma doadora vira lançamento
        # do mês; quem não bater vira SALDO com a data original (mesma regra
        # de sempre) - ver comentário em config.SALDO_INICIAL_ABAS.

        df = df.rename(columns={
            cfg["coluna_nome"]: "nome_bruto",
            cfg["coluna_valor"]: "valor",
            cfg["coluna_data"]: "data",
        })
        df["data"] = df["data"].apply(converter_data_excel)
        df = df.dropna(subset=["nome_bruto", "valor"])
        df["nome_bruto"] = df["nome_bruto"].astype(str).str.strip()
        # remove linhas de rodapé de totais (ex: "TOTAL"), que não são
        # doadoras - sem isso, o valor do rodapé (soma de tudo) entrava
        # como se fosse mais uma pessoa pendente, duplicando o saldo
        df = df[~df["nome_bruto"].str.upper().isin(["TOTAL", "TOTAIS"])]
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
        df = df.dropna(subset=["valor"])
        df["categoria"] = cfg["categoria"]
        print(f"  [ok] lendo aba '{aba_real}' do saldo inicial ({len(df)} linhas) como {cfg['categoria']}")
        partes.append(df[["nome_bruto", "valor", "data", "categoria"]])

    if not partes:
        return pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])
    return pd.concat(partes, ignore_index=True)


def pre_semear_controle_saldo_com_saldo_inicial(df_saldo_inicial, pasta_controle):
    """Antes de rodar o matching, grava no controle de SALDO a data ORIGINAL
    (mais antiga) de cada item do saldo inicial que ainda não tinha entrada
    lá - assim, se o item não bater com nenhuma doadora, a contagem dos 3
    meses até virar AJUSTE já começa a contar da data certa, não de hoje."""
    if df_saldo_inicial.empty:
        return

    df = df_saldo_inicial.copy()
    df["mes_origem"] = pd.to_datetime(df["data"], errors="coerce").dt.strftime("%Y-%m")
    df = df.dropna(subset=["mes_origem"])
    if df.empty:
        return
    primeiro_mes_por_par = df.groupby(["nome_bruto", "categoria"])["mes_origem"].min().reset_index()

    controle = carregar_controle_saldo(pasta_controle)
    ja_existentes = set()
    if not controle.empty:
        ja_existentes = set(zip(controle["nome_bruto"].str.lower(), controle["categoria"]))

    novas_linhas = []
    for _, linha in primeiro_mes_por_par.iterrows():
        chave = (linha["nome_bruto"].lower(), linha["categoria"])
        if chave in ja_existentes:
            continue  # já tem uma entrada de execução anterior - não sobrescreve
        novas_linhas.append({
            "nome_bruto": linha["nome_bruto"],
            "categoria": linha["categoria"],
            "primeiro_mes_pendente": linha["mes_origem"],
            "status": "SALDO",
        })

    if novas_linhas:
        controle = pd.concat([controle, pd.DataFrame(novas_linhas)], ignore_index=True)
        salvar_controle_saldo(controle, pasta_controle)
    print(f"  [info] {len(novas_linhas)} nomes/categorias do saldo inicial semeados no controle "
          f"com a data de origem correta")


# ---------------------------------------------------------------------------
# 4. MATCHING APROXIMADO + CONTROLE DE SALDO / AJUSTE
#    (só se aplica aos custos por doadora - os rateados já vêm com o nome
#    exato da doadora, então sempre batem)
# ---------------------------------------------------------------------------

def encontrar_doadora(nome_bruto, nomes_master):
    if not nomes_master:
        return None, 0
    # fuzz.ratio (comparação direta) em vez de WRatio: o WRatio usa
    # comparação por CONJUNTO de palavras, o que faz nomes que só têm um
    # sobrenome comum em comum (ex: "Martins", "Ferreira") baterem com
    # score alto mesmo sendo pessoas completamente diferentes.
    resultado = process.extractOne(
    nome_bruto, nomes_master, scorer=fuzz.ratio, processor=utils.default_process)
    if resultado is None:
        return None, 0
    nome_encontrado, score, _ = resultado
    if score >= config.SCORE_MINIMO_MATCH:
        return nome_encontrado, score
    return None, score


def carregar_controle_saldo(pasta_controle):
    caminho = os.path.join(pasta_controle, config.ARQUIVO_CONTROLE_SALDO)
    if os.path.exists(caminho):
        return pd.read_csv(caminho, dtype=str)
    return pd.DataFrame(columns=["nome_bruto", "categoria", "primeiro_mes_pendente", "status"])


def salvar_controle_saldo(df, pasta_controle):
    os.makedirs(pasta_controle, exist_ok=True)
    caminho = os.path.join(pasta_controle, config.ARQUIVO_CONTROLE_SALDO)
    df.to_csv(caminho, index=False)


def meses_entre(mes_inicial_str, mes_atual_str):
    inicial = datetime.strptime(mes_inicial_str, "%Y-%m")
    atual = datetime.strptime(mes_atual_str, "%Y-%m")
    diff = relativedelta(atual, inicial)
    return diff.years * 12 + diff.months


def classificar_custos(df_custos, nomes_master, nome_para_data, mes_referencia, pasta_controle):
    controle = carregar_controle_saldo(pasta_controle)
    linhas_classificadas = []

    for _, linha in df_custos.iterrows():
        # categoria "somente_saldo" (config.FONTES_DE_CUSTO): entra SEMPRE como SALDO. Não passa
        # pelo matching (não vincula a nenhuma doadora, mesmo que o nome/PIN do arquivo exista
        # na lista) e não entra no controle de saldo, então nunca envelhece pra AJUSTE.
        if config.FONTES_DE_CUSTO.get(linha["categoria"], {}).get("somente_saldo"):
            linhas_classificadas.append({
                **linha.to_dict(),
                "status": "SALDO",
                "matched": False,
                "doadora_identificada": "",
                "score_match": 0,
                "observacao": "somente saldo (não vincula doadora)",
            })
            continue

        nome_encontrado, score = encontrar_doadora(linha["nome_bruto"], nomes_master)

        if nome_encontrado:
            data_pertencimento = nome_para_data.get(nome_encontrado, "")
            linhas_classificadas.append({
                **linha.to_dict(),
                "status": data_pertencimento,  # data (coluna K) em vez de "DOADORA"
                "matched": True,
                "doadora_identificada": nome_encontrado,
                "score_match": score,
                "observacao": "",
            })
            controle = controle[~(
                (controle["nome_bruto"].str.lower() == linha["nome_bruto"].lower())
                & (controle["categoria"] == linha["categoria"])
            )]
            continue

        existente = controle[
            (controle["nome_bruto"].str.lower() == linha["nome_bruto"].lower())
            & (controle["categoria"] == linha["categoria"])
        ]

        if existente.empty:
            controle = pd.concat([controle, pd.DataFrame([{
                "nome_bruto": linha["nome_bruto"],
                "categoria": linha["categoria"],
                "primeiro_mes_pendente": mes_referencia,
                "status": "SALDO",
            }])], ignore_index=True)
            linhas_classificadas.append({
                **linha.to_dict(),
                "status": "SALDO",
                "matched": False,
                "doadora_identificada": "",
                "score_match": score,
                "observacao": "1º mês pendente",
            })
        else:
            primeiro_mes = existente.iloc[0]["primeiro_mes_pendente"]
            qtd_meses = meses_entre(primeiro_mes, mes_referencia)
            if qtd_meses >= config.MESES_LIMITE_SALDO:
                data_ajuste = datetime.strptime(mes_referencia, "%Y-%m").strftime("%d/%m/%Y")
                linhas_classificadas.append({
                    **linha.to_dict(),
                    "status": f"AJUSTES {data_ajuste}",
                    "matched": False,
                    "doadora_identificada": "",
                    "score_match": score,
                    "observacao": f"pendente desde {primeiro_mes} ({qtd_meses} meses)",
                })
                idx = existente.index[0]
                controle.loc[idx, "status"] = "AJUSTADO"
            else:
                linhas_classificadas.append({
                    **linha.to_dict(),
                    "status": "SALDO",
                    "matched": False,
                    "doadora_identificada": "",
                    "score_match": score,
                    "observacao": f"pendente desde {primeiro_mes} ({qtd_meses} meses)",
                })

    return pd.DataFrame(linhas_classificadas), controle


# ---------------------------------------------------------------------------
# 5. HISTÓRICO CONSOLIDADO POR DOADORA/MÊS (para a aba "Posição Consolidada")
#    e CÓPIA DA ABA "Controle Doadoras"
# ---------------------------------------------------------------------------

NOME_ARQUIVO_HISTORICO_CUSTOS = "historico_custos_doadora.csv"


def atualizar_historico_custos_doadora(df_classificado, mes_referencia, pasta_controle):
    """Guarda, mês a mês, quanto cada doadora custou em cada categoria.
    Isso alimenta a aba 'Posição Consolidada' do relatório. Só entram
    lançamentos que bateram com uma doadora (matched == True);
    SALDO/AJUSTES não têm doadora identificada."""
    novo = df_classificado[df_classificado["matched"]][
        ["doadora_identificada", "categoria", "valor"]
    ].copy()
    novo = novo.groupby(["doadora_identificada", "categoria"], as_index=False)["valor"].sum()
    novo["mes_referencia"] = mes_referencia
    novo = novo.rename(columns={"doadora_identificada": "doadora"})

    # tudo que caiu em AJUSTE neste mês (seja de custo de agosto, setembro...)
    # entra numa linha só, "AJUSTES DIVERSOS", somada por categoria
    ajustes = df_classificado[df_classificado["status"].astype(str).str.startswith("AJUSTES")]
    if not ajustes.empty:
        linha_ajustes = ajustes.groupby("categoria", as_index=False)["valor"].sum()
        linha_ajustes["doadora"] = config.NOME_LINHA_AJUSTES_DIVERSOS
        linha_ajustes["mes_referencia"] = mes_referencia
        novo = pd.concat([novo, linha_ajustes[["doadora", "categoria", "valor", "mes_referencia"]]], ignore_index=True)

    caminho = os.path.join(pasta_controle, NOME_ARQUIVO_HISTORICO_CUSTOS)
    os.makedirs(pasta_controle, exist_ok=True)
    if os.path.exists(caminho):
        historico = pd.read_csv(caminho)
        # remove qualquer lançamento antigo deste mesmo mês antes de gravar de
        # novo, para que rodar o script 2x no mesmo mês não duplique valores
        historico = historico[historico["mes_referencia"] != mes_referencia]
        historico = pd.concat([historico, novo], ignore_index=True)
    else:
        historico = novo
    historico = semear_ajustes_diversos_historicos(historico, mes_referencia)
    historico.to_csv(caminho, index=False)
    return historico


def semear_ajustes_diversos_historicos(historico, mes_referencia):
    """Meses que foram fechados ANTES da automação (config.AJUSTES_DIVERSOS_
    HISTORICO, hoje só agosto/2026) entram na linha AJUSTES DIVERSOS 1 vez: só
    se aquele mês ainda não tiver nenhuma linha de AJUSTES DIVERSOS no
    histórico. Nunca mexe no mês que está sendo processado (esse vem do que a
    reconciliação classificou agora)."""
    novas = []
    for mes, por_categoria in config.AJUSTES_DIVERSOS_HISTORICO.items():
        if mes == mes_referencia:
            continue
        ja_tem = ((historico["doadora"] == config.NOME_LINHA_AJUSTES_DIVERSOS)
                  & (historico["mes_referencia"] == mes)).any()
        if ja_tem:
            continue
        for categoria, valor in por_categoria.items():
            novas.append({"doadora": config.NOME_LINHA_AJUSTES_DIVERSOS, "categoria": categoria,
                          "valor": valor, "mes_referencia": mes})
    if novas:
        print(f"  [info] AJUSTES DIVERSOS de meses anteriores à automação adicionados ao histórico "
              f"({', '.join(sorted({n['mes_referencia'] for n in novas}))}) - só acontece 1 vez")
        historico = pd.concat([historico, pd.DataFrame(novas)], ignore_index=True)
    return historico


def gerar_posicao_consolidada(ws, historico):
    """Monta a aba com TODAS as doadoras que já passaram pela empresa,
    um bloco de colunas por mês (custo por categoria + total do mês),
    e uma coluna de total geral no final - no espírito da 'Posição 3 via',
    mas calculada a partir do histórico que o próprio script acumula."""
    if historico.empty:
        ws.append(["Ainda não há histórico suficiente - rode o script por alguns meses."])
        return

    meses = sorted(historico["mes_referencia"].unique())
    categorias = sorted(historico["categoria"].unique())
    # AJUSTES DIVERSOS não é doadora: fica sempre na última linha, em negrito
    doadoras = sorted(d for d in historico["doadora"].unique() if d != config.NOME_LINHA_AJUSTES_DIVERSOS)
    if (historico["doadora"] == config.NOME_LINHA_AJUSTES_DIVERSOS).any():
        doadoras.append(config.NOME_LINHA_AJUSTES_DIVERSOS)

    pivot = historico.pivot_table(
        index="doadora", columns=["mes_referencia", "categoria"], values="valor",
        aggfunc="sum", fill_value=0
    )

    # linha 1: nome do mês (mesclado sobre o bloco de categorias + total dele)
    # linha 2: nome da categoria / "TOTAL MÊS"
    linha1 = ["Doadora"]
    linha2 = [""]
    for mes in meses:
        linha1.extend([mes] + [""] * len(categorias))  # +1 para a col. "TOTAL MÊS"
        linha2.extend(categorias + ["TOTAL MÊS"])
    linha1.append("")
    linha2.append("TOTAL GERAL")
    ws.append(linha1)
    ws.append(linha2)

    col_atual = 2
    for mes in meses:
        n_col = len(categorias) + 1
        if n_col > 1:
            ws.merge_cells(start_row=1, start_column=col_atual, end_row=1, end_column=col_atual + n_col - 1)
        col_atual += n_col

    for doadora in doadoras:
        linha = [doadora]
        total_geral = 0
        for mes in meses:
            total_mes = 0
            for cat in categorias:
                valor = pivot.loc[doadora, (mes, cat)] if (mes, cat) in pivot.columns else 0
                valor = float(valor) if pd.notna(valor) else 0
                linha.append(valor)
                total_mes += valor
            linha.append(total_mes)
            total_geral += total_mes
        linha.append(total_geral)
        ws.append(linha)

    # estilo do cabeçalho (2 linhas) e largura das colunas
    for row in (1, 2):
        for col in range(1, len(linha1) + 1):
            celula = ws.cell(row=row, column=col)
            celula.font = Font(bold=True, color="FFFFFF", name="Arial")
            celula.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            celula.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 32
    for col in range(2, len(linha1) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16

    # formato numérico brasileiro nas células de valor (todas as colunas
    # menos a 1ª, que é o nome da doadora), para poder somar/usar fórmula
    for row in range(3, ws.max_row + 1):
        for col in range(2, len(linha1) + 1):
            celula = ws.cell(row=row, column=col)
            if isinstance(celula.value, (int, float)):
                celula.number_format = "#,##0.00"

    for row in range(3, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == config.NOME_LINHA_AJUSTES_DIVERSOS:
            for col in range(1, len(linha1) + 1):
                ws.cell(row=row, column=col).font = Font(bold=True, name="Arial")

    ws.freeze_panes = "B3"


NOME_ARQUIVO_CONTROLE_DOADORAS = "controle_doadoras_completo.csv"


def carregar_seed_controle_doadoras(caminho_arquivo, aba="Controle Doadoras"):
    """Lê o arquivo de Controle de Doadoras original PULANDO A 1ª LINHA
    (linha de título) - usado só na primeira execução do script, para
    semear o histórico com as doadoras que já existiam antes da automação."""
    # cabeçalho detectado (a lista nova não tem mais a linha de título no topo)
    linha_cabecalho = _detectar_linha_cabecalho(caminho_arquivo, [
        config.MASTER_COLUNA_CODIGO, config.MASTER_COLUNA_NOME,
        config.MASTER_COLUNA_LOTES_DOADOS, config.MASTER_COLUNA_LOTES_CONGELADOS])
    df = pd.read_excel(caminho_arquivo, sheet_name=aba, header=linha_cabecalho)
    df = df.rename(columns={
        config.MASTER_COLUNA_CODIGO: "codigo",
        config.MASTER_COLUNA_NOME: "nome",
        config.MASTER_COLUNA_LOTES_DOADOS: "lotes_doados",
        config.MASTER_COLUNA_LOTES_CONGELADOS: "lotes_congelados",
        config.MASTER_COLUNA_MINILOTES: "minilotes",
        config.MASTER_COLUNA_MES: "data_pertencimento",
    })
    if "minilotes" not in df.columns:
        df["minilotes"] = 0
    df["minilotes"] = pd.to_numeric(df["minilotes"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["codigo", "nome"])
    df = df[~df["nome"].astype(str).str.contains(">", regex=False)]
    df["nome"] = df["nome"].astype(str).str.strip()
    df["codigo"] = df["codigo"].astype(str).str.strip()
    df["data_pertencimento"] = pd.to_datetime(df["data_pertencimento"], errors="coerce").dt.strftime("%d/%m/%Y")
    return df[COLUNAS_CONTROLE_DOADORAS].astype(str)


COLUNAS_CONTROLE_DOADORAS = ["codigo", "nome", "lotes_doados", "lotes_congelados", "minilotes", "data_pertencimento"]


def inicializar_ou_atualizar_controle_doadoras(caminho_seed, df_mes, pasta_controle):
    """
    A planilha de Controle de Doadoras (--controle-doadoras) só é usada UMA
    VEZ, na primeira execução do script, para semear o histórico. A partir
    daí, o próprio script mantém e atualiza esse histórico sozinho, usando
    as doadoras que chegam todo mês pela 'lista_doadoras' - não é mais
    necessário (nem é usado) passar --controle-doadoras de novo depois.
    """
    caminho = os.path.join(pasta_controle, NOME_ARQUIVO_CONTROLE_DOADORAS)
    if os.path.exists(caminho):
        completo = pd.read_csv(caminho, dtype=str)
        if "minilotes" not in completo.columns:
            # histórico gravado antes dos mini lotes existirem: esses meses não tinham mini lote
            completo["minilotes"] = "0"
        if caminho_seed:
            print("  [info] controle de doadoras já foi inicializado antes - "
                  "ignorando --controle-doadoras (só é usado na 1ª execução)")
    elif caminho_seed:
        print(f"  [info] 1ª execução - semeando controle de doadoras a partir de '{caminho_seed}'")
        completo = carregar_seed_controle_doadoras(caminho_seed)
    else:
        print("  [info] 1ª execução sem --controle-doadoras - controle de doadoras "
              "vai começar vazio e ser preenchido com as próximas 'lista_doadoras'")
        completo = pd.DataFrame(columns=COLUNAS_CONTROLE_DOADORAS)

    novo = df_mes[COLUNAS_CONTROLE_DOADORAS].astype(str)
    # A lista é a fonte da verdade do mês: o(s) mês(es) que ela traz SUBSTITUI(M) o que já havia
    # no histórico. Isso evita duplicar se o script rodar 2x e também limpa a linha antiga quando
    # um PIN ou um nome é corrigido na lista. NÃO se deduplica por PIN: o PIN pode vir repetido
    # na lista por erro de digitação, e quando a deduplicação era por PIN + mês a segunda doadora
    # com o mesmo PIN fazia a primeira SUMIR do histórico (e dos lotes do Controle Estoque).
    meses_da_lista = set(novo["data_pertencimento"]) - {"nan", "NaT", ""}
    completo = completo[COLUNAS_CONTROLE_DOADORAS]
    completo = completo[~completo["data_pertencimento"].isin(meses_da_lista)]
    completo = pd.concat([completo, novo], ignore_index=True)

    os.makedirs(pasta_controle, exist_ok=True)
    completo.to_csv(caminho, index=False)
    return completo


def escrever_aba_controle_doadoras(ws, completo):
    tabela = completo.rename(columns={
        "codigo": "Código (PIN)",
        "nome": "Nome da doadora",
        "lotes_doados": "Lotes (Fresco)",
        "lotes_congelados": "Lotes (Congelados/Banco)",
        "minilotes": "Mini lotes",
        "data_pertencimento": "Data (Mês)",
    })
    escrever_dataframe(ws, tabela)
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# 6. ABA "Lançamento Contábil"
# ---------------------------------------------------------------------------

def montar_texto_lancamento(nome, categoria, mes_referencia):
    primeiro_nome = str(nome).strip().split()[0].upper() if str(nome).strip() else ""
    rotulo_categoria = config.CATEGORIA_ROTULO_CONTABIL.get(categoria, categoria.replace("_", " "))
    ano, mes = mes_referencia.split("-")
    mes_ano = f"{config.MESES_PT[int(mes)].upper()}/{ano}"
    return f"Trf. Est. 3 {primeiro_nome}_{rotulo_categoria} - {mes_ano}"


def gerar_lancamentos_contabeis(df_classificado, mes_referencia):
    """Um lançamento por linha, para tudo que é do mês (doadora identificada)
    ou virou AJUSTE do mês - SALDO pendente não entra, pois ainda não foi
    atribuído a ninguém. O valor fica em coluna própria (número real do
    Excel), não dentro do texto do lançamento."""
    elegiveis = df_classificado[
        df_classificado["matched"] | df_classificado["status"].str.startswith("AJUSTES")
    ].copy()
    if elegiveis.empty:
        return pd.DataFrame(columns=["lancamento", "doadora", "categoria", "valor", "data"])

    elegiveis["nome_para_lancamento"] = elegiveis["doadora_identificada"].where(
        elegiveis["doadora_identificada"].astype(bool), elegiveis["nome_bruto"]
    )
    elegiveis["lancamento"] = elegiveis.apply(
        lambda r: montar_texto_lancamento(r["nome_para_lancamento"], r["categoria"], mes_referencia),
        axis=1,
    )
    elegiveis = elegiveis.rename(columns={"nome_para_lancamento": "doadora"})

    # consolida: quando a mesma doadora tem mais de uma linha na mesma
    # categoria este mês (ex: 2 corridas de Táxi), soma tudo num
    # lançamento só, em vez de uma linha por transação - a data fica sendo
    # a mais recente do grupo
    consolidado = elegiveis.groupby(["lancamento", "doadora", "categoria"], as_index=False).agg(
        valor=("valor", "sum"), data=("data", "max")
    )
    return consolidado[["lancamento", "doadora", "categoria", "valor", "data"]]


# ---------------------------------------------------------------------------
# 7. RELATÓRIO FINAL EM EXCEL
# ---------------------------------------------------------------------------

def estilizar_cabecalho(ws, n_colunas):
    for col in range(1, n_colunas + 1):
        celula = ws.cell(row=1, column=col)
        celula.font = Font(bold=True, color="FFFFFF", name="Arial")
        celula.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        celula.alignment = Alignment(horizontal="center")
    for col in range(1, n_colunas + 1):
        ws.column_dimensions[get_column_letter(col)].width = 24


def escrever_dataframe(ws, df):
    ws.append(list(df.columns))
    # colunas monetárias e de lotes viram número real do Excel (não texto),
    # para poder somar/usar fórmula direto na planilha
    colunas_moeda = {i for i, c in enumerate(df.columns) if str(c).strip().lower() in ("valor", "total_mes", "total_geral")}
    colunas_lotes = {i for i, c in enumerate(df.columns) if "lotes" in str(c).strip().lower()}
    formatos_numericos = {i: "#,##0.00" for i in colunas_moeda}
    formatos_numericos.update({i: "0" for i in colunas_lotes})

    for _, linha in df.iterrows():
        valores_linha = []
        for i, v in enumerate(linha):
            if pd.isna(v) or v == "":
                valores_linha.append("")
            elif i in formatos_numericos:
                # colunas numéricas: tenta converter mesmo se veio como texto
                # do CSV de controle (ex: "2.0"), para virar número de verdade
                try:
                    valores_linha.append(float(v))
                except (ValueError, TypeError):
                    valores_linha.append(str(v))
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                valores_linha.append(float(v))
            else:
                valores_linha.append(str(v))
        ws.append(valores_linha)

    estilizar_cabecalho(ws, len(df.columns))

    for col_idx, formato in formatos_numericos.items():
        col_letra = get_column_letter(col_idx + 1)
        for linha_num in range(2, ws.max_row + 1):
            celula = ws[f"{col_letra}{linha_num}"]
            if isinstance(celula.value, (int, float)):
                # '#,##0.00'/'0' são códigos internos do Excel; aparecem com
                # ponto de milhar e vírgula decimal quando o Excel está em
                # português (Brasil)
                celula.number_format = formato


PADRAO_ARQUIVO_CONFERENCIA = re.compile(r"^conferencia_doadoras_(\d{4}-\d{2})\.xlsx$", re.IGNORECASE)

# prefixo das linhas "AJUSTE" que o SALDO pendente carrega do mês anterior (ver
# carregar_linhas_ajuste_anterior). A conferência reconhece esse prefixo.
PREFIXO_LINHA_AJUSTE_ANTERIOR = "(sem doadora - AJUSTE mantido do mês anterior)"


def encontrar_conferencia_anterior(pasta_saida, mes_referencia, caminho_informado=None):
    """Acha o arquivo da conferência do mês anterior - o conferencia_doadoras_
    AAAA-MM.xlsx MAIS RECENTE da pasta de saída com mês menor que o mês sendo
    processado (assim, se um mês foi pulado, pega o último que existe). Se a
    usuária informou o caminho (--conferencia-anterior), usa esse. Devolve None
    se não houver nenhum (1ª vez usando a automação)."""
    if caminho_informado:
        if not os.path.exists(caminho_informado):
            raise FileNotFoundError(f"--conferencia-anterior: não encontrei '{caminho_informado}'")
        return caminho_informado
    if not os.path.isdir(pasta_saida):
        return None
    candidatos = []
    for nome in os.listdir(pasta_saida):
        m = PADRAO_ARQUIVO_CONFERENCIA.match(nome)
        if m and m.group(1) < mes_referencia:
            candidatos.append((m.group(1), os.path.join(pasta_saida, nome)))
    return max(candidatos)[1] if candidatos else None


def carregar_linhas_ajuste_anterior(caminho_conferencia):
    """Lê, do SALDO pendente da conferência do mês anterior (já com as edições
    manuais da usuária), as linhas de AJUSTE: as que têm 'AJUSTE' na coluna
    'data' e nenhum nome de doadora. Devolve um DataFrame (categoria interna,
    valor) - uma linha por linha de AJUSTE, sem somar."""
    rotulo_para_interno = {v: k for k, v in config.CATEGORIA_ROTULO_CONTABIL.items()}
    ws = load_workbook(caminho_conferencia, data_only=True)
    if "SALDO pendente" not in ws.sheetnames:
        print(f"  [aviso] '{caminho_conferencia}' não tem a aba 'SALDO pendente' - nenhuma linha de AJUSTE foi mantida")
        return pd.DataFrame(columns=["categoria", "valor"])
    ws = ws["SALDO pendente"]
    linhas, sem_valor = [], 0
    for linha in ws.iter_rows(min_row=2, values_only=True):
        nome, categoria, valor, data = linha[0], linha[1], linha[2], linha[3]
        if categoria and str(data).strip().upper() == "AJUSTE" and nome in (None, ""):
            if isinstance(valor, (int, float)):
                linhas.append({"categoria": rotulo_para_interno.get(categoria, categoria), "valor": float(valor)})
            else:
                sem_valor += 1
    if sem_valor:
        print(f"  [aviso] {sem_valor} linha(s) de AJUSTE de '{os.path.basename(caminho_conferencia)}' estão sem valor "
              f"calculado (fórmula nunca aberta no Excel?) - abra e salve o arquivo no Excel e rode de novo")
    return pd.DataFrame(linhas, columns=["categoria", "valor"])


def montar_linhas_ajuste_anterior(df_ajuste, caminho_conferencia):
    """Linhas no formato do SALDO pendente da reconciliação (o valor é mantido
    EXATAMENTE como a usuária deixou no mês anterior)."""
    if df_ajuste is None or df_ajuste.empty:
        return pd.DataFrame(columns=["nome_bruto", "categoria", "valor", "data", "score_match", "observacao"])
    origem = os.path.basename(caminho_conferencia)
    return pd.DataFrame([{
        "nome_bruto": PREFIXO_LINHA_AJUSTE_ANTERIOR, "categoria": l["categoria"], "valor": l["valor"],
        "data": "AJUSTE", "score_match": None,
        "observacao": f"AJUSTE mantido de '{origem}' - retirar manualmente do saldo",
    } for _, l in df_ajuste.iterrows()])


def gerar_linhas_saldo_legado(pasta_controle):
    """Saldo agregado de meses anteriores à automação (sem doadora
    individual - só o total por categoria, vindo da planilha antiga). Só
    entra 1 vez (marcador em config.ARQUIVO_SALDO_LEGADO_PROCESSADO) -
    depois disso já está incorporado na aba 'SALDO pendente' pra sempre,
    não precisa semear de novo."""
    caminho_marcador = os.path.join(pasta_controle, config.ARQUIVO_SALDO_LEGADO_PROCESSADO)
    if os.path.exists(caminho_marcador):
        return pd.DataFrame(columns=["nome_bruto", "categoria", "valor", "data", "score_match", "observacao"])
    linhas = [{
        "nome_bruto": "(sem doadora - saldo agregado de meses anteriores)",
        "categoria": categoria,
        "valor": valor,
        "data": None,
        "score_match": None,
        "observacao": "saldo de meses anteriores à automação, sem doadora individual identificada na planilha antiga",
    } for categoria, valor in config.SALDO_LEGADO_MESES_ANTERIORES.items()]
    os.makedirs(pasta_controle, exist_ok=True)
    with open(caminho_marcador, "w", encoding="utf-8") as f:
        f.write("Saldo legado de meses anteriores semeado - ver config.SALDO_LEGADO_MESES_ANTERIORES\n")
    print(f"  [info] saldo legado de meses anteriores (sem doadora) adicionado à aba 'SALDO pendente' "
          f"({len(linhas)} categorias) - só acontece 1 vez")
    return pd.DataFrame(linhas)


def gerar_relatorio(df_classificado, mes_referencia, historico_consolidado, controle_doadoras_completo, pasta_saida="saida", pasta_controle="controle", linhas_ajuste_anterior=None):
    wb = Workbook()
    wb.remove(wb.active)
    for categoria in sorted(df_classificado["categoria"].unique()):
        aba = wb.create_sheet(categoria[:31])
        subset = df_classificado[df_classificado["categoria"] == categoria][
            ["nome_bruto", "doadora_identificada", "valor", "data", "status", "score_match", "observacao"]
        ]
        escrever_dataframe(aba, subset)

    saldo = df_classificado[df_classificado["status"] == "SALDO"][
        ["nome_bruto", "categoria", "valor", "data", "score_match", "observacao"]
    ]
    saldo_legado = gerar_linhas_saldo_legado(pasta_controle)
    saldo = pd.concat([saldo, saldo_legado], ignore_index=True)
    # linhas de AJUSTE que a usuária criou no SALDO pendente do mês anterior:
    # repetidas aqui, com o mesmo valor, pra ela retirar manualmente do saldo
    if linhas_ajuste_anterior is not None and not linhas_ajuste_anterior.empty:
        saldo = pd.concat([saldo, linhas_ajuste_anterior], ignore_index=True)
    escrever_dataframe(wb.create_sheet("SALDO pendente"), saldo)

    # ajustes do mês - consolidado por doadora+categoria (soma o valor
    # quando a mesma doadora tem mais de uma linha virando ajuste na mesma
    # categoria), igual ao Lançamento Contábil. SALDO pendente (acima)
    # continua linha por linha de propósito - só os ajustes são somados.
    ajuste = df_classificado[df_classificado["status"].str.startswith("AJUSTES")][
        ["nome_bruto", "categoria", "valor", "data", "status", "score_match", "observacao"]
    ]
    if not ajuste.empty:
        ajuste = ajuste.groupby(["nome_bruto", "categoria", "status"], as_index=False).agg(
            valor=("valor", "sum"), data=("data", "max"),
            score_match=("score_match", "max"), observacao=("observacao", "first"),
        )[["nome_bruto", "categoria", "valor", "data", "status", "score_match", "observacao"]]
    escrever_dataframe(wb.create_sheet("AJUSTE do mês"), ajuste)

    # lançamento contábil: 1 linha por lançamento do mês (doadora identificada
    # ou ajuste do mês) - SALDO pendente não entra aqui
    lancamentos = gerar_lancamentos_contabeis(df_classificado, mes_referencia)
    escrever_dataframe(wb.create_sheet("Lançamento Contábil"), lancamentos)

    # todas as doadoras que já passaram pela empresa, consolidadas por mês
    gerar_posicao_consolidada(wb.create_sheet("Posição Consolidada"), historico_consolidado)

    # controle de doadoras acumulado (semeado 1x, atualizado todo mês)
    escrever_aba_controle_doadoras(wb.create_sheet("Controle Doadoras"), controle_doadoras_completo)

    caminho_saida = f"{pasta_saida}/relatorio_reconciliacao_{mes_referencia}.xlsx"
    os.makedirs(pasta_saida, exist_ok=True)
    wb.save(caminho_saida)
    return caminho_saida


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Reconciliação mensal de custos de doadoras")
    parser.add_argument("--mes", required=True, help="Mês de referência, formato AAAA-MM, ex: 2026-07")
    parser.add_argument("--lista-doadoras", required=True, help="Planilha do estagiário com as doadoras do mês")
    parser.add_argument("--custos", required=True, help="Pasta com os arquivos de custo de cada fornecedor")
    parser.add_argument("--medicamentos-liberado", choices=["sim", "nao"], default="nao",
                         help="Se o fechamento contábil de medicamentos (4BIO) já liberou este mês")
    parser.add_argument("--controle-doadoras", default=None,
                         help="Caminho do arquivo com a aba 'Controle Doadoras' - só é usado na "
                              "1ª execução do script, para semear o histórico. Nas execuções "
                              "seguintes esse argumento é ignorado (histórico já existe).")
    parser.add_argument("--modo-teste", action="store_true",
                         help="Usa uma pasta de controle separada ('controle_teste/') em vez de "
                              "'controle/', para testar o script sem mexer no histórico real "
                              "(SALDO, cadastro mestre, controle de doadoras). Rode com essa flag "
                              "sempre que quiser experimentar sem contaminar os dados de produção.")
    parser.add_argument("--saldo-inicial", default=None,
                         help="Caminho do Excel com os custos que já estavam em aberto (SALDO) antes "
                              "de começar a usar a automação (uma aba por fornecedor). Só é usado na "
                              "1ª execução - depois disso é ignorado (já foi incorporado ao controle).")
    parser.add_argument("--saldo-adicional", default=None,
                         help="Mesmo formato do --saldo-inicial (uma aba por fornecedor), mas para "
                              "corrigir itens esquecidos depois - ex: uma listagem que devia ter "
                              "entrado num mês passado e não entrou. Cada arquivo diferente só é "
                              "processado 1 vez (controlado por um marcador com o nome do arquivo), "
                              "mas você pode usar --saldo-adicional várias vezes ao longo do tempo "
                              "com arquivos diferentes, ao contrário do --saldo-inicial (que é só 1x "
                              "para sempre). Os itens entram com a DATA ORIGINAL de cada linha, então "
                              "a contagem dos 3 meses até virar AJUSTE conta certo desde o início.")
    parser.add_argument("--conferencia-anterior", default=None,
                         help="Caminho da conferência do mês anterior (conferencia_doadoras_AAAA-MM.xlsx, já "
                              "com as suas edições manuais). Dela vêm as linhas de AJUSTE do SALDO pendente, "
                              "que são mantidas no mês novo pra você retirar manualmente do saldo. Padrão: o "
                              "conferencia_doadoras_*.xlsx mais recente (mês anterior ao processado) da pasta saida/.")
    args = parser.parse_args()

    pasta_controle = "controle_teste" if args.modo_teste else "controle"
    if args.modo_teste:
        print(f"[MODO TESTE] usando pasta '{pasta_controle}/' - o histórico real em 'controle/' "
              f"não será alterado.")

    print(f"== Reconciliação {args.mes} ==")

    print("1. Carregando lista de doadoras do mês...")
    df_mes = carregar_lista_doadoras_do_mes(args.lista_doadoras, args.mes)
    print(f"   {len(df_mes)} doadoras nesta planilha")

    print("2. Atualizando cadastro mestre histórico...")
    historico = atualizar_cadastro_mestre(os.path.join(pasta_controle, "cadastro_mestre_historico.csv"), df_mes)
    nomes_master = historico["nome"].dropna().unique().tolist()
    # data (coluna K) à qual cada doadora pertence, para usar como "status"
    # no lugar de "DOADORA" - se ela aparece em mais de um mês, usa a mais recente
    historico_ordenado = historico.sort_values("mes_referencia").drop_duplicates(subset="nome", keep="last")
    nome_para_data = dict(zip(historico_ordenado["nome"], historico_ordenado["data_pertencimento"]))
    print(f"   {len(nomes_master)} nomes distintos no cadastro mestre acumulado")

    print("3. Carregando arquivos de custo por doadora (Uber, Táxi, Cerba, Ajuda de Custo, Medicamentos)...")
    rodar_medicamentos = args.medicamentos_liberado == "sim"
    df_custos_doadora = carregar_todos_custos_por_doadora(args.custos, args.mes, rodar_medicamentos)
    print(f"   {len(df_custos_doadora)} lançamentos carregados")

    print("4. Carregando custos rateados (Médicos / Insumos / MGR)...")
    df_custos_rateados = carregar_custos_rateados(args.custos, df_mes, args.mes)
    print(f"   {len(df_custos_rateados)} linhas geradas pelo rateio (INSUMOS/MGR)")
    df_medicos_rateado = carregar_medicos_rateado(args.custos, df_mes, args.mes)
    print(f"   {len(df_medicos_rateado)} linhas geradas pelo rateio de MEDICOS")

    df_saldo_inicial = pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])
    if args.saldo_inicial:
        caminho_marcador = os.path.join(pasta_controle, "saldo_inicial_processado.txt")
        if os.path.exists(caminho_marcador):
            print("  [info] saldo inicial já foi processado antes - ignorando --saldo-inicial "
                  "(só é usado na 1ª execução)")
        else:
            print(f"4b. Carregando saldo inicial de '{args.saldo_inicial}' (só acontece 1x)...")
            df_saldo_inicial = carregar_saldo_inicial(args.saldo_inicial)
            pre_semear_controle_saldo_com_saldo_inicial(df_saldo_inicial, pasta_controle)
            os.makedirs(pasta_controle, exist_ok=True)
            with open(caminho_marcador, "w", encoding="utf-8") as f:
                f.write(f"Saldo inicial processado a partir de: {args.saldo_inicial}\n"
                        f"Na execução de {args.mes}.\n")
            print(f"   {len(df_saldo_inicial)} linhas do saldo inicial vão entrar no matching deste mês")

    df_saldo_adicional = pd.DataFrame(columns=["nome_bruto", "valor", "data", "categoria"])
    if args.saldo_adicional:
        nome_sanitizado = re.sub(r"[^A-Za-z0-9]+", "_", os.path.basename(args.saldo_adicional))
        caminho_marcador_adicional = os.path.join(pasta_controle, f"saldo_adicional_processado__{nome_sanitizado}.txt")
        if os.path.exists(caminho_marcador_adicional):
            print(f"  [info] '{args.saldo_adicional}' já foi processado antes via --saldo-adicional - ignorando")
        else:
            print(f"4c. Carregando saldo adicional de '{args.saldo_adicional}' (só acontece 1x para este arquivo)...")
            df_saldo_adicional = carregar_saldo_inicial(args.saldo_adicional)
            pre_semear_controle_saldo_com_saldo_inicial(df_saldo_adicional, pasta_controle)
            os.makedirs(pasta_controle, exist_ok=True)
            with open(caminho_marcador_adicional, "w", encoding="utf-8") as f:
                f.write(f"Saldo adicional processado a partir de: {args.saldo_adicional}\n"
                        f"Na execução de {args.mes}.\n")
            print(f"   {len(df_saldo_adicional)} linhas do saldo adicional vão entrar no matching deste mês")

    df_custos = pd.concat([df_custos_doadora, df_custos_rateados, df_medicos_rateado, df_saldo_inicial, df_saldo_adicional], ignore_index=True)

    print("5. Fazendo o matching e classificando SALDO/AJUSTES...")
    df_classificado, controle_atualizado = classificar_custos(
        df_custos, nomes_master, nome_para_data, args.mes, pasta_controle
    )
    salvar_controle_saldo(controle_atualizado, pasta_controle)

    resumo_status = pd.Series({
        "doadora identificada": int(df_classificado["matched"].sum()),
        "SALDO": int((df_classificado["status"] == "SALDO").sum()),
        "AJUSTES": int(df_classificado["status"].str.startswith("AJUSTES").sum()),
    })
    print("   Resultado do matching:")
    for status, qtd in resumo_status.items():
        print(f"     {status}: {qtd}")

    print("6. Atualizando controle de doadoras (histórico completo)...")
    controle_doadoras_completo = inicializar_ou_atualizar_controle_doadoras(
        args.controle_doadoras, df_mes, pasta_controle
    )
    print(f"   {len(controle_doadoras_completo)} linhas no histórico de doadoras")

    print("7. Gerando relatório Excel...")
    historico_consolidado = atualizar_historico_custos_doadora(df_classificado, args.mes, pasta_controle)
    pasta_saida = "saida_teste" if args.modo_teste else "saida"
    caminho_conf_anterior = encontrar_conferencia_anterior(pasta_saida, args.mes, args.conferencia_anterior)
    linhas_ajuste_anterior = None
    if caminho_conf_anterior:
        print(f"   conferência do mês anterior: '{caminho_conf_anterior}'")
        df_aj = carregar_linhas_ajuste_anterior(caminho_conf_anterior)
        linhas_ajuste_anterior = montar_linhas_ajuste_anterior(df_aj, caminho_conf_anterior)
        if len(df_aj):
            print(f"   {len(df_aj)} linha(s) de AJUSTE mantida(s) no SALDO pendente (R$ {df_aj['valor'].sum():,.2f}) - "
                  f"retire manualmente do saldo")
    else:
        print("   (nenhuma conferência de mês anterior encontrada em '" + pasta_saida + "/' - sem linhas de AJUSTE pra manter)")
    caminho = gerar_relatorio(df_classificado, args.mes, historico_consolidado, controle_doadoras_completo, pasta_saida, pasta_controle, linhas_ajuste_anterior)
    print(f"   Relatório salvo em: {caminho}")


if __name__ == "__main__":
    main()
