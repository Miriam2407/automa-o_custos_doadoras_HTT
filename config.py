# ============================================================================
# CONFIGURAÇÃO CENTRAL - ajuste aqui, sem precisar mexer no resto do código
# ============================================================================

# Quantos meses um item pode ficar como "SALDO" antes de virar "AJUSTE"
MESES_LIMITE_SALDO = 3

# Score mínimo (0-100) para considerar que um nome bate com uma doadora da
# lista mestre, usando fuzz.ratio (comparação direta caractere-a-caractere,
# NÃO por conjunto de palavras). Testado com casos reais (exemplos abaixo com nomes fictícios, scores calculados de verdade): matches corretos
# ficam entre 83-100 (inclui casos como "Carolina Maria de Jesus
# Carolina" -> "Carolina Maria de Jesus", nome repetido no campo -
# score 83,64); matches errados ficam abaixo disso - ex: "Clarice Helena
# Lispector" bateu incorretamente com "Clarisse Mila Lispector" com score
# 80,85 (pessoas diferentes, mas nomes com
# estrutura parecida e o mesmo sobrenome). 83 dá margem para não pegar
# esse caso, mantendo os matches corretos observados. Ainda assim, a
# margem entre certo e errado pode ser pequena - use a coluna "score_match"
# no relatório para revisar manualmente qualquer match abaixo de ~90.
SCORE_MINIMO_MATCH = 83

# --------------------------------------------------------------------------
# FONTES QUE VÊM COM NOME DE DOADORA (passam pelo matching e pela lógica
# de SALDO -> AJUSTE após 3 meses): Transporte (Uber/Táxi), Exames
# (Cerba), Ajuda de Custo, e Medicamentos (quando liberado).
#
# Campos:
#   arquivo_padrao -> palavras-chave para reconhecer o arquivo pelo nome
#   coluna_nome    -> coluna com o nome da doadora
#   coluna_valor   -> coluna com o valor cobrado/pago
#   coluna_data    -> coluna com a data (referência no relatório)
#   header_row     -> linha (0-indexada) onde está o cabeçalho de verdade
#                      (várias fontes têm linhas de título antes dele)
#   sheet_name     -> (opcional) nome fixo da aba a ler
#   aba_por_mes    -> (opcional) True se o arquivo tem 1 aba por mês
#                      (ex: "Agosto2026") - a aba é escolhida automaticamente
#                      a partir do --mes informado na linha de comando
#   filtro         -> (opcional) só considera linhas onde essa coluna tem
#                      esse valor exato
# --------------------------------------------------------------------------
FONTES_DE_CUSTO = {
    "TRANSPORTE_UBER": {
        "arquivo_padrao": ["uber"],
        "coluna_nome": "Nome do convidado",   # mudou de "Nome.1" pra uma coluna
                                                # própria a partir de setembro/2026 -
                                                # "Nome"/"Sobrenome" agora são da
                                                # funcionária que pediu a corrida
        # aceita as duas versões do export (a 1ª que existir no arquivo vale):
        # "Valor da transação: BRL" (export direto da Uber) ou "Valor" (planilha
        # de outubro/2026, já tratada)
        "coluna_valor": ["Valor da transação: BRL", "Valor"],
        "coluna_data": "Data",
        "header_row": 0,
        # a coluna é achada sem diferenciar maiúscula/minúscula ("Tipo" = "TIPO")
        "filtro": {"coluna": "Tipo", "valores": ["Doadora", "Ovodoação"]},
    },
    "TRANSPORTE_TAXI": {
        "arquivo_padrao": ["taxi"],
        "coluna_nome": "Passageiro",
        "coluna_valor": "Valor",
        "coluna_data": "Data",
        "header_row": 1,  # linha 0 é só o título "R.Transportes"
    },
    "EXAMES_CERBA": {
        "arquivo_padrao": ["cerba"],
        "coluna_nome": "NOME DO PACIENTE",
        "coluna_valor": "VALOR",   # valor por exame/linha; o script soma por
                                    # doadora sozinho. A coluna "VALOR TOTAL"
                                    # do arquivo tem fórmula do Excel sem
                                    # cache, por isso não é usada direto.
        "coluna_data": "DATA SOLICITAÇÃO",
        "header_row": 2,  # linhas 0-1 são título/convênio do relatório
        "sheet_name": "Exames Encaminhados",
    },
    "AJUDA_DE_CUSTO": {
        # repasse direto à doadora (Início / Coleta / Pós Coleta / Completo).
        # Categoria nova, à parte das 7 originais - confirmado com a usuária.
        "arquivo_padrao": ["ajuda_de_custo", "ajuda de custo"],
        "coluna_nome": "Nome",
        "coluna_valor": "Valor",
        "coluna_data": "Data envio",
        "header_row": 4,
        "aba_por_mes": True,
    },
    # Medicamentos fica registrado aqui mas só roda quando o fechamento
    # contábil liberar (veja --medicamentos-liberado no script principal).
    "MEDICAMENTOS_4BIO": {
        "arquivo_padrao": ["medicamento", "4bio"],
        "coluna_nome": "Paciente",
        "coluna_valor": "Total",
        "coluna_data": "Data",
        "header_row": 0,
        "sheet_name": "Planilha1",
    },
    # TESTES GENÉTICOS - exame de um laboratório de genética. Por enquanto entra SÓ em saldo: nunca é ligado
    # a uma doadora (mesmo que o PIN do arquivo exista na lista) e nunca envelhece pra
    # AJUSTE - ver "somente_saldo" em classificar_custos (reconciliacao.py). Quando as
    # movimentações forem definidas, é só tirar o "somente_saldo" (ou virar False).
    "TESTES_GENETICOS": {
        "arquivo_padrao": ["genetica"],
        "sheet_name": "Planilha1",   # a 1ª aba do arquivo (Planilha2) é uma tabela dinâmica
                                     # (soma por nota fiscal) - não é ela que se lê
        "header_row": 0,
        "coluna_nome": "NomePaciente",   # não é nome: vem um PIN ("DOADORA PIN 100801")
        "coluna_valor": "Valor bruto",
        "coluna_data": "mês",
        "somente_saldo": True,
    },
}

# --------------------------------------------------------------------------
# CUSTOS RATEADOS POR DOADORA (tipo "agregado" na origem, mas lançado por
# doadora no relatório): o CUSTOS RATEADOS.xlsx traz Material de Embriologia
# (Insumos) e MGR como total fechado do mês, sem nome, na aba "CUSTO".
# Médicos tem sua própria aba agora (veja ARQUIVO_MEDICOS_RATEADO abaixo),
# por causa da regra do mês defasado.
#
# Regra de negócio (confirmada com a usuária):
#   MATERIAL DE EMBRIOLOGIA (INSUMOS)-> NÃO faz conta nenhuma: cada doadora
#                                        recebe o valor exatamente como está
#                                        no arquivo - o valor do arquivo já
#                                        é o valor POR doadora (a soma de
#                                        todas as doadoras é que dá o total)
#   MGR                              -> idem: sem conta, valor por doadora
#                                        igual ao do arquivo
#
# O resultado dessa conta é lançado igualmente para CADA doadora do mês
# (não precisa/não tem como fazer matching por nome aqui, já que o arquivo
# não traz nome nenhum).
#
# Isso significa que essas categorias NUNCA caem em SALDO/AJUSTE: elas
# são sempre atribuídas a alguém (a lista de doadoras do mês).
#
# mapa_categoria: traduz o texto que vem na coluna de categoria do arquivo
# para (nome interno da categoria, operação: "dividir" ou "nenhuma").
# --------------------------------------------------------------------------
ARQUIVO_CUSTOS_RATEADOS = {
    "CUSTOS_RATEADOS": {
        # Aceita o nome com espaço ou com underscore.
        "arquivo_padrao": ["custos_rateados", "custos rateados"],
        "coluna_categoria": "CUSTO",
        "coluna_valor": "VALOR",
        "coluna_data": "DATA",
        "header_row": 0,
        "sheet_name": "CUSTO",
        "mapa_categoria": {
            "MATERIAL DE EMBRIOLOGIA": ("INSUMOS", "nenhuma"),
            "INSUMOS": ("INSUMOS", "nenhuma"),
            "MGR": ("MGR_ANESTESIA", "nenhuma"),
        },
    },
}

# Nome esperado do arquivo de custos rateados (usado nos avisos) e nomes ANTIGOS
# aceitos só pra avisar "renomeie o arquivo" (lista vazia = nenhum).
NOME_ARQUIVO_CUSTOS_RATEADOS = "CUSTOS RATEADOS.xlsx"
ARQUIVO_CUSTOS_RATEADOS_NOMES_ANTIGOS = []

# --------------------------------------------------------------------------
# MÉDICOS - aba própria "MEDICOS" dentro do mesmo arquivo CUSTOS RATEADOS.xlsx,
# com uma linha por mês (CUSTO, VALOR, DATA, STATUS).
#
# Regra de negócio (confirmada com a usuária, setembro/2026): o valor do
# MÊS ATUAL sempre entra como SALDO (ainda não é rateado) - só é rateado
# no mês seguinte, quando finalmente sai do SALDO. Ou seja: pra ratear o
# mês M, a automação usa o(s) valor(es) que ainda estiverem com
# STATUS="SALDO" de mês(es) ANTERIOR(es) a M (normalmente só M-1, mas se
# ficou algum mês esquecido pra trás, soma todos - isso já aconteceu:
# agosto foi consumido por engano e julho ficou parado em SALDO).
#
# A automação só LÊ o STATUS - quem mantém/atualiza SALDO -> CONSUMIDO na
# planilha, e quem adiciona a linha nova de cada mês (sempre como SALDO),
# é a usuária. A automação avisa no console o que consumiu, pra lembrete.
# --------------------------------------------------------------------------
ARQUIVO_MEDICOS_RATEADO = {
    "arquivo_padrao": ["custos_rateados", "custos rateados"],
    "sheet_name": "MEDICOS",
    "coluna_custo": "CUSTO",
    "coluna_valor": "VALOR",
    "coluna_data": "DATA",
    "coluna_status": "STATUS",
    "header_row": 0,
    "categoria_interna": "MEDICOS",
}

# Nomes dos meses em português, para montar o nome da aba em arquivos com
# "aba_por_mes": True (ex: mês 7 de 2026 -> tenta "Julho2026" e "Julho26")
MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

# --------------------------------------------------------------------------
# ABA "Lançamento Contábil": rótulo de cada categoria usado no texto do
# lançamento (ex: "Trf. Est. 3 DIANA_AJUDA DE CUSTO - JULHO/2026 - 250").
# Ajuste aqui se quiser mudar como uma categoria aparece nesse texto.
# --------------------------------------------------------------------------
CATEGORIA_ROTULO_CONTABIL = {
    "TRANSPORTE_UBER": "UBER",
    "TRANSPORTE_TAXI": "TAXI",
    "EXAMES_CERBA": "CERBA",
    "AJUDA_DE_CUSTO": "AJUDA DE CUSTO",
    "MEDICAMENTOS_4BIO": "MEDICAMENTOS",
    "MEDICOS": "MEDICOS",
    "INSUMOS": "INSUMOS",
    "MGR_ANESTESIA": "MGR",
    "TESTES_GENETICOS": "TESTES GENETICOS",
}

# --------------------------------------------------------------------------
# CONFERÊNCIA COM O TOTVS (conta contábil)
# --------------------------------------------------------------------------
# A contabilidade lança no TOTVS o mesmo texto que geramos na aba
# "Lançamento Contábil" (ex: "Trf. Est. 3 FRIDA_MGR - JULHO/2026"), então
# extraímos o rótulo da categoria de dentro do HISTORICO para casar
# automaticamente com a nossa categoria interna.
#
# TOTVS_ROTULO_PARA_CATEGORIA: mapeia o rótulo como aparece no TOTVS (pode
# ser abreviado/diferente do nosso CATEGORIA_ROTULO_CONTABIL) para a
# categoria interna. Inclui "DOE" -> MEDICOS porque foi observado no
# extrato de exemplo que a contabilidade abrevia "MEDICOS" como "DOE" no
# TOTVS - CONFIRME se está certo e ajuste aqui se não estiver.
TOTVS_ROTULO_PARA_CATEGORIA = {
    "UBER": "TRANSPORTE_UBER",
    "TAXI": "TRANSPORTE_TAXI",
    "CERBA": "EXAMES_CERBA",
    "AJUDA DE CUSTO": "AJUDA_DE_CUSTO",
    "MEDICAMENTOS": "MEDICAMENTOS_4BIO",
    "MEDICOS": "MEDICOS",
    "DOE": "MEDICOS",  # abreviação observada da contabilidade p/ MEDICOS - confirmar
    "INSUMOS": "INSUMOS",
    "MGR": "MGR_ANESTESIA",
    "TESTES GENETICOS": "TESTES_GENETICOS",
    "TESTES GENÉTICOS": "TESTES_GENETICOS",
    "LAB GENETICO": "TESTES_GENETICOS",
}

# Lançamentos do TOTVS que NÃO seguem o padrão "Trf. Est. 3 NOME_CATEGORIA"
# (provisões, notas fiscais, itens de medicamento) são classificados por
# palavra-chave no texto do HISTORICO. Ajuste/complete conforme aparecerem
# novos tipos de lançamento nos extratos do TOTVS.
TOTVS_PALAVRAS_CHAVE_CATEGORIA = {
    "MGR_ANESTESIA": ["MGR"],
    "TRANSPORTE_UBER": ["UBER"],
    "TRANSPORTE_TAXI": ["TAXI", "USETAXI"],
    # "RECL. " = reclassificação do valor de cada médico/clínica (ex: "RECL.
    # DRA. NISE SILVEIRA - CLINICA EXEMPLO 09_2026") - o ponto depois de RECL é de
    # propósito, pra NÃO pegar "RECL_MATERIAL..." (que é INSUMOS)
    "MEDICOS": ["MEDICOS", "MÉDICOS", "OVODOACAO", "RECL. "],
    "AJUDA_DE_CUSTO": ["DOADORAS PROGRAMA"],
    # além do "4BIO", entram os remédios que o TOTVS lança pelo nome do
    # produto (sem nenhum rótulo de fornecedor). As marcas são o que
    # apareceu em setembro/2026; as palavras de apresentação (COMPRIMIDO,
    # SERINGA, CARTUCHO) servem de rede de segurança pra remédio novo - se
    # entrar um medicamento de nome diferente e ele cair em
    # NAO_CLASSIFICADO, é só incluir a marca aqui.
    "MEDICAMENTOS_4BIO": ["4BIO", "ORGALUTRAN", "PUREGON", "FOSTIMON", "GONAPEPTYL",
                          "MENOPUR", "DUPHASTON", "CETROTIDE", "COMPRIMIDO", "SERINGA", "CARTUCHO"],
    # reclassificação do material de embriologia (ex: "RECL_MATERIAL
    # EMBRIOLOGIA_DOE_07_2026 10DOADORAS")
    "INSUMOS": ["RECL_MATERIAL", "MATERIAL EMBRIOLOGIA"],
    # nota fiscal do laboratório da Cerba (razão social do laboratório)
    "EXAMES_CERBA": ["LABORATORIO EXEMPLO"],
    # nota fiscal do laboratório de genética (ainda não apareceu no TOTVS - ajuste se o texto vier diferente)
    "TESTES_GENETICOS": ["LAB GENETICO"],
}

# Nome do arquivo persistido com o extrato acumulado da conta do TOTVS
ARQUIVO_TOTVS_ACUMULADO = "totvs_acumulado.csv"

# --------------------------------------------------------------------------
# SALDO INICIAL CONTÁBIL (TOTVS) - valores reais que já estavam na conta
# contábil antes de começar a usar a automação, por categoria (confirmado
# com a usuária a partir do extrato real do TOTVS). Só é usado UMA VEZ - na
# 1ª vez que --totvs é passado para conferencia_doadoras.py - e entra como
# uma linha "SALDO INICIAL <CATEGORIA>" no acumulado. Depois disso, o
# próprio acumulado (totvs_acumulado.csv) carrega esse valor para sempre
# junto com os lançamentos reais - não precisa informar de novo.
SALDO_INICIAL_CONTABIL_TOTVS = {
    "EXAMES_CERBA": 5000.00,
    "AJUDA_DE_CUSTO": 8000.00,
    "MEDICAMENTOS_4BIO": 3000.00,
    "TRANSPORTE_UBER": 400.00,
    "MEDICOS": 6000.00,
    "TRANSPORTE_TAXI": 1200.00,
}
ARQUIVO_SALDO_INICIAL_TOTVS_PROCESSADO = "saldo_inicial_totvs_processado.txt"

# --------------------------------------------------------------------------
# SALDO LEGADO DE MESES ANTERIORES (sem doadora individual) - valores que já
# estavam pendentes na planilha antiga (antes da automação), mas nunca foram
# quebrados por pessoa - só o total por categoria mesmo. Confirmado com a
# usuária (fonte: aba "SALDO pendente" do relatório de referência, linhas
# "MESES ANTERIORES"). Como não tem nome pra casar com doadora nenhuma,
# entra direto na aba "SALDO pendente" do reconciliacao.py como um saldo
# agregado, só 1 vez (ver ARQUIVO_SALDO_LEGADO_PROCESSADO).
SALDO_LEGADO_MESES_ANTERIORES = {
    "AJUDA_DE_CUSTO": 4000.0,
    "EXAMES_CERBA": 2500.0,
    "TRANSPORTE_UBER": 300.0,
    "MEDICAMENTOS_4BIO": 1800.0,
    "MEDICOS": 6000.0,
}
ARQUIVO_SALDO_LEGADO_PROCESSADO = "saldo_legado_meses_anteriores_processado.txt"
# marcador PRÓPRIO do conferencia_doadoras.py - separado de propósito do marcador do
# reconciliacao.py acima, senão quem rodasse primeiro consumiria o marcador e o outro
# relatório nunca mostraria as linhas "MESES ANTERIORES"
ARQUIVO_SALDO_LEGADO_CONFERENCIA_PROCESSADO = "saldo_legado_meses_anteriores_conferencia_processado.txt"

# --------------------------------------------------------------------------
# AJUSTE ACUMULADO (sem doadora individual) - total já convertido em ajuste
# nos meses anteriores à automação, por categoria (mesma fonte que
# SALDO_LEGADO_MESES_ANTERIORES: aba "SALDO pendente" do relatório de
# referência, linhas "AJUSTE"). Só usado no conferencia_doadoras.py, pra
# manter visível o total de ajuste histórico junto com o saldo pendente.
SALDO_LEGADO_AJUSTE_ACUMULADO = {
    "AJUDA_DE_CUSTO": 1500.0,
    "TRANSPORTE_TAXI": 600.0,
    "TRANSPORTE_UBER": 250.0,
}
ARQUIVO_AJUSTE_LEGADO_PROCESSADO = "ajuste_legado_processado.txt"

# Colunas esperadas na planilha mensal que o estagiário preenche
# (cabeçalho está na linha 2 do arquivo, por isso MASTER_HEADER_ROW = 1)
MASTER_HEADER_ROW = 1
MASTER_COLUNA_CODIGO = "Número PIN"
MASTER_COLUNA_NOME = "Nome da doadora"
MASTER_COLUNA_LOTES_DOADOS = "Fresco"
MASTER_COLUNA_LOTES_CONGELADOS = "Lotes"
MASTER_COLUNA_MES = "Mês"  # coluna K - data à qual a doadora pertence
# Mini lote: nova modalidade de venda para lotes pequenos (menos de 6 óvulos), coluna
# que apareceu na lista de doadoras entre "Fresco" e "Total". É OPCIONAL: lista antiga,
# sem essa coluna, continua sendo lida normalmente (mini lotes = 0).
MASTER_COLUNA_MINILOTES = "Mini lote"
# True = os mini lotes entram na "qte de lotes" (coluna I do Controle Estoque), do mesmo
# jeito que a coluna "Total" da lista (que soma Lotes + Fresco + Mini lote). Se a decisão
# for contar só os lotes inteiros, mude para False.
MINILOTES_ENTRAM_NA_QTE_DE_LOTES = True

# Arquivo que guarda o histórico de SALDO entre os meses (não apagar!)
ARQUIVO_CONTROLE_SALDO = "controle_saldo.csv"

# --------------------------------------------------------------------------
# SALDO INICIAL (--saldo-inicial): usado só na primeira vez que a
# automação roda, para trazer os custos que já estavam em aberto ANTES de
# começar a usar o sistema (ex: Saldo_em_aberto_agosto_doadoras.xlsx).
# Cada aba do arquivo é um fornecedor diferente.
#
# Esses itens passam pela MESMA regra de sempre: se baterem com uma
# doadora agora, entram no Lançamento Contábil deste mês; se não baterem,
# viram SALDO usando a DATA ORIGINAL de cada item (não a data de hoje),
# para a contagem dos 3 meses ficar correta desde o início.
# --------------------------------------------------------------------------
# A coluna "UTILIZADO" NÃO é mais usada como filtro (ver reconciliacao.py /
# carregar_saldo_inicial): confirmado com a usuária que hoje ela só serve de
# anotação manual (data em que o item "deveria ter sido utilizado"; itens
# realmente em aberto ficam com o texto "SALDO" nessa coluna) - não é um
# critério para decidir o que entra no matching. TODAS as linhas com nome e
# valor da aba entram no matching; quem bate vira lançamento do mês, quem não
# bate vira SALDO (com a data original), exatamente como qualquer outro custo.
SALDO_INICIAL_ABAS = {
    "Taxi": {"categoria": "TRANSPORTE_TAXI", "coluna_valor": "VALOR", "coluna_nome": "NOME", "coluna_data": "DATA"},
    "Cerba": {"categoria": "EXAMES_CERBA", "coluna_valor": "VALOR", "coluna_nome": "NOME DO PACIENTE", "coluna_data": "DATA"},
    "Ajuda de Custo": {"categoria": "AJUDA_DE_CUSTO", "coluna_valor": "VL TOTAL", "coluna_nome": "NOME", "coluna_data": "DATA"},
    "Uber": {"categoria": "TRANSPORTE_UBER", "coluna_valor": "VALOR", "coluna_nome": "NOME", "coluna_data": "DATA"},
    "Medicamentos": {"categoria": "MEDICAMENTOS_4BIO", "coluna_valor": "VL TOTAL", "coluna_nome": "PACIENTE/DOADORA", "coluna_data": "DATA"},
}


# --------------------------------------------------------------------------
# AJUSTES DIVERSOS na aba "Posição Consolidada" (reconciliacao.py): sempre que
# algo cai em AJUSTE num mês, o valor entra numa linha chamada "AJUSTES
# DIVERSOS", por categoria, no bloco daquele mês. Daqui pra frente isso é
# automático (vem do que a reconciliação classifica como AJUSTE).
#
# Meses fechados ANTES da automação podem ser informados aqui, no formato
# {"AAAA-MM": {categoria: valor}} - entram 1x no histórico. Vazio = nenhum.
NOME_LINHA_AJUSTES_DIVERSOS = "AJUSTES DIVERSOS"
AJUSTES_DIVERSOS_HISTORICO = {}
