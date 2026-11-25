"""
Gerador de dados 100% FICTÍCIOS para a demonstração
=====================================================

Cria, em demo/dados_entrada/, os arquivos de entrada que a automação espera - nos mesmos
layouts das planilhas reais, mas com personagens em vez de pessoas:

  * doadoras  -> mulheres importantes da história (Marie Curie, Frida Kahlo, ...) e
                 super-heroínas (Diana Prince, Carol Danvers, ...). Nomes de pessoas
                 falecidas e de personagens de ficção, usados só como exemplo.
  * "estranhas" -> nomes que aparecem nas planilhas de custo mas NUNCA viram doadora
                 (Selina Kyle, Peggy Carter, ...): servem pra mostrar o fluxo
                 SALDO -> AJUSTE depois de 3 meses.
  * testes genéticos -> fornecedor que só entra em SALDO (PINs no lugar de nomes, sem vínculo)
  * erros de digitação de propósito ("Marie Currie", "Frida Kalo") -> mostram o
                 matching aproximado (fuzzy) achando a doadora certa.

Nenhum valor, nome, CPF/CNPJ ou endereço aqui é real. A semente do sorteio é fixa: rodar
de novo gera exatamente os mesmos arquivos.
"""

import datetime as dt
import os
import random
import calendar

from openpyxl import Workbook

PASTA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados_entrada")

MESES_PT = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
            7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}

# ---------------------------------------------------------------------------------------
# Elenco
# ---------------------------------------------------------------------------------------
DOADORAS = {
    "2026-08": ["Marie Curie", "Ada Lovelace", "Frida Kahlo", "Rosa Parks",
                "Diana Prince", "Carol Danvers", "Natasha Romanoff", "Ororo Munroe"],
    "2026-09": ["Clarice Lispector", "Tarsila do Amaral", "Nise da Silveira", "Cleópatra",
                "Barbara Gordon", "Jean Grey", "Kamala Khan", "Hope van Dyne", "Rosalind Franklin"],
}

# nomes que aparecem nas planilhas de custo do mês mas nunca são doadoras ("--" = sem nome)
ESTRANHAS = {
    "2026-08": ["Selina Kyle", "Zatanna Zatara", "--"],
    "2026-09": ["Gamora", "Nebula", "--"],
}

# como a planilha de cada fornecedor às vezes escreve o nome (matching aproximado resolve)
VARIANTES = {
    "Marie Curie": ["Marie Currie", "Marie Curie "], "Frida Kahlo": ["Frida Kalo"],
    "Natasha Romanoff": ["Natasha Romanov"], "Ororo Munroe": ["Ororo Monroe"],
    "Carol Danvers": ["Carol Denvers"], "Kamala Khan": ["Kamala Kahn"],
    "Tarsila do Amaral": ["Tarsila Do Amaral"], "Nise da Silveira": ["Nise da Silveira "],
    "Rosalind Franklin": ["Rosalind Franklyn"], "Clarice Lispector": ["Clarice Lispector "],
}

EXAMES = [8.80, 10.80, 15.70, 25.50, 72.38, 96.40, 170.50, 188.37]
CORRIDAS_TAXI = [45.00, 80.00, 120.00, 150.00, 230.00, 320.00]
AJUDAS = [250.0, 1450.0, 2250.0, 3700.0]

# saldo em aberto de ANTES da automação: (nome, valor, data). Os de junho viram AJUSTE
# em setembro (3 meses); os de julho continuam SALDO; os que batem com doadora viram
# lançamento do mês.
SALDO_EM_ABERTO = {
    "Cerba": [("Jessica Jones", 320.50, dt.datetime(2026, 6, 8)), ("Jessica Jones", 118.20, dt.datetime(2026, 6, 9)),
              ("Peggy Carter", 96.40, dt.datetime(2026, 7, 14)), ("Shuri", 210.00, dt.datetime(2026, 7, 5)),
              ("Diana Prince", 150.00, dt.datetime(2026, 7, 10)), ("Marie Currie", 88.30, dt.datetime(2026, 7, 22))],
    "Taxi": [("Okoye", 180.00, dt.datetime(2026, 6, 12)), ("Peggy Carter", 95.00, dt.datetime(2026, 6, 15)),
             ("Shuri", 140.00, dt.datetime(2026, 7, 18))],
    "Ajuda de Custo": [("Jessica Jones", 1450.00, dt.datetime(2026, 6, 18)), ("Okoye", 250.00, dt.datetime(2026, 7, 20))],
    "Uber": [("Peggy Carter", 42.30, dt.datetime(2026, 6, 3)), ("--", 61.90, dt.datetime(2026, 6, 11))],
    "Medicamentos": [("Jessica Jones", 640.00, dt.datetime(2026, 6, 2)), ("Peggy Carter", 385.50, dt.datetime(2026, 7, 14)),
                     ("Ororo Munroe", 120.00, dt.datetime(2026, 7, 12))],
}


def _rng(mes):
    return random.Random(f"demo-{mes}")   # semente de texto: mesma sequência em qualquer máquina


def _data(mes, rng, com_hora=False):
    ano, m = map(int, mes.split("-"))
    dia = rng.randint(1, min(28, calendar.monthrange(ano, m)[1]))
    d = dt.datetime(ano, m, dia)
    return d + dt.timedelta(hours=rng.randint(7, 18), minutes=rng.randint(0, 59)) if com_hora else d


def _nome(base, rng):
    """às vezes devolve uma grafia diferente da lista de doadoras"""
    if base in VARIANTES and rng.random() < 0.35:
        return rng.choice(VARIANTES[base])
    return base


def _quem_aparece(mes, rng):
    """doadoras do mês + as dos meses anteriores (custo de uma doadora antiga pode chegar depois)
    + as 'estranhas' do mês"""
    atuais = DOADORAS[mes]
    anteriores = [d for m, lst in DOADORAS.items() if m < mes for d in lst]
    return atuais, anteriores, ESTRANHAS[mes]


# ---------------------------------------------------------------------------------------
# Arquivos
# ---------------------------------------------------------------------------------------

def escrever_lista(mes, pasta):
    """Agosto: layout ANTIGO (um mês por arquivo). Setembro: layout NOVO - coluna 'Mini lote' e os meses
    acumulados na mesma aba (a automação usa só as linhas do mês processado)."""
    ano, m = map(int, mes.split("-"))
    novo_layout = mes >= "2026-09"
    wb = Workbook()
    ws = wb.active
    ws.title = "Controle Doadoras"
    cab = ["Data (e-mail)", "Número PIN", "Nome da doadora", "Lotes", "Fresco"]
    cab += ["Mini lote", "Total"] if novo_layout else ["Total"]
    ws.append(cab + ["backup", "Obs", "Status", "Ok controle?", "Mês"])
    rng = _rng(mes + "lista")
    meses = [mm for mm in DOADORAS if mm <= mes] if novo_layout else [mes]
    for mm_ in meses:
        a_, m_ = map(int, mm_.split("-"))
        for i, nome in enumerate(DOADORAS[mm_]):
            cong, fresco = rng.randint(0, 6), rng.randint(0, 2)
            mini = 1 if (novo_layout and i % 4 == 0) else 0
            linha = [dt.datetime(a_, m_, rng.randint(1, 28)), pin_da_doadora(mm_, i), nome, cong, fresco]
            linha += [mini, None] if novo_layout else [cong + fresco]
            ws.append(linha + [None, None, "OK", "OK", dt.datetime(a_, m_, 1)])
    if novo_layout:       # linha com PIN e SEM nome (a automação avisa e ignora)
        ws.append([dt.datetime(ano, m, 20), 100999, None, 1, 0, 0, None, None, None, None, None, dt.datetime(ano, m, 1)])
    ws.append([None, None, "Total Ciclo >>>"])        # rodapé que a automação descarta
    wb.save(os.path.join(pasta, f"lista_doadoras_{mes}.xlsx"))


def pin_da_doadora(mes, indice):
    a, m = map(int, mes.split("-"))
    return 100000 + m * 100 + indice


def escrever_uber(mes, pasta):
    rng = _rng(mes + "uber")
    atuais, anteriores, estranhas = _quem_aparece(mes, rng)
    linhas = []
    for nome in atuais + rng.sample(anteriores, min(2, len(anteriores))):
        for _ in range(rng.randint(1, 3)):
            linhas.append((_data(mes, rng, True), round(rng.uniform(14, 95), 2), _nome(nome, rng), "Doadora"))
    for nome in estranhas:
        for _ in range(rng.randint(1, 2)):
            linhas.append((_data(mes, rng, True), round(rng.uniform(12, 70), 2), nome, "Doadora"))
    linhas.append((_data(mes, rng, True), 38.50, "Funcionária Exemplo", "Funcionária"))   # filtrada: não é doadora
    linhas.sort(key=lambda x: x[0])
    wb = Workbook()
    ws = wb.active
    ws.title = "relatorio"
    if mes >= "2026-09":      # layout novo do export: colunas TIPO e Valor (a automação aceita os dois)
        ws.append(["Data", "Nome", "Sobrenome", "ID do funcionário", "Grupo", "Tipo de transação",
                   "Valor", "Nome do convidado", "TIPO"])
    else:
        ws.append(["Data", "Nome", "Sobrenome", "ID do funcionário", "Grupo", "Tipo de transação",
                   "Valor da transação: BRL", "Nome do convidado", "Tipo"])
    for data, valor, convidado, tipo in linhas:
        ws.append([data, "ATENDIMENTO", "CLINICA EXEMPLO", "PROGRAMA DOADORAS", "LOGISTICA", "Fare", valor, convidado, tipo])
    wb.save(os.path.join(pasta, "UBER.xlsx"))


def escrever_taxi(mes, pasta):
    rng = _rng(mes + "taxi")
    atuais, anteriores, estranhas = _quem_aparece(mes, rng)
    linhas = []
    for nome in rng.sample(atuais, max(3, len(atuais) // 2)) + estranhas[:1]:
        for _ in range(rng.randint(1, 2)):
            linhas.append((_data(mes, rng), rng.choice(CORRIDAS_TAXI), _nome(nome, rng)))
    linhas.sort(key=lambda x: x[0])
    wb = Workbook()
    ws = wb.active
    ws.title = "Corridas"
    ws.append(["Transportes - Táxi (dados de demonstração)"])
    ws.append(["Data", "Horário", "Passageiro", "Endereço de Origem", "Endereço de Destino ", "Valor"])
    for data, valor, nome in linhas:
        ws.append([data, "08:30", nome, "Rua das Flores, 100", "Av. Exemplo, 2000", valor])
    ws.append([None, None, None, None, "TOTAL:", round(sum(l[1] for l in linhas), 2)])
    wb.save(os.path.join(pasta, "TAXI.xlsx"))


def escrever_cerba(mes, pasta):
    rng = _rng(mes + "cerba")
    atuais, anteriores, estranhas = _quem_aparece(mes, rng)
    wb = Workbook()
    ws = wb.active
    ws.title = "Exames Encaminhados"
    ws.append(["Relatório de exames encaminhados (dados de demonstração)"])
    ws.append(["Convênio: CLINICA EXEMPLO"])
    ws.append(["GUIA", "DATA SOLICITAÇÃO", "NOME DO PACIENTE", "EXAME", "VALOR"])
    n_pacientes, n_guias = 0, 0
    for nome in atuais + estranhas[:2]:
        n_pacientes += 1
        data = _data(mes, rng, True)
        for k in range(rng.randint(3, 7)):
            n_guias += 1
            ws.append([f"G{n_guias:05d}", data + dt.timedelta(seconds=k * 3), _nome(nome, rng), f"EX{rng.randint(100, 999)}", rng.choice(EXAMES)])
    ws.append(["Total de Pacientes:", n_pacientes, None, "Total de Guias:", n_guias])
    wb.save(os.path.join(pasta, "CERBA.xlsx"))


def escrever_ajuda(mes, pasta):
    ano, m = map(int, mes.split("-"))
    rng = _rng(mes + "ajuda")
    atuais, anteriores, estranhas = _quem_aparece(mes, rng)
    wb = Workbook()
    ws = wb.active
    ws.title = f"{MESES_PT[m]}{ano}"
    ws.append([])
    ws.append([None, None, None, "Voluntárias"])
    ws.append([])
    ws.append([])
    ws.append([None, "Data envio", "Data do rebibo", "PIX", "Nome", "Obs", "Valor", "Lib.", "Obs"])
    total = 0.0
    for nome in rng.sample(atuais, max(3, len(atuais) // 2)) + estranhas[:1]:
        valor = rng.choice(AJUDAS)
        total += valor
        d = _data(mes, rng)
        ws.append([None, d, d, None, _nome(nome, rng), rng.choice(["Início", "Coleta", "Pós Coleta", "Completo"]), valor, None, None])
    ws.append([None, None, None, None, None, "TOTAL", total])
    wb.save(os.path.join(pasta, "AJUDA DE CUSTO.xlsx"))


def escrever_genetica(mes, pasta):
    """Fornecedor de TESTES GENÉTICOS: o arquivo traz PINs no lugar de nomes e é ACUMULADO (setembro repete
    agosto). Os PINs são de doadoras que existem na lista - de propósito: a categoria entra SÓ em saldo e
    nunca é vinculada a ninguém. Há também uma linha com o nome EXATO de uma doadora, pelo mesmo motivo."""
    rng = _rng(mes + "genetica")
    meses = [mm for mm in DOADORAS if mm <= mes]
    wb = Workbook()
    resumo = wb.active
    resumo.title = "Planilha2"                       # aba de resumo (tabela dinâmica): NÃO é a que se lê
    dados = wb.create_sheet("Planilha1")
    dados.append(["mês", "NomePaciente", "Procedimentos", "exame", "QUANTIDADE", "Valor bruto", "nota fiscal"])
    por_nota = {}
    formatos = ["DOADORA PIN {}", "OD PIN {}", "PIN DOADORA {}", "DOADORA {}"]
    for mm_ in meses:
        a_, m_ = map(int, mm_.split("-"))
        nota = 5100 + m_
        for i in sorted(rng.sample(range(len(DOADORAS[mm_])), 6)):
            valor = rng.choice([1350, 2000, 3000])
            dados.append([dt.datetime(a_, m_, 1), rng.choice(formatos).format(pin_da_doadora(mm_, i)),
                          "GEN001", "PAINEL GENETICO", 1, valor, nota])
            por_nota[nota] = por_nota.get(nota, 0) + valor
    if mes >= "2026-09":
        dados.append([dt.datetime(2026, 9, 1), "Diana Prince", "GEN001", "PAINEL GENETICO", 1, 2000, 5109])
        por_nota[5109] = por_nota.get(5109, 0) + 2000
    resumo.append([])
    resumo.append(["Rótulos de Linha", "Soma de Valor bruto"])
    for nota, total in por_nota.items():
        resumo.append([nota, total])
    resumo.append(["Total Geral", sum(por_nota.values())])
    wb.save(os.path.join(pasta, "GENETICA.xlsx"))


def escrever_custos_rateados(mes, pasta):
    ano, m = map(int, mes.split("-"))
    ultimo = dt.datetime(ano, m, calendar.monthrange(ano, m)[1])
    wb = Workbook()
    ws = wb.active
    ws.title = "CUSTO"
    ws.append(["CUSTO", "VALOR", "DATA"])
    ws.append(["MATERIAL DE EMBRIOLOGIA", 1850.00, ultimo])     # valor POR doadora
    ws.append(["MGR", 480.00, ultimo])                           # valor POR doadora
    wm = wb.create_sheet("MEDICOS")
    wm.append(["CUSTO", "VALOR", "DATA", "STATUS"])
    # regra do mês defasado: o valor do mês atual fica em SALDO; é o do mês ANTERIOR que é rateado
    historico = [("2026-07", 15200.00), ("2026-08", 16100.00), ("2026-09", 16850.00)]
    for mes_linha, valor in historico:
        if mes_linha > mes:
            continue
        a, mm = map(int, mes_linha.split("-"))
        fim = dt.datetime(a, mm, calendar.monthrange(a, mm)[1])
        status = "SALDO" if mes_linha >= mes or mes_linha == _mes_anterior(mes) else "CONSUMIDO"
        wm.append(["MÉDICOS", valor, fim, status])
    wb.save(os.path.join(pasta, "CUSTOS RATEADOS.xlsx"))


def _mes_anterior(mes):
    a, m = map(int, mes.split("-"))
    return f"{a - 1}-12" if m == 1 else f"{a}-{m - 1:02d}"


def escrever_saldo_em_aberto(pasta):
    cols = {"Cerba": ("VALOR", "NOME DO PACIENTE"), "Taxi": ("VALOR", "NOME"), "Ajuda de Custo": ("VL TOTAL", "NOME"),
            "Uber": ("VALOR", "NOME"), "Medicamentos": ("VL TOTAL", "PACIENTE/DOADORA")}
    wb = Workbook()
    wb.remove(wb.active)
    for aba, (col_valor, col_nome) in cols.items():
        ws = wb.create_sheet(aba)
        ws.append([col_valor, col_nome, "UTILIZADO", "DATA"])
        for nome, valor, data in SALDO_EM_ABERTO[aba]:
            ws.append([valor, nome, "SALDO", data])
        ws.append([None])
        ws.append([round(sum(v for _, v, _ in SALDO_EM_ABERTO[aba]), 2), "TOTAL"])
    wb.save(os.path.join(pasta, "saldo_em_aberto_demo.xlsx"))


def gerar(meses=("2026-08", "2026-09")):
    os.makedirs(PASTA, exist_ok=True)
    for mes in meses:
        pasta = os.path.join(PASTA, mes)
        os.makedirs(pasta, exist_ok=True)
        escrever_lista(mes, pasta)
        escrever_uber(mes, pasta)
        escrever_taxi(mes, pasta)
        escrever_cerba(mes, pasta)
        escrever_ajuda(mes, pasta)
        escrever_genetica(mes, pasta)
        escrever_custos_rateados(mes, pasta)
    escrever_saldo_em_aberto(PASTA)
    print(f"Dados fictícios gerados em '{PASTA}'")


if __name__ == "__main__":
    gerar()
