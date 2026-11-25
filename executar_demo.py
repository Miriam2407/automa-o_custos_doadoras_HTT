"""
Executa a DEMONSTRAÇÃO completa com dados fictícios (2 meses, em sequência)
==============================================================================

    python demo/executar_demo.py

O que acontece (tudo dentro da pasta demo/, nada fora dela):

  1. apaga demo/controle e demo/saida (começa do zero) e gera os dados fictícios;
  2. AGOSTO/2026:  reconciliação  ->  extrato do TOTVS (gerado a partir dos lançamentos)
                   ->  conferência  ->  "ajustes manuais" simulados;
  3. SETEMBRO/2026: reconciliação (com o saldo em aberto de antes da automação)
                   ->  extrato do TOTVS  ->  conferência, que PARTE do arquivo de agosto
                   (mantendo os ajustes) e acrescenta só os lançamentos novos.

Os "ajustes manuais" são o trabalho que a analista faz à mão no Excel: linhas AJUSTE no
SALDO pendente e na aba TOTVS pra a diferença do Resumo chegar a zero. Aqui o script
simula isso, só pra demonstração.

Resultado em demo/saida/: relatorio_reconciliacao_AAAA-MM.xlsx e conferencia_doadoras_AAAA-MM.xlsx
"""

import calendar
import datetime as dt
import os
import random
import shutil
import subprocess
import sys

import openpyxl
import pandas as pd
from openpyxl import Workbook

DEMO = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(DEMO)
sys.path.insert(0, DEMO)
import gerar_dados_demo  # noqa: E402

MESES = ["2026-08", "2026-09"]
ABREV = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
ROTULOS = ["AJUDA DE CUSTO", "CERBA", "INSUMOS", "MEDICAMENTOS", "MEDICOS", "MGR", "TAXI", "UBER"]


def rodar(script, *args):
    """Roda um script do projeto, mostra a saída na tela E guarda em demo/logs/ (serve de print)."""
    cmd = [sys.executable, os.path.join(RAIZ, script), *args]
    print("\n$ python", script, " ".join(args))
    resultado = subprocess.run(cmd, cwd=DEMO, capture_output=True, text=True, encoding="utf-8")
    print(resultado.stdout, end="")
    if resultado.returncode != 0:
        print(resultado.stderr)
        raise SystemExit(f"{script} terminou com erro")
    os.makedirs(os.path.join(DEMO, "logs"), exist_ok=True)
    mes = args[args.index("--mes") + 1]
    with open(os.path.join(DEMO, "logs", f"{os.path.splitext(script)[0]}_{mes}.txt"), "w", encoding="utf-8") as f:
        f.write(resultado.stdout)


def formatar_saldo(valor):
    texto = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{texto} {'D' if valor >= 0 else 'C'}"


_SALDO_CORRENTE = [50000.0]      # o saldo do extrato de um mês continua de onde o do mês anterior parou


def gerar_extrato_totvs(mes):
    """Extrato fictício do TOTVS: provisões, notas fiscais e os 'Trf. Est. 3' que a contabilidade
    lança a partir da nossa aba Lançamento Contábil (a maioria - alguns ficam de fora de propósito)."""
    lanc = pd.read_excel(os.path.join(DEMO, "saida", f"relatorio_reconciliacao_{mes}.xlsx"), sheet_name="Lançamento Contábil")
    lanc = lanc[lanc["lancamento"].notna()]
    ano, m = map(int, mes.split("-"))
    ultimo = dt.datetime(ano, m, calendar.monthrange(ano, m)[1])
    rng = random.Random(int(mes.replace("-", "")) + 7)
    yy = ano % 100
    linhas = [
        (dt.datetime(ano, m, 1), f"EST. PROV. UBER {ABREV[m]}.{yy} UNIDADE EXEMPLO", 0.0, 1500.00),
        (dt.datetime(ano, m, 1), f"PROV. MGR {ABREV[m]}.{yy} UNIDADE EXEMPLO", 0.0, 5400.00),
        (ultimo - dt.timedelta(days=3), "NF 1021 - LABORATORIO EXEMPLO LTDA CNPJ/CPF 00000000000000", 2500.00, 0.0),
        (ultimo - dt.timedelta(days=3), "NF 88 - TAXI EXEMPLO ME CNPJ/CPF 00000000000000", 600.00, 0.0),
    ]
    for _, r in lanc.iterrows():
        if rng.random() < 0.85:
            linhas.append((ultimo, str(r["lancamento"]), 0.0, round(float(r["valor"]), 2)))
    linhas.sort(key=lambda x: x[0])

    wb = Workbook()
    ws = wb.active
    ws.title = "Lançamentos Contábeis"
    ws.append(["Lançamentos Contábeis"])
    ws.append(["CONTA", "DATA", "LOTE/SUB/DOC/LINHA", "HISTORICO", "C/PARTIDA", "FILIAL DE ORIGEM", "C CUSTO",
               "ITEM CONTA", "COD CL VAL", "DEBITO", "CREDITO", "SALDO ATUAL"])
    saldo = _SALDO_CORRENTE[0]
    for i, (data, historico, debito, credito) in enumerate(linhas, start=1):
        saldo += debito - credito
        ws.append(["CONTA-DEMO", data, f"{mes.replace('-', '')}{i:09d}", historico, None, "000001", "3.1.1.7", None, None,
                   debito, credito, formatar_saldo(saldo)])
    _SALDO_CORRENTE[0] = saldo
    caminho = os.path.join(DEMO, "dados_entrada", mes, "extrato_totvs.xlsx")
    wb.save(caminho)
    return caminho


def _somar_por_categoria(ws, col_categoria, col_valor):
    soma = {}
    for r in range(2, ws.max_row + 1):
        cat, valor = ws.cell(row=r, column=col_categoria).value, ws.cell(row=r, column=col_valor).value
        if cat and isinstance(valor, (int, float)):
            soma[cat] = soma.get(cat, 0.0) + float(valor)
    return soma


def simular_ajustes_manuais(mes, ajustes_saldo=None):
    """O que a analista faz à mão no Excel: (1) cria linhas AJUSTE no SALDO pendente e (2) acrescenta
    linhas AJUSTE na aba TOTVS pra diferença do Resumo ficar em zero."""
    caminho = os.path.join(DEMO, "saida", f"conferencia_doadoras_{mes}.xlsx")
    wb = openpyxl.load_workbook(caminho)
    sp, tv = wb["SALDO pendente"], wb["TOTVS"]

    for categoria, valor in (ajustes_saldo or {}).items():
        sp.append([None, categoria, valor, "AJUSTE", None])
        sp.cell(row=sp.max_row, column=3).number_format = "#,##0.00"

    saldo = _somar_por_categoria(sp, col_categoria=2, col_valor=3)
    contab = _somar_por_categoria(tv, col_categoria=6, col_valor=5)
    adicionadas = []
    for rotulo in ROTULOS:
        diferenca = round(saldo.get(rotulo, 0.0) - contab.get(rotulo, 0.0), 2)
        if abs(diferenca) > 0.005:
            tv.append([None, "AJUSTE", None, None, diferenca, rotulo, None, None, None])
            tv.cell(row=tv.max_row, column=5).number_format = "#,##0.00"
            adicionadas.append((rotulo, diferenca))
    wb.save(caminho)
    print(f"   [simulação] {len(ajustes_saldo or {})} linha(s) AJUSTE no SALDO pendente e "
          f"{len(adicionadas)} linha(s) AJUSTE na aba TOTVS (pra zerar o Resumo)")


def main():
    for pasta in ("controle", "saida", "logs"):
        shutil.rmtree(os.path.join(DEMO, pasta), ignore_errors=True)
    gerar_dados_demo.gerar(MESES)

    for mes in MESES:
        print(f"\n{'=' * 78}\n  {mes}\n{'=' * 78}")
        args = ["--mes", mes, "--lista-doadoras", f"dados_entrada/{mes}/lista_doadoras_{mes}.xlsx",
                "--custos", f"dados_entrada/{mes}", "--medicamentos-liberado", "nao"]
        if mes == MESES[1]:   # o saldo em aberto de antes da automação entra no 2º mês
            args += ["--saldo-inicial", "dados_entrada/saldo_em_aberto_demo.xlsx"]
        rodar("reconciliacao.py", *args)

        extrato = gerar_extrato_totvs(mes)
        rodar("conferencia_doadoras.py", "--mes", mes, "--totvs", os.path.relpath(extrato, DEMO))

        simular_ajustes_manuais(mes, ajustes_saldo={"CERBA": 500.00, "AJUDA DE CUSTO": 1500.00} if mes == MESES[0] else None)

    print(f"\nPronto. Relatórios em: {os.path.join(DEMO, 'saida')}")


if __name__ == "__main__":
    main()
