"""
Gera as IMAGENS (prints) da demonstração em docs/imagens/
==========================================================

    python demo/executar_demo.py     # antes: gera os relatórios fictícios
    python demo/gerar_prints.py      # depois: transforma abas escolhidas em PNG

Precisa do LibreOffice (comando 'soffice') e do 'pdftoppm' (poppler) instalados - só pra gerar
as imagens; a automação em si NÃO depende deles. Se você preferir, tire os prints à mão
abrindo os arquivos de demo/saida/ no Excel (veja docs/GUIA_DE_PRINTS.md).

Como funciona: o LibreOffice recalcula as fórmulas (o Resumo usa SUMIF) -> cada aba escolhida
vira uma "vitrine" (só as linhas que interessam, com cabeçalho A,B,C.../1,2,3... como no Excel)
-> PDF -> PNG recortado. A imagem do log reproduz a caixa preta da interface com o texto REAL
gravado em demo/logs/.
"""

import copy
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter
from PIL import Image, ImageChops, ImageDraw, ImageFont

DEMO = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(DEMO, "saida")
IMAGENS = os.path.join(os.path.dirname(DEMO), "docs", "imagens")

# (nome da imagem, arquivo, aba, quais linhas mostrar)
#   "todas" | ("topo", n) | ("topo_e_fim", n_topo, n_fim) | "totvs"  - sempre inclui o cabeçalho (linha 1). As linhas NÃO escolhidas
#   ficam ocultas (como num filtro do Excel), então a numeração de linhas da imagem é a verdadeira.
VITRINES = [
    ("01_saldo_pendente", "relatorio_reconciliacao_2026-09.xlsx", "SALDO pendente", ("topo_e_fim", 14, 8)),
    ("02_ajuste_do_mes", "relatorio_reconciliacao_2026-09.xlsx", "AJUSTE do mês", "todas"),
    ("03_lancamento_contabil", "relatorio_reconciliacao_2026-09.xlsx", "Lançamento Contábil", ("topo", 16)),
    ("04_posicao_consolidada", "relatorio_reconciliacao_2026-09.xlsx", "Posição Consolidada", "todas"),
    ("05_resumo_conferencia", "conferencia_doadoras_2026-09.xlsx", "Resumo", "todas"),
    ("06_totvs_com_historico", "conferencia_doadoras_2026-09.xlsx", "TOTVS", "totvs"),
    ("08_testes_geneticos", "relatorio_reconciliacao_2026-09.xlsx", "TESTES_GENETICOS", ("topo", 16)),
]


def recalcular(arquivo, pasta_tmp):
    """LibreOffice abre e salva de novo -> as fórmulas (SUMIF) ganham valor calculado."""
    destino = os.path.join(pasta_tmp, "recalculado")
    os.makedirs(destino, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", destino, arquivo],
                   check=True, capture_output=True, timeout=240)
    return os.path.join(destino, os.path.basename(arquivo))


def formatar_br(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def escolher_linhas(ws, regra):
    ultima = ws.max_row
    if regra == "todas":
        return list(range(1, ultima + 1))
    if regra == "totvs":
        # cabeçalho + saldos iniciais + 1as linhas + TODAS as linhas AJUSTE + 1as linhas depois do 1º bloco de AJUSTE
        sementes, normais, ajustes = [], [], []
        for r in range(2, ultima + 1):
            historico, origem = ws.cell(row=r, column=2).value, ws.cell(row=r, column=7).value
            if not historico:
                continue
            (sementes if origem == "Saldo Inicial" else ajustes if historico == "AJUSTE" else normais).append(r)
        depois_do_1o_ajuste = [r for r in normais if r > min(ajustes)][:3]
        return sorted(set([1] + sementes + normais[:3] + ajustes + depois_do_1o_ajuste))
    if regra[0] == "topo_e_fim":
        topo = list(range(1, min(ultima, regra[1] + 1) + 1))
        return topo + list(range(max(ultima - regra[2] + 1, topo[-1] + 1), ultima + 1))
    return list(range(1, min(ultima, regra[1] + 1) + 1))


def montar_vitrine(caminho_recalculado, aba, regra, destino_xlsx):
    origem = openpyxl.load_workbook(caminho_recalculado, data_only=True)[aba]
    mostrar = escolher_linhas(origem, regra)
    cabecalho = {c: str(origem.cell(row=1, column=c).value or "").strip() for c in range(1, origem.max_column + 1)}
    n_colunas = max(c for r in mostrar for c in range(1, origem.max_column + 1)
                    if origem.cell(row=r, column=c).value not in (None, ""))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = aba
    for r in mostrar:                       # cada linha vai pra MESMA posição que tem no arquivo real
        for c in range(1, n_colunas + 1):
            c_origem = origem.cell(row=r, column=c)
            valor = c_origem.value
            destino = ws.cell(row=r, column=c)
            if isinstance(valor, (int, float)) and not isinstance(valor, bool) and r > 1 and cabecalho[c] == "score_match":
                destino.value = f"{valor:.1f}".replace(".", ",")      # só na imagem: 1 casa decimal
                destino.alignment = Alignment(horizontal="right")
            elif isinstance(valor, (int, float)) and not isinstance(valor, bool) and (
                    "0.00" in (c_origem.number_format or "") or (r > 1 and cabecalho[c] == "valor")):
                destino.value = formatar_br(valor)            # padrão brasileiro, como o Excel em pt-BR
                destino.alignment = Alignment(horizontal="right")
            else:
                destino.value = valor
                if c_origem.alignment is not None:
                    destino.alignment = Alignment(horizontal=c_origem.alignment.horizontal)
            if c_origem.has_style:
                destino.font, destino.fill = copy.copy(c_origem.font), copy.copy(c_origem.fill)
    for r in range(1, origem.max_row + 1):
        if r not in mostrar:
            ws.row_dimensions[r].hidden = True
    # largura pelo conteúdo (o original usa largura fixa e cortaria os textos longos)
    for c in range(1, n_colunas + 1):
        maior = max((len(str(ws.cell(row=r, column=c).value or "")) for r in mostrar), default=8)
        ws.column_dimensions[get_column_letter(c)].width = min(max(maior + 5, 12), 95)
    # aparência de "print do Excel": grade e cabeçalhos de linha/coluna
    ws.print_options.gridLines = True
    ws.print_options.headings = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8                                # A3
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = ws.page_margins.bottom = 0.2
    wb.save(destino_xlsx)


def xlsx_para_png(xlsx, png, pasta_tmp, dpi=150):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", pasta_tmp, xlsx], check=True, capture_output=True, timeout=240)
    pdf = os.path.join(pasta_tmp, os.path.splitext(os.path.basename(xlsx))[0] + ".pdf")
    prefixo = os.path.join(pasta_tmp, "pagina")
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", "1", "-l", "1", pdf, prefixo], check=True)
    pagina = [f for f in os.listdir(pasta_tmp) if f.startswith("pagina") and f.endswith(".png")][0]
    imagem = Image.open(os.path.join(pasta_tmp, pagina)).convert("RGB")
    fundo = Image.new("RGB", imagem.size, (255, 255, 255))
    caixa = ImageChops.difference(imagem, fundo).getbbox()
    if caixa:
        e, t, d, b = caixa
        imagem = imagem.crop((max(e - 14, 0), max(t - 14, 0), min(d + 14, imagem.width), min(b + 14, imagem.height)))
    imagem.save(png, optimize=True)
    os.remove(os.path.join(pasta_tmp, pagina))


def imagem_do_log(arquivo_log, png, titulo="Saída da execução"):
    """Reproduz a caixa preta com texto verde da interface, com o texto REAL do log da demo."""
    with open(arquivo_log, encoding="utf-8") as f:
        linhas = [l.rstrip() for l in f.read().splitlines()]
    largura_max = 112
    quebradas = []
    for l in linhas:
        while len(l) > largura_max:
            quebradas.append(l[:largura_max])
            l = "      " + l[largura_max:]
        quebradas.append(l)
    fonte = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf", 17)
    fonte_titulo = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
    alt_linha, margem, barra = 24, 18, 38
    largura = int(max(fonte.getlength(l) for l in quebradas)) + 2 * margem
    altura = barra + margem * 2 + alt_linha * len(quebradas)
    img = Image.new("RGB", (largura, altura), (17, 17, 17))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, largura, barra], fill=(46, 84, 150))
    d.text((margem, 9), titulo, font=fonte_titulo, fill=(255, 255, 255))
    y = barra + margem
    for l in quebradas:
        cor = (255, 196, 61) if l.lstrip().startswith("[aviso]") or l.lstrip().startswith("[info]") else (0, 255, 0)
        d.text((margem, y), l, font=fonte, fill=cor)
        y += alt_linha
    img.save(png, optimize=True)


def main():
    for ferramenta in ("soffice", "pdftoppm"):
        if shutil.which(ferramenta) is None:
            sys.exit(f"Preciso do '{ferramenta}' instalado só pra gerar as imagens (veja o cabeçalho deste arquivo).")
    if not os.path.isdir(SAIDA):
        sys.exit("Rode antes: python demo/executar_demo.py")
    os.makedirs(IMAGENS, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        recalculados = {}
        for nome, arquivo, aba, regra in VITRINES:
            if arquivo not in recalculados:
                recalculados[arquivo] = recalcular(os.path.join(SAIDA, arquivo), tmp)
            vitrine = os.path.join(tmp, f"{nome}.xlsx")
            montar_vitrine(recalculados[arquivo], aba, regra, vitrine)
            xlsx_para_png(vitrine, os.path.join(IMAGENS, f"{nome}.png"), tmp)
            print(f"  {nome}.png  ({aba})")
    imagem_do_log(os.path.join(DEMO, "logs", "reconciliacao_2026-09.txt"), os.path.join(IMAGENS, "07_log_reconciliacao.png"))
    print("  07_log_reconciliacao.png  (log da execução)")
    print(f"\nImagens em: {IMAGENS}")


if __name__ == "__main__":
    main()
