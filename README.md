# Automação de reconciliação de custos de doadoras

Automação em Python que transforma a conciliação **mensal e manual** de custos (transporte por app,
táxi, laboratório, ajuda de custo, medicamentos…) em um processo repetível: liga cada custo à pessoa
certa, gera o lançamento contábil, acompanha o que ficou pendente e **confere tudo contra o extrato do ERP**.

> 🦸‍♀️ **Todos os dados deste repositório são fictícios.** As "doadoras" da demonstração são mulheres
> importantes da história (Marie Curie, Frida Kahlo, Ada Lovelace…) e super-heroínas (Diana Prince,
> Carol Danvers, Natasha Romanoff…). Nenhum nome, valor ou documento é real.

---

## O problema

Todo mês chegam planilhas de várias fontes, **cada uma com um layout**, com o nome da pessoa digitado de
jeitos diferentes (`Marie Currie` × `Marie Curie`). Era preciso ligar cada custo à pessoa certa, lançar na
contabilidade, acompanhar o que não tinha dono e conferir o saldo com o ERP — tudo à mão, todo mês.

## O que a automação faz

| | |
|---|---|
| 📥 **Lê fontes diferentes** | Uber, táxi, laboratório, ajuda de custo, medicamentos, testes genéticos… cada uma com seu layout (aba, linha do cabeçalho e colunas ficam em `config.py`). Os nomes de coluna e de aba são achados **sem diferenciar maiúscula, acento ou espaço nas pontas** (`Tipo` = `TIPO`), e quando falta algo o erro diz o que foi procurado e o que o arquivo tem. |
| 📋 **Lista de doadoras: layout antigo e novo** | Aceita a lista antiga (um mês por arquivo) e a nova (coluna **Mini lote** e meses **acumulados** na mesma aba). Só entram as linhas do mês processado, porque esse número é o divisor do rateio. Avisa quando um **PIN se repete** para pessoas diferentes e quando há linha **sem nome**. |
| 🔎 **Liga o custo à pessoa** | Matching aproximado de nomes (`rapidfuzz`) com limite de similaridade configurável. A coluna `score_match` fica no relatório para revisar o que ficou perto do limite. |
| ⏳ **SALDO → AJUSTE** | Custo sem dono vira *SALDO*; depois de 3 meses vira *AJUSTE*, contando a partir da **data original** (não da data em que foi lido). |
| ➗ **Rateia custos sem nome** | Insumos, MGR e médicos vêm como total; são distribuídos entre as doadoras do mês. Médicos segue a **regra do mês defasado** (o valor do mês fica em saldo; o do mês anterior é consumido). |
| 🚧 **Categoria só-saldo** | Uma fonte pode ser marcada `somente_saldo` no `config.py`: ela **nunca é vinculada a ninguém** (nem por PIN, nem por nome idêntico) e **nunca vira ajuste**. Usei para *testes genéticos*, enquanto as movimentações não estão definidas. |
| 🧾 **Lançamento contábil** | Um lançamento por pessoa e categoria, já consolidado (`Trf. Est. 3 DIANA_TAXI - SETEMBRO/2026`). |
| 📊 **Posição consolidada** | Histórico mês a mês de quanto cada pessoa custou, com uma linha **AJUSTES DIVERSOS** somando o que caiu em ajuste. |
| ✅ **Confere com o ERP** | O *Resumo* compara o saldo pendente com o extrato do ERP usando **fórmulas `SUMIF`** (qualquer edição manual entra sozinha). |
| 🧠 **Lembra do mês anterior** | A aba do ERP do mês novo **parte da aba do mês anterior**, mantendo os ajustes feitos à mão, e só acrescenta os lançamentos novos (sem duplicar). |
| 🖥️ **Interface gráfica** | Quem não programa roda com dois cliques (`executar.bat`). |

## Como funciona

```
 planilhas de custo (Uber · táxi · laboratório · ajuda de custo · medicamentos)  +  lista de doadoras do mês
                                           │
                                           ▼
                              ┌─────────────────────────┐
                              │     reconciliacao.py    │  matching aproximado → SALDO → (3 meses) → AJUSTE
                              └─────────────────────────┘
                                           │
                    relatorio_reconciliacao_AAAA-MM.xlsx   (por categoria · SALDO · AJUSTE ·
                                           │                Lançamento Contábil · Posição Consolidada)
                                           ▼
   extrato do ERP ──►  ┌──────────────────────────┐  ◄── conferência do mês anterior
                       │  conferencia_doadoras.py │      (ajustes manuais mantidos)
                       └──────────────────────────┘
                                           │
                    conferencia_doadoras_AAAA-MM.xlsx   (SALDO pendente · TOTVS · Resumo com SUMIF)
```

## Veja funcionando (sem instalar nada além do Python)

```bash
pip install -r requirements.txt
python demo/executar_demo.py
```

O script gera os dados fictícios, roda **dois meses em sequência** (agosto e setembro/2026) e grava os
relatórios em `demo/saida/`. Os resultados já estão versionados no repositório; as imagens abaixo vêm deles.

Os dois meses de propósito têm **layouts diferentes**, como acontece na vida real: agosto usa a lista antiga; setembro usa a lista
nova (com *Mini lote* e agosto acumulado na mesma aba), o Uber com colunas `TIPO`/`Valor` e o extrato do ERP com a aba sem acento.

### Setembro/2026 em imagens

**Log da execução** — leitura, matching e classificação (9 doadoras; 152 lançamentos: 107 identificados, 38 em saldo, 7 em ajuste):

![Log da execução](docs/imagens/07_log_reconciliacao.png)

**SALDO pendente** — quem não casou com nenhuma doadora (Gamora, Nebula…) e, no fim, as linhas de AJUSTE
mantidas do mês anterior para retirar manualmente do saldo:

![SALDO pendente](docs/imagens/01_saldo_pendente.png)

**Testes genéticos** — o arquivo do fornecedor traz PINs no lugar de nomes e é acumulado (setembro repete agosto). Os PINs são de
doadoras que **existem na lista**, e a última linha tem o nome **exato** de uma doadora: mesmo assim nada é vinculado e tudo fica em
SALDO, como pedido:

![Testes genéticos](docs/imagens/08_testes_geneticos.png)

**AJUSTE do mês** — itens de junho que completaram 3 meses sem dono:

![AJUSTE do mês](docs/imagens/02_ajuste_do_mes.png)

**Lançamento contábil** — um lançamento por doadora e categoria. Repare nos médicos: o valor consumido em
setembro é o de agosto (data 31/08), pela regra do mês defasado:

![Lançamento contábil](docs/imagens/03_lancamento_contabil.png)

**Posição consolidada** — histórico por mês, com a linha **AJUSTES DIVERSOS** (em negrito, no fim):

![Posição consolidada](docs/imagens/04_posicao_consolidada.png)

**Resumo da conferência** — saldo pendente × ERP, com `SUMIF`. A linha *Testes Genéticos* aparece como pendência (R$ 31.350,00 de
diferença) porque ainda não houve movimentação no ERP; as demais zeram depois dos ajustes:

![Resumo](docs/imagens/05_resumo_conferencia.png)

**Aba do ERP com histórico** — saldos iniciais, ajustes do mês anterior preservados, lançamentos novos e a
coluna de saldo do extrato:

![ERP com histórico](docs/imagens/06_totvs_com_historico.png)

> Os "ajustes manuais" que zeram o Resumo são **simulados** por `demo/executar_demo.py`: na rotina real é a
> analista quem os faz no Excel. O ponto é que a automação os **respeita e carrega** de um mês para o outro.

## Usando com os seus dados

1. Coloque as planilhas do mês numa pasta e rode a interface (`executar.bat`) ou os scripts:
   ```bash
   python reconciliacao.py --mes 2026-09 --lista-doadoras lista.xlsx --custos pasta_do_mes
   python conferencia_doadoras.py --mes 2026-09 --totvs extrato_erp.xlsx
   ```
2. O que cada fonte tem de **nome de arquivo, aba, linha de cabeçalho e colunas** está em [`config.py`](config.py) —
   é o único lugar que precisa mudar para adaptar a outro layout.
3. A pasta `controle/` guarda a memória entre os meses (cadastro de doadoras, saldos pendentes, histórico). **Não apague.**

Parâmetros úteis: `--modo-teste` (usa `controle_teste/` e `saida_teste/`, sem tocar no histórico real),
`--saldo-inicial` (custos em aberto de antes da automação), `--saldo-adicional` (itens esquecidos de um mês
passado), `--conferencia-anterior` (se o arquivo do mês anterior não estiver em `saida/`).

## Estrutura

```
reconciliacao.py          # lê custos, casa com doadoras, SALDO/AJUSTE, relatório de reconciliação
conferencia_doadoras.py   # confere com o ERP: SALDO pendente, aba do ERP, Resumo (SUMIF), estoque
config.py                 # TODO o que é específico de layout/regra: arquivos, colunas, limites, rótulos
gui_automacao.py          # interface gráfica (tkinter)
executar.bat / instalar.bat
demo/                     # dados fictícios + execução de 2 meses + gerador das imagens
docs/imagens/             # prints da demonstração
```

## Decisões de projeto

- **Falhar alto, não em silêncio.** Arquivo não encontrado, coluna ou aba com nome diferente, relatório de outro mês:
  tudo vira aviso ou erro claro no log, nunca um valor zerado sem explicação.
- **Rodar de novo não duplica.** Reprocessar o mesmo mês não duplica o histórico de custos nem o extrato do ERP.
- **A automação respeita o trabalho manual.** O que a analista ajusta à mão no Excel é carregado para o mês
  seguinte, e o Resumo usa fórmulas em vez de números calculados em Python.
- **Um "match" errado é pior que nenhum.** Por isso o limite de similaridade e a coluna `score_match`.
- **A lista é a fonte da verdade do mês.** Reprocessar um mês substitui as linhas dele no histórico (não duplica) e corrige PIN ou nome
  alterado na lista. O PIN **não** é usado para eliminar linhas: repetido por engano, ele chegou a fazer uma doadora sumir do histórico.
- **Avisar em vez de adivinhar.** PIN repetido, linha sem nome, lista de outro mês: a automação mantém o que pode e diz no log o que viu.

## Limitações

- Projeto nascido de uma rotina interna; os layouts de entrada são os de lá, adaptáveis via `config.py`.
- A interface e os `.bat` são para Windows. Os scripts em si rodam em qualquer sistema com Python.
- Testado com Python 3.12.

## Tecnologias

Python · pandas · openpyxl · rapidfuzz · tkinter

## Licença

[MIT](LICENSE)
