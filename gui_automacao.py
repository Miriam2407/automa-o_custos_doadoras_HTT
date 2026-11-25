"""
Interface gráfica da automação de doadoras
=============================================

Tela visual para rodar reconciliacao.py e conferencia_doadoras.py sem
precisar usar linha de comando. Abra pelo executar.bat.

Usa só bibliotecas que já vêm com o Python (tkinter) - não precisa
instalar nada além do que já está no requirements.txt.
"""

import os
import sys
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# garante que os módulos do projeto são encontrados mesmo se o .bat abrir
# a partir de outra pasta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import reconciliacao
import conferencia_doadoras


# ---------------------------------------------------------------------------
# Utilitário: roda uma função main(argv) numa thread separada e manda a
# saída (print) para a caixa de texto da interface, sem travar a tela
# ---------------------------------------------------------------------------

class RedirecionadorSaida:
    """Substitui sys.stdout temporariamente para mandar os prints para a fila."""
    def __init__(self, fila):
        self.fila = fila

    def write(self, texto):
        if texto:
            self.fila.put(texto)

    def flush(self):
        pass


def rodar_em_thread(func_main, argv, fila, callback_fim):
    def alvo():
        stdout_original = sys.stdout
        sys.stdout = RedirecionadorSaida(fila)
        sucesso = True
        try:
            sys.argv = argv
            func_main()
        except SystemExit:
            pass  # argparse chama sys.exit em erro de argumento - já foi logado
        except Exception:
            fila.put("\n--- ERRO ---\n")
            fila.put(traceback.format_exc())
            sucesso = False
        finally:
            sys.stdout = stdout_original
            # NUNCA mexer em widgets do Tkinter direto da thread de fundo -
            # só sinaliza o fim pela fila; quem atualiza a tela é o loop
            # principal (_poll_fila), que roda na thread certa
            fila.put(("__FIM__", sucesso, callback_fim))

    threading.Thread(target=alvo, daemon=True).start()


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------

class LinhaArquivo(ttk.Frame):
    """Uma linha com rótulo + campo de texto + botão de procurar."""
    def __init__(self, master, rotulo, modo="arquivo", filetypes=(("Excel", "*.xlsx"),)):
        super().__init__(master)
        self.modo = modo
        self.filetypes = filetypes
        ttk.Label(self, text=rotulo, width=26, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.var = tk.StringVar()
        ttk.Entry(self, textvariable=self.var, width=55).grid(row=0, column=1, sticky="we")
        ttk.Button(self, text="Procurar...", command=self.procurar).grid(row=0, column=2, padx=(6, 0))
        self.columnconfigure(1, weight=1)

    def procurar(self):
        if self.modo == "arquivo":
            caminho = filedialog.askopenfilename(filetypes=self.filetypes)
        else:
            caminho = filedialog.askdirectory()
        if caminho:
            self.var.set(caminho)

    def get(self):
        return self.var.get().strip()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação de Reconciliação de Custos de Doadoras")
        self.geometry("880x700")
        self.minsize(760, 600)

        self.fila_log = queue.Queue()
        self._montar_layout()
        self.after(100, self._poll_fila)

    # -- layout ----------------------------------------------------------

    def _montar_layout(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="x", padx=10, pady=10)

        self._montar_aba_reconciliacao(notebook)
        self._montar_aba_conferencia(notebook)

        # log compartilhado
        frame_log = ttk.LabelFrame(self, text="Saída da execução")
        frame_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.texto_log = tk.Text(frame_log, wrap="word", state="disabled", bg="#111", fg="#0f0", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frame_log, command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scrollbar.set)
        self.texto_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame_botoes_log = ttk.Frame(self)
        frame_botoes_log.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(frame_botoes_log, text="Limpar log", command=self._limpar_log).pack(side="left")
        ttk.Button(frame_botoes_log, text="Abrir pasta de saída (saida/)", command=self._abrir_pasta_saida).pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x", side="bottom")

    def _montar_aba_reconciliacao(self, notebook):
        aba = ttk.Frame(notebook, padding=12)
        notebook.add(aba, text="Reconciliação Mensal")

        self.rec_mes = self._linha_texto(aba, "Mês (AAAA-MM):", 0)
        self.rec_lista = LinhaArquivo(aba, "Lista de doadoras do mês:")
        self.rec_lista.grid(row=1, column=0, columnspan=2, sticky="we", pady=4)
        self.rec_custos = LinhaArquivo(aba, "Pasta com os arquivos de custo:", modo="pasta")
        self.rec_custos.grid(row=2, column=0, columnspan=2, sticky="we", pady=4)
        self.rec_controle_doadoras = LinhaArquivo(aba, "Controle de Doadoras (só 1ª vez):")
        self.rec_controle_doadoras.grid(row=3, column=0, columnspan=2, sticky="we", pady=4)
        self.rec_saldo_inicial = LinhaArquivo(aba, "Saldo Inicial (só 1ª vez):")
        self.rec_saldo_inicial.grid(row=4, column=0, columnspan=2, sticky="we", pady=4)
        self.rec_saldo_adicional = LinhaArquivo(aba, "Saldo Adicional (itens esquecidos):")
        self.rec_saldo_adicional.grid(row=5, column=0, columnspan=2, sticky="we", pady=4)
        self.rec_conf_anterior = LinhaArquivo(aba, "Conferência do mês anterior (opcional):")
        self.rec_conf_anterior.grid(row=6, column=0, columnspan=2, sticky="we", pady=4)

        self.rec_medicamentos = tk.BooleanVar()
        ttk.Checkbutton(aba, text="Medicamentos já liberados pelo fechamento contábil este mês",
                         variable=self.rec_medicamentos).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 2))

        self.rec_modo_teste = tk.BooleanVar()
        ttk.Checkbutton(aba, text="Modo teste (não altera o histórico real)",
                         variable=self.rec_modo_teste).grid(row=8, column=0, columnspan=2, sticky="w")

        ttk.Button(aba, text="▶  Rodar Reconciliação Mensal", command=self._rodar_reconciliacao
                   ).grid(row=9, column=0, columnspan=2, sticky="we", pady=(16, 0))

        aba.columnconfigure(0, weight=1)
        aba.columnconfigure(1, weight=1)

    def _montar_aba_conferencia(self, notebook):
        aba = ttk.Frame(notebook, padding=12)
        notebook.add(aba, text="Conferência Doadoras")

        ttk.Label(aba, text="O SALDO pendente vem direto do relatório da Reconciliação do mês (nada é recalculado) - "
                            "rode a Reconciliação antes, no mesmo modo (teste ou normal).",
                  wraplength=780, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self.conf_mes = self._linha_texto(aba, "Mês (AAAA-MM):", 1)
        self.conf_relatorio = LinhaArquivo(aba, "Relatório da Reconciliação (opcional):")
        self.conf_relatorio.grid(row=2, column=0, columnspan=2, sticky="we", pady=4)
        self.conf_totvs = LinhaArquivo(aba, "Extrato TOTVS (opcional):")
        self.conf_totvs.grid(row=3, column=0, columnspan=2, sticky="we", pady=4)
        self.conf_controle_estoque = LinhaArquivo(aba, "Controle Estoque Único (opcional):")
        self.conf_controle_estoque.grid(row=4, column=0, columnspan=2, sticky="we", pady=4)
        self.conf_anterior = LinhaArquivo(aba, "Conferência do mês anterior (opcional):")
        self.conf_anterior.grid(row=5, column=0, columnspan=2, sticky="we", pady=4)

        self.conf_modo_teste = tk.BooleanVar()
        ttk.Checkbutton(aba, text="Modo teste (usar o mesmo modo da Reconciliação)",
                         variable=self.conf_modo_teste).grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        ttk.Button(aba, text="▶  Rodar Conferência Doadoras", command=self._rodar_conferencia
                   ).grid(row=7, column=0, columnspan=2, sticky="we", pady=(16, 0))

        aba.columnconfigure(0, weight=1)
        aba.columnconfigure(1, weight=1)

    def _linha_texto(self, master, rotulo, row):
        ttk.Label(master, text=rotulo).grid(row=row, column=0, sticky="w")
        var = tk.StringVar()
        ttk.Entry(master, textvariable=var, width=20).grid(row=row, column=1, sticky="w")
        return var

    # -- ações -------------------------------------------------------------

    def _rodar_reconciliacao(self):
        mes = self.rec_mes.get().strip()
        lista = self.rec_lista.get()
        custos = self.rec_custos.get()
        if not mes or not lista or not custos:
            messagebox.showwarning("Campos obrigatórios", "Preencha Mês, Lista de doadoras e Pasta de custos.")
            return

        argv = ["reconciliacao.py", "--mes", mes, "--lista-doadoras", lista, "--custos", custos,
                "--medicamentos-liberado", "sim" if self.rec_medicamentos.get() else "nao"]
        if self.rec_controle_doadoras.get():
            argv += ["--controle-doadoras", self.rec_controle_doadoras.get()]
        if self.rec_saldo_inicial.get():
            argv += ["--saldo-inicial", self.rec_saldo_inicial.get()]
        if self.rec_saldo_adicional.get():
            argv += ["--saldo-adicional", self.rec_saldo_adicional.get()]
        if self.rec_conf_anterior.get():
            argv += ["--conferencia-anterior", self.rec_conf_anterior.get()]
        if self.rec_modo_teste.get():
            argv += ["--modo-teste"]

        self._limpar_log()
        self.status_var.set("Rodando Reconciliação Mensal...")
        rodar_em_thread(reconciliacao.main, argv, self.fila_log, self._ao_terminar)

    def _rodar_conferencia(self):
        mes = self.conf_mes.get().strip()
        if not mes:
            messagebox.showwarning("Campo obrigatório", "Preencha o Mês (AAAA-MM).")
            return

        argv = ["conferencia_doadoras.py", "--mes", mes]
        if self.conf_relatorio.get():
            argv += ["--relatorio-reconciliacao", self.conf_relatorio.get()]
        if self.conf_totvs.get():
            argv += ["--totvs", self.conf_totvs.get()]
        if self.conf_controle_estoque.get():
            argv += ["--controle-estoque", self.conf_controle_estoque.get()]
        if self.conf_anterior.get():
            argv += ["--conferencia-anterior", self.conf_anterior.get()]
        if self.conf_modo_teste.get():
            argv += ["--modo-teste"]

        self._limpar_log()
        self.status_var.set("Rodando Conferência Doadoras...")
        rodar_em_thread(conferencia_doadoras.main, argv, self.fila_log, self._ao_terminar)

    def _ao_terminar(self, sucesso):
        texto = "Concluído com sucesso." if sucesso else "Terminou com erro - veja o log."
        self.status_var.set(texto)
        if sucesso:
            messagebox.showinfo("Pronto", "Execução concluída! Confira a pasta 'saida/'.")

    def _limpar_log(self):
        self.texto_log.configure(state="normal")
        self.texto_log.delete("1.0", "end")
        self.texto_log.configure(state="disabled")

    def _abrir_pasta_saida(self):
        pasta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saida")
        os.makedirs(pasta, exist_ok=True)
        os.startfile(pasta)  # Windows

    def _poll_fila(self):
        try:
            while True:
                item = self.fila_log.get_nowait()
                if item is None:
                    continue
                if isinstance(item, tuple) and item[0] == "__FIM__":
                    _, sucesso, callback_fim = item
                    callback_fim(sucesso)
                    continue
                self.texto_log.configure(state="normal")
                self.texto_log.insert("end", item)
                self.texto_log.see("end")
                self.texto_log.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_fila)


if __name__ == "__main__":
    App().mainloop()
