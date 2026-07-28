import os
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from modulo_Sap import verificar_sap_aberto


DEFAULT_PALETAS = {
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#f57c00",
        "hover": "#e65100",
        "input": "#1f3a60",
    }
}


def obter_diretorio():
    import sys
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


SCRIPT_IQ09 = os.path.join(obter_diretorio(), "SCRIPT DA IQ09.vbs")


class AbaExtrator:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando_iq09 = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)

        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            widget.configure(
                bg=c["input"],
                fg=c["texto"],
                insertbackground=c["texto"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=c["container"],
                highlightcolor=c["accent"],
            )

        botoes_principais = [self.btn_extrair, self.btn_iq09]
        botoes_secundarios = [self.btn_copiar_sem, self.btn_copiar_com, self.btn_limpar]

        for btn in botoes_principais:
            btn.configure(bg=c["accent"], fg="white", activebackground=c["hover"], activeforeground="white")

        for btn in botoes_secundarios:
            btn.configure(bg=c["container"], fg=c["texto"], activebackground=c["hover"], activeforeground="white")

    def extrair_seriais(self):
        texto = self.txt_entrada.get("1.0", tk.END)
        seriais = self._extrair_seriais(texto)

        if not seriais:
            messagebox.showwarning("Aviso", "Nenhum serial encontrado para extrair.")
            return

        sem_ponto_virgula = "\n".join(seriais)
        com_ponto_virgula = "\n".join(f"{serial};" for serial in seriais)

        self._set_texto(self.txt_sem_ponto_virgula, sem_ponto_virgula)
        self._set_texto(self.txt_com_ponto_virgula, com_ponto_virgula)
        self._status(f"{len(seriais)} seriais extraidos.")

    def copiar_sem_ponto_virgula(self):
        self._copiar_texto(self.txt_sem_ponto_virgula, "Lista sem ponto e virgula copiada.")

    def copiar_com_ponto_virgula(self):
        self._copiar_texto(self.txt_com_ponto_virgula, "Lista com ponto e virgula copiada.")

    def abrir_na_iq09(self):
        if self.executando_iq09:
            messagebox.showwarning("Aviso", "A IQ09 ja esta em execucao.")
            return

        seriais = self._obter_lista_sem_ponto_virgula()
        if not seriais:
            self.extrair_seriais()
            seriais = self._obter_lista_sem_ponto_virgula()

        if not seriais:
            return

        if not os.path.exists(SCRIPT_IQ09):
            messagebox.showerror("Erro", f"Script nao encontrado:\n{SCRIPT_IQ09}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para abrir a IQ09.\nAbra o SAP e tente novamente.")
            return

        # A IQ09 precisa da tabela limpa, sem ponto e virgula.
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append("\r\n".join(seriais))
        self.parent_frame.update()

        self.executando_iq09 = True
        self.btn_iq09.configure(state="disabled")
        self._status(f"Abrindo IQ09 com {len(seriais)} seriais sem ponto e virgula...")

        thread = threading.Thread(target=self._executar_iq09, daemon=True)
        thread.start()

    def limpar(self):
        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            self._set_texto(widget, "")
        self._status("Extrator limpo.")

    def _executar_iq09(self):
        script_temp = None
        try:
            script_temp = self._criar_script_iq09_temporario()
            resultado = subprocess.run(
                ["cscript.exe", "//nologo", script_temp],
                capture_output=True,
                text=True,
                timeout=120,
            )

            if resultado.returncode == 0:
                self._status_threadsafe("IQ09 aberta com a lista sem ponto e virgula.")
            else:
                erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", erro))
                self._status_threadsafe("Falha ao abrir IQ09.")
        except subprocess.TimeoutExpired:
            self.parent_frame.after(0, lambda: messagebox.showwarning("Aviso", "Tempo limite atingido ao abrir a IQ09."))
            self._status_threadsafe("IQ09 excedeu o tempo limite.")
        except Exception as e:
            self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", str(e)))
            self._status_threadsafe("Erro ao executar script da IQ09.")
        finally:
            if script_temp:
                try:
                    os.remove(script_temp)
                except OSError:
                    pass
            self.parent_frame.after(0, self._finalizar_iq09)

    def _criar_script_iq09_temporario(self):
        with open(SCRIPT_IQ09, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        conteudo = self._inserir_reset_sap(conteudo)

        fd, caminho = tempfile.mkstemp(prefix="iq09_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="mbcs") as f:
            f.write(conteudo)
        return caminho

    def _inserir_reset_sap(self, conteudo):
        if "Reset automatico inserido pelo Conecta Hub" in conteudo:
            return conteudo

        ponto_ancora = 'session.findById("wnd[0]").maximize'
        bloco_reset = (
            "' Reset automatico inserido pelo Conecta Hub\n"
            'On Error Resume Next\n'
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nS000"\n'
            'session.findById("wnd[0]").sendVKey 0\n'
            'WScript.Sleep 500\n'
            'If session.Children.Count > 1 Then\n'
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press\n'
            'End If\n'
            'On Error GoTo 0\n'
        )

        if ponto_ancora in conteudo:
            return conteudo.replace(ponto_ancora, ponto_ancora + "\n" + bloco_reset, 1)

        return bloco_reset + conteudo

    def _extrair_seriais(self, texto):
        seriais = []

        for linha in texto.splitlines():
            linha = linha.strip()
            if not linha:
                continue

            match_intervalo = re.search(r"(\d+)\s*(?:a|A|-|ate|at[eé])\s*(\d+)", linha, flags=re.IGNORECASE)
            if match_intervalo:
                inicio_txt, fim_txt = match_intervalo.groups()
                inicio = int(inicio_txt)
                fim = int(fim_txt)
                passo = 1 if fim >= inicio else -1
                largura = max(len(inicio_txt), len(fim_txt))
                preservar_zeros = inicio_txt.startswith("0") or fim_txt.startswith("0")

                for numero in range(inicio, fim + passo, passo):
                    seriais.append(str(numero).zfill(largura) if preservar_zeros else str(numero))
                continue

            numeros = re.findall(r"\d+", linha.replace(";", " "))
            seriais.extend(numeros)

        return self._remover_duplicados_mantendo_ordem(seriais)

    def _remover_duplicados_mantendo_ordem(self, seriais):
        vistos = set()
        resultado = []
        for serial in seriais:
            if serial not in vistos:
                vistos.add(serial)
                resultado.append(serial)
        return resultado

    def _obter_lista_sem_ponto_virgula(self):
        texto = self.txt_sem_ponto_virgula.get("1.0", tk.END)
        return [linha.strip().replace(";", "") for linha in texto.splitlines() if linha.strip()]

    def _copiar_texto(self, widget, mensagem):
        texto = widget.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nao ha conteudo para copiar.")
            return

        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        self.parent_frame.update()
        self._status(mensagem)

    def _construir_interface(self):
        container = ttk.Frame(self.parent_frame)
        container.pack(fill="both", expand=True, padx=40, pady=20)

        ttk.Label(container, text="\U0001f522 Cole os intervalos ou seriais SAP abaixo:").pack(anchor="w")

        frame_entrada = ttk.Frame(container)
        frame_entrada.pack(fill="both", expand=True, pady=(6, 12))

        self.txt_entrada = tk.Text(frame_entrada, height=8, font=("Consolas", 10), wrap="none")
        scroll_entrada_y = ttk.Scrollbar(frame_entrada, orient="vertical", command=self.txt_entrada.yview)
        self.txt_entrada.configure(yscrollcommand=scroll_entrada_y.set)
        self.txt_entrada.pack(side="left", fill="both", expand=True)
        scroll_entrada_y.pack(side="right", fill="y")

        self.btn_extrair = tk.Button(
            container,
            text="\U0001f50d EXTRAIR SERIAL",
            command=self.extrair_seriais,
            font=("Segoe UI", 11, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=10,
        )
        self.btn_extrair.pack(fill="x", pady=(0, 14))

        frame_resultados = ttk.Frame(container)
        frame_resultados.pack(fill="both", expand=True)
        frame_resultados.columnconfigure(0, weight=1)
        frame_resultados.columnconfigure(1, weight=1)
        frame_resultados.rowconfigure(1, weight=1)

        ttk.Label(frame_resultados, text='Tabela IQ09 sem ";"').grid(row=0, column=0, sticky="w")
        ttk.Label(frame_resultados, text='Formatado com ";"').grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.txt_sem_ponto_virgula = tk.Text(frame_resultados, height=10, font=("Consolas", 10), wrap="none")
        self.txt_sem_ponto_virgula.grid(row=1, column=0, sticky="nsew", pady=(6, 8))

        self.txt_com_ponto_virgula = tk.Text(frame_resultados, height=10, font=("Consolas", 10), wrap="none")
        self.txt_com_ponto_virgula.grid(row=1, column=1, sticky="nsew", padx=(12, 0), pady=(6, 8))

        frame_botoes_resultado = ttk.Frame(frame_resultados)
        frame_botoes_resultado.grid(row=2, column=0, columnspan=2, sticky="ew")
        frame_botoes_resultado.columnconfigure(0, weight=1)
        frame_botoes_resultado.columnconfigure(1, weight=1)

        self.btn_copiar_sem = tk.Button(
            frame_botoes_resultado,
            text='\U0001f4cb Copiar tabela sem ";"',
            command=self.copiar_sem_ponto_virgula,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=8,
        )
        self.btn_copiar_sem.grid(row=0, column=0, sticky="ew")

        self.btn_copiar_com = tk.Button(
            frame_botoes_resultado,
            text='\U0001f4cb Copiar com ";"',
            command=self.copiar_com_ponto_virgula,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=8,
        )
        self.btn_copiar_com.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        frame_acoes = ttk.Frame(container)
        frame_acoes.pack(fill="x", pady=(14, 0))

        self.btn_iq09 = tk.Button(
            frame_acoes,
            text='\U0001f5a5 ABRIR NA IQ09 usando a tabela sem ";"',
            command=self.abrir_na_iq09,
            font=("Segoe UI", 11, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=10,
        )
        self.btn_iq09.pack(side="left", fill="x", expand=True)

        self.btn_limpar = tk.Button(
            frame_acoes,
            text="\U0001f9f9 Limpar",
            command=self.limpar,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            padx=18,
            pady=10,
        )
        self.btn_limpar.pack(side="left", padx=(10, 0))

        for btn, principal in [
            (self.btn_extrair, True),
            (self.btn_iq09, True),
            (self.btn_copiar_sem, False),
            (self.btn_copiar_com, False),
            (self.btn_limpar, False),
        ]:
            btn.bind("<Enter>", lambda e, b=btn, p=principal: self._botao_hover(b, p))
            btn.bind("<Leave>", lambda e, b=btn, p=principal: self._botao_normal(b, p))

    def _set_texto(self, widget, texto):
        widget.delete("1.0", tk.END)
        widget.insert("1.0", texto)

    def _finalizar_iq09(self):
        self.executando_iq09 = False
        self.btn_iq09.configure(state="normal")

    def _status(self, texto):
        if self.label_status_global:
            self.label_status_global.configure(text=f" {texto}")

    def _status_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._status(texto))

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

    def _botao_hover(self, btn, principal=False):
        c = self._cores()
        btn.configure(bg=c["hover"], fg="white")

    def _botao_normal(self, btn, principal=False):
        c = self._cores()
        if principal:
            btn.configure(bg=c["accent"], fg="white")
        else:
            btn.configure(bg=c["container"], fg=c["texto"])


def construir_aba_extrator(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrator(parent_frame, label_status_global, paletas, obter_tema_atual)
