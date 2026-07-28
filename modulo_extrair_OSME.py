import os
import re
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk

import pythoncom
import win32com.client

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


SCRIPT_OSME = os.path.join(obter_diretorio(), "Sciprt OSME2.vbs")


class AbaExtrairOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("\U0001f4c4 Aba Extrator OSME pronta.")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)

        for widget in [self.txt_recursos, self.txt_log]:
            widget.configure(
                bg=c["input"],
                fg=c["texto"],
                insertbackground=c["texto"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=c["container"],
                highlightcolor=c["accent"],
            )

        self.btn_extrair.configure(
            bg=c["accent"],
            fg="white",
            activebackground=c["hover"],
            activeforeground="white",
        )
        self.btn_limpar_log.configure(
            bg=c["container"],
            fg=c["texto"],
            activebackground=c["hover"],
            activeforeground="white",
        )
        self.check_imprimir_excel.configure(
            bg=c["bg"],
            fg=c["texto"],
            selectcolor=c["input"],
            activebackground=c["bg"],
            activeforeground=c["texto"],
        )

    def iniciar_extracao(self):
        if self.executando:
            messagebox.showwarning("Aviso", "A extracao OSME ja esta em execucao.")
            return

        recursos = [linha.strip() for linha in self.txt_recursos.get("1.0", tk.END).splitlines() if linha.strip()]
        if not recursos:
            messagebox.showwarning("Aviso", "Insira ao menos uma equipe/recurso.")
            return

        if not os.path.exists(SCRIPT_OSME):
            messagebox.showerror("Erro", f"Script nao encontrado:\n{SCRIPT_OSME}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para realizar a extracao OSME.\nAbra o SAP e tente novamente.")
            return

        self.executando = True
        self.btn_extrair.configure(state="disabled")
        self._status("Extracao OSME em execucao...")
        self._log("\n\U0001f4c4 Iniciando extracao OSME...")

        thread = threading.Thread(target=self._executar_recursos, args=(recursos,), daemon=True)
        thread.start()

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("\U0001f9f9 Log limpo.")

    def _executar_recursos(self, recursos):
        try:
            for recurso in recursos:
                self._log_threadsafe(f"\n\u23f3 Consultando OSME para: {recurso}...")
                script_temp = self._criar_script_temporario(recurso)
                self._log_threadsafe("\U0001f50d Executando extracao OSME no SAP...")

                resultado = subprocess.run(
                    ["cscript.exe", "//nologo", script_temp],
                    capture_output=True,
                    text=True,
                    timeout=240,
                )

                saida = (resultado.stdout or "").strip()
                if saida:
                    for linha in saida.splitlines():
                        self._log_threadsafe(linha.strip())

                if resultado.returncode == 0:
                    if self.imprimir_excel_var.get():
                        self._log_threadsafe("\U0001f5a8 Enviando planilha do Excel para impressao...")
                        self._imprimir_excel_aberto()
                    self._log_threadsafe(f"\u2705 Extracao OSME finalizada para: {recurso}")
                else:
                    erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                    self._log_threadsafe(f"\u26a0 Falha na extracao OSME de {recurso}: {erro}")

                try:
                    os.remove(script_temp)
                except OSError:
                    pass

            self._status_threadsafe("Extracao OSME finalizada.")
            self._log_threadsafe("\n\u2705 Processo OSME concluido.")
        except subprocess.TimeoutExpired:
            self._log_threadsafe("\u26a0 Tempo limite atingido ao executar o script OSME.")
            self._status_threadsafe("Extracao OSME interrompida por tempo limite.")
        except Exception as e:
            self._log_threadsafe(f"\u26a0 Erro inesperado: {e}")
            self._status_threadsafe("Erro na extracao OSME.")
        finally:
            self.parent_frame.after(0, self._finalizar_execucao)

    def _criar_script_temporario(self, recurso):
        with open(SCRIPT_OSME, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        conteudo = self._inserir_reset_sap(conteudo)
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.text\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo,
            count=1,
            flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.caretPosition\s*=\s*)\d+',
            lambda match: f"{match.group(1)}{len(recurso)}",
            conteudo,
            count=1,
            flags=re.IGNORECASE,
        )

        fd, caminho = tempfile.mkstemp(prefix="osme_", suffix=".vbs", text=True)
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

    def _imprimir_excel_aberto(self):
        pythoncom.CoInitialize()

        excel = self._aguardar_excel()
        if not excel:
            self._log_threadsafe("\u26a0 Nao encontrei Excel via COM. Tentando imprimir pela janela ativa...")
            if self._imprimir_excel_por_atalho():
                self._log_threadsafe("\u2705 Comando de impressao enviado pela janela do Excel.")
            else:
                self._log_threadsafe("\u26a0 Nao consegui ativar a janela do Excel para imprimir.")
            return

        workbook = self._aguardar_workbook(excel)
        if not workbook:
            self._log_threadsafe("\u26a0 Excel encontrado, mas nenhuma pasta de trabalho ativa ficou disponivel.")
            return

        try:
            excel.DisplayAlerts = False
            workbook.ActiveSheet.PageSetup.Orientation = 2
            workbook.ActiveSheet.PrintOut()
            self._log_threadsafe("\u2705 Impressao em paisagem enviada para a impressora padrao.")
        except Exception as e:
            self._log_threadsafe(f"\u26a0 Excel encontrado, mas a impressao via COM falhou: {e}")
            if self._imprimir_excel_por_atalho():
                self._log_threadsafe("\u2705 Comando de impressao enviado pela janela do Excel.")

    def _aguardar_excel(self, tentativas=20, intervalo=0.5):
        for _ in range(tentativas):
            try:
                return win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                time.sleep(intervalo)
        return None

    def _aguardar_workbook(self, excel, tentativas=20, intervalo=0.5):
        for _ in range(tentativas):
            try:
                if excel.Workbooks.Count > 0:
                    return excel.ActiveWorkbook or excel.Workbooks(excel.Workbooks.Count)
            except Exception:
                pass
            time.sleep(intervalo)
        return None

    def _imprimir_excel_por_atalho(self, tentativas=20, intervalo=0.5):
        shell = win32com.client.Dispatch("WScript.Shell")
        titulos = ["Microsoft Excel", "Excel"]

        for _ in range(tentativas):
            for titulo in titulos:
                try:
                    if shell.AppActivate(titulo):
                        time.sleep(0.5)
                        shell.SendKeys("%p")
                        time.sleep(0.3)
                        shell.SendKeys("o")
                        time.sleep(0.3)
                        shell.SendKeys("l")
                        time.sleep(0.5)
                        shell.SendKeys("^p")
                        time.sleep(1.2)
                        shell.SendKeys("{ENTER}")
                        return True
                except Exception:
                    pass
            time.sleep(intervalo)

        return False

    def _construir_interface(self):
        container = ttk.Frame(self.parent_frame)
        container.pack(fill="both", expand=True, padx=40, pady=24)

        ttk.Label(container, text="\U0001f4c4 Insira os recursos/equipes para extrair OSME (um por linha):").pack(anchor="w")

        frame_recursos = ttk.Frame(container)
        frame_recursos.pack(fill="x", pady=(6, 14))

        self.txt_recursos = tk.Text(frame_recursos, height=8, font=("Consolas", 10), wrap="none")
        scroll_recursos = ttk.Scrollbar(frame_recursos, orient="vertical", command=self.txt_recursos.yview)
        self.txt_recursos.configure(yscrollcommand=scroll_recursos.set)
        self.txt_recursos.pack(side="left", fill="x", expand=True)
        scroll_recursos.pack(side="right", fill="y")

        frame_botoes = ttk.Frame(container)
        frame_botoes.pack(fill="x", pady=(0, 18))

        self.btn_extrair = tk.Button(
            frame_botoes,
            text="\U0001f4c4 Iniciar extracao OSME",
            command=self.iniciar_extracao,
            font=("Segoe UI", 11, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=10,
        )
        self.btn_extrair.pack(side="left", fill="x", expand=True)

        self.imprimir_excel_var = tk.BooleanVar(value=False)
        self.check_imprimir_excel = tk.Checkbutton(
            frame_botoes,
            text="\U0001f5a8 Imprimir Excel ao finalizar",
            variable=self.imprimir_excel_var,
            font=("Segoe UI", 9, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
        )
        self.check_imprimir_excel.pack(side="left", padx=(10, 0), ipady=8)

        self.btn_limpar_log = tk.Button(
            frame_botoes,
            text="\U0001f9f9 Limpar log",
            command=self.limpar_log,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            padx=18,
            pady=10,
        )
        self.btn_limpar_log.pack(side="left", padx=(10, 0))

        ttk.Label(container, text="\U0001f4cb Log de execucao:").pack(anchor="w")

        frame_log = ttk.Frame(container)
        frame_log.pack(fill="both", expand=True, pady=(6, 0))

        self.txt_log = tk.Text(frame_log, height=12, font=("Segoe UI", 10), wrap="word", state="disabled")
        scroll_log = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll_log.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        scroll_log.pack(side="right", fill="y")

        self.btn_extrair.bind("<Enter>", lambda e: self._botao_hover(self.btn_extrair, principal=True))
        self.btn_extrair.bind("<Leave>", lambda e: self._botao_normal(self.btn_extrair, principal=True))
        self.btn_limpar_log.bind("<Enter>", lambda e: self._botao_hover(self.btn_limpar_log))
        self.btn_limpar_log.bind("<Leave>", lambda e: self._botao_normal(self.btn_limpar_log))

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_extrair.configure(state="normal")

    def _log(self, texto):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, texto + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._log(texto))

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


def construir_aba_extrair_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrairOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
