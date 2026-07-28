import json
import os
import subprocess
import time
import tkinter as tk
from tkinter import messagebox, ttk

import pythoncom
import win32com.client


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


def verificar_sap_aberto():
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if SapGuiAuto:
            return True
    except Exception:
        pass
    return False


ARQUIVO_LOGINS = os.path.join(obter_diretorio(), "logins_sap.json")
logins_salvos = {}


def carregar_logins():
    global logins_salvos
    LOGINS_PADRAO = {
        "Adriana": {"usuario": "BR0425314228", "senha": "Cnt*11092027"},
        "Junior": {"usuario": "BR0287830298", "senha": "Cnt*11092027"},
        "Pedro": {"usuario": "BR0498361748", "senha": "Cnt*11092027"},
    }
    if os.path.exists(ARQUIVO_LOGINS):
        try:
            with open(ARQUIVO_LOGINS, "r", encoding="utf-8") as f:
                logins_salvos = json.load(f)
        except Exception:
            logins_salvos = {}
        if not logins_salvos:
            logins_salvos = dict(LOGINS_PADRAO)
            salvar_logins()
    else:
        logins_salvos = dict(LOGINS_PADRAO)
        salvar_logins()


def salvar_logins():
    try:
        with open(ARQUIVO_LOGINS, "w", encoding="utf-8") as f:
            json.dump(logins_salvos, f, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Erro", f"Nao foi possivel salvar os logins:\n{str(e)}")


def logar_sap(usuario, senha, nome, label_status=None):
    try:
        pythoncom.CoInitialize()

        SapGuiAuto = None
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            pass

        if not SapGuiAuto:
            caminho_sap = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            if not os.path.exists(caminho_sap):
                messagebox.showerror("Erro", "Executavel do SAP nao encontrado no caminho padrao.")
                return
            subprocess.Popen(caminho_sap)

            tentativas = 0
            while tentativas < 10 and not SapGuiAuto:
                time.sleep(1)
                try:
                    pythoncom.CoInitialize()
                    SapGuiAuto = win32com.client.GetObject("SAPGUI")
                except Exception:
                    tentativas += 1

            if not SapGuiAuto:
                messagebox.showerror("Erro", "O SAP demorou para abrir. Tente novamente.")
                return

        app = SapGuiAuto.GetScriptingEngine

        try:
            connection = app.OpenConnection("Enel_SP_RP1_CCS_Produtivo [VPN Palo Alto]", True)
        except Exception:
            connection = app.Children(0)

        session = connection.Children(0)
        time.sleep(1)

        try:
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = ""
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = ""
        except Exception:
            pass

        try:
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha
            session.findById("wnd[0]").sendVKey(0)
        except Exception:
            try:
                session.findById("wnd[0]/usr/txtRSYST-BNAME[0,0]").Text = usuario
                session.findById("wnd[0]/usr/pwdRSYST-BCODE[0,0]").Text = senha
                session.findById("wnd[0]").sendVKey(0)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao preencher dados. O SAP esta na tela inicial?\n\nDetalhe: {str(e)}")
                return

        if label_status:
            label_status.config(text=f" Usuario SAP logado: {nome} | {usuario}")
        else:
            messagebox.showinfo("Sucesso", f"Usuario {nome} conectado no SAP!")

    except Exception as e:
        messagebox.showerror("Erro Inesperado", f"Nao foi possivel integrar com o SAP:\n{str(e)}")


def logout_sap(label_status=None):
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        connection = app.Children(0)
        session = connection.Children(0)

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
        session.findById("wnd[0]").sendVKey(0)

        if label_status:
            label_status.config(text=" SAP desconectado.")
    except Exception:
        messagebox.showinfo("Aviso", "O SAP ja esta fechado ou inativo.")


def obter_info_sap_atual():
    try:
        pythoncom.CoInitialize()
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        connection = app.Children(0)
        session = connection.Children(0)
        info = session.Info

        sistema = getattr(info, "SystemName", "") or "SAP"
        cliente = getattr(info, "Client", "") or "-"
        usuario = getattr(info, "User", "") or "-"
        transacao = getattr(info, "Transaction", "") or "-"

        return f"SAP: {sistema} | Cliente {cliente} | Usuario {usuario} | Tx {transacao}"
    except Exception:
        return "SAP: desconectado"


class AbaSap:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.botoes_usuario = []

        carregar_logins()
        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(style="TFrame")

        self.entry_busca.configure(
            bg=c["input"],
            fg=c["texto"],
            insertbackground=c["texto"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=c["container"],
            highlightcolor=c["accent"],
        )
        self.canvas.configure(bg=c["bg"], highlightbackground=c["container"])
        self.btn_gerenciar.configure(
            bg=c["container"],
            fg=c["texto"],
            activebackground=c["hover"],
            activeforeground="white",
        )
        self.btn_logout.configure(
            bg=c["accent"],
            fg="white",
            activebackground=c["hover"],
            activeforeground="white",
        )

        self.atualizar_botoes(self.entry_busca.get())

    def atualizar_botoes(self, busca=""):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.botoes_usuario = []
        c = self._cores()

        for nome in sorted(logins_salvos.keys()):
            if busca.lower() not in nome.lower():
                continue

            dados = logins_salvos[nome]
            btn = tk.Button(
                self.scrollable_frame,
                text=f"\U0001f464 Conectar como {nome}",
                font=("Segoe UI", 12, "bold"),
                bg=c["container"],
                fg=c["texto"],
                activebackground=c["hover"],
                activeforeground="white",
                relief="flat",
                cursor="hand2",
                pady=12,
                command=lambda n=nome, u=dados["usuario"], s=dados["senha"]: logar_sap(
                    u,
                    s,
                    n,
                    self.label_status_global,
                ),
            )
            btn.pack(fill="x", pady=4)
            btn.bind("<Enter>", lambda e, b=btn: self._botao_hover(b))
            btn.bind("<Leave>", lambda e, b=btn: self._botao_normal(b))
            self.botoes_usuario.append(btn)

    def abrir_gerenciador(self):
        c = self._cores()
        pop = tk.Toplevel()
        pop.title("Gerenciar Usuarios SAP")
        pop.geometry("380x450")
        pop.resizable(False, False)
        pop.grab_set()
        pop.configure(bg=c["bg"])

        tk.Label(
            pop,
            text="Usuarios cadastrados:",
            font=("Segoe UI", 10, "bold"),
            bg=c["bg"],
            fg=c["texto"],
        ).pack(anchor="w", padx=20, pady=(15, 5))

        listbox_pop = tk.Listbox(
            pop,
            height=6,
            font=("Segoe UI", 10),
            bd=0,
            relief="flat",
            bg=c["input"],
            fg=c["texto"],
            selectbackground=c["accent"],
            selectforeground="white",
            highlightthickness=1,
            highlightbackground=c["container"],
            highlightcolor=c["accent"],
        )
        listbox_pop.pack(fill="x", padx=20)

        frame_form = tk.Frame(pop, bg=c["bg"])
        frame_form.pack(fill="both", expand=True, padx=20, pady=10)

        entry_nome = self._campo_popup(frame_form, "Nome:", c)
        entry_user = self._campo_popup(frame_form, "Usuario SAP:", c)
        entry_senha = self._campo_popup(frame_form, "Senha SAP:", c, show="*")

        def atualizar_lista_pop():
            listbox_pop.delete(0, tk.END)
            for n in sorted(logins_salvos.keys()):
                listbox_pop.insert(tk.END, n)

        def selecionar_pop(event):
            if not listbox_pop.curselection():
                return
            nome_sel = listbox_pop.get(listbox_pop.curselection())
            dados = logins_salvos.get(nome_sel)
            if dados:
                entry_nome.delete(0, tk.END)
                entry_nome.insert(0, nome_sel)
                entry_user.delete(0, tk.END)
                entry_user.insert(0, dados["usuario"])
                entry_senha.delete(0, tk.END)
                entry_senha.insert(0, dados["senha"])

        def btn_salvar():
            n = entry_nome.get().strip()
            u = entry_user.get().strip()
            s = entry_senha.get().strip()
            if not n or not u or not s:
                messagebox.showwarning("Aviso", "Preencha Nome, Usuario e Senha.", parent=pop)
                return

            logins_salvos[n] = {"usuario": u, "senha": s}
            salvar_logins()
            atualizar_lista_pop()
            self.atualizar_botoes(self.entry_busca.get())
            messagebox.showinfo("Sucesso", f"Usuario {n} salvo com sucesso!", parent=pop)

        def btn_excluir():
            n = entry_nome.get().strip()
            if n in logins_salvos:
                del logins_salvos[n]
                salvar_logins()
                atualizar_lista_pop()
                self.atualizar_botoes(self.entry_busca.get())
                for entry in [entry_nome, entry_user, entry_senha]:
                    entry.delete(0, tk.END)
                messagebox.showinfo("Sucesso", "Usuario excluido.", parent=pop)

        listbox_pop.bind("<<ListboxSelect>>", selecionar_pop)
        atualizar_lista_pop()

        btn_salvar_ui = tk.Button(
            frame_form,
            text="Salvar / Adicionar",
            command=btn_salvar,
            font=("Segoe UI", 10, "bold"),
            bg=c["accent"],
            fg="white",
            activebackground=c["hover"],
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            pady=8,
        )
        btn_salvar_ui.pack(fill="x", pady=(12, 4))

        btn_excluir_ui = tk.Button(
            frame_form,
            text="Excluir selecionado",
            command=btn_excluir,
            font=("Segoe UI", 10, "bold"),
            bg=c["container"],
            fg=c["texto"],
            activebackground=c["hover"],
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            pady=8,
        )
        btn_excluir_ui.pack(fill="x", pady=4)

    def _construir_interface(self):
        self.frame_topo = ttk.Frame(self.parent_frame)
        self.frame_topo.pack(fill="x", padx=40, pady=(30, 10))

        ttk.Label(self.frame_topo, text="Buscar:").pack(side="left")

        self.entry_busca = tk.Entry(self.frame_topo, font=("Segoe UI", 11))
        self.entry_busca.pack(side="left", fill="x", expand=True, padx=10)

        self.btn_gerenciar = tk.Button(
            self.frame_topo,
            text="\u2699 Gerenciar Usuarios",
            command=self.abrir_gerenciador,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            relief="flat",
            pady=4,
            padx=10,
        )
        self.btn_gerenciar.pack(side="right")

        self.frame_lista = ttk.Frame(self.parent_frame)
        self.frame_lista.pack(fill="both", expand=True, padx=40, pady=10)

        self.canvas = tk.Canvas(self.frame_lista, bd=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.frame_lista, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_janela = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_janela, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.entry_busca.bind("<KeyRelease>", lambda e: self.atualizar_botoes(self.entry_busca.get()))

        self.btn_logout = tk.Button(
            self.parent_frame,
            text="\u23fb Encerrar Sessao (Logout)",
            command=lambda: logout_sap(self.label_status_global),
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            pady=12,
        )
        self.btn_logout.pack(fill="x", padx=40, pady=(10, 30))

    def _campo_popup(self, parent, texto, cores, show=None):
        tk.Label(parent, text=texto, bg=cores["bg"], fg=cores["texto"], font=("Segoe UI", 10, "bold")).pack(anchor="w")
        entry = tk.Entry(
            parent,
            font=("Segoe UI", 11),
            show=show,
            bg=cores["input"],
            fg=cores["texto"],
            insertbackground=cores["texto"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=cores["container"],
            highlightcolor=cores["accent"],
        )
        entry.pack(fill="x", pady=(0, 8))
        return entry

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

    def _botao_hover(self, btn):
        c = self._cores()
        btn.configure(bg=c["hover"], fg="white")

    def _botao_normal(self, btn):
        c = self._cores()
        btn.configure(bg=c["container"], fg=c["texto"])


def construir_aba_sap(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaSap(parent_frame, label_status_global, paletas, obter_tema_atual)


if __name__ == "__main__":
    janela_teste = tk.Tk()
    janela_teste.title("Teste Isolado - Modulo SAP")
    janela_teste.geometry("700x500")

    frame_principal = ttk.Frame(janela_teste)
    frame_principal.pack(fill="both", expand=True)

    lbl_teste_status = tk.Label(janela_teste, text=" SAP desconectado", font=("Segoe UI", 10, "bold"), anchor="w")
    lbl_teste_status.pack(fill="x", side="bottom", padx=10, pady=10)

    construir_aba_sap(frame_principal, lbl_teste_status)
    janela_teste.mainloop()
