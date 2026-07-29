import os
import json
import subprocess
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import messagebox, ttk
import win32api
import win32print

import pandas as pd
from gerador import gerar_pdf_placa_tradicional, gerar_html_para_impressao


def imprimir_direto_na_impressora(deposito, material, descricao, qtd, unidade, lote="", tamanho="Grande", qr_topo=""):
    impressora = win32print.GetDefaultPrinter()
    html_file = gerar_html_para_impressao(
        deposito, material, descricao, qtd, unidade,
        lote=lote, tamanho=tamanho, qr_topo=qr_topo,
    )
    os.startfile(html_file)
    return True, impressora


CAMINHO_PLANILHA_CONTROLE = r"\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm"
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwfZ0U67WaIoXw9Cy1jz9BISe4X2PFYbx2k-nX34m49JeiFZUuK_BQlTogjw8OICa0h/exec"


class AbaEtiquetas:
    def __init__(self, parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status=None):
        self.parent_frame = parent_frame
        self.root = root
        self.arquivo_banco = arquivo_banco
        self.paletas = paletas
        self.obter_tema_atual = obter_tema_atual
        self.label_status = label_status

        self.banco = pd.DataFrame()
        self.janela_sugestoes = None
        self.lista_sugestoes = None
        self.seriais_lote = []

        self._construir_interface()
        self.carregar_banco_dados()
        self.atualizar_preview()

    def carregar_banco_dados(self):
        try:
            self.banco = pd.read_excel(self.arquivo_banco)
            self.banco.columns = self.banco.columns.str.strip()
            if "Material" in self.banco.columns:
                self.banco["Material"] = self.banco["Material"].astype(str).str.strip()
            return True
        except Exception:
            self.banco = pd.DataFrame()
            return False

    def forcar_atualizacao_banco(self):
        if self.carregar_banco_dados():
            messagebox.showinfo("Sucesso", "Banco de dados atualizado com sucesso!")
        else:
            messagebox.showwarning("Aviso", f"Nao foi possivel ler o arquivo:\n{self.arquivo_banco}")

    def buscar_lote(self):
        lote = self.campo_lote.get().strip()
        if not lote:
            messagebox.showwarning("Aviso", "Digite o numero do lote primeiro.")
            return

        try:
            df = pd.read_excel(CAMINHO_PLANILHA_CONTROLE, sheet_name="ESTOQUE", engine="openpyxl")
            df.columns = df.columns.str.strip()

            if "LOTE" not in df.columns:
                messagebox.showerror("Erro", "Coluna 'LOTE' nao encontrada na planilha.")
                return

            df["LOTE"] = df["LOTE"].astype(str).str.strip()
            filtrado = df[df["LOTE"] == lote]

            if filtrado.empty:
                messagebox.showwarning("Aviso", f"Lote '{lote}' nao encontrado na planilha.")
                return

            self.seriais_lote = filtrado["SERIE"].astype(str).tolist()
            qtd = len(self.seriais_lote)

            self.campo_qtd.delete(0, tk.END)
            self.campo_qtd.insert(0, str(qtd))

            if "MATERIAL" in filtrado.columns:
                material = filtrado.iloc[0]["MATERIAL"]
                self.campo_material.delete(0, tk.END)
                self.campo_material.insert(0, str(material))
                self.buscar_material()

            if "DEPOSITO" in filtrado.columns:
                deposito = str(filtrado.iloc[0]["DEPOSITO"]).strip()
                if deposito in ["CL03", "CL04", "CS06"]:
                    self.deposito_var.set(deposito)

            self.atualizar_preview()
            self._atualizar_status(f"Lote {lote}: {qtd} seriais encontrados.")

        except Exception as e:
            messagebox.showerror("Erro ao Ler Planilha", f"Nao foi possivel ler a planilha de controle:\n{str(e)}")

    def gerar_doc_google(self):
        if not self.seriais_lote:
            messagebox.showwarning("Aviso", "Busque um lote primeiro antes de gerar o documento.")
            return

        if not GOOGLE_APPS_SCRIPT_URL:
            messagebox.showwarning("Aviso", "Configure a URL do Google Apps Script primeiro.")
            return

        lote = self.campo_lote.get().strip()
        material = self.campo_material.get().strip()
        descricao = self.campo_descricao.get().strip()
        deposito = self.deposito_var.get()

        dados = {
            "lote": lote,
            "material": material,
            "descricao": descricao,
            "deposito": deposito,
            "seriais": self.seriais_lote,
            "quantidade": len(self.seriais_lote)
        }

        try:
            dados_json = json.dumps(dados).encode("utf-8")
            requisicao = urllib.request.Request(
                GOOGLE_APPS_SCRIPT_URL,
                data=dados_json,
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(requisicao, timeout=30) as resposta:
                resultado = json.loads(resposta.read().decode("utf-8"))
                link = resultado.get("link", "")
                if link:
                    self.campo_qr_topo.delete(0, tk.END)
                    self.campo_qr_topo.insert(0, link)
                    self.atualizar_preview()
                    self._atualizar_status(f"Google Doc criado: {link}")
                    messagebox.showinfo("Sucesso", f"Documento criado com sucesso!\n\nLink: {link}")
                else:
                    messagebox.showerror("Erro", "Nenhum link retornado pelo Google Apps Script.")
        except urllib.error.URLError:
            messagebox.showerror("Erro", "Nao foi possivel conectar com Google Apps Script.\nVerifique a URL e sua conexao com a internet.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar documento:\n{str(e)}")

    def aplicar_tema(self, tema_nome):
        c = self.paletas[tema_nome]

        for campo in [self.campo_material, self.campo_descricao, self.campo_lote, self.campo_qtd, self.campo_qr_topo]:
            campo.configure(
                bg=c["input"],
                fg=c["texto"],
                insertbackground=c["texto"],
                bd=0,
                highlightthickness=1,
                highlightbackground=c["container"],
                highlightcolor=c["accent"],
            )

        self.check_abrir.configure(
            bg=c["bg"],
            fg=c["texto"],
            selectcolor=c["input"],
            activebackground=c["bg"],
            activeforeground=c["texto"],
        )
        self.check_imprimir.configure(
            bg=c["bg"],
            fg=c["texto"],
            selectcolor=c["input"],
            activebackground=c["bg"],
            activeforeground=c["texto"],
        )

        self.btn_gerar.configure(bg=c["accent"], fg="white", activebackground=c["hover"], activeforeground="white")
        self.btn_atualizar_banco.configure(
            bg=c["container"],
            fg=c["texto"],
            activebackground=c["hover"],
            activeforeground="white",
        )
        self.btn_buscar_lote.configure(
            bg=c["container"],
            fg=c["texto"],
            activebackground=c["hover"],
            activeforeground="white",
        )
        self.btn_gerar_doc.configure(
            bg=c["accent"],
            fg="white",
            activebackground=c["hover"],
            activeforeground="white",
        )

        if self.lista_sugestoes and self.lista_sugestoes.winfo_exists():
            self.lista_sugestoes.configure(
                bg=c["input"],
                fg=c["texto"],
                selectbackground=c["accent"],
                selectforeground="white",
                highlightbackground=c["accent"],
            )

    def fechar_sugestoes(self):
        if self.janela_sugestoes and self.janela_sugestoes.winfo_exists():
            self.janela_sugestoes.destroy()
        self.janela_sugestoes = None
        self.lista_sugestoes = None

    def verificar_fechar(self, event=None):
        self.root.after(150, self._executar_fechar_se_fora)

    def _executar_fechar_se_fora(self):
        if not self.janela_sugestoes:
            return

        foco = self.root.focus_get()
        if foco != self.campo_material and foco != self.lista_sugestoes:
            self.fechar_sugestoes()

    def filtrar_sugestoes(self, event):
        self.atualizar_preview()

        if event.keysym == "Down" and self.janela_sugestoes:
            self.lista_sugestoes.focus_set()
            self.lista_sugestoes.selection_set(0)
            return

        if event.keysym in ["Up", "Return", "Escape", "Tab"]:
            return

        texto_digitado = self.campo_material.get().strip()
        if not texto_digitado:
            self.fechar_sugestoes()
            self.campo_descricao.delete(0, tk.END)
            return

        if self.banco.empty or "Material" not in self.banco.columns:
            self.fechar_sugestoes()
            return

        resultados = self.banco[self.banco["Material"].str.contains(texto_digitado, case=False, na=False)]
        opcoes = list(resultados["Material"].unique())[:8]
        if not opcoes:
            self.fechar_sugestoes()
            return

        c = self.paletas[self.obter_tema_atual()]
        if not self.janela_sugestoes or not self.janela_sugestoes.winfo_exists():
            self.janela_sugestoes = tk.Toplevel(self.root)
            self.janela_sugestoes.wm_overrideredirect(True)
            self.janela_sugestoes.wm_attributes("-topmost", True)

            self.lista_sugestoes = tk.Listbox(
                self.janela_sugestoes,
                bg=c["input"],
                fg=c["texto"],
                selectbackground=c["accent"],
                selectforeground="white",
                bd=0,
                highlightthickness=1,
                highlightbackground=c["accent"],
                font=("Segoe UI", 10),
                activestyle="none",
            )
            self.lista_sugestoes.pack(fill="both", expand=True)
            self.lista_sugestoes.bind("<ButtonRelease-1>", self.preencher_selecao)
            self.lista_sugestoes.bind("<Return>", self.preencher_selecao)
            self.lista_sugestoes.bind("<FocusOut>", self.verificar_fechar)

        self.lista_sugestoes.delete(0, tk.END)
        for opcao in opcoes:
            self.lista_sugestoes.insert(tk.END, opcao)

        x = self.campo_material.winfo_rootx()
        y = self.campo_material.winfo_rooty() + self.campo_material.winfo_height() + 2
        w = self.campo_material.winfo_width()
        h = len(opcoes) * 22 + 4
        self.janela_sugestoes.geometry(f"{w}x{h}+{x}+{y}")

    def preencher_selecao(self, event=None):
        if self.lista_sugestoes and self.lista_sugestoes.curselection():
            selecionado = self.lista_sugestoes.get(self.lista_sugestoes.curselection())
            self.campo_material.delete(0, tk.END)
            self.campo_material.insert(0, selecionado)
            self.buscar_material()
            self.fechar_sugestoes()
            self.campo_qtd.focus()

    def buscar_material(self, event=None):
        self.campo_descricao.delete(0, tk.END)

        if self.banco.empty or "Material" not in self.banco.columns:
            self.atualizar_preview()
            return

        res = self.banco[self.banco["Material"] == str(self.campo_material.get().strip())]
        coluna_descricao = self._obter_coluna_descricao()

        if not res.empty and coluna_descricao:
            self.campo_descricao.insert(0, res.iloc[0][coluna_descricao])

        self.atualizar_preview()

    def limpar_campos(self):
        self.fechar_sugestoes()
        self.campo_lote.delete(0, tk.END)
        self.campo_qr_topo.delete(0, tk.END)
        self.campo_lote.focus()
        self.atualizar_preview()

    def executar_geracao(self):
        mat_cod = self.campo_material.get().strip()
        lote_val = self.campo_lote.get().strip()
        dep_val = self.deposito_var.get()
        qtd_val = self.campo_qtd.get().strip()
        qr_topo_val = self.campo_qr_topo.get().strip()
        descricao = self.campo_descricao.get().strip()
        unidade = self.combo_unidade.get()
        tamanho = self.combo_tamanho.get()

        erro = self._validar_dados(dep_val)
        if erro:
            messagebox.showwarning("Aviso", erro)
            return

        try:
            arquivo = gerar_pdf_placa_tradicional(
                dep_val,
                mat_cod,
                descricao,
                qtd_val,
                unidade,
                lote=lote_val,
                tamanho=tamanho,
                qr_topo=qr_topo_val,
            )

            material_status = mat_cod if mat_cod else "identificacao livre"

            if self.check_imprimir_var.get():
                ok, resultado = imprimir_direto_na_impressora(
                    dep_val, mat_cod, descricao, qtd_val, unidade,
                    lote=lote_val, tamanho=tamanho, qr_topo=qr_topo_val,
                )
                if ok:
                    self._atualizar_status(f" Impresso direto em: {resultado}")
                else:
                    html_file = gerar_html_para_impressao(
                        dep_val, mat_cod, descricao, qtd_val, unidade,
                        lote=lote_val, tamanho=tamanho, qr_topo=qr_topo_val,
                    )
                    os.startfile(html_file)
                    self._atualizar_status(f" Fallback navegador: {resultado}")
            else:
                if self.check_abrir_var.get():
                    os.startfile(arquivo)
                self._atualizar_status(f" PDF gerado: {material_status}")

            messagebox.showinfo("Sucesso", "Etiqueta gerada com sucesso!")
            self.limpar_campos()
        except Exception as e:
            messagebox.showerror("Erro ao Gerar", str(e))

    def atualizar_preview(self, event=None):
        self.preview.delete("all")
        dep = self.deposito_var.get()
        tamanho = self.combo_tamanho.get()
        mat = self.campo_material.get().strip()
        desc = self.campo_descricao.get().strip()
        quantidade = self.campo_qtd.get().strip()
        un = self.combo_unidade.get()
        lt = self.campo_lote.get().strip()
        qr_t = self.campo_qr_topo.get().strip()

        mapa_largura = {"Pequena": 250, "M\u00e9dia": 350, "Grande": 450}
        largura = mapa_largura.get(tamanho, 450)
        x1, x2 = 240 - (largura // 2), 240 + (largura // 2)

        self.preview.create_rectangle(x1, 15, x2, 305, fill="white", outline="black", width=2)
        self.preview.create_text(
            x2 - 10,
            30,
            text=f"TAM: {tamanho.upper()}",
            font=("Segoe UI", 8, "bold"),
            fill="gray",
            anchor="ne",
        )
        self.preview.create_line(x1, 65, x2, 65, fill="black", width=1)
        self.preview.create_text(
            x1 + 15,
            40,
            text="CONECTA LOG",
            font=("Segoe UI", 11, "bold"),
            fill="black",
            anchor="w",
        )

        if qr_t:
            self.preview.create_rectangle(x2 - 45, 20, x2 - 15, 50, fill="black", outline="black")
            self.preview.create_rectangle(x2 - 41, 24, x2 - 19, 46, fill="white", outline="white")
            self.preview.create_rectangle(x2 - 35, 30, x2 - 25, 40, fill="black", outline="black")
            self.preview.create_text(
                x2 - 53,
                40,
                text=f"DEP: {dep}",
                font=("Segoe UI", 12, "bold"),
                fill="black",
                anchor="e",
            )
        else:
            self.preview.create_text(
                x2 - 15,
                40,
                text=f"DEP: {dep}",
                font=("Segoe UI", 12, "bold"),
                fill="black",
                anchor="e",
            )

        self.preview.create_text(240, 120, text=dep, font=("Segoe UI", 48, "bold"), fill="black", anchor="center")
        self.preview.create_text(
            240,
            175,
            text=mat if mat else "-----",
            font=("Segoe UI", 22, "bold"),
            fill="black",
            anchor="center",
        )

        desc_resumida = desc[:45] + "..." if len(desc) > 45 else desc
        self.preview.create_text(
            240,
            210,
            text=desc_resumida if desc else "[Aguardando Descricao]",
            font=("Segoe UI", 11, "bold"),
            fill="black" if desc else "gray",
            anchor="center",
        )
        self.preview.create_text(
            240,
            245,
            text=f"QUANTIDADE: {quantidade} {un}" if quantidade else "QUANTIDADE: -",
            font=("Segoe UI", 14, "bold"),
            fill="black",
            anchor="center",
        )
        if lt:
            self.preview.create_text(
                240,
                270,
                text=f"LOTE: {lt}",
                font=("Segoe UI", 14, "bold"),
                fill="black",
                anchor="center",
            )

        if mat:
            for x in range(240 - 80, 240 + 80, 4):
                largura_linha = 3 if x % 3 == 0 else 1
                self.preview.create_line(x, 285, x, 300, fill="black", width=largura_linha)

    def _construir_interface(self):
        coluna_esquerda = ttk.Frame(self.parent_frame)
        coluna_esquerda.pack(side="left", fill="both", expand=True, padx=5, pady=10)

        coluna_direita = ttk.Frame(self.parent_frame)
        coluna_direita.pack(side="right", fill="both", padx=5, pady=10)

        grupo_config = ttk.LabelFrame(coluna_esquerda, text=" Informacoes de Controle ", padding=12)
        grupo_config.pack(fill="x", pady=4)

        ttk.Label(grupo_config, text="Deposito:").grid(row=0, column=0, pady=4)
        self.deposito_var = ttk.Combobox(grupo_config, values=["CL03", "CL04", "CS06"], state="readonly", width=10)
        self.deposito_var.set("CL03")
        self.deposito_var.grid(row=0, column=1, padx=6)
        self.deposito_var.bind("<<ComboboxSelected>>", self.atualizar_preview)

        ttk.Label(grupo_config, text="Tamanho:").grid(row=0, column=2, padx=6)
        self.combo_tamanho = ttk.Combobox(
            grupo_config,
            values=["Pequena", "M\u00e9dia", "Grande"],
            state="readonly",
            width=10,
        )
        self.combo_tamanho.set("Grande")
        self.combo_tamanho.grid(row=0, column=3, padx=6)
        self.combo_tamanho.bind("<<ComboboxSelected>>", self.atualizar_preview)

        ttk.Label(grupo_config, text="No. Lote:").grid(row=0, column=4, padx=6)
        self.campo_lote = tk.Entry(grupo_config, width=14, font=("Segoe UI", 10))
        self.campo_lote.grid(row=0, column=5)
        self.campo_lote.bind("<KeyRelease>", self.atualizar_preview)

        self.btn_buscar_lote = tk.Button(
            grupo_config,
            text="Buscar Lote",
            command=self.buscar_lote,
            font=("Segoe UI", 9, "bold"),
            bd=1,
            relief="flat",
            cursor="hand2",
        )
        self.btn_buscar_lote.grid(row=0, column=6, padx=6)

        ttk.Label(grupo_config, text="QR Topo Dir:").grid(row=1, column=0, sticky="w", pady=6)
        self.campo_qr_topo = tk.Entry(grupo_config, width=44, font=("Segoe UI", 10))
        self.campo_qr_topo.grid(row=1, column=1, columnspan=5, sticky="w", padx=6, pady=6)
        self.campo_qr_topo.bind("<KeyRelease>", self.atualizar_preview)

        grupo_prod = ttk.LabelFrame(coluna_esquerda, text=" Identificacao do Material ", padding=12)
        grupo_prod.pack(fill="x", pady=4)

        ttk.Label(grupo_prod, text="Cod. Material:").grid(row=0, column=0, sticky="w", pady=4)
        self.campo_material = tk.Entry(grupo_prod, width=18, font=("Segoe UI", 10))
        self.campo_material.grid(row=0, column=1, sticky="w", pady=4)
        self.campo_material.bind("<KeyRelease>", self.filtrar_sugestoes)
        self.campo_material.bind("<FocusOut>", self.verificar_fechar)
        self.campo_material.bind("<Return>", self.buscar_material)

        self.btn_atualizar_banco = tk.Button(
            grupo_prod,
            text="\U0001f504 Atualizar Banco",
            command=self.forcar_atualizacao_banco,
            font=("Segoe UI", 9, "bold"),
            bd=1,
            relief="flat",
            cursor="hand2",
        )
        self.btn_atualizar_banco.grid(row=0, column=2, padx=10, sticky="w")
        self.btn_atualizar_banco.bind("<Enter>", lambda e: self._mouse_entrou_botao(self.btn_atualizar_banco))
        self.btn_atualizar_banco.bind("<Leave>", lambda e: self._mouse_saiu_botao_secundario(self.btn_atualizar_banco))

        self.btn_gerar_doc = tk.Button(
            grupo_prod,
            text="Gerar Doc Google",
            command=self.gerar_doc_google,
            font=("Segoe UI", 9, "bold"),
            bd=1,
            relief="flat",
            cursor="hand2",
        )
        self.btn_gerar_doc.grid(row=0, column=3, padx=10, sticky="w")
        self.btn_gerar_doc.bind("<Enter>", lambda e: self._mouse_entrou_botao(self.btn_gerar_doc))
        self.btn_gerar_doc.bind("<Leave>", lambda e: self._mouse_saiu_botao_secundario(self.btn_gerar_doc))

        ttk.Label(grupo_prod, text="Descricao:").grid(row=1, column=0, sticky="w", pady=4)
        self.campo_descricao = tk.Entry(grupo_prod, width=64, font=("Segoe UI", 10))
        self.campo_descricao.grid(row=1, column=1, columnspan=5, sticky="w", pady=4)
        self.campo_descricao.bind("<KeyRelease>", self.atualizar_preview)

        grupo_inv = ttk.LabelFrame(coluna_esquerda, text=" Dados de Inventario ", padding=12)
        grupo_inv.pack(fill="x", pady=4)

        ttk.Label(grupo_inv, text="Qtd:").grid(row=0, column=0, pady=4)
        self.campo_qtd = tk.Entry(grupo_inv, width=14, font=("Segoe UI", 10))
        self.campo_qtd.grid(row=0, column=1)
        self.campo_qtd.bind("<KeyRelease>", self.atualizar_preview)

        ttk.Label(grupo_inv, text="Unidade:").grid(row=0, column=2, padx=6)
        self.combo_unidade = ttk.Combobox(grupo_inv, values=["P\u00c7", "P\u00c7S", "M", "CE", "KG"], width=10)
        self.combo_unidade.set("P\u00c7S")
        self.combo_unidade.grid(row=0, column=3)
        self.combo_unidade.bind("<<ComboboxSelected>>", self.atualizar_preview)

        frame_opcoes = ttk.Frame(coluna_esquerda, padding=5)
        frame_opcoes.pack(fill="x", pady=5)

        self.check_abrir_var = tk.BooleanVar(value=True)
        self.check_abrir = tk.Checkbutton(
            frame_opcoes,
            text="\U0001f441 Abrir PDF criado",
            variable=self.check_abrir_var,
            font=("Segoe UI", 9),
            bd=0,
            relief="flat",
        )
        self.check_abrir.pack(side="left", padx=10)

        self.check_imprimir_var = tk.BooleanVar(value=False)
        self.check_imprimir = tk.Checkbutton(
            frame_opcoes,
            text="\U0001f5a8 Mandar direto para impressora",
            variable=self.check_imprimir_var,
            font=("Segoe UI", 9),
            bd=0,
            relief="flat",
        )
        self.check_imprimir.pack(side="left", padx=10)

        self.btn_gerar = tk.Button(
            coluna_esquerda,
            text="\U0001f3f7 GERAR ETIQUETA",
            command=self.executar_geracao,
            font=("Segoe UI", 11, "bold"),
            bd=0,
            cursor="hand2",
            pady=8,
        )
        self.btn_gerar.pack(fill="x", pady=10)
        self.btn_gerar.bind("<Enter>", lambda e: self._mouse_entrou_botao(self.btn_gerar))
        self.btn_gerar.bind("<Leave>", lambda e: self._mouse_saiu_botao(self.btn_gerar))

        grupo_preview = ttk.LabelFrame(coluna_direita, text=" Miniatura de Impressao em Tempo Real ", padding=5)
        grupo_preview.pack(fill="both", expand=True)

        self.preview = tk.Canvas(grupo_preview, width=480, height=320, bg="#edf2f7", bd=0, highlightthickness=0)
        self.preview.pack(fill="both", expand=True, padx=5, pady=5)

    def _obter_coluna_descricao(self):
        for coluna in self.banco.columns:
            if str(coluna).strip().lower() in ["descricao", "descri\u00e7\u00e3o"]:
                return coluna
        return None

    def _validar_dados(self, deposito):
        if not deposito:
            return "Selecione o deposito."
        return None

    def _atualizar_status(self, texto):
        if self.label_status:
            self.label_status.configure(text=texto)

    def _mouse_entrou_botao(self, btn):
        c = self.paletas[self.obter_tema_atual()]
        btn.configure(bg=c["hover"])

    def _mouse_saiu_botao(self, btn):
        c = self.paletas[self.obter_tema_atual()]
        btn.configure(bg=c["accent"])

    def _mouse_saiu_botao_secundario(self, btn):
        c = self.paletas[self.obter_tema_atual()]
        btn.configure(bg=c["container"])


def construir_aba_etiquetas(parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status=None):
    return AbaEtiquetas(parent_frame, root, arquivo_banco, paletas, obter_tema_atual, label_status)
