import os
import sys
import tkinter as tk
from tkinter import ttk

from modulo_Sap import construir_aba_sap, obter_info_sap_atual
from modulo_bandeirada import construir_aba_bandeirada
from modulo_conferencia import construir_aba_conferencia
from modulo_etiquetas import construir_aba_etiquetas
from modulo_extrair_OSME import construir_aba_extrair_osme
from modulo_extrator import construir_aba_extrator
from modulo_chat import JanelaChat



def obter_caminho_recurso(nome_arquivo):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nome_arquivo)


def obter_diretorio_executavel():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(obter_diretorio_executavel(), "config_tema_hub.txt")

def obter_arquivo_banco():
    caminho_exe = os.path.join(obter_diretorio_executavel(), "banco_de_dados.xlsx")
    if os.path.exists(caminho_exe):
        return caminho_exe
    return obter_caminho_recurso("banco_de_dados.xlsx")

ARQUIVO_BANCO = obter_arquivo_banco()

PALETAS = {
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#f57c00",
        "hover": "#e65100",
        "input": "#1f3a60",
        "scrollbar": "#5f7fa8",
        "scrollbar_hover": "#7fa2cf",
        "scrollbar_trough": "#10233d",
    },
    "Dark Mode": {
        "bg": "#1a202c",
        "container": "#2d3748",
        "texto": "#ffffff",
        "accent": "#319795",
        "hover": "#2c7a7b",
        "input": "#3a4454",
        "scrollbar": "#64748b",
        "scrollbar_hover": "#94a3b8",
        "scrollbar_trough": "#202938",
    },
    "Light Clean": {
        "bg": "#ffffff",
        "container": "#f8f9fa",
        "texto": "#2d3748",
        "accent": "#319795",
        "hover": "#2c7a7b",
        "input": "#edf2f7",
        "scrollbar": "#94a3b8",
        "scrollbar_hover": "#64748b",
        "scrollbar_trough": "#e2e8f0",
    },
    "Corporate Slate": {
        "bg": "#111827",
        "container": "#1f2937",
        "texto": "#f9fafb",
        "accent": "#2563eb",
        "hover": "#1d4ed8",
        "input": "#374151",
        "scrollbar": "#6b7280",
        "scrollbar_hover": "#9ca3af",
        "scrollbar_trough": "#172033",
    },
    "Office Light": {
        "bg": "#f3f6fb",
        "container": "#ffffff",
        "texto": "#1f2937",
        "accent": "#2563eb",
        "hover": "#1d4ed8",
        "input": "#e8eef7",
        "scrollbar": "#a7b4c8",
        "scrollbar_hover": "#6b7f99",
        "scrollbar_trough": "#dce5f2",
    },
    "Emerald Pro": {
        "bg": "#0f1f1b",
        "container": "#18352e",
        "texto": "#f8fafc",
        "accent": "#10b981",
        "hover": "#059669",
        "input": "#21483e",
        "scrollbar": "#6ba99a",
        "scrollbar_hover": "#8fd4c3",
        "scrollbar_trough": "#132b25",
    },
    "Graphite Orange": {
        "bg": "#171717",
        "container": "#262626",
        "texto": "#fafafa",
        "accent": "#f97316",
        "hover": "#ea580c",
        "input": "#3f3f46",
        "scrollbar": "#737373",
        "scrollbar_hover": "#a3a3a3",
        "scrollbar_trough": "#202020",
    },
    "Royal Indigo": {
        "bg": "#17152b",
        "container": "#262447",
        "texto": "#ffffff",
        "accent": "#8b5cf6",
        "hover": "#7c3aed",
        "input": "#34315f",
        "scrollbar": "#8179bd",
        "scrollbar_hover": "#a69cf0",
        "scrollbar_trough": "#201d3b",
    },
}

etiquetas_view = None
sap_view = None
bandeirada_view = None
extrator_view = None
osme_view = None
conferencia_view = None


def carregar_tema_salvo():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                tema = f.read().strip()
                if tema in PALETAS:
                    return tema
        except Exception:
            pass
    return "Conecta Blue"


def salvar_tema_atual(tema_nome):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(tema_nome)
    except Exception:
        pass


def obter_tema_atual():
    return combo_temas.get()


def aplicar_icone_janela(janela):
    try:
        janela.iconbitmap(obter_caminho_recurso("icone_lua.ico"))
    except Exception:
        pass

    try:
        icone_taskbar = tk.PhotoImage(file=obter_caminho_recurso("icone_lua.png"))
        janela.iconphoto(True, icone_taskbar)
        janela.icone_taskbar = icone_taskbar
    except Exception:
        pass


def aplicar_tema(tema_nome):
    c = PALETAS[tema_nome]
    scrollbar = c.get("scrollbar", c["input"])
    scrollbar_hover = c.get("scrollbar_hover", c["accent"])
    scrollbar_trough = c.get("scrollbar_trough", c["bg"])
    root.configure(bg=c["bg"])
    salvar_tema_atual(tema_nome)

    style.configure("TFrame", background=c["bg"])
    style.configure("TLabel", background=c["bg"], foreground=c["texto"])
    style.configure("TLabelframe", background=c["bg"], foreground=c["accent"])
    style.configure(
        "TLabelframe.Label",
        background=c["bg"],
        foreground=c["accent"],
        font=("Segoe UI", 10, "bold"),
    )
    style.configure("TNotebook", background=c["bg"])
    style.configure(
        "TNotebook.Tab",
        background=c["container"],
        foreground=c["texto"],
        font=("Segoe UI", 10, "bold"),
        padding=[15, 5],
    )
    style.map("TNotebook.Tab", background=[("selected", c["accent"])], foreground=[("selected", "white")])

    try:
        style.layout(
            "Vertical.TScrollbar",
            [
                (
                    "Vertical.Scrollbar.trough",
                    {
                        "sticky": "ns",
                        "children": [("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
                    },
                )
            ],
        )
        style.layout(
            "Horizontal.TScrollbar",
            [
                (
                    "Horizontal.Scrollbar.trough",
                    {
                        "sticky": "we",
                        "children": [("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
                    },
                )
            ],
        )
    except tk.TclError:
        pass

    for scrollbar_style in ["Vertical.TScrollbar", "Horizontal.TScrollbar", "TScrollbar"]:
        style.configure(
            scrollbar_style,
            gripcount=0,
            background=scrollbar,
            darkcolor=scrollbar,
            lightcolor=scrollbar,
            troughcolor=scrollbar_trough,
            bordercolor=scrollbar_trough,
            arrowcolor=c["texto"],
            relief="flat",
            borderwidth=0,
            width=12,
            arrowsize=10,
        )
        style.map(
            scrollbar_style,
            background=[("pressed", c["accent"]), ("active", scrollbar_hover)],
            troughcolor=[("active", scrollbar_trough)],
            arrowcolor=[("active", "white")],
        )

    for widget_type in ["TCombobox", "TSpinbox"]:
        style.configure(
            widget_type,
            fieldbackground=c["input"],
            background=c["input"],
            foreground=c["texto"],
            bordercolor=c["container"],
            arrowcolor=c["accent"],
        )
        style.map(
            widget_type,
            fieldbackground=[("readonly", c["input"])],
            foreground=[("readonly", c["texto"])],
            background=[("readonly", c["input"])],
        )

    root.option_add("*TCombobox*Listbox.background", c["input"])
    root.option_add("*TCombobox*Listbox.foreground", c["texto"])
    root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])

    bar_status.configure(bg=c["container"])
    lbl_status.configure(bg=c["container"], fg=c["texto"])
    lbl_sap_info.configure(bg=c["container"], fg=c["texto"])
    lbl_creditos.configure(bg=c["container"], fg=c["texto"])
    lbl_logo.configure(bg=c["bg"], fg=c["texto"])

    if etiquetas_view:
        etiquetas_view.aplicar_tema(tema_nome)
    if sap_view:
        sap_view.aplicar_tema(tema_nome)
    if bandeirada_view:
        bandeirada_view.aplicar_tema(tema_nome)
    if extrator_view:
        extrator_view.aplicar_tema(tema_nome)
    if osme_view:
        osme_view.aplicar_tema(tema_nome)
    if conferencia_view:
        conferencia_view.aplicar_tema(tema_nome)


def atualizar_info_sap():
    lbl_sap_info.configure(text=obter_info_sap_atual())
    root.after(10000, atualizar_info_sap)


root = tk.Tk()
root.title("Conecta Hub 2.0 - Hub de Ferramentas")
root.geometry("1160x700")
aplicar_icone_janela(root)

style = ttk.Style()
style.theme_use("clam")
tema_inicial = carregar_tema_salvo()

frame_topo = ttk.Frame(root, padding=10)
frame_topo.pack(fill=tk.X)

caminho_img = obter_caminho_recurso("logo_cntc.png")
try:
    img_logo = tk.PhotoImage(file=caminho_img).subsample(3, 3)
    lbl_logo = tk.Label(frame_topo, image=img_logo, borderwidth=0)
    lbl_logo.image = img_logo
except Exception:
    lbl_logo = tk.Label(frame_topo, text="CONECTA EMPREENDIMENTOS", font=("Segoe UI", 14, "bold"), borderwidth=0)
lbl_logo.pack(side=tk.LEFT, padx=10)

ttk.Label(frame_topo, text="Conecta Hub 2.0", font=("Segoe UI", 16, "bold")).pack(side="left", padx=10)

combo_temas = ttk.Combobox(frame_topo, values=list(PALETAS.keys()), state="readonly", width=14)
combo_temas.set(tema_inicial)
combo_temas.pack(side=tk.RIGHT, pady=10)
combo_temas.bind("<<ComboboxSelected>>", lambda e: aplicar_tema(combo_temas.get()))

notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

tab_gerador = ttk.Frame(notebook)
notebook.add(tab_gerador, text="\U0001f3ed Gerador de Etiquetas")

tab_extrator = ttk.Frame(notebook)
notebook.add(tab_extrator, text="\U0001f522 Extrator de Seriais")

tab_sap = ttk.Frame(notebook)
notebook.add(tab_sap, text="\U0001f510 Login SAP")

tab_bandeirada = ttk.Frame(notebook)
notebook.add(tab_bandeirada, text="\u2691 BANDEIRADA")

tab_osme = ttk.Frame(notebook)
notebook.add(tab_osme, text="\U0001f4c4 Extrator OSME")

tab_conferencia = ttk.Frame(notebook)
notebook.add(tab_conferencia, text="\U0001f4e6 Conferencia")

bar_status = tk.Frame(root, bd=1, relief="sunken")
bar_status.pack(side="bottom", fill="x")

lbl_status = tk.Label(bar_status, text=" Modulo Gerador Ativo | Conecta Hub 2.0", font=("Segoe UI", 9, "bold"))
lbl_status.pack(side="left", padx=5, pady=3)

lbl_sap_info = tk.Label(bar_status, text="SAP: verificando...", font=("Segoe UI", 9, "bold"))
lbl_sap_info.pack(side="left", padx=18, pady=3)

lbl_creditos = tk.Label(bar_status, text="2026 Desenvolvido por 🟢 Anderson Vieira™", font=("Segoe UI", 9, "italic"))
lbl_creditos.pack(side="right", padx=10, pady=3)

etiquetas_view = construir_aba_etiquetas(tab_gerador, root, ARQUIVO_BANCO, PALETAS, obter_tema_atual, lbl_status)
extrator_view = construir_aba_extrator(tab_extrator, lbl_status, PALETAS, obter_tema_atual)
sap_view = construir_aba_sap(tab_sap, lbl_status, PALETAS, obter_tema_atual)
bandeirada_view = construir_aba_bandeirada(tab_bandeirada, lbl_status, PALETAS, obter_tema_atual)
osme_view = construir_aba_extrair_osme(tab_osme, lbl_status, PALETAS, obter_tema_atual)
conferencia_view = construir_aba_conferencia(tab_conferencia, lbl_status, PALETAS, obter_tema_atual)

chat_janela_ref = [None]

def _trocar_aba(tab):
    notebook.select(tab)

def _buscar_lote_chat(lote):
    _trocar_aba(tab_gerador)
    etiquetas_view.campo_lote.delete(0, tk.END)
    etiquetas_view.campo_lote.insert(0, lote)
    etiquetas_view.buscar_lote()
    return {
        "lote": lote,
        "deposito": etiquetas_view.deposito_var.get(),
        "material": etiquetas_view.campo_material.get(),
        "quantidade": etiquetas_view.campo_qtd.get(),
        "seriais_encontrados": len(etiquetas_view.seriais_lote),
        "primeiros_seriais": etiquetas_view.seriais_lote[:5]
    }

def _gerar_doc_chat(lote):
    _trocar_aba(tab_gerador)
    etiquetas_view.campo_lote.delete(0, tk.END)
    etiquetas_view.campo_lote.insert(0, lote)
    etiquetas_view.buscar_lote()
    etiquetas_view.gerar_doc_google()
    link = etiquetas_view.campo_qr_topo.get()
    if link:
        return {"lote": lote, "link": link, "status": "Documento criado com sucesso"}
    return {"erro": "Nao foi possivel criar o documento"}

def _gerar_etiqueta_chat(deposito, material="", descricao="", qtd="", unidade="PCS", lote=""):
    _trocar_aba(tab_gerador)
    from gerador import gerar_pdf_placa_tradicional
    try:
        caminho = gerar_pdf_placa_tradicional(
            deposito, material, descricao, qtd, unidade,
            lote=lote, tamanho="Grande"
        )
        return {"status": "PDF gerado com sucesso", "arquivo": caminho, "deposito": deposito, "material": material}
    except Exception as e:
        return {"erro": str(e)}

def _consultar_material_chat(codigo):
    _trocar_aba(tab_gerador)
    etiquetas_view.campo_material.delete(0, tk.END)
    etiquetas_view.campo_material.insert(0, codigo)
    etiquetas_view.buscar_material()
    desc = etiquetas_view.campo_descricao.get()
    if desc:
        return {"codigo": codigo, "descricao": desc}
    return {"codigo": codigo, "descricao": "Nao encontrada no banco de dados"}

def _listar_lotes_chat():
    import pandas as pd
    try:
        caminho = r"\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm"
        df = pd.read_excel(caminho, sheet_name="ESTOQUE", engine="openpyxl")
        df.columns = df.columns.str.strip()
        if "LOTE" in df.columns:
            lotes = df["LOTE"].dropna().unique().tolist()
            lotes_str = [str(l).strip() for l in lotes]
            return {"total": len(lotes_str), "lotes": sorted(lotes_str)}
        return {"erro": "Coluna LOTE nao encontrada"}
    except Exception as e:
        return {"erro": str(e)}

def _extrair_seriais_chat(texto):
    _trocar_aba(tab_extrator)
    extrator_view.txt_entrada.delete("1.0", tk.END)
    extrator_view.txt_entrada.insert("1.0", texto)
    extrator_view.extrair_seriais()
    resultado = extrator_view.txt_sem_ponto_virgula.get("1.0", tk.END).strip()
    lista = [l.strip() for l in resultado.splitlines() if l.strip()]
    return {"total": len(lista), "seriais": lista}

def _login_sap_chat(usuario):
    _trocar_aba(tab_sap)
    try:
        from modulo_Sap import logar_sap, logins_salvos, carregar_logins
        carregar_logins()
        nome_lower = usuario.lower().strip()
        encontrado = None
        for nome_salvo, dados in logins_salvos.items():
            if nome_salvo.lower() == nome_lower:
                encontrado = dados
                break
        if not encontrado:
            for nome_salvo, dados in logins_salvos.items():
                if nome_lower in nome_salvo.lower():
                    encontrado = dados
                    usuario = nome_salvo
                    break
        if not encontrado:
            nomes = list(logins_salvos.keys())
            return {"erro": f"Usuario '{usuario}' nao encontrado. Usuarios disponiveis: {', '.join(nomes)}"}
        logar_sap(encontrado["usuario"], encontrado["senha"], usuario, lbl_status)
        return {"status": f"Login SAP realizado para {usuario} ({encontrado['usuario']})"}
    except Exception as e:
        return {"erro": str(e)}

def _info_sap_chat():
    try:
        info = obter_info_sap_atual()
        return {"info": info}
    except Exception as e:
        return {"erro": str(e)}

def _copiar_seriais_chat():
    _trocar_aba(tab_extrator)
    extrator_view.copiar_sem_ponto_virgula()
    return {"status": "Serais copiados para a area de transferencia"}

funcoes_chat = {
    "buscar_lote": _buscar_lote_chat,
    "gerar_doc_google": _gerar_doc_chat,
    "gerar_etiqueta": _gerar_etiqueta_chat,
    "consultar_material": _consultar_material_chat,
    "listar_lotes": _listar_lotes_chat,
    "extrair_seriais": _extrair_seriais_chat,
    "login_sap": _login_sap_chat,
    "info_sap": _info_sap_chat,
    "copiar_seriais": _copiar_seriais_chat,
}

def abrir_chat():
    if chat_janela_ref[0] and chat_janela_ref[0].winfo_exists():
        chat_janela_ref[0].lift()
        chat_janela_ref[0].focus_force()
        return
    chat_popup = JanelaChat(root, PALETAS, obter_tema_atual, funcoes_chat)
    chat_janela_ref[0] = chat_popup.janela

btn_chat = tk.Button(
    root, text="💬", font=("Segoe UI", 16),
    bg="#f57c00", fg="white", bd=0, width=3, height=1,
    command=abrir_chat, cursor="hand2",
    activebackground="#e65100", relief="flat"
)
btn_chat.place(relx=1.0, rely=1.0, x=-70, y=-45)

aplicar_tema(tema_inicial)
atualizar_info_sap()
root.mainloop()
