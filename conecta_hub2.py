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

aplicar_tema(tema_inicial)
atualizar_info_sap()
root.mainloop()
