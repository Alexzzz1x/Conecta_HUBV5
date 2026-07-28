from collections import Counter
import re
import tkinter as tk
from tkinter import messagebox, ttk


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


class AbaConferencia:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.relatorio_texto = ""

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._atualizar_cards(0, 0, 0, 0, 0)
        self._atualizar_banner("Aguardando conferencia", "neutro")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)

        for widget in [self.txt_base, self.txt_fisico, self.txt_faltantes, self.txt_extras, self.txt_presentes]:
            widget.configure(
                bg=c["input"],
                fg=c["texto"],
                insertbackground=c["texto"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=c["container"],
                highlightcolor=c["accent"],
            )

        for frame in self.cards:
            frame.configure(bg=c["container"])

        for label in self.card_labels:
            label.configure(bg=c["container"], fg=c["texto"])

        self.card_base_valor.configure(bg=c["container"], fg="#8b5cf6")
        self.card_fisico_valor.configure(bg=c["container"], fg="#3b82f6")
        self.card_ok_valor.configure(bg=c["container"], fg="#10b981")
        self.card_faltantes_valor.configure(bg=c["container"], fg="#ef4444")
        self.card_extras_valor.configure(bg=c["container"], fg="#f59e0b")
        self.lbl_banner.configure(bg=c["container"], fg=c["texto"])
        for frame in [self.frame_faltantes, self.frame_extras, self.frame_presentes]:
            frame.configure(bg=c["container"])
        self.lbl_faltantes_titulo.configure(bg=c["container"], fg="#ef4444")
        self.lbl_extras_titulo.configure(bg=c["container"], fg="#f59e0b")
        self.lbl_presentes_titulo.configure(bg=c["container"], fg="#10b981")
        self.lbl_faltantes_qtd.configure(bg=c["container"], fg="#ef4444")
        self.lbl_extras_qtd.configure(bg=c["container"], fg="#f59e0b")
        self.lbl_presentes_qtd.configure(bg=c["container"], fg="#10b981")

        for btn, principal in [
            (self.btn_comparar, True),
            (self.btn_copiar, False),
            (self.btn_limpar, False),
        ]:
            if principal:
                btn.configure(bg=c["accent"], fg="white", activebackground=c["hover"], activeforeground="white")
            else:
                btn.configure(bg=c["container"], fg=c["texto"], activebackground=c["hover"], activeforeground="white")

    def comparar(self):
        base = self._extrair_itens(self.txt_base.get("1.0", tk.END))
        fisico = self._extrair_itens(self.txt_fisico.get("1.0", tk.END))

        if not base and not fisico:
            messagebox.showwarning("Aviso", "Cole os materiais/seriais recebidos e coletados antes de comparar.")
            return

        cont_base = Counter(base)
        cont_fisico = Counter(fisico)

        presentes = []
        faltantes = []
        extras = []

        for item in sorted(set(cont_base) | set(cont_fisico)):
            qtd_base = cont_base[item]
            qtd_fisico = cont_fisico[item]
            qtd_ok = min(qtd_base, qtd_fisico)

            if qtd_ok:
                presentes.extend([item] * qtd_ok)
            if qtd_base > qtd_fisico:
                faltantes.extend([item] * (qtd_base - qtd_fisico))
            if qtd_fisico > qtd_base:
                extras.extend([item] * (qtd_fisico - qtd_base))

        self._atualizar_cards(len(base), len(fisico), len(presentes), len(faltantes), len(extras))
        self._gerar_relatorio(base, fisico, presentes, faltantes, extras)

        if faltantes or extras:
            self._atualizar_banner(f"DIVERGENCIA: {len(faltantes)} faltantes e {len(extras)} extras", "erro")
            self._status(f"Conferencia finalizada: {len(faltantes)} faltantes e {len(extras)} extras.")
        else:
            self._atualizar_banner("OK: recebimento conferido sem divergencias", "ok")
            self._status("Conferencia finalizada: quantidades batem.")

    def copiar_relatorio(self):
        texto = getattr(self, "relatorio_texto", "").strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nao ha relatorio para copiar.")
            return

        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        self.parent_frame.update()
        self._status("Relatorio de conferencia copiado.")

    def limpar(self):
        for widget in [self.txt_base, self.txt_fisico, self.txt_faltantes, self.txt_extras, self.txt_presentes]:
            self._set_texto(widget, "")
        self.relatorio_texto = ""
        self._atualizar_cards(0, 0, 0, 0, 0)
        self._atualizar_resultados_visuais([], [], [])
        self._atualizar_banner("Aguardando conferencia", "neutro")
        self._status("Conferencia limpa.")

    def _extrair_itens(self, texto):
        itens = []
        for linha in texto.splitlines():
            linha = linha.strip().replace(";", "")
            if not linha:
                continue

            partes = re.split(r"[\t, ]+", linha)
            for parte in partes:
                item = parte.strip().replace(";", "")
                if item:
                    itens.append(item.upper())

        return itens

    def _gerar_relatorio(self, base, fisico, presentes, faltantes, extras):
        status = "OK - RECEBIMENTO CONFERIDO" if not faltantes and not extras else "DIVERGENCIA ENCONTRADA"
        linha = "=" * 72
        sublinha = "-" * 72

        linhas = [
            linha,
            " RELATORIO DE CONFERENCIA DE RECEBIMENTO",
            linha,
            "",
            f" STATUS: {status}",
            "",
            " RESUMO",
            sublinha,
            f" SAP/Base recebido : {len(base):>5}",
            f" Fisico coletado   : {len(fisico):>5}",
            f" OK presentes      : {len(presentes):>5}",
            f" Faltantes fisico  : {len(faltantes):>5}",
            f" Extras fisico     : {len(extras):>5}",
            "",
        ]
        linhas.extend(self._linhas_secao("FALTANTES NO FISICO", faltantes, "Nenhum faltante."))
        linhas.append("")
        linhas.extend(self._linhas_secao("EXTRAS NO FISICO", extras, "Nenhum extra."))
        linhas.append("")
        linhas.extend(self._linhas_secao("OK / PRESENTES", presentes, "Nenhum item confirmado."))
        linhas.append("")
        linhas.append(linha)
        self.relatorio_texto = "\n".join(linhas)
        self._atualizar_resultados_visuais(faltantes, extras, presentes)

    def _linhas_secao(self, titulo, itens, texto_vazio):
        sublinha = "-" * 72
        if not itens:
            return [titulo, sublinha, f"  {texto_vazio}"]

        contagem = Counter(itens)
        linhas = [titulo, sublinha]
        for item in sorted(contagem):
            qtd = contagem[item]
            linhas.append(f"  {item}" if qtd == 1 else f"  {item:<30} x{qtd}")
        return linhas

    def _construir_interface(self):
        container = ttk.Frame(self.parent_frame)
        container.pack(fill="both", expand=True, padx=18, pady=12)
        container.columnconfigure(0, weight=0)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        self._construir_cards(container)

        painel_esquerdo = ttk.Frame(container)
        painel_esquerdo.grid(row=1, column=0, sticky="ns", padx=(0, 12))

        ttk.Label(painel_esquerdo, text="\U0001f4e5 1. Material recebido no SAP/Base").pack(anchor="w")
        frame_base, self.txt_base = self._text_area(painel_esquerdo, height=12)
        frame_base.pack(fill="both", expand=True, pady=(6, 12))

        ttk.Label(painel_esquerdo, text="\U0001f4e6 2. Material coletado no fisico").pack(anchor="w")
        frame_fisico, self.txt_fisico = self._text_area(painel_esquerdo, height=12)
        frame_fisico.pack(fill="both", expand=True, pady=(6, 12))

        self.btn_comparar = tk.Button(
            painel_esquerdo,
            text="\U0001f50d COMPARAR RECEBIMENTO",
            command=self.comparar,
            font=("Segoe UI", 11, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            pady=12,
        )
        self.btn_comparar.pack(fill="x")

        painel_relatorio = ttk.Frame(container)
        painel_relatorio.grid(row=1, column=1, sticky="nsew")
        painel_relatorio.rowconfigure(1, weight=0)
        painel_relatorio.columnconfigure(0, weight=1)

        frame_titulo = ttk.Frame(painel_relatorio)
        frame_titulo.grid(row=0, column=0, sticky="ew")
        frame_titulo.columnconfigure(0, weight=1)

        ttk.Label(frame_titulo, text="\U0001f4cb Relatorio de conferencia").grid(row=0, column=0, sticky="w")

        self.btn_copiar = tk.Button(
            frame_titulo,
            text="\U0001f4cb Log",
            command=self.copiar_relatorio,
            font=("Segoe UI", 9, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=5,
        )
        self.btn_copiar.grid(row=0, column=1, padx=(8, 0))

        self.btn_limpar = tk.Button(
            frame_titulo,
            text="\U0001f9f9 Limpar",
            command=self.limpar,
            font=("Segoe UI", 9, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=5,
        )
        self.btn_limpar.grid(row=0, column=2, padx=(8, 0))

        self.lbl_banner = tk.Label(
            painel_relatorio,
            text="Aguardando conferencia",
            font=("Segoe UI", 11, "bold"),
            anchor="center",
            padx=12,
            pady=8,
        )
        self.lbl_banner.grid(row=1, column=0, sticky="ew", pady=(8, 0))

        frame_relatorio = ttk.Frame(painel_relatorio)
        frame_relatorio.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        painel_relatorio.rowconfigure(2, weight=1)
        frame_relatorio.columnconfigure(0, weight=1)
        frame_relatorio.columnconfigure(1, weight=1)
        frame_relatorio.columnconfigure(2, weight=1)
        frame_relatorio.rowconfigure(0, weight=1)

        (
            self.frame_faltantes,
            self.lbl_faltantes_titulo,
            self.lbl_faltantes_qtd,
            self.txt_faltantes,
        ) = self._resultado_card(frame_relatorio, 0, "[!] FALTANTES", "#ef4444")
        (
            self.frame_extras,
            self.lbl_extras_titulo,
            self.lbl_extras_qtd,
            self.txt_extras,
        ) = self._resultado_card(frame_relatorio, 2, "[+] EXTRAS", "#f59e0b")
        (
            self.frame_presentes,
            self.lbl_presentes_titulo,
            self.lbl_presentes_qtd,
            self.txt_presentes,
        ) = self._resultado_card(frame_relatorio, 1, "[OK] PRESENTES", "#10b981")

        for btn, principal in [(self.btn_comparar, True), (self.btn_copiar, False), (self.btn_limpar, False)]:
            btn.bind("<Enter>", lambda e, b=btn: self._botao_hover(b))
            btn.bind("<Leave>", lambda e, b=btn, p=principal: self._botao_normal(b, p))

    def _construir_cards(self, parent):
        frame_cards = ttk.Frame(parent)
        frame_cards.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        self.cards = []
        self.card_labels = []

        dados = [
            ("SAP\nBase", "#8b5cf6"),
            ("Fisico\nRecebido", "#3b82f6"),
            ("OK\nPresentes", "#10b981"),
            ("Faltantes\nFisico", "#ef4444"),
            ("Extras\nFisico", "#f59e0b"),
        ]

        valores = []
        for indice, (titulo, cor) in enumerate(dados):
            frame_cards.columnconfigure(indice, weight=1)
            card = tk.Frame(frame_cards, bd=0, padx=12, pady=8)
            card.grid(row=0, column=indice, sticky="ew", padx=(0 if indice == 0 else 6, 0))

            label = tk.Label(card, text=titulo, font=("Segoe UI", 8, "bold"))
            label.pack()
            valor = tk.Label(card, text="0", font=("Segoe UI", 22, "bold"), fg=cor)
            valor.pack()

            self.cards.append(card)
            self.card_labels.append(label)
            valores.append(valor)

        (
            self.card_base_valor,
            self.card_fisico_valor,
            self.card_ok_valor,
            self.card_faltantes_valor,
            self.card_extras_valor,
        ) = valores

    def _text_area(self, parent, height):
        frame = ttk.Frame(parent)
        texto = tk.Text(frame, width=34, height=height, font=("Consolas", 10), wrap="none")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=texto.yview)
        texto.configure(yscrollcommand=scroll.set)
        texto.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        return frame, texto

    def _atualizar_cards(self, base, fisico, ok, faltantes, extras):
        self.card_base_valor.configure(text=str(base))
        self.card_fisico_valor.configure(text=str(fisico))
        self.card_ok_valor.configure(text=str(ok))
        self.card_faltantes_valor.configure(text=str(faltantes))
        self.card_extras_valor.configure(text=str(extras))

    def _set_texto(self, widget, texto):
        widget.delete("1.0", tk.END)
        widget.insert("1.0", texto)

    def _resultado_card(self, parent, column, titulo, cor):
        frame = tk.Frame(parent, padx=10, pady=10)
        frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
        frame.rowconfigure(2, weight=1)
        frame.columnconfigure(0, weight=1)

        lbl_titulo = tk.Label(frame, text=titulo, font=("Segoe UI", 11, "bold"), fg=cor)
        lbl_titulo.grid(row=0, column=0)

        lbl_qtd = tk.Label(frame, text="0", font=("Segoe UI", 28, "bold"), fg=cor)
        lbl_qtd.grid(row=1, column=0, pady=(0, 8))

        texto = tk.Text(frame, height=18, font=("Consolas", 10), wrap="none", bd=0)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=texto.yview)
        texto.configure(yscrollcommand=scroll.set)
        texto.grid(row=2, column=0, sticky="nsew")
        scroll.grid(row=2, column=1, sticky="ns")
        return frame, lbl_titulo, lbl_qtd, texto

    def _atualizar_resultados_visuais(self, faltantes, extras, presentes):
        self.lbl_faltantes_qtd.configure(text=str(len(faltantes)))
        self.lbl_extras_qtd.configure(text=str(len(extras)))
        self.lbl_presentes_qtd.configure(text=str(len(presentes)))
        self._set_texto(self.txt_faltantes, self._texto_lista(faltantes, "Nenhum faltante."))
        self._set_texto(self.txt_extras, self._texto_lista(extras, "Nenhum extra."))
        self._set_texto(self.txt_presentes, self._texto_lista(presentes, "Nenhum item confirmado."))

    def _texto_lista(self, itens, texto_vazio):
        if not itens:
            return texto_vazio
        contagem = Counter(itens)
        linhas = []
        for item in sorted(contagem):
            qtd = contagem[item]
            linhas.append(item if qtd == 1 else f"{item}    x{qtd}")
        return "\n".join(linhas)

    def _atualizar_banner(self, texto, tipo):
        c = self._cores()
        cores = {
            "ok": ("#0f5132", "#d1fae5"),
            "erro": ("#7f1d1d", "#fee2e2"),
            "neutro": (c["container"], c["texto"]),
        }
        bg, fg = cores.get(tipo, cores["neutro"])
        self.lbl_banner.configure(text=texto, bg=bg, fg=fg)

    def _status(self, texto):
        if self.label_status_global:
            self.label_status_global.configure(text=f" {texto}")

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

    def _botao_hover(self, btn):
        c = self._cores()
        btn.configure(bg=c["hover"], fg="white")

    def _botao_normal(self, btn, principal=False):
        c = self._cores()
        if principal:
            btn.configure(bg=c["accent"], fg="white")
        else:
            btn.configure(bg=c["container"], fg=c["texto"])


def construir_aba_conferencia(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaConferencia(parent_frame, label_status_global, paletas, obter_tema_atual)
