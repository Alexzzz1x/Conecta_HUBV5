import os
import json
import threading
import customtkinter as ctk
import requests


CAMINHO_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_chat.txt")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

MODELOS_OPENROUTER = [
    {"id": "deepseek/deepseek-chat-v3-0324:free", "nome": "DeepSeek V3 (Free)"},
    {"id": "meta-llama/llama-3.3-70b-instruct:free", "nome": "Llama 3.3 70B (Free)"},
    {"id": "google/gemma-3-27b-it:free", "nome": "Gemma 3 27B (Free)"},
    {"id": "nvidia/llama-3.1-nemotron-70b-instruct:free", "nome": "Nemotron 70B (Free)"},
    {"id": "xiaomi/mimo-v2.5", "nome": "MiMo V2.5"},
]

SYSTEM_PROMPT = """Voce e o assistente virtual do Conecta Hub 3.0, um sistema de gerenciamento de estoque e etiquetas da Conecta Empreendimentos.

Voce tem acesso as seguintes funcoes:

GERADOR DE ETIQUETAS:
- buscar_lote(lote): Busca informacoes de um lote na planilha de controle.
- gerar_doc_google(lote): Cria um documento no Google Docs com os seriais do lote e retorna o link.
- gerar_etiqueta(deposito, material, descricao, qtd, unidade, lote): Gera PDF da etiqueta para impressao.
- consultar_material(codigo): Busca descricao de um material pelo codigo.

EXTRATOR DE SERIAIS:
- extrair_seriais(texto): Extrai numeros de serie de um texto.
- copiar_seriais(): Copia os seriais extraidos para a area de transferencia.

SAP:
- login_sap(usuario): Faz login no SAP com credenciais salvas. Use APENAS: "Anderson", "Adriana", "Junior", "Pedro".
- info_sap(): Retorna informacoes do SAP atual.

GERAIS:
- listar_lotes(): Lista todos os lotes disponiveis na planilha.

REGRAS:
- Responda em portugues brasileiro, seja conciso e direto
- Use ferramentas quando o usuario pedir algo que elas podem fazer
- Se o usuario digitar apenas um lote (ex: CL03-002), busque esse lote
- Para login SAP use apenas os nomes: Anderson, Adriana, Junior, Pedro
- Ao executar uma funcao confirme em frase curta"""

TOOLS = [
    {
        "functionDeclarations": [
            {"name": "buscar_lote", "description": "Busca informacoes de um lote na planilha de controle de seriais", "parameters": {"type": "OBJECT", "properties": {"lote": {"type": "STRING", "description": "Numero do lote (ex: CL03-002)"}}, "required": ["lote"]}},
            {"name": "gerar_doc_google", "description": "Cria um documento no Google Docs com os seriais do lote e retorna o link", "parameters": {"type": "OBJECT", "properties": {"lote": {"type": "STRING", "description": "Numero do lote"}}, "required": ["lote"]}},
            {"name": "gerar_etiqueta", "description": "Gera uma etiqueta em PDF para impressao", "parameters": {"type": "OBJECT", "properties": {"deposito": {"type": "STRING", "description": "Codigo do deposito (CL03, CL04 ou CS06)"}, "material": {"type": "STRING", "description": "Codigo do material"}, "descricao": {"type": "STRING", "description": "Descricao do material"}, "qtd": {"type": "STRING", "description": "Quantidade"}, "unidade": {"type": "STRING", "description": "Unidade (PCS, M, CE, KG)"}, "lote": {"type": "STRING", "description": "Numero do lote"}}, "required": ["deposito"]}},
            {"name": "consultar_material", "description": "Busca a descricao de um material pelo codigo", "parameters": {"type": "OBJECT", "properties": {"codigo": {"type": "STRING", "description": "Codigo do material"}}, "required": ["codigo"]}},
            {"name": "extrair_seriais", "description": "Extrai numeros de serie de um texto, suportando intervalos", "parameters": {"type": "OBJECT", "properties": {"texto": {"type": "STRING", "description": "Texto contendo seriais"}}, "required": ["texto"]}},
            {"name": "listar_lotes", "description": "Lista todos os lotes disponiveis na planilha de controle", "parameters": {"type": "OBJECT", "properties": {}}},
            {"name": "login_sap", "description": "Faz login no SAP com o usuario informado", "parameters": {"type": "OBJECT", "properties": {"usuario": {"type": "STRING", "description": "Nome do usuario SAP"}}, "required": ["usuario"]}},
            {"name": "copiar_seriais", "description": "Copia os seriais extraidos para a area de transferencia", "parameters": {"type": "OBJECT", "properties": {}}},
            {"name": "info_sap", "description": "Retorna informacoes do SAP atual", "parameters": {"type": "OBJECT", "properties": {}}},
        ]
    }
]


class ChatBot:
    def __init__(self):
        self.config = self._carregar_config()
        self.historico = []
        self.contador = 0
        self.erros = 0

    def _carregar_config(self):
        padrao = {"api_key": "", "modelo": "deepseek/deepseek-chat-v3-0324:free"}
        if os.path.exists(CAMINHO_CONFIG):
            try:
                with open(CAMINHO_CONFIG, "r", encoding="utf-8") as f:
                    salva = json.loads(f.read().strip())
                    padrao.update(salva)
            except Exception:
                pass
        return padrao

    def salvar_config(self):
        with open(CAMINHO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)

    def obter_api_key(self):
        return self.config.get("api_key", "")

    def obter_modelo(self):
        return self.config.get("modelo", "deepseek/deepseek-chat-v3-0324:free")

    def _converter_tools_para_openai(self):
        tools = []
        for bloco in TOOLS:
            for func in bloco.get("functionDeclarations", []):
                tools.append({"type": "function", "function": {"name": func["name"], "description": func["description"], "parameters": func["parameters"]}})
        return tools

    def _chamar_api(self, api_key, model_id, mensagem):
        mensagens = [{"role": "system", "content": SYSTEM_PROMPT}]
        for msg in self.historico[-10:]:
            role = "assistant" if msg.get("role") == "model" else "user"
            text = ""
            for p in msg.get("parts", []):
                if "text" in p:
                    text = p["text"]
            if text:
                mensagens.append({"role": role, "content": text})
        mensagens.append({"role": "user", "content": mensagem})

        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            json={"model": model_id, "messages": mensagens, "tools": self._converter_tools_para_openai(), "tool_choice": "auto"},
            timeout=(10, 30),
        )
        resp.raise_for_status()
        return resp.json()

    def _chamar_api_rapida(self, api_key, model_id, mensagem):
        mensagens = [{"role": "system", "content": "Resuma em 1-2 frases em portugues."}, {"role": "user", "content": mensagem}]
        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            json={"model": model_id, "messages": mensagens},
            timeout=(5, 15),
        )
        resp.raise_for_status()
        return resp.json()

    def _adicionar_historico(self, role, texto):
        self.historico.append({"role": role, "parts": [{"text": texto}]})
        if len(self.historico) > 20:
            self.historico = self.historico[-20:]

    def processar_mensagem(self, mensagem, funcoes_disponiveis=None):
        api_key = self.obter_api_key()
        modelo = self.obter_modelo()
        if not api_key:
            return "Nenhuma API key configurada. Clique no ⚙ para configurar."
        try:
            resposta = self._chamar_api(api_key, modelo, mensagem)
            return self._processar_resposta(resposta, mensagem, funcoes_disponiveis)
        except requests.exceptions.HTTPError as e:
            self.erros += 1
            cod = e.response.status_code if e.response else 0
            if cod == 429:
                return "Limite de requisicoes atingido. Aguarde um momento."
            elif cod in (401, 403):
                return "Chave API invalida. Verifique no ⚙."
            elif cod == 404:
                return f"Modelo '{modelo}' nao encontrado. Troque no ⚙."
            return f"Erro {cod}: {str(e)[:200]}"
        except requests.exceptions.ConnectionError:
            self.erros += 1
            return "Erro de conexao. Verifique sua internet."
        except Exception as e:
            self.erros += 1
            return f"Erro: {str(e)[:100]}"

    def _processar_resposta(self, resposta, mensagem, funcoes_disponiveis):
        choices = resposta.get("choices", [])
        if not choices:
            return "Nao consegui processar sua mensagem."
        msg = choices[0].get("message", {})
        self.contador += 1

        if msg.get("tool_calls"):
            for tc in msg["tool_calls"]:
                func_info = tc.get("function", {})
                nome_func = func_info.get("name", "")
                try:
                    args = json.loads(func_info.get("arguments", "{}"))
                except Exception:
                    args = {}
                resultado = self._executar_funcao(nome_func, args, funcoes_disponiveis)
                resultado_str = json.dumps(resultado, ensure_ascii=False)
                self._adicionar_historico("user", mensagem)
                self._adicionar_historico("model", f"[Chamou funcao {nome_func}]")

                if "erro" in resultado:
                    return f"Erro ao executar {nome_func}: {resultado['erro']}"

                try:
                    fu = self._chamar_api_rapida(self.obter_api_key(), self.obter_modelo(), f"Funcao {nome_func} executada. Resultado: {resultado_str}")
                    fu_text = fu.get("choices", [{}])[0].get("message", {}).get("content", "")
                    if fu_text:
                        self._adicionar_historico("model", fu_text)
                        return fu_text
                except Exception:
                    pass

                resumo = resultado.get("status", "") or resultado.get("descricao", "") or resultado_str[:200]
                return f"✓ {nome_func}: {resumo}"

        content = msg.get("content", "")
        if content:
            self._adicionar_historico("user", mensagem)
            self._adicionar_historico("model", content)
            return content
        return "Nao consegui gerar uma resposta."

    def _executar_funcao(self, nome, args, funcoes):
        if not funcoes:
            return {"erro": "Funcoes do programa nao disponiveis"}
        func = funcoes.get(nome)
        if not func:
            return {"erro": f"Funcao '{nome}' nao encontrada"}
        try:
            return func(**args)
        except Exception as e:
            return {"erro": str(e)}


class JanelaChat:
    def __init__(self, root, paletas, obter_tema_atual, funcoes_disponiveis=None):
        self.root = root
        self.paletas = paletas
        self.obter_tema_atual = obter_tema_atual
        self.funcoes_disponiveis = funcoes_disponiveis or {}
        self.chatbot = ChatBot()

        self.janela = ctk.CTkToplevel(root)
        self.janela.title("Conecta Hub Chat")
        self.janela.geometry("440x650")
        self.janela.minsize(380, 450)
        self.janela.protocol("WM_DELETE_WINDOW", self._fechar)

        self._construir_interface()
        self._enviar_boas_vindas()

    def _construir_interface(self):
        frame_topo = ctk.CTkFrame(self.janela, fg_color="#1a1a2e", corner_radius=0)
        frame_topo.pack(fill="x")

        ctk.CTkLabel(frame_topo, text="Conecta Hub Chat", font=("Segoe UI", 15, "bold"), text_color="#00d4ff").pack(side="left", padx=12, pady=10)

        ctk.CTkButton(frame_topo, text="⚙", width=32, height=32, font=("Segoe UI", 14), fg_color="transparent", hover_color="#2a2a4a", command=self._abrir_config).pack(side="right", padx=10)

        self.frame_mensagens = ctk.CTkScrollableFrame(self.janela, fg_color="#0f0f23", corner_radius=0)
        self.frame_mensagens.pack(fill="both", expand=True, padx=0, pady=0)

        frame_input = ctk.CTkFrame(self.janela, fg_color="#1a1a2e", corner_radius=0)
        frame_input.pack(fill="x", side="bottom")

        self.entry_msg = ctk.CTkEntry(frame_input, placeholder_text="Digite sua mensagem...", font=("Segoe UI", 12), height=40, fg_color="#16213e", border_color="#00d4ff", text_color="white")
        self.entry_msg.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        self.entry_msg.bind("<Return>", self._enviar_mensagem)

        self.btn_enviar = ctk.CTkButton(frame_input, text="➤", width=40, height=40, font=("Segoe UI", 14, "bold"), fg_color="#00d4ff", hover_color="#0099cc", text_color="#0f0f23", command=self._enviar_mensagem)
        self.btn_enviar.pack(side="right", padx=(0, 10), pady=10)

        self.frame_status = ctk.CTkFrame(self.janela, fg_color="#1a1a2e", corner_radius=0, height=28)
        self.frame_status.pack(fill="x", side="bottom")
        self.frame_status.pack_propagate(False)
        self._atualizar_status()

    def _atualizar_status(self):
        for w in self.frame_status.winfo_children():
            w.destroy()
        api_key = self.chatbot.obter_api_key()
        modelo = self.chatbot.obter_modelo()
        if api_key:
            modelo_nome = ""
            for m in MODELOS_OPENROUTER:
                if m["id"] == modelo:
                    modelo_nome = m["nome"]
                    break
            cor = "#00ff88" if self.chatbot.erros < 3 else "#ff4444"
            status = "●" if self.chatbot.erros < 3 else "●"
            texto = f"{status} {modelo_nome} | {self.chatbot.contador} req"
        else:
            cor = "#ff4444"
            texto = "● API nao configurada — clique ⚙"
        ctk.CTkLabel(self.frame_status, text=texto, font=("Segoe UI", 10), text_color=cor).pack(side="left", padx=10, pady=4)

    def _enviar_boas_vindas(self):
        api_key = self.chatbot.obter_api_key()
        if not api_key:
            self._adicionar_mensagem_bot(
                "Bem-vindo ao Conecta Hub Chat!\n\n"
                "Clique no ⚙ e cole sua API key do OpenRouter.\n"
                "Gratuita em: https://openrouter.ai/keys\n\n"
                "Modelos disponiveis:\n"
                "• DeepSeek V3 (Free)\n"
                "• Llama 3.3 70B (Free)\n"
                "• Gemma 3 27B (Free)\n"
                "• Nemotron 70B (Free)\n"
                "• MiMo V2.5"
            )
        else:
            modelo = self.chatbot.obter_modelo()
            modelo_nome = ""
            for m in MODELOS_OPENROUTER:
                if m["id"] == modelo:
                    modelo_nome = m["nome"]
                    break
            self._adicionar_mensagem_bot(
                f"Modelo ativo: {modelo_nome}\n\n"
                f"Posso ajudar com:\n"
                f"• Buscar informacoes de lotes\n"
                f"• Gerar documentos no Google Drive\n"
                f"• Consultar materiais\n"
                f"• Gerar etiquetas (PDF)\n"
                f"• Extrair seriais\n"
                f"• Login no SAP\n\n"
                f"Como posso ajudar?"
            )

    def _adicionar_mensagem_bot(self, texto):
        frame = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=4)

        bubble = ctk.CTkFrame(frame, fg_color="#16213e", corner_radius=12)
        bubble.pack(anchor="w", fill="x")

        ctk.CTkLabel(bubble, text="🤖", font=("Segoe UI", 16), text_color="#00d4ff").pack(anchor="w", padx=10, pady=(8, 0))
        ctk.CTkLabel(bubble, text=texto, font=("Segoe UI", 11), text_color="#e0e0e0", wraplength=340, justify="left", anchor="w").pack(anchor="w", fill="x", padx=10, pady=(2, 10))

        self.janela.after(50, lambda: self.frame_mensagens._parent_canvas.yview_moveto(1.0))

    def _mostrar_pensando(self):
        self.frame_pensando = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        self.frame_pensando.pack(fill="x", padx=10, pady=4)

        bubble = ctk.CTkFrame(self.frame_pensando, fg_color="#16213e", corner_radius=12)
        bubble.pack(anchor="w", fill="x")

        ctk.CTkLabel(bubble, text="🤖", font=("Segoe UI", 16), text_color="#00d4ff").pack(anchor="w", padx=10, pady=(8, 0))
        self.lbl_pensando = ctk.CTkLabel(bubble, text="Digitando", font=("Segoe UI", 11, "italic"), text_color="#888888")
        self.lbl_pensando.pack(anchor="w", padx=10, pady=(2, 10))

        self._animacao_pensando = 0
        self._animar_pensando()

    def _animar_pensando(self):
        if not hasattr(self, 'frame_pensando') or not self.frame_pensando.winfo_exists():
            return
        pontos = "." * (self._animacao_pensando % 4)
        self.lbl_pensando.configure(text=f"Digitando{pontos}")
        self._animacao_pensando += 1
        self.janela.after(400, self._animar_pensando)

    def _remover_pensando(self):
        if hasattr(self, 'frame_pensando') and self.frame_pensando.winfo_exists():
            self.frame_pensando.destroy()

    def _adicionar_mensagem_usuario(self, texto):
        frame = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=4)

        bubble = ctk.CTkFrame(frame, fg_color="#00d4ff", corner_radius=12)
        bubble.pack(anchor="e")

        ctk.CTkLabel(bubble, text=texto, font=("Segoe UI", 11), text_color="#0f0f23", wraplength=340, justify="left", anchor="w").pack(padx=10, pady=8)

        self.janela.after(50, lambda: self.frame_mensagens._parent_canvas.yview_moveto(1.0))

    def _enviar_mensagem(self, event=None):
        texto = self.entry_msg.get().strip()
        if not texto:
            return
        self.entry_msg.delete(0, "end")
        self._adicionar_mensagem_usuario(texto)
        self.btn_enviar.configure(state="disabled")
        self.entry_msg.configure(state="disabled")
        self._mostrar_pensando()
        thread = threading.Thread(target=self._processar_em_thread, args=(texto,), daemon=True)
        thread.start()

    def _processar_em_thread(self, texto):
        try:
            resposta = self.chatbot.processar_mensagem(texto, self.funcoes_disponiveis)
        except Exception as e:
            resposta = f"Erro inesperado: {str(e)}"
        finally:
            try:
                self.janela.after(0, lambda: self._remover_pensando())
                self.janela.after(50, lambda r=resposta: self._adicionar_mensagem_bot(r))
                self.janela.after(100, lambda: self._atualizar_status())
                self.janela.after(150, lambda: self.btn_enviar.configure(state="normal"))
                self.janela.after(150, lambda: self.entry_msg.configure(state="normal"))
            except Exception:
                pass

    def _abrir_config(self):
        janela_cfg = ctk.CTkToplevel(self.janela)
        janela_cfg.title("Configuracao — OpenRouter")
        janela_cfg.geometry("420x320")
        janela_cfg.configure(fg_color="#0f0f23")
        janela_cfg.transient(self.janela)
        janela_cfg.grab_set()

        ctk.CTkLabel(janela_cfg, text="OpenRouter API", font=("Segoe UI", 16, "bold"), text_color="#00d4ff").pack(pady=(20, 5))
        ctk.CTkLabel(janela_cfg, text="Pegue sua key gratis em:\nhttps://openrouter.ai/keys", font=("Segoe UI", 10), text_color="#888888").pack()

        ctk.CTkLabel(janela_cfg, text="API Key:", font=("Segoe UI", 11, "bold"), text_color="white").pack(pady=(15, 3))
        entry_key = ctk.CTkEntry(janela_cfg, font=("Consolas", 11), width=350, height=36, fg_color="#16213e", border_color="#00d4ff")
        entry_key.pack()
        entry_key.insert(0, self.chatbot.obter_api_key())

        ctk.CTkLabel(janela_cfg, text="Modelo:", font=("Segoe UI", 11, "bold"), text_color="white").pack(pady=(10, 3))
        modelo_var = ctk.StringVar(value=self.chatbot.obter_modelo())
        opcoes = [m["id"] for m in MODELOS_OPENROUTER]
        ctk.CTkOptionMenu(janela_cfg, variable=modelo_var, values=opcoes, font=("Segoe UI", 10), width=350, fg_color="#16213e", button_color="#00d4ff", button_hover_color="#0099cc").pack()

        def salvar():
            self.chatbot.config["api_key"] = entry_key.get().strip()
            self.chatbot.config["modelo"] = modelo_var.get()
            self.chatbot.salvar_config()
            self.chatbot.contador = 0
            self.chatbot.erros = 0
            self._atualizar_status()
            janela_cfg.destroy()

        ctk.CTkButton(janela_cfg, text="Salvar", font=("Segoe UI", 11, "bold"), fg_color="#00d4ff", hover_color="#0099cc", text_color="#0f0f23", width=200, height=36, command=salvar).pack(pady=20)

    def _fechar(self):
        self.janela.destroy()

    def aplicar_tema(self, tema_nome):
        pass
