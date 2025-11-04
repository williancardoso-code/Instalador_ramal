# ====================== IMPORTS ======================
import os, sys, time, shutil, tempfile, subprocess, threading
import tkinter as tk
from tkinter import ttk, messagebox
import requests
from pywinauto import Application, Desktop
import pyautogui
from PIL import Image, ImageTk
import time

# ====================== CONSTANTES ======================
USUARIO_API_PADRAO = "contatowilliancardoso@gmail.com"
TOKEN_FIXO = "6ddf7ebb-8d57-4d60-b63c-87cfcb853ace"
LISTAR_RAMAIS_URL = "https://novosistema.atendas.com.br/suite/api/listar_ramais"

SENHA_INICIAL = "Atendas@*2510"
SERVIDOR_PADRAO = "novosistema.atendas.com.br"
APP_EXEC = r"%LOCALAPPDATA%\Atendas\Atendas.exe"

BASE_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(__file__)
ATENDAS_EXE = os.path.join(BASE_DIR, "Atendas-3.21.6.exe")
IMG_MENU = os.path.join(BASE_DIR, "imagens_config", "menu_tres_riscos.png")
LOGO_PNG = os.path.join(BASE_DIR, "imagens_config", "logo-atendas.png")
ICON_PNG = os.path.join(BASE_DIR, "imagens_config", "ponto.png")

DARK_BG = "#1F2630"
FG_TXT  = "#E8EEF7"
WARN_YELLOW = "#FFD700"

# ====================== API ======================
def listar_ramais_api(usuario_api: str, token: str, cliente_id: int, pos_inicio=0):
    headers = {"usuario": usuario_api, "token": token, "Accept": "application/json"}
    params = {"cliente_id": int(cliente_id), "pos_registro_inicial": int(pos_inicio)}
    r = requests.get(LISTAR_RAMAIS_URL, headers=headers, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    if not isinstance(data, dict) or "dados" not in data:
        raise RuntimeError(f"JSON inesperado: {str(data)[:400]}")
    return data["dados"]

# ====================== INSTALAÇÃO / ABERTURA ======================
def instalar_atendas_silencioso():
    if not os.path.exists(ATENDAS_EXE):
        messagebox.showerror("Instalador não encontrado", f"Não achei: {ATENDAS_EXE}")
        return False
    try:
        tmp = os.path.join(tempfile.gettempdir(), os.path.basename(ATENDAS_EXE))
        shutil.copyfile(ATENDAS_EXE, tmp)
        print("🛠 Instalando Atendas silenciosamente...")
        subprocess.run([tmp, "/S"], check=True)
        time.sleep(6)
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao instalar Atendas:\n{e}")
        return False

def abrir_atendas():
    os.startfile(os.path.expandvars(APP_EXEC))
    print("⏳ Abrindo Atendas…")
    time.sleep(5)

def conectar_app():
    return Application(backend="uia").connect(path="Atendas.exe")

def resolver_janela_principal(app: Application):
    win = app.window(title_re=r"^Atendas.*")
    win.wait("exists enabled visible", timeout=5)
    return win

# ====================== SUPORTE DE AUTOMACAO ======================
def abrir_menu_tres_riscos(app):
    janela = resolver_janela_principal(app)
    print("📂 Abrindo menu…")
    try:
        btn = janela.child_window(control_type="Button", found_index=0)
        btn.click_input()
        time.sleep(0.4)
        return True
    except:
        pass
    try:
        pos = pyautogui.locateCenterOnScreen(IMG_MENU, confidence=0.75)
        if pos:
            pyautogui.click(pos)
            time.sleep(0.4)
            return True
    except:
        pass
    print("⚠️ Não consegui abrir o menu.")
    return False

def _menu_tem_item(app, pattern_regex):
    try:
        popup = app.window(control_type="Menu")
        item = popup.child_window(title_re=pattern_regex, control_type="MenuItem")
        item.wrapper_object()
        return True
    except Exception:
        return False

def _click_menu_item(app, pattern_regex):
    for _ in range(2):
        try:
            popup = app.window(control_type="Menu")
            item = popup.child_window(title_re=pattern_regex, control_type="MenuItem")
            item.click_input()
            return True
        except Exception:
            abrir_menu_tres_riscos(app)
    return False

# ====================== SENHA (corrigida) ======================
def preencher_senha_apos_acao(app, acao_click, senha=SENHA_INICIAL, tentativas=16, pausa=0.25):
    print("🔐 Preenchendo senha inicial…")
    janela = resolver_janela_principal(app)
    main_handle = janela.handle
    before = set(w.handle for w in Desktop(backend="uia").windows(process=app.process) if w.is_visible())

    if not acao_click():
        print("❌ Ação de clique no menu falhou.")
        return False

    # ✅ NOVA LINHA: pequena pausa para o popup renderizar completamente
    time.sleep(1.5)

    # Tentativa principal
    for _ in range(tentativas):
        for w in Desktop(backend="uia").windows(process=app.process):
            if not w.is_visible() or w.handle in before:
                continue
            edits = [e for e in w.descendants(control_type="Edit") if e.is_visible()]
            if edits:
                try:
                    edt = edits[0]
                    edt.type_keys("^a{BACKSPACE}", pause=0.01)
                    edt.type_keys(senha, with_spaces=True, pause=0.01)
                    time.sleep(0.2)
                    w.type_keys("{ENTER}")
                    print("✅ Senha enviada (popup detectado).")
                    return True
                except:
                    pass
        time.sleep(pausa)

    # Fallback universal (campo na janela principal)
    print("⚙️ Tentando fallback de senha (modo amplo)…")
    for _ in range(12):
        for w in Desktop(backend="uia").windows(process=app.process):
            if not w.is_visible():
                continue
            try:
                edits = [e for e in w.descendants(control_type="Edit") if e.is_visible()]
            except Exception:
                edits = []
            for edt in edits:
                try:
                    edt.type_keys("^a{BACKSPACE}", pause=0.01)
                    edt.type_keys(senha, with_spaces=True, pause=0.01)
                    time.sleep(0.2)
                    pyautogui.press("enter")
                    print("✅ Senha enviada (fallback amplo).")
                    return True
                except Exception:
                    continue
        time.sleep(pausa)

    print("⚠️ Nenhum campo de senha encontrado.")
    return False

# ====================== CAMPOS / SALVAR ======================
def preencher_campos(app, cred):
    janela = resolver_janela_principal(app)
    campos = [
        (("Nome da Conta","Nome da Conta"), cred.get("nome_conta") or cred["login"]),
        (("Nome de Exibição","Nome de Exibicao"), cred.get("nome_exibicao") or cred["login"]),
        (("Usuário","Usuario"), cred["login"]),
        (("Login","Login"), cred["login"]),
        (("Senha","Senha"), cred["senha"]),
        (("Servidor SIP","Servidor SIP"), cred.get("servidor_sip")),
        (("Proxy SIP","Proxy SIP"), cred.get("proxy_sip")),
        (("Domínio","Dominio"), cred.get("dominio")),
    ]
    print("✏️ Preenchendo campos…")
    for (t1, t2), valor in campos:
        for t in (t1, t2):
            try:
                janela.child_window(title=t, control_type="Edit").set_text(valor)
                break
            except:
                continue
    return True


def salvar_confirmar(app):
    janela = resolver_janela_principal(app)
    try:
        janela.child_window(title="Salvar", control_type="Button").click_input()
        print("💾 Salvando configuração...")
        time.sleep(0.5)

        try:
            janela.child_window(title="OK", control_type="Button").click_input()
            print("🔘 Botão OK clicado.")
        except:
            pass

        # Mensagem aparece exatamente após o salvamento real
        messagebox.showinfo("Concluído", "✅ Ramal configurado com sucesso!")
        print("🎉 Configuração concluída!")
        return True
    except Exception as e:
        print(f"⚠️ Erro ao salvar/confirmar: {e}")
        return False


# ====================== GUI ======================
class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Instalador Atendas")
        self.configure(bg=DARK_BG)
        self.geometry("450x240")
        self.resizable(False, False)

        if os.path.exists(ICON_PNG):
            self.iconphoto(False, tk.PhotoImage(file=ICON_PNG))

        header = tk.Frame(self, bg=DARK_BG)
        header.pack(pady=(20, 10))
        if os.path.exists(LOGO_PNG):
            img = Image.open(LOGO_PNG)
            img = img.resize((260, int(260 * img.height / img.width)), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(header, image=self.logo_img, bg=DARK_BG).pack()

        body = tk.Frame(self, bg=DARK_BG)
        body.pack(pady=(0, 10))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TLabel", background=DARK_BG, foreground=FG_TXT, font=("Segoe UI", 10))
        style.configure("Dark.TEntry", fieldbackground="#2A3340", foreground=FG_TXT, insertcolor=FG_TXT)
        style.configure("Dark.TButton", font=("Segoe UI", 10, "bold"))

        ttk.Label(body, text="Digite seu código:", style="Dark.TLabel").grid(row=0, column=0, padx=(0, 10))
        self.e_codigo = ttk.Entry(body, width=28, style="Dark.TEntry")
        self.e_codigo.grid(row=0, column=1)
        self.btn_install = ttk.Button(body, text="Instalar", style="Dark.TButton", command=self.on_install)
        self.btn_install.grid(row=0, column=2, padx=(10, 0))

        self.status_lbl = tk.Label(self, text="", bg=DARK_BG, fg=FG_TXT, font=("Segoe UI", 11, "bold"))
        self.status_lbl.pack(pady=(8, 0))

    def exibir_alerta_permissao(self):
        alerta = tk.Toplevel(self)
        alerta.title("Atenção – Permissão do Windows")
        alerta.geometry("460x180")
        alerta.configure(bg=DARK_BG)
        tk.Label(alerta, text="⚠️ NÃO MEXA NO MOUSE DURANTE A INSTALAÇÃO",
                 fg=WARN_YELLOW, bg=DARK_BG, font=("Segoe UI", 11, "bold")).pack(pady=(20,10))
        tk.Label(alerta, text="Ao clicar em 'Permitir' na permissão do Windows,\naguarde até o fim da instalação automática.",
                 fg=FG_TXT, bg=DARK_BG, wraplength=420, justify="center").pack(pady=(5,20))
        ttk.Button(alerta, text="Prosseguir", command=alerta.destroy).pack()
        alerta.grab_set()
        alerta.wait_window()

    def _encerrar(self):
        try:
            self.destroy()
        finally:
            sys.exit(0)

    def _buscar_ramal(self, cliente_id: int, ramal_id: int):
        pos = 0
        while pos <= 80:
            dados = listar_ramais_api(USUARIO_API_PADRAO, TOKEN_FIXO, cliente_id, pos)
            if not dados:
                break
            for d in dados:
                try:
                    rid = int(d.get("ramal_id", -1))
                except:
                    rid = -1
                if rid == int(ramal_id):
                    return d
            if len(dados) < 20:
                break
            pos += 20
        return None

    def _executar_instalacao(self, cliente_id, ramal_id):
        self._animando = True
        threading.Thread(target=self._animar_carregamento, daemon=True).start()

        try:
            d = self._buscar_ramal(cliente_id, ramal_id)
            if not d:
                messagebox.showwarning("Não encontrado", f"Ramal {ramal_id} não localizado para a unidade {cliente_id}.")
                return

            login = d.get("usuario_autenticacao") or d.get("numero")
            senha = d.get("senha_sip") or SENHA_INICIAL

            cred = {
                "login": login,
                "senha": senha,
                "nome_conta": d.get("nome") or login,
                "nome_exibicao": d.get("nome") or login,
                "dominio": SERVIDOR_PADRAO,
                "servidor_sip": SERVIDOR_PADRAO,
                "proxy_sip": SERVIDOR_PADRAO,
            }

            if not instalar_atendas_silencioso():
                return
            abrir_atendas()
            app = conectar_app()

            if not abrir_menu_tres_riscos(app):
                return

            if _menu_tem_item(app, r".*Editar Conta.*"):
                resp = messagebox.askyesno("Ramal já configurado",
                    f"Já existe um ramal configurado nesta máquina.\n\nDeseja substituir pelas credenciais de '{login}'?")
                if not resp:
                    messagebox.showinfo("Cancelado", "Operação cancelada. Mantendo o ramal atual.")
                    return
                ok_pwd = preencher_senha_apos_acao(app, acao_click=lambda: _click_menu_item(app, r".*Editar Conta.*"))
            else:
                ok_pwd = preencher_senha_apos_acao(app, acao_click=lambda: _click_menu_item(app, r".*Adicionar Conta.*"))

            if not ok_pwd:
                return

            preencher_campos(app, cred)
            salvar_confirmar(app)

        finally:
            self._animando = False
            self.status_lbl.config(text="")
            self.after(400, self._encerrar)

    def _animar_carregamento(self):
        pontos = ["", ".", "..", "..."]
        i = 0
        while getattr(self, "_animando", False):
            self.status_lbl.config(text=f"⏳ Instalando, aguarde{pontos[i % len(pontos)]}")
            i += 1
            time.sleep(0.5)

    def on_install(self):
        codigo = self.e_codigo.get().strip()
        if not codigo or "-" not in codigo:
            messagebox.showwarning("Atenção", "Formato inválido.\nExemplo: 123-456")
            return
        cliente_id, ramal_id = map(int, codigo.split("-", 1))
        self.exibir_alerta_permissao()
        self.status_lbl.config(text="⏳ Instalando, aguarde")
        threading.Thread(target=self._executar_instalacao, args=(cliente_id, ramal_id), daemon=True).start()

# ====================== MAIN ======================
if __name__ == "__main__":
    AppGUI().mainloop()
