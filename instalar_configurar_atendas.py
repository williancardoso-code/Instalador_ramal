# ====================== IMPORTS ======================
import os, sys, time, shutil, tempfile, subprocess, re
import tkinter as tk
from tkinter import ttk, messagebox
import requests
from pywinauto import Application, Desktop
import pyautogui
import ctypes

# ====================== CONSTANTES ======================
USUARIO_API_PADRAO = "contatowilliancardoso@gmail.com"   # oculto na UI
TOKEN_FIXO = "6ddf7ebb-8d57-4d60-b63c-87cfcb853ace"      # oculto na UI
LISTAR_RAMAIS_URL = "https://novosistema.atendas.com.br/suite/api/listar_ramais"

SENHA_INICIAL = "Atendas@*2510"
SERVIDOR_PADRAO = "novosistema.atendas.com.br"
APP_EXEC = r"%LOCALAPPDATA%\Atendas\Atendas.exe"

BASE_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(__file__)
ATENDAS_EXE = os.path.join(BASE_DIR, "Atendas-3.21.6.exe")
IMG_MENU = os.path.join(BASE_DIR, "imagens_config", "menu_tres_riscos.png")
LOGO_PNG = os.path.join(BASE_DIR, "imagens_config", "logo-atendas.png")

DARK_BG = "#1F2630"   # cinza-azulado escuro
FG_TXT  = "#E8EEF7"

# ====================== FECHAR APP (NOVO) ======================
def finalizar_e_sair(janela_tk, delay_ms=400):
    """Fecha a janela principal do instalador após um pequeno delay (ms)."""
    try:
        janela_tk.after(delay_ms, janela_tk.destroy)
    except Exception:
        try:
            janela_tk.destroy()
        except Exception:
            pass

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
        time.sleep(8)
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao instalar Atendas:\n{e}")
        return False

def abrir_atendas():
    os.startfile(os.path.expandvars(APP_EXEC))
    print("⏳ Abrindo Atendas…")
    time.sleep(5)

def conectar_janela_principal():
    app = Application(backend="uia").connect(path="Atendas.exe")
    janela = app.window(best_match="Atendas")
    return app, janela

# ====================== SUPORTE: VERIFICAÇÃO / MENU ======================
def extrair_login_do_titulo(janela):
    try:
        titulo = (janela.window_text() or "").strip()
    except Exception:
        return None
    m = re.match(r"^Atendas\s*-\s*([^\s]+)", titulo)
    return m.group(1) if m else None

def verificar_conta_existente(janela):
    login = extrair_login_do_titulo(janela)
    return (login is not None), login

def abrir_menu_tres_riscos(janela):
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

def clicar_adicionar_conta(app):
    print("➕ Selecionando 'Adicionar Conta…'")
    try:
        popup = app.window(control_type="Menu")
        popup.child_window(title_re=".*Adicionar Conta.*", control_type="MenuItem").click_input()
        time.sleep(0.6)
        return True
    except Exception as e:
        print(f"⚠️ Não encontrei 'Adicionar Conta…': {e}")
        return False

def clicar_editar_conta(app):
    print("📝 Selecionando 'Editar Conta…'")
    try:
        popup = app.window(control_type="Menu")
        for pat in [".*Editar Conta.*", ".*Gerenciar Conta.*", ".*Configurar Conta.*", ".*Conta.*Editar.*"]:
            try:
                popup.child_window(title_re=pat, control_type="MenuItem").click_input()
                time.sleep(0.6)
                return True
            except Exception:
                continue
    except Exception as e:
        print(f"⚠️ Não encontrei 'Editar Conta…': {e}")
    return False

# ====================== SENHA (robusto) ======================
def preencher_qualquer_senha(app, janela, senha=SENHA_INICIAL):
    print("🔐 Preenchendo senha…")
    main_handle = janela.handle
    proc = app.process

    # 1) popups
    for w in Desktop(backend="uia").windows(process=proc):
        if not w.is_visible() or w.handle == main_handle:
            continue
        try:
            edits = [e for e in w.descendants(control_type="Edit") if e.is_visible()]
        except Exception:
            edits = []
        if edits:
            try:
                edt = edits[0]
                try:
                    edt.set_text(senha)
                except Exception:
                    edt.type_keys("^a{BACKSPACE}", pause=0.01)
                    edt.type_keys(senha, with_spaces=True, pause=0.01)
                time.sleep(0.15)
                try:
                    w.type_keys("{ENTER}")
                except Exception:
                    pyautogui.press("enter")
                time.sleep(0.6)
                print("✅ Senha enviada (popup).")
                return True
            except Exception:
                pass

    # 2) janela principal
    try:
        edits = [e for e in janela.descendants(control_type="Edit") if e.is_visible()]
    except Exception:
        edits = []
    if edits:
        try:
            edt = edits[0]
            try:
                edt.set_text(senha)
            except Exception:
                edt.type_keys("^a{BACKSPACE}", pause=0.01)
                edt.type_keys(senha, with_spaces=True, pause=0.01)
            time.sleep(0.15)
            try:
                janela.type_keys("{ENTER}")
            except Exception:
                pyautogui.press("enter")
            time.sleep(0.6)
            print("✅ Senha enviada (janela).")
            return True
        except Exception:
            pass

    print("⚠️ Não achei campo de senha.")
    return False

def inserir_senha_inicial_e_enter(janela):
    print("🔑 Inserindo senha inicial…")
    try:
        caixa = janela.child_window(control_type="Edit", found_index=0)
        caixa.set_text(SENHA_INICIAL)
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.6)
        return True
    except:
        pass
    try:
        app = Application(backend="uia").connect(path="Atendas.exe")
        return preencher_qualquer_senha(app, janela, SENHA_INICIAL)
    except Exception:
        print("⚠️ Não consegui inserir a senha inicial.")
        return False

# ====================== CAMPOS / SALVAR ======================
def preencher_campos(janela, cred):
    nome_conta   = cred.get("nome_conta") or cred["login"]
    nome_exib    = cred.get("nome_exibicao") or cred["login"]
    usuario      = cred["login"]
    login_val    = cred["login"]
    senha_val    = cred["senha"]
    dominio      = cred.get("dominio", SERVIDOR_PADRAO)
    servidor_sip = cred.get("servidor_sip", dominio)
    proxy_sip    = cred.get("proxy_sip", dominio)

    campos = [
        (("Nome da Conta","Nome da Conta"),           nome_conta),
        (("Nome de Exibição","Nome de Exibicao"),     nome_exib),
        (("Usuário","Usuario"),                        usuario),
        (("Login","Login"),                            login_val),
        (("Senha","Senha"),                            senha_val),
        (("Servidor SIP","Servidor SIP"),              servidor_sip),
        (("Proxy SIP","Proxy SIP"),                    proxy_sip),
        (("Domínio","Dominio"),                        dominio),
    ]
    print("✏️ Preenchendo campos…")
    for (t1, t2), valor in campos:
        ok = False
        for t in (t1, t2):
            try:
                janela.child_window(title=t, control_type="Edit").set_text(valor)
                print(f"   • {t}: OK")
                ok = True
                break
            except:
                continue
        if not ok:
            print(f"⚠️ Não consegui preencher '{t1}'")

def salvar_confirmar(janela):
    try:
        janela.child_window(title="Salvar", control_type="Button").click_input()
        time.sleep(0.4)
        try:
            janela.child_window(title="OK", control_type="Button").click_input()
        except:
            pass
        print("🎉 Configuração concluída!")
        return True
    except:
        print("⚠️ Erro ao salvar/confirmar.")
        return False

# ====================== GUI – Instalação por CÓDIGO ======================
class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        # Janela principal
        self.title("")  # sem texto no título
        self.configure(bg=DARK_BG)
        self.geometry("640x320")
        self.resizable(False, False)

        # ==== Ícone personalizado ====
        try:
            icon_path = os.path.join(BASE_DIR, "imagens_config", "ponto.png")
            if os.path.exists(icon_path):
                self.iconphoto(False, tk.PhotoImage(file=icon_path))
        except Exception as e:
            print(f"⚠️ Não consegui aplicar o ícone: {e}")

        # ---- Cabeçalho só com a LOGO ----
        header = tk.Frame(self, bg=DARK_BG, height=70)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)  # mantém altura

        self.logo_img = None
        if os.path.exists(LOGO_PNG):
            try:
                img = tk.PhotoImage(file=LOGO_PNG)
                if img.height() > 48:
                    scale = max(1, img.height() // 48)
                    img = img.subsample(scale, scale)
                self.logo_img = img
                tk.Label(header, image=self.logo_img, bg=DARK_BG).pack(side="left", padx=14, pady=8)
            except Exception as e:
                messagebox.showwarning("Logo", f"Falha ao carregar logo:\n{e}")

        # ---- Corpo do formulário ----
        body = tk.Frame(self, bg=DARK_BG)
        body.pack(fill="both", expand=True, padx=20, pady=(6, 8))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Dark.TLabel", background=DARK_BG, foreground=FG_TXT, font=("Segoe UI", 10))
        style.configure("Dark.TEntry", fieldbackground="#2A3340", foreground=FG_TXT, insertcolor=FG_TXT)
        style.configure("Dark.TButton", font=("Segoe UI", 10, "bold"))

        ttk.Label(body, text="Digite seu código:", style="Dark.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 10))
        self.e_codigo = ttk.Entry(body, width=28, style="Dark.TEntry")
        self.e_codigo.grid(row=0, column=1, sticky="w")
        self.e_codigo.focus_set()

        self.btn_install = ttk.Button(body, text="Instalar", style="Dark.TButton", command=self.on_install)
        self.btn_install.grid(row=0, column=2, padx=(16, 0))

    # ---- helpers ----
    def _parse_codigo(self, s: str):
        s = (s or "").strip()
        if not s:
            return None
        if "-" in s:
            a, b = s.split("-", 1)
        elif "," in s:
            a, b = s.split(",", 1)
        elif " " in s:
            a, b = s.split(None, 1)
        else:
            return None
        try:
            return int(a), int(b)
        except:
            return None

    def _buscar_ramal(self, cliente_id: int, ramal_id: int):
        dados = listar_ramais_api(USUARIO_API_PADRAO, TOKEN_FIXO, cliente_id, 0)
        for d in dados:
            if int(d.get("ramal_id", -1)) == int(ramal_id):
                return d
        return None

    # ---- ação principal ----
    def on_install(self):
        parsed = self._parse_codigo(self.e_codigo.get())
        if not parsed:
            messagebox.showwarning("Atenção","Formato inválido.\nExemplo: __-__")
            return
        cliente_id, ramal_id = parsed

        try:
            d = self._buscar_ramal(cliente_id, ramal_id)
        except Exception as e:
            messagebox.showerror("Erro na API", f"Não foi possível buscar o ramal:\n{e}")
            return
        if not d:
            messagebox.showwarning("Não encontrado", f"Ramal {ramal_id} não localizado para a unidade {cliente_id}.")
            return

        login = d.get("usuario_autenticacao") or d.get("numero")
        if not login:
            messagebox.showwarning("Atenção", "Este ramal não possui 'usuario_autenticacao'.")
            return
        senha = d.get("senha_sip") or ""

        cred = {
            "login": login,
            "senha": senha if senha else SENHA_INICIAL,
            "nome_conta": d.get("nome") or login,
            "nome_exibicao": d.get("nome") or login,
            "dominio": SERVIDOR_PADRAO,
            "servidor_sip": SERVIDOR_PADRAO,
            "proxy_sip": SERVIDOR_PADRAO,
        }

        # ===== Fluxo =====
        if not instalar_atendas_silencioso(): return
        abrir_atendas()
        app, janela = conectar_janela_principal()

        # Verifica se já existe ramal configurado
        existe, login_atual = verificar_conta_existente(janela)
        if existe and login_atual:
            if not messagebox.askyesno(
                "Ramal já configurado",
                f"Já existe um ramal configurado nesta máquina:\n\n"
                f"• Ramal atual: {login_atual}\n\n"
                f"Deseja substituir por: {login} ?"
            ):
                messagebox.showinfo("Informação", "Operação cancelada. Mantendo o ramal atual.")
                return

            # Substituição: Editar Conta…
            if abrir_menu_tres_riscos(janela) and clicar_editar_conta(app):
                if not inserir_senha_inicial_e_enter(janela): return
                preencher_campos(janela, cred)
                if salvar_confirmar(janela):
                    messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' substituído com sucesso!")
                    finalizar_e_sair(self)
                    return

            # Plano B (se necessário): desbloquear via Adicionar, ESC, depois Editar
            if not abrir_menu_tres_riscos(janela): return
            if not clicar_adicionar_conta(app): return
            if not inserir_senha_inicial_e_enter(janela): return
            try:
                pyautogui.press("esc")
                time.sleep(0.3)
            except Exception:
                pass
            if not abrir_menu_tres_riscos(janela): return
            if not clicar_editar_conta(app): return
            if not inserir_senha_inicial_e_enter(janela): return
            preencher_campos(janela, cred)
            if salvar_confirmar(janela):
                messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' substituído com sucesso!")
                finalizar_e_sair(self)
                return

        # Primeira instalação
        if not abrir_menu_tres_riscos(janela): return
        if not clicar_adicionar_conta(app): return
        if not inserir_senha_inicial_e_enter(janela): return
        preencher_campos(janela, cred)
        if salvar_confirmar(janela):
            messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' configurado com sucesso!")
            finalizar_e_sair(self)
            return

# ====================== MAIN ======================
if __name__ == "__main__":
    AppGUI().mainloop()
