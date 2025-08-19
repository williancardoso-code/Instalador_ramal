# ====================== IMPORTS ======================
import os, sys, time, shutil, tempfile, subprocess
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import requests
from pywinauto import Application
import pyautogui

# ====================== CONSTANTES ======================
USUARIO_API_PADRAO = "contatowilliancardoso@gmail.com"   # oculto na UI
TOKEN_FIXO = "6ddf7ebb-8d57-4d60-b63c-87cfcb853ace"      # oculto na UI
LISTAR_RAMAIS_URL = "https://novosistema.atendas.com.br/suite/api/listar_ramais?cliente_id=10&pos_registro_inicial=0"


SENHA_INICIAL = "Atendas@*2510"
SERVIDOR_PADRAO = "novosistema.atendas.com.br"
APP_EXEC = r"%LOCALAPPDATA%\Atendas\Atendas.exe"

BASE_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(__file__)
ATENDAS_EXE = os.path.join(BASE_DIR, "Atendas-3.21.6.exe")
IMG_MENU = os.path.join(BASE_DIR, "imagens_config", "menu_tres_riscos.png")

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
        time.sleep(10)
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao instalar Atendas:\n{e}")
        return False

def abrir_atendas():
    os.startfile(os.path.expandvars(APP_EXEC))
    print("⏳ Abrindo Atendas…")
    time.sleep(6)

def conectar_janela_principal():
    app = Application(backend="uia").connect(path="Atendas.exe")
    janela = app.window(best_match="Atendas")
    return app, janela

# ====================== AÇÕES NA UI DO ATENDAS ======================
def abrir_menu_tres_riscos(janela):
    print("📂 Abrindo menu…")
    try:
        btn = janela.child_window(control_type="Button", found_index=0)
        btn.click_input()
        time.sleep(0.5)
        return True
    except:
        pass
    try:
        pos = pyautogui.locateCenterOnScreen(IMG_MENU, confidence=0.75)
        if pos:
            pyautogui.click(pos)
            time.sleep(0.5)
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
        time.sleep(0.7)
        return True
    except Exception as e:
        print(f"⚠️ Não encontrei 'Adicionar Conta…': {e}")
        return False

def inserir_senha_inicial_e_enter(janela):
    print("🔑 Inserindo senha inicial…")
    try:
        caixa = janela.child_window(control_type="Edit", found_index=0)
        caixa.set_text(SENHA_INICIAL)
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(1.0)
        return True
    except:
        print("⚠️ Não consegui inserir a senha inicial.")
        return False

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
        time.sleep(0.5)
        try:
            janela.child_window(title="OK", control_type="Button").click_input()
        except:
            pass
        print("🎉 Configuração concluída!")
        return True
    except:
        print("⚠️ Erro ao salvar/confirmar.")
        return False

# ====================== GUI ======================
class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Atendas – Selecionar Ramal pela Unidade")
        self.geometry("820x560")  # janela mais alta/larga
        self.resizable(False, False)

        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Label(top, text="Unidade (Cliente ID):").grid(row=0, column=0, sticky="w")
        self.e_cli = ttk.Entry(top, width=12)
        self.e_cli.grid(row=0, column=1, sticky="w", padx=(6, 12))
        self.e_cli.insert(0, "9")

        self.btn_load = ttk.Button(top, text="Carregar ramais", command=self.on_load)
        self.btn_load.grid(row=0, column=2, sticky="w")

        # Botão de instalar AGORA no topo, sempre visível
        self.btn_install = ttk.Button(top, text="Instalar ramal selecionado",
                                      command=self.on_install, state="disabled")
        self.btn_install.grid(row=0, column=3, sticky="e", padx=(16, 0))

        # Tabela
        cols = ("nome", "usuario", "tipo")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        self.tree.heading("nome", text="Nome")
        self.tree.heading("usuario", text="Usuário (SIP)")
        self.tree.heading("tipo", text="Tipo (PA/PABX)")
        self.tree.column("nome", width=380)
        self.tree.column("usuario", width=250)
        self.tree.column("tipo", width=120, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        # Habilita botão ao selecionar e permite duplo-clique para instalar
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._toggle_install_button())
        self.tree.bind("<Double-1>", lambda e: self.on_install())

        self.ramais = []

    def _toggle_install_button(self):
        self.btn_install.config(state=("normal" if self.tree.selection() else "disabled"))

    def on_load(self):
        user = USUARIO_API_PADRAO
        token = TOKEN_FIXO
        try:
            cliente_id = int(self.e_cli.get().strip())
        except:
            messagebox.showwarning("Atenção", "Cliente ID inválido.")
            return
        try:
            dados = listar_ramais_api(user, token, cliente_id, 0)
        except Exception as e:
            messagebox.showerror("Erro na API", f"Falha ao listar ramais:\n{e}")
            return

        def tipo_label(v):  # ajuste se necessário (2=PA, demais=PABX)
            return "PA" if str(v) == "2" else "PABX"

        # organiza e popula; usa iid = índice para mapear seleção -> dados
        self.ramais = sorted(dados, key=lambda d: (d.get("nome") or "", d.get("usuario_autenticacao") or ""))
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, d in enumerate(self.ramais):
            nome = d.get("nome") or ""
            u = d.get("usuario_autenticacao") or ""
            t = tipo_label(d.get("tipo"))
            self.tree.insert("", "end", iid=str(idx), values=(nome, u, t))

        self._toggle_install_button()
        messagebox.showinfo("Diagnóstico", f"Ramais carregados: {len(self.ramais)}")

    def on_install(self):
        if not self.tree.selection():
            messagebox.showwarning("Atenção", "Selecione um ramal da lista.")
            return
        iid = self.tree.selection()[0]    # iid = índice que inserimos
        d = self.ramais[int(iid)]

        login = d.get("usuario_autenticacao") or d.get("numero")
        if not login:
            messagebox.showwarning("Atenção", "Este ramal não possui 'usuario_autenticacao'.")
            return

        senha = d.get("senha_sip") or ""
        if not senha:
            senha = simpledialog.askstring(
                "Senha do SIP",
                f"Informe a senha SIP para '{login}':",
                show="•",
                parent=self
            )
            if not senha:
                return

        cred = {
            "login": login,
            "senha": senha,
            "nome_conta": d.get("nome") or login,
            "nome_exibicao": d.get("nome") or login,
            "dominio": SERVIDOR_PADRAO,
            "servidor_sip": SERVIDOR_PADRAO,
            "proxy_sip": SERVIDOR_PADRAO,
        }

        # Fluxo de instalação + configuração
        if not instalar_atendas_silencioso(): return
        abrir_atendas()
        app, janela = conectar_janela_principal()
        if not abrir_menu_tres_riscos(janela): return
        if not clicar_adicionar_conta(app): return
        if not inserir_senha_inicial_e_enter(janela): return
        preencher_campos(janela, cred)
        if salvar_confirmar(janela):
            messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' configurado com sucesso!")

# ====================== MAIN ======================
if __name__ == "__main__":
    AppGUI().mainloop()
