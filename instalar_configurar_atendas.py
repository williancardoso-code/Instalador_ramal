# ====================== IMPORTS ======================
import os, sys, time, shutil, tempfile, subprocess, re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import requests
from pywinauto import Application, Desktop
import pyautogui

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

# ====================== API ======================
def _headers():
    return {"usuario": USUARIO_API_PADRAO, "token": TOKEN_FIXO, "Accept": "application/json"}

def listar_ramais_api(cliente_id: int, pos_inicio=0):
    """Busca uma página de ramais para o cliente_id."""
    params = {"cliente_id": int(cliente_id), "pos_registro_inicial": int(pos_inicio)}
    r = requests.get(LISTAR_RAMAIS_URL, headers=_headers(), params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    if not isinstance(data, dict) or "dados" not in data:
        raise RuntimeError(f"JSON inesperado: {str(data)[:400]}")
    return data["dados"]

def listar_ramais_todos(cliente_id: int, max_registros: int = 50000):
    """Varre todas as páginas até acabar (ou atingir max_registros)."""
    todos, pos = [], 0
    while pos < max_registros:
        try:
            pagina = listar_ramais_api(cliente_id, pos_inicio=pos)
        except requests.HTTPError:
            break
        if not pagina:
            break
        todos.extend(pagina)
        passo = len(pagina) if len(pagina) > 0 else 100
        pos += passo
    return todos

def _get_int(d, *keys):
    for k in keys:
        if k in d and d[k] is not None:
            try:
                return int(str(d[k]).strip())
            except Exception:
                pass
    return None

def _get_str(d, *keys):
    for k in keys:
        if k in d and d[k] is not None:
            return str(d[k]).strip()
    return ""

def resolver_codigo_para_ramal(code: str):
    """
    Código = clienteID + ramalID (apenas dígitos). Testa TODAS as divisões do código.
    Para cada cliente_id candidato, pagina TODOS os ramais e procura o ramal_id.
    Retorna (cliente_id, dict_ramal) se encontrar.
    """
    code = (code or "").strip()
    if not code.isdigit():
        raise ValueError("O código deve conter apenas números (clienteID + ramalID).")

    for i in range(1, len(code)):
        cid = int(code[:i])
        rid = int(code[i:])

        ramais = listar_ramais_todos(cid)
        if not ramais:
            continue

        # Tenta id do ramal por diferentes chaves
        for d in ramais:
            ramal_id = _get_int(d, "id", "ramal_id", "codigo")
            if ramal_id is not None and ramal_id == rid:
                return cid, d

        # Fallback: alguns retornam "numero" como identificador
        for d in ramais:
            numero = _get_int(d, "numero")
            if numero is not None and numero == rid:
                return cid, d

    raise LookupError("Não encontrei um ramal para este código (clienteID+ramalID).")

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
    # Usa wildcard, pq o título pode ser "Atendas - <login> Nome"
    janela = app.window(title_re="^Atendas.*")
    return app, janela

# ====================== SUPORTE: VERIFICAÇÃO/SUBSTITUIÇÃO ======================
def extrair_login_do_titulo(janela):
    """Se o título for 'Atendas - atendas-1004 Fulano', retorna 'atendas-1004'; senão, None."""
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
    # 1) pywinauto direto (mais rápido)
    try:
        btn = janela.child_window(control_type="Button", found_index=0)
        btn.click_input()
        time.sleep(0.4)
        return True
    except:
        pass
    # 2) fallback por imagem (robusto)
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
        time.sleep(0.5)
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
                time.sleep(0.5)
                return True
            except Exception:
                continue
    except Exception as e:
        print(f"⚠️ Não encontrei 'Editar Conta…': {e}")
    return False

# ====== SENHA: usa exatamente o padrão que funcionava no seu fluxo ======
def inserir_senha_inicial_e_enter(janela):
    """
    Método rápido (o mesmo que estava funcionando antes):
    - tenta o primeiro Edit visível na janela principal,
    - se falhar, tenta em popups do mesmo processo,
    - ENTER no final.
    """
    print("🔑 Inserindo senha inicial…")
    # 1) tentar direto na janela principal
    try:
        caixa = janela.child_window(control_type="Edit", found_index=0)
        caixa.set_text(SENHA_INICIAL)
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.6)
        return True
    except:
        pass
    # 2) tentar nos popups (mesmo processo)
    try:
        app = Application(backend="uia").connect(path="Atendas.exe")
        main_handle = None
        try:
            main_handle = janela.handle
        except Exception:
            pass
        for w in Desktop(backend="uia").windows(process=app.process):
            if not w.is_visible():
                continue
            if main_handle and w.handle == main_handle:
                continue
            try:
                edits = [e for e in w.descendants(control_type="Edit") if e.is_visible()]
            except Exception:
                edits = []
            if edits:
                try:
                    edt = edits[0]
                    try:
                        edt.set_text(SENHA_INICIAL)
                    except Exception:
                        edt.type_keys("^a{BACKSPACE}", pause=0.02)
                        edt.type_keys(SENHA_INICIAL, with_spaces=True, pause=0.02)
                    time.sleep(0.2)
                    try:
                        w.type_keys("{ENTER}")
                    except Exception:
                        pyautogui.press("enter")
                    time.sleep(0.6)
                    return True
                except Exception:
                    continue
    except Exception:
        pass
    print("⚠️ Não consegui inserir a senha inicial.")
    return False

# ====================== CAMPOS / SALVAR ======================
def _preencher_por_rotulo(janela, rotulo, valor):
    """Tenta preencher pelo título acessível do campo."""
    try:
        campo = janela.child_window(title=rotulo, control_type="Edit")
        campo.set_text(valor)
        return True
    except Exception:
        return False

def _preencher_por_tab_sequence(janela, valores_seq):
    """
    Fallback rápido por ordem de TAB:
      1 Nome da Conta
      2 Nome de Exibição
      3 Usuário
      4 Login
      5 Senha
      6 Servidor SIP
      7 Proxy SIP
      8 Domínio
    Parte do primeiro Edit visível e sai tabulando.
    """
    try:
        edits = [e for e in janela.descendants(control_type="Edit") if e.is_visible()]
    except Exception:
        edits = []

    if not edits:
        return False

    # foca o primeiro e vai preenchendo
    try:
        edits[0].set_focus()
    except Exception:
        pass

    ordem = valores_seq[:]  # copia
    for idx, text in enumerate(ordem):
        # limpa, digita, TAB
        try:
            # usar teclado para ser mais compatível
            pyautogui.hotkey('ctrl', 'a'); time.sleep(0.05)
            pyautogui.press('backspace'); time.sleep(0.05)
            if text:
                pyautogui.typewrite(text, interval=0.01)
            time.sleep(0.05)
            if idx < len(ordem) - 1:
                pyautogui.press('tab'); time.sleep(0.05)
        except Exception:
            return False
    return True

def preencher_campos(janela, cred):
    """
    1) Tenta por rótulo (como no seu script base).
    2) Se algum falhar, usa fallback por TAB sequence (rápido).
    """
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
    falhou_algum = False
    for (t1, t2), valor in campos:
        ok = _preencher_por_rotulo(janela, t1, valor) or _preencher_por_rotulo(janela, t2, valor)
        if ok:
            print(f"   • {t1}: OK")
        else:
            print(f"   • {t1}: fallback por TAB será usado")
            falhou_algum = True

    if falhou_algum:
        print("↪️ Aplicando fallback por ordem de TAB…")
        seq = [nome_conta, nome_exib, usuario, login_val, senha_val, servidor_sip, proxy_sip, dominio]
        if _preencher_por_tab_sequence(janela, seq):
            print("   • Fallback por TAB: OK")
        else:
            print("⚠️ Fallback por TAB falhou (verifique o foco/ordem dos campos).")

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

# ====================== GUI – POR CÓDIGO ======================
class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Atendas – Instalação por Código (clienteID+ramalID)")
        self.geometry("560x160")
        self.resizable(False, False)

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Digite seu código:").grid(row=0, column=0, sticky="w")
        self.e_code = ttk.Entry(frm, width=28)
        self.e_code.grid(row=0, column=1, sticky="w", padx=(8, 12))
        self.e_code.focus_set()

        self.btn_install = ttk.Button(frm, text="Instalar/Configurar", command=self.on_install)
        self.btn_install.grid(row=1, column=1, sticky="e", pady=(12,0))

    def on_install(self):
        code = (self.e_code.get() or "").strip()
        if not code:
            messagebox.showwarning("Atenção", "Informe o código (clienteID + ramalID).")
            return

        try:
            cliente_id, d = resolver_codigo_para_ramal(code)
        except Exception as e:
            messagebox.showerror("Não encontrado", f"Não consegui resolver o código '{code}':\n{e}")
            return

        login = _get_str(d, "usuario_autenticacao", "numero")
        if not login:
            messagebox.showwarning("Atenção", "Este ramal não possui 'usuario_autenticacao'/'numero'.")
            return

        senha = _get_str(d, "senha_sip")
        if not senha:
            senha = simpledialog.askstring("Senha do SIP", f"Informe a senha SIP para '{login}':", show="•", parent=self)
            if not senha:
                return

        cred = {
            "login": login,
            "senha": senha,
            "nome_conta": _get_str(d, "nome") or login,
            "nome_exibicao": _get_str(d, "nome") or login,
            "dominio": SERVIDOR_PADRAO,
            "servidor_sip": SERVIDOR_PADRAO,
            "proxy_sip": SERVIDOR_PADRAO,
        }

        # ===== Fluxo: instalar/abrir, verificar, substituir ou adicionar =====
        if not instalar_atendas_silencioso(): return
        abrir_atendas()
        app, janela = conectar_janela_principal()

        existe, login_atual = verificar_conta_existente(janela)
        if existe and login_atual:
            if not messagebox.askyesno(
                "Ramal já configurado",
                f"Já existe um ramal nesta máquina:\n\n"
                f"• Ramal atual: {login_atual}\n\n"
                f"Deseja substituir pelos dados selecionados?"
            ):
                messagebox.showinfo("Informação", "Operação cancelada. Mantendo o ramal atual.")
                return

            # Tentar editar direto
            if abrir_menu_tres_riscos(janela) and clicar_editar_conta(app):
                if not inserir_senha_inicial_e_enter(janela): return
                preencher_campos(janela, cred)
                if salvar_confirmar(janela):
                    messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' substituído com sucesso!")
                return

            # Plano B: desbloquear via Adicionar, cancelar, depois Editar
            if not abrir_menu_tres_riscos(janela): return
            if not clicar_adicionar_conta(app): return
            if not inserir_senha_inicial_e_enter(janela): return
            try:
                pyautogui.press("esc"); time.sleep(0.3)
            except Exception:
                pass
            if not abrir_menu_tres_riscos(janela): return
            if not clicar_editar_conta(app):
                messagebox.showerror("Erro", "Não encontrei 'Editar Conta…' no menu.")
                return
            if not inserir_senha_inicial_e_enter(janela): return
            preencher_campos(janela, cred)
            if salvar_confirmar(janela):
                messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' substituído com sucesso!")
            return

        # Primeira instalação
        if not abrir_menu_tres_riscos(janela): return
        if not clicar_adicionar_conta(app): return
        if not inserir_senha_inicial_e_enter(janela): return
        preencher_campos(janela, cred)
        if salvar_confirmar(janela):
            messagebox.showinfo("Concluído", f"Ramal '{cred['login']}' configurado com sucesso!")

# ====================== MAIN ======================
if __name__ == "__main__":
    AppGUI().mainloop()
