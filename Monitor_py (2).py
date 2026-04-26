import streamlit as st
import streamlit.components.v1 as components
import json
import threading
from datetime import datetime, timedelta
try:
    import pyotp
except:
    pass
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os, re, io, calendar, time, base64
import plotly.express as px
from datetime import date, timedelta
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from cloud_driver import criar_driver


# ==============================================================================
# SISTEMA DE HISTÓRICO — Aresta PRO
# ==============================================================================

HISTORICO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aresta_historico.json")

def historico_salvar(funcao, status, detalhes=None):
    """Salva um registro no histórico. Remove registros com mais de 24h automaticamente."""
    try:
        # Carrega histórico existente
        registros = []
        if os.path.exists(HISTORICO_FILE):
            try:
                with open(HISTORICO_FILE, "r", encoding="utf-8") as f:
                    registros = json.load(f)
            except:
                registros = []

        # Remove registros com mais de 24h
        agora = datetime.now()
        registros = [
            r for r in registros
            if datetime.fromisoformat(r["data_hora"]) > agora - timedelta(hours=24)
        ]

        # Adiciona novo registro
        novo = {
            "data_hora": agora.isoformat(),
            "funcao": funcao,
            "status": status,
            "detalhes": detalhes or {}
        }
        registros.append(novo)

        # Salva
        with open(HISTORICO_FILE, "w", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False, indent=2)
    except Exception as e:
        pass  # Nunca deixa o histórico quebrar o sistema principal

def historico_carregar():
    """Carrega registros das últimas 24h."""
    try:
        if not os.path.exists(HISTORICO_FILE):
            return []
        with open(HISTORICO_FILE, "r", encoding="utf-8") as f:
            registros = json.load(f)
        agora = datetime.now()
        return [
            r for r in registros
            if datetime.fromisoformat(r["data_hora"]) > agora - timedelta(hours=24)
        ]
    except:
        return []


# ==============================================================================
# WATCHDOG — mata processo se travar mais de 5 minutos
# ==============================================================================

def _executar_com_watchdog(func, timeout_segundos=300):
    """
    Executa uma função com timeout.
    Se travar por mais de timeout_segundos, cancela e registra o erro.
    """
    resultado = {"erro": None, "concluido": False}

    def _wrapper():
        try:
            func()
            resultado["concluido"] = True
        except Exception as e:
            resultado["erro"] = str(e)

    thread = threading.Thread(target=_wrapper, daemon=True)
    thread.start()
    thread.join(timeout=timeout_segundos)

    if thread.is_alive():
        _log_aresta(f"WATCHDOG: Função travou após {timeout_segundos}s — cancelando", "ERRO")
        historico_salvar(
            funcao="Watchdog",
            status=f"❌ Processo cancelado por timeout ({timeout_segundos}s)",
            detalhes={"mensagem": "O robô travou e foi cancelado automaticamente"}
        )
        return False, "Processo cancelado por timeout"

    if resultado["erro"]:
        _log_aresta(f"WATCHDOG: Erro capturado: {resultado['erro']}", "ERRO")
        return False, resultado["erro"]

    return True, None


# ==============================================================================
# RECONEXÃO AUTOMÁTICA AMANDA
# ==============================================================================

def _verificar_sessao_amanda(driver, wait, amanda_url, login, senha, status):
    """
    Verifica se a sessão do Amanda ainda está ativa.
    Se expirou, faz login novamente sem reiniciar o browser.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    import time
    try:
        # Se a URL atual não contém o domínio do Amanda ou voltou para login
        url_atual = driver.current_url
        if "login" in url_atual.lower() or amanda_url.split("/")[2] not in url_atual:
            status.markdown("🔄 Sessão Amanda expirada — reconectando...")
            _log_aresta("Sessão Amanda expirada — reconectando automaticamente")
            driver.get(amanda_url)
            time.sleep(3)
            try:
                campo_email = driver.find_element(By.CSS_SELECTOR,
                    "input[type='email'], input[name='email'], #email")
                campo_email.clear()
                campo_email.send_keys(login)
                campo_senha = driver.find_element(By.CSS_SELECTOR,
                    "input[type='password']")
                campo_senha.clear()
                campo_senha.send_keys(senha)
                driver.find_element(By.CSS_SELECTOR,
                    "button[type='submit']").click()
                time.sleep(4)
                status.markdown("✅ Reconectado ao Amanda!")
                _log_aresta("Reconexão Amanda bem-sucedida")
            except:
                pass
    except:
        pass

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Aresta PRO", layout="wide")

# --- 2. CSS CLEAN LIGHT CLÁSSICO ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    * { font-family: 'Inter', Arial, sans-serif !important; }

    /* FUNDO GERAL */
    .stApp { background-color: #f5f7fa !important; color: #333333 !important; }
    .stApp > div { background-color: #f5f7fa !important; }
    [data-testid="stAppViewContainer"] { background-color: #f5f7fa !important; }
    [data-testid="stHeader"] { background-color: #ffffff !important; border-bottom: 0.5px solid #e2e6ed !important; }
    [data-testid="stToolbar"] { background-color: #ffffff !important; }
    [data-testid="block-container"] { background-color: #f5f7fa !important; }

    /* TEXTOS */
    h1, h2, h3 { color: #185FA5 !important; font-weight: 700 !important; text-align: center !important; }
    p, span, label, div { color: #333333 !important; }
    [data-testid="stMarkdownContainer"] p { color: #333333 !important; }
    [data-testid="stImage"] img { display: block; margin-left: auto; margin-right: auto; }

    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background: #ffffff !important;
        border-right: 0.5px solid #e2e6ed !important;
    }
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p { color: #333333 !important; font-weight: 500 !important; }
    section[data-testid="stSidebar"] h1 { color: #185FA5 !important; font-weight: 700 !important; }
    section[data-testid="stSidebar"] input { color: #333333 !important; }
    section[data-testid="stSidebar"] .stSelectbox label { color: #185FA5 !important; font-weight: 600 !important; }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #185FA5 !important; font-weight: 600 !important; }

    /* DROPDOWN do selectbox */
    [data-baseweb="popover"] { background-color: #ffffff !important; border: 0.5px solid #e2e6ed !important; border-radius: 8px !important; }
    [data-baseweb="menu"] { background-color: #ffffff !important; }
    [data-baseweb="menu"] ul { background-color: #ffffff !important; }
    [data-baseweb="menu"] li { background-color: #ffffff !important; color: #333333 !important; font-weight: 500 !important; }
    [data-baseweb="menu"] li:hover { background-color: #EBF4FF !important; color: #185FA5 !important; }
    [data-baseweb="option"] { background-color: #ffffff !important; color: #333333 !important; font-weight: 500 !important; }
    [data-baseweb="option"]:hover { background-color: #EBF4FF !important; color: #185FA5 !important; }
    [role="option"] { background-color: #ffffff !important; color: #333333 !important; font-weight: 500 !important; }
    [role="option"]:hover { background-color: #EBF4FF !important; color: #185FA5 !important; }
    [aria-selected="true"][role="option"] { background-color: #EBF4FF !important; color: #185FA5 !important; font-weight: 600 !important; }

    /* BOTÕES */
    .stButton > button {
        background: #185FA5 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: 0.5px solid #1a6db5 !important;
        font-weight: 600 !important;
        width: 100% !important;
        transition: background 0.2s, box-shadow 0.2s !important;
        box-shadow: none !important;
    }
    .stButton > button:hover {
        background: #0C447C !important;
        color: #ffffff !important;
        border-color: #0C447C !important;
    }

    /* TABS */
    .stTabs [data-baseweb="tab-list"] { background-color: #ffffff !important; border-radius: 8px; padding: 4px; border: 0.5px solid #e2e6ed !important; }
    .stTabs [data-baseweb="tab"] { color: #888888 !important; font-weight: 600 !important; }
    .stTabs [aria-selected="true"] { background-color: #185FA5 !important; color: #ffffff !important; border-radius: 6px !important; }
    .stTabs [data-baseweb="tab-panel"] { background-color: #f5f7fa !important; }

    /* MÉTRICAS */
    [data-testid="metric-container"] { background: #ffffff !important; border-left: 4px solid #185FA5 !important; border-radius: 8px !important; padding: 12px !important; border-top: 0.5px solid #e2e6ed !important; border-right: 0.5px solid #e2e6ed !important; border-bottom: 0.5px solid #e2e6ed !important; }
    [data-testid="metric-container"] label { color: #888888 !important; font-weight: 600 !important; text-transform: uppercase !important; font-size: 11px !important; letter-spacing: 0.5px !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #185FA5 !important; font-weight: 700 !important; }

    /* INPUTS */
    .stTextInput input { background: #ffffff !important; color: #333333 !important; border: 0.5px solid #d0d5dd !important; border-radius: 6px !important; }
    .stTextInput input:focus { border: 1.5px solid #185FA5 !important; }
    .stTextInput label { color: #555555 !important; font-weight: 500 !important; }
    .stSelectbox > div > div { background: #ffffff !important; color: #333333 !important; border: 0.5px solid #d0d5dd !important; border-radius: 6px !important; }
    .stSelectbox label { color: #555555 !important; font-weight: 500 !important; }
    .stDateInput input { background: #ffffff !important; color: #333333 !important; border: 0.5px solid #d0d5dd !important; border-radius: 6px !important; }
    .stDateInput label { color: #555555 !important; font-weight: 500 !important; }
    .stNumberInput input { background: #ffffff !important; color: #333333 !important; border: 0.5px solid #d0d5dd !important; border-radius: 6px !important; }
    .stNumberInput label { color: #555555 !important; font-weight: 500 !important; }
    .stCheckbox label { color: #333333 !important; }
    .stMultiSelect > div { background: #ffffff !important; color: #333333 !important; border: 0.5px solid #d0d5dd !important; border-radius: 6px !important; }
    .stMultiSelect label { color: #555555 !important; font-weight: 500 !important; }
    .stToggle label { color: #333333 !important; }

    /* PROGRESS BAR */
    .stProgress > div > div { background: #185FA5 !important; }

    /* INFO / SUCCESS / ERROR / WARNING */
    .stSuccess { border-left: 4px solid #1a8a4a !important; background: #f0faf4 !important; color: #333333 !important; }
    .stInfo    { border-left: 4px solid #185FA5 !important; background: #EBF4FF !important; color: #333333 !important; }
    .stWarning { border-left: 4px solid #b8860b !important; background: #FFFBEB !important; color: #333333 !important; }
    .stError   { border-left: 4px solid #c0392b !important; background: #FEF2F2 !important; color: #333333 !important; }

    /* DATAFRAME */
    [data-testid="stDataFrame"] { border: 0.5px solid #e2e6ed !important; border-radius: 8px !important; overflow: hidden !important; }
    [data-testid="stDataFrame"] th { background: #f8f9fb !important; color: #185FA5 !important; font-weight: 600 !important; }
    [data-testid="stDataFrame"] td { background: #ffffff !important; color: #333333 !important; }

    /* EXPANDER */
    .streamlit-expanderHeader { background: #f0f4f8 !important; color: #185FA5 !important; border-radius: 8px !important; border: 0.5px solid #e2e6ed !important; }
    .streamlit-expanderContent { background: #ffffff !important; border: 0.5px solid #e2e6ed !important; border-radius: 0 0 8px 8px !important; }

    /* SEPARADOR */
    hr { border-color: #e2e6ed !important; }
    </style>
""", unsafe_allow_html=True)

# --- 3. LOGIN E ESTADO ---
if 'dados_busca' not in st.session_state: st.session_state.dados_busca = None
if 'dados_busca_caravaggio' not in st.session_state: st.session_state.dados_busca_caravaggio = None
if 'dados_busca_hamburgo' not in st.session_state: st.session_state.dados_busca_hamburgo = None
if 'dados_busca_terrazzo' not in st.session_state: st.session_state.dados_busca_terrazzo = None
if 'dados_busca_serra_negra' not in st.session_state: st.session_state.dados_busca_serra_negra = None
if 'dados_busca_canoeiros' not in st.session_state: st.session_state.dados_busca_canoeiros = None
if 'omnibees_status' not in st.session_state: st.session_state.omnibees_status = None
if 'amanda_resultado' not in st.session_state: st.session_state.amanda_resultado = None
if 'dados_hits_atual' not in st.session_state: st.session_state.dados_hits_atual = None
if 'dados_hits_anterior' not in st.session_state: st.session_state.dados_hits_anterior = None
if "autenticado" not in st.session_state: st.session_state.autenticado = False


if not st.session_state.autenticado:
    c2 = st.columns([1,2,1])[1]
    with c2:
        if os.path.exists("logo.png"):
            st.markdown(
                f'<div style="display:flex; justify-content:center; margin-bottom:20px;">' +
                f'<img src="data:image/png;base64,{__import__("base64").b64encode(open("logo.png","rb").read()).decode()}" width="200" style="border-radius:12px;"/>' +
                f'</div>',
                unsafe_allow_html=True
            )
        st.title("🔒 ACESSO")
        user = st.text_input("Usuário")
        pw = st.text_input("Senha", type="password")
        if st.button("ENTRAR"):
            if user == "easy" and pw == "easy2026":
                st.session_state.autenticado = True
                st.rerun()
            else: st.error("Dados incorretos.")
    st.stop()

# --- 4. NAVEGAÇÃO DE ABAS ---
if os.path.exists("logo.png"):
    st.sidebar.markdown(
        f'<div style="display:flex; justify-content:center; padding: 10px 0;">'
        f'<img src="data:image/png;base64,{__import__("base64").b64encode(open("logo.png","rb").read()).decode()}" width="160" style="border-radius:12px;"/>'
        f'</div>',
        unsafe_allow_html=True
    )
st.sidebar.markdown(
    '<p style="text-align:center; color:#185FA5; font-size:18px; font-weight:800; letter-spacing:1.5px; margin-top:8px; margin-bottom:4px;">ARESTA PRO</p>',
    unsafe_allow_html=True
)

aba_selecionada = st.sidebar.selectbox("Selecione a Ferramenta:", [
    "🎯 Shopper Alto da Boa Vista",
    "🏕️ Shopper Villa Caravaggio",
    "🍺 Shopper Hamburgo",
    "🏨 Shopper Terrazzo",
    "🏖️ Shopper Serra Negra",
    "🏞️ Shopper Canoeiros",
    "📈 Hits - Alto da Boa Vista",
    "🐝 Omnibees",
    "📊 Monitor de Pick-up",
    "🤖 Fecho Automático Amanda",
    "📋 Histórico"
])
st.sidebar.markdown("---")

def _capturar_2fa_authenticator(status):
    """Gera o código 2FA automaticamente via pyotp."""
    try:
        import pyotp
        totp = pyotp.TOTP("GY2DSYZZGAZWMYZQGY4Q")
        codigo = totp.now()
        status.markdown(f"✅ 2FA gerado: {codigo[:3]} {codigo[3:]}")
        _log_aresta(f"2FA gerado via pyotp: {codigo[:3]}***")
        return codigo
    except Exception as e:
        status.markdown(f"⚠️ Erro ao gerar 2FA: {e}")
        _log_aresta(f"ERRO ao gerar 2FA: {e}", "ERRO")
        return None



# ==============================================================================
# FUNÇÃO: Log de execuções — aresta_log.txt
# ==============================================================================

def _log_aresta(mensagem, nivel="INFO"):
    """Registra no arquivo aresta_log.txt com data/hora."""
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aresta_log.txt")
        linha = f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] [{nivel}] {mensagem}\n"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(linha)
    except:
        pass


# ==============================================================================
# FUNÇÃO: Fechar popups do Omnibees (antes de navegar)
# ==============================================================================

def _fechar_popups_omnibees(driver, status=None):
    """
    Varre a página e fecha qualquer popup/modal que esteja aberto.
    Chamada após o login E após cada refresh no Omnibees.
    """
    from selenium.webdriver.common.by import By
    import time
    try:
        fechou = driver.execute_script("""
            // Tenta fechar qualquer modal/popup visível
            var seletores = [
                'button[aria-label="Close"]',
                'button[aria-label="Fechar"]',
                'button.close',
                '.modal-close',
                '.btn-close',
                'lib-icon[name="close"]',
                'button[class*="close"]',
                '.modal .close',
                '.modal-dialog button.close',
                'div[class*="modal"] button[class*="close"]',
                'lib-button[text="Fechar"] button',
                'lib-button[text="OK"] button',
                'lib-button[text="Não mostrar"] button',
            ];
            for (var sel of seletores) {
                var els = document.querySelectorAll(sel);
                for (var el of els) {
                    if (el.offsetParent !== null) {
                        el.click();
                        return true;
                    }
                }
            }
            // Tenta pelo texto do botão
            var btns = document.querySelectorAll('button, lib-button');
            for (var btn of btns) {
                var txt = (btn.innerText || '').trim().toLowerCase();
                if ((txt === 'fechar' || txt === 'ok' || txt === 'não mostrar' || 
                     txt === 'close' || txt === 'dismiss') && btn.offsetParent !== null) {
                    btn.click();
                    return true;
                }
            }
            return false;
        """)
        if fechou:
            if status:
                status.markdown("✅ Popup Omnibees fechado!")
            _log_aresta("Popup Omnibees fechado automaticamente")
            time.sleep(1)
    except:
        pass


# ==============================================================================
# FUNÇÃO: Capturar 2FA com retry automático (trata código expirado)
# ==============================================================================

def _capturar_2fa_com_retry(status, max_tentativas=3):
    """
    Captura o código 2FA com até 3 tentativas.
    O código expira a cada 30s — se o Omnibees rejeitar, captura novamente.
    """
    import time as t_retry
    for tentativa in range(1, max_tentativas + 1):
        if tentativa > 1:
            status.markdown(f"🔄 Tentativa {tentativa} de capturar 2FA...")
            t_retry.sleep(2)
        codigo = _capturar_2fa_authenticator(status)
        if codigo and len(codigo) == 6:
            _log_aresta(f"2FA capturado na tentativa {tentativa}: {codigo[:3]}***")
            return codigo
    _log_aresta("ERRO: Não foi possível capturar 2FA após 3 tentativas", "ERRO")
    return None

# ==============================================================================
# FUNÇÕES AUXILIARES OMNIBEES — ESCOPO GLOBAL
# ==============================================================================

def _login_omnibees(driver, wait, config, codigo_2fa, status):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    import time
    status.markdown("🔄 Abrindo Omnibees...")
    driver.get("https://myhotel2.omnibees.com/#/home")
    time.sleep(2)
    driver.refresh()
    time.sleep(2)
    try:
        popup_close = driver.find_element(By.CSS_SELECTOR,
            "button.close, .modal-close, [aria-label='Close'], [aria-label='Fechar'], "
            ".btn-close, button[class*='close'], .close-button, lib-icon[name='close']")
        popup_close.click()
        time.sleep(2)
    except:
        pass
    try:
        campo_login = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
            "input[type='text'], input[name='username'], #username")))
        time.sleep(5)
        campo_login.clear()
        campo_login.send_keys(config["login"])
        time.sleep(0.5)
        campo_senha = driver.find_element(By.CSS_SELECTOR,
            "input[type='password'], input[name='password'], #password")
        campo_senha.clear()
        campo_senha.send_keys(config["senha"])
        time.sleep(0.5)
        driver.find_element(By.CSS_SELECTOR,
            "button[type='submit'], input[type='submit']").click()
        status.markdown("🔐 Login realizado, aguardando autenticador...")
        time.sleep(3)
    except:
        status.markdown("🔄 Já logado, continuando...")
    try:
        campo_2fa = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
            "input[name='code'], input[placeholder*='código'], "
            "input[placeholder*='code'], input[maxlength='6']")))
        campo_2fa.clear()
        campo_2fa.send_keys(codigo_2fa)
        time.sleep(0.3)
        driver.find_element(By.CSS_SELECTOR,
            "button[type='submit'], input[type='submit']").click()
        status.markdown("✅ Autenticador validado!")
        time.sleep(4)
        driver.refresh()
        time.sleep(6)
    except:
        status.markdown("🔄 2FA não solicitado, continuando...")
        time.sleep(4)
    # Aguarda 2s após login e atualiza para fechar popup que pode aparecer
    status.markdown("🔄 Atualizando página após login...")
    time.sleep(2)
    driver.refresh()
    time.sleep(3)
    _fechar_popups_omnibees(driver, status)  # Fecha popup se aparecer após refresh
    status.markdown("📋 Abrindo Tarifários e Disponibilidade...")
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH,
            "//*[contains(text(), 'Tarifários e Disponibilidade')]"))).click()
        time.sleep(2)
    except:
        driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var el of els) {
                if (el.innerText && el.innerText.trim() === 'Tarifários e Disponibilidade') {
                    el.click(); break;
                }
            }
        """)
        time.sleep(2)
    status.markdown("💰 Abrindo Preços e Disponibilidade...")
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH,
            "//*[contains(text(), 'Preços e Disponibilidade')]"))).click()
        time.sleep(2)
    except:
        driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var el of els) {
                if (el.innerText && el.innerText.trim() === 'Preços e Disponibilidade'
                        && el.children.length === 0) {
                    el.click(); break;
                }
            }
        """)
        time.sleep(2)


def _selecionar_datas(driver, wait, data_ini, data_fim, status):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    import time
    status.markdown("📅 Adicionando datas...")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
        "#btn-updateRates-button-dates-add-new"))).click()
    time.sleep(0.8)
    status.markdown("📅 Abrindo calendário...")
    try:
        btn_cal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
            "lib-date-range button")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_cal)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", btn_cal)
    except:
        driver.execute_script(
            "var btn = document.querySelector('lib-date-range button'); if (btn) btn.click();")
    time.sleep(0.8)

    def _navegar_mes_calendario(driver, data_alvo, max_cliques=24):
        """Navega o calendário até o mês/ano correto clicando na seta >"""
        import time
        meses_pt = {1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio",
                    6:"Junho", 7:"Julho", 8:"Agosto", 9:"Setembro",
                    10:"Outubro", 11:"Novembro", 12:"Dezembro"}
        mes_alvo = meses_pt[data_alvo.month].strip().lower()
        ano_alvo = str(data_alvo.year)

        for _ in range(max_cliques):
            # Lê TODOS os meses visíveis no calendário (pode mostrar 2 meses)
            meses_visiveis = driver.execute_script("""
                var resultado = [];
                var spans = document.querySelectorAll('span.selected-year-month');
                for (var s of spans) {
                    var m = s.querySelector('span.month');
                    var a = s.querySelector('span.year');
                    if (m && a) {
                        resultado.push({
                            mes: m.innerText.trim().toLowerCase(),
                            ano: a.innerText.trim()
                        });
                    }
                }
                // Fallback se não encontrar
                if (resultado.length === 0) {
                    var m = document.querySelector('span.month');
                    var a = document.querySelector('span.year');
                    if (m && a) resultado.push({
                        mes: m.innerText.trim().toLowerCase(),
                        ano: a.innerText.trim()
                    });
                }
                return resultado;
            """)

            # Verifica se o mês alvo está em qualquer um dos meses visíveis
            for mv in meses_visiveis:
                if mes_alvo in mv['mes'] and ano_alvo == mv['ano']:
                    return True

            # Clica na seta > para avançar o mês
            clicou = driver.execute_script("""
                var icon = document.querySelector('[data-cy="btn_next-icon_angle-down"]');
                if (icon) {
                    var btn = icon.closest('button') || icon.parentElement;
                    if (btn) { btn.click(); return true; }
                }
                return false;
            """)
            time.sleep(0.8)
            if not clicou:
                break
        return False

    for dt_omni in [data_ini, data_fim]:
        # Navega até o mês correto antes de clicar no dia
        status.markdown(f"📅 Navegando para {dt_omni.strftime('%m/%Y')}...")
        _navegar_mes_calendario(driver, dt_omni)
        time.sleep(0.3)

        dia_str_z = str(dt_omni.day).zfill(2)
        dia_str   = str(dt_omni.day)
        mes_str_z = str(dt_omni.month).zfill(2)
        mes_str   = str(dt_omni.month)
        clicou = False
        for cy in [
            f"{dia_str_z}-month_id_-_{mes_str_z}",
            f"{dia_str}-month_id_-_{mes_str}",
            f"{dia_str_z}-month_id_-_{mes_str}",
            f"{dia_str}-month_id_-_{mes_str_z}",
        ]:
            try:
                el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    f"span[data-cy='{cy}']")))
                driver.execute_script("arguments[0].click();", el)
                clicou = True
                break
            except:
                pass
        if not clicou:
            driver.execute_script("""
                var spans = document.querySelectorAll('span[data-cy]');
                for (var s of spans) {
                    var cy = s.getAttribute('data-cy') || '';
                    if (cy.startsWith(arguments[0]) || cy.startsWith(arguments[1])) {
                        s.click(); break;
                    }
                }
            """, f"{dia_str}-month_id_-_", f"{dia_str_z}-month_id_-_")
        time.sleep(0.4)
    time.sleep(0.5)
    try:
        ok_btn = driver.find_element(By.CSS_SELECTOR,
            "div[class*='cdk-overlay-pane'] div[class*='esp-footer'] lib-button:first-child button, "
            "div[class*='esp-footer'] lib-button:first-child button")
        driver.execute_script("arguments[0].click();", ok_btn)
    except:
        driver.execute_script("""
            var btns = document.querySelectorAll('lib-button button, button');
            for (var b of btns) {
                var txt = (b.innerText || '').trim().toUpperCase();
                if (txt === 'OK' || txt === 'CONFIRMAR' || txt === 'APLICAR') { b.click(); break; }
            }
        """)
    time.sleep(0.8)


def _clicar_checkbox_texto(driver, texto):
    driver.execute_script("""
        (function(texto) {
            var walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null, false);
            var node;
            while (node = walker.nextNode()) {
                if (node.textContent.trim() === texto) {
                    var parent = node.parentElement;
                    for (var i = 0; i < 8; i++) {
                        if (!parent) break;
                        var chk = parent.querySelector('input[type="checkbox"]');
                        if (chk) {
                            chk.scrollIntoView({block: 'center'});
                            chk.click();
                            return;
                        }
                        parent = parent.parentElement;
                    }
                }
            }
        })(arguments[0]);
    """, texto)
    import time
    time.sleep(0.4)


TARIFARIOS_DATA_CY = {
    "Trf Bancorbras":                  "update_treeview-union_rates_package_tree-item_click_-_Trf Bancorbras",
    "CLUBE MONTREAL":                  "update_treeview-union_rates_package_tree-item_click_-_CLUBE MONTREAL",
    "Tarifa Flexível":                 "update_treeview-union_rates_package_tree-item_click_-_Tarifa Flexivel",
    "Condição Exclusiva":              "update_treeview-union_rates_package_tree-item_click_-_Condição Exclusiva",
    "Programa Preferencial - Orinter": "update_treeview-union_rates_package_tree-item_click_-_Programa Preferencial - Orinter",
    "Tarifa Não Reembolsável.":        "update_treeview-union_rates_package_tree-item_click_-_Tarifa Não Reembolsável.",
    "Tarifa Site":                     "update_treeview-union_rates_package_tree-item_click_-_Tarifa Site",
}
SELECIONAR_TODOS_CY = "update_treeview-union_rates_package_tree-select_all_checkbox-"


def _selecionar_tarifarios(driver, wait, sel_todos, tarifarios_sel, status):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    import time
    status.markdown("🏷️ Abrindo painel de tarifários...")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
        "body > ob-root > div > ob-layout > div > div > div > div > ob-rate-availability "
        "> div.h-100 > lib-frame > div.frm-inside > div.wrapper.overflow-auto > div.row "
        "> div:nth-child(2) > div > div.button-area"))).click()
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
        "#updateRates-button-rate-packages-sidepanel-save")))
    time.sleep(1.5)

    def _clicar_por_data_cy(cy):
        el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            f'input[data-cy="{cy}"]')))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", el)
        time.sleep(0.4)

    if sel_todos:
        status.markdown("🏷️ Selecionando todos os tarifários...")
        # Clica no checkbox "Selecionar Todos os Tarifários e Pacotes" via TreeWalker
        _clicar_checkbox_texto(driver, "Selecionar Todos os Tarifários e Pacotes")
        time.sleep(0.5)
    else:
        for tar in tarifarios_sel:
            status.markdown(f"🏷️ Selecionando: {tar}")
            cy = TARIFARIOS_DATA_CY.get(tar)
            if cy:
                try:
                    # Clica no div.tree-item pelo data-cy (elemento clicável do Angular)
                    el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                        f'div[data-cy="{cy}"]')))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    time.sleep(0.3)
                    inner = el.find_element(By.CSS_SELECTOR, "div[clicktype='1']")
                    driver.execute_script("arguments[0].click();", inner)
                except:
                    # Fallback: TreeWalker pelo texto
                    _clicar_checkbox_texto(driver, tar)
            else:
                _clicar_checkbox_texto(driver, tar)
            time.sleep(0.5)

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
        "#updateRates-button-rate-packages-sidepanel-save"))).click()
    time.sleep(1)


def _selecionar_quartos(driver, wait, sel_todos, quartos_sel, status):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    import time
    status.markdown("\U0001f6cf\ufe0f Selecionando quartos...")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
        "body > ob-root > div > ob-layout > div > div > div > div > ob-rate-availability "
        "> div.h-100 > lib-frame > div.frm-inside > div.wrapper.overflow-auto > div.row "
        "> div:nth-child(3) > div > div.button-area"))).click()
    # Aguarda painel carregar
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
        'input[data-cy="room_type_list-select_all_chk-"]')))
    time.sleep(1)

    chk_todos = driver.find_element(By.CSS_SELECTOR,
        'input[data-cy="room_type_list-select_all_chk-"]')

    if sel_todos:
        # Marca "Selecionar Todos" se ainda não estiver marcado
        if not chk_todos.is_selected():
            driver.execute_script("arguments[0].click();", chk_todos)
        time.sleep(0.5)
    else:
        # 1. Garante que NADA está marcado — desmarca tudo
        if chk_todos.is_selected():
            driver.execute_script("arguments[0].click();", chk_todos)  # desmarca
            time.sleep(0.5)
        # Se não estava marcado, verifica se tem algum item marcado individualmente
        # Clica duas vezes para garantir estado limpo
        driver.execute_script("arguments[0].click();", chk_todos)  # marca todos
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", chk_todos)  # desmarca todos
        time.sleep(0.5)

        # 2. Seleciona cada quarto pelo div com data-cy exato
        for qrt in quartos_sel:
            status.markdown(f"\U0001f6cf\ufe0f Selecionando: {qrt}")
            try:
                el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    f'div[data-cy="update_rate_availability-room_types_sidepanel-checkbox_-_{qrt}"]'
                )))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", el)
            except:
                _clicar_checkbox_texto(driver, qrt)
            time.sleep(0.4)

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
        "#updateRates-button-rate-packages-sidepanel-save"))).click()
    time.sleep(0.5)


def gerar_excel_formatado(tabela_final, sua_col, sheet_name='Monitoramento'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        tabela_final.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        cor_header    = PatternFill("solid", fgColor="002D62")
        cor_sua_col   = PatternFill("solid", fgColor="BBDEFB")
        cor_media     = PatternFill("solid", fgColor="E3F2FD")
        cor_esgotado  = PatternFill("solid", fgColor="FFCDD2")
        fonte_header  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        fonte_sua_col = Font(bold=True, color="002D62", name="Calibri", size=11)
        fonte_normal  = Font(name="Calibri", size=11)
        alinhamento_c = Alignment(horizontal="center", vertical="center")
        borda = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        colunas = list(tabela_final.columns)

        for col_idx, col_name in enumerate(colunas, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = cor_header
            cell.font = fonte_header
            cell.alignment = alinhamento_c
            cell.border = borda

        for row_idx in range(2, ws.max_row + 1):
            for col_idx, col_name in enumerate(colunas, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = alinhamento_c
                cell.border = borda
                val = cell.value

                if col_name == 'Data':
                    cell.font = Font(bold=True, name="Calibri", size=11)
                elif col_name == sua_col:
                    cell.fill = cor_sua_col
                    cell.font = fonte_sua_col
                    if val and str(val) != 'Esgotado':
                        try:
                            cell.value = float(str(val).replace('.', '').replace(',', '.'))
                            cell.number_format = 'R$ #,##0'
                        except: pass
                    elif str(val) == 'Esgotado':
                        cell.fill = cor_esgotado
                        cell.font = Font(bold=True, color="C62828", name="Calibri", size=11)
                elif col_name == 'Market Sync (Média)':
                    cell.fill = cor_media
                    cell.font = Font(bold=True, name="Calibri", size=11)
                    if val:
                        try:
                            cell.value = float(str(val))
                            cell.number_format = 'R$ #,##0.00'
                        except: pass
                else:
                    cell.font = fonte_normal
                    if val and str(val) != 'Esgotado':
                        try:
                            cell.value = float(str(val).replace('.', '').replace(',', '.'))
                            cell.number_format = 'R$ #,##0'
                        except: pass
                    elif str(val) == 'Esgotado':
                        cell.fill = cor_esgotado
                        cell.font = Font(bold=True, color="C62828", name="Calibri", size=11)

        for col_idx, col_name in enumerate(colunas, 1):
            max_len = max(len(str(col_name)), 12)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

        ws.freeze_panes = "A2"

    output.seek(0)
    return output

# --- FUNÇÃO: Executa varredura com Edge (por nome) ---
def executar_varredura(datas_para_busca, concorrentes, url_base_busca=None):
    todas_buscas = []
    driver = criar_driver()
    wait = WebDriverWait(driver, 8)
    status = st.empty()
    prog = st.progress(0)
    total = len(datas_para_busca) * len(concorrentes)
    cont = 0

    for d_in, d_out in datas_para_busca:
        for hotel in concorrentes:
            status.markdown(f"📡 Pesquisando: **{d_in.strftime('%d/%m')}** | {hotel}")
            url = (f"https://www.booking.com/searchresults.pt-br.html?ss={hotel.replace(' ', '+')}"
                   f"&checkin={d_in.isoformat()}&checkout={d_out.isoformat()}"
                   f"&group_adults=2&no_rooms=1&selected_currency=BRL&lang=pt-br")
            driver.get(url)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="property-card"]')))
                card = driver.find_element(By.CSS_SELECTOR, '[data-testid="property-card"]')
                preco_elem = card.find_element(By.CSS_SELECTOR, '[data-testid="price-and-discounted-price"]')
                valor = re.findall(r'\d+', preco_elem.text.replace('.', '').replace(',', ''))
                preco_final = int(valor[-1]) if valor else "Esgotado"
            except:
                preco_final = "Esgotado"
            todas_buscas.append({"Data": d_in.strftime("%d/%m/%Y"), "Hotel": hotel, "Preço": preco_final})
            cont += 1
            prog.progress(cont / total)

    driver.quit()
    status.success("✅ Varredura concluída!")
    df_resultado = pd.DataFrame(todas_buscas)
    historico_salvar(
        funcao="Shopper Booking",
        status="✅ Sucesso",
        detalhes={
            "hotel": concorrentes[0] if concorrentes else "",
            "total_buscas": len(todas_buscas),
            "datas": f"{datas_para_busca[0][0].strftime('%d/%m/%Y')} a {datas_para_busca[-1][0].strftime('%d/%m/%Y')}"
        }
    )
    return df_resultado

# --- FUNÇÃO: Executa varredura por URL direta (slug do Booking) ---
def executar_varredura_por_url(datas_para_busca, slugs):
    """
    slugs = dicionário {nome_hotel: slug_booking}
    Ex: {"Serra Negra Pousada Spa": "serra-negra-pousada-spa"}
    """
    todas_buscas = []
    driver = criar_driver()
    wait = WebDriverWait(driver, 8)
    status = st.empty()
    prog = st.progress(0)
    total = len(datas_para_busca) * len(slugs)
    cont = 0

    for d_in, d_out in datas_para_busca:
        for nome, slug in slugs.items():
            status.markdown(f"📡 Pesquisando: **{d_in.strftime('%d/%m')}** | {nome}")
            url = (f"https://www.booking.com/hotel/br/{slug}.pt-br.html"
                   f"?checkin={d_in.isoformat()}&checkout={d_out.isoformat()}"
                   f"&group_adults=2&no_rooms=1&selected_currency=BRL")
            driver.get(url)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                    '[data-testid="price-and-discounted-price"], .hprt-price-price, [class*="priceValue"]')))
                preco_elem = driver.find_element(By.CSS_SELECTOR,
                    '[data-testid="price-and-discounted-price"], .hprt-price-price, [class*="priceValue"]')
                valor = re.findall(r'\d+', preco_elem.text.replace('.', '').replace(',', ''))
                preco_final = int(valor[-1]) if valor else "Esgotado"
            except:
                preco_final = "Esgotado"
            todas_buscas.append({"Data": d_in.strftime("%d/%m/%Y"), "Hotel": nome, "Preço": preco_final})
            cont += 1
            prog.progress(cont / total)

    driver.quit()
    status.success("✅ Varredura concluída!")
    df_resultado = pd.DataFrame(todas_buscas)
    historico_salvar(
        funcao="Shopper Booking (URL)",
        status="✅ Sucesso",
        detalhes={
            "hotel": list(slugs.keys())[0] if slugs else "",
            "total_buscas": len(todas_buscas),
            "datas": f"{datas_para_busca[0][0].strftime('%d/%m/%Y')} a {datas_para_busca[-1][0].strftime('%d/%m/%Y')}"
        }
    )
    return df_resultado

# --- FUNÇÃO: Monta e exibe tabela de resultados ---
def exibir_tabela_resultados(df, concorrentes, sua_col, key_download):
    df['Preço_Num'] = pd.to_numeric(df['Preço'], errors='coerce')
    medias = df.groupby('Data')['Preço_Num'].mean().reset_index()
    medias.columns = ['Data', 'Market Sync (Média)']
    df_pivot = df.pivot(index='Data', columns='Hotel', values='Preço').reset_index()
    outros = [h for h in concorrentes if h != sua_col]
    colunas_ordenadas = ['Data', sua_col] + outros
    colunas_disp = [c for c in colunas_ordenadas if c in df_pivot.columns]
    tabela_final = df_pivot[colunas_disp].merge(medias, on='Data')

    st.dataframe(
        tabela_final.style.format(subset=['Market Sync (Média)'], precision=2)
        .background_gradient(subset=['Market Sync (Média)'], cmap='YlGnBu')
        .set_properties(subset=[sua_col] if sua_col in tabela_final.columns else [],
                        **{'background-color': '#e1f5fe', 'font-weight': 'bold', 'border': '1px solid #002D62'}),
        use_container_width=True
    )

    excel_file = gerar_excel_formatado(tabela_final, sua_col)
    st.download_button(
        label="📥 Baixar Planilha",
        data=excel_file,
        file_name=f"{key_download}_{date.today().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key_download
    )

# ==============================================================================
# ABA 1: MONITORAMENTO DE SHOPPER - ALTO DA BOA VISTA
# ==============================================================================

# ==============================================================================
# FUNÇÃO: Seleção de datas para Shoppers (3 modos)
# ==============================================================================
def _shopper_selecionar_datas(key_prefix):
    hoje = date.today()
    modo = st.sidebar.radio(
        "📅 Modo de Busca:",
        ["🗓️ Data Específica", "📆 Varredura Mensal", "🏖️ Finais de Semana", "📆 Múltiplos Períodos", "🎉 Feriados", "📋 Semana (2-3-4 noites)"],
        key=f"modo_{key_prefix}"
    )

    if modo == "🗓️ Data Específica":
        c_in  = st.sidebar.date_input("Check-in",  hoje, key=f"cin_{key_prefix}", format="DD/MM/YYYY")
        c_out = st.sidebar.date_input("Check-out", c_in + timedelta(days=1), key=f"cout_{key_prefix}", format="DD/MM/YYYY")
        return [(c_in, c_out)]

    elif modo == "📆 Varredura Mensal":
        mes_sel = st.sidebar.selectbox("Mês",
                    list(range(1,13)),
                    format_func=lambda m: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m-1],
                    index=hoje.month - 1, key=f"mes_m_{key_prefix}")
        ano_sel = st.sidebar.selectbox("Ano", [2025, 2026, 2027], index=1, key=f"ano_m_{key_prefix}")
        _, ult = calendar.monthrange(ano_sel, mes_sel)
        dia_ini = hoje.day if (mes_sel == hoje.month and ano_sel == hoje.year) else 1
        datas = [(date(ano_sel, mes_sel, d), date(ano_sel, mes_sel, d) + timedelta(days=1)) for d in range(dia_ini, ult + 1)]
        st.sidebar.info(f"📋 {len(datas)} dias — {dia_ini:02d}/{mes_sel:02d} a {ult:02d}/{mes_sel:02d}/{ano_sel}")
        return datas

    elif modo == "🏖️ Finais de Semana":
        mes_sel = st.sidebar.selectbox("Mês",
                    list(range(1,13)),
                    format_func=lambda m: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m-1],
                    index=hoje.month - 1, key=f"mes_fs_{key_prefix}")
        ano_sel = st.sidebar.selectbox("Ano", [2025, 2026, 2027], index=1, key=f"ano_fs_{key_prefix}")
        _, ult = calendar.monthrange(ano_sel, mes_sel)
        # Monta períodos Sex→Seg (check-in sexta, check-out segunda)
        finais = []
        for d in range(1, ult + 1):
            dt = date(ano_sel, mes_sel, d)
            if dt.weekday() == 4:  # Sexta
                check_in  = dt
                check_out = dt + timedelta(days=2)  # Dom
                finais.append((check_in, check_out))
        # Exibir resumo
        resumo = "  \n".join([f"• {ci.strftime('%d/%m')} (sex) → {co.strftime('%d/%m')} (dom)" for ci, co in finais])
        st.sidebar.info(f"📋 {len(finais)} final(is) de semana:\n{resumo}")
        return finais

    elif modo == "🎉 Feriados":
        from datetime import date as _date
        FERIADOS_PADRAO = [
            {"nome": "Páscoa",            "cin": _date(2026, 4,  2), "cout": _date(2026, 4,  6)},
            {"nome": "Tiradentes",        "cin": _date(2026, 4, 21), "cout": _date(2026, 4, 22)},
            {"nome": "Dia do Trabalho",   "cin": _date(2026, 5,  1), "cout": _date(2026, 5,  2)},
            {"nome": "Corpus Christi",    "cin": _date(2026, 6,  4), "cout": _date(2026, 6,  7)},
            {"nome": "Independência",     "cin": _date(2026, 9,  5), "cout": _date(2026, 9,  7)},
            {"nome": "N. Sra. Aparecida", "cin": _date(2026, 10, 10), "cout": _date(2026, 10, 12)},
            {"nome": "Finados",           "cin": _date(2026, 11,  1), "cout": _date(2026, 11,  2)},
            {"nome": "Proclamação",       "cin": _date(2026, 11, 14), "cout": _date(2026, 11, 16)},
            {"nome": "Natal",             "cin": _date(2026, 12, 24), "cout": _date(2026, 12, 27)},
            {"nome": "Réveillon",         "cin": _date(2026, 12, 31), "cout": _date(2027,  1,  2)},
        ]
        sk = f"feriados_{key_prefix}"
        if sk not in st.session_state:
            st.session_state[sk] = [dict(f) for f in FERIADOS_PADRAO]
        st.session_state[sk] = [f for f in st.session_state[sk] if f["cout"] >= hoje]
        feriados = st.session_state[sk]
        st.sidebar.markdown("**✏️ Editar feriados:**")
        for i, f in enumerate(feriados):
            st.sidebar.markdown(f"**{f['nome']}**")
            f["nome"] = st.sidebar.text_input("Nome", f["nome"], key=f"fn_{key_prefix}_{i}", label_visibility="collapsed")
            col_ci, col_co = st.sidebar.columns(2)
            f["cin"]  = col_ci.date_input("Check-in",  f["cin"],  key=f"fci_{key_prefix}_{i}", format="DD/MM/YYYY")
            f["cout"] = col_co.date_input("Check-out", f["cout"], key=f"fco_{key_prefix}_{i}", format="DD/MM/YYYY")
            if st.sidebar.button("🗑️ Remover", key=f"fdel_{key_prefix}_{i}", use_container_width=True):
                st.session_state[sk].pop(i)
                st.rerun()
            st.sidebar.markdown("---")
        col1, col2 = st.sidebar.columns(2)
        if col1.button("➕ Adicionar", key=f"fadd_{key_prefix}", use_container_width=True):
            st.session_state[sk].append({"nome": "Novo feriado", "cin": hoje, "cout": hoje + timedelta(days=1)})
            st.rerun()
        if col2.button("🔄 Restaurar", key=f"freset_{key_prefix}", use_container_width=True):
            st.session_state[sk] = [dict(f) for f in FERIADOS_PADRAO if f["cout"] >= hoje]
            st.rerun()
        datas = [(f["cin"], f["cout"]) for f in feriados if f["cin"] < f["cout"]]
        resumo = "  \n".join([f"• {f['nome']}: {f['cin'].strftime('%d/%m')} → {f['cout'].strftime('%d/%m')}" for f in feriados if f["cin"] < f["cout"]])
        st.sidebar.info(f"📋 {len(datas)} feriado(s):\n{resumo}")
        return datas

    elif modo == "📋 Semana (2-3-4 noites)":
        FERIADOS_EXCLUIR = [
            date(2026, 4,  2), date(2026, 4,  3), date(2026, 4,  4), date(2026, 4,  5), date(2026, 4,  6),
            date(2026, 4, 21), date(2026, 5,  1),
            date(2026, 6,  4), date(2026, 6,  5), date(2026, 6,  6), date(2026, 6,  7),
            date(2026, 9,  5), date(2026, 9,  6), date(2026, 9,  7),
            date(2026, 10, 10), date(2026, 10, 11), date(2026, 10, 12),
            date(2026, 11,  1), date(2026, 11,  2),
            date(2026, 11, 14), date(2026, 11, 15), date(2026, 11, 16),
            date(2026, 12, 24), date(2026, 12, 25), date(2026, 12, 26), date(2026, 12, 27),
            date(2026, 12, 31), date(2027,  1,  1),
        ]
        mes_sel = st.sidebar.selectbox("Mês", list(range(1, 13)),
                    format_func=lambda m: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][m-1],
                    index=hoje.month - 1, key=f"mes_sem_{key_prefix}")
        ano_sel = st.sidebar.selectbox("Ano", [2025, 2026, 2027], index=1, key=f"ano_sem_{key_prefix}")
        _, ult = calendar.monthrange(ano_sel, mes_sel)
        segundas_validas = []
        for d in range(1, ult + 1):
            dt = date(ano_sel, mes_sel, d)
            if dt.weekday() != 0: continue
            semana = [dt + timedelta(days=x) for x in range(5)]
            if any(dia in FERIADOS_EXCLUIR for dia in semana): continue
            segundas_validas.append(dt)
        if not segundas_validas:
            st.sidebar.warning("⚠️ Nenhuma segunda sem feriado neste mês.")
            return []
        opcoes = {seg.strftime("%d/%m/%Y (%A)").replace("Monday","Segunda"): seg for seg in segundas_validas}
        escolha = st.sidebar.selectbox("Segunda-feira:", list(opcoes.keys()), key=f"seg_{key_prefix}")
        seg_escolhida = opcoes[escolha]
        sex = seg_escolhida + timedelta(days=4)
        datas = [
            (seg_escolhida,                     sex),
            (seg_escolhida + timedelta(days=1), sex),
            (seg_escolhida + timedelta(days=2), sex),
        ]
        nomes = ["4 noites (seg→sex)", "3 noites (ter→sex)", "2 noites (qua→sex)"]
        resumo = "  \n".join([f"• {n}: {ci.strftime('%d/%m')} → {co.strftime('%d/%m')}" for n, (ci, co) in zip(nomes, datas)])
        st.sidebar.info(f"📋 Semana de {seg_escolhida.strftime('%d/%m/%Y')} — 3 buscas:\n{resumo}")
        return datas

    else:  # Múltiplos Períodos
        if f"periodos_{key_prefix}" not in st.session_state:
            st.session_state[f"periodos_{key_prefix}"] = [
                {"ini": hoje, "fim": hoje + timedelta(days=1), "tipo": "dia_a_dia"}
            ]
        periodos = st.session_state[f"periodos_{key_prefix}"]

        st.sidebar.markdown("**Períodos adicionados:**")
        to_remove = []
        for i, p in enumerate(periodos):
            st.sidebar.markdown(f"**Período {i+1}**")
            ini = st.sidebar.date_input(f"Início {i+1}", p["ini"], key=f"pini_{key_prefix}_{i}", format="DD/MM/YYYY")
            fim = st.sidebar.date_input(f"Fim {i+1}",    p["fim"], key=f"pfim_{key_prefix}_{i}", format="DD/MM/YYYY")
            periodos[i]["ini"] = ini
            periodos[i]["fim"] = fim

            # Toggle: Dia a dia ou Período inteiro
            tipo_atual = p.get("tipo", "dia_a_dia")
            tipo_sel = st.sidebar.radio(
                f"Busca do período {i+1}:",
                ["📅 Dia a dia", "📆 Período inteiro"],
                index=0 if tipo_atual == "dia_a_dia" else 1,
                key=f"tipo_{key_prefix}_{i}",
                horizontal=True
            )
            periodos[i]["tipo"] = "dia_a_dia" if tipo_sel == "📅 Dia a dia" else "periodo_inteiro"

            if st.sidebar.button(f"❌ Remover período {i+1}", key=f"del_{key_prefix}_{i}", use_container_width=True):
                to_remove.append(i)
            st.sidebar.markdown("---")

        for i in reversed(to_remove):
            periodos.pop(i)
            st.rerun()

        if st.sidebar.button("➕ Adicionar Período", key=f"add_{key_prefix}", use_container_width=True):
            periodos.append({"ini": hoje, "fim": hoje + timedelta(days=1), "tipo": "dia_a_dia"})
            st.rerun()

        # Gera lista de datas de cada período conforme o tipo escolhido
        datas = []
        resumo_linhas = []
        for p in periodos:
            ini, fim = p["ini"], p["fim"]
            tipo = p.get("tipo", "dia_a_dia")
            if ini > fim:
                continue
            if tipo == "dia_a_dia":
                # Uma busca por dia: check-in D, check-out D+1
                # Ex: 11 a 15 → (11→12), (12→13), (13→14), (14→15)
                d = ini
                count = 0
                while d < fim:
                    datas.append((d, d + timedelta(days=1)))
                    d += timedelta(days=1)
                    count += 1
                resumo_linhas.append(f"• {ini.strftime('%d/%m')} a {fim.strftime('%d/%m')} — {count} dia(s) 📅")
            else:
                # Período inteiro: check-in = ini, check-out = fim
                # Ex: 11 a 15 → uma busca (11→15)
                datas.append((ini, fim))
                resumo_linhas.append(f"• {ini.strftime('%d/%m')} → {fim.strftime('%d/%m')} — período único 📆")

        resumo = "  \n".join(resumo_linhas)
        st.sidebar.info(f"📋 {len(periodos)} período(s) → {len(datas)} busca(s)\n{resumo}")
        return datas
# ==============================================================================
# SHOPPER ALTO DA BOA VISTA — funções especiais (quartos + nome)
# ==============================================================================
def _buscar_quartos_hotel_booking(driver, wait, url_hotel, n=5):
    import time as _t
    quartos = []
    nomes_vistos = set()
    precos_vistos = set()
    try:
        driver.get(url_hotel)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'tr.js-rt-block-row')))
        _t.sleep(2)
        for linha in driver.find_elements(By.CSS_SELECTOR, 'tr.js-rt-block-row'):
            if len(quartos) >= n:
                break
            nome = ""
            try:
                nome = " ".join(linha.find_element(By.CSS_SELECTOR,
                    'span.hprt-roomtype-icon-link').text.strip().split())
            except:
                continue
            if not nome or nome.lower() in nomes_vistos:
                continue
            preco = "Esgotado"
            try:
                el = linha.find_element(By.CSS_SELECTOR,
                    'div.bui-price-display__value span.prco-valign-middle-helper')
                nums = re.findall(r'\d+', el.text.strip().replace('.', '').replace(',', ''))
                if nums:
                    preco = int(nums[-1])
            except:
                try:
                    opt = linha.find_element(By.CSS_SELECTOR,
                        'select.hprt-nos-select option[value="1"]')
                    nums = re.findall(r'\d+', opt.text.replace('.', '').replace(',', ''))
                    if nums:
                        preco = int(nums[-1])
                except:
                    pass
            if isinstance(preco, int) and preco in precos_vistos:
                continue
            nomes_vistos.add(nome.lower())
            if isinstance(preco, int):
                precos_vistos.add(preco)
            quartos.append({"nome": nome, "preco": preco})
    except:
        pass
    return quartos[:n]

def executar_varredura_alto(datas_para_busca, concorrentes, sua_col, slug_alto):
    import time as _t
    SLUGS_CONC = {
        "Pousada Villa Capivary Campos do Jordão": "pousada-villa-capivary",
        "Pousada Da Pedra":                        "pousada-da-pedra",
        "Villa Amistà Campos do Jordão":           "villa-amista-campos-do-jordao",
        "Pousada Boutique Figueira da Serra":      "pousada-figueira-da-serra",
        "L.A.H. Hostellerie":                      "l-a-h-hostellerie",
        "Carballo Hotel & Spa":                    "carballo-amp-spa",
        "Hotel Boutique QUEBRA-NOZ":               "quebra-noz",
    }
    driver = criar_driver()
    wait   = WebDriverWait(driver, 12)
    status = st.empty()
    prog   = st.progress(0)
    outros = [c for c in concorrentes if c != sua_col]
    total  = len(datas_para_busca) * (len(outros) + 1)
    cont   = 0
    todas  = []
    for d_in, d_out in datas_para_busca:
        label = f"{d_in.strftime('%d/%m/%Y')} → {d_out.strftime('%d/%m/%Y')}"
        linha = {"Data": label}
        for hotel in outros:
            status.markdown(f"📡 **{label}** | {hotel[:35]}...")
            slug = SLUGS_CONC.get(hotel, "")
            url  = (f"https://www.booking.com/hotel/br/{slug}.pt-br.html"
                    f"?checkin={d_in.isoformat()}&checkout={d_out.isoformat()}"
                    f"&group_adults=2&no_rooms=1&selected_currency=BRL") if slug else (
                   f"https://www.booking.com/searchresults.pt-br.html"
                   f"?ss={hotel.replace(' ', '+')}"
                   f"&checkin={d_in.isoformat()}&checkout={d_out.isoformat()}"
                   f"&group_adults=2&no_rooms=1&selected_currency=BRL&lang=pt-br")
            driver.get(url)
            preco = "Esgotado"
            nome_conc = ""
            try:
                if slug:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'tr.js-rt-block-row')))
                    melhor = None
                    for lc in driver.find_elements(By.CSS_SELECTOR, 'tr.js-rt-block-row'):
                        try:
                            pe = lc.find_element(By.CSS_SELECTOR,
                                'div.bui-price-display__value span.prco-valign-middle-helper')
                            nums = re.findall(r'\d+', pe.text.strip().replace('.','').replace(',',''))
                            if nums:
                                v = int(nums[-1])
                                if melhor is None or v < melhor:
                                    melhor = v
                                    try:
                                        nome_conc = " ".join(lc.find_element(By.CSS_SELECTOR,
                                            'span.hprt-roomtype-icon-link').text.strip().split())
                                    except:
                                        pass
                        except:
                            pass
                    if melhor:
                        preco = melhor
                else:
                    wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, '[data-testid="property-card"]')))
                    card = driver.find_element(By.CSS_SELECTOR, '[data-testid="property-card"]')
                    pe   = card.find_element(By.CSS_SELECTOR,
                        '[data-testid="price-and-discounted-price"]')
                    nums = re.findall(r'\d+', pe.text.replace('.','').replace(',',''))
                    if nums:
                        preco = int(nums[-1])
            except:
                pass
            linha[hotel] = preco
            linha[f"{hotel}__quarto"] = nome_conc
            cont += 1
            prog.progress(cont / total)
        status.markdown(f"🏨 **{label}** | Alto da Boa Vista — buscando 5 quartos...")
        url_alto = (f"https://www.booking.com/hotel/br/{slug_alto}.pt-br.html"
                    f"?checkin={d_in.isoformat()}&checkout={d_out.isoformat()}"
                    f"&group_adults=2&no_rooms=1&selected_currency=BRL&lang=pt-br")
        quartos = _buscar_quartos_hotel_booking(driver, wait, url_alto, n=5)
        for i, q in enumerate(quartos):
            linha[f"Alto_Q{i+1}_nome"]  = q["nome"]
            linha[f"Alto_Q{i+1}_preco"] = q["preco"]
        for i in range(len(quartos), 5):
            linha[f"Alto_Q{i+1}_nome"]  = "Esgotado"
            linha[f"Alto_Q{i+1}_preco"] = "Esgotado"
        cont += 1
        prog.progress(cont / total)
        todas.append(linha)
    driver.quit()
    status.success("✅ Varredura concluída!")
    df = pd.DataFrame(todas)
    historico_salvar(funcao="Shopper Alto da Boa Vista",
                     status="✅ Sucesso",
                     detalhes={"total_datas": len(datas_para_busca)})
    return df

def exibir_tabela_alto(df, cols_conc, key_download):
    if df is None or df.empty:
        return
    for _, row in df.iterrows():
        st.markdown(f"#### 📅 {row['Data']}")
        col_esq, col_dir = st.columns([1, 1])
        with col_esq:
            st.markdown("**Concorrentes**")
            dados = []
            for c in cols_conc:
                val       = row.get(c, "—")
                nome_q    = row.get(f"{c}__quarto", "")
                preco_str = f"R$ {int(val):,}".replace(",", ".") if isinstance(val, (int, float)) else str(val)
                label_c   = f"{c}  ({nome_q})" if nome_q else c
                dados.append({"Hotel": label_c, "Valor": preco_str})
            if dados:
                st.dataframe(pd.DataFrame(dados).set_index("Hotel"), use_container_width=True)
        with col_dir:
            st.markdown("**Alto da Boa Vista**")
            dados_alto = []
            for i in range(1, 6):
                nome  = row.get(f"Alto_Q{i}_nome",  "Esgotado")
                preco = row.get(f"Alto_Q{i}_preco", "Esgotado")
                preco_str = f"R$ {int(preco):,}".replace(",", ".") if isinstance(preco, (int, float)) else str(preco)
                dados_alto.append({"Quarto": nome, "Valor": preco_str})
            if dados_alto:
                st.dataframe(pd.DataFrame(dados_alto).set_index("Quarto"), use_container_width=True)
        st.markdown("---")
    excel_file = _gerar_excel_alto(df, cols_conc)
    st.download_button(label="📥 Baixar Planilha", data=excel_file,
        file_name=f"shopper_alto_{date.today().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key_download)

def _gerar_excel_alto(df, cols_conc):
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    output    = io.BytesIO()
    cyan_fill = PatternFill("solid", fgColor="00B0F0")
    azul_fill = PatternFill("solid", fgColor="002D62")
    rosa_fill = PatternFill("solid", fgColor="FF9999")
    f_cyan  = Font(bold=True, color="000000", name="Calibri", size=11)
    f_azul  = Font(bold=True, color="FFFFFF",  name="Calibri", size=11)
    f_esgot = Font(bold=True, color="C62828",  name="Calibri", size=11)
    f_norm  = Font(name="Calibri", size=11)
    alin_c  = Alignment(horizontal="center", vertical="center")
    alin_e  = Alignment(horizontal="left",   vertical="center")
    alin_d  = Alignment(horizontal="right",  vertical="center")
    borda   = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin",  color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"))
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet(title="Shopper Alto da Boa Vista")
        linha_atual = 1
        for idx, (_, row) in enumerate(df.iterrows()):
            if idx > 0:
                linha_atual += 1
            label = row["Data"]
            try:
                partes = label.split("→")
                d1 = datetime.strptime(partes[0].strip(), "%d/%m/%Y")
                d2 = datetime.strptime(partes[1].strip(), "%d/%m/%Y")
                noites = (d2 - d1).days
                header = f"{d1.strftime('%d')} a {d2.strftime('%d/%m')} - {noites} noite{'s' if noites > 1 else ''}"
            except:
                header = label
            c = ws.cell(row=linha_atual, column=1, value=header)
            c.fill=cyan_fill; c.font=f_cyan; c.alignment=alin_c; c.border=borda
            c = ws.cell(row=linha_atual, column=2, value="valor")
            c.fill=azul_fill; c.font=f_azul; c.alignment=alin_c; c.border=borda
            ws.cell(row=linha_atual, column=3).border = borda
            c = ws.cell(row=linha_atual, column=4, value="Alto Da Boa Vista")
            c.fill=azul_fill; c.font=f_azul; c.alignment=alin_c; c.border=borda
            ws.merge_cells(start_row=linha_atual, start_column=4,
                           end_row=linha_atual,   end_column=6)
            linha_atual += 1
            for ri, hotel in enumerate(cols_conc):
                lc      = linha_atual + ri
                nome_q  = row.get(f"{hotel}__quarto", "")
                label_c = f"{hotel}  ({nome_q})" if nome_q else hotel
                c = ws.cell(row=lc, column=1, value=label_c)
                c.font=f_norm; c.alignment=alin_e; c.border=borda
                val = row.get(hotel, "Esgotado")
                c = ws.cell(row=lc, column=2)
                c.border=borda; c.alignment=alin_c
                if val == "Esgotado" or not isinstance(val, (int, float)):
                    c.value = "Esgotado"; c.fill = rosa_fill; c.font = f_esgot
                else:
                    c.value = val; c.number_format = "R$ #,##0"; c.font = f_norm
                ws.cell(row=lc, column=3).border = borda
                for col in [4, 5, 6]:
                    ws.cell(row=lc, column=col).border = borda
            for qi in range(1, 6):
                lq    = linha_atual + qi - 1
                nome  = row.get(f"Alto_Q{qi}_nome",  "Esgotado")
                preco = row.get(f"Alto_Q{qi}_preco", "Esgotado")
                c = ws.cell(row=lq, column=4); c.border=borda; c.alignment=alin_d
                if nome == "Esgotado" or preco == "Esgotado":
                    c.value = "Esgotado"; c.fill = rosa_fill; c.font = f_esgot
                    ws.cell(row=lq, column=5).border = borda
                    ws.cell(row=lq, column=6).border = borda
                else:
                    c.value = preco; c.number_format = "R$ #,##0"; c.font = f_norm
                    c2 = ws.cell(row=lq, column=5, value=" - ")
                    c2.font=f_norm; c2.alignment=alin_c; c2.border=borda
                    c3 = ws.cell(row=lq, column=6, value=nome)
                    c3.font=f_norm; c3.alignment=alin_e; c3.border=borda
            linha_atual += len(cols_conc)
        ws.column_dimensions["A"].width = 32
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 3
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 4
        ws.column_dimensions["F"].width = 42
        ws.freeze_panes = "A2"
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    output.seek(0)
    return output

if aba_selecionada == "🎯 Shopper Alto da Boa Vista":
    st.title("🎯 MONITORAMENTO DE SHOPPER")
    st.subheader("Pousada Alto Da Boa Vista Campos do Jordão")
    CONCORRENTES_ALVO = [
        "Pousada Villa Capivary Campos do Jordão",
        "Pousada Da Pedra",
        "Villa Amistà Campos do Jordão",
        "Pousada Boutique Figueira da Serra",
        "L.A.H. Hostellerie",
        "Carballo Hotel & Spa",
        "Hotel Boutique QUEBRA-NOZ",
    ]
    SLUG_ALTO = "pousada-alto-da-boa-vista"
    datas_para_busca = _shopper_selecionar_datas("s1")
    if st.button("🚀 INICIAR VARREDURA", key="btn1"):
        st.session_state.dados_busca = None
        try:
            st.session_state.dados_busca = executar_varredura_alto(
                datas_para_busca, CONCORRENTES_ALVO,
                "Pousada Alto Da Boa Vista Campos do Jordão", SLUG_ALTO
            )
        except Exception as e:
            st.error(f"Erro: {e}")
    if st.session_state.dados_busca is not None:
        exibir_tabela_alto(st.session_state.dados_busca, CONCORRENTES_ALVO, "shopper_alto_boa_vista")
# ==============================================================================
# ABA 2: SHOPPER VILLA CARAVAGGIO
# ==============================================================================
elif aba_selecionada == "🏕️ Shopper Villa Caravaggio":
    st.title("🏕️ SHOPPER VILLA CARAVAGGIO")
    st.subheader("Chalés Villa Caravaggio - by Easy Hotéis")

    CONCORRENTES_CARAVAGGIO = [
        "Chalés Villa Caravaggio - by Easy Hotéis",
        "Pousada Rabo do Lagarto",
        "Glamping Pedra Azul",
        "Cabana Raposo",
    ]

    URLS_CARAVAGGIO = {
        "Chalés Villa Caravaggio - by Easy Hotéis": "chales-villa-caravaggio",
        "Pousada Rabo do Lagarto":                  "pousada-rabo-do-lagarto",
        "Glamping Pedra Azul":                      "glamping-pedra-azul-domingos-martins1",
        "Cabana Raposo":                            "cabana-raposo-em-santa-teresa",
    }

    SUA_COL_CAR = "Chalés Villa Caravaggio - by Easy Hotéis"

    datas_para_busca2 = _shopper_selecionar_datas("s2")

    if st.button("🚀 INICIAR VARREDURA", key="btn2"):
        st.session_state.dados_busca_caravaggio = None
        try:
            st.session_state.dados_busca_caravaggio = executar_varredura(datas_para_busca2, CONCORRENTES_CARAVAGGIO)
        except Exception as e:
            st.error(f"Erro: {e}")

    if st.session_state.dados_busca_caravaggio is not None:
        exibir_tabela_resultados(st.session_state.dados_busca_caravaggio, CONCORRENTES_CARAVAGGIO, SUA_COL_CAR, "shopper_villa_caravaggio")

# ==============================================================================
# ABA 3: SHOPPER HAMBURGO
# ==============================================================================
elif aba_selecionada == "🍺 Shopper Hamburgo":
    st.title("🍺 SHOPPER HAMBURGO")
    st.subheader("Hamburgo Palace Hotel - Balneário Camboriú")

    CONCORRENTES_HAMBURGO = [
        "Hamburgo Palace Hotel",
        "Hotel HM",
        "Camboriú Praia",
        "San Marino Cassino Hotel",
        "Bella Camboriú",
        "Ibis Balneario Camboriu",
    ]

    SUA_COL_HAM = "Hamburgo Palace Hotel"

    datas_para_busca3 = _shopper_selecionar_datas("s3")

    if st.button("🚀 INICIAR VARREDURA", key="btn3"):
        st.session_state.dados_busca_hamburgo = None
        try:
            st.session_state.dados_busca_hamburgo = executar_varredura(datas_para_busca3, CONCORRENTES_HAMBURGO)
        except Exception as e:
            st.error(f"Erro: {e}")

    if st.session_state.dados_busca_hamburgo is not None:
        exibir_tabela_resultados(st.session_state.dados_busca_hamburgo, CONCORRENTES_HAMBURGO, SUA_COL_HAM, "shopper_hamburgo")

# ==============================================================================
# ABA 4: SHOPPER TERRAZZO
# ==============================================================================
elif aba_selecionada == "🏨 Shopper Terrazzo":
    st.title("🏨 SHOPPER TERRAZZO")
    st.subheader("Terrazzo Bonjardim - Campos do Jordão")

    CONCORRENTES_TERRAZZO = [
        "Terrazzo Bonjardim",
        "Hotel Moinho Itália",
        "Hotel Vila Dom Bosco",
        "Altitude Lodge Hotel",
        "La Vita Pousada de Charme",
    ]

    URLS_TERRAZZO = {
        "Terrazzo Bonjardim":         "pousada-terrazzo-bonjardim-campos-do-jordao",
        "Hotel Moinho Itália":        "o-moinho",
        "Hotel Vila Dom Bosco":       "vila-dom-bosco",
        "Altitude Lodge Hotel":       "altitude-lodge",
        "La Vita Pousada de Charme":  "la-vie-pousada-de-charme",
    }

    SUA_COL_TER = "Terrazzo Bonjardim"

    datas_para_busca4 = _shopper_selecionar_datas("s4")

    if st.button("🚀 INICIAR VARREDURA", key="btn4"):
        st.session_state.dados_busca_terrazzo = None
        try:
            st.session_state.dados_busca_terrazzo = executar_varredura(datas_para_busca4, CONCORRENTES_TERRAZZO)
        except Exception as e:
            st.error(f"Erro: {e}")

    if st.session_state.dados_busca_terrazzo is not None:
        exibir_tabela_resultados(st.session_state.dados_busca_terrazzo, CONCORRENTES_TERRAZZO, SUA_COL_TER, "shopper_terrazzo")

# ==============================================================================
# ABA 5: SHOPPER SERRA NEGRA
# ==============================================================================
elif aba_selecionada == "🏖️ Shopper Serra Negra":
    st.title("🏖️ SHOPPER SERRA NEGRA")
    st.subheader("Serra Negra Pousada & Spa")

    CONCORRENTES_SERRA_NEGRA = [
        "Serra Negra Pousada Spa - by Easy Hotéis",
        "Duas Praias Hotel Pousada",
        "Hotel Pousada Caminho da Praia",
        "Hotel Diamantina - em Guarapari",
    ]

    SLUGS_SERRA_NEGRA = {
        "Serra Negra Pousada Spa - by Easy Hotéis": "serra-negra-pousada-spa",
        "Duas Praias Hotel Pousada":                "pousada-duas-praias",
        "Hotel Pousada Caminho da Praia":           "pousada-caminho-da-praia-guarapari",
        "Hotel Diamantina - em Guarapari":          "diamantina-guarapari",
    }

    SUA_COL_SN = "Serra Negra Pousada Spa - by Easy Hotéis"

    datas_para_busca5 = _shopper_selecionar_datas("s5")

    if st.button("🚀 INICIAR VARREDURA", key="btn5"):
        st.session_state.dados_busca_serra_negra = None
        try:
            st.session_state.dados_busca_serra_negra = executar_varredura(datas_para_busca5, CONCORRENTES_SERRA_NEGRA)
        except Exception as e:
            st.error(f"Erro: {e}")

    if st.session_state.dados_busca_serra_negra is not None:
        exibir_tabela_resultados(st.session_state.dados_busca_serra_negra, CONCORRENTES_SERRA_NEGRA, SUA_COL_SN, "shopper_serra_negra")

# ==============================================================================
# ABA 11: SHOPPER CANOEIROS
# ==============================================================================
elif aba_selecionada == "🏞️ Shopper Canoeiros":
    st.title("🏞️ SHOPPER CANOEIROS")
    st.subheader("Hotel Canoeiros")

    CONCORRENTES_CANOEIROS = [
        "Hotel Canoeiros",
        "Pousada Grande Rio",
        "Hotel Mundial Pirapora",
        "Villacoco",
    ]

    SUA_COL_CANOEIROS = "Hotel Canoeiros"

    datas_para_busca_canoeiros = _shopper_selecionar_datas("canoeiros")

    if st.button("🚀 INICIAR VARREDURA", key="btn_canoeiros"):
        st.session_state.dados_busca_canoeiros = None
        try:
            st.session_state.dados_busca_canoeiros = executar_varredura(
                datas_para_busca_canoeiros,
                CONCORRENTES_CANOEIROS
            )
        except Exception as e:
            st.error(f"Erro: {e}")

    if st.session_state.dados_busca_canoeiros is not None:
        exibir_tabela_resultados(
            st.session_state.dados_busca_canoeiros,
            CONCORRENTES_CANOEIROS,
            SUA_COL_CANOEIROS,
            "shopper_canoeiros"
        )

# ==============================================================================
# ABA 6: HITS - ALTO DA BOA VISTA (COMPARAÇÃO ANO VS ANO)
# ==============================================================================
elif aba_selecionada == "📈 Hits - Alto da Boa Vista":
    st.title("📈 HITS - HISTÓRICO E PREVISÃO")
    st.subheader("Pousada Alto Da Boa Vista")

    URL_HITS = "https://altodaboavista.hitspms.net/#/home"
    hoje = date.today()
    data_ini = st.sidebar.date_input("📅 Início do Período", date(hoje.year, hoje.month, 1), key="hits_ini", format="DD/MM/YYYY")
    data_fim = st.sidebar.date_input("📅 Fim do Período", date(hoje.year, hoje.month, hoje.day), key="hits_fim", format="DD/MM/YYYY")

    data_ini_ant = date(data_ini.year - 1, data_ini.month, data_ini.day)
    data_fim_ant = date(data_fim.year - 1, data_fim.month, data_fim.day)
    st.sidebar.markdown(f"📅 **Ano anterior:** {data_ini_ant.strftime('%d/%m/%Y')} a {data_fim_ant.strftime('%d/%m/%Y')}")

    def limpar_pct(val):
        if not val: return 0.0
        import re as re2
        match = re2.search(r'\(([\d,.]+)%?\)', str(val))
        if match: return float(match.group(1).replace(',', '.'))
        val = re2.sub(r'[^\d,.]', '', str(val).replace('%', ''))
        val = val.replace(',', '.')
        try: return float(val)
        except: return 0.0

    def limpar_valor(val):
        if not val or val in ['-', '', 'N/A', '$0,00', '$0.00']: return 0.0
        import re as re2
        val = re2.sub(r'[^\d,.]', '', str(val))
        if ',' in val and '.' in val: val = val.replace('.', '').replace(',', '.')
        elif ',' in val: val = val.replace(',', '.')
        try: return float(val)
        except: return 0.0

    def buscar_hits(d_ini, d_fim, d_ini_ant, d_fim_ant):
        from selenium.webdriver.common.keys import Keys
        driver = criar_driver()
        wait = WebDriverWait(driver, 15)
        status = st.empty()

        def selecionar_data_e_extrair(d_i, d_f, label):
            status.markdown(f"📅 Configurando datas {label}...")
            try:
                wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".block-ui-overlay")))
            except:
                time.sleep(1)
            try:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "#one-search-filters-container > div.active-filters > button:nth-child(1)"
                ))).click()
            except:
                driver.execute_script("var btn = document.querySelector('#one-search-filters-container button'); if (btn) btn.click();")
            time.sleep(1)
            try:
                campo_data = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                    "#one-search-modal-content > div > div > input")))
                campo_data.clear()
                campo_data.send_keys(f"{d_i.strftime('%d/%m/%y')} - {d_f.strftime('%d/%m/%y')}")
                time.sleep(0.3)
                campo_data.send_keys(Keys.RETURN)
                time.sleep(1)
            except: pass
            try:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "body > div.modal.fade.in > div > div > div.modal-footer.one-search-modal-footer > button > em"
                ))).click()
            except:
                driver.execute_script("var btns = document.querySelectorAll('.modal-footer button'); if (btns.length > 0) btns[0].click();")
            time.sleep(2.5)
            status.markdown(f"📥 Lendo dados {label}...")
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                    "#historyAndForecastReport > div:nth-child(1) > history-and-forecast-report-list > div.history-and-forecast-of-revenue-and-occupation-report-container > div.table-max-width > div"
                )))
            except:
                time.sleep(2)
            time.sleep(1.5)
            return driver.execute_script("""
                var resultado = [];
                var tabelas = document.querySelectorAll('#historyAndForecastReport table');
                tabelas.forEach(function(tabela) {
                    var linhas = tabela.querySelectorAll('tbody tr');
                    linhas.forEach(function(tr) {
                        var colunas = tr.querySelectorAll('td');
                        if (colunas.length > 5) {
                            var linha = [];
                            colunas.forEach(function(td) { linha.push(td.innerText.trim()); });
                            resultado.push(linha);
                        }
                    });
                });
                return resultado;
            """)

        try:
            status.markdown("🔄 Abrindo Hits...")
            driver.get(URL_HITS)
            time.sleep(3)
            driver.refresh()
            time.sleep(5)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "header")))
                time.sleep(1)
            except:
                time.sleep(2)

            try:
                campo_email = driver.find_element(By.CSS_SELECTOR, "#Email")
                campo_email.clear()
                campo_email.send_keys("daniel@easyhoteis.com")
                campo_senha = driver.find_element(By.CSS_SELECTOR, "#Password")
                campo_senha.clear()
                campo_senha.send_keys("@Livia92")
                driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
                status.markdown("🔐 Fazendo login...")
                time.sleep(5)
                try: wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".block-ui-overlay")))
                except: time.sleep(1)
            except:
                try: wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".block-ui-overlay")))
                except: time.sleep(1)

            status.markdown("🔍 Abrindo relatório...")
            try:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "body > div:nth-child(4) > div > header > div.col-xs-12.no-padding.my-general > div.pull-right.float-none.box-conf-menu > a.nav-link.search-all-menu-icon.cursor-pointer.no-selection.pull-left.active > div > em"
                ))).click()
            except:
                driver.execute_script("var el = document.querySelector('a.search-all-menu-icon em'); if (el) el.click();")
            time.sleep(1)

            try:
                campo_busca = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                    "input[type='search'], input[type='text'][placeholder*='esquis'], .search-input input, input.search")))
                campo_busca.clear()
                campo_busca.send_keys("Histórico e previsão das receitas e ocupações")
            except:
                driver.execute_script("""
                    var inputs = document.querySelectorAll('input');
                    for (var i = 0; i < inputs.length; i++) {
                        if (inputs[i].offsetParent !== null) {
                            inputs[i].value = 'Histórico e previsão das receitas e ocupações';
                            inputs[i].dispatchEvent(new Event('input', {bubbles: true}));
                            break;
                        }
                    }
                """)
            time.sleep(1)

            try:
                resultado = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Histórico e previsão das receitas')]")))
                resultado.click()
            except:
                driver.execute_script("""
                    var els = document.querySelectorAll('*');
                    for (var i = 0; i < els.length; i++) {
                        if (els[i].innerText && els[i].innerText.indexOf('Histórico e previsão das receitas') >= 0 && els[i].children.length === 0) {
                            els[i].click(); break;
                        }
                    }
                """)
            time.sleep(1.5)

            dados_atual = selecionar_data_e_extrair(d_ini, d_fim, str(d_ini.year))
            status.success(f"✅ Dados {d_ini.year} carregados!")
            dados_anterior = selecionar_data_e_extrair(d_ini_ant, d_fim_ant, str(d_ini_ant.year))
            status.success(f"✅ Dados {d_ini_ant.year} carregados!")
            historico_salvar(
                funcao="PMS — Histórico e Previsão",
                status="✅ Sucesso",
                detalhes={
                    "hotel": "Alto da Boa Vista",
                    "data_ini": d_ini.strftime('%d/%m/%Y'),
                    "data_fim": d_fim.strftime('%d/%m/%Y'),
                    "mensagem": ""
                }
            )
            driver.quit()
            return dados_atual, dados_anterior

        except Exception as e:
            driver.quit()
            status.error(f"❌ Erro: {e}")
            historico_salvar(
                funcao="PMS — Histórico e Previsão",
                status=f"❌ Erro: {str(e)[:100]}",
                detalhes={"hotel": "Alto da Boa Vista"}
            )
            return None, None

    def dados_para_df(dados_raw):
        registros = []
        for linha in dados_raw:
            if len(linha) >= 15 and '/' in str(linha[0]):
                try:
                    registros.append({
                        'Data':       linha[0],
                        'Ocupação':   limpar_pct(linha[3]),
                        'Rec_Hosp':   limpar_valor(linha[14]) if len(linha) > 14 else 0.0,
                        'Diaria_Med': limpar_valor(linha[17]) if len(linha) > 17 else 0.0,
                        'RevPAR':     limpar_valor(linha[18]) if len(linha) > 18 else 0.0,
                        'Pax':        limpar_valor(linha[7])  if len(linha) > 7  else 0.0,
                    })
                except: pass
        return pd.DataFrame(registros) if registros else pd.DataFrame()

    def variacao(atual, anterior):
        if anterior == 0: return 0.0
        return ((atual - anterior) / anterior) * 100

    def seta(v):
        if v > 0: return f"▲ +{v:.1f}%"
        elif v < 0: return f"▼ {v:.1f}%"
        return "= 0%"

    if st.button("🚀 BUSCAR DADOS", key="btn_hits"):
        st.session_state.dados_hits_atual = None
        st.session_state.dados_hits_anterior = None
        dados_atual, dados_ant = buscar_hits(data_ini, data_fim, data_ini_ant, data_fim_ant)
        if dados_atual: st.session_state.dados_hits_atual = dados_atual
        if dados_ant: st.session_state.dados_hits_anterior = dados_ant

    if st.session_state.dados_hits_atual and st.session_state.dados_hits_anterior:
        df_at = dados_para_df(st.session_state.dados_hits_atual)
        df_an = dados_para_df(st.session_state.dados_hits_anterior)

        if not df_at.empty and not df_an.empty:
            ano_at = data_ini.year
            ano_an = data_ini_ant.year
            label_at = f"{data_ini.strftime('%d/%m')} a {data_fim.strftime('%d/%m')}/{ano_at}"
            label_an = f"{data_ini_ant.strftime('%d/%m')} a {data_fim_ant.strftime('%d/%m')}/{ano_an}"

            rec_at  = df_at['Rec_Hosp'].sum()
            rec_an  = df_an['Rec_Hosp'].sum()
            ocp_at  = df_at['Ocupação'].mean()
            ocp_an  = df_an['Ocupação'].mean()
            dm_at   = df_at[df_at['Diaria_Med'] > 0]['Diaria_Med'].mean()
            dm_an   = df_an[df_an['Diaria_Med'] > 0]['Diaria_Med'].mean()
            rev_at  = df_at[df_at['RevPAR'] > 0]['RevPAR'].mean()
            rev_an  = df_an[df_an['RevPAR'] > 0]['RevPAR'].mean()

            st.markdown("### 📊 Comparação do Período")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("💰 Receita Total", f"R$ {rec_at:,.2f}", f"{seta(variacao(rec_at, rec_an))} vs {ano_an}")
            col2.metric("🏨 Ocupação Média", f"{ocp_at:.1f}%", f"{seta(variacao(ocp_at, ocp_an))} vs {ano_an}")
            col3.metric("📈 Diária Média", f"R$ {dm_at:,.2f}", f"{seta(variacao(dm_at, dm_an))} vs {ano_an}")
            col4.metric("📊 RevPAR Médio", f"R$ {rev_at:,.2f}", f"{seta(variacao(rev_at, rev_an))} vs {ano_an}")
            st.markdown("---")

            st.markdown("### 📋 Resumo Comparativo")
            df_resumo = pd.DataFrame({
                'Indicador':    ['Receita Total', 'Ocupação Média', 'Diária Média', 'RevPAR Médio'],
                label_an:       [f"R$ {rec_an:,.2f}", f"{ocp_an:.1f}%", f"R$ {dm_an:,.2f}", f"R$ {rev_an:,.2f}"],
                label_at:       [f"R$ {rec_at:,.2f}", f"{ocp_at:.1f}%", f"R$ {dm_at:,.2f}", f"R$ {rev_at:,.2f}"],
                'Variação':     [seta(variacao(rec_at, rec_an)), seta(variacao(ocp_at, ocp_an)),
                                 seta(variacao(dm_at, dm_an)), seta(variacao(rev_at, rev_an))],
            })
            st.dataframe(df_resumo, use_container_width=True, hide_index=True)
            st.markdown("---")

            df_at['Ano'] = str(ano_at)
            df_an['Ano'] = str(ano_an)
            df_at['Dia'] = range(1, len(df_at) + 1)
            df_an['Dia'] = range(1, len(df_an) + 1)
            df_graf = pd.concat([df_at, df_an])

            fig_rec = px.bar(df_graf, x='Dia', y='Rec_Hosp', color='Ano', barmode='group',
                title="💰 Receita Diária — Comparação Ano vs Ano",
                labels={'Rec_Hosp': 'Receita (R$)', 'Dia': 'Dia do Período'},
                color_discrete_map={str(ano_at): '#002D62', str(ano_an): '#90CAF9'})
            fig_rec.update_layout(plot_bgcolor='white', paper_bgcolor='white',
                yaxis=dict(tickprefix='R$ ', tickformat=',.0f'))
            st.plotly_chart(fig_rec, use_container_width=True)

            fig_ocp = px.line(df_graf, x='Dia', y='Ocupação', color='Ano', markers=True,
                title="🏨 Ocupação — Comparação Ano vs Ano",
                labels={'Ocupação': 'Ocupação (%)', 'Dia': 'Dia do Período'},
                color_discrete_map={str(ano_at): '#D32F2F', str(ano_an): '#FFCDD2'})
            fig_ocp.update_layout(plot_bgcolor='white', paper_bgcolor='white',
                yaxis=dict(ticksuffix='%'))
            st.plotly_chart(fig_ocp, use_container_width=True)

            fig_dm = px.line(df_graf[df_graf['Diaria_Med'] > 0], x='Dia', y='Diaria_Med', color='Ano', markers=True,
                title="📈 Diária Média — Comparação Ano vs Ano",
                labels={'Diaria_Med': 'Diária Média (R$)', 'Dia': 'Dia do Período'},
                color_discrete_map={str(ano_at): '#002D62', str(ano_an): '#90CAF9'})
            fig_dm.update_layout(plot_bgcolor='white', paper_bgcolor='white',
                yaxis=dict(tickprefix='R$ ', tickformat=',.0f'))
            st.plotly_chart(fig_dm, use_container_width=True)

            fig_rev = px.line(df_graf[df_graf['RevPAR'] > 0], x='Dia', y='RevPAR', color='Ano', markers=True,
                title="📊 RevPAR — Comparação Ano vs Ano",
                labels={'RevPAR': 'RevPAR (R$)', 'Dia': 'Dia do Período'},
                color_discrete_map={str(ano_at): '#388E3C', str(ano_an): '#A5D6A7'})
            fig_rev.update_layout(plot_bgcolor='white', paper_bgcolor='white',
                yaxis=dict(tickprefix='R$ ', tickformat=',.0f'))
            st.plotly_chart(fig_rev, use_container_width=True)

            col_dl1, col_dl2 = st.columns(2)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resumo.to_excel(writer, index=False, sheet_name='Comparativo')
                df_at[['Data','Ocupação','Rec_Hosp','Diaria_Med','RevPAR','Pax']].to_excel(writer, index=False, sheet_name=f'Dados {ano_at}')
                df_an[['Data','Ocupação','Rec_Hosp','Diaria_Med','RevPAR','Pax']].to_excel(writer, index=False, sheet_name=f'Dados {ano_an}')
            output.seek(0)
            with col_dl1:
                st.download_button(label="📥 Baixar Excel", data=output,
                    file_name=f"hits_comparativo_{ano_an}_vs_{ano_at}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_hits_excel")

            try:
                import docx as docx_lib
                from docx.shared import Inches
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.enum.table import WD_TABLE_ALIGNMENT
                import tempfile
                with tempfile.TemporaryDirectory() as tmpdir:
                    img_rec = os.path.join(tmpdir, 'receita.png')
                    img_ocp = os.path.join(tmpdir, 'ocupacao.png')
                    img_dm  = os.path.join(tmpdir, 'diaria.png')
                    img_rev = os.path.join(tmpdir, 'revpar.png')
                    fig_rec.write_image(img_rec, width=900, height=400)
                    fig_ocp.write_image(img_ocp, width=900, height=400)
                    fig_dm.write_image(img_dm, width=900, height=400)
                    fig_rev.write_image(img_rev, width=900, height=400)
                    doc = docx_lib.Document()
                    t = doc.add_heading('HITS - HISTÓRICO E PREVISÃO', 0)
                    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    s = doc.add_heading('Pousada Alto Da Boa Vista', 2)
                    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p = doc.add_paragraph(f"Período: {label_an}  vs  {label_at}")
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    doc.add_paragraph()
                    doc.add_heading('Resumo Comparativo', 1)
                    tb = doc.add_table(rows=1, cols=4)
                    tb.style = 'Table Grid'
                    hdr = tb.rows[0].cells
                    for i, txt in enumerate(['Indicador', label_an, label_at, 'Variação']):
                        hdr[i].text = txt
                        hdr[i].paragraphs[0].runs[0].bold = True
                    for linha in [
                        ['Receita Total', f"R$ {rec_an:,.2f}", f"R$ {rec_at:,.2f}", seta(variacao(rec_at, rec_an))],
                        ['Ocupação Média', f"{ocp_an:.1f}%", f"{ocp_at:.1f}%", seta(variacao(ocp_at, ocp_an))],
                        ['Diária Média', f"R$ {dm_an:,.2f}", f"R$ {dm_at:,.2f}", seta(variacao(dm_at, dm_an))],
                        ['RevPAR Médio', f"R$ {rev_an:,.2f}", f"R$ {rev_at:,.2f}", seta(variacao(rev_at, rev_an))],
                    ]:
                        r = tb.add_row().cells
                        for i, v in enumerate(linha): r[i].text = v
                    doc.add_paragraph()
                    for tg, ip in [('Receita Diária', img_rec), ('Ocupação', img_ocp), ('Diária Média', img_dm), ('RevPAR', img_rev)]:
                        doc.add_heading(tg, 2)
                        doc.add_picture(ip, width=Inches(6))
                        doc.add_paragraph()
                    word_output = io.BytesIO()
                    doc.save(word_output)
                    word_output.seek(0)
                with col_dl2:
                    st.download_button(label="📄 Baixar Word (com gráficos)", data=word_output,
                        file_name=f"hits_comparativo_{ano_an}_vs_{ano_at}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_hits_word")
            except:
                with col_dl2:
                    st.warning("Word: instale `pip install python-docx kaleido`")
        else:
            st.warning("⚠️ Não foi possível extrair os dados. Tente novamente.")

# ==============================================================================
# ==============================================================================
# ABA 7: OMNIBEES
# ==============================================================================
elif aba_selecionada == "🐝 Omnibees":
    st.title("🐝 OMNIBEES")
    st.subheader("Gestão de Disponibilidade — Alto da Boa Vista")

    # ── Configuração do hotel ──────────────────────────────────────────────────
    HOTEIS_OMNIBEES = {
        "Alto da Boa Vista": {
            "login": "daniel.altodaboavista",
            "senha": "@Altodaboavista2026#",
            "tarifarios": [
                "Trf Bancorbras",
                "CLUBE MONTREAL",
                "Tarifa Flexível",
                "Condição Exclusiva",
                "Programa Preferencial - Orinter",
                "Tarifa Não Reembolsável.",
                "Tarifa Site",
            ],
            "quartos": [
                "Duplex Superior",
                "Suite Presidencial com Varanda",
                "Suíte Presidencial",
                "Quarto Duplo com Varanda",
                "Quarto Duplo Luxo com Sauna",
                "Chalé Exclusivo com Cozinha e Sala de Estar",
                "Chalé Família Exclusivo com Cozinha",
                "Chalé Duplex com Varanda",
                "Duplex com Piscina e lareira",
                "Chalé Duplex com Ofurô",
                "Suíte com Varanda",
                "Chalé Deluxe com Banheira e Lareira",
                "Quarto Duplo com Banheira",
                "Duplo com Lareira",
            ]
        }
    }

    # ── Tabela de BARs ─────────────────────────────────────────────────────────
    # Estrutura: { "NOME_BAR": { "NOME_QUARTO_OMNIBEES": { "campo_pax": valor } } }
    # Quartos simples (por quarto) → campo "2 Pax"
    # Presidencial Dbl → "2 Pax" | Tpl → "3 Pax" | Qpl → "4 Pax"
    # Crianças → sempre 0
    BARS = {
        "RACK": {
            "Duplex Superior":                             {"1 Pax": 6625, "2 Pax": 6625, "3 Pax": 6625, "4 Pax": 6625, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 5300, "3 Pax": 5963, "4 Pax": 6625, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 4240, "3 Pax": 4770, "4 Pax": 5300, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 4240, "2 Pax": 4240},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 6625, "2 Pax": 6625},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 6625, "2 Pax": 6625, "3 Pax": 6625, "4 Pax": 6625, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 5300, "2 Pax": 5300, "3 Pax": 5300, "4 Pax": 5300, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 5300, "2 Pax": 5300, "3 Pax": 5300, "4 Pax": 5300, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 6625, "2 Pax": 6625, "3 Pax": 6625, "4 Pax": 6625, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 5300, "2 Pax": 5300, "3 Pax": 5300},
            "Suíte com Varanda":                           {"1 Pax": 4240, "2 Pax": 4240},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 5300, "2 Pax": 5300, "3 Pax": 5300, "4 Pax": 5300, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 4240, "2 Pax": 4240},
            "Duplo com lareira":                           {"1 Pax": 3710, "2 Pax": 3710},
        },
        "BAR 1": {
            "Duplex Superior":                             {"1 Pax": 5963, "2 Pax": 5963, "3 Pax": 5963, "4 Pax": 5963, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 4770, "3 Pax": 5366, "4 Pax": 5963, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 3816, "3 Pax": 4293, "4 Pax": 4770, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 3816, "2 Pax": 3816},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 5963, "2 Pax": 5963},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 5963, "2 Pax": 5963, "3 Pax": 5963, "4 Pax": 5963, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 4770, "2 Pax": 4770, "3 Pax": 4770, "4 Pax": 4770, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 4770, "2 Pax": 4770, "3 Pax": 4770, "4 Pax": 4770, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 5963, "2 Pax": 5963, "3 Pax": 5963, "4 Pax": 5963, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 4770, "2 Pax": 4770, "3 Pax": 4770},
            "Suíte com Varanda":                           {"1 Pax": 3816, "2 Pax": 3816},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 4770, "2 Pax": 4770, "3 Pax": 4770, "4 Pax": 4770, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 3816, "2 Pax": 3816},
            "Duplo com lareira":                           {"1 Pax": 3339, "2 Pax": 3339},
        },
        "BAR 2": {
            "Duplex Superior":                             {"1 Pax": 5366, "2 Pax": 5366, "3 Pax": 5366, "4 Pax": 5366, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 4293, "3 Pax": 4830, "4 Pax": 5366, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 3434, "3 Pax": 3864, "4 Pax": 4293, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 3434, "2 Pax": 3434},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 5366, "2 Pax": 5366},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 5366, "2 Pax": 5366, "3 Pax": 5366, "4 Pax": 5366, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 4293, "2 Pax": 4293, "3 Pax": 4293, "4 Pax": 4293, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 4293, "2 Pax": 4293, "3 Pax": 4293, "4 Pax": 4293, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 5366, "2 Pax": 5366, "3 Pax": 5366, "4 Pax": 5366, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 4293, "2 Pax": 4293, "3 Pax": 4293},
            "Suíte com Varanda":                           {"1 Pax": 3434, "2 Pax": 3434},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 4293, "2 Pax": 4293, "3 Pax": 4293, "4 Pax": 4293, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 3434, "2 Pax": 3434},
            "Duplo com lareira":                           {"1 Pax": 3005, "2 Pax": 3005},
        },
        "BAR 3": {
            "Duplex Superior":                             {"1 Pax": 4830, "2 Pax": 4830, "3 Pax": 4830, "4 Pax": 4830, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 3864, "3 Pax": 4347, "4 Pax": 4830, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 3091, "3 Pax": 3477, "4 Pax": 3864, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 3091, "2 Pax": 3091},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 4830, "2 Pax": 4830},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 4830, "2 Pax": 4830, "3 Pax": 4830, "4 Pax": 4830, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 3864, "2 Pax": 3864, "3 Pax": 3864, "4 Pax": 3864, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 3864, "2 Pax": 3864, "3 Pax": 3864, "4 Pax": 3864, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 4830, "2 Pax": 4830, "3 Pax": 4830, "4 Pax": 4830, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 3864, "2 Pax": 3864, "3 Pax": 3864},
            "Suíte com Varanda":                           {"1 Pax": 3091, "2 Pax": 3091},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 3864, "2 Pax": 3864, "3 Pax": 3864, "4 Pax": 3864, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 3091, "2 Pax": 3091},
            "Duplo com lareira":                           {"1 Pax": 2705, "2 Pax": 2705},
        },
        "BAR 4": {
            "Duplex Superior":                             {"1 Pax": 4347, "2 Pax": 4347, "3 Pax": 4347, "4 Pax": 4347, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 3477, "3 Pax": 3912, "4 Pax": 4347, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 2782, "3 Pax": 3130, "4 Pax": 3477, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 2782, "2 Pax": 2782},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 4347, "2 Pax": 4347},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 4347, "2 Pax": 4347, "3 Pax": 4347, "4 Pax": 4347, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 3477, "2 Pax": 3477, "3 Pax": 3477, "4 Pax": 3477, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 3477, "2 Pax": 3477, "3 Pax": 3477, "4 Pax": 3477, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 4347, "2 Pax": 4347, "3 Pax": 4347, "4 Pax": 4347, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 3477, "2 Pax": 3477, "3 Pax": 3477},
            "Suíte com Varanda":                           {"1 Pax": 2782, "2 Pax": 2782},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 3477, "2 Pax": 3477, "3 Pax": 3477, "4 Pax": 3477, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 2782, "2 Pax": 2782},
            "Duplo com lareira":                           {"1 Pax": 2434, "2 Pax": 2434},
        },
        "BAR 5": {
            "Duplex Superior":                             {"1 Pax": 3912, "2 Pax": 3912, "3 Pax": 3912, "4 Pax": 3912, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 3130, "3 Pax": 3521, "4 Pax": 3912, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 2504, "3 Pax": 2817, "4 Pax": 3130, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 2504, "2 Pax": 2504},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 3912, "2 Pax": 3912},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 3912, "2 Pax": 3912, "3 Pax": 3912, "4 Pax": 3912, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 3130, "2 Pax": 3130, "3 Pax": 3130, "4 Pax": 3130, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 3130, "2 Pax": 3130, "3 Pax": 3130, "4 Pax": 3130, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 3912, "2 Pax": 3912, "3 Pax": 3912, "4 Pax": 3912, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 3130, "2 Pax": 3130, "3 Pax": 3130},
            "Suíte com Varanda":                           {"1 Pax": 2504, "2 Pax": 2504},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 3130, "2 Pax": 3130, "3 Pax": 3130, "4 Pax": 3130, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 2504, "2 Pax": 2504},
            "Duplo com lareira":                           {"1 Pax": 2191, "2 Pax": 2191},
        },
        "BAR 6": {
            "Duplex Superior":                             {"1 Pax": 3521, "2 Pax": 3521, "3 Pax": 3521, "4 Pax": 3521, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 2817, "3 Pax": 3169, "4 Pax": 3521, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 2253, "3 Pax": 2535, "4 Pax": 2817, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 2253, "2 Pax": 2253},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 3521, "2 Pax": 3521},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 3521, "2 Pax": 3521, "3 Pax": 3521, "4 Pax": 3521, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 2817, "2 Pax": 2817, "3 Pax": 2817, "4 Pax": 2817, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 2817, "2 Pax": 2817, "3 Pax": 2817, "4 Pax": 2817, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 3521, "2 Pax": 3521, "3 Pax": 3521, "4 Pax": 3521, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 2817, "2 Pax": 2817, "3 Pax": 2817},
            "Suíte com Varanda":                           {"1 Pax": 2253, "2 Pax": 2253},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 2817, "2 Pax": 2817, "3 Pax": 2817, "4 Pax": 2817, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 2253, "2 Pax": 2253},
            "Duplo com lareira":                           {"1 Pax": 1972, "2 Pax": 1972},
        },
        "BAR 7": {
            "Duplex Superior":                             {"1 Pax": 3169, "2 Pax": 3169, "3 Pax": 3169, "4 Pax": 3169, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 2535, "3 Pax": 2852, "4 Pax": 3169, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 2028, "3 Pax": 2281, "4 Pax": 2535, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 2028, "2 Pax": 2028},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 3169, "2 Pax": 3169},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 3169, "2 Pax": 3169, "3 Pax": 3169, "4 Pax": 3169, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 2535, "2 Pax": 2535, "3 Pax": 2535, "4 Pax": 2535, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 2535, "2 Pax": 2535, "3 Pax": 2535, "4 Pax": 2535, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 3169, "2 Pax": 3169, "3 Pax": 3169, "4 Pax": 3169, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 2535, "2 Pax": 2535, "3 Pax": 2535},
            "Suíte com Varanda":                           {"1 Pax": 2028, "2 Pax": 2028},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 2535, "2 Pax": 2535, "3 Pax": 2535, "4 Pax": 2535, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 2028, "2 Pax": 2028},
            "Duplo com lareira":                           {"1 Pax": 1774, "2 Pax": 1774},
        },
        "BAR 8": {
            "Duplex Superior":                             {"1 Pax": 2852, "2 Pax": 2852, "3 Pax": 2852, "4 Pax": 2852, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 2281, "3 Pax": 2567, "4 Pax": 2852, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 1825, "3 Pax": 2053, "4 Pax": 2281, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 1825, "2 Pax": 1825},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 2852, "2 Pax": 2852},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 2852, "2 Pax": 2852, "3 Pax": 2852, "4 Pax": 2852, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 2281, "2 Pax": 2281, "3 Pax": 2281, "4 Pax": 2281, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 2281, "2 Pax": 2281, "3 Pax": 2281, "4 Pax": 2281, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 2852, "2 Pax": 2852, "3 Pax": 2852, "4 Pax": 2852, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 2281, "2 Pax": 2281, "3 Pax": 2281},
            "Suíte com Varanda":                           {"1 Pax": 1825, "2 Pax": 1825},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 2281, "2 Pax": 2281, "3 Pax": 2281, "4 Pax": 2281, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 1825, "2 Pax": 1825},
            "Duplo com lareira":                           {"1 Pax": 1597, "2 Pax": 1597},
        },
        "BAR 9": {
            "Duplex Superior":                             {"1 Pax": 2567, "2 Pax": 2567, "3 Pax": 2567, "4 Pax": 2567, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 2053, "3 Pax": 2310, "4 Pax": 2567, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 1643, "3 Pax": 1848, "4 Pax": 2053, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 1643, "2 Pax": 1643},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 2567, "2 Pax": 2567},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 2567, "2 Pax": 2567, "3 Pax": 2567, "4 Pax": 2567, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 2053, "2 Pax": 2053, "3 Pax": 2053, "4 Pax": 2053, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 2053, "2 Pax": 2053, "3 Pax": 2053, "4 Pax": 2053, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 2567, "2 Pax": 2567, "3 Pax": 2567, "4 Pax": 2567, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 2053, "2 Pax": 2053, "3 Pax": 2053},
            "Suíte com Varanda":                           {"1 Pax": 1643, "2 Pax": 1643},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 2053, "2 Pax": 2053, "3 Pax": 2053, "4 Pax": 2053, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 1643, "2 Pax": 1643},
            "Duplo com lareira":                           {"1 Pax": 1437, "2 Pax": 1437},
        },
        "BAR 10": {
            "Duplex Superior":                             {"1 Pax": 2310, "2 Pax": 2310, "3 Pax": 2310, "4 Pax": 2310, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 1848, "3 Pax": 2079, "4 Pax": 2310, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 1478, "3 Pax": 1663, "4 Pax": 1848, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 1478, "2 Pax": 1478},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 2310, "2 Pax": 2310},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 2310, "2 Pax": 2310, "3 Pax": 2310, "4 Pax": 2310, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 1848, "2 Pax": 1848, "3 Pax": 1848, "4 Pax": 1848, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 1848, "2 Pax": 1848, "3 Pax": 1848, "4 Pax": 1848, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 2310, "2 Pax": 2310, "3 Pax": 2310, "4 Pax": 2310, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 1848, "2 Pax": 1848, "3 Pax": 1848},
            "Suíte com Varanda":                           {"1 Pax": 1478, "2 Pax": 1478},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 1848, "2 Pax": 1848, "3 Pax": 1848, "4 Pax": 1848, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 1478, "2 Pax": 1478},
            "Duplo com lareira":                           {"1 Pax": 1294, "2 Pax": 1294},
        },
        "BAR 11": {
            "Duplex Superior":                             {"1 Pax": 2079, "2 Pax": 2079, "3 Pax": 2079, "4 Pax": 2079, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 1663, "3 Pax": 1871, "4 Pax": 2079, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 1331, "3 Pax": 1497, "4 Pax": 1663, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 1331, "2 Pax": 1331},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 2079, "2 Pax": 2079},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 2079, "2 Pax": 2079, "3 Pax": 2079, "4 Pax": 2079, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 1663, "2 Pax": 1663, "3 Pax": 1663, "4 Pax": 1663, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 1663, "2 Pax": 1663, "3 Pax": 1663, "4 Pax": 1663, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 2079, "2 Pax": 2079, "3 Pax": 2079, "4 Pax": 2079, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 1663, "2 Pax": 1663, "3 Pax": 1663},
            "Suíte com Varanda":                           {"1 Pax": 1331, "2 Pax": 1331},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 1663, "2 Pax": 1663, "3 Pax": 1663, "4 Pax": 1663, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 1331, "2 Pax": 1331},
            "Duplo com lareira":                           {"1 Pax": 1164, "2 Pax": 1164},
        },
        "BAR 12": {
            "Duplex Superior":                             {"1 Pax": 1871, "2 Pax": 1871, "3 Pax": 1871, "4 Pax": 1871, "1 Criança": 0, "2 Crianças": 0},
            "Suite presidencial com Varanda":              {"2 Pax": 1497, "3 Pax": 1684, "4 Pax": 1871, "1 Criança": 0, "2 Crianças": 0},
            "Suíte Presidencial":                          {"2 Pax": 1198, "3 Pax": 1347, "4 Pax": 1497, "1 Criança": 0, "2 Crianças": 0},
            "Quarto duplo com varanda":                    {"1 Pax": 1198, "2 Pax": 1198},
            "Quarto duplo luxo com sauna":                 {"1 Pax": 1871, "2 Pax": 1871},
            "Chalé Exclusivo com Cozinha e Sala de Estar": {"1 Pax": 1871, "2 Pax": 1871, "3 Pax": 1871, "4 Pax": 1871, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Família Exclusivo com Cozinha":         {"1 Pax": 1497, "2 Pax": 1497, "3 Pax": 1497, "4 Pax": 1497, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Chalé Duplex com Varanda":                    {"1 Pax": 1497, "2 Pax": 1497, "3 Pax": 1497, "4 Pax": 1497, "1 Criança": 0, "2 Crianças": 0},
            "Duplex com piscina e lareira":                {"1 Pax": 1871, "2 Pax": 1871, "3 Pax": 1871, "4 Pax": 1871, "1 Criança": 0, "2 Crianças": 0},
            "Chalé Duplex com Ofurô":                      {"1 Pax": 1497, "2 Pax": 1497, "3 Pax": 1497},
            "Suíte com Varanda":                           {"1 Pax": 1198, "2 Pax": 1198},
            "Chalé Deluxe com Banheira e Lareira":         {"1 Pax": 1497, "2 Pax": 1497, "3 Pax": 1497, "4 Pax": 1497, "1 Criança": 0, "2 Crianças": 0, "3 Crianças": 0},
            "Quarto Duplo com banheira":                   {"1 Pax": 1198, "2 Pax": 1198},
            "Duplo com lareira":                           {"1 Pax": 1048, "2 Pax": 1048},
        },
    }


    hotel_sel = st.sidebar.selectbox("🏨 Hotel:", list(HOTEIS_OMNIBEES.keys()), key="omni_hotel")
    config = HOTEIS_OMNIBEES[hotel_sel]

    # ── Sub-abas ───────────────────────────────────────────────────────────────
    omni_tab1, omni_tab2 = st.tabs(["🔒 Fechar / Abrir Disponibilidade", "💰 Atualizar Tarifas"])

    # ── Funções auxiliares compartilhadas ─────────────────────────────────────

    # =========================================================
    # SUB-ABA 1: FECHAR / ABRIR DISPONIBILIDADE
    # =========================================================
    with omni_tab1:
        hoje = date.today()
        col1, col2 = st.columns(2)
        with col1:
            data_ini_omni = st.date_input("📅 Data Início", hoje, key="omni_ini", format="DD/MM/YYYY")
        with col2:
            data_fim_omni = st.date_input("📅 Data Fim", hoje, key="omni_fim", format="DD/MM/YYYY")

        st.markdown("---")
        st.markdown("### 🏷️ Tarifários")
        sel_todos_tar = st.checkbox("✅ Selecionar Todos os Tarifários", value=True, key="omni_todos_tar")
        tarifarios_sel = []
        if sel_todos_tar:
            tarifarios_sel = config["tarifarios"]
            for t in config["tarifarios"]:
                st.checkbox(t, value=True, disabled=True, key=f"tar_{t}")
        else:
            cols_tar = st.columns(2)
            for i, t in enumerate(config["tarifarios"]):
                with cols_tar[i % 2]:
                    if st.checkbox(t, value=False, key=f"tar_{t}"):
                        tarifarios_sel.append(t)

        st.markdown("---")
        st.markdown("### 🛏️ Tipos de Quartos")
        sel_todos_qrt = st.checkbox("✅ Selecionar Todos os Quartos", value=True, key="omni_todos_qrt")
        quartos_sel = []
        if sel_todos_qrt:
            quartos_sel = config["quartos"]
            for q in config["quartos"]:
                st.checkbox(q, value=True, disabled=True, key=f"qrt_{q}")
        else:
            cols_qrt = st.columns(2)
            for i, q in enumerate(config["quartos"]):
                with cols_qrt[i % 2]:
                    if st.checkbox(q, value=False, key=f"qrt_{q}"):
                        quartos_sel.append(q)

        st.markdown("---")
        codigo_2fa = ""  # Capturado automaticamente pelo robô
        st.markdown("---")
        col_btn1, col_btn2 = st.columns(2)

        def executar_omnibees(acao):
            if not tarifarios_sel:
                st.error("❌ Selecione pelo menos um tarifário!")
                return
            if not quartos_sel:
                st.error("❌ Selecione pelo menos um tipo de quarto!")
                return
            status_pre = st.empty()
            codigo_2fa = _capturar_2fa_authenticator(status_pre)
            if not codigo_2fa or len(codigo_2fa) != 6:
                st.error("❌ Não foi possível capturar o código 2FA automaticamente!")
                return
            driver = criar_driver()
            wait = WebDriverWait(driver, 15)
            status = st.empty()
            try:
                _login_omnibees(driver, wait, config, codigo_2fa, status)
                _selecionar_datas(driver, wait, data_ini_omni, data_fim_omni, status)
                _selecionar_tarifarios(driver, wait, sel_todos_tar, tarifarios_sel, status)
                _selecionar_quartos(driver, wait, sel_todos_qrt, quartos_sel, status)
                status.markdown("🔒 Abrindo aba Fechar/Abrir Vendas...")
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "#ob-ui-update-rates-packages-tab > li:nth-child(3)"))).click()
                time.sleep(1)
                status.markdown(f"{'🔒' if acao == 'fechar' else '🟢'} Selecionando {'Fechar' if acao == 'fechar' else 'Abrir'} Vendas...")
                if acao == "fechar":
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                        "#rd-updateRates-input-close-open-sales-2"))).click()
                else:
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                        "#rd-updateRates-input-close-open-sales-3"))).click()
                time.sleep(1)
                status.markdown("💾 Salvando...")
                wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//button[contains(text(), 'Salvar')]"))).click()
                time.sleep(2)
                driver.quit()
                emoji = "🔒" if acao == "fechar" else "🟢"
                acao_txt = "fechada" if acao == "fechar" else "aberta"
                status.success(f"{emoji} Disponibilidade {acao_txt} com sucesso! "
                               f"{data_ini_omni.strftime('%d/%m/%Y')} a {data_fim_omni.strftime('%d/%m/%Y')}")
                st.session_state.omnibees_status = (f"{emoji} {acao_txt.capitalize()} em "
                    f"{data_ini_omni.strftime('%d/%m/%Y')} - {data_fim_omni.strftime('%d/%m/%Y')}")
                historico_salvar(
                    funcao=f"Omnibees — {'Fechar' if acao == 'fechar' else 'Abrir'} Disponibilidade",
                    status="✅ Sucesso",
                    detalhes={
                        "hotel": hotel_sel,
                        "quartos": ", ".join(quartos_sel) if not sel_todos_qrt else "Todos",
                        "tarifarios": ", ".join(tarifarios_sel) if not sel_todos_tar else "Todos",
                        "data_ini": data_ini_omni.strftime('%d/%m/%Y'),
                        "data_fim": data_fim_omni.strftime('%d/%m/%Y'),
                        "mensagem": ""
                    }
                )
            except Exception as e:
                try: driver.quit()
                except: pass
                status.error(f"❌ Erro: {e}")
                historico_salvar(
                    funcao=f"Omnibees — {'Fechar' if acao == 'fechar' else 'Abrir'} Disponibilidade",
                    status=f"❌ Erro: {str(e)[:100]}",
                    detalhes={"hotel": hotel_sel}
                )

        with col_btn1:
            if st.button("🔒 FECHAR DISPONIBILIDADE", key="omni_fechar", use_container_width=True):
                executar_omnibees("fechar")
        with col_btn2:
            if st.button("🟢 ABRIR DISPONIBILIDADE", key="omni_abrir", use_container_width=True):
                executar_omnibees("abrir")

        if st.session_state.omnibees_status:
            st.info(f"Última ação: {st.session_state.omnibees_status}")

    # =========================================================
    # SUB-ABA 2: ATUALIZAR TARIFAS (BARs)
    # =========================================================
    with omni_tab2:
        st.markdown("### 💰 Atualizar Tarifas por BAR")

        # Toggle salvar
        salvar_ativado = st.toggle(
            "💾 Ativar Salvar automaticamente",
            value=False,
            key="toggle_salvar_bar",
            help="Quando desativado, o robô preenche os valores mas NÃO salva. Ative apenas quando tiver certeza."
        )
        if salvar_ativado:
            st.warning("⚠️ Salvar ATIVADO — o robô irá salvar automaticamente após preencher!")
        else:
            st.info("🔒 Salvar DESATIVADO — modo de teste. O robô preenche mas não salva.")

        hoje2 = date.today()
        col1t, col2t = st.columns(2)
        with col1t:
            data_ini_tar = st.date_input("📅 Data Início", hoje2, key="tar_ini", format="DD/MM/YYYY")
        with col2t:
            data_fim_tar = st.date_input("📅 Data Fim", hoje2, key="tar_fim", format="DD/MM/YYYY")

        st.markdown("---")

        # Seleção de Tarifários
        st.markdown("### 🏷️ Tarifários")
        sel_todos_tar2 = st.checkbox("✅ Selecionar Todos os Tarifários", value=True, key="omni_todos_tar2")
        tarifarios_sel2 = []
        if sel_todos_tar2:
            tarifarios_sel2 = config["tarifarios"]
            for t in config["tarifarios"]:
                st.checkbox(t, value=True, disabled=True, key=f"tar2_{t}")
        else:
            cols_tar2 = st.columns(2)
            for i, t in enumerate(config["tarifarios"]):
                with cols_tar2[i % 2]:
                    if st.checkbox(t, value=False, key=f"tar2_{t}"):
                        tarifarios_sel2.append(t)

        st.markdown("---")

        # Seleção da BAR
        st.markdown("### 🎯 Selecione a BAR")
        bar_selecionada = st.selectbox("BAR:", list(BARS.keys()), key="bar_sel")

        # Preview dos valores da BAR selecionada
        with st.expander("👁️ Ver tarifas da BAR selecionada"):
            dados_preview = []
            for quarto, campos in BARS[bar_selecionada].items():
                linha = {"Quarto": quarto}
                linha.update(campos)
                dados_preview.append(linha)
            df_preview = pd.DataFrame(dados_preview).fillna("").astype(str).replace("0", "-")
            st.dataframe(df_preview, use_container_width=True, hide_index=True)

        st.markdown("---")
        codigo_2fa_tar = ""  # Capturado automaticamente pelo robô
        st.markdown("---")

        def executar_atualizar_tarifas():
            status_pre2 = st.empty()
            codigo_2fa_tar = _capturar_2fa_com_retry(status_pre2)
            if not codigo_2fa_tar or len(codigo_2fa_tar) != 6:
                st.error("❌ Não foi possível capturar o código 2FA automaticamente!")
                _log_aresta("ERRO: 2FA não capturado na aba Omnibees tarifas", "ERRO")
                return

            tarifas_bar = BARS[bar_selecionada]

            driver = criar_driver()
            wait = WebDriverWait(driver, 15)
            status = st.empty()

            try:
                _login_omnibees(driver, wait, config, codigo_2fa_tar, status)
                _selecionar_datas(driver, wait, data_ini_tar, data_fim_tar, status)
                _selecionar_tarifarios(driver, wait, sel_todos_tar2, tarifarios_sel2, status)
                _selecionar_quartos(driver, wait, True, config["quartos"], status)

                # Garante que está na aba "Preços e Allotment"
                status.markdown("💰 Abrindo aba Preços e Allotment...")
                try:
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                        "#ob-ui-update-rates-packages-tab > li:nth-child(1)"))).click()
                    time.sleep(1.5)
                except:
                    pass

                # Mapa de label → data-cy do input (extraído do HTML do Omnibees)
                PAX_DATA_CY = {
                    "1 Pax":      "reservation_limit_multiple-price_decimal_1_1-input",
                    "2 Pax":      "reservation_limit_multiple-price_decimal_2_1-input",
                    "3 Pax":      "reservation_limit_multiple-price_decimal_3_1-input",
                    "4 Pax":      "reservation_limit_multiple-price_decimal_4_1-input",
                    "1 Criança":  "reservation_limit_multiple-price_decimal_1_2-input",
                    "2 Crianças": "reservation_limit_multiple-price_decimal_2_2-input",
                    "3 Crianças": "reservation_limit_multiple-price_decimal_3_2-input",
                }

                # Preenche campo por campo, quarto por quarto
                from selenium.webdriver.common.keys import Keys

                for nome_quarto, campos_pax in tarifas_bar.items():
                    status.markdown(f"🛏️ Preenchendo **{nome_quarto}** — {bar_selecionada}...")

                    # Localiza o bloco correto — match exato case-insensitive
                    bloco_quarto = driver.execute_script("""
                        var nomeAlvo = arguments[0].toLowerCase().trim();
                        var headers = document.querySelectorAll('div.header-rate-room');
                        // Primeiro tenta match exato
                        for (var h of headers) {
                            if (h.innerText.toLowerCase().trim() === nomeAlvo) {
                                var p = h.parentElement;
                                for (var i = 0; i < 10; i++) {
                                    if (!p) break;
                                    if (p.querySelector('input[data-cy*="price_decimal"]')) return p;
                                    p = p.parentElement;
                                }
                            }
                        }
                        // Fallback: match parcial
                        for (var h of headers) {
                            if (h.innerText.toLowerCase().trim().indexOf(nomeAlvo) !== -1) {
                                var p = h.parentElement;
                                for (var i = 0; i < 10; i++) {
                                    if (!p) break;
                                    if (p.querySelector('input[data-cy*="price_decimal"]')) return p;
                                    p = p.parentElement;
                                }
                            }
                        }
                        return null;
                    """, nome_quarto)

                    if not bloco_quarto:
                        status.markdown(f"⚠️ **{nome_quarto}** não encontrado, pulando...")
                        continue

                    # Rola até o bloco ficar visível
                    driver.execute_script("""
                        arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});
                        window.scrollBy(0, -80);
                    """, bloco_quarto)
                    time.sleep(0.6)

                    for label_pax, valor in campos_pax.items():
                        cy = PAX_DATA_CY.get(label_pax)
                        if not cy:
                            continue
                        valor_str = str(int(valor))
                        try:
                            inp = bloco_quarto.find_element(By.CSS_SELECTOR,
                                f'input[data-cy="{cy}"]')
                            driver.execute_script(
                                "arguments[0].scrollIntoView({block:'center', behavior:'instant'});", inp)
                            time.sleep(0.2)
                            inp.click()
                            time.sleep(0.2)
                            inp.send_keys(Keys.CONTROL + "a")
                            time.sleep(0.1)
                            inp.send_keys(Keys.DELETE)
                            time.sleep(0.1)
                            inp.send_keys(valor_str)
                            time.sleep(0.2)
                            inp.send_keys(Keys.TAB)
                            time.sleep(0.3)
                        except:
                            pass
                    time.sleep(0.4)

                if salvar_ativado:
                    # Clica no botão Salvar
                    status.markdown("💾 Salvando tarifas...")
                    time.sleep(2)
                    try:
                        btn_salvar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                            'button[data-cy="save_btn-input"]')))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_salvar)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", btn_salvar)
                        time.sleep(2)
                        # Confirma popup "Sim, Continuar"
                        status.markdown("⏳ Aguardando popup de confirmação...")
                        time.sleep(1.5)
                        try:
                            btn_sim = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                'span[data-cy="button-label_-_Sim, Continuar"]'
                            )))
                            driver.execute_script("arguments[0].click();", btn_sim)
                            status.markdown("✅ Popup confirmado!")
                            time.sleep(2)
                        except:
                            pass
                        status.success(
                            f"✅ {bar_selecionada} salva com sucesso! "
                            f"Período: {data_ini_tar.strftime('%d/%m/%Y')} a {data_fim_tar.strftime('%d/%m/%Y')}")
                    except Exception as e_save:
                        status.error(f"❌ Erro ao salvar: {e_save}")
                else:
                    status.success(
                        f"✅ {bar_selecionada} preenchida! Salvar está DESATIVADO. "
                        f"Confira os valores no Omnibees antes de salvar.")
                driver.quit()
                st.session_state.omnibees_status = (
                    f"💰 {bar_selecionada} salva: "
                    f"{data_ini_tar.strftime('%d/%m/%Y')} - {data_fim_tar.strftime('%d/%m/%Y')}")
                historico_salvar(
                    funcao=f"Omnibees — Atualizar Tarifas",
                    status="✅ Salva com sucesso",
                    detalhes={
                        "hotel": hotel_sel,
                        "bar": bar_selecionada,
                        "data_ini": data_ini_tar.strftime('%d/%m/%Y'),
                        "data_fim": data_fim_tar.strftime('%d/%m/%Y'),
                        "mensagem": ""
                    }
                )

            except Exception as e:
                driver.quit()
                status.error(f"❌ Erro: {e}")

        if st.button("💰 APLICAR TARIFAS NO OMNIBEES", key="omni_aplicar_bar", use_container_width=True):
            executar_atualizar_tarifas()

        if st.session_state.omnibees_status:
            st.info(f"Última ação: {st.session_state.omnibees_status}")

# ABA 8: MONITOR DE PICK-UP
# ==============================================================================
elif aba_selecionada == "📊 Monitor de Pick-up":
    st.title("📊 MONITOR DE PICK-UP")

    PICKUP_FILE = "Pick-up_sniper.xlsx"

    def _pickup_ler_dados(sheet_name):
        if not os.path.exists(PICKUP_FILE):
            return None, {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(PICKUP_FILE, read_only=True, data_only=True)
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            meses = {0: {}, 1: {}, 2: {}}
            for mi, off in enumerate([0, 8, 16]):
                for row in rows[2:34]:
                    c = off
                    if row[c] and hasattr(row[c], 'day'):
                        dia = row[c].day
                        hr  = row[c+3] if isinstance(row[c+3], (int, float)) else 0
                        ho  = row[c+4] if isinstance(row[c+4], (int, float)) else 0
                        or_ = row[c+1] if isinstance(row[c+1], (int, float)) else 0
                        oo  = row[c+2] if isinstance(row[c+2], (int, float)) else 0
                        pu_rec_raw = row[c+5] if isinstance(row[c+5], (int, float)) else (hr - or_)
                        pu_occ_raw = row[c+6] if isinstance(row[c+6], (int, float)) else (ho - oo)
                        if pu_rec_raw == 0 and hr != or_:
                            pu_rec_raw = hr - or_
                        if pu_occ_raw == 0 and ho != oo:
                            pu_occ_raw = ho - oo
                        meses[mi][dia] = {
                            'rec':    float(hr),
                            'occ':    float(ho),
                            'pu_rec': float(pu_rec_raw),
                            'pu_occ': float(pu_occ_raw),
                        }
            # Diária média do Normandie: Abril=F35, Maio=N36, Junho=V36
            def _safe(r, c):
                try:
                    v = rows[r][c]
                    return float(v) if isinstance(v, (int, float)) else 0.0
                except: return 0.0
            dm_normandie = {
                0: _safe(34, 5),   # F35
                1: _safe(35, 13),  # N36
                2: _safe(35, 21),  # V36
            }
            return meses, dm_normandie
        except:
            return None, {}

    def _pickup_ler_metas():
        if not os.path.exists(PICKUP_FILE):
            return {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(PICKUP_FILE, read_only=True, data_only=True)
            ws = wb['Alto da Boa Vista Metas']
            rows = list(ws.iter_rows(values_only=True))
            def v(r, c):
                try:
                    val = rows[r][c]
                    if val in ('#DIV/0!', '#REF!', '#VALUE!', '#N/A', '#NAME?'): return 0.0
                    return float(val) if val is not None else 0.0
                except: return 0.0
            return {
                # Abril: L4=dias_faltam, L7=receita, L8=occ, L9=dm (índice 0-based: 3,6,7,8)
                0: {
                    'rec_acum': v(6,2),  'rec_meta': v(6,3),  'falta_dia': v(6,5),  'dias_faltam': v(3,3),
                    'occ_acum': v(7,2),  'occ_meta': v(7,3),
                    'dm_acum':  v(8,2),  'dm_meta':  v(8,3),
                },
                # Maio: L11=dias_faltam, L14=receita (índice: 10,13,14,15)
                1: {
                    'rec_acum': v(13,2), 'rec_meta': v(13,3), 'falta_dia': v(13,5), 'dias_faltam': v(10,3),
                    'occ_acum': v(14,2), 'occ_meta': v(14,3),
                    'dm_acum':  v(15,2), 'dm_meta':  v(15,3),
                },
                # Junho: L18=dias_faltam, L21=receita (índice: 17,20,21,22)
                2: {
                    'rec_acum': v(20,2), 'rec_meta': v(20,3), 'falta_dia': v(20,5), 'dias_faltam': v(17,3),
                    'occ_acum': v(21,2), 'occ_meta': v(21,3),
                    'dm_acum':  v(22,2), 'dm_meta':  v(22,3),
                },
            }
        except: return {}

    def _pickup_render(meses, metas, cfg_meses, mes_idx):
        cfg      = cfg_meses[mes_idx]
        dias     = meses.get(mes_idx, {})
        meta     = metas.get(mes_idx, {})
        mes_nome = cfg['nome']
        total_rec = sum(d['rec']    for d in dias.values())
        total_occ = sum(d['occ']    for d in dias.values()) / len(dias) if dias else 0
        pu_rec    = sum(d['pu_rec'] for d in dias.values())
        r_acum = meta.get('rec_acum', 0); r_meta = meta.get('rec_meta', 0)
        o_acum_raw = meta.get('occ_acum', 0)
        o_acum = o_acum_raw * 100 if o_acum_raw <= 1 else o_acum_raw
        o_meta_raw = meta.get('occ_meta', 0)
        o_meta = o_meta_raw * 100 if o_meta_raw <= 1 else o_meta_raw
        dm_acum    = meta.get('dm_acum', 0)
        dm_meta    = meta.get('dm_meta', 0)
        dias_faltam = int(meta.get('dias_faltam', 0))
        falta_dia   = meta.get('falta_dia', 0)
        pct_r  = round((r_acum / r_meta * 100) if r_meta else 0, 1)
        pct_o  = round((o_acum / o_meta * 100) if o_meta else 0, 1)
        pct_dm = round((dm_acum / dm_meta * 100) if dm_meta else 0, 1)
        pu_cor = "#4ade80" if pu_rec > 0 else "#f87171" if pu_rec < 0 else "#e5e7eb"

        def brl(val):
            neg = val < 0
            s = f"{abs(val):,.2f}".replace(",","X").replace(".",",").replace("X",".")
            return f"-R$ {s}" if neg else f"R$ {s}"

        html = """<style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
        .pu-wrap{background:linear-gradient(135deg,#0a0a0a 0%,#1a1a2e 50%,#16213e 100%) !important;border-radius:16px;padding:20px;font-family:'Inter',sans-serif;}
        .pu-header{display:flex;align-items:stretch;gap:16px;margin-bottom:20px;}
        .pu-title-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:14px;padding:20px 28px;min-width:160px;display:flex;flex-direction:column;justify-content:center;}
        .pu-title-txt{font-size:36px;font-weight:900;color:#fff;line-height:1.1;letter-spacing:-1px;}
        .pu-mes-txt{font-size:20px;font-weight:800;color:#e63946;margin-top:6px;text-transform:uppercase;letter-spacing:1px;}
        .pu-total-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:14px;padding:16px 22px;flex:1;display:flex;flex-direction:column;justify-content:center;}
        .pu-total-lbl{font-size:11px;color:rgba(255,255,255,0.5);text-transform:uppercase;letter-spacing:.1em;font-weight:600;}
        .pu-total-val{font-size:22px;font-weight:800;color:#fff;margin-top:4px;}
        .pu-total-val.pos{color:#4ade80;} .pu-total-val.neg{color:#f87171;}
        .pu-metas{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px;}
        .pu-meta-card{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;}
        .pu-meta-lbl{font-size:10px;color:rgba(255,255,255,0.45);text-transform:uppercase;letter-spacing:.08em;font-weight:600;}
        .pu-meta-val{font-size:18px;font-weight:800;color:#fff;margin-top:4px;}
        .pu-meta-sub{font-size:11px;color:rgba(255,255,255,0.4);margin-top:2px;}
        .pu-pct{font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;display:inline-block;margin-top:4px;}
        .pu-bar-bg{height:3px;background:rgba(255,255,255,0.1);border-radius:2px;margin-top:8px;}
        .pu-bar-fg{height:3px;border-radius:2px;}
        .pu-alertas{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px;}
        .pu-alerta{border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:12px;}
        .pu-alerta-ico{font-size:30px;line-height:1;}
        .pu-alerta-lbl{font-size:10px;text-transform:uppercase;letter-spacing:.07em;font-weight:600;opacity:.7;}
        .pu-alerta-val{font-size:22px;font-weight:900;line-height:1.1;margin-top:2px;}
        .pu-wdays{display:grid;grid-template-columns:repeat(7,1fr);gap:5px;margin-bottom:5px;}
        .pu-wd{text-align:center;font-size:10px;font-weight:700;color:rgba(255,255,255,0.35);text-transform:uppercase;padding:3px 0;letter-spacing:.05em;}
        .pu-wd.wkd{color:#e63946;}
        .pu-cal{display:grid;grid-template-columns:repeat(7,1fr);gap:5px;}
        .pu-dc{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.08);border-radius:10px;padding:10px 9px;min-height:120px;}
        .pu-dc.em{background:transparent;border-color:transparent;}
        .pu-dc.wk{border-top:3px solid #e63946;background:rgba(230,57,70,0.08);}
        .pu-dc.pp{border-top:3px solid #4ade80;background:rgba(74,222,128,0.06);}
        .pu-dc.pn{border-top:3px solid #f87171;background:rgba(248,113,113,0.06);}
        .pu-dn{font-size:15px;font-weight:900;color:#fff;margin-bottom:8px;}
        .pu-dn.wc{color:#e63946;}
        .pu-lbl{font-size:9px;color:rgba(255,255,255,0.38);text-transform:uppercase;letter-spacing:.04em;margin-top:3px;}
        .pu-val{font-size:11px;font-weight:700;color:#e5e7eb;}
        .pu-val.p{color:#4ade80;} .pu-val.n{color:#f87171;} .pu-val.z{color:rgba(255,255,255,0.25);}
        .pu-dv{height:1px;background:rgba(255,255,255,0.06);margin:4px 0;}
        </style>
        <div class="pu-wrap">"""

        # HEADER — título + pick-up total + resumo
        pu_cls_hdr = "pos" if pu_rec > 0 else "neg" if pu_rec < 0 else ""
        html += f"""
        <div class="pu-header">
          <div class="pu-title-box">
            <div class="pu-title-txt">PICK-<br>UP<br>DIÁRIO</div>
            <div class="pu-mes-txt">{mes_nome}</div>
          </div>
          <div class="pu-total-box">
            <div class="pu-total-lbl">Pick-up total receita</div>
            <div class="pu-total-val {pu_cls_hdr}">{("+" if pu_rec > 0 else "") + brl(pu_rec)}</div>
            <div style="margin-top:14px;">
              <div class="pu-total-lbl">Pick-up total ocupação</div>
              <div class="pu-total-val {pu_cls_hdr}" style="font-size:18px;">{("+" if pu_rec>0 else "")}{sum(d["pu_occ"] for d in dias.values()):.2f}%</div>
            </div>
          </div>
          <div class="pu-total-box">
            <div class="pu-total-lbl">Receita total do mês</div>
            <div class="pu-total-val">{brl(total_rec)}</div>
            <div style="margin-top:14px;">
              <div class="pu-total-lbl">Ocupação média</div>
              <div class="pu-total-val" style="font-size:18px;">{total_occ:.2f}%</div>
            </div>
          </div>
          <div class="pu-total-box">
            <div class="pu-total-lbl">Diária média</div>
            <div class="pu-total-val">{brl(dm_acum) if dm_acum and dm_acum > 0 else "R$ —"}</div>
          </div>
        </div>"""

        # ALERTAS dias faltam + falta por dia
        if dias_faltam > 0 or falta_dia > 0:
            falta_html = ""
            if falta_dia and falta_dia > 0 and dias_faltam > 0:
                falta_html = f"""<div class="pu-alerta" style="background:rgba(74,222,128,0.12);border:1px solid rgba(74,222,128,0.2);">
                  <div class="pu-alerta-ico">🎯</div>
                  <div>
                    <div class="pu-alerta-lbl" style="color:rgba(74,222,128,0.7);">Faturar/dia p/ bater meta</div>
                    <div class="pu-alerta-val" style="color:#4ade80;">{brl(falta_dia)}<span style="font-size:13px;font-weight:500;color:rgba(74,222,128,0.6);"> /dia</span></div>
                  </div>
                </div>"""
            elif r_acum >= r_meta and r_meta > 0:
                falta_html = f"""<div class="pu-alerta" style="background:rgba(74,222,128,0.12);border:1px solid rgba(74,222,128,0.2);">
                  <div class="pu-alerta-ico">🏆</div>
                  <div>
                    <div class="pu-alerta-lbl" style="color:rgba(74,222,128,0.7);">Meta atingida!</div>
                    <div class="pu-alerta-val" style="color:#4ade80;">+{brl(r_acum - r_meta)}</div>
                  </div>
                </div>"""
            if dias_faltam > 0:
                html += f"""<div class="pu-alertas">
                  <div class="pu-alerta" style="background:rgba(255,215,0,0.08);border:1px solid rgba(255,215,0,0.2);">
                    <div class="pu-alerta-ico">📅</div>
                    <div>
                      <div class="pu-alerta-lbl" style="color:rgba(255,215,0,0.7);">Dias restantes no mês</div>
                      <div class="pu-alerta-val" style="color:#FFD700;">{dias_faltam} {"dia" if dias_faltam==1 else "dias"}</div>
                    </div>
                  </div>
                  {falta_html}
                </div>"""

        # METAS
        if r_meta > 0 or o_meta > 0:
            html += f"""<div class="pu-metas">
              <div class="pu-meta-card">
                <div class="pu-meta-lbl">Receita Acumulada</div>
                <div class="pu-meta-val">{brl(r_acum)}</div>
                <div class="pu-meta-sub">Meta: {brl(r_meta)}</div>
                <span class="pu-pct" style="background:{"rgba(74,222,128,0.15)" if pct_r>=100 else "rgba(251,146,60,0.15)"};color:{"#4ade80" if pct_r>=100 else "#fb923c"};">{pct_r}%</span>
                <div class="pu-bar-bg"><div class="pu-bar-fg" style="width:{min(pct_r,100)}%;background:{"#4ade80" if pct_r>=100 else "#fb923c"};"></div></div>
              </div>
              <div class="pu-meta-card">
                <div class="pu-meta-lbl">Ocupação Média</div>
                <div class="pu-meta-val">{o_acum:.2f}%</div>
                <div class="pu-meta-sub">Meta: {o_meta:.2f}%</div>
                <span class="pu-pct" style="background:{"rgba(74,222,128,0.15)" if pct_o>=100 else "rgba(251,146,60,0.15)"};color:{"#4ade80" if pct_o>=100 else "#fb923c"};">{pct_o}%</span>
                <div class="pu-bar-bg"><div class="pu-bar-fg" style="width:{min(pct_o,100)}%;background:{"#4ade80" if pct_o>=100 else "#fb923c"};"></div></div>
              </div>
              <div class="pu-meta-card">
                <div class="pu-meta-lbl">Diária Média</div>
                <div class="pu-meta-val">{brl(dm_acum)}</div>
                <div class="pu-meta-sub">Meta: {brl(dm_meta)}</div>
                <span class="pu-pct" style="background:{"rgba(74,222,128,0.15)" if pct_dm>=100 else "rgba(251,146,60,0.15)"};color:{"#4ade80" if pct_dm>=100 else "#fb923c"};">{pct_dm}%</span>
                <div class="pu-bar-bg"><div class="pu-bar-fg" style="width:{min(pct_dm,100)}%;background:{"#4ade80" if pct_dm>=100 else "#fb923c"};"></div></div>
              </div>
            </div>"""

        # CALENDÁRIO — dias da semana
        html += '<div class="pu-wdays">'
        for d in ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"]:
            wk_cls = "wkd" if d in ["Dom","Sáb"] else ""
            html += f'<div class="pu-wd {wk_cls}">{d}</div>'
        html += '</div><div class="pu-cal">'

        for _ in range(cfg['start']):
            html += '<div class="pu-dc em"></div>'

        for day in range(1, cfg['days'] + 1):
            dow    = (cfg['start'] + day - 1) % 7
            wk     = dow == 0 or dow == 6
            info   = dias.get(day, {'rec':0,'occ':0,'pu_rec':0,'pu_occ':0})
            pu     = info['pu_rec']
            pu_occ = info['pu_occ']
            cls    = 'pu-dc' + (' wk' if wk else ' pp' if pu>0 else ' pn' if pu<0 else '')
            pc     = 'p' if pu>0 else 'n' if pu<0 else 'z'
            ps     = '+' if pu>0 else ''
            oc     = 'p' if pu_occ>0 else 'n' if pu_occ<0 else 'z'
            os_    = '+' if pu_occ>0 else ''
            html += f"""<div class="{cls}">
              <div class="pu-dn {"wc" if wk else ""}">{day}</div>
              <div class="pu-lbl">Receita</div>
              <div class="pu-val">{brl(info["rec"]) if info["rec"] else "R$ 0"}</div>
              <div class="pu-dv"></div>
              <div class="pu-lbl">Occ. do dia</div>
              <div class="pu-val">{info["occ"]:.2f}%</div>
              <div class="pu-dv"></div>
              <div class="pu-lbl">Pick-up Rec.</div>
              <div class="pu-val {pc}">{ps}{brl(pu)}</div>
              <div class="pu-dv"></div>
              <div class="pu-lbl">Pick-up Occ.</div>
              <div class="pu-val {oc}">{os_}{abs(pu_occ):.2f}%</div>
            </div>"""

        html += "</div></div>"
        return html

    CFG_MESES = [
        {'nome': 'Abril',  'days': 30, 'start': 3},
        {'nome': 'Maio',   'days': 31, 'start': 5},
        {'nome': 'Junho',  'days': 30, 'start': 1},
    ]


    def _pickup_gerar_html_pdf(meses, metas, cfg_meses, nome_hotel):
        """Gera HTML com os 3 meses empilhados, pronto para imprimir como PDF com fundo escuro."""
        css_global = """
        <style>
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0a0a0a !important; padding: 16px; }
        .mes-bloco { page-break-after: always; margin-bottom: 30px; }
        .mes-bloco:last-child { page-break-after: avoid; }
        @media print {
            body { background: #0a0a0a !important; }
            .mes-bloco { page-break-after: always; }
            .mes-bloco:last-child { page-break-after: avoid; }
        }
        </style>
        """
        html_completo = f"<html><head><meta charset='utf-8'>{css_global}</head><body>"
        for mi, cfg in enumerate(cfg_meses):
            bloco = _pickup_render(meses, metas, cfg_meses, mi)
            html_completo += f"<div class='mes-bloco'>{bloco}</div>"
        html_completo += "</body></html>"
        return html_completo

    if not os.path.exists(PICKUP_FILE):
        st.error(f"❌ Arquivo '{PICKUP_FILE}' não encontrado na pasta do programa.")
        st.info("Coloque o arquivo 'Pick-up_sniper.xlsx' na mesma pasta do monitor.py")
    else:
        hotel_tab1, hotel_tab2, hotel_tab3 = st.tabs(["🏨 Alto da Boa Vista", "🏩 Normandie", "🏡 Chales Villa Caravaggio"])

        with hotel_tab1:
            meses_abv, _ = _pickup_ler_dados("Alto da Boa Vista")
            metas_abv = _pickup_ler_metas()
            if meses_abv:
                if 'abv_mes' not in st.session_state: st.session_state['abv_mes'] = 0
                cols_m = st.columns(3)
                for i, cfg in enumerate(CFG_MESES):
                    if cols_m[i].button(f"📅 {cfg['nome']}", key=f"abv_{i}", use_container_width=True):
                        st.session_state['abv_mes'] = i
                mi = st.session_state['abv_mes']
                st.markdown(f"### {CFG_MESES[mi]['nome']} 2026 — Pousada Alto da Boa Vista")
                components.html(_pickup_render(meses_abv, metas_abv, CFG_MESES, mi), height=1000, scrolling=True)
                st.markdown("---")
                html_pdf_abv = _pickup_gerar_html_pdf(meses_abv, metas_abv, CFG_MESES, "Pousada Alto da Boa Vista")
                b64_abv = base64.b64encode(html_pdf_abv.encode("utf-8")).decode()
                href_abv = f'data:text/html;base64,{b64_abv}'
                st.markdown(f'''<a href="{href_abv}" download="pickup_alto_da_boa_vista.html" target="_blank"><button style="background:#002D62;color:white;border:none;padding:10px 22px;border-radius:6px;font-weight:bold;font-size:14px;cursor:pointer;width:100%;">📄 Exportar PDF — Alto da Boa Vista</button></a>''', unsafe_allow_html=True)
                if st.button("📊 GERAR PPTX — Alto da Boa Vista", key="btn_pptx_alto", use_container_width=True):
                    import subprocess, datetime as _dt
                    _pasta    = os.path.dirname(os.path.abspath(__file__))
                    _script   = os.path.join(_pasta, "gerar_pickup_pptx.py")
                    _xlsx     = os.path.join(_pasta, "Pick-up_sniper.xlsx")
                    _tmpl     = os.path.join(_pasta, "pickup_alto_template.pptx")
                    _saida    = os.path.join(_pasta, f"pickup_alto_{_dt.date.today().strftime('%d%m%Y')}.pptx")
                    if not os.path.exists(_script): st.error(f"❌ Script não encontrado: {_script}")
                    elif not os.path.exists(_xlsx):  st.error(f"❌ Planilha não encontrada: {_xlsx}")
                    elif not os.path.exists(_tmpl):  st.error(f"❌ Template não encontrado: {_tmpl}")
                    else:
                        _sp = st.empty(); _sp.markdown("⏳ Gerando PPTX...")
                        _r = subprocess.run(["python", _script, _xlsx, _tmpl, "alto", _saida], capture_output=True, text=True, timeout=180)
                        if _r.returncode == 0:
                            with open(_saida, "rb") as _fp:
                                st.download_button(f"📥 Baixar {os.path.basename(_saida)}", _fp.read(), os.path.basename(_saida), "application/vnd.openxmlformats-officedocument.presentationml.presentation", key="dl_pptx_alto")
                            _sp.success(f"✅ PPTX gerado!")
                        else:
                            _sp.error("❌ Erro ao gerar PPTX")
                            st.code(_r.stderr[-500:] if _r.stderr else _r.stdout[-500:])
            else:
                st.error("❌ Não foi possível carregar dados.")

        with hotel_tab2:
            meses_norm, dm_norm = _pickup_ler_dados("Normandie")
            if meses_norm:
                if 'norm_mes' not in st.session_state: st.session_state['norm_mes'] = 0
                cols_n = st.columns(3)
                for i, cfg in enumerate(CFG_MESES):
                    if cols_n[i].button(f"📅 {cfg['nome']}", key=f"norm_{i}", use_container_width=True):
                        st.session_state['norm_mes'] = i
                mi2 = st.session_state['norm_mes']
                st.markdown(f"### {CFG_MESES[mi2]['nome']} 2026 — Hotel Normandie")
                # Injeta diária média lida da planilha (F35, N36, V36)
                metas_norm_inj = {
                    i: {'dm_acum': dm_norm.get(i, 0), 'rec_acum': 0, 'rec_meta': 0,
                        'occ_acum': 0, 'occ_meta': 0, 'dm_meta': 0, 'dias_faltam': 0, 'falta_dia': 0}
                    for i in range(3)
                }
                components.html(_pickup_render(meses_norm, metas_norm_inj, CFG_MESES, mi2), height=1000, scrolling=True)
                st.markdown("---")
                html_pdf_norm = _pickup_gerar_html_pdf(meses_norm, metas_norm_inj, CFG_MESES, "Hotel Normandie")
                b64_norm = base64.b64encode(html_pdf_norm.encode("utf-8")).decode()
                href_norm = f'data:text/html;base64,{b64_norm}'
                st.markdown(f'''<a href="{href_norm}" download="pickup_normandie.html" target="_blank"><button style="background:#002D62;color:white;border:none;padding:10px 22px;border-radius:6px;font-weight:bold;font-size:14px;cursor:pointer;width:100%;">📄 Exportar PDF — Hotel Normandie</button></a>''', unsafe_allow_html=True)
                if st.button("📊 GERAR PPTX — Hotel Normandie", key="btn_pptx_norm", use_container_width=True):
                    import subprocess, datetime as _dt
                    _pasta_n  = os.path.dirname(os.path.abspath(__file__))
                    _script_n = os.path.join(_pasta_n, "gerar_pickup_pptx.py")
                    _xlsx_n   = os.path.join(_pasta_n, "Pick-up_sniper.xlsx")
                    _tmpl_n   = os.path.join(_pasta_n, "pickup_normandie_template.pptx")
                    _saida_n  = os.path.join(_pasta_n, f"pickup_normandie_{_dt.date.today().strftime('%d%m%Y')}.pptx")
                    if not os.path.exists(_script_n): st.error(f"❌ Script não encontrado: {_script_n}")
                    elif not os.path.exists(_xlsx_n):  st.error(f"❌ Planilha não encontrada: {_xlsx_n}")
                    elif not os.path.exists(_tmpl_n):  st.error(f"❌ Template não encontrado: {_tmpl_n}")
                    else:
                        _sp_n = st.empty(); _sp_n.markdown("⏳ Gerando PPTX...")
                        _r_n = subprocess.run(["python", _script_n, _xlsx_n, _tmpl_n, "normandie", _saida_n], capture_output=True, text=True, timeout=180)
                        if _r_n.returncode == 0:
                            with open(_saida_n, "rb") as _fp_n:
                                st.download_button(f"📥 Baixar {os.path.basename(_saida_n)}", _fp_n.read(), os.path.basename(_saida_n), "application/vnd.openxmlformats-officedocument.presentationml.presentation", key="dl_pptx_norm")
                            _sp_n.success(f"✅ PPTX gerado!")
                        else:
                            _sp_n.error("❌ Erro ao gerar PPTX")
                            st.code(_r_n.stderr[-500:] if _r_n.stderr else _r_n.stdout[-500:])
            else:
                st.error("❌ Não foi possível carregar dados.")

        with hotel_tab3:
            meses_chales, dm_chales_dict = _pickup_ler_dados("CHALES VILLA CARAVAGGIO")
            if meses_chales:
                if 'chales_mes' not in st.session_state: st.session_state['chales_mes'] = 0
                cols_c = st.columns(3)
                for i, cfg in enumerate(CFG_MESES):
                    if cols_c[i].button(f"📅 {cfg['nome']}", key=f"chales_{i}", use_container_width=True):
                        st.session_state['chales_mes'] = i
                mi3 = st.session_state['chales_mes']
                st.markdown(f"### {CFG_MESES[mi3]['nome']} 2026 — Chales Villa Caravaggio")

                def _chales_ler_dm():
                    if not os.path.exists(PICKUP_FILE):
                        return {}
                    try:
                        from openpyxl import load_workbook
                        wb = load_workbook(PICKUP_FILE, read_only=True, data_only=True)
                        ws = wb['CHALES VILLA CARAVAGGIO']
                        rows = list(ws.iter_rows(values_only=True))
                        def _safe(r, c):
                            try:
                                v = rows[r][c]
                                return float(v) if isinstance(v, (int, float)) else 0.0
                            except: return 0.0
                        return {
                            0: _safe(34, 4),   # E35
                            1: _safe(35, 12),  # M36
                            2: _safe(34, 21),  # V35
                        }
                    except:
                        return {}

                dm_chales = _chales_ler_dm()
                metas_chales_inj = {
                    i: {
                        'dm_acum': dm_chales.get(i, 0),
                        'rec_acum': 0, 'rec_meta': 0,
                        'occ_acum': 0, 'occ_meta': 0,
                        'dm_meta': 0, 'dias_faltam': 0, 'falta_dia': 0
                    }
                    for i in range(3)
                }

                components.html(_pickup_render(meses_chales, metas_chales_inj, CFG_MESES, mi3), height=1000, scrolling=True)
                st.markdown("---")

                html_pdf_chales = _pickup_gerar_html_pdf(meses_chales, metas_chales_inj, CFG_MESES, "Chales Villa Caravaggio")
                b64_chales = base64.b64encode(html_pdf_chales.encode("utf-8")).decode()
                href_chales = f'data:text/html;base64,{b64_chales}'
                st.markdown(f'''<a href="{href_chales}" download="pickup_chales_villa_caravaggio.html" target="_blank"><button style="background:#002D62;color:white;border:none;padding:10px 22px;border-radius:6px;font-weight:bold;font-size:14px;cursor:pointer;width:100%;">📄 Exportar PDF — Chales Villa Caravaggio</button></a>''', unsafe_allow_html=True)

                if st.button("📊 GERAR PPTX — Chales Villa Caravaggio", key="btn_pptx_chales", use_container_width=True):
                    import subprocess, datetime as _dt
                    _pasta_c  = os.path.dirname(os.path.abspath(__file__))
                    _script_c = os.path.join(_pasta_c, "gerar_pickup_pptx.py")
                    _xlsx_c   = os.path.join(_pasta_c, "Pick-up_sniper.xlsx")
                    _tmpl_c   = os.path.join(_pasta_c, "pickup_chales_vilaa_caravaggio_template.pptx")
                    _saida_c  = os.path.join(_pasta_c, f"pickup_chales_{_dt.date.today().strftime('%d%m%Y')}.pptx")
                    if not os.path.exists(_script_c): st.error(f"❌ Script não encontrado: {_script_c}")
                    elif not os.path.exists(_xlsx_c):  st.error(f"❌ Planilha não encontrada: {_xlsx_c}")
                    elif not os.path.exists(_tmpl_c):  st.error(f"❌ Template não encontrado: {_tmpl_c}")
                    else:
                        _sp_c = st.empty(); _sp_c.markdown("⏳ Gerando PPTX...")
                        _r_c = subprocess.run(
                            ["python", _script_c, _xlsx_c, _tmpl_c, "chales", _saida_c],
                            capture_output=True, text=True, timeout=180
                        )
                        if _r_c.returncode == 0:
                            with open(_saida_c, "rb") as _fp_c:
                                st.download_button(
                                    f"📥 Baixar {os.path.basename(_saida_c)}",
                                    _fp_c.read(),
                                    os.path.basename(_saida_c),
                                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    key="dl_pptx_chales"
                                )
                            _sp_c.success(f"✅ PPTX gerado!")
                        else:
                            _sp_c.error("❌ Erro ao gerar PPTX")
                            st.code(_r_c.stderr[-500:] if _r_c.stderr else _r_c.stdout[-500:])
            else:
                st.error("❌ Não foi possível carregar dados.")


# ABA 9: FECHO AUTOMÁTICO AMANDA
# ==============================================================================
elif aba_selecionada == "🤖 Fecho Automático Amanda":
    st.title("🤖 FECHO AUTOMÁTICO AMANDA")
    st.subheader("Lê ticket → Interpreta com IA → Fecha no Omnibees → Avisa no Telegram")

    AMANDA_URL       = "https://chat2.appamanda.com.br/tickets"
    AMANDA_LOGIN     = "daniel@easyhoteis.com"
    AMANDA_SENHA     = "123456"
    CONTATO_ALVO     = "5511999837787"
    TELEGRAM_BOT     = "https://uncoopered-collectivistic-fern.ngrok-free.dev"
    TELEGRAM_CHAT_ID = "7591049082"
    CLAUDE_API_KEY   = "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"

    TODOS_QUARTOS_AMANDA = [
        "Duplex Superior",
        "Suite Presidencial com Varanda",
        "Suíte Presidencial",
        "Quarto Duplo com Varanda",
        "Quarto Duplo Luxo com Sauna",
        "Chalé Exclusivo com Cozinha e Sala de Estar",
        "Chalé Família Exclusivo com Cozinha",
        "Chalé Duplex com Varanda",
        "Duplex com Piscina e lareira",
        "Chalé Duplex com Ofurô",
        "Suíte com Varanda",
        "Chalé Deluxe com Banheira e Lareira",
        "Quarto Duplo com Banheira",
        "Duplo com Lareira",
    ]

    if 'amanda_resultado' not in st.session_state:
        st.session_state.amanda_resultado = None

    st.markdown("### 📋 Como funciona")
    st.info(
        "1. Clique em **EXECUTAR**\n"
        "2. O robô entra no Amanda e lê a mensagem do contato 5511999837787 \n"
        "3. A IA interpreta a mensagem e extrai datas e quartos \n"
        "4. O robô captura o código 2FA automaticamente \n"
        "5. Executa o fecho no Omnibees e te avisa no Telegram"
    )
    st.markdown("---")

    def interpretar_mensagem_amanda(mensagem):
        """Interpreta a mensagem usando regex puro — sem depender de API."""
        import re as re_amanda
        from datetime import date as date_cls

        msg = mensagem.lower().strip()

        # ── Mapeamento de quartos ─────────────────────────────────────────────
        MAPA_QUARTOS = [
            (["duplex superior", "duplex sup"],                         "Duplex Superior"),
            (["suite presidencial com varanda", "suite pres varanda",
              "suíte presidencial com varanda"],                        "Suite Presidencial com Varanda"),
            (["suite presidencial", "suíte presidencial",
              "suite pres", "suíte pres"],                              "Suíte Presidencial"),
            (["quarto duplo com varanda", "duplo com varanda",
              "duplo varanda"],                                         "Quarto Duplo com Varanda"),
            (["quarto duplo luxo", "duplo luxo", "luxo sauna"],         "Quarto Duplo Luxo com Sauna"),
            (["chalé exclusivo", "chale exclusivo",
              "exclusivo cozinha"],                                     "Chalé Exclusivo com Cozinha e Sala de Estar"),
            (["chalé família", "chale familia",
              "familia cozinha"],                                       "Chalé Família Exclusivo com Cozinha"),
            (["chalé duplex com varanda", "chale duplex varanda"],      "Chalé Duplex com Varanda"),
            (["duplex piscina", "piscina lareira",
              "duplex com piscina"],                                    "Duplex com Piscina e lareira"),
            (["chalé duplex com ofurô", "chale duplex ofuro",
              "chalé ofurô", "ofuro"],                                  "Chalé Duplex com Ofurô"),
            (["suíte com varanda", "suite com varanda",
              "suite varanda"],                                         "Suíte com Varanda"),
            (["chalé deluxe", "chale deluxe", "banheira lareira",
              "deluxe banheira"],                                       "Chalé Deluxe com Banheira e Lareira"),
            (["quarto duplo com banheira", "duplo banheira"],           "Quarto Duplo com Banheira"),
            (["duplo com lareira", "duplo lareira"],                    "Duplo com Lareira"),
        ]

        # ── Extrai quartos ────────────────────────────────────────────────────
        todos_quartos = any(p in msg for p in ["todos", "tudo", "todas", "all"])
        quartos = []
        if not todos_quartos:
            for aliases, nome_oficial in MAPA_QUARTOS:
                for alias in aliases:
                    if alias in msg:
                        if nome_oficial not in quartos:
                            quartos.append(nome_oficial)
                        break

        # ── Extrai datas ──────────────────────────────────────────────────────
        # Padrões: 30/03, 30/03/2026, 30/03/26
        datas = re_amanda.findall(r"(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?", mensagem)
        ano_atual = date_cls.today().year

        datas_obj = []
        for d in datas:
            try:
                dia = int(d[0])
                mes = int(d[1])
                ano = int(d[2]) if d[2] else ano_atual
                if ano < 100:
                    ano += 2000
                datas_obj.append(date_cls(ano, mes, dia))
            except:
                pass

        if not datas_obj:
            return None

        data_ini = min(datas_obj)
        data_fim = max(datas_obj)

        return {
            "data_ini": data_ini.strftime("%d/%m/%Y"),
            "data_fim": data_fim.strftime("%d/%m/%Y"),
            "quartos": quartos,
            "todos_quartos": todos_quartos
        }

    def enviar_telegram_amanda(mensagem):
        import urllib.request, json as jsonlib
        try:
            payload = jsonlib.dumps({
                "chat_id": TELEGRAM_CHAT_ID,
                "text": mensagem
            }).encode("utf-8")
            req = urllib.request.Request(
                f"{TELEGRAM_BOT}/webhook/7c29bc12-edf9-400e-a1ed-f5e40e9ef126",
                data=payload,
                headers={"Content-Type": "application/json"}
            )
            urllib.request.urlopen(req, timeout=5)
        except:
            pass


    def enviar_mensagem_amanda_chat(driver, wait, mensagem):
        """Digita e envia mensagem no chat do Amanda."""
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys
        import time
        try:
            campo = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                "textarea.MuiInputBase-input:not([aria-hidden])"
            )))
            campo.click()
            time.sleep(0.3)
            campo.send_keys(mensagem)
            time.sleep(0.3)
            campo.send_keys(Keys.RETURN)
            time.sleep(1)
        except Exception as e:
            pass

    def executar_fecho_amanda():
        status = st.empty()
        prog   = st.progress(0)
        _log_aresta("Iniciando fecho automático Amanda")

        # ── PASSO 1: Abre Amanda e lê mensagem ───────────────────────────────
        status.markdown("🌐 Abrindo Amanda...")
        driver_a = criar_driver()
        wait_a = WebDriverWait(driver_a, 20)
        mensagem_cliente = None

        try:
            driver_a.get(AMANDA_URL)
            time.sleep(3)

            # Login
            try:
                campo_email = driver_a.find_element(By.CSS_SELECTOR,
                    "input[type='email'], input[name='email'], #email")
                campo_email.clear()
                campo_email.send_keys(AMANDA_LOGIN)
                campo_senha = driver_a.find_element(By.CSS_SELECTOR,
                    "input[type='password']")
                campo_senha.clear()
                campo_senha.send_keys(AMANDA_SENHA)
                driver_a.find_element(By.CSS_SELECTOR,
                    "button[type='submit']").click()
                status.markdown("🔐 Login Amanda...")
                time.sleep(4)
            except:
                status.markdown("🔄 Já logado no Amanda...")
                time.sleep(2)

            prog.progress(10)

            # Verifica e reconecta sessão se necessário
            _verificar_sessao_amanda(driver_a, wait_a, AMANDA_URL, AMANDA_LOGIN, AMANDA_SENHA, status)

            # Aguarda 3s para os popups aparecerem após login
            status.markdown("🔔 Aguardando popups...")
            time.sleep(3)

            # Popup 1: clica em "Agora não"
            try:
                btn1 = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[contains(@class,'MuiButton-label') and normalize-space(text())='Agora não']/.."
                )))
                driver_a.execute_script("arguments[0].click();", btn1)
                status.markdown("✅ Popup 1 fechado...")
                time.sleep(1)
            except:
                pass

            # Popup 2: clica em "Fechar"
            try:
                btn2 = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[contains(@class,'MuiButton-label') and normalize-space(text())='Fechar']/.."
                )))
                driver_a.execute_script("arguments[0].click();", btn2)
                status.markdown("✅ Popup 2 fechado...")
                time.sleep(1)
            except:
                pass

            prog.progress(20)

            # Clica no botão de notificações pelo seletor CSS exato
            status.markdown("🔔 Abrindo notificações...")
            try:
                btn_notif = wait_a.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "#root > div.app-container.light > header > div > div.navbar-actions > div.notifications > button"
                )))
                driver_a.execute_script("arguments[0].click();", btn_notif)
                time.sleep(2)
            except:
                pass

            prog.progress(30)

            # Clica no ticket do contato 5511999837787
            status.markdown(f"🔍 Abrindo ticket de {CONTATO_ALVO}...")
            try:
                ticket = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    f"//span[contains(@class,'MuiTypography-body2') and contains(text(),'{CONTATO_ALVO}')]"
                    f"/ancestor::li | "
                    f"//span[contains(@class,'MuiTypography-body2') and contains(text(),'{CONTATO_ALVO}')]"
                    f"/ancestor::div[contains(@class,'MuiListItem')]"
                )))
                driver_a.execute_script("arguments[0].click();", ticket)
                time.sleep(3)
            except:
                status.error(f"❌ Ticket de {CONTATO_ALVO} não encontrado!")
                driver_a.quit()
                return

            prog.progress(35)

            # Passo 5: Clica em "Aceitar" — botão no topo da conversa
            status.markdown("✅ Aceitando ticket...")
            try:
                btn_aceitar = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[contains(@class,'MuiButton-label') and normalize-space(text())='Aceitar']/.."
                )))
                btn_aceitar.click()
                time.sleep(2)
            except:
                pass

            # Passo 6: Aguarda popup "Aceitar Chat" e clica no dropdown "Selecionar setor"
            status.markdown("📂 Selecionando setor...")
            try:
                btn_setor = wait_a.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "div.MuiDialogContent-root div.MuiSelect-root.MuiSelect-select"
                )))
                btn_setor.click()
                time.sleep(1.5)
            except:
                pass

            # Passo 7: Clica em "Disponibilidade ( URGENTE )" pelo texto exato
            status.markdown("🚨 Selecionando Disponibilidade URGENTE...")
            try:
                opt_urgente = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    "//li[contains(text(),'Disponibilidade') and contains(text(),'URGENTE')]"
                )))
                opt_urgente.click()
                time.sleep(1)
            except:
                pass

            # Passo 8: Clica em "iniciar"
            status.markdown("▶️ Iniciando atendimento...")
            try:
                btn_iniciar = wait_a.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[contains(@class,'MuiButton-label') and normalize-space(text())='iniciar']/.."
                )))
                btn_iniciar.click()
                time.sleep(2)
            except:
                pass

            prog.progress(45)

            # PRIMEIRO lê a mensagem ANTES de enviar qualquer aviso
            status.markdown("📖 Aguardando mensagens carregarem...")
            try:
                wait_a.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                    "#messagesList"
                )))
            except:
                time.sleep(5)

            time.sleep(2)

            status.markdown("📖 Lendo mensagem de fecho...")
            time.sleep(2)
            mensagem_cliente = driver_a.execute_script("""
                var lista = document.getElementById("messagesList");
                if (!lista) return null;
                // Pega todos os divs filhos diretos do messagesList
                var filhos = Array.from(lista.querySelectorAll(":scope > div"));
                // Filtra só os que têm texto real (ignora separadores)
                var mensagens = filhos.filter(function(el) {
                    return el.innerText && el.innerText.trim().length > 2;
                });
                // Percorre de trás para frente pulando mensagens do robô
                // Robô = tem SVG de check duplo no span de horário
                for (var i = mensagens.length - 1; i >= 0; i--) {
                    var msg = mensagens[i];
                    var temCheckDuplo = msg.querySelector("svg.jss432, svg.jss384, svg.jss382");
                    if (temCheckDuplo) continue;
                    // Ignora mensagens do robô e do sistema
                    if (msg.innerText.indexOf("Vamos efetuar o fecho") !== -1) continue;
                    if (msg.innerText.indexOf("Fecho realizado com sucesso") !== -1) continue;
                    if (msg.innerText.indexOf("Lili") !== -1) continue;
                    if (msg.innerText.indexOf("Solicitação de fecho de disponibilidade") !== -1) continue;
                    if (msg.innerText.indexOf("Suporte com tarifas") !== -1) continue;
                    if (msg.innerText.indexOf("Falar com um atendente") !== -1) continue;
                    // É mensagem do cliente — pega o texto sem o horário
                    var clone = msg.cloneNode(true);
                    clone.querySelectorAll("button, span").forEach(function(el) {
                        if (el.querySelector("svg") || /^\d{1,2}:\d{2}$/.test(el.innerText.trim())) {
                            el.remove();
                        }
                    });
                    var txt = clone.innerText.trim();
                    if (txt && txt.length > 2) return txt;
                }
                return null;
            """)

            # AGORA envia mensagem de aviso (após ter lido a mensagem do cliente)
            status.markdown("💬 Enviando mensagem de aviso no chat...")
            try:
                enviar_mensagem_amanda_chat(driver_a, wait_a,
                    "Olá! Vamos efetuar o fecho de disponibilidade agora. Aguarde!")
            except:
                pass

            driver_a.quit()

            if not mensagem_cliente:
                status.error("❌ Não foi possível ler a mensagem do ticket!")
                return

            status.markdown(f"📩 Mensagem: **{mensagem_cliente}**")
            _log_aresta(f"Mensagem lida: {mensagem_cliente}")
            prog.progress(50)

        except Exception as e:
            try: driver_a.quit()
            except: pass
            status.error(f"❌ Erro no Amanda: {e}")
            return

                # ── PASSO 2: Interpreta com IA ────────────────────────────────────────
        status.markdown("🧠 Interpretando com IA...")
        resultado = interpretar_mensagem_amanda(mensagem_cliente)

        if not resultado or not resultado.get("data_ini"):
            status.error(f"❌ Não consegui interpretar: '{mensagem_cliente}'")
            return

        data_ini_str = resultado["data_ini"]
        data_fim_str = resultado["data_fim"]
        todos_q      = resultado.get("todos_quartos", False)
        quartos_msg  = resultado.get("quartos", [])

        if todos_q or not quartos_msg:
            quartos_fecho = TODOS_QUARTOS_AMANDA
            quartos_txt   = "Todos os quartos"
        else:
            quartos_fecho = quartos_msg
            quartos_txt   = ", ".join(quartos_msg)

        try:
            data_ini_omni = date(int(data_ini_str[6:]), int(data_ini_str[3:5]), int(data_ini_str[:2]))
            data_fim_omni = date(int(data_fim_str[6:]), int(data_fim_str[3:5]), int(data_fim_str[:2]))
        except:
            status.error(f"❌ Data inválida: {data_ini_str} / {data_fim_str}")
            return

        status.markdown(f"✅ **{quartos_txt}** | **{data_ini_str}** a **{data_fim_str}**")
        prog.progress(55)

        # ── PASSO 3: Captura 2FA ──────────────────────────────────────────────
        status.markdown("🔐 Capturando código 2FA...")
        codigo_2fa = _capturar_2fa_com_retry(status)
        if not codigo_2fa or len(codigo_2fa) != 6:
            status.error("❌ Não foi possível capturar o código 2FA!")
            _log_aresta("ERRO: 2FA não capturado no fecho Amanda", "ERRO")
            return
        prog.progress(65)

        # ── PASSO 4: Executa fecho no Omnibees ───────────────────────────────
        config_omni = {
            "login": "daniel.altodaboavista",
            "senha": "@Altodaboavista2026#",
            "tarifarios": [
                "Trf Bancorbras", "CLUBE MONTREAL", "Tarifa Flexível",
                "Condição Exclusiva", "Programa Preferencial - Orinter",
                "Tarifa Não Reembolsável.", "Tarifa Site",
            ]
        }

        status.markdown("🐝 Abrindo Omnibees...")
        driver_o = criar_driver()
        wait_o = WebDriverWait(driver_o, 15)

        try:
            _login_omnibees(driver_o, wait_o, config_omni, codigo_2fa, status)
            prog.progress(73)

            _selecionar_datas(driver_o, wait_o, data_ini_omni, data_fim_omni, status)
            prog.progress(80)

            _selecionar_tarifarios(driver_o, wait_o, True, config_omni["tarifarios"], status)
            prog.progress(86)

            sel_todos_q = todos_q or not quartos_msg
            _selecionar_quartos(driver_o, wait_o, sel_todos_q, quartos_fecho, status)
            prog.progress(92)

            # Fechar vendas
            status.markdown("🔒 Abrindo aba Fechar/Abrir Vendas...")
            try:
                wait_o.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "#ob-ui-update-rates-packages-tab > li:nth-child(3)"))).click()
                time.sleep(1.5)
            except Exception as e:
                status.markdown(f"⚠️ Aba Fechar/Abrir: {e}")

            status.markdown("🔒 Selecionando Fechar Vendas...")
            try:
                wait_o.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "#rd-updateRates-input-close-open-sales-2"))).click()
                time.sleep(1.5)
            except Exception as e:
                status.markdown(f"⚠️ Fechar Vendas: {e}")

            # ⚠️ SALVAR DESATIVADO — em modo de teste
            status.markdown("✅ Fecho selecionado! Salvar desativado (modo teste).")
            time.sleep(2)

            try:
                driver_o.quit()
            except:
                pass
            prog.progress(100)

            # ── PASSO 5: Envia confirmação no chat do Amanda ─────────────────
            status.markdown("💬 Enviando confirmação no Amanda...")
            try:
                driver_conf = criar_driver()
                wait_conf = WebDriverWait(driver_conf, 15)
                driver_conf.get(AMANDA_URL)
                time.sleep(3)

                # Login
                try:
                    campo_email = driver_conf.find_element(By.CSS_SELECTOR,
                        "input[type='email'], input[name='email'], #email")
                    campo_email.clear()
                    campo_email.send_keys(AMANDA_LOGIN)
                    campo_senha = driver_conf.find_element(By.CSS_SELECTOR,
                        "input[type='password']")
                    campo_senha.clear()
                    campo_senha.send_keys(AMANDA_SENHA)
                    driver_conf.find_element(By.CSS_SELECTOR,
                        "button[type='submit']").click()
                    status.markdown("🔐 Login Amanda (confirmação)...")
                    time.sleep(4)
                except:
                    time.sleep(2)

                # Fecha popups com timeout curto (3s)
                wait_popup = WebDriverWait(driver_conf, 3)
                try:
                    btn1 = wait_popup.until(EC.element_to_be_clickable((By.XPATH,
                        "//span[contains(@class,'MuiButton-label') and normalize-space(text())='Agora não']/.."
                    )))
                    driver_conf.execute_script("arguments[0].click();", btn1)
                    time.sleep(0.8)
                except:
                    pass
                try:
                    btn2 = wait_popup.until(EC.element_to_be_clickable((By.XPATH,
                        "//span[contains(@class,'MuiButton-label') and normalize-space(text())='Fechar']/.."
                    )))
                    driver_conf.execute_script("arguments[0].click();", btn2)
                    time.sleep(0.8)
                except:
                    pass

                # Clica no ticket pelo número do contato via JavaScript
                status.markdown("🔍 Abrindo ticket para confirmação...")
                driver_conf.execute_script("""
                    var contato = arguments[0];
                    var spans = document.querySelectorAll('span.MuiTypography-body2');
                    for (var s of spans) {
                        if (s.innerText.indexOf(contato) !== -1) {
                            // Sobe até o div com role=button
                            var parent = s.parentElement;
                            for (var i = 0; i < 8; i++) {
                                if (!parent) break;
                                if (parent.getAttribute('role') === 'button') {
                                    parent.click();
                                    return;
                                }
                                parent = parent.parentElement;
                            }
                        }
                    }
                """, CONTATO_ALVO)
                time.sleep(2)

                # Envia mensagem de confirmação
                msg_confirmacao = (
                    f"✅ Fecho realizado com sucesso! "
                    f"Período: {data_ini_str} a {data_fim_str} | "
                    f"Quarto: {quartos_txt}"
                )
                enviar_mensagem_amanda_chat(driver_conf, wait_conf, msg_confirmacao)
                time.sleep(1)

                # Fecha o ticket clicando no X vermelho da lista lateral
                status.markdown("🔒 Fechando ticket...")
                try:
                    driver_conf.execute_script("""
                        var btns = document.querySelectorAll('button');
                        for (var btn of btns) {
                            var svg = btn.querySelector('svg');
                            if (svg && svg.getAttribute('title') === 'Fechar') {
                                btn.click();
                                return;
                            }
                        }
                    """)
                    time.sleep(1)
                except:
                    pass

                driver_conf.quit()
                status.markdown("✅ Confirmação enviada e ticket fechado!")
            except Exception as e:
                status.markdown(f"⚠️ Não foi possível enviar confirmação no Amanda: {e}")

            # ── PASSO 6: Avisa no Telegram ────────────────────────────────────
            msg = (
                f"✅ FECHO EXECUTADO — Alto da Boa Vista\n"
                f"📅 Período: {data_ini_str} a {data_fim_str}\n"
                f"🛏️ Quartos: {quartos_txt}\n"
                f"🏷️ Tarifários: Todos\n"
                f"📩 Mensagem: {mensagem_cliente}"
            )
            enviar_telegram_amanda(msg)
            status.success(f"✅ Fecho concluído! Amanda e Telegram notificados!")
            _log_aresta(f"SUCESSO: Fecho realizado | Quartos: {quartos_txt} | Período: {data_ini_str} a {data_fim_str}")
            st.session_state.amanda_resultado = msg
            historico_salvar(
                funcao="Fecho Automático Amanda",
                status="✅ Sucesso",
                detalhes={
                    "hotel": "Alto da Boa Vista",
                    "quartos": quartos_txt,
                    "data_ini": data_ini_str,
                    "data_fim": data_fim_str,
                    "mensagem": mensagem_cliente
                }
            )

        except Exception as e:
            try: driver_o.quit()
            except: pass
            status.error(f"❌ Erro no Omnibees: {e}")
            _log_aresta(f"ERRO no Omnibees: {e}", "ERRO")
            enviar_telegram_amanda(f"❌ Erro: {e}\nMensagem: {mensagem_cliente}")
            historico_salvar(
                funcao="Fecho Automático Amanda",
                status=f"❌ Erro: {str(e)[:100]}",
                detalhes={
                    "hotel": "Alto da Boa Vista",
                    "mensagem": mensagem_cliente if mensagem_cliente else ""
                }
            )

    if st.button("🚀 EXECUTAR FECHO AUTOMÁTICO", key="amanda_executar", use_container_width=True):
        sucesso, erro_msg = _executar_com_watchdog(executar_fecho_amanda, timeout_segundos=300)
        if not sucesso and erro_msg == "Processo cancelado por timeout":
            st.error("⚠️ O robô travou e foi cancelado automaticamente após 5 minutos. Verifique o log.")

    if st.session_state.amanda_resultado:
        st.markdown("---")
        st.markdown("### 📋 Último fecho executado")
        st.info(st.session_state.amanda_resultado)

# ==============================================================================
# ABA HISTÓRICO
# ==============================================================================
elif aba_selecionada == "📋 Histórico":
    st.title("📋 HISTÓRICO DE EXECUÇÕES")
    st.subheader("Registros das últimas 24 horas")

    registros = historico_carregar()

    col_ref, col_limpar = st.columns([4, 1])
    with col_ref:
        st.markdown(f"**{len(registros)} registro(s) nas últimas 24h**")
    with col_limpar:
        if st.button("🗑️ Limpar", key="limpar_historico"):
            try:
                if os.path.exists(HISTORICO_FILE):
                    os.remove(HISTORICO_FILE)
                st.success("Histórico limpo!")
                st.rerun()
            except:
                pass

    st.markdown("---")

    if not registros:
        st.info("Nenhuma execução registrada nas últimas 24 horas.")
    else:
        # Exibe do mais recente para o mais antigo
        for r in reversed(registros):
            try:
                dt = datetime.fromisoformat(r["data_hora"])
                dt_str = dt.strftime("%d/%m/%Y %H:%M:%S")
                status_icon = "✅" if "✅" in r["status"] else "❌"
                d = r.get("detalhes", {})

                # Cores adaptadas ao tema
                ok = "✅" in r["status"]
                cor_fundo = "rgba(74,222,128,0.12)" if ok else "rgba(248,113,113,0.12)"
                cor_borda  = "#4ade80"  if ok else "#f87171"
                cor_titulo = "#4ade80"  if ok else "#f87171"
                cor_texto  = "#e5e7eb"
                cor_hora   = "#9ca3af"
                cor_label  = "#FFD700"

                linhas_html = []
                if d.get('hotel'):    linhas_html.append(f'<span style="color:{cor_label}"><b>Hotel:</b></span> {d["hotel"]}')
                if d.get('quartos'):  linhas_html.append(f'<span style="color:{cor_label}"><b>Quarto(s):</b></span> {d["quartos"]}')
                if d.get('bar'):      linhas_html.append(f'<span style="color:{cor_label}"><b>BAR:</b></span> {d["bar"]}')
                if d.get('data_ini'): linhas_html.append(f'<span style="color:{cor_label}"><b>Período:</b></span> {d["data_ini"]} a {d["data_fim"]}')
                if d.get('datas'):    linhas_html.append(f'<span style="color:{cor_label}"><b>Datas:</b></span> {d["datas"]}')
                if d.get('mensagem'): linhas_html.append(f'<span style="color:{cor_label}"><b>Mensagem:</b></span> <i>{d["mensagem"]}</i>')
                corpo = "<br>".join(linhas_html)

                html_card = (
                    f'<div style="background:{cor_fundo};border-left:5px solid {cor_borda};'
                    f'border-radius:10px;padding:16px 20px;margin-bottom:12px;">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">'
                    f'<span style="font-size:16px;font-weight:800;color:{cor_titulo};">{status_icon} {r["funcao"]}</span>'
                    f'<span style="font-size:12px;color:{cor_hora};">🕐 {dt_str}</span>'
                    f'</div>'
                    f'<div style="font-size:13px;color:{cor_texto};line-height:1.9;">'
                    f'<span style="color:{cor_label}"><b>Status:</b></span> {r["status"]}<br>{corpo}'
                    f'</div></div>'
                )
                st.markdown(html_card, unsafe_allow_html=True)
            except:
                pass