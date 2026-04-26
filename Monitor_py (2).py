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
