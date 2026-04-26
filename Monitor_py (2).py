import streamlit as st
import streamlit.components.v1 as components
import json
import threading
from datetime import datetime, timedelta
try:
    import pyotp
except:
    pass
import requests
import pandas as pd
import os, re, io, calendar, time, base64
import plotly.express as px
from datetime import date, timedelta
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


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


# ==============================================================================
# FUNÇÕES DE BUSCA — via Booking.com Demand API (rápido e confiável)
# ==============================================================================

BOOKING_API_KEY = "fd3e7724-043d-4b2b-bd73-6bc05b012b02"
BOOKING_AFFILIATE_ID = "2871199"
BOOKING_API_BASE = "https://demandapi.booking.com/3.1"

def _booking_api_headers():
    """Retorna headers de autenticação para a Demand API."""
    return {
        "Authorization": f"Bearer {BOOKING_API_KEY}",
        "Content-Type": "application/json",
        "X-Affiliate-Id": BOOKING_AFFILIATE_ID,
    }

@st.cache_data(ttl=86400)
def _buscar_hotel_id_por_nome(nome_hotel):
    """
    Busca o accommodation ID do Booking.com pelo nome do hotel.
    Usa cache de 24h para não repetir buscas.
    """
    try:
        resp = requests.post(
            f"{BOOKING_API_BASE}/accommodations/search",
            headers=_booking_api_headers(),
            json={
                "booker": {"country": "br", "platform": "desktop"},
                "checkin": (date.today() + timedelta(days=30)).isoformat(),
                "checkout": (date.today() + timedelta(days=31)).isoformat(),
                "guests": {"number_of_adults": 2, "number_of_rooms": 1},
                "extras": ["products"],
                "name": nome_hotel,
                "rows": 5,
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        results = data.get("data", [])
        if results:
            # Retorna o primeiro resultado (melhor match)
            return results[0].get("id")
    except Exception as e:
        st.warning(f"⚠️ Erro ao buscar ID de '{nome_hotel}': {e}")
    return None


def _buscar_disponibilidade_api(accommodation_id, d_in, d_out):
    """
    Busca disponibilidade e preço via Demand API.
    Retorna dict com preço e info do quarto, ou None.
    """
    try:
        resp = requests.post(
            f"{BOOKING_API_BASE}/accommodations/availability",
            headers=_booking_api_headers(),
            json={
                "accommodation": accommodation_id,
                "booker": {"country": "br", "platform": "desktop"},
                "checkin": d_in.isoformat(),
                "checkout": d_out.isoformat(),
                "guests": {"number_of_adults": 2, "number_of_rooms": 1},
                "extras": ["extra_charges"],
                "currency": "BRL",
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json().get("data", {})
        
        products = data.get("products", [])
        if not products:
            return None
        
        # Pega a recomendação (melhor preço) ou o primeiro produto
        recommendation = data.get("recommendation", {})
        
        melhor_preco = None
        melhor_nome = ""
        
        for product in products:
            preco_info = product.get("price", {})
            preco = preco_info.get("book", preco_info.get("total"))
            nome_quarto = product.get("room_name", product.get("name", ""))
            
            if preco is not None:
                preco_val = float(preco)
                if melhor_preco is None or preco_val < melhor_preco:
                    melhor_preco = preco_val
                    melhor_nome = nome_quarto
        
        if melhor_preco:
            return {"preco": int(melhor_preco), "quarto": melhor_nome}
        return None
        
    except Exception:
        return None


def _buscar_quartos_api(accommodation_id, d_in, d_out, n=5):
    """
    Busca todos os quartos e preços de um hotel via API.
    Retorna lista de dicts: [{"nome": str, "preco": int|str}]
    """
    try:
        resp = requests.post(
            f"{BOOKING_API_BASE}/accommodations/availability",
            headers=_booking_api_headers(),
            json={
                "accommodation": accommodation_id,
                "booker": {"country": "br", "platform": "desktop"},
                "checkin": d_in.isoformat(),
                "checkout": d_out.isoformat(),
                "guests": {"number_of_adults": 2, "number_of_rooms": 1},
                "extras": ["extra_charges"],
                "currency": "BRL",
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json().get("data", {})
        products = data.get("products", [])
        
        quartos = []
        nomes_vistos = set()
        
        for product in products:
            nome = product.get("room_name", product.get("name", ""))
            if not nome or nome.lower() in nomes_vistos:
                continue
            
            preco_info = product.get("price", {})
            preco = preco_info.get("book", preco_info.get("total"))
            
            if preco is not None:
                quartos.append({"nome": nome, "preco": int(float(preco))})
            else:
                quartos.append({"nome": nome, "preco": "Esgotado"})
            
            nomes_vistos.add(nome.lower())
            if len(quartos) >= n:
                break
        
        return quartos[:n]
    except:
        return []


# --- FUNÇÃO: Executa varredura por nome (API) ---
def executar_varredura(datas_para_busca, concorrentes, url_base_busca=None):
    todas_buscas = []
    status = st.empty()
    prog = st.progress(0)
    total = len(datas_para_busca) * len(concorrentes)
    cont = 0

    # Primeiro, busca os IDs de todos os hotéis
    status.markdown("🔍 Buscando IDs dos hotéis na API do Booking...")
    hotel_ids = {}
    for hotel in concorrentes:
        hid = _buscar_hotel_id_por_nome(hotel)
        hotel_ids[hotel] = hid
        if hid:
            status.markdown(f"✅ {hotel} → ID: {hid}")
        else:
            status.markdown(f"⚠️ {hotel} → Não encontrado")
        time.sleep(0.3)  # rate limit

    for d_in, d_out in datas_para_busca:
        for hotel in concorrentes:
            status.markdown(f"📡 Pesquisando: **{d_in.strftime('%d/%m')}** | {hotel}")
            hid = hotel_ids.get(hotel)
            preco_final = "Esgotado"
            
            if hid:
                resultado = _buscar_disponibilidade_api(hid, d_in, d_out)
                if resultado:
                    preco_final = resultado["preco"]
            
            todas_buscas.append({"Data": d_in.strftime("%d/%m/%Y"), "Hotel": hotel, "Preço": preco_final})
            cont += 1
            prog.progress(cont / total)
            time.sleep(0.2)  # rate limit

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


# --- FUNÇÃO: Executa varredura por URL direta / slug (API) ---
def executar_varredura_por_url(datas_para_busca, slugs):
    """Mantém compatibilidade — busca por nome via API."""
    todas_buscas = []
    status = st.empty()
    prog = st.progress(0)
    total = len(datas_para_busca) * len(slugs)
    cont = 0

    status.markdown("🔍 Buscando IDs dos hotéis na API do Booking...")
    hotel_ids = {}
    for nome in slugs.keys():
        hid = _buscar_hotel_id_por_nome(nome)
        hotel_ids[nome] = hid
        time.sleep(0.3)

    for d_in, d_out in datas_para_busca:
        for nome, slug in slugs.items():
            status.markdown(f"📡 Pesquisando: **{d_in.strftime('%d/%m')}** | {nome}")
            hid = hotel_ids.get(nome)
            preco_final = "Esgotado"
            
            if hid:
                resultado = _buscar_disponibilidade_api(hid, d_in, d_out)
                if resultado:
                    preco_final = resultado["preco"]
            
            todas_buscas.append({"Data": d_in.strftime("%d/%m/%Y"), "Hotel": nome, "Preço": preco_final})
            cont += 1
            prog.progress(cont / total)
            time.sleep(0.2)

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
# SHOPPER ALTO DA BOA VISTA — funções especiais (quartos + nome) via API
# ==============================================================================

def executar_varredura_alto(datas_para_busca, concorrentes, sua_col, slug_alto):
    status = st.empty()
    prog   = st.progress(0)
    outros = [c for c in concorrentes if c != sua_col]
    total  = len(datas_para_busca) * (len(outros) + 1)
    cont   = 0
    todas  = []

    # Busca IDs de todos os hotéis
    status.markdown("🔍 Buscando IDs dos hotéis na API do Booking...")
    hotel_ids = {}
    for hotel in outros:
        hid = _buscar_hotel_id_por_nome(hotel)
        hotel_ids[hotel] = hid
        time.sleep(0.3)
    
    alto_id = _buscar_hotel_id_por_nome(sua_col)
    time.sleep(0.3)

    for d_in, d_out in datas_para_busca:
        label = f"{d_in.strftime('%d/%m/%Y')} → {d_out.strftime('%d/%m/%Y')}"
        linha = {"Data": label}

        for hotel in outros:
            status.markdown(f"📡 **{label}** | {hotel[:35]}...")
            hid = hotel_ids.get(hotel)
            preco = "Esgotado"
            nome_conc = ""
            
            if hid:
                resultado = _buscar_disponibilidade_api(hid, d_in, d_out)
                if resultado:
                    preco = resultado["preco"]
                    nome_conc = resultado.get("quarto", "")
            
            linha[hotel] = preco
            linha[f"{hotel}__quarto"] = nome_conc
            cont += 1
            prog.progress(cont / total)
            time.sleep(0.2)

        # Busca quartos do Alto da Boa Vista
        status.markdown(f"🏨 **{label}** | Alto da Boa Vista — buscando quartos...")
        quartos = []
        if alto_id:
            quartos = _buscar_quartos_api(alto_id, d_in, d_out, n=5)
        
        for i, q in enumerate(quartos):
            linha[f"Alto_Q{i+1}_nome"]  = q["nome"]
            linha[f"Alto_Q{i+1}_preco"] = q["preco"]
        for i in range(len(quartos), 5):
            linha[f"Alto_Q{i+1}_nome"]  = "Esgotado"
            linha[f"Alto_Q{i+1}_preco"] = "Esgotado"
        cont += 1
        prog.progress(cont / total)
        todas.append(linha)

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
