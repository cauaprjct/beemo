"""Streamlit interface for Beemo — Office Agent.

This module provides a web interface for users to interact with the Agent
through natural language commands to manipulate Office files.
"""

import base64
import streamlit as st
from pathlib import Path
from typing import List

from src.factory import create_agent
from src.models import AgentResponse
from src.logging_config import get_logger

logger = get_logger(__name__)


def initialize_session_state():
    """Initialize Streamlit session state variables."""
    if "agent" not in st.session_state:
        st.session_state.agent = None
    if "history" not in st.session_state:
        st.session_state.history = []
    if "discovered_files" not in st.session_state:
        st.session_state.discovered_files = []
    if "selected_files" not in st.session_state:
        st.session_state.selected_files = []
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [
            {
                "role": "assistant",
                "content": (
                    "Olá! Sou o **Beemo** 🎮\n\n"
                    "Converse naturalmente — faço tanto **perguntas** quanto **ações**:\n"
                    "- *\"O que tem nesse arquivo?\"* → analiso e respondo\n"
                    "- *\"Converta o gráfico para pizza\"* → executo na hora\n"
                    "- *\"Ordene os dados pela coluna Data\"* → modifica o arquivo\n\n"
                    "Selecione os arquivos no painel esquerdo e comece!"
                ),
                "is_chat": True,
                "response": None,
            }
        ]


def initialize_agent():
    """Initialize the Agent using factory."""
    if st.session_state.agent is None:
        try:
            with st.spinner("Inicializando Agent..."):
                st.session_state.agent = create_agent()
                logger.info("Agent inicializado com sucesso via Streamlit")
        except Exception as e:
            st.error(f"Erro ao inicializar Agent: {e}")
            logger.error(f"Erro ao inicializar Agent: {e}", exc_info=True)
            st.stop()


def _strip_json_wrapper_ui(text: str) -> str:
    """Strip JSON wrappers like {"response": "..."} from model replies.

    Safety net in the UI layer to ensure chat messages display as plain text.
    """
    if not text:
        return text
    stripped = text.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        try:
            import json
            data = json.loads(stripped)
            if isinstance(data, dict):
                for key in ("response", "message", "text", "reply", "answer", "content"):
                    if key in data and isinstance(data[key], str):
                        return data[key]
                if len(data) == 1:
                    val = next(iter(data.values()))
                    if isinstance(val, str):
                        return val
        except (ValueError, TypeError):
            pass
    return text


def discover_files() -> List[str]:
    """Discover available Office files using the Agent's file scanner."""
    try:
        files = st.session_state.agent.file_scanner.scan_office_files()
        logger.info(f"Arquivos descobertos: {len(files)}")
        return files
    except Exception as e:
        st.error(f"Erro ao descobrir arquivos: {e}")
        logger.error(f"Erro ao descobrir arquivos: {e}", exc_info=True)
        return []


def display_file_selector(files: List[str]):
    """Display file selection interface with checkboxes.

    Args:
        files: List of discovered file paths
    """
    if not files:
        st.info(
            "Nenhum arquivo Office encontrado.\n\n"
            "Use o botão **📤 Enviar arquivos** acima para adicionar "
            "arquivos, ou peça ao Beemo para criar um novo!"
        )
        return

    st.subheader("📁 Arquivos Disponíveis")
    st.caption(f"{len(files)} arquivo(s) encontrado(s)")

    # Mostrar botões apenas se houver múltiplos arquivos
    if len(files) > 1:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Selecionar Todos", use_container_width=True):
                st.session_state.selected_files = files.copy()
                for _f in files:
                    st.session_state[f"file_{_f}"] = True
                st.rerun()
        with col2:
            if st.button("❌ Limpar Seleção", use_container_width=True):
                st.session_state.selected_files = []
                for _f in files:
                    st.session_state[f"file_{_f}"] = False
                st.rerun()

    _EXT_ICONS = {'.docx': '📝', '.xlsx': '📊', '.pptx': '📑', '.pdf': '📕', '.doc': '📝', '.xls': '📊', '.ppt': '📑'}
    for file_path in files:
        file_name = Path(file_path).name
        ext = Path(file_path).suffix.lower()
        icon = _EXT_ICONS.get(ext, '📄')
        is_selected = file_path in st.session_state.selected_files
        checked = st.checkbox(
            f"{icon} {file_name}",
            value=is_selected,
            key=f"file_{file_path}",
            help=str(Path(file_path).parent),
        )
        if checked and file_path not in st.session_state.selected_files:
            st.session_state.selected_files.append(file_path)
        elif not checked and file_path in st.session_state.selected_files:
            st.session_state.selected_files.remove(file_path)


def main():
    """Main entry point for the Streamlit application."""
    from PIL import Image as _PILImage
    _beemo_path = Path(__file__).parent / "beemo.png"
    _page_icon = _PILImage.open(_beemo_path) if _beemo_path.exists() else "🎮"
    st.set_page_config(
        page_title="Beemo",
        page_icon=_page_icon,
        layout="wide",
    )

    initialize_session_state()
    initialize_agent()

    # ── CSS base: aplica a ambos os temas ──────────────────────────────────────
    st.markdown("""
    <style>
    /* Ocultar chrome do Streamlit: menu, footer, âncoras nos headings */
    #MainMenu { visibility: hidden !important; }
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    .stDeployButton { display: none !important; }
    [data-testid="StyledLinkIconContainer"] { display: none !important; }

    /* Reduzir padding do topo da área principal */
    .main .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1.5rem !important;
    }

    /* Botões em formato pill com transições suaves */
    .stButton > button {
        border-radius: 20px !important;
        font-weight: 500 !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
        border: none !important;
        padding: 0.35rem 1.1rem !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
    }
    .stButton > button:active { transform: translateY(0) !important; }
    .stButton > button:focus,
    .stButton > button:focus-visible,
    .stButton > button:active {
        outline: none !important;
        box-shadow: none !important;
        border: none !important;
    }
    /* Kill focus ring on the container wrapper too */
    .stButton:focus-within,
    .stButton:focus,
    .stButton > div:focus-within {
        outline: none !important;
        box-shadow: none !important;
    }

    /* Inputs arredondados */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea {
        border-radius: 10px !important;
    }

    /* Selectbox arredondado */
    .stSelectbox [data-baseweb="select"] > div:first-child {
        border-radius: 10px !important;
    }

    /* Chat input pill */
    [data-testid="stChatInput"] { border-radius: 24px !important; overflow: hidden !important; }

    /* Quick undo/redo bar */
    .undo-redo-bar {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 8px 12px;
        margin-bottom: 4px;
        border-radius: 12px;
        background: rgba(31, 119, 180, 0.08);
        border: 1px solid rgba(31, 119, 180, 0.2);
    }
    .undo-redo-bar .stButton > button {
        font-size: 18px !important;
        min-height: 36px !important;
        padding: 4px 10px !important;
        border-radius: 10px !important;
    }
    .undo-redo-bar .stButton > button:disabled {
        background-color: #d0d0d0 !important;
        color: #888 !important;
        border: 1px solid #bbb !important;
    }

    /* Chat messages arredondadas */
    [data-testid="stChatMessage"] {
        border-radius: 14px !important;
        margin-bottom: 0.4rem !important;
    }

    /* Expanders e alerts arredondados */
    [data-testid="stExpander"] { border-radius: 10px !important; overflow: hidden !important; }
    [data-testid="stAlert"]    { border-radius: 10px !important; }

    /* Sidebar: section headers menores e discretos */
    [data-testid="stSidebar"] h2 {
        font-size: 0.7rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.1em !important;
        opacity: 0.5 !important;
        margin-top: 0.6rem !important;
        margin-bottom: 0.1rem !important;
    }
    [data-testid="stSidebar"] h3 {
        font-size: 0.65rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.08em !important;
        opacity: 0.45 !important;
        margin-top: 0.4rem !important;
        margin-bottom: 0.1rem !important;
    }

    /* Model badge (estrutura — cor definida por tema) */
    .model-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.72rem;
        font-weight: 600;
        border: 1px solid;
        vertical-align: middle;
        margin-left: 6px;
    }

    /* Header compacto */
    .app-header { margin-bottom: 0.6rem; }
    .app-title-row {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
        margin-bottom: 2px;
    }
    .app-title {
        font-size: 1.9rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }
    .app-subtitle {
        font-size: 0.82rem;
        margin: 0;
        opacity: 0.6;
    }
    </style>
    """, unsafe_allow_html=True)

    # Aplicar tema escuro via CSS se selecionado
    # (kept on next line to avoid merge conflict below)
    if "theme" in st.session_state and st.session_state.theme == "dark":
        st.markdown("""
        <style>
        /* Tema Escuro */
        .stApp {
            background-color: #0e1117;
            color: #fafafa;
        }
        
        /* Sidebar - nuclear option: force dark on everything */
        [data-testid="stSidebar"],
        [data-testid="stSidebar"] *,
        [data-testid="stSidebar"] > div,
        [data-testid="stSidebar"] > div > div,
        [data-testid="stSidebar"] > div > div > div,
        [data-testid="stSidebar"] section,
        [data-testid="stSidebar"] section > div,
        [data-testid="stSidebarContent"],
        [data-testid="stSidebarUserContent"],
        [data-testid="stSidebarNav"],
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div,
        section[data-testid="stSidebar"] > div > div {
            background-color: #262730 !important;
            background: #262730 !important;
        }
        
        /* Sidebar border/resize handle */
        [data-testid="stSidebar"]::before,
        [data-testid="stSidebar"]::after,
        [data-testid="stSidebar"] > div::before,
        [data-testid="stSidebar"] > div::after,
        [data-testid="stSidebar"] > div > div::before,
        [data-testid="stSidebar"] > div > div::after {
            background-color: #262730 !important;
            background: #262730 !important;
            border-color: #262730 !important;
        }
        
        /* Text color in sidebar */
        [data-testid="stSidebar"] * {
            color: #fafafa !important;
        }
        
        /* Override any white backgrounds with attribute selectors */
        [data-testid="stSidebar"] [style*="background"],
        [data-testid="stSidebar"] [class*="css-"] {
            background-color: #262730 !important;
            background: #262730 !important;
        }
        
        /* Force iframe and embed backgrounds */
        iframe, embed {
            background-color: #0e1117 !important;
        }
        
        /* Target specific Streamlit emotion cache classes for sidebar */
        .st-emotion-cache-1gwvy71,
        .st-emotion-cache-6qob1r,
        .st-emotion-cache-1cypcdb,
        .st-emotion-cache-16txtl3,
        .st-emotion-cache-vk3wp9,
        .st-emotion-cache-1rtdyuf,
        .st-emotion-cache-1wmy9hl,
        .st-emotion-cache-ffhzg2,
        .st-emotion-cache-1avcm0n,
        .st-emotion-cache-18ni7ap,
        .st-emotion-cache-1dp5vir,
        .st-emotion-cache-z5fcl4,
        .st-emotion-cache-1v0mbdj,
        .st-emotion-cache-uf99v8,
        .st-emotion-cache-r421ms {
            background-color: #262730 !important;
            background: #262730 !important;
        }
        
        /* Texto principal */
        p, span, div, label {
            color: #fafafa !important;
        }
        
        /* Headers */
        h1, h2, h3, h4, h5, h6 {
            color: #fafafa !important;
        }
        
        /* Inputs e textareas */
        .stTextInput > div > div > input,
        .stTextArea > div > div > textarea {
            background-color: #262730 !important;
            color: #fafafa !important;
            border-color: #4a4a4a !important;
        }
        
        .stTextInput label,
        .stTextArea label {
            color: #fafafa !important;
        }
        
        /* Placeholder text */
        .stTextInput input::placeholder,
        .stTextArea textarea::placeholder {
            color: #a0a0a0 !important;
        }
        
        /* Botões */
        .stButton > button {
            background-color: #1f77b4 !important;
            color: white !important;
            border: none !important;
            outline: none !important;
        }
        
        .stButton > button:hover {
            background-color: #1557a0 !important;
            border: none !important;
        }
        
        .stButton > button:focus,
        .stButton > button:focus-visible,
        .stButton > button:active {
            outline: none !important;
            outline-offset: 0 !important;
            box-shadow: none !important;
            border: none !important;
        }
        .stButton:focus-within {
            outline: none !important;
            box-shadow: none !important;
        }
        
        /* Botões desabilitados */
        .stButton > button:disabled {
            background-color: #4a4a4a !important;
            color: #808080 !important;
        }
        
        /* Expanders */
        .streamlit-expanderHeader {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        details summary {
            color: #fafafa !important;
        }
        
        /* Métricas */
        [data-testid="stMetricValue"] {
            color: #fafafa !important;
        }
        
        [data-testid="stMetricLabel"] {
            color: #fafafa !important;
        }
        
        /* Alertas */
        .stAlert {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        /* Checkboxes */
        .stCheckbox label {
            color: #fafafa !important;
        }
        
        /* Selectbox */
        .stSelectbox > div > div {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        .stSelectbox label {
            color: #fafafa !important;
        }
        
        .stSelectbox [data-baseweb="select"] {
            background-color: #262730 !important;
        }
        
        .stSelectbox [data-baseweb="select"] > div {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        /* Dropdown menu */
        [role="listbox"] {
            background-color: #262730 !important;
        }
        
        [role="option"] {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        [role="option"]:hover {
            background-color: #1f77b4 !important;
        }
        
        /* Dividers */
        hr {
            border-color: #4a4a4a !important;
        }
        
        /* Captions */
        .stCaption {
            color: #a0a0a0 !important;
        }
        
        /* Markdown */
        .stMarkdown {
            color: #fafafa !important;
        }
        
        /* Code blocks */
        code {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        pre {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        /* Links */
        a {
            color: #4da6ff !important;
        }
        
        a:hover {
            color: #80bfff !important;
        }
        
        /* Spinner */
        .stSpinner > div {
            border-top-color: #1f77b4 !important;
        }
        
        /* Success/Error/Warning/Info messages */
        .stSuccess {
            background-color: #1a4d2e !important;
            color: #90ee90 !important;
        }
        
        .stError {
            background-color: #4d1a1a !important;
            color: #ff6b6b !important;
        }
        
        .stWarning {
            background-color: #4d3d1a !important;
            color: #ffd93d !important;
        }
        
        .stInfo {
            background-color: #1a3a4d !important;
            color: #6bb6ff !important;
        }
        
        /* Chat messages */
        [data-testid="stChatMessage"] {
            background-color: #1e1e2e !important;
            border: 1px solid #3a3a4a !important;
        }
        
        [data-testid="stChatMessageContent"] {
            color: #fafafa !important;
        }
        
        [data-testid="stChatMessageContent"] p,
        [data-testid="stChatMessageContent"] span,
        [data-testid="stChatMessageContent"] div {
            color: #fafafa !important;
        }
        
        /* Chat input */
        [data-testid="stChatInput"] {
            background-color: #262730 !important;
            border-color: #4a4a4a !important;
        }
        
        [data-testid="stChatInput"] textarea,
        [data-testid="stChatInputTextArea"] {
            background-color: #262730 !important;
            color: #fafafa !important;
            border-color: #4a4a4a !important;
        }
        
        [data-testid="stChatInput"] textarea::placeholder {
            color: #a0a0a0 !important;
        }
        
        /* Bottom container (chat input area) */
        [data-testid="stBottom"] {
            background-color: #0e1117 !important;
        }
        
        [data-testid="stBottomBlockContainer"] {
            background-color: #0e1117 !important;
        }
        
        /* Scrollbars - global */
        * {
            scrollbar-width: thin;
            scrollbar-color: #4a4a4a #1e1e2e;
        }
        
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
            background-color: #1e1e2e;
        }
        
        ::-webkit-scrollbar-track {
            background: #1e1e2e;
        }
        
        ::-webkit-scrollbar-thumb {
            background: #4a4a4a;
            border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: #5a5a5a;
        }
        
        /* Sidebar scrollbar - hide default white scrollbar */
        [data-testid="stSidebar"],
        [data-testid="stSidebar"] > div,
        [data-testid="stSidebar"] > div > div,
        [data-testid="stSidebarContent"],
        [data-testid="stSidebarUserContent"] {
            scrollbar-width: thin !important;
            scrollbar-color: #4a4a4a #262730 !important;
        }
        
        [data-testid="stSidebar"] ::-webkit-scrollbar,
        [data-testid="stSidebar"] > div ::-webkit-scrollbar {
            width: 6px !important;
            background: #262730 !important;
        }
        
        [data-testid="stSidebar"] ::-webkit-scrollbar-track,
        [data-testid="stSidebar"] > div ::-webkit-scrollbar-track {
            background: #262730 !important;
        }
        
        [data-testid="stSidebar"] ::-webkit-scrollbar-thumb,
        [data-testid="stSidebar"] > div ::-webkit-scrollbar-thumb {
            background: #4a4a4a !important;
            border-radius: 3px !important;
        }
        
        /* Sidebar resize handle / border - make it darker */
        [data-testid="stSidebar"]::after,
        [data-testid="stSidebar"] > div::after {
            background-color: #3a3a4a !important;
            width: 1px !important;
        }
        
        [data-testid="stSidebarUserContent"] {
            background-color: #262730 !important;
        }
        
        /* Sidebar resize drag handle */
        [data-testid="stSidebarResizer"],
        [data-testid="collapsedControl"] {
            background-color: #262730 !important;
        }
        
        /* Sidebar gutter - the white bar between sidebar and main */
        .st-emotion-cache-1gwvy71,
        .st-emotion-cache-6qob1r,
        .st-emotion-cache-1cypcdb,
        .st-emotion-cache-eczf16,
        .st-emotion-cache-16txtl3,
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="baseButton-header"] {
            background-color: #262730 !important;
            background: #262730 !important;
        }
        
        /* Force all section backgrounds */
        section[data-testid="stSidebar"] {
            background: #262730 !important;
        }
        
        section[data-testid="stSidebar"] > div {
            background: #262730 !important;
        }
        
        /* The inner scroll container */
        .css-1d391kg, .css-1lcbmhc, .css-12oz5g7 {
            background-color: #262730 !important;
        }
        
        /* Generic sidebar div backgrounds */
        [data-testid="stSidebar"] div[class*="css-"] {
            background-color: transparent !important;
        }
        
        [data-testid="stSidebar"] > div:first-child {
            background-color: #262730 !important;
        }
        
        /* Tooltips */
        [data-baseweb="tooltip"] {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        /* Multiselect */
        [data-baseweb="tag"] {
            background-color: #1f77b4 !important;
            color: #fafafa !important;
        }
        
        /* Tables / Dataframes */
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {
            background-color: #1e1e2e !important;
        }
        
        [data-testid="stDataFrame"] th,
        [data-testid="stTable"] th {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        [data-testid="stDataFrame"] td,
        [data-testid="stTable"] td {
            background-color: #1e1e2e !important;
            color: #fafafa !important;
        }
        
        /* Form submit button */
        [data-testid="stFormSubmitButton"] button {
            background-color: #1f77b4 !important;
            color: white !important;
        }
        
        /* Number input */
        [data-testid="stNumberInput"] input {
            background-color: #262730 !important;
            color: #fafafa !important;
            border-color: #4a4a4a !important;
        }
        
        /* Radio buttons */
        [data-testid="stRadio"] label {
            color: #fafafa !important;
        }
        
        /* File uploader */
        [data-testid="stFileUploader"] {
            background-color: #262730 !important;
            border-color: #4a4a4a !important;
        }
        
        [data-testid="stFileUploader"] label {
            color: #fafafa !important;
        }
        
        /* JSON viewer */
        .react-json-view {
            background-color: #262730 !important;
        }
        
        /* Block container backgrounds */
        [data-testid="stVerticalBlock"] {
            background-color: transparent !important;
        }
        
        /* Main container */
        .main .block-container {
            background-color: #0e1117 !important;
        }
        
        /* Sidebar specific fixes */
        [data-testid="stSidebar"] [data-testid="stCheckbox"] label span {
            color: #fafafa !important;
        }
        
        [data-testid="stSidebar"] .stCheckbox label p {
            color: #fafafa !important;
        }
        
        /* Info boxes in sidebar */
        [data-testid="stSidebar"] [data-testid="stAlert"],
        [data-testid="stSidebar"] .stAlert {
            background-color: #1a3a4d !important;
            color: #fafafa !important;
        }
        
        [data-testid="stSidebar"] [data-testid="stAlert"] p,
        [data-testid="stSidebar"] .stAlert p {
            color: #fafafa !important;
        }
        
        /* Expander headers and content in sidebar */
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            background-color: #262730 !important;
            border-color: #4a4a4a !important;
        }
        
        [data-testid="stSidebar"] [data-testid="stExpander"] summary,
        [data-testid="stSidebar"] [data-testid="stExpander"] summary span {
            color: #fafafa !important;
        }
        
        [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stText"],
        [data-testid="stSidebar"] .stText {
            color: #fafafa !important;
        }
        
        /* Caption text */
        [data-testid="stSidebar"] [data-testid="stCaption"],
        [data-testid="stSidebar"] .stCaption,
        [data-testid="stSidebar"] small {
            color: #a0a0a0 !important;
        }
        
        /* All text elements in sidebar */
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label {
            color: #fafafa !important;
        }
        
        /* Headers in sidebar */
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #fafafa !important;
        }
        
        /* Buttons in sidebar */
        [data-testid="stSidebar"] .stButton button {
            background-color: #1f77b4 !important;
            color: white !important;
            border: none !important;
        }
        
        [data-testid="stSidebar"] .stButton button:disabled {
            background-color: #3a3a4a !important;
            color: #808080 !important;
        }
        
        /* st.text() elements */
        [data-testid="stText"] {
            color: #fafafa !important;
        }

        /* Model badge — tema escuro */
        .model-badge {
            background: rgba(100, 160, 230, 0.15) !important;
            color: #90c0f0 !important;
            border-color: rgba(100, 160, 230, 0.3) !important;
        }

        /* App header — cores do tema escuro */
        .app-title { color: #fafafa !important; }
        .app-subtitle { color: #c0c0c0 !important; }
        </style>
        """, unsafe_allow_html=True)

    # Aplicar tema claro via CSS (necessário para sobrescrever dark mode do OS/browser)
    if "theme" not in st.session_state or st.session_state.theme == "light":
        st.markdown("""
        <style>
        /* Tema Claro — força cores claras independente do dark mode do SO */
        .stApp {
            background-color: #ffffff !important;
            color: #31333f !important;
        }
        .main .block-container {
            background-color: #ffffff !important;
        }
        /* Sidebar */
        [data-testid="stSidebar"],
        [data-testid="stSidebar"] *,
        [data-testid="stSidebar"] > div,
        [data-testid="stSidebar"] > div > div,
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div,
        [data-testid="stSidebarContent"],
        [data-testid="stSidebarUserContent"] {
            background-color: #f0f2f6 !important;
            background: #f0f2f6 !important;
        }
        [data-testid="stSidebar"] * {
            color: #31333f !important;
        }
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #31333f !important;
        }
        /* Text */
        p, span, div, label { color: #31333f !important; }
        h1, h2, h3, h4, h5, h6 { color: #31333f !important; }
        [data-testid="stText"] { color: #31333f !important; }
        .stMarkdown { color: #31333f !important; }
        /* Inputs */
        .stTextInput > div > div > input,
        .stTextArea > div > div > textarea {
            background-color: #ffffff !important;
            color: #31333f !important;
            border-color: #cccccc !important;
        }
        /* Buttons */
        .stButton > button {
            background-color: #1f77b4 !important;
            color: white !important;
        }
        .stButton > button:disabled {
            background-color: #e0e0e0 !important;
            color: #999999 !important;
        }
        /* Selectbox */
        .stSelectbox > div > div,
        .stSelectbox [data-baseweb="select"],
        .stSelectbox [data-baseweb="select"] > div {
            background-color: #ffffff !important;
            color: #31333f !important;
        }
        [role="listbox"] { background-color: #ffffff !important; }
        [role="option"] {
            background-color: #ffffff !important;
            color: #31333f !important;
        }
        [role="option"]:hover { background-color: #e8f0fe !important; }
        /* Expanders */
        .streamlit-expanderHeader { background-color: #f0f2f6 !important; color: #31333f !important; }
        details summary { color: #31333f !important; }
        /* Chat messages */
        [data-testid="stChatMessage"] {
            background-color: #f8f9fa !important;
            border: 1px solid #e0e0e0 !important;
        }
        [data-testid="stChatMessageContent"],
        [data-testid="stChatMessageContent"] p,
        [data-testid="stChatMessageContent"] span,
        [data-testid="stChatMessageContent"] div { color: #31333f !important; }
        /* Chat input — todos os wrappers possíveis */
        [data-testid="stChatInput"],
        [data-testid="stChatInput"] *,
        [data-testid="stChatInput"] > div,
        [data-testid="stChatInput"] > div > div,
        [data-testid="stChatInputContainer"],
        [data-testid="stChatInputContainer"] *,
        [data-testid="stChatInputTextArea"],
        .stChatInput,
        .stChatInput * {
            background-color: #ffffff !important;
            background: #ffffff !important;
            color: #31333f !important;
            border-color: #cccccc !important;
        }
        [data-testid="stChatInput"] textarea,
        [data-testid="stChatInputTextArea"],
        textarea[data-testid="stChatInputTextArea"] {
            background-color: #ffffff !important;
            color: #31333f !important;
            border-color: #cccccc !important;
        }
        [data-testid="stChatInput"] textarea::placeholder {
            color: #999999 !important;
        }
        /* Bottom bar (área fixa no rodapé com o input) */
        [data-testid="stBottom"],
        [data-testid="stBottom"] *,
        [data-testid="stBottom"] > div,
        [data-testid="stBottomBlockContainer"],
        [data-testid="stBottomBlockContainer"] *,
        .stBottom,
        .stBottom * {
            background-color: #ffffff !important;
            background: #ffffff !important;
        }
        /* Alerts */
        .stAlert { background-color: #f0f2f6 !important; color: #31333f !important; }
        /* Code */
        code { background-color: #f0f2f6 !important; color: #31333f !important; }
        pre  { background-color: #f0f2f6 !important; color: #31333f !important; }
        /* Links */
        a { color: #1f77b4 !important; }
        a:hover { color: #155d8a !important; }
        /* Dividers */
        hr { border-color: #e0e0e0 !important; }
        /* Captions */
        .stCaption { color: #6c757d !important; }
        /* Metrics */
        [data-testid="stMetricValue"],
        [data-testid="stMetricLabel"] { color: #31333f !important; }
        /* Scrollbars */
        * { scrollbar-width: thin; scrollbar-color: #cccccc #f0f2f6; }
        ::-webkit-scrollbar { width: 8px; height: 8px; background-color: #f0f2f6; }
        ::-webkit-scrollbar-track { background: #f0f2f6; }
        ::-webkit-scrollbar-thumb { background: #cccccc; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #aaaaaa; }

        /* Model badge — tema claro */
        .model-badge {
            background: rgba(31, 119, 180, 0.1) !important;
            color: #1a6fa0 !important;
            border-color: rgba(31, 119, 180, 0.28) !important;
        }

        /* App header — cores do tema claro */
        .app-title { color: #1a1a2e !important; }
        .app-subtitle { color: #555 !important; }
        </style>
        """, unsafe_allow_html=True)

    # Header compacto com badge do modelo
    model_badge_html = ""
    if st.session_state.agent:
        client = st.session_state.agent.gemini_client
        model_badge_html = f'<span class="model-badge">⚡ {client.active_model_name}</span>'

    _beemo_path = Path(__file__).parent / "beemo.png"
    _beemo_img_html = ""
    if _beemo_path.exists():
        with open(_beemo_path, "rb") as _f:
            _b64 = base64.b64encode(_f.read()).decode()
        _beemo_img_html = (
            f'<img src="data:image/png;base64,{_b64}" '
            f'height="56" style="vertical-align:middle; margin-right:10px; '
            f'border-radius:8px;">'
        )

    st.markdown(f"""
    <div class="app-header">
        <div class="app-title-row">
            <h1 class="app-title">{_beemo_img_html}Beemo</h1>
            {model_badge_html}
        </div>
        <p class="app-subtitle">Seu assistente Office com comandos em linguagem natural</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar – file discovery and selection
    with st.sidebar:
        # Seletor de tema no topo
        st.markdown("### 🎨 Tema")
        
        # Inicializar tema no session_state
        if "theme" not in st.session_state:
            st.session_state.theme = "light"
        
        # Toggle de tema
        col1, col2 = st.columns([3, 1])
        with col1:
            theme_option = st.selectbox(
                "Selecione o tema:",
                options=["light", "dark"],
                format_func=lambda x: "☀️ Claro" if x == "light" else "🌙 Escuro",
                index=0 if st.session_state.theme == "light" else 1,
                key="theme_selector",
                label_visibility="collapsed"
            )
        
        # Atualizar tema se mudou
        if theme_option != st.session_state.theme:
            st.session_state.theme = theme_option
            st.rerun()
        
        st.divider()
        
        st.header("⚙️ Arquivos")

        # ── File uploader — allows user to add files to the working directory ──
        _root = Path(st.session_state.agent.config.root_path) if st.session_state.agent else None
        if _root:
            _root.mkdir(parents=True, exist_ok=True)  # auto-create ROOT_PATH
        uploaded_files = st.file_uploader(
            "📤 Enviar arquivos",
            type=["xlsx", "docx", "pptx", "pdf"],
            accept_multiple_files=True,
            help="Envie arquivos Office para o diretório de trabalho do Beemo",
            key="file_uploader",
        )
        if uploaded_files and _root:
            _any_new = False
            for uf in uploaded_files:
                dest = _root / uf.name
                if not dest.exists():
                    dest.write_bytes(uf.getvalue())
                    logger.info(f"Arquivo enviado: {dest}")
                    _any_new = True
            if _any_new:
                st.toast("📤 Arquivo(s) enviado(s)!", icon="✅")
                st.session_state.discovered_files = discover_files()
                st.rerun()

        if st.button("🔄 Atualizar Lista", help="Atualiza a lista de arquivos Office disponíveis"):
            with st.spinner("Descobrindo arquivos..."):
                st.session_state.discovered_files = discover_files()
                st.session_state.selected_files = []

        if not st.session_state.discovered_files:
            st.session_state.discovered_files = discover_files()

        display_file_selector(st.session_state.discovered_files)

        if _root:
            st.caption(f"📂 `{_root}`")
        
        # Seção de versionamento
        st.divider()
        st.header("⏱️ Histórico de Versões")
        
        if st.session_state.agent and st.session_state.agent.version_manager:
            version_manager = st.session_state.agent.version_manager
            files_with_history = version_manager.get_all_files_with_history()
            
            if files_with_history:
                st.caption(f"{len(files_with_history)} arquivo(s) com histórico")
                
                for rel_path in files_with_history:
                    full_path = str(version_manager.root_path / rel_path)
                    file_name = Path(rel_path).name
                    history = version_manager.get_history(full_path)
                    
                    if history:
                        with st.expander(f"📄 {file_name}"):
                            st.caption(f"{len(history.versions)} versão(ões)")
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                can_undo = version_manager.can_undo(full_path)
                                if st.button(
                                    "↩️ Desfazer",
                                    key=f"undo_{rel_path}",
                                    disabled=not can_undo,
                                    use_container_width=True,
                                    help="Desfaz a última operação realizada neste arquivo"
                                ):
                                    response = st.session_state.agent.undo_file(full_path)
                                    if response.success:
                                        st.success(response.message)
                                        st.rerun()
                                    else:
                                        st.error(response.message)
                            
                            with col2:
                                can_redo = version_manager.can_redo(full_path)
                                if st.button(
                                    "↪️ Refazer",
                                    key=f"redo_{rel_path}",
                                    disabled=not can_redo,
                                    use_container_width=True,
                                    help="Refaz a última operação desfeita neste arquivo"
                                ):
                                    response = st.session_state.agent.redo_file(full_path)
                                    if response.success:
                                        st.success(response.message)
                                        st.rerun()
                                    else:
                                        st.error(response.message)
                            
                            # Mostrar histórico de versões
                            st.caption("Histórico:")
                            for i, version in enumerate(reversed(history.versions)):
                                is_current = (len(history.versions) - 1 - i) == history.current_index
                                marker = "→ " if is_current else "  "
                                timestamp = version.timestamp[:19].replace('T', ' ')
                                
                                # Adicionar contexto e ícone da operação
                                operation_icon = ""
                                operation_text = version.operation
                                
                                if version.operation == 'add_chart':
                                    operation_icon = "📊"
                                    operation_text = "Adicionou gráfico"
                                elif version.operation == 'delete_chart':
                                    operation_icon = "🗑️"
                                    operation_text = "Removeu gráfico"
                                elif version.operation == 'update_chart':
                                    operation_icon = "📊"
                                    operation_text = "Alterou tipo de gráfico"
                                elif version.operation == 'resize_chart':
                                    operation_icon = "📐"
                                    operation_text = "Redimensionou gráfico"
                                elif version.operation == 'move_chart':
                                    operation_icon = "↔️"
                                    operation_text = "Moveu gráfico"
                                elif version.operation == 'sort':
                                    operation_icon = "🔄"
                                    operation_text = "Ordenou dados"
                                elif version.operation == 'update':
                                    operation_icon = "✏️"
                                    operation_text = "Atualizou células"
                                elif version.operation == 'format':
                                    operation_icon = "🎨"
                                    operation_text = "Formatou células"
                                elif version.operation == 'formula':
                                    operation_icon = "🔢"
                                    operation_text = "Adicionou fórmulas"
                                elif version.operation == 'append':
                                    operation_icon = "➕"
                                    operation_text = "Adicionou linhas"
                                elif version.operation == 'create':
                                    operation_icon = "📄"
                                    operation_text = "Criou arquivo"
                                elif version.operation == 'delete_sheet':
                                    operation_icon = "🗑️"
                                    operation_text = "Removeu planilha"
                                elif version.operation == 'delete_rows':
                                    operation_icon = "➖"
                                    operation_text = "Removeu linhas"
                                elif version.operation == 'merge':
                                    operation_icon = "🔗"
                                    operation_text = "Mesclou células"
                                else:
                                    operation_icon = "📝"
                                    operation_text = version.operation
                                
                                st.text(f"{marker}{operation_icon} {timestamp} - {operation_text}")
            else:
                st.info("Nenhum arquivo com histórico de versões ainda")
        else:
            st.info("Versionamento não disponível")
        
        # Seção de cache
        st.divider()
        st.header("⚡ Cache de Respostas")
        
        if st.session_state.agent and st.session_state.agent.response_cache:
            cache = st.session_state.agent.response_cache
            
            if cache.enabled:
                # Estatísticas
                stats = cache.get_stats()
                
                # Resumo simples
                st.info(f"💾 {stats.total_entries} respostas salvas | Taxa de acerto: {stats.hit_rate:.1%}")
                
                # Modo avançado
                show_advanced = st.checkbox("Mostrar estatísticas avançadas", value=False)
                
                if show_advanced:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Entradas", stats.total_entries)
                        st.metric("Hits", stats.total_hits)
                    with col2:
                        st.metric("Taxa de Hit", f"{stats.hit_rate:.1%}")
                        st.metric("Tamanho", f"{stats.cache_size_bytes / 1024:.1f} KB")
                    
                    # Entradas recentes
                    recent = cache.get_recent_entries(limit=5)
                    if recent:
                        with st.expander("📋 Entradas Recentes"):
                            for entry in recent:
                                timestamp = entry['timestamp'][:19].replace('T', ' ')
                                
                                st.text(f"🕐 {timestamp}")
                                st.caption(f"Prompt: {entry['prompt_preview']}...")
                                st.caption(f"Hits: {entry['hit_count']}")
                                st.divider()
                
                # Controles
                if st.button("🗑️ Limpar Cache", use_container_width=True, help="Remove todas as respostas em cache"):
                    cache.invalidate()
                    st.success("Cache limpo com sucesso!")
                    st.rerun()
            else:
                st.info("Cache desabilitado")
                st.caption("Configure CACHE_ENABLED=true no .env para habilitar")
        else:
            st.info("Cache não disponível")
        
        # Limpar conversa
        if st.session_state.chat_messages:
            st.divider()
            if st.button(
                "🗑️ Limpar Conversa",
                use_container_width=True,
                help="Apaga todas as mensagens da conversa atual"
            ):
                st.session_state.chat_messages = [
                    {
                        "role": "assistant",
                        "content": (
                            "Olá! Sou o **Beemo** 🎮\n\n"
                            "Converse naturalmente — faço tanto **perguntas** quanto **ações**:\n"
                            "- *\"O que tem nesse arquivo?\"* → analiso e respondo\n"
                            "- *\"Converta o gráfico para pizza\"* → executo na hora\n"
                            "- *\"Ordene os dados pela coluna Data\"* → modifica o arquivo\n\n"
                            "Selecione os arquivos no painel esquerdo e comece!"
                        ),
                        "is_chat": True,
                        "response": None,
                    }
                ]
                st.rerun()

        # Seção de ajuda
        st.divider()

        with st.expander("❓ Ajuda"):
            st.markdown("""
            ### Como usar:
            1. Selecione um ou mais arquivos no painel
            2. Digite na caixa de chat no fundo da tela
            3. A resposta chega automaticamente — sem botão de enviar

            ### Perguntas (Chat Mode):
            - "O que tem nesse arquivo?"
            - "Quantas linhas há na planilha?"
            - "Explique os dados"
            - "Tem algum gráfico?"

            ### Comandos (Action Mode — executa na hora):
            - "Converta o gráfico para pizza"
            - "Ordene os dados pela coluna Data"
            - "Adicione um gráfico de barras na posição H2"
            - "Formate o cabeçalho com negrito e fundo azul"
            - "Delete a linha 5"

            ### Funcionalidades:
            - ✅ Excel: criar, atualizar, formatar, gráficos, ordenar
            - ✅ Word: criar, formatar, tabelas, listas
            - ✅ PowerPoint: slides, layouts, tabelas
            - ✅ PDF: criar, extrair, mesclar, dividir

            ### Dicas:
            - Use ↩️/↪️ no histórico de versões para desfazer/refazer
            - Feche o Excel antes de modificar arquivos (Windows)
            """)

    # Main area — chat interface
    selected = st.session_state.selected_files if st.session_state.selected_files else None
    if selected:
        st.caption(f"📎 {len(selected)} arquivo(s) selecionado(s) nesta conversa")

    # Display chat history
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            if msg["role"] == "user":
                st.write(msg["content"])
            elif msg.get("is_chat"):
                content = msg["content"].strip()
                # Detect accidental JSON leak (action that wasn't recognized/executed)
                if content.startswith("{") and '"actions"' in content:
                    st.warning(
                        "⚠️ Recebi um comando JSON mas não consegui executá-lo. "
                        "Tente reformular o pedido."
                    )
                    with st.expander("Detalhes técnicos"):
                        st.code(content, language="json")
                else:
                    # Safety net: strip JSON wrappers like {"response":"..."} that
                    # slipped through the agent's _strip_json_wrapper
                    content = _strip_json_wrapper_ui(content)
                    st.markdown(content)
            else:
                response = msg.get("response")
                if response and response.success:
                    st.success(f"✅ {response.message}")
                    if response.files_modified:
                        unique_files = list(dict.fromkeys(response.files_modified))
                        st.write("**Arquivos modificados:**")
                        for f in unique_files:
                            st.write(f"- {Path(f).name}")
                    if response.batch_result:
                        batch = response.batch_result
                        for res in batch.action_results:
                            icon = "✅" if res.success else "❌"
                            with st.expander(
                                f"{icon} Ação {res.action_index + 1}: {res.operation} — {Path(res.target_file).name}",
                                expanded=not res.success
                            ):
                                if not res.success and res.error:
                                    st.error(res.error)
                elif response and not response.success:
                    st.error(f"❌ {response.message}")
                    if response.error:
                        with st.expander("Detalhes do erro"):
                            st.code(response.error)
                else:
                    st.write(msg["content"])

    # ── Quick Undo/Redo bar (acima do input) ────────────────────────────────
    if st.session_state.agent and st.session_state.agent.version_manager:
        _vm = st.session_state.agent.version_manager
        _last_modified: list = []
        for _msg in reversed(st.session_state.chat_messages):
            _r = _msg.get("response")
            if _r and not getattr(_r, 'is_chat', True) and getattr(_r, 'files_modified', None):
                _last_modified = list(dict.fromkeys(_r.files_modified))
                break

        _bars = [(f, _vm.can_undo(f), _vm.can_redo(f)) for f in _last_modified]
        _bars = [(f, u, r) for f, u, r in _bars if u or r]

        if _bars:
            st.markdown('<div class="undo-redo-bar">', unsafe_allow_html=True)
            for _fp, _can_u, _can_r in _bars:
                _fname = Path(_fp).name
                _cu, _cr, _clabel, _ = st.columns([0.18, 0.18, 1.6, 4])
                with _cu:
                    if st.button(
                        "↩", key=f"qu_{_fp}",
                        disabled=not _can_u,
                        help=f"Desfazer última ação em {_fname}",
                        use_container_width=True
                    ):
                        _res = st.session_state.agent.undo_file(_fp)
                        if _res.success:
                            st.toast(f"↩ Desfeito em {_fname}")
                        else:
                            st.toast(f"Não foi possível desfazer", icon="❌")
                        st.rerun()
                with _cr:
                    if st.button(
                        "↪", key=f"qr_{_fp}",
                        disabled=not _can_r,
                        help=f"Refazer última ação desfeita em {_fname}",
                        use_container_width=True
                    ):
                        _res = st.session_state.agent.redo_file(_fp)
                        if _res.success:
                            st.toast(f"↪ Refeito em {_fname}")
                        else:
                            st.toast(f"Não foi possível refazer", icon="❌")
                        st.rerun()
                with _clabel:
                    st.caption(f"_{_fname}_")
            st.markdown('</div>', unsafe_allow_html=True)

    # Chat input (sticky at the bottom)
    if user_input := st.chat_input(
        "Pergunte algo, dê um comando ou converse naturalmente...",
        key="chat_input"
    ):
        # Append user message immediately
        st.session_state.chat_messages.append({
            "role": "user",
            "content": user_input,
            "is_chat": True,
            "response": None
        })

        # Build history for context (all previous turns)
        chat_history = [
            {"role": m["role"], "content": m["content"]}
            for m in st.session_state.chat_messages[:-1]
        ]

        # Process via chat agent
        with st.spinner("🤖 Pensando..."):
            response: AgentResponse = st.session_state.agent.process_chat_message(
                user_input,
                chat_history=chat_history,
                selected_files=selected
            )

        # Append assistant response
        st.session_state.chat_messages.append({
            "role": "assistant",
            "content": response.message,
            "is_chat": response.is_chat,
            "response": response
        })

        # Keep legacy history for version manager sidebar
        if not response.is_chat:
            st.session_state.history.append({
                "prompt": user_input,
                "success": response.success,
                "message": response.message,
                "files_modified": response.files_modified,
                "error": response.error,
                "is_batch": response.batch_result is not None,
                "batch_summary": (
                    f"{response.batch_result.successful_actions}/{response.batch_result.total_actions}"
                    if response.batch_result else None
                )
            })

        st.rerun()

if __name__ == "__main__":
    main()
