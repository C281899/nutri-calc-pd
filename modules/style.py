"""
style.py — Design system for NutriCalc P&D
New brand identity: deep-blue technical palette with amber P&D accent.
"""
import streamlit as st

_TOKENS = """
:root {
    --bg-app:         #0A0F1A;
    --bg-panel:       #0D1520;
    --bg-card:        #111B28;
    --bg-input:       #141E2A;
    --border:         #1A2A40;
    --border-focus:   #378ADD;
    --border-input:   #1E2E42;
    --text-primary:   #E8F0F8;
    --text-secondary: #85B7EB;
    --text-muted:     #4A6080;
    --accent:         #378ADD;
    --accent-dark:    #185FA5;
    --accent-dim:     #0A1828;
    --brand-dark:     #042C53;
    --brand-mid:      #0C447C;
    --amber:          #EF9F27;
    --amber-bg:       #5A3A00;
    --success:        #2ECC71;
    --success-bg:     #0A2010;
    --error:          #F85149;
    --error-bg:       #1A0808;
    --radius-sm:      4px;
    --radius-md:      8px;
    --radius-lg:      12px;
}
"""

_GLOBAL_CSS = """
/* ── Base ─────────────────────────────────────────────────────────────────── */
html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--bg-app) !important;
}
.main .block-container {
    padding: 1.25rem 1.75rem 3rem !important;
    max-width: 1280px;
}

/* ── Top header bar ───────────────────────────────────────────────────────── */
[data-testid="stHeader"] {
    background-color: var(--bg-panel) !important;
    border-bottom: 0.5px solid var(--border) !important;
}

/* ── Headings ─────────────────────────────────────────────────────────────── */
h1, h2, h3, h4, h5, h6 {
    color: var(--text-primary) !important;
    font-weight: 600 !important;
    letter-spacing: -0.02em;
}

/* ── Sidebar ──────────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background-color: var(--bg-panel) !important;
    border-right: 0.5px solid var(--border) !important;
}
[data-testid="stSidebarContent"] {
    padding: 0 !important;
}

/* ── Tabs ─────────────────────────────────────────────────────────────────── */
[data-baseweb="tab-list"] {
    background-color: transparent !important;
    border-bottom: 0.5px solid var(--border) !important;
    gap: 0 !important;
}
[data-baseweb="tab"] {
    background-color: transparent !important;
    color: var(--text-muted) !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 10px 18px !important;
    border-bottom: 2px solid transparent !important;
    transition: color 0.15s !important;
}
[aria-selected="true"][data-baseweb="tab"] {
    color: var(--text-primary) !important;
    border-bottom: 2px solid var(--accent) !important;
    background-color: transparent !important;
}
[data-baseweb="tab-highlight"] {
    background-color: var(--accent) !important;
    height: 2px !important;
}
[data-baseweb="tab-border"] {
    background-color: var(--border) !important;
}

/* ── Inputs ───────────────────────────────────────────────────────────────── */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stTextArea"] textarea {
    background-color: var(--bg-input) !important;
    border: 0.5px solid var(--border-input) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text-secondary) !important;
    font-size: 12px !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
    border-color: var(--border-focus) !important;
    box-shadow: 0 0 0 2px rgba(55,138,221,0.15) !important;
}
[data-baseweb="select"] > div {
    background-color: var(--bg-input) !important;
    border-color: var(--border-input) !important;
}

/* ── Labels ───────────────────────────────────────────────────────────────── */
[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stTextArea"] label,
.stRadio label {
    font-size: 11px !important;
    font-weight: 500 !important;
    color: var(--text-secondary) !important;
    letter-spacing: 0.3px !important;
}

/* ── Buttons ──────────────────────────────────────────────────────────────── */
.stButton > button {
    background-color: var(--accent-dark) !important;
    border: none !important;
    color: #ffffff !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    border-radius: var(--radius-sm) !important;
    padding: 8px 16px !important;
    transition: background-color 0.15s !important;
    white-space: nowrap !important;
    min-width: fit-content !important;
}
.stButton > button:hover {
    background-color: var(--brand-mid) !important;
}
.stButton > button[kind="primary"],
.stButton > button[data-testid*="primary"] {
    background-color: var(--accent) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"]:hover {
    background-color: var(--accent-dark) !important;
}

/* ── Download button ──────────────────────────────────────────────────────── */
.stDownloadButton > button {
    background-color: var(--accent) !important;
    border: none !important;
    color: #ffffff !important;
    font-weight: 700 !important;
}

/* ── Dataframe ────────────────────────────────────────────────────────────── */
.stDataFrame {
    background-color: var(--bg-panel) !important;
    border: 0.5px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
}

/* ── Expander ─────────────────────────────────────────────────────────────── */
[data-testid="stExpander"],
details[data-testid="stExpander"] {
    background-color: var(--bg-card) !important;
    border: 0.5px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
}
details[data-testid="stExpander"] summary {
    color: var(--text-secondary) !important;
    font-size: 12px !important;
    font-weight: 500 !important;
}

/* ── Alerts ───────────────────────────────────────────────────────────────── */
[data-testid="stAlert"] {
    border-radius: var(--radius-md) !important;
}

/* ── Divider ──────────────────────────────────────────────────────────────── */
hr {
    border-color: var(--border) !important;
    margin: 1rem 0 !important;
}

/* ── Caption ──────────────────────────────────────────────────────────────── */
/* Legendas trazem informação de trabalho (peso total, porção, avisos):
   precisam de contraste legível sobre o fundo escuro. */
small, .stCaption, [data-testid="stCaptionContainer"] * {
    color: #7FA3CC !important;
    font-size: 11px !important;
}

/* ── File uploader ────────────────────────────────────────────────────────── */
section[data-testid="stFileUploadDropzone"],
[data-testid="stFileUploader"] {
    background-color: var(--bg-input) !important;
    border: 1px dashed var(--border-input) !important;
    border-radius: var(--radius-md) !important;
}

/* ── Radio ────────────────────────────────────────────────────────────────── */
.stRadio [data-testid="stMarkdownContainer"] p {
    font-size: 12px !important;
    color: var(--text-secondary) !important;
}

/* ── Scrollbar ────────────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg-app); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--accent); }

/* ── Chip / badge classes ─────────────────────────────────────────────────── */
.chip-green {
    background: #0A2010; color: #2ECC71;
    border: 0.5px solid #1A5030;
    padding: 3px 10px; border-radius: 12px;
    font-size: 10px; font-weight: 700; letter-spacing: 0.5px;
    display: inline-block;
}
.chip-blue {
    background: #0A1828; color: #378ADD;
    border: 0.5px solid #185FA5;
    padding: 3px 10px; border-radius: 12px;
    font-size: 10px; font-weight: 700;
    display: inline-block;
}
.chip-gray {
    background: #141E2A; color: #85B7EB;
    border: 0.5px solid #1A2A40;
    padding: 3px 10px; border-radius: 12px;
    font-size: 10px; font-weight: 700;
    display: inline-block;
}
.chip-amber {
    background: #5A3A00; color: #EF9F27;
    border: 0.5px solid #EF9F27;
    padding: 3px 10px; border-radius: 12px;
    font-size: 10px; font-weight: 700; letter-spacing: 1px;
    display: inline-block;
}
"""

_SVG_MOLECULE = (
    '<svg viewBox="0 0 80 80" width="{s}" height="{s}" xmlns="http://www.w3.org/2000/svg">'
    '<rect width="80" height="80" rx="{rx}" fill="#0C447C"/>'
    '<ellipse cx="40" cy="40" rx="22" ry="8" fill="none" stroke="#85B7EB" stroke-width="{sw}"/>'
    '<ellipse cx="40" cy="40" rx="22" ry="8" fill="none" stroke="#85B7EB" stroke-width="{sw}" transform="rotate(60 40 40)"/>'
    '<ellipse cx="40" cy="40" rx="22" ry="8" fill="none" stroke="#85B7EB" stroke-width="{sw}" transform="rotate(120 40 40)"/>'
    '<circle cx="40" cy="40" r="7" fill="#378ADD"/>'
    '<circle cx="62" cy="40" r="3.5" fill="#B5D4F4"/>'
    '<circle cx="18" cy="40" r="3.5" fill="#B5D4F4"/>'
    '<circle cx="29" cy="21" r="3.5" fill="#B5D4F4"/>'
    '<circle cx="51" cy="21" r="3.5" fill="#B5D4F4"/>'
    '<circle cx="29" cy="59" r="3.5" fill="#B5D4F4"/>'
    '<circle cx="51" cy="59" r="3.5" fill="#B5D4F4"/>'
    '</svg>'
)


def _svg(size: int, rx: int = 16, sw: float = 2) -> str:
    return _SVG_MOLECULE.format(s=size, rx=rx, sw=sw)


def inject_global_css():
    """Injects the NutriCalc design system into the Streamlit page."""
    st.markdown(f"<style>{_TOKENS}{_GLOBAL_CSS}</style>", unsafe_allow_html=True)


def render_page_header(db_count: int = 0):
    """Renders the branded application header with molecule logo and status chips."""
    count_chip = f'<span class="chip-gray">{db_count:,} ALIMENTOS</span>' if db_count else ""
    st.markdown(f"""
    <div style="
        background:#0D1520;
        border:0.5px solid #1A2A40;
        border-radius:12px;
        padding:16px 24px;
        display:flex;
        align-items:center;
        justify-content:space-between;
        flex-wrap:wrap;
        gap:10px;
        margin-bottom:16px;
    ">
      <div style="display:flex;align-items:center;gap:14px;">
        {_svg(44)}
        <div>
          <div style="font-size:22px;font-weight:900;color:#fff;line-height:1.1;
                      font-family:system-ui,-apple-system,sans-serif;">
            NUTRI<span style="font-weight:300;color:#85B7EB;">Calc</span
            ><span style="color:#EF9F27;font-size:14px;font-weight:700;
                          letter-spacing:3px;margin-left:6px;">P&amp;D</span>
          </div>
          <div style="font-size:10px;color:#4A6080;margin-top:3px;letter-spacing:0.5px;">
            Tabelas Nutricionais &middot; Padrão ANVISA RDC 429/2020 &middot; Uso Interno P&amp;D
          </div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <span class="chip-green">&#9679; SISTEMA ATIVO</span>
        <span class="chip-blue">TACO 4ª ED.</span>
        {count_chip}
      </div>
    </div>
    """, unsafe_allow_html=True)


def render_section_label(text: str):
    """Renders an uppercase section label with bottom border."""
    st.markdown(f"""
    <div style="
        font-size:10px;
        font-weight:700;
        letter-spacing:2.5px;
        color:#378ADD;
        text-transform:uppercase;
        margin:16px 0 8px;
        padding-bottom:8px;
        border-bottom:0.5px solid #1A2A40;
    ">{text}</div>
    """, unsafe_allow_html=True)


def render_card(content_html: str, accent: bool = False):
    """Wraps HTML content in a styled dark card."""
    border = "border-left:3px solid #378ADD;" if accent else ""
    st.markdown(
        f'<div style="background:#111B28;border:0.5px solid #1A2A40;'
        f'border-radius:8px;padding:1rem 1.25rem;margin-bottom:0.75rem;{border}">'
        f'{content_html}</div>',
        unsafe_allow_html=True,
    )


def render_divider():
    """Renders a styled horizontal rule."""
    st.markdown('<hr>', unsafe_allow_html=True)


def render_legal_notice():
    """Renders the ANVISA legal disclaimer card."""
    st.markdown("""
    <div style="
        background:#1A1200;
        border:0.5px solid #5A3A00;
        border-left:3px solid #EF9F27;
        border-radius:8px;
        padding:0.65rem 1rem;
        color:#EF9F27;
        font-size:0.8rem;
        line-height:1.6;
        margin-top:0.75rem;
    ">
        <strong>Aviso Legal:</strong> Este cálculo é uma estimativa para uso interno em P&amp;D.
        <strong>Obrigatório:</strong> validação por nutricionista habilitado (CRN) antes do uso
        em rótulo comercial para venda. Valores podem diferir de análise laboratorial.
    </div>
    """, unsafe_allow_html=True)


def render_fonte_badge(fonte: str) -> str:
    """Returns an inline HTML chip for TACO or FORNECEDOR source."""
    if fonte == "TACO":
        return '<span class="chip-blue">TACO</span>'
    return '<span class="chip-green">Fornecedor</span>'


def render_sidebar(db_count: int = 0):
    """Renders the fully branded sidebar content."""
    st.markdown(f"""
    <div style="padding:20px 16px 16px;border-bottom:0.5px solid #1A2A40;">
      <div style="display:flex;align-items:center;gap:10px;">
        {_svg(36, rx=14, sw=2.5)}
        <div>
          <div style="font-size:16px;font-weight:900;color:#fff;line-height:1.1;
                      font-family:system-ui,-apple-system,sans-serif;">
            NUTRI<span style="font-weight:300;color:#85B7EB;">CALC</span>
          </div>
          <div style="display:flex;align-items:center;gap:6px;margin-top:2px;">
            <div style="width:14px;height:1px;background:#378ADD;"></div>
            <span style="font-size:9px;letter-spacing:3px;color:#EF9F27;font-weight:700;">P&amp;D</span>
            <div style="width:14px;height:1px;background:#378ADD;"></div>
          </div>
        </div>
      </div>
      <div style="font-size:10px;color:#4A6080;margin-top:6px;">v2.2 &middot; Uso Interno</div>
    </div>

    <div style="padding:14px 16px 8px;">
      <div style="font-size:10px;font-weight:700;letter-spacing:2.5px;color:#378ADD;margin-bottom:10px;">
        STATUS DO SISTEMA
      </div>
      <div style="font-size:12px;color:#85B7EB;margin-bottom:5px;">&#9679; Banco ativo: {db_count:,} alimentos</div>
      <div style="font-size:12px;color:#85B7EB;margin-bottom:5px;">&#9635; Padrão: ANVISA RDC 429/2020</div>
      <div style="font-size:12px;color:#85B7EB;">&#11041; Fontes: TACO 4ª Ed. + TBCA/USDA</div>
    </div>
    <div style="height:0.5px;background:#1A2A40;margin:0 16px;"></div>

    <div style="padding:12px 16px;">
      <div style="font-size:10px;font-weight:700;letter-spacing:2.5px;color:#378ADD;margin-bottom:8px;">
        AVISO LEGAL
      </div>
      <div style="font-size:11px;color:#4A6080;line-height:1.6;">
        Estimativa para uso interno em P&amp;D.<br>
        Obrigatório: validação por nutricionista<br>
        habilitado (CRN) antes de uso em rótulo comercial.
      </div>
    </div>
    """, unsafe_allow_html=True)
