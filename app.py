import streamlit as st
import streamlit.components.v1 as components
import os
import io
import re
import html
import base64
import datetime
from PIL import Image
from docxtpl import DocxTemplate
import docx

# -----------------------------------------------------------------------------
# Asset Loader & Base64 Caching
# -----------------------------------------------------------------------------
TEMPLATE_FILENAME = "abstract_template.docx"
BG_IMAGE_FILENAME = "background.jpg"
LOGO_FILENAME = "logo_marag.png"
LOGO_THUMB_FILENAME = "logo_thumb.png"
TEMPLATE_SEAL_FILENAME = "template_seal.png"

@st.cache_data
def get_base64_of_file(path: str) -> str:
    """Read a local binary file and return its base64 encoded string."""
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""

logo_img = Image.open(LOGO_FILENAME) if os.path.exists(LOGO_FILENAME) else None
bg_b64 = get_base64_of_file(BG_IMAGE_FILENAME)
logo_b64 = get_base64_of_file(LOGO_FILENAME)

# Create optimized 160x160 logo thumbnail if it doesn't exist to ensure fast rendering
if not os.path.exists(LOGO_THUMB_FILENAME) and os.path.exists(LOGO_FILENAME):
    try:
        im = Image.open(LOGO_FILENAME)
        im.thumbnail((160, 160), Image.Resampling.LANCZOS)
        im.save(LOGO_THUMB_FILENAME, format="PNG", optimize=True)
    except Exception:
        pass

thumb_b64 = get_base64_of_file(LOGO_THUMB_FILENAME) or logo_b64
seal_b64 = get_base64_of_file(TEMPLATE_SEAL_FILENAME) or thumb_b64

# -----------------------------------------------------------------------------
# Page Configuration (Enforced Dark Mode)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Abstract of Canvass Generator",
    page_icon=logo_img,
    layout="wide",
    initial_sidebar_state="expanded"
)

DEFAULT_SAMPLE_DATA = {
    "ac_number": "AC –0433H-10-26",
    "opened_date": "October 26, 2026",
    "purpose": "FOR: PAYMENT FOR THE PURCHASED OF TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD IN LGU- MARAGONDON, CAVITE",
    "bidder_1_name": "RBK PRINTING SERVICES",
    "bidder_1_location": "Bucal 1, Maragondon, Cavite",
    "bidder_1_amount": "₱ 1,000.00",
    "bidder_2_name": "CAVITE PRINT EXPRESS",
    "bidder_2_location": "Naic, Cavite",
    "bidder_2_amount": "₱ 1,250.00",
    "bidder_3_name": "MARAGONDON GRAPHICS & PRINTS",
    "bidder_3_location": "Poblacion 1A, Maragondon, Cavite",
    "bidder_3_amount": "₱ 1,500.00",
    "auto_clean_purpose": True,
}

# -----------------------------------------------------------------------------
# Enforce Dark Theme & Compact Responsive CSS
# -----------------------------------------------------------------------------
css_rules = f"""
<style>
    /* Enforce Dark Scheme Everywhere & Remove Light Mode */
    :root {{
        color-scheme: dark !important;
    }}
    
    /* Hide Theme Settings from menu */
    div[data-testid="stToolbar"] button[aria-label="Settings"],
    div[data-testid="stToolbar"] ul li:has(button[aria-label*="theme"]),
    div[data-testid="stToolbar"] ul li:has(button[aria-label*="Theme"]) {{
        display: none !important;
    }}

    /* Font: Arial for web UI, strictly excluding Material Icons */
    html, body, p, 
    h1, h2, h3, h4, h5, h6, 
    input, textarea, select, label, table, td, th {{
        font-family: Arial, Helvetica, sans-serif !important;
    }}
    
    button:not([data-testid="stSidebarCollapseButton"]) {{
        font-family: Arial, Helvetica, sans-serif !important;
    }}

    /* Strictly preserve Material Symbols for sidebar collapse toggle and icons */
    [data-testid*="Icon"], 
    [class*="material-symbols"], 
    [data-testid="stIconMaterial"],
    [data-testid="stSidebarCollapseButton"] span,
    [data-testid="stSidebarCollapseButton"] i,
    span.material-symbols-rounded,
    span.material-symbols-outlined {{
        font-family: "Material Symbols Rounded", "Material Symbols Outlined", sans-serif !important;
        font-style: normal !important;
        letter-spacing: normal !important;
        text-transform: none !important;
        white-space: nowrap !important;
        word-wrap: normal !important;
        direction: ltr !important;
    }}

    /* Background Image: Dimmed & Blurred over full app */
    .stApp {{
        background: linear-gradient(rgba(11, 19, 43, 0.75), rgba(11, 19, 43, 0.75)), 
                    url("data:image/jpeg;base64,{bg_b64}") no-repeat center center fixed !important;
        background-size: cover !important;
    }}
    
    .stApp::before {{
        content: "";
        position: fixed;
        top: -20px;
        left: -20px;
        width: calc(100vw + 40px);
        height: calc(100vh + 40px);
        background-image: url("data:image/jpeg;base64,{bg_b64}");
        background-size: cover;
        background-position: center center;
        background-repeat: no-repeat;
        filter: blur(8px) brightness(0.38);
        -webkit-filter: blur(8px) brightness(0.38);
        z-index: 0;
        pointer-events: none;
    }}
    
    /* Layout: Strictly Center the Main Content Area on Sidebar Toggle */
    [data-testid="stAppViewContainer"] {{
        background: transparent !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}

    [data-testid="stMain"] {{
        background: transparent !important;
        position: relative;
        z-index: 1;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        width: 100% !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}

    [data-testid="stMainBlockContainer"],
    section.main,
    .main {{
        background: transparent !important;
        width: 100% !important;
        display: flex !important;
        justify-content: center !important;
    }}
    
    /* Center container smoothly with margin auto */
    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 1100px !important;
        width: 100% !important;
        margin-left: auto !important;
        margin-right: auto !important;
        transition: max-width 0.3s ease, margin 0.3s ease, padding 0.3s ease !important;
    }}
    
    [data-testid="stHeader"] {{
        background: transparent !important;
        height: 38px !important;
        z-index: 2;
    }}
    
    /* Green and Blue Gradient Sidebar: NO HORIZONTAL SCROLL */
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div,
    section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
    section[data-testid="stSidebar"] .block-container {{
        overflow-x: hidden !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }}
    
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #0b2545 0%, #0d4251 45%, #064e3b 100%) !important;
        border-right: 1px solid rgba(56, 189, 248, 0.25) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    section[data-testid="stSidebar"] * {{
        color: #f8fafc !important;
        box-sizing: border-box !important;
    }}
    section[data-testid="stSidebar"] .block-container {{
        padding-top: 1.5rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
    }}
    
    .sidebar-heading {{
        color: #6ee7b7 !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.8px !important;
        border-bottom: 1px solid rgba(110, 231, 183, 0.3) !important;
        padding-bottom: 6px !important;
        margin-bottom: 12px !important;
    }}
    
    /* Interactive History Card Styling */
    .history-item {{
        background: rgba(15, 23, 42, 0.6) !important;
        border: 1px solid rgba(56, 189, 248, 0.25) !important;
        border-left: 3.5px solid #10b981 !important;
        border-radius: 5px !important;
        padding: 9px 11px !important;
        margin-bottom: 6px !important;
        box-sizing: border-box !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    .history-item:hover {{
        background: rgba(15, 23, 42, 0.9) !important;
        border-color: #38bdf8 !important;
        border-left-color: #38bdf8 !important;
        transform: translateX(3px) !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3) !important;
    }}
    .history-ac {{
        font-size: 12.5px !important;
        font-weight: 700 !important;
        color: #38bdf8 !important;
        margin-bottom: 2px !important;
        word-break: break-all !important;
    }}
    .history-winner {{
        font-size: 11.5px !important;
        color: #f1f5f9 !important;
        line-height: 1.35 !important;
    }}
    .history-meta {{
        font-size: 10.5px !important;
        color: #94a3b8 !important;
        margin-top: 4px !important;
        display: flex;
        justify-content: space-between;
    }}
    
    /* Header Bar: Compact Blue-Green Gradient */
    .app-header-bar {{
        background: linear-gradient(135deg, #0f3460 0%, #1a5fb4 45%, #059669 100%);
        color: #ffffff;
        padding: 10px 16px;
        border-radius: 8px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.35);
        border: 1px solid rgba(255, 255, 255, 0.15);
        transition: box-shadow 0.25s ease, transform 0.25s ease;
    }}
    .app-header-bar:hover {{
        box-shadow: 0 4px 14px rgba(13, 148, 136, 0.3);
    }}
    .app-header-logo {{
        height: 44px;
        width: auto;
        object-fit: contain;
        flex-shrink: 0;
        transition: transform 0.25s ease;
    }}
    .app-header-bar:hover .app-header-logo {{
        transform: scale(1.05);
    }}
    .app-header-title {{
        margin: 0;
        font-size: 19px;
        font-weight: 700;
        color: #ffffff;
        letter-spacing: -0.3px;
        line-height: 1.2;
    }}
    .app-header-subtitle {{
        margin: 2px 0 0 0;
        font-size: 11px;
        color: #d1fae5;
        letter-spacing: 0.2px;
        font-weight: 400;
    }}
    
    /* Panel Title */
    .panel-title {{
        font-size: 13px !important;
        font-weight: 700 !important;
        color: #38bdf8 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        border-bottom: 1px solid rgba(56, 189, 248, 0.25) !important;
        padding-bottom: 4px !important;
        margin-bottom: 10px !important;
    }}
    
    /* High-Contrast Interactive Inputs */
    .stTextInput input, .stTextArea textarea {{
        background-color: rgba(30, 41, 59, 0.88) !important;
        color: #ffffff !important;
        border: 1px solid #475569 !important;
        font-size: 13px !important;
        padding: 4px 8px !important;
        min-height: 32px !important;
        height: 32px !important;
        border-radius: 5px !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease !important;
    }}
    .stTextArea textarea {{
        min-height: 60px !important;
        height: auto !important;
    }}
    .stTextInput input:hover, .stTextArea textarea:hover {{
        border-color: rgba(56, 189, 248, 0.6) !important;
        background-color: rgba(30, 41, 59, 0.96) !important;
    }}
    .stTextInput input:focus, .stTextArea textarea:focus {{
        border-color: #38bdf8 !important;
        background-color: rgba(15, 23, 42, 0.98) !important;
        box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25) !important;
    }}
    
    label p {{
        font-size: 12px !important;
        font-weight: 600 !important;
        color: #e2e8f0 !important;
        margin-bottom: 2px !important;
    }}
    
    /* Primary Gradient Button */
    button[kind="primary"] {{
        background: linear-gradient(135deg, #1d4ed8 0%, #0d9488 50%, #16a34a 100%) !important;
        border: none !important;
        color: #ffffff !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        padding: 7px 16px !important;
        border-radius: 6px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.25) !important;
        cursor: pointer !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    button[kind="primary"]:hover {{
        background: linear-gradient(135deg, #2563eb 0%, #0f766e 50%, #15803d 100%) !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35) !important;
    }}
    button[kind="primary"]:active {{
        transform: translateY(0.5px) !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.2) !important;
    }}
    
    /* Secondary Action Buttons */
    button[kind="secondary"] {{
        background: rgba(30, 41, 59, 0.75) !important;
        border: 1px solid #475569 !important;
        color: #f1f5f9 !important;
        font-size: 11.5px !important;
        padding: 4px 10px !important;
        border-radius: 5px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
    }}
    button[kind="secondary"]:hover {{
        background: rgba(51, 65, 85, 0.95) !important;
        border-color: #38bdf8 !important;
        color: #38bdf8 !important;
        transform: translateY(-1px) !important;
    }}
    button[kind="secondary"]:active {{
        transform: translateY(0.5px) !important;
    }}
    
    /* Clean Divider */
    hr {{
        margin: 6px 0 !important;
        border-color: rgba(148, 163, 184, 0.2) !important;
    }}
    
    /* Amount Display in 2-Column Mode */
    .amount-display {{
        font-size: 14px;
        font-weight: 700;
        color: #34d399;
        line-height: 32px;
        transition: color 0.2s ease;
    }}
    
    /* Validation Warning Boxes */
    .val-box-error {{
        background: rgba(153, 27, 27, 0.4) !important;
        border: 1px solid #ef4444 !important;
        border-left: 4px solid #f87171 !important;
        padding: 8px 12px !important;
        border-radius: 4px !important;
        font-size: 12.5px !important;
        color: #fecaca !important;
        margin: 8px 0 !important;
    }}
    .val-box-success {{
        background: rgba(6, 78, 59, 0.5) !important;
        border: 1px solid #10b981 !important;
        border-left: 4px solid #34d399 !important;
        padding: 6px 10px !important;
        border-radius: 4px !important;
        font-size: 12px !important;
        color: #d1fae5 !important;
        margin: 6px 0 !important;
    }}
    
    /* Award Clause Highlight */
    .award-summary-box {{
        background: rgba(6, 78, 59, 0.6) !important;
        border: 1px solid #10b981 !important;
        border-left: 4px solid #34d399 !important;
        padding: 8px 12px !important;
        border-radius: 4px !important;
        font-size: 12.5px !important;
        color: #ecfdf5 !important;
        margin-top: 6px !important;
        transition: border-color 0.25s ease, background 0.25s ease;
    }}
    
    /* -------------------------------------------------------------------------
       Realistic Printing Preview Sheet (Authentic US Legal 8.5 x 14 Page Style)
       ------------------------------------------------------------------------- */
    .print-preview-container {{
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        width: 100% !important;
        margin: 12px 0 24px 0 !important;
        overflow-x: auto !important;
        padding-bottom: 12px !important;
    }}
    
    .print-preview-topbar {{
        width: 100% !important;
        max-width: 800px !important;
        min-width: 680px !important;
        background: #1e293b !important;
        border: 1px solid #334155 !important;
        border-bottom: 2px solid #0d9488 !important;
        border-radius: 6px 6px 0 0 !important;
        padding: 8px 16px !important;
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        color: #cbd5e1 !important;
        font-size: 12px !important;
        box-sizing: border-box !important;
    }}
    .print-preview-topbar span {{
        font-weight: 600 !important;
        letter-spacing: 0.5px !important;
        color: #38bdf8 !important;
    }}
    
    .print-preview-sheet {{
        width: 100% !important;
        max-width: 800px !important;
        min-width: 680px !important;
        min-height: 1050px !important;
        background-color: #ffffff !important;
        background: #ffffff !important;
        color: #000000 !important;
        border: 1px solid #cbd5e1 !important;
        border-top: none !important;
        border-radius: 0 0 6px 6px !important;
        padding: 52px 64px !important;
        box-shadow: 0 14px 40px rgba(0, 0, 0, 0.65) !important;
        font-family: 'Times New Roman', Times, serif !important;
        box-sizing: border-box !important;
        line-height: 1.35 !important;
        display: block !important;
        visibility: visible !important;
    }}
    .print-preview-sheet,
    .print-preview-sheet * {{
        color: #000000 !important;
        font-family: 'Times New Roman', Times, serif !important;
    }}
</style>
"""
st.markdown(css_rules, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Client-Side Auto-Apply Script (Dispatches change event so user doesn't need to press Enter)
# -----------------------------------------------------------------------------
components.html("""
<script>
const pDoc = window.parent.document;
function initAutoCommit() {
    const fields = pDoc.querySelectorAll('input[type="text"], textarea');
    fields.forEach(f => {
        if (f._acDone) return;
        f._acDone = true;
        let debTimer = null;
        f.addEventListener('input', () => {
            clearTimeout(debTimer);
            debTimer = setTimeout(() => {
                if (pDoc.activeElement === f) {
                    const start = f.selectionStart;
                    const end = f.selectionEnd;
                    f.blur();
                    setTimeout(() => {
                        f.focus();
                        try { f.setSelectionRange(start, end); } catch(e){}
                    }, 35);
                }
            }, 650);
        });
    });
}
initAutoCommit();
setInterval(initAutoCommit, 700);
</script>
""", height=0, width=0)

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def parse_bid_amount(val_str):
    """Extract numeric bid value from string."""
    if val_str is None:
        return None
    val_clean = re.sub(r"[^\d.]", "", str(val_str))
    if not val_clean:
        return None
    try:
        return float(val_clean)
    except ValueError:
        return None

def format_bid_amount(val_str, currency_symbol="₱"):
    """Format string or number to standard currency format (₱ 1,000.00)."""
    val_str = str(val_str).strip() if val_str else ""
    if not val_str:
        return ""
    num = parse_bid_amount(val_str)
    if num is not None:
        return f"{currency_symbol} {num:,.2f}"
    return val_str

def clean_purpose_text(text: str) -> str:
    """Trim boilerplate to match template structure."""
    if not text:
        return ""
    cleaned = text.strip()
    cleaned = re.sub(r'^(?:FOR:\s*)?(?:PAYMENT\s+FOR\s+THE\s+PURCHASE[D]?\s+OF\s*)?', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'[\s,.-]*IN\s+LGU[- ]*MARAGONDON[,\s]*CAVITE\.?$', '', cleaned, flags=re.IGNORECASE)
    return cleaned.strip(" ,.-")

# -----------------------------------------------------------------------------
# Session State Initialization
# -----------------------------------------------------------------------------
if "form_data" not in st.session_state:
    st.session_state.form_data = DEFAULT_SAMPLE_DATA.copy()

if "generated_doc_bytes" not in st.session_state:
    st.session_state.generated_doc_bytes = None

if "generated_context" not in st.session_state:
    st.session_state.generated_context = None

if "abstract_history" not in st.session_state:
    st.session_state.abstract_history = [
        {
            "ac_number": "AC –0433H-10-26",
            "winning_bidder": "RBK PRINTING SERVICES",
            "winning_location": "Bucal 1, Maragondon, Cavite",
            "winning_amount": "₱ 1,000.00",
            "opened_date": "October 26, 2026",
            "purpose": "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD",
            "time_str": "Initial"
        }
    ]

def load_sample_data():
    st.session_state.form_data = DEFAULT_SAMPLE_DATA.copy()
    st.session_state["b1_name"] = DEFAULT_SAMPLE_DATA["bidder_1_name"]
    st.session_state["b1_loc"] = DEFAULT_SAMPLE_DATA["bidder_1_location"]
    st.session_state["b1_amt"] = DEFAULT_SAMPLE_DATA["bidder_1_amount"]
    st.session_state["b2_name"] = DEFAULT_SAMPLE_DATA["bidder_2_name"]
    st.session_state["b2_loc"] = DEFAULT_SAMPLE_DATA["bidder_2_location"]
    st.session_state["b2_amt"] = DEFAULT_SAMPLE_DATA["bidder_2_amount"]
    st.session_state["b3_name"] = DEFAULT_SAMPLE_DATA["bidder_3_name"]
    st.session_state["b3_loc"] = DEFAULT_SAMPLE_DATA["bidder_3_location"]
    st.session_state["b3_amt"] = DEFAULT_SAMPLE_DATA["bidder_3_amount"]
    st.session_state.generated_doc_bytes = None
    st.session_state.generated_context = None
    st.toast("Sample data loaded.")

def clear_form_data():
    st.session_state.form_data = {
        "ac_number": "",
        "opened_date": "October 26, 2026",
        "purpose": "",
        "bidder_1_name": "",
        "bidder_1_location": "",
        "bidder_1_amount": "",
        "bidder_2_name": "",
        "bidder_2_location": "",
        "bidder_2_amount": "",
        "bidder_3_name": "",
        "bidder_3_location": "",
        "bidder_3_amount": "",
        "auto_clean_purpose": True,
    }
    st.session_state["b1_name"] = ""
    st.session_state["b1_loc"] = ""
    st.session_state["b1_amt"] = ""
    st.session_state["b2_name"] = ""
    st.session_state["b2_loc"] = ""
    st.session_state["b2_amt"] = ""
    st.session_state["b3_name"] = ""
    st.session_state["b3_loc"] = ""
    st.session_state["b3_amt"] = ""
    st.session_state.generated_doc_bytes = None
    st.session_state.generated_context = None
    st.toast("Form cleared.")

def load_history_entry(entry):
    """Load a past history entry into the form."""
    st.session_state.form_data["ac_number"] = entry.get("ac_number", "")
    st.session_state.form_data["opened_date"] = entry.get("opened_date", "")
    st.session_state.form_data["purpose"] = entry.get("purpose", "")
    st.session_state["b1_name"] = entry.get("winning_bidder", "")
    st.session_state["b1_loc"] = entry.get("winning_location", "")
    st.session_state["b1_amt"] = entry.get("winning_amount", "")
    st.session_state.generated_doc_bytes = entry.get("doc_bytes", None)
    st.toast(f"Loaded {entry.get('ac_number', '')} into form.")

def clear_history():
    st.session_state.abstract_history = []
    st.toast("History cleared.")

def auto_arrange_bidders(b1_tuple, b2_tuple, b3_tuple):
    """Sort the 3 bidders from lowest to highest numerical bid amount."""
    bidders_list = [b1_tuple, b2_tuple, b3_tuple]
    sorted_list = sorted(bidders_list, key=lambda x: x["num"] if x["num"] is not None else 999999999999.0)
    
    st.session_state["b1_name"] = sorted_list[0]["name"]
    st.session_state["b1_loc"] = sorted_list[0]["loc"]
    st.session_state["b1_amt"] = sorted_list[0]["amt"]
    
    st.session_state["b2_name"] = sorted_list[1]["name"]
    st.session_state["b2_loc"] = sorted_list[1]["loc"]
    st.session_state["b2_amt"] = sorted_list[1]["amt"]
    
    st.session_state["b3_name"] = sorted_list[2]["name"]
    st.session_state["b3_loc"] = sorted_list[2]["loc"]
    st.session_state["b3_amt"] = sorted_list[2]["amt"]
    st.toast("Bidders auto-arranged from lowest (winner) to highest.")

# -----------------------------------------------------------------------------
# Sidebar: Clean History of Abstract Number and Winner Only
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<div class='sidebar-heading'>Abstract History</div>", unsafe_allow_html=True)
    
    history_items = st.session_state.abstract_history
    if not history_items:
        st.caption("No abstracts recorded yet.")
    else:
        for idx, item in enumerate(history_items):
            ac_num = item.get("ac_number", "—")
            winner = item.get("winning_bidder", "—")
            amount = item.get("winning_amount", "")
            time_tag = item.get("time_str", "")
            
            st.markdown(f"""
            <div class="history-item">
                <div class="history-ac">{ac_num}</div>
                <div class="history-winner"><strong>Winner:</strong> {winner}</div>
                <div class="history-meta">
                    <span>{amount}</span>
                    <span>{time_tag}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("Load into Form", key=f"hist_btn_{idx}", use_container_width=True):
                load_history_entry(item)
                st.rerun()

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Clear History", use_container_width=True, on_click=clear_history):
            pass

# -----------------------------------------------------------------------------
# Main Header Bar with Logo and Blue-Green Gradient
# -----------------------------------------------------------------------------
logo_img_tag = f'<img src="data:image/png;base64,{logo_b64}" class="app-header-logo" alt="Maragondon Logo">' if logo_b64 else ''

st.markdown(f"""
<div class="app-header-bar">
    {logo_img_tag}
    <div>
        <div class="app-header-title">Abstract of Canvass Generator</div>
        <div class="app-header-subtitle">Republic of the Philippines • Province of Cavite • Municipality of Maragondon • Committee on Awards</div>
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Section 1: General Canvass Information
# -----------------------------------------------------------------------------
with st.container():
    col_hdr, col_acts = st.columns([7, 3])
    with col_hdr:
        st.markdown("<div class='panel-title'>1. General Canvass Information</div>", unsafe_allow_html=True)
    with col_acts:
        act_c1, act_c2 = st.columns(2)
        with act_c1:
            st.button("Load Sample", use_container_width=True, on_click=load_sample_data)
        with act_c2:
            st.button("Clear Form", use_container_width=True, on_click=clear_form_data)
    
    col_meta1, col_meta2 = st.columns(2)
    with col_meta1:
        ac_number_input = st.text_input(
            "Abstract Number",
            value=st.session_state.form_data.get("ac_number", ""),
            placeholder="e.g. AC –0433H-10-26"
        )
    with col_meta2:
        opened_date_input = st.text_input(
            "Opened Date (Look for Closing Date)",
            value=st.session_state.form_data.get("opened_date", "October 26, 2026"),
            placeholder="e.g. October 26, 2026"
        )

    raw_purpose_input = st.text_area(
        "Purpose / Description",
        value=st.session_state.form_data.get("purpose", ""),
        placeholder="e.g. FOR: PAYMENT FOR THE PURCHASED OF TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD IN LGU- MARAGONDON, CAVITE",
        height=60
    )

    auto_clean = st.checkbox(
        "Auto-trim boilerplate prefix/suffix (matches template wording)",
        value=st.session_state.form_data.get("auto_clean_purpose", True)
    )

    if auto_clean:
        final_purpose = clean_purpose_text(raw_purpose_input)
    else:
        final_purpose = raw_purpose_input.strip()

    st.caption(f"Document Line Preview: FOR: PAYMENT FOR THE PURCHASED OF **{final_purpose or '[Description]'}** IN LGU MARAGONDON, CAVITE.")

# -----------------------------------------------------------------------------
# Section 2: Bidders Section (Strict 2-Column Mode: Bidders & Amount)
# -----------------------------------------------------------------------------
with st.container():
    st.markdown("<div class='panel-title' style='margin-top: 10px;'>2. Bidders & Bid Amounts (Inside Table in Word Document)</div>", unsafe_allow_html=True)

    # Table Header Row: 2 Columns Only
    h_col1, h_col2 = st.columns([7, 3])
    with h_col1:
        st.markdown("**NAME OF BIDDERS & LOCATION**")
    with h_col2:
        st.markdown("**AMOUNT OF BID**")
    
    st.markdown("<hr style='margin: 4px 0 8px 0;'>", unsafe_allow_html=True)

    # --- BIDDER 1 ROW ---
    r1_c1, r1_c2 = st.columns([7, 3])
    with r1_c1:
        st.caption("Bidder 1 (Winning Bidder Slot)")
        b1_name = st.text_input(
            "Bidder 1 Name",
            value=st.session_state.get("b1_name", st.session_state.form_data.get("bidder_1_name", "")),
            placeholder="Bidder 1 Name (e.g. RBK PRINTING SERVICES)",
            key="b1_name",
            label_visibility="collapsed"
        )
        b1_loc = st.text_input(
            "Bidder 1 Location",
            value=st.session_state.get("b1_loc", st.session_state.form_data.get("bidder_1_location", "")),
            placeholder="Bidder 1 Location (e.g. Bucal 1, Maragondon, Cavite)",
            key="b1_loc",
            label_visibility="collapsed"
        )
    with r1_c2:
        st.caption("Amount 1")
        b1_amt_raw = st.text_input(
            "Bidder 1 Amount",
            value=st.session_state.get("b1_amt", st.session_state.form_data.get("bidder_1_amount", "")),
            placeholder="e.g. ₱ 1,000.00",
            key="b1_amt",
            label_visibility="collapsed"
        )
        b1_num = parse_bid_amount(b1_amt_raw)
        b1_amt_formatted = format_bid_amount(b1_amt_raw)
        if b1_num is not None:
            st.markdown(f"<span class='amount-display'>{b1_amt_formatted}</span>", unsafe_allow_html=True)

    st.markdown("<hr style='margin: 4px 0 8px 0;'>", unsafe_allow_html=True)

    # --- BIDDER 2 ROW ---
    r2_c1, r2_c2 = st.columns([7, 3])
    with r2_c1:
        st.caption("Bidder 2")
        b2_name = st.text_input(
            "Bidder 2 Name",
            value=st.session_state.get("b2_name", st.session_state.form_data.get("bidder_2_name", "")),
            placeholder="Bidder 2 Name (e.g. CAVITE PRINT EXPRESS)",
            key="b2_name",
            label_visibility="collapsed"
        )
        b2_loc = st.text_input(
            "Bidder 2 Location",
            value=st.session_state.get("b2_loc", st.session_state.form_data.get("bidder_2_location", "")),
            placeholder="Bidder 2 Location (e.g. Naic, Cavite)",
            key="b2_loc",
            label_visibility="collapsed"
        )
    with r2_c2:
        st.caption("Amount 2")
        b2_amt_raw = st.text_input(
            "Bidder 2 Amount",
            value=st.session_state.get("b2_amt", st.session_state.form_data.get("bidder_2_amount", "")),
            placeholder="e.g. ₱ 1,250.00",
            key="b2_amt",
            label_visibility="collapsed"
        )
        b2_num = parse_bid_amount(b2_amt_raw)
        b2_amt_formatted = format_bid_amount(b2_amt_raw)
        if b2_num is not None:
            st.markdown(f"<span class='amount-display'>{b2_amt_formatted}</span>", unsafe_allow_html=True)

    st.markdown("<hr style='margin: 4px 0 8px 0;'>", unsafe_allow_html=True)

    # --- BIDDER 3 ROW ---
    r3_c1, r3_c2 = st.columns([7, 3])
    with r3_c1:
        st.caption("Bidder 3")
        b3_name = st.text_input(
            "Bidder 3 Name",
            value=st.session_state.get("b3_name", st.session_state.form_data.get("bidder_3_name", "")),
            placeholder="Bidder 3 Name (e.g. MARAGONDON GRAPHICS & PRINTS)",
            key="b3_name",
            label_visibility="collapsed"
        )
        b3_loc = st.text_input(
            "Bidder 3 Location",
            value=st.session_state.get("b3_loc", st.session_state.form_data.get("bidder_3_location", "")),
            placeholder="Bidder 3 Location (e.g. Poblacion 1A, Maragondon, Cavite)",
            key="b3_loc",
            label_visibility="collapsed"
        )
    with r3_c2:
        st.caption("Amount 3")
        b3_amt_raw = st.text_input(
            "Bidder 3 Amount",
            value=st.session_state.get("b3_amt", st.session_state.form_data.get("bidder_3_amount", "")),
            placeholder="e.g. ₱ 1,500.00",
            key="b3_amt",
            label_visibility="collapsed"
        )
        b3_num = parse_bid_amount(b3_amt_raw)
        b3_amt_formatted = format_bid_amount(b3_amt_raw)
        if b3_num is not None:
            st.markdown(f"<span class='amount-display'>{b3_amt_formatted}</span>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Bid Amount Validation & Lowest-to-Higher Checker
    # -------------------------------------------------------------------------
    b1_tuple = {"name": b1_name, "loc": b1_loc, "amt": b1_amt_formatted, "num": b1_num}
    b2_tuple = {"name": b2_name, "loc": b2_loc, "amt": b2_amt_formatted, "num": b2_num}
    b3_tuple = {"name": b3_name, "loc": b3_loc, "amt": b3_amt_formatted, "num": b3_num}
    
    validation_errors = []
    
    if b1_num is not None and b2_num is not None and b1_num > b2_num:
        validation_errors.append(f"Bidder 1 ({b1_amt_formatted}) is higher than Bidder 2 ({b2_amt_formatted}). Bidder 1 (winner) must have the lowest amount.")
    if b2_num is not None and b3_num is not None and b2_num > b3_num:
        validation_errors.append(f"Bidder 2 ({b2_amt_formatted}) is higher than Bidder 3 ({b3_amt_formatted}). Amounts must be entered from lowest to higher.")
    if b1_num is not None and b3_num is not None and b1_num > b3_num:
        validation_errors.append(f"Bidder 1 ({b1_amt_formatted}) is higher than Bidder 3 ({b3_amt_formatted}). Bidder 1 (winner) must have the lowest amount.")
    
    if validation_errors:
        st.markdown(f"""
        <div class="val-box-error">
            <strong>Validation Warning:</strong> The winning bidder must have the lowest bid, and amounts must be entered from lowest to higher.<br>
            {'<br>'.join(validation_errors)}
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("Auto-Arrange Bidders (Lowest to Highest)", use_container_width=True):
            auto_arrange_bidders(b1_tuple, b2_tuple, b3_tuple)
            st.rerun()
    elif b1_num is not None and b2_num is not None and b3_num is not None:
        st.markdown("""
        <div class="val-box-success">
            <strong>Validated:</strong> Bidders are correctly ordered from lowest (winner) to higher.
        </div>
        """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Section 3: Award Section
# -----------------------------------------------------------------------------
with st.container():
    st.markdown("<div class='panel-title' style='margin-top: 10px;'>3. Award Section</div>", unsafe_allow_html=True)

    bidders_data = [
        {"index": 1, "name": b1_name, "location": b1_loc, "num": b1_num, "amt_str": b1_amt_formatted},
        {"index": 2, "name": b2_name, "location": b2_loc, "num": b2_num, "amt_str": b2_amt_formatted},
        {"index": 3, "name": b3_name, "location": b3_loc, "num": b3_num, "amt_str": b3_amt_formatted},
    ]

    valid_bids = [b for b in bidders_data if b["num"] is not None and b["name"].strip()]
    if valid_bids:
        lowest_bid = min(valid_bids, key=lambda x: x["num"])
        lowest_summary_text = f"Bidder {lowest_bid['index']} ({lowest_bid['name']} - {lowest_bid['amt_str']})"
    else:
        lowest_bid = bidders_data[0]
        lowest_summary_text = "Bidder 1"

    aw_col1, aw_col2 = st.columns([1, 1])
    with aw_col1:
        award_choice = st.selectbox(
            "Winning Bidder Selection",
            options=[
                f"Auto: Lowest Bidder [{lowest_summary_text}]",
                f"Bidder 1 ({b1_name or 'Bidder 1'})",
                f"Bidder 2 ({b2_name or 'Bidder 2'})",
                f"Bidder 3 ({b3_name or 'Bidder 3'})",
                "Custom Override"
            ],
            index=0
        )

    if award_choice.startswith("Auto"):
        winning_name = lowest_bid["name"]
        winning_loc = lowest_bid["location"]
        winning_amt = lowest_bid["amt_str"]
        winning_num = lowest_bid["num"]
    elif "Bidder 1" in award_choice:
        winning_name = b1_name
        winning_loc = b1_loc
        winning_amt = b1_amt_formatted
        winning_num = b1_num
    elif "Bidder 2" in award_choice:
        winning_name = b2_name
        winning_loc = b2_loc
        winning_amt = b2_amt_formatted
        winning_num = b2_num
    elif "Bidder 3" in award_choice:
        winning_name = b3_name
        winning_loc = b3_loc
        winning_amt = b3_amt_formatted
        winning_num = b3_num
    else:
        winning_name = st.text_input("Custom Winner Name", value=b1_name)
        winning_loc = st.text_input("Custom Winner Location", value=b1_loc)
        winning_amt = ""
        winning_num = b1_num

    winner_is_valid = True
    if winning_num is not None:
        for b in bidders_data:
            if b["num"] is not None and b["name"] != winning_name and winning_num > b["num"]:
                winner_is_valid = False
                break

    with aw_col2:
        st.markdown(f"""
        <div style="background: rgba(30, 41, 59, 0.8); border: 1px solid rgba(148, 163, 184, 0.3); padding: 8px 12px; border-radius: 5px; font-size: 12.5px; color: #f8fafc;">
            <strong>Awardee:</strong> <span style="color:#38bdf8; font-weight:700;">{winning_name or '—'}</span><br>
            <strong>Location:</strong> {winning_loc or '—'}<br>
            <strong>Amount:</strong> <span style="color:#34d399; font-weight:700;">{winning_amt or '—'}</span>
        </div>
        """, unsafe_allow_html=True)

    if not winner_is_valid:
        st.markdown(f"""
        <div class="val-box-error">
            <strong>Invalid Winner:</strong> The selected winner ({winning_name}) has a bid amount of {winning_amt}, which is higher than one of the competitors. The winner cannot have a higher amount than the competitors.
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="award-summary-box">
        <strong>Official Award Clause in Word Document:</strong><br>
        <em>Awarded to the lowest bidder <strong>{winning_name or '[Winning Bidder]'}</strong>, 
        <strong>{winning_loc or '[Location]'}</strong> same being reasonable and most advantageous to the government.</em>
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Document Generation Button
# -----------------------------------------------------------------------------
st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
generate_clicked = st.button("Generate Document", type="primary", use_container_width=True)

if generate_clicked:
    errors = []
    if not ac_number_input.strip():
        errors.append("Abstract Number is required.")
    if not final_purpose.strip():
        errors.append("Purpose / Description is required.")
    if not b1_name.strip():
        errors.append("Bidder 1 Name is required.")
    if not winning_name.strip():
        errors.append("Winning Bidder Name is required.")
    
    if not winner_is_valid:
        errors.append("The winning bidder cannot have a higher bid amount than competing bidders. Please adjust amounts or use Auto-Arrange.")
        
    template_path = os.path.join(os.getcwd(), TEMPLATE_FILENAME)
    if not os.path.exists(template_path):
        errors.append(f"Template file '{TEMPLATE_FILENAME}' not found in folder.")

    if errors:
        for err in errors:
            st.error(err)
    else:
        try:
            context = {
                "purpose": final_purpose,
                "opened_date": opened_date_input.strip(),
                "bidder_1_name": b1_name.strip(),
                "bidder_1_location": b1_loc.strip(),
                "bidder_1_amount": b1_amt_formatted,
                "bidder_2_name": b2_name.strip(),
                "bidder_2_location": b2_loc.strip(),
                "bidder_2_amount": b2_amt_formatted,
                "bidder_3_name": b3_name.strip(),
                "bidder_3_location": b3_loc.strip(),
                "bidder_3_amount": b3_amt_formatted,
                "winning_bidder_name": winning_name.strip(),
                "winning_bidder_location": winning_loc.strip(),
                "ac_number": ac_number_input.strip(),
            }
            
            tpl = DocxTemplate(template_path)
            tpl.render(context, autoescape=True)
            
            out_stream = io.BytesIO()
            tpl.save(out_stream)
            doc_bytes = out_stream.getvalue()
            
            st.session_state.generated_doc_bytes = doc_bytes
            st.session_state.generated_context = context
            st.session_state.generated_ac_number = ac_number_input.strip()
            
            now_str = datetime.datetime.now().strftime("%b %d, %I:%M %p")
            history_entry = {
                "ac_number": ac_number_input.strip(),
                "winning_bidder": winning_name.strip(),
                "winning_location": winning_loc.strip(),
                "winning_amount": winning_amt,
                "opened_date": opened_date_input.strip(),
                "purpose": final_purpose,
                "time_str": now_str,
                "doc_bytes": doc_bytes
            }
            if not st.session_state.abstract_history or st.session_state.abstract_history[0].get("ac_number") != ac_number_input.strip():
                st.session_state.abstract_history.insert(0, history_entry)
            
            st.success("Document successfully generated.")
            
        except Exception as e:
            st.error(f"Failed to generate document: {str(e)}")

# -----------------------------------------------------------------------------
# Section 4: Realistic Printing Preview Sheet (Exact 1:1 Print Replica)
# -----------------------------------------------------------------------------
if st.session_state.generated_doc_bytes is not None:
    st.markdown("<hr style='margin: 18px 0 12px 0;'>", unsafe_allow_html=True)
    
    ctx = st.session_state.generated_context
    clean_ac = re.sub(r"[^\w\-]", "_", ctx["ac_number"]).strip("_")
    download_filename = f"Abstract_of_Canvass_{clean_ac or 'filled'}.docx"
    
    # HTML-escape all text for safe browser preview rendering without entity corruption
    esc_ctx = {k: html.escape(str(v)) if v is not None else "" for k, v in ctx.items()}
    
    # Top Action Bar with Download Button
    exp_col1, exp_col2 = st.columns([7, 3])
    with exp_col1:
        st.markdown("<div class='panel-title' style='margin: 0;'>Printing Preview</div>", unsafe_allow_html=True)
        st.caption("Realistic 1:1 view of the document as it will appear when printed on paper.")
    with exp_col2:
        st.download_button(
            label="Download Word Document (.docx)",
            data=st.session_state.generated_doc_bytes,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    
    # HTML Print Preview Sheet (using st.html for reliable DOM rendering)
    sheet_html = f"""
    <div class="print-preview-container">
        <div class="print-preview-topbar">
            <div>Page Size: <span>US Legal (8.5 × 14 in)</span> • Print Layout: <span>100% Full View</span></div>
            <div>Status: <span style="color:#10b981;">Ready to Print / Download</span></div>
        </div>
        <div class="print-preview-sheet" style="background-color: #ffffff !important; color: #000000 !important; display: block !important;">
            <!-- Header Block: Municipal Seal floating at top-left, Centered Header Text -->
            <div style="position: relative; min-height: 85px; margin-bottom: 24px;">
                <img src="data:image/png;base64,{seal_b64}" style="position: absolute; left: 8px; top: -6px; width: 88px; height: 88px; object-fit: contain;" alt="Municipal Seal">
                <div style="text-align: center; font-size: 15px; line-height: 1.35; padding-top: 4px;">
                    <div>Republic of the Philippines</div>
                    <div>Province of Cavite</div>
                    <div>MUNICIPALITY OF MARAGONDON</div>
                </div>
            </div>
            
            <!-- Document Title (Abstract of Canvass) -->
            <div style="text-align: center; font-size: 16px; font-weight: 700; letter-spacing: 0.5px; margin: 26px 0 20px 0;">
                ABSTRACT OF CANVASS
            </div>
            
            <!-- Purpose Line (Exact template format: justified, regular font, no underline) -->
            <div style="font-size: 15px; text-align: justify; line-height: 1.4; margin-bottom: 18px;">
                FOR: PAYMENT FOR THE PURCHASED OF {esc_ctx['purpose']} IN LGU MARAGONDON, CAVITE.
            </div>
            
            <!-- Opened Date (Positioned on the right side with tabbed indent) -->
            <div style="display: flex; justify-content: flex-end; margin-bottom: 22px;">
                <div style="font-size: 15px; line-height: 1.4; text-align: left; min-width: 295px;">
                    <div>Opened at the Office of the Municipal</div>
                    <div style="padding-left: 14px;">Mayor on {esc_ctx['opened_date']}</div>
                </div>
            </div>
            
            <!-- Bidders & Bid Amounts Table (Exact 1:1 replica: BORDERLESS 2-column table) -->
            <table style="width: 100%; border-collapse: collapse; border: none; margin: 16px 0 26px 0;">
                <thead>
                    <tr>
                        <th style="width: 68%; padding: 4px 0; text-align: left; font-size: 15px; font-weight: 700; border: none;">
                            <span style="padding-left: 18px;">NAME OF BIDDERS</span>
                        </th>
                        <th style="width: 32%; padding: 4px 0; text-align: left; font-size: 15px; font-weight: 700; border: none;">
                            AMOUNT OF BID
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; border: none;">
                            <div style="font-size: 15px; padding-left: 18px;">1. {esc_ctx['bidder_1_name']}</div>
                            <div style="font-size: 15px; padding-left: 42px; margin-top: 2px;">{esc_ctx['bidder_1_location']}</div>
                        </td>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; text-align: left; font-size: 15px; border: none;">
                            {esc_ctx['bidder_1_amount']}
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; border: none;">
                            <div style="font-size: 15px; padding-left: 18px;">2. {esc_ctx['bidder_2_name']}</div>
                            <div style="font-size: 15px; padding-left: 42px; margin-top: 2px;">{esc_ctx['bidder_2_location']}</div>
                        </td>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; text-align: left; font-size: 15px; border: none;">
                            {esc_ctx['bidder_2_amount']}
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; border: none;">
                            <div style="font-size: 15px; padding-left: 18px;">3. {esc_ctx['bidder_3_name']}</div>
                            <div style="font-size: 15px; padding-left: 42px; margin-top: 2px;">{esc_ctx['bidder_3_location']}</div>
                        </td>
                        <td style="padding: 7px 0 4px 0; vertical-align: top; text-align: left; font-size: 15px; border: none;">
                            {esc_ctx['bidder_3_amount']}
                        </td>
                    </tr>
                </tbody>
            </table>
            
            <!-- Committee on Awards Block (Exact template format: centered, bold names, no underlines) -->
            <div style="margin-top: 28px; text-align: center;">
                <div style="font-weight: 700; font-size: 15px; margin-bottom: 26px;">
                    COMMITTEE ON AWARDS
                </div>
                
                <div style="margin-bottom: 24px;">
                    <div style="font-weight: 700; font-size: 15px;">ENGR. ALDWIN JESTER A. ANCAYAN</div>
                    <div style="font-size: 13.5px; margin-top: 2px;">BAC CHAIRMAN/MUN. ENGINEER</div>
                </div>
                
                <div style="margin-bottom: 24px;">
                    <div style="font-weight: 700; font-size: 15px;">CHERRYL I. ESQUILLO</div>
                    <div style="font-size: 13.5px; margin-top: 2px;">BAC VICE-CHAIRMAN/MUN. BUDGET OFFICER</div>
                </div>
                
                <div style="display: flex; justify-content: space-between; margin-bottom: 24px; padding: 0 10px;">
                    <div style="text-align: center; width: 48%;">
                        <div style="font-weight: 700; font-size: 15px;">REGINO S. REYES</div>
                        <div style="font-size: 13.5px; margin-top: 2px;">BAC MEMBER/MUN. AGRICULTURIST</div>
                    </div>
                    <div style="text-align: center; width: 48%;">
                        <div style="font-weight: 700; font-size: 15px;">GERARDO Z. RIELO JR.</div>
                        <div style="font-size: 13.5px; margin-top: 2px;">BAC MEMBER/MDRRMO</div>
                    </div>
                </div>
                
                <div style="margin-bottom: 26px;">
                    <div style="font-weight: 700; font-size: 15px;">ENGR. ENP. PERCIVAL M. MALIMBAN</div>
                    <div style="font-size: 13.5px; margin-top: 2px;">BAC MEMBER/MUN. PLANNING &amp; DEV'T COORDINATOR</div>
                </div>
            </div>
            
            <!-- Award Clause (Exact template format: winner is bold, no underlines) -->
            <div style="margin-top: 22px; font-size: 15px; text-align: justify; line-height: 1.45;">
                Awarded to the lowest bidder <strong>{esc_ctx['winning_bidder_name']}</strong>, {esc_ctx['winning_bidder_location']} same being reasonable and most advantageous to the government.
            </div>
            
            <!-- Abstract Number (Exact template placement: bottom left, regular weight) -->
            <div style="text-align: left; margin-top: 34px; font-size: 15px;">
                {esc_ctx['ac_number']}
            </div>
        </div>
    </div>
    """
    st.html(sheet_html)
