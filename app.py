import streamlit as st
import streamlit.components.v1 as components
import os
import io
import re
import html
import base64
import datetime
import calendar
import time
import json
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
    initial_sidebar_state="auto"
)

DEFAULT_SAMPLE_DATA = {
    "ac_code": "0433H",
    "ac_month": "10",
    "ac_year": "26",
    "ac_number": "AC-0433H-10-26",
    "rfq_closing_date": datetime.date(2026, 10, 26),
    "opened_date": "October 26, 2026",
    "purpose_desc": "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD",
    "include_lgu_suffix": True,
    "purpose": "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD IN LGU MARAGONDON, CAVITE.",
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
    
    /* Hide Deploy button, Three dots menu, and right-hand toolbar widgets only */
    [data-testid="stAppDeployButton"],
    [data-testid="stDeployButton"],
    .stAppDeployButton,
    .stDeployButton,
    button[data-testid="stAppDeployButton"],
    button[data-testid="stDeployButton"],
    #MainMenu,
    [data-testid="stMainMenu"],
    [data-testid="stStatusWidget"],
    [data-testid="stToolbarActions"],
    [data-testid="stHeaderActionElements"] {{
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0px !important;
        width: 0px !important;
        padding: 0 !important;
        margin: 0 !important;
    }}

    /* Transparent header bar that keeps the sidebar toggle controls fully functional */
    header[data-testid="stHeader"],
    .stApp > header {{
        background: transparent !important;
        height: 48px !important;
        min-height: 48px !important;
        border: none !important;
        box-shadow: none !important;
        z-index: 50 !important;
        pointer-events: none !important;
    }}
    header[data-testid="stHeader"] button,
    .stApp > header button {{
        pointer-events: auto !important;
    }}

    /* Toolbar container: transparent and passes through clicks, keeping stExpandSidebarButton active */
    [data-testid="stToolbar"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        pointer-events: none !important;
        height: 48px !important;
    }}

    /* Sidebar collapse / expand controls: ALWAYS visible, interactive, and responsive */
    [data-testid="stExpandSidebarButton"],
    button[data-testid="stExpandSidebarButton"],
    [data-testid="stSidebarCollapseButton"],
    button[data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"] button,
    section[data-testid="stSidebar"] button,
    [data-testid="stSidebarCollapsedControl"],
    div[data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {{
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        pointer-events: auto !important;
        z-index: 999999 !important;
    }}
    
    [data-testid="stExpandSidebarButton"],
    button[data-testid="stExpandSidebarButton"],
    [data-testid="stSidebarCollapsedControl"],
    div[data-testid="stSidebarCollapsedControl"] {{
        position: fixed !important;
        top: 8px !important;
        left: 12px !important;
    }}
    
    /* Ensure sidebar collapse outer wrapper has no border to prevent double-border artifact */
    [data-testid="stSidebarCollapseButton"],
    div[data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"],
    div[data-testid="stSidebarHeader"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
        outline: none !important;
    }}
    
    /* Clean single border strictly on the interactive button */
    [data-testid="stExpandSidebarButton"],
    button[data-testid="stExpandSidebarButton"],
    [data-testid="stSidebarCollapseButton"] button,
    [data-testid="stSidebarHeader"] button,
    button[aria-label="Collapse sidebar"],
    button[aria-label="Expand sidebar"] {{
        background: rgba(15, 23, 42, 0.85) !important;
        border: 1px solid rgba(56, 189, 248, 0.45) !important;
        border-radius: 6px !important;
        color: #38bdf8 !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.5) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        cursor: pointer !important;
        padding: 5px 8px !important;
        outline: none !important;
    }}
    [data-testid="stExpandSidebarButton"]:hover,
    button[data-testid="stExpandSidebarButton"]:hover,
    [data-testid="stSidebarCollapseButton"] button:hover,
    [data-testid="stSidebarHeader"] button:hover,
    button[aria-label="Collapse sidebar"]:hover,
    button[aria-label="Expand sidebar"]:hover {{
        background: rgba(30, 41, 59, 0.95) !important;
        border-color: #34d399 !important;
        color: #6ee7b7 !important;
        transform: scale(1.08) !important;
        box-shadow: 0 4px 16px rgba(56, 189, 248, 0.4) !important;
    }}
    
    /* Ensure no nested icon or span inside the button gets its own border */
    [data-testid="stSidebarCollapseButton"] button *,
    [data-testid="stExpandSidebarButton"] * {{
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        outline: none !important;
    }}

    /* Remove the split-second 'Press Enter to apply' tooltip */
    [data-testid="stInputInstructions"],
    div[data-testid="InputInstructions"] {{
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
    
    /* Layout: Magnetic Pull Effect to the Center on Sidebar Toggle */
    [data-testid="stAppViewContainer"] {{
        background: transparent !important;
        transition: all 0.45s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}

    /* Intermediate flex wrapper & main display smooth animation transitions */
    [data-testid="stAppViewContainer"] > div:not([data-testid="stSidebar"]),
    [data-testid="stAppViewContainer"] div:has(> [data-testid="stMain"]),
    header[data-testid="stHeader"] {{
        transition: margin 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    transform 0.4s cubic-bezier(0.16, 1, 0.3, 1) !important;
        will-change: margin, width;
    }}

    [data-testid="stMain"] {{
        background: transparent !important;
        position: relative;
        z-index: 1;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        transition: margin 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    transform 0.4s cubic-bezier(0.16, 1, 0.3, 1) !important;
        will-change: margin, width;
    }}

    [data-testid="stMainBlockContainer"],
    section.main,
    .main {{
        background: transparent !important;
        width: 100% !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: flex-start !important;
        height: auto !important;
        min-height: 100% !important;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    
    /* Center container smoothly with magnetic pull */
    .block-container {{
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 1080px !important;
        width: 100% !important;
        margin: 0 auto !important;
        transition: max-width 0.4s cubic-bezier(0.16, 1, 0.3, 1), 
                    margin 0.4s cubic-bezier(0.16, 1, 0.3, 1), 
                    padding 0.4s cubic-bezier(0.16, 1, 0.3, 1) !important;
        will-change: max-width, margin;
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
        z-index: 1000 !important;
        transition: transform 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    margin 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    min-width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    max-width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    opacity 0.3s ease !important;
        will-change: transform, width, margin;
    }}
    
    /* Sidebar collapsed state: cleanly collapse with smooth slide transition */
    section[data-testid="stSidebar"][data-collapsed="true"],
    section[data-testid="stSidebar"][aria-expanded="false"] {{
        width: 0px !important;
        min-width: 0px !important;
        max-width: 0px !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        overflow: hidden !important;
        pointer-events: none !important;
        transform: translateX(-100%) !important;
        opacity: 0 !important;
        transition: transform 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    margin 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    min-width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    max-width 0.4s cubic-bezier(0.16, 1, 0.3, 1),
                    opacity 0.25s ease !important;
    }}
    
    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) > div:not([data-testid="stSidebar"]),
    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) div:has(> [data-testid="stMain"]),
    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) [data-testid="stMain"],
    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) header[data-testid="stHeader"],
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) > div:not([data-testid="stSidebar"]),
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) div:has(> [data-testid="stMain"]),
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) [data-testid="stMain"],
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) header[data-testid="stHeader"] {{
        margin-left: 0 !important;
        transform: none !important;
        left: 0 !important;
        width: 100% !important;
        max-width: 100% !important;
        transition: margin 0.45s cubic-bezier(0.16, 1, 0.3, 1), width 0.45s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}

    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) .block-container,
    [data-testid="stAppViewContainer"]:has(section[data-testid="stSidebar"][data-collapsed="true"]) [data-testid="stMainBlockContainer"],
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) .block-container,
    [data-testid="stAppViewContainer"]:has([data-testid="stExpandSidebarButton"][aria-expanded="false"]) [data-testid="stMainBlockContainer"] {{
        margin-left: auto !important;
        margin-right: auto !important;
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
        background: rgba(15, 23, 42, 0.65) !important;
        border: 1px solid rgba(56, 189, 248, 0.25) !important;
        border-left: 3.5px solid #10b981 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
        background-clip: padding-box !important;
        padding: 9px 11px !important;
        margin-bottom: 4px !important;
        box-sizing: border-box !important;
        cursor: pointer !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    .history-item:hover {{
        background: rgba(15, 23, 42, 0.92) !important;
        border: 1px solid #38bdf8 !important;
        border-left: 3.5px solid #38bdf8 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
        background-clip: padding-box !important;
        transform: translateX(3px) !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4), 0 0 10px rgba(56, 189, 248, 0.2) !important;
    }}
    .history-item.is-expanded {{
        border-left: 3.5px solid #38bdf8 !important;
        border-bottom-left-radius: 0 !important;
        border-bottom-right-radius: 0 !important;
        margin-bottom: 0 !important;
    }}
    .history-item.is-expanded:hover {{
        border-bottom-left-radius: 0 !important;
        border-bottom-right-radius: 0 !important;
        border-radius: 8px 8px 0 0 !important;
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
        margin-bottom: 2px !important;
    }}
    .history-meta {{
        font-size: 10.5px !important;
        color: #94a3b8 !important;
        margin-top: 4px !important;
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
    }}
    .history-amount {{
        font-weight: 700 !important;
        color: #34d399 !important;
        font-size: 11.5px !important;
    }}
    .history-top-row {{
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 6px !important;
        margin-bottom: 3px !important;
    }}
    .history-date-badge {{
        font-size: 9.5px !important;
        color: #94a3b8 !important;
        background: rgba(15, 23, 42, 0.6) !important;
        padding: 1px 5px !important;
        border-radius: 4px !important;
        border: 1px solid rgba(100, 116, 139, 0.25) !important;
        white-space: nowrap !important;
    }}
    .history-extra-info {{
        background: rgba(10, 16, 32, 0.95) !important;
        border: 1px solid rgba(56, 189, 248, 0.3) !important;
        border-top: none !important;
        border-left: 3.5px solid #38bdf8 !important;
        border-radius: 0 0 8px 8px !important;
        padding: 9px 11px 11px 11px !important;
        margin-bottom: 6px !important;
        font-size: 11px !important;
        color: #cbd5e1 !important;
        line-height: 1.4 !important;
        overflow: hidden !important;
        background-clip: padding-box !important;
    }}
    .history-extra-info-row {{
        margin-bottom: 4px !important;
        word-break: break-word !important;
    }}
    .history-extra-info-row strong {{
        color: #7dd3fc !important;
    }}
    .history-hidden-toggle,
    .stSidebar [class*="st-key-hist_tog_"] {{
        position: absolute !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0 !important;
        width: 0 !important;
        overflow: hidden !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    
    /* Compact Sidebar Action Buttons */
    .stSidebar [data-testid="stButton"] button {{
        height: 28px !important;
        min-height: 28px !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        padding: 2px 8px !important;
        border-radius: 6px !important;
        letter-spacing: 0.3px !important;
        margin: 0 !important;
        line-height: 1 !important;
    }}
    
    /* 60fps Hardware-Accelerated Animations for Header & Logo */
    @keyframes headerGlow60fps {{
        0% {{
            background-position: 0% 50%;
            border-color: rgba(56, 189, 248, 0.45);
            box-shadow: 0 4px 18px rgba(14, 165, 233, 0.25), 0 0 15px rgba(16, 185, 129, 0.15);
        }}
        50% {{
            background-position: 100% 50%;
            border-color: rgba(52, 211, 153, 0.6);
            box-shadow: 0 4px 24px rgba(16, 185, 129, 0.35), 0 0 25px rgba(56, 189, 248, 0.25);
        }}
        100% {{
            background-position: 0% 50%;
            border-color: rgba(56, 189, 248, 0.45);
            box-shadow: 0 4px 18px rgba(14, 165, 233, 0.25), 0 0 15px rgba(16, 185, 129, 0.15);
        }}
    }}

    @keyframes logoFloatPulse60fps {{
        0% {{
            transform: translateY(0px) scale(1) translateZ(0);
            filter: drop-shadow(0 0 6px rgba(56, 189, 248, 0.6)) drop-shadow(0 0 12px rgba(16, 185, 129, 0.35));
        }}
        50% {{
            transform: translateY(-2.5px) scale(1.05) translateZ(0);
            filter: drop-shadow(0 0 14px rgba(56, 189, 248, 0.9)) drop-shadow(0 0 22px rgba(52, 211, 153, 0.7));
        }}
        100% {{
            transform: translateY(0px) scale(1) translateZ(0);
            filter: drop-shadow(0 0 6px rgba(56, 189, 248, 0.6)) drop-shadow(0 0 12px rgba(16, 185, 129, 0.35));
        }}
    }}

    @keyframes titleShimmer60fps {{
        0% {{
            background-position: -200% 0;
        }}
        100% {{
            background-position: 200% 0;
        }}
    }}

    /* Header Bar: Animated 60fps Blue-Green Ambient Gradient */
    .app-header-bar {{
        background: linear-gradient(135deg, #091e3a 0%, #0d4251 35%, #064e3b 70%, #091e3a 100%) !important;
        background-size: 250% 250% !important;
        animation: headerGlow60fps 7s ease-in-out infinite alternate !important;
        color: #ffffff;
        padding: 13px 22px;
        border-radius: 10px;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 16px;
        border: 1px solid rgba(56, 189, 248, 0.35);
        will-change: background-position, box-shadow, border-color;
        transform: translateZ(0);
    }}
    .app-header-logo-wrapper {{
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }}
    .app-header-logo {{
        height: 52px;
        width: auto;
        object-fit: contain;
        flex-shrink: 0;
        animation: logoFloatPulse60fps 3.2s ease-in-out infinite alternate !important;
        will-change: transform, filter;
        transform: translateZ(0);
    }}
    .app-header-title {{
        margin: 0;
        font-size: 26px !important;
        font-weight: 800 !important;
        text-transform: uppercase !important;
        letter-spacing: 1.5px !important;
        line-height: 1.15;
        background: linear-gradient(90deg, #ffffff 0%, #38bdf8 30%, #34d399 70%, #ffffff 100%);
        background-size: 200% auto;
        color: #ffffff !important;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: titleShimmer60fps 5s linear infinite !important;
        will-change: background-position;
    }}
    .app-header-subtitle {{
        margin: 2px 0 0 0;
        font-size: 11px;
        color: #d1fae5;
        letter-spacing: 0.2px;
        font-weight: 400;
    }}
    .app-header-credit {{
        margin: 2px 0 0 0 !important;
        font-family: Arial, sans-serif !important;
        font-size: 11px !important;
        font-weight: 600 !important;
        color: #7dd3fc !important;
        letter-spacing: 0.35px !important;
        line-height: 1.25 !important;
        opacity: 0.95 !important;
        text-shadow: 0 1px 4px rgba(0, 0, 0, 0.4) !important;
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
    
    /* Unified Field Labels */
    .unified-field-label,
    label p {{
        font-family: Arial, Helvetica, sans-serif !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        color: #e2e8f0 !important;
        margin-bottom: 4px !important;
        line-height: 18px !important;
    }}
    .unified-field-label {{
        display: block !important;
        padding-top: 1px !important;
    }}

    /* Compact Single-Border Highlight Inputs: Text Inputs & Date Inputs */
    div[data-testid="stTextInputRootElement"],
    div[data-testid="stDateInputField"],
    div[data-baseweb="input"],
    div[data-baseweb="base-input"] {{
        background-color: rgba(30, 41, 59, 0.85) !important;
        border: 1px solid #334155 !important;
        border-radius: 5px !important;
        min-height: 34px !important;
        height: 34px !important;
        max-height: 34px !important;
        box-sizing: border-box !important;
        box-shadow: none !important;
        display: flex !important;
        align-items: center !important;
        padding: 0 !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease !important;
    }}
    div[data-testid="stTextInputRootElement"]:hover,
    div[data-testid="stDateInputField"]:hover,
    div[data-baseweb="input"]:hover,
    div[data-baseweb="base-input"]:hover {{
        border-color: #0ea5e9 !important;
        background-color: rgba(30, 41, 59, 0.95) !important;
    }}
    div[data-testid="stTextInputRootElement"]:focus-within,
    div[data-testid="stDateInputField"]:focus-within,
    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="base-input"]:focus-within {{
        border-color: #38bdf8 !important;
        box-shadow: 0 0 0 1px #38bdf8 !important;
        background-color: rgba(15, 23, 42, 0.98) !important;
    }}

    /* Inner text input element: NO double border, unified typography */
    input[data-testid="stTextInputField"],
    .stTextInput input,
    div[data-baseweb="input"] input {{
        background: transparent !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        color: #ffffff !important;
        font-family: Arial, Helvetica, sans-serif !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        padding: 0 8px !important;
        min-height: 32px !important;
        height: 32px !important;
        line-height: 32px !important;
        box-sizing: border-box !important;
    }}
    
    /* Date Input: Inner React Aria Field & Spinbuttons */
    div[data-testid="stDateInputField"] {{
        padding-left: 8px !important;
        padding-right: 6px !important;
    }}
    div[data-testid="stDateInputFieldsScroller"],
    div[data-testid="stDateInputField"] div.react-aria-DateField,
    div[data-testid="stDateInputField"] div[role="group"] {{
        font-family: Arial, Helvetica, sans-serif !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        color: #ffffff !important;
        display: flex !important;
        align-items: center !important;
        height: 32px !important;
        line-height: 32px !important;
    }}
    div[data-testid="stDateInputField"] span[role="spinbutton"],
    div[data-testid="stDateInputField"] span[data-type="literal"],
    div[data-testid="stDateInputField"] span {{
        font-family: Arial, Helvetica, sans-serif !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        color: #ffffff !important;
        line-height: 32px !important;
    }}
    div[data-testid="stDateInputField"] span[role="spinbutton"] {{
        outline: none !important;
        padding: 0 1px !important;
    }}
    div[data-testid="stDateInputField"] span[role="spinbutton"]:focus {{
        background-color: rgba(56, 189, 248, 0.3) !important;
        border-radius: 2px !important;
        color: #ffffff !important;
    }}
    button[data-testid="stDateInputCalendarButton"] {{
        background: transparent !important;
        border: none !important;
        color: #94a3b8 !important;
        padding: 0 4px !important;
        margin: 0 !important;
        cursor: pointer !important;
        display: flex !important;
        align-items: center !important;
        height: 32px !important;
    }}
    button[data-testid="stDateInputCalendarButton"]:hover {{
        color: #38bdf8 !important;
    }}

    /* BaseWeb Calendar Popover Styling */
    div[data-baseweb="popover"] {{
        z-index: 999999 !important;
    }}
    div[data-baseweb="calendar"] {{
        background-color: #0f172a !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.7) !important;
        padding: 8px !important;
    }}
    div[data-baseweb="calendar"] header {{
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        pointer-events: auto !important;
        height: auto !important;
        min-height: 38px !important;
        align-items: center !important;
        justify-content: space-between !important;
        padding-bottom: 6px !important;
        border-bottom: 1px solid rgba(56, 189, 248, 0.2) !important;
        margin-bottom: 6px !important;
    }}
    div[data-baseweb="calendar"] header * {{
        visibility: visible !important;
        opacity: 1 !important;
        color: #f8fafc !important;
    }}
    div[data-baseweb="calendar"] select {{
        background-color: #1e293b !important;
        color: #f8fafc !important;
        border: 1px solid #334155 !important;
        border-radius: 4px !important;
        padding: 2px 6px !important;
    }}

    /* Compact Textarea with single border highlight */
    div[data-baseweb="textarea"] {{
        background-color: rgba(30, 41, 59, 0.85) !important;
        border: 1px solid #334155 !important;
        border-radius: 5px !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease !important;
        box-shadow: none !important;
    }}
    div[data-baseweb="textarea"]:hover {{
        border-color: #0ea5e9 !important;
        background-color: rgba(30, 41, 59, 0.95) !important;
    }}
    div[data-baseweb="textarea"]:focus-within {{
        border-color: #38bdf8 !important;
        box-shadow: 0 0 0 1px #38bdf8 !important;
        background-color: rgba(15, 23, 42, 0.98) !important;
    }}
    .stTextArea textarea,
    div[data-baseweb="textarea"] textarea {{
        background: transparent !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        color: #ffffff !important;
        font-size: 13px !important;
        font-family: Arial, sans-serif !important;
        padding: 6px 8px !important;
        min-height: 56px !important;
        text-transform: uppercase !important;
    }}

    /* Document Line Preview Container */
    .doc-line-preview-container {{
        font-family: Arial, sans-serif !important;
        font-size: 13px !important;
        color: #94a3b8 !important;
        margin-top: 6px !important;
        margin-bottom: 8px !important;
    }}
    .doc-line-preview-bold {{
        font-weight: 700 !important;
        color: #ffffff !important;
        text-transform: uppercase !important;
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

    /* =========================================================================
       Mobile Responsive Adaptations (Smartphones & Tablets <= 768px)
       ========================================================================= */
    @media (max-width: 768px) {{
        /* Prevent horizontal page overflow on mobile */
        html, body, .stApp {{
            overflow-x: hidden !important;
            max-width: 100vw !important;
        }}

        /* Ensure normal block flow and zero vertical clipping on mobile */
        [data-testid="stMainBlockContainer"],
        section.main,
        .main {{
            display: block !important;
            height: auto !important;
            min-height: 100% !important;
            overflow-x: hidden !important;
        }}

        /* Container Spacing for Mobile */
        .block-container {{
            padding-top: 3.4rem !important;
            padding-bottom: 1.5rem !important;
            padding-left: 0.65rem !important;
            padding-right: 0.65rem !important;
            max-width: 100% !important;
        }}
        
        /* Floating sidebar expand button mobile position */
        [data-testid="stExpandSidebarButton"],
        button[data-testid="stExpandSidebarButton"] {{
            position: fixed !important;
            top: 10px !important;
            left: 10px !important;
            z-index: 99999 !important;
            background: rgba(15, 23, 42, 0.9) !important;
            border: 1px solid rgba(56, 189, 248, 0.4) !important;
            border-radius: 6px !important;
            padding: 4px 8px !important;
        }}
        
        /* Ensure full width layout on mobile viewports */
        [data-testid="stAppViewContainer"] > div:not([data-testid="stSidebar"]),
        [data-testid="stAppViewContainer"] div:has(> [data-testid="stMain"]),
        [data-testid="stMain"] {{
            margin-left: 0 !important;
            width: 100% !important;
            max-width: 100% !important;
        }}
        
        /* Sidebar on mobile as an overlay drawer */
        section[data-testid="stSidebar"] {{
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            height: 100vh !important;
            width: 85vw !important;
            max-width: 320px !important;
            z-index: 999999 !important;
            box-shadow: 10px 0 30px rgba(0, 0, 0, 0.8) !important;
        }}
        section[data-testid="stSidebar"][data-collapsed="true"],
        section[data-testid="stSidebar"][aria-expanded="false"] {{
            transform: translateX(-100%) !important;
            width: 0px !important;
            max-width: 0px !important;
            pointer-events: none !important;
        }}
        
        /* Header Bar: Compact on Mobile */
        .app-header-bar {{
            padding: 9px 12px !important;
            gap: 10px !important;
            margin-bottom: 12px !important;
            border-radius: 8px !important;
        }}
        .app-header-logo {{
            height: 38px !important;
        }}
        .app-header-title {{
            font-size: 16px !important;
            letter-spacing: 0.7px !important;
            line-height: 1.25 !important;
        }}
        .app-header-credit {{
            font-size: 9.5px !important;
            letter-spacing: 0.2px !important;
        }}
        
        /* Section 1 Header: Stack Title and Action Buttons */
        [data-testid="stHorizontalBlock"]:has(.panel-title) {{
            display: flex !important;
            flex-direction: column !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.panel-title) > div[data-testid="stColumn"] {{
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.panel-title) [data-testid="stHorizontalBlock"] {{
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.panel-title) [data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {{
            width: auto !important;
            flex: 1 1 0 !important;
            min-width: 0 !important;
        }}
        
        /* Section 1 Metadata: Stack Abstract Number and Date Pickers vertically */
        [data-testid="stHorizontalBlock"]:has(.unified-field-label) {{
            display: flex !important;
            flex-direction: column !important;
            gap: 12px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.unified-field-label) > div[data-testid="stColumn"] {{
            width: 100% !important;
            min-width: 100% !important;
            max-width: 100% !important;
            flex: 1 1 100% !important;
        }}
        
        /* Keep Abstract Number 4 micro-inputs in a clean horizontal row on mobile */
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) {{
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            gap: 4px !important;
            align-items: center !important;
            width: 100% !important;
            max-width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) > div[data-testid="stColumn"] {{
            min-width: 0 !important;
        }}
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) > div[data-testid="stColumn"]:first-child {{
            flex: 0 0 46px !important;
            width: 46px !important;
            min-width: 46px !important;
            max-width: 46px !important;
        }}
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) > div[data-testid="stColumn"]:nth-child(2) {{
            flex: 2 1 0 !important;
        }}
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) > div[data-testid="stColumn"]:nth-child(3),
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) > div[data-testid="stColumn"]:nth-child(4) {{
            flex: 1 1 0 !important;
        }}
        [data-testid="stHorizontalBlock"]:not(:has(.unified-field-label)):has(input[aria-label="Prefix"]) input {{
            font-size: 13px !important;
            padding: 5px 2px !important;
            text-align: center !important;
            width: 100% !important;
        }}
        
        /* Section 1 Checkboxes: Wrap labels cleanly without ellipsis truncation */
        [data-testid="stHorizontalBlock"]:has([data-testid="stCheckbox"]) {{
            display: flex !important;
            flex-direction: column !important;
            gap: 6px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has([data-testid="stCheckbox"]) > div[data-testid="stColumn"] {{
            width: 100% !important;
        }}
        [data-testid="stCheckbox"] label,
        [data-testid="stCheckbox"] label span,
        [data-testid="stCheckbox"] label div,
        [data-testid="stCheckbox"] label p {{
            white-space: normal !important;
            word-break: break-word !important;
            overflow: visible !important;
            text-overflow: unset !important;
            font-size: 12px !important;
            line-height: 1.35 !important;
        }}
        
        /* Document Line Preview Card on Mobile */
        .doc-line-preview-container {{
            font-size: 12.5px !important;
            padding: 8px 10px !important;
            line-height: 1.4 !important;
        }}
        
        /* Section 2 Column Header Row: match 5.8:4.2 proportion */
        [data-testid="stHorizontalBlock"]:has(.bidder-col-head) {{
            display: flex !important;
            flex-direction: row !important;
            gap: 8px !important;
            width: 100% !important;
            margin-bottom: 2px !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.bidder-col-head) > div[data-testid="stColumn"]:first-child {{
            flex: 5.8 1 0 !important;
            min-width: 0 !important;
        }}
        [data-testid="stHorizontalBlock"]:has(.bidder-col-head) > div[data-testid="stColumn"]:last-child {{
            flex: 4.2 1 0 !important;
            min-width: 0 !important;
        }}
        .bidder-col-head {{
            font-size: 11px !important;
            color: #94a3b8 !important;
            letter-spacing: 0.3px !important;
        }}
        
        /* Section 2: Proportional Bidders and Amounts on Mobile */
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 1 Name"]),
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 2 Name"]),
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 3 Name"]) {{
            display: flex !important;
            flex-direction: row !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 1 Name"]) > div[data-testid="stColumn"]:first-child,
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 2 Name"]) > div[data-testid="stColumn"]:first-child,
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 3 Name"]) > div[data-testid="stColumn"]:first-child {{
            flex: 5.8 1 0 !important;
            min-width: 0 !important;
        }}
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 1 Name"]) > div[data-testid="stColumn"]:last-child,
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 2 Name"]) > div[data-testid="stColumn"]:last-child,
        [data-testid="stHorizontalBlock"]:has(input[aria-label="Bidder 3 Name"]) > div[data-testid="stColumn"]:last-child {{
            flex: 4.2 1 0 !important;
            min-width: 0 !important;
        }}
        .amount-display {{
            font-size: 13px !important;
            line-height: 28px !important;
        }}
        
        /* Section 3 Award Card on Mobile */
        .award-summary-box {{
            font-size: 12px !important;
            padding: 8px 10px !important;
            line-height: 1.4 !important;
        }}
        
        /* Section 4 Action Buttons: Full-width stacked for easy tapping */
        [data-testid="stHorizontalBlock"]:has(button[kind="primaryDownload"]) {{
            display: flex !important;
            flex-direction: column !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(button[kind="primaryDownload"]) > div[data-testid="stColumn"] {{
            width: 100% !important;
            min-width: 100% !important;
        }}
        [data-testid="stHorizontalBlock"]:has(button[kind="primaryDownload"]) iframe[title="st.iframe"] {{
            width: 100% !important;
            height: 44px !important;
        }}
        
        /* Document Preview Paper container on Mobile */
        img[alt="Official Document Preview"] {{
            width: 100% !important;
            max-width: 100% !important;
            height: auto !important;
        }}
    }}
    
    @media (max-width: 480px) {{
        .app-header-bar {{
            padding: 7px 10px !important;
            gap: 8px !important;
        }}
        .app-header-logo {{
            height: 32px !important;
        }}
        .app-header-title {{
            font-size: 14px !important;
            letter-spacing: 0.5px !important;
        }}
        .app-header-credit {{
            font-size: 8.5px !important;
            letter-spacing: 0.1px !important;
        }}
        .panel-title {{
            font-size: 14px !important;
        }}
    }}
    /* =========================================================================
       Creative Municipal Loading Screen & Progress Bar
       ========================================================================= */
    /* Fullscreen Frameless Semi-Transparent (80%) Overlay in the Middle of Screen */
    .gen-modal-overlay {{
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        bottom: 0 !important;
        width: 100vw !important;
        height: 100vh !important;
        background: rgba(11, 19, 43, 0.8) !important;
        backdrop-filter: blur(8px) !important;
        -webkit-backdrop-filter: blur(8px) !important;
        z-index: 99999999 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        animation: overlayFadeIn 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        pointer-events: all !important;
    }}
    
    @keyframes overlayFadeIn {{
        from {{ opacity: 0; }}
        to {{ opacity: 1; }}
    }}
    
    .gen-overlay-content {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
        margin: auto !important;
        text-align: center !important;
        max-width: 520px !important;
        width: 90% !important;
        animation: contentScaleUp 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    
    @keyframes contentScaleUp {{
        from {{ transform: scale(0.95); opacity: 0; }}
        to {{ transform: scale(1); opacity: 1; }}
    }}

    /* =========================================================================
       Official Document Preview Modal Overlay (Hover over Main Page)
       ========================================================================= */
    .preview-modal-overlay {{
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        bottom: 0 !important;
        width: 100vw !important;
        height: 100vh !important;
        background: rgba(11, 19, 43, 0.85) !important;
        backdrop-filter: blur(10px) !important;
        -webkit-backdrop-filter: blur(10px) !important;
        z-index: 9999999 !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: flex-start !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        padding: 62px 12px 40px 12px !important;
        box-sizing: border-box !important;
        animation: overlayFadeIn 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    .preview-modal-overlay.is-hidden {{
        display: none !important;
    }}
    
    .preview-top-bar {{
        position: sticky !important;
        top: 0 !important;
        z-index: 100 !important;
        width: 100% !important;
        max-width: 900px !important;
        background: rgba(15, 23, 42, 0.94) !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        border: 1px solid rgba(56, 189, 248, 0.35) !important;
        border-radius: 12px !important;
        padding: 10px 16px !important;
        margin-bottom: 16px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 12px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.7) !important;
    }}
    
    .preview-title-box {{
        display: flex !important;
        align-items: center !important;
        gap: 10px !important;
    }}
    
    .preview-bar-logo {{
        width: 36px !important;
        height: 36px !important;
        border-radius: 50% !important;
        object-fit: contain !important;
    }}
    
    .preview-bar-heading {{
        font-family: Arial, sans-serif !important;
        font-size: 13.5px !important;
        font-weight: 800 !important;
        letter-spacing: 0.6px !important;
        color: #f8fafc !important;
        line-height: 1.2 !important;
    }}
    
    .preview-bar-sub {{
        font-family: Arial, sans-serif !important;
        font-size: 11px !important;
        color: #38bdf8 !important;
        font-weight: 600 !important;
    }}
    
    .preview-actions-group {{
        display: flex !important;
        align-items: center !important;
        gap: 8px !important;
    }}
    
    .btn-overlay-dl {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        color: #ffffff !important;
        font-family: Arial, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13.5px !important;
        border: 1px solid rgba(56, 189, 248, 0.7) !important;
        border-radius: 8px !important;
        padding: 0 16px !important;
        height: 38px !important;
        text-decoration: none !important;
        cursor: pointer !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.35) !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
        white-space: nowrap !important;
    }}
    .btn-overlay-dl:hover {{
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
        border: 1px solid #7dd3fc !important;
        border-radius: 8px !important;
        box-shadow: 0 6px 18px rgba(14, 165, 233, 0.55) !important;
        color: #ffffff !important;
    }}
    
    .btn-overlay-print {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        color: #ffffff !important;
        font-family: Arial, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13.5px !important;
        border: 1px solid rgba(52, 211, 153, 0.8) !important;
        border-radius: 8px !important;
        padding: 0 16px !important;
        height: 38px !important;
        cursor: pointer !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35) !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
        outline: none !important;
        white-space: nowrap !important;
    }}
    .btn-overlay-print:hover {{
        background: linear-gradient(135deg, #0d9488 0%, #10b981 50%, #34d399 100%) !important;
        border: 1px solid #6ee7b7 !important;
        border-radius: 8px !important;
        outline: none !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.55) !important;
    }}
    
    .btn-overlay-close {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: rgba(30, 41, 59, 0.8) !important;
        color: #e2e8f0 !important;
        font-family: Arial, sans-serif !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        border: 1px solid rgba(148, 163, 184, 0.45) !important;
        border-radius: 8px !important;
        padding: 0 14px !important;
        height: 38px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        white-space: nowrap !important;
    }}
    .btn-overlay-close:hover {{
        background: rgba(51, 65, 85, 0.95) !important;
        border-radius: 8px !important;
        color: #ffffff !important;
    }}
    
    .preview-sheet-wrapper {{
        display: flex !important;
        justify-content: center !important;
        width: 100% !important;
        padding-bottom: 30px !important;
    }}
    
    .preview-paper-card {{
        width: 100% !important;
        max-width: 860px !important;
        background: #ffffff !important;
        border-radius: 6px !important;
        box-shadow: 0 16px 50px rgba(0, 0, 0, 0.8), 0 2px 12px rgba(0, 0, 0, 0.5) !important;
        border: 1px solid rgba(148, 163, 184, 0.4) !important;
        overflow: hidden !important;
    }}
    
    .preview-paper-img {{
        width: 100% !important;
        height: auto !important;
        display: block !important;
        object-fit: contain !important;
        pointer-events: none !important;
        user-select: none !important;
    }}
    
    /* Prompt Modal: Ask User to Generate Another One */
    .prompt-modal-backdrop {{
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        bottom: 0 !important;
        width: 100vw !important;
        height: 100vh !important;
        background: rgba(5, 10, 25, 0.85) !important;
        backdrop-filter: blur(8px) !important;
        -webkit-backdrop-filter: blur(8px) !important;
        z-index: 99999999 !important;
        display: none !important;
        align-items: center !important;
        justify-content: center !important;
        overflow: hidden !important;
        overscroll-behavior: none !important;
        touch-action: none !important;
        animation: promptFadeIn 0.25s ease-out !important;
    }}
    .prompt-modal-backdrop.is-open {{
        display: flex !important;
        overflow: hidden !important;
        overscroll-behavior: none !important;
        touch-action: none !important;
    }}
    
    .preview-modal-overlay.prompt-active {{
        overflow: hidden !important;
        overscroll-behavior: none !important;
        touch-action: none !important;
    }}
    
    html.prompt-active, body.prompt-active {{
        overflow: hidden !important;
        overscroll-behavior: none !important;
        touch-action: none !important;
    }}
    
    @keyframes promptFadeIn {{
        from {{ opacity: 0; }}
        to {{ opacity: 1; }}
    }}
    
    .prompt-dialog-box {{
        background: radial-gradient(circle at 50% 20%, rgba(15, 23, 42, 0.98) 0%, rgba(10, 15, 30, 0.99) 100%) !important;
        border: 1px solid rgba(56, 189, 248, 0.45) !important;
        border-radius: 14px !important;
        padding: 26px 24px 22px 24px !important;
        text-align: center !important;
        max-width: 460px !important;
        width: 90% !important;
        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.9), 0 0 30px rgba(56, 189, 248, 0.2) !important;
        animation: promptPop 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }}
    
    @keyframes promptPop {{
        from {{ transform: scale(0.92); opacity: 0; }}
        to {{ transform: scale(1); opacity: 1; }}
    }}
    
    .prompt-seal-logo {{
        width: 58px !important;
        height: 58px !important;
        border-radius: 50% !important;
        margin-bottom: 12px !important;
        box-shadow: 0 0 16px rgba(16, 185, 129, 0.4) !important;
    }}
    
    .prompt-title {{
        font-family: Arial, sans-serif !important;
        font-size: 17px !important;
        font-weight: 800 !important;
        color: #f8fafc !important;
        margin-bottom: 8px !important;
    }}
    
    .prompt-desc {{
        font-family: Arial, sans-serif !important;
        font-size: 13px !important;
        color: #94a3b8 !important;
        line-height: 1.45 !important;
        margin-bottom: 20px !important;
    }}
    
    .prompt-btn-group {{
        display: flex !important;
        gap: 10px !important;
        justify-content: center !important;
    }}
    
    .prompt-btn-primary {{
        background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        color: #ffffff !important;
        font-family: Arial, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13.5px !important;
        border: 1px solid #34d399 !important;
        border-radius: 8px !important;
        padding: 10px 18px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        outline: none !important;
    }}
    .prompt-btn-primary:hover {{
        background: linear-gradient(135deg, #0d9488 0%, #10b981 50%, #34d399 100%) !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.5) !important;
    }}
    
    .prompt-btn-secondary {{
        background: rgba(30, 41, 59, 0.8) !important;
        color: #cbd5e1 !important;
        font-family: Arial, sans-serif !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
        border: 1px solid #475569 !important;
        border-radius: 8px !important;
        padding: 10px 18px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        outline: none !important;
    }}
    .prompt-btn-secondary:hover {{
        background: rgba(51, 65, 85, 0.9) !important;
        color: #ffffff !important;
        border-radius: 8px !important;
    }}
    
    .btn-reopen-preview {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        color: #ffffff !important;
        font-family: Arial, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        border: 1px solid rgba(56, 189, 248, 0.6) !important;
        border-radius: 8px !important;
        padding: 8px 14px !important;
        cursor: pointer !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.35) !important;
        transition: all 0.2s ease !important;
        white-space: nowrap !important;
    }}
    .btn-reopen-preview:hover {{
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
        border-radius: 8px !important;
        box-shadow: 0 6px 18px rgba(14, 165, 233, 0.5) !important;
    }}
    
    /* Responsive Overrides for Preview Overlay & Confirmation Prompt */
    @media (max-width: 768px) {{
        .preview-modal-overlay {{
            padding: 54px 8px 30px 8px !important;
            width: 100vw !important;
            max-width: 100vw !important;
            overflow-x: hidden !important;
        }}
        
        .preview-top-bar {{
            padding: 10px 12px !important;
            margin-bottom: 12px !important;
            border-radius: 10px !important;
            flex-direction: column !important;
            align-items: stretch !important;
            gap: 10px !important;
            width: 100% !important;
        }}
        
        .preview-title-box {{
            width: 100% !important;
            justify-content: flex-start !important;
        }}
        
        .preview-bar-heading {{
            font-size: 12.5px !important;
        }}
        
        .preview-bar-sub {{
            font-size: 10.5px !important;
        }}
        
        .preview-actions-group {{
            display: flex !important;
            flex-wrap: wrap !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        
        .btn-overlay-dl {{
            flex: 1 1 100% !important;
            width: 100% !important;
            height: 38px !important;
            font-size: 12.5px !important;
            padding: 0 10px !important;
            justify-content: center !important;
            text-align: center !important;
        }}
        
        .btn-overlay-print {{
            flex: 1 1 calc(58% - 4px) !important;
            height: 38px !important;
            font-size: 12.5px !important;
            padding: 0 10px !important;
            justify-content: center !important;
        }}
        
        .btn-overlay-close {{
            flex: 1 1 calc(42% - 4px) !important;
            height: 38px !important;
            font-size: 12.5px !important;
            padding: 0 10px !important;
            justify-content: center !important;
        }}
        
        .preview-sheet-wrapper {{
            width: 100% !important;
            padding-bottom: 20px !important;
        }}
        
        .preview-paper-card {{
            width: 100% !important;
            max-width: 100% !important;
            border-radius: 4px !important;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.7) !important;
        }}
        
        .preview-paper-img {{
            width: 100% !important;
            height: auto !important;
            max-width: 100% !important;
        }}
        
        .prompt-dialog-box {{
            width: 92% !important;
            max-width: 95% !important;
            padding: 20px 16px 18px 16px !important;
        }}
        
        .prompt-title {{
            font-size: 15.5px !important;
        }}
        
        .prompt-desc {{
            font-size: 12px !important;
            margin-bottom: 16px !important;
        }}
        
        .prompt-btn-group {{
            flex-direction: column !important;
            gap: 8px !important;
            width: 100% !important;
        }}
        
        .prompt-btn-primary, .prompt-btn-secondary {{
            width: 100% !important;
            font-size: 13px !important;
            padding: 10px !important;
            text-align: center !important;
            justify-content: center !important;
        }}
        
        .btn-reopen-preview {{
            width: 100% !important;
            font-size: 12.5px !important;
            justify-content: center !important;
            margin-top: 6px !important;
        }}
    }}
    
    @media (max-width: 480px) {{
        .preview-modal-overlay {{
            padding: 50px 6px 24px 6px !important;
        }}
        .preview-top-bar {{
            padding: 8px 10px !important;
            gap: 8px !important;
        }}
        .preview-bar-logo {{
            width: 28px !important;
            height: 28px !important;
        }}
        .preview-bar-heading {{
            font-size: 11.5px !important;
            letter-spacing: 0.3px !important;
        }}
        .preview-bar-sub {{
            font-size: 10px !important;
        }}
        .btn-overlay-dl {{
            font-size: 12px !important;
            height: 36px !important;
        }}
        .btn-overlay-print, .btn-overlay-close {{
            font-size: 12px !important;
            height: 36px !important;
        }}
    }}
    
    .gen-logo-wrapper {{
        position: relative;
        width: 76px;
        height: 76px;
        margin: 0 auto 12px auto;
        display: flex;
        align-items: center;
        justify-content: center;
    }}
    
    .gen-halo-ring {{
        position: absolute;
        inset: -6px;
        border-radius: 50%;
        pointer-events: none;
    }}
    
    .halo-anim {{
        background: conic-gradient(from 0deg, rgba(16, 185, 129, 0.9), rgba(6, 182, 212, 0.9), rgba(59, 130, 246, 0.9), rgba(16, 185, 129, 0.9));
        animation: haloSpin 2.2s linear infinite;
        filter: blur(5px);
        opacity: 0.85;
    }}
    
    .halo-finished {{
        background: radial-gradient(circle, rgba(16, 185, 129, 0.95) 0%, rgba(5, 150, 105, 0.5) 65%, transparent 100%);
        filter: blur(7px);
        animation: finishPulse 1.2s ease-in-out infinite alternate;
    }}
    
    @keyframes haloSpin {{
        0% {{ transform: rotate(0deg); }}
        100% {{ transform: rotate(360deg); }}
    }}
    
    @keyframes finishPulse {{
        0% {{ transform: scale(0.95); opacity: 0.75; }}
        100% {{ transform: scale(1.18); opacity: 1; }}
    }}
    
    .loading-logo {{
        position: relative;
        z-index: 2;
        width: 70px;
        height: 70px;
        border-radius: 50%;
        object-fit: contain;
        background: rgba(15, 23, 42, 0.92);
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.6);
    }}
    
    .logo-pulsing {{
        animation: logoBreathe 1.8s ease-in-out infinite alternate;
    }}
    
    .logo-finished {{
        animation: logoCelebrate 0.55s ease-out forwards;
    }}
    
    @keyframes logoBreathe {{
        0% {{ transform: scale(0.95); filter: drop-shadow(0 0 4px rgba(56, 189, 248, 0.4)); }}
        100% {{ transform: scale(1.05); filter: drop-shadow(0 0 14px rgba(16, 185, 129, 0.85)); }}
    }}
    
    @keyframes logoCelebrate {{
        0% {{ transform: scale(0.95); }}
        50% {{ transform: scale(1.14); filter: drop-shadow(0 0 20px rgba(52, 211, 153, 1)); }}
        100% {{ transform: scale(1.03); filter: drop-shadow(0 0 12px rgba(52, 211, 153, 0.85)); }}
    }}
    
    .gen-header-title {{
        font-family: Arial, sans-serif;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 1.8px;
        color: #94a3b8;
        text-transform: uppercase;
        margin-bottom: 3px;
    }}
    
    .gen-header-sub {{
        font-family: Arial, sans-serif;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: 0.9px;
        color: #f1f5f9;
        text-transform: uppercase;
        margin-bottom: 12px;
    }}
    
    .gen-status-box {{
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(56, 189, 248, 0.25);
        border-radius: 8px;
        padding: 9px 14px;
        margin: 0 auto 14px auto;
        max-width: 460px;
    }}
    
    .gen-status-text {{
        font-family: Arial, sans-serif;
        font-size: 13.5px;
        font-weight: 700;
        letter-spacing: 0.4px;
        margin-bottom: 3px;
    }}
    
    .gen-subtitle-text {{
        font-family: Arial, sans-serif;
        font-size: 11.5px;
        color: #94a3b8;
        line-height: 1.35;
    }}
    
    .gen-progress-wrapper {{
        display: flex;
        align-items: center;
        gap: 12px;
        max-width: 500px;
        margin: 0 auto 14px auto;
    }}
    
    .gen-progress-track {{
        flex: 1;
        height: 11px;
        background: rgba(30, 41, 59, 0.95);
        border-radius: 999px;
        overflow: hidden;
        border: 1px solid rgba(71, 85, 105, 0.6);
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.5);
        position: relative;
    }}
    
    .gen-progress-bar {{
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #059669 0%, #10b981 50%, #06b6d4 100%);
        box-shadow: 0 0 10px rgba(16, 185, 129, 0.6);
        transition: width 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
        overflow: hidden;
    }}
    
    .bar-finished {{
        background: linear-gradient(90deg, #10b981 0%, #34d399 100%) !important;
        box-shadow: 0 0 16px rgba(52, 211, 153, 0.95) !important;
    }}
    
    .gen-shimmer {{
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(90deg, transparent 0%, rgba(255, 255, 255, 0.4) 50%, transparent 100%);
        animation: progressShimmer 1.4s infinite;
        transform: translateX(-100%);
    }}
    
    @keyframes progressShimmer {{
        100% {{ transform: translateX(100%); }}
    }}
    
    .gen-pct-label {{
        font-family: Arial, sans-serif;
        font-size: 13.5px;
        font-weight: 800;
        color: #38bdf8;
        min-width: 44px;
        text-align: right;
    }}
    
    .gen-badges-row {{
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 6px;
        max-width: 500px;
        margin: 0 auto;
    }}
    
    .step-badge {{
        font-family: Arial, sans-serif;
        font-size: 11px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 999px;
        transition: all 0.25s ease;
    }}
    
    .badge-done {{
        background: rgba(16, 185, 129, 0.22);
        color: #34d399;
        border: 1px solid rgba(16, 185, 129, 0.5);
        box-shadow: 0 0 8px rgba(16, 185, 129, 0.2);
    }}
    
    .badge-active {{
        background: rgba(6, 182, 212, 0.25);
        color: #38bdf8;
        border: 1px solid rgba(56, 189, 248, 0.65);
        box-shadow: 0 0 10px rgba(56, 189, 248, 0.35);
        animation: badgePulse 1.2s ease-in-out infinite alternate;
    }}
    
    @keyframes badgePulse {{
        0% {{ transform: scale(0.97); }}
        100% {{ transform: scale(1.03); }}
    }}
    
    .badge-pending {{
        background: rgba(30, 41, 59, 0.55);
        color: #64748b;
        border: 1px solid rgba(51, 65, 85, 0.5);
    }}
    
    @media (max-width: 480px) {{
        .gen-overlay-content {{
            max-width: 95% !important;
        }}
        .gen-logo-wrapper {{
            width: 60px;
            height: 60px;
        }}
        .loading-logo {{
            width: 56px;
            height: 56px;
        }}
        .gen-header-title {{
            font-size: 9.5px;
        }}
        .gen-header-sub {{
            font-size: 12px;
        }}
        .gen-status-text {{
            font-size: 12px;
        }}
        .step-badge {{
            font-size: 10px;
            padding: 3px 7px;
        }}
    }}

</style>
"""
st.markdown(css_rules, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def render_generation_progress_card(pct: int, status_text: str, subtitle: str, current_step: int, finished: bool = False) -> str:
    """Generates the frameless, 80% semi-transparent overlay loading screen centered on the main page."""
    steps_data = [
        ("1. Data Check", 1),
        ("2. Word .docx", 2),
        ("3. PDF Engine", 3),
        ("4. 1:1 Preview", 4)
    ]
    
    badges_html = []
    for label, s_num in steps_data:
        if finished or s_num < current_step:
            cls = "step-badge badge-done"
            prefix = "✓ "
        elif s_num == current_step:
            cls = "step-badge badge-active"
            prefix = ""
        else:
            cls = "step-badge badge-pending"
            prefix = ""
        badges_html.append(f'<span class="{cls}">{prefix}{label}</span>')
    
    badges_row = "".join(badges_html)
    logo_cls = "loading-logo logo-finished" if finished else "loading-logo logo-pulsing"
    status_color = "#34d399" if finished else "#38bdf8"
    
    return f"""<div class="gen-modal-overlay">
<div class="gen-overlay-content">
<div class="gen-logo-wrapper">
<div class="gen-halo-ring {'halo-finished' if finished else 'halo-anim'}"></div>
<img src="data:image/png;base64,{thumb_b64}" class="{logo_cls}" alt="Municipality of Maragondon" />
</div>
<div class="gen-header-title">MUNICIPALITY OF MARAGONDON</div>
<div class="gen-header-sub">OFFICIAL ABSTRACT OF CANVASS ENGINE</div>
<div class="gen-status-box">
<div class="gen-status-text" style="color: {status_color};">{status_text}</div>
<div class="gen-subtitle-text">{subtitle}</div>
</div>
<div class="gen-progress-wrapper">
<div class="gen-progress-track">
<div class="gen-progress-bar {'bar-finished' if finished else ''}" style="width: {pct}%;">
<div class="gen-shimmer"></div>
</div>
</div>
<div class="gen-pct-label">{pct}%</div>
</div>
<div class="gen-badges-row">{badges_row}</div>
</div>
</div>"""

def sync_rfq_date_from_ac():
    """Update rfq_closing_date whenever user edits month or year in Abstract Number."""
    month_val = str(st.session_state.get("ac_month", "")).strip()
    year_val = str(st.session_state.get("ac_year", "")).strip()
    curr_date = st.session_state.get("rfq_closing_date")
    if not isinstance(curr_date, datetime.date):
        curr_date = datetime.date(2026, 10, 26)
    
    try:
        m = int(month_val) if month_val else curr_date.month
        if not (1 <= m <= 12):
            return
            
        if year_val:
            y_int = int(year_val)
            y = 2000 + y_int if y_int < 100 else y_int
        else:
            y = curr_date.year
            
        max_day = calendar.monthrange(y, m)[1]
        day = min(curr_date.day, max_day)
        
        st.session_state["rfq_closing_date"] = datetime.date(y, m, day)
        st.session_state["ac_month"] = f"{m:02d}"
        st.session_state["ac_year"] = f"{y % 100:02d}"
    except Exception:
        pass

def sync_ac_from_rfq_date():
    """Update month and year in Abstract Number whenever user selects a date on the calendar."""
    selected_date = st.session_state.get("rfq_closing_date")
    if isinstance(selected_date, datetime.date):
        st.session_state["ac_month"] = f"{selected_date.month:02d}"
        st.session_state["ac_year"] = f"{selected_date.year % 100:02d}"

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

def clean_purpose_description(text: str) -> str:
    """Trim boilerplate prefixes and location suffixes from user entered description."""
    if not text:
        return ""
    cleaned = text.strip()
    # Strip any accidental prefix user might paste (e.g. FOR: PAYMENT FOR THE or PAYMENT FOR THE or FOR:)
    cleaned = re.sub(r'^(?:FOR:\s*)?PAYMENT\s+FOR\s+THE\s*|^FOR:\s*', '', cleaned, flags=re.IGNORECASE).strip()
    # Strip any accidental trailing location suffix so it doesn't duplicate
    cleaned = re.sub(r'[\s,.-]*IN\s+LGU[- ]*MARAGONDON[,\s]*CAVITE\.?$', '', cleaned, flags=re.IGNORECASE).strip()
    return cleaned.strip(" ,.-")

def generate_pdf_from_docx_bytes(docx_bytes: bytes) -> bytes:
    """Convert generated docx bytes to PDF bytes using LibreOffice (Linux/Cloud) or Word COM (Windows)."""
    if not docx_bytes:
        return b""
    import tempfile, time, subprocess, shutil, sys
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_docx = os.path.join(tmp_dir, "document.docx")
        tmp_pdf = os.path.join(tmp_dir, "document.pdf")
        with open(tmp_docx, "wb") as f:
            f.write(docx_bytes)
        
        # 1. Try LibreOffice / soffice (Standard on Linux & Streamlit Cloud)
        soffice_bin = shutil.which("libreoffice") or shutil.which("soffice")
        if soffice_bin:
            try:
                env = os.environ.copy()
                env["HOME"] = tmp_dir
                profile_dir = os.path.join(tmp_dir, "lo_profile")
                os.makedirs(profile_dir, exist_ok=True)
                profile_uri = "file:///" + profile_dir.replace("\\", "/").lstrip("/")
                cmd = [
                    soffice_bin,
                    "--headless",
                    "--invisible",
                    "--nologo",
                    "--nodefault",
                    "--nofirststartwizard",
                    "--norestore",
                    f"-env:UserInstallation={profile_uri}",
                    "--convert-to", "pdf",
                    "--outdir", tmp_dir,
                    os.path.abspath(tmp_docx)
                ]
                proc = subprocess.run(cmd, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
                if proc.returncode != 0:
                    print("LibreOffice error:", proc.stderr.decode("utf-8", errors="ignore"))
                for _ in range(25):
                    if os.path.exists(tmp_pdf) and os.path.getsize(tmp_pdf) > 0:
                        break
                    time.sleep(0.1)
                if os.path.exists(tmp_pdf) and os.path.getsize(tmp_pdf) > 0:
                    with open(tmp_pdf, "rb") as f:
                        return f.read()
            except Exception as e:
                print("LibreOffice conversion failed:", repr(e))

        # 2. Try Windows MS Word COM (on local Windows)
        if sys.platform == "win32":
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                word = None
                doc = None
                try:
                    try:
                        word = win32com.client.DispatchEx("Word.Application")
                    except Exception:
                        word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = False
                    doc = word.Documents.Open(os.path.abspath(tmp_docx), ReadOnly=True)
                    try:
                        doc.SaveAs2(os.path.abspath(tmp_pdf), FileFormat=17)
                    except Exception:
                        doc.SaveAs(os.path.abspath(tmp_pdf), FileFormat=17)
                    doc.Close(False)
                    doc = None
                finally:
                    if doc:
                        try:
                            doc.Close(False)
                        except Exception:
                            pass
                    if word:
                        try:
                            word.Quit()
                        except Exception:
                            pass
                    pythoncom.CoUninitialize()
                
                for _ in range(20):
                    if os.path.exists(tmp_pdf) and os.path.getsize(tmp_pdf) > 0:
                        break
                    time.sleep(0.1)
                    
                if os.path.exists(tmp_pdf) and os.path.getsize(tmp_pdf) > 0:
                    with open(tmp_pdf, "rb") as f:
                        return f.read()
            except Exception as e:
                print("Windows Word COM conversion failed:", repr(e))
            
    return b""

def build_html_document_fallback(ctx: dict, logo_b64: str) -> str:
    """Build a 1:1 crisp HTML paper document layout fallback when PDF conversion is unavailable."""
    if not ctx:
        return ""
    b1_name = ctx.get("bidder_1_name", "")
    b1_loc = ctx.get("bidder_1_location", "")
    b1_amt = ctx.get("bidder_1_amount", "")
    b2_name = ctx.get("bidder_2_name", "")
    b2_loc = ctx.get("bidder_2_location", "")
    b2_amt = ctx.get("bidder_2_amount", "")
    b3_name = ctx.get("bidder_3_name", "")
    b3_loc = ctx.get("bidder_3_location", "")
    b3_amt = ctx.get("bidder_3_amount", "")
    winner_name = ctx.get("winning_bidder_name", b1_name)
    winner_loc = ctx.get("winning_bidder_location", b1_loc)
    logo_tag = f'<img src="data:image/png;base64,{logo_b64}" style="width: 72px; height: 72px; object-fit: contain; margin-bottom: 4px;" alt="Seal" />' if logo_b64 else ''
    
    html = f"""<div class="doc-paper-fallback" style="background:#ffffff;color:#000000;font-family:'Times New Roman',Times,serif;padding:36px 42px;box-sizing:border-box;width:100%;max-width:820px;margin:0 auto;line-height:1.35;box-shadow:0 4px 20px rgba(0,0,0,0.5);border-radius:2px;">
<div style="text-align:center;margin-bottom:14px;">
{logo_tag}
<div style="font-size:13px;">Republic of the Philippines</div>
<div style="font-size:13px;">Province of Cavite</div>
<div style="font-size:14px;font-weight:bold;letter-spacing:0.5px;">MUNICIPALITY OF MARAGONDON</div>
<div style="font-size:17px;font-weight:bold;margin-top:14px;text-decoration:underline;letter-spacing:1px;">ABSTRACT OF CANVASS</div>
</div>
<div style="font-size:13px;margin-bottom:12px;font-weight:bold;">FOR: PAYMENT FOR THE {ctx.get('purpose', '')}</div>
<div style="text-align:right;font-size:12.5px;margin-bottom:16px;line-height:1.4;">
<div>Opened at the Office of the Municipal</div>
<div>Mayor on <u>{ctx.get('opened_date', '')}</u></div>
</div>
<table style="width:100%;border-collapse:collapse;border:1.5px solid #000000;margin-bottom:18px;font-size:12.5px;">
<thead>
<tr style="background:#f8fafc;font-weight:bold;text-align:center;">
<th style="border:1px solid #000000;padding:8px 10px;width:62%;">NAME OF BIDDERS</th>
<th style="border:1px solid #000000;padding:8px 10px;width:38%;">AMOUNT OF BID</th>
</tr>
</thead>
<tbody>
<tr>
<td style="border:1px solid #000000;padding:8px 10px;vertical-align:top;"><strong>1. {b1_name}</strong><br><span style="font-size:11.5px;padding-left:14px;color:#222;">{b1_loc}</span></td>
<td style="border:1px solid #000000;padding:8px 10px;text-align:right;vertical-align:top;font-weight:bold;">{b1_amt}</td>
</tr>
<tr>
<td style="border:1px solid #000000;padding:8px 10px;vertical-align:top;"><strong>2. {b2_name}</strong><br><span style="font-size:11.5px;padding-left:14px;color:#222;">{b2_loc}</span></td>
<td style="border:1px solid #000000;padding:8px 10px;text-align:right;vertical-align:top;font-weight:bold;">{b2_amt}</td>
</tr>
<tr>
<td style="border:1px solid #000000;padding:8px 10px;vertical-align:top;"><strong>3. {b3_name}</strong><br><span style="font-size:11.5px;padding-left:14px;color:#222;">{b3_loc}</span></td>
<td style="border:1px solid #000000;padding:8px 10px;text-align:right;vertical-align:top;font-weight:bold;">{b3_amt}</td>
</tr>
</tbody>
</table>
<div style="font-size:12.5px;margin-bottom:22px;text-align:justify;line-height:1.45;">
Awarded to the lowest bidder <strong>{winner_name}</strong>, <strong>{winner_loc}</strong> same being reasonable and most advantageous to the government.
</div>
<div style="margin-top:14px;margin-bottom:24px;">
<div style="text-align:center;font-size:12px;font-weight:bold;margin-bottom:18px;letter-spacing:0.5px;">COMMITTEE ON AWARDS</div>
<div style="display:flex;justify-content:space-around;text-align:center;font-size:11px;margin-bottom:16px;">
<div>
<div style="font-weight:bold;border-bottom:1px solid #000;padding-bottom:2px;">ENGR. ALDWIN JESTER A. ANCAYAN</div>
<div style="font-size:10px;color:#333;">BAC CHAIRMAN/MUN. ENGINEER</div>
</div>
<div>
<div style="font-weight:bold;border-bottom:1px solid #000;padding-bottom:2px;">CHERRYL I. ESQUILLO</div>
<div style="font-size:10px;color:#333;">BAC VICE-CHAIRMAN/MUN. BUDGET OFFICER</div>
</div>
</div>
<div style="display:flex;justify-content:space-around;text-align:center;font-size:11px;margin-bottom:16px;">
<div>
<div style="font-weight:bold;border-bottom:1px solid #000;padding-bottom:2px;">REGINO S. REYES</div>
<div style="font-size:10px;color:#333;">BAC MEMBER/MUN. AGRICULTURIST</div>
</div>
<div>
<div style="font-weight:bold;border-bottom:1px solid #000;padding-bottom:2px;">GERARDO Z. RIELO JR.</div>
<div style="font-size:10px;color:#333;">BAC MEMBER/MDRRMO</div>
</div>
</div>
<div style="text-align:center;font-size:11px;">
<div style="font-weight:bold;border-bottom:1px solid #000;display:inline-block;padding-bottom:2px;">ENGR. ENP. PERCIVAL M. MALIMBAN</div>
<div style="font-size:10px;color:#333;">BAC MEMBER/MUN. PLANNING &amp; DEV’T COORDINATOR</div>
</div>
</div>
<div style="font-size:12px;font-weight:bold;color:#1e293b;border-top:1px dashed #cbd5e1;padding-top:8px;">
{ctx.get('ac_number', '')}
</div>
</div>"""
    return "".join(l.strip() for l in html.splitlines() if l.strip())

def render_pdf_to_preview_png(pdf_bytes: bytes) -> bytes:
    """Render the first page of PDF bytes to a crisp, high-DPI PNG image."""
    if not pdf_bytes:
        return b""
    try:
        try:
            import fitz
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        except Exception:
            import pymupdf
            doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        if len(doc) > 0:
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            return pix.tobytes("png")
    except Exception as e:
        print("render_pdf_to_preview_png error:", repr(e))
    return b""

# -----------------------------------------------------------------------------
# Session State Initialization
# -----------------------------------------------------------------------------
if "form_data" not in st.session_state:
    st.session_state.form_data = DEFAULT_SAMPLE_DATA.copy()

if "generated_doc_bytes" not in st.session_state:
    st.session_state.generated_doc_bytes = None

if "generated_pdf_bytes" not in st.session_state:
    st.session_state.generated_pdf_bytes = None

if "generated_preview_png_bytes" not in st.session_state:
    st.session_state.generated_preview_png_bytes = None

if "generated_context" not in st.session_state:
    st.session_state.generated_context = None

if "rfq_closing_date" not in st.session_state:
    st.session_state["rfq_closing_date"] = datetime.date(2026, 10, 26)

if "ac_code" not in st.session_state:
    st.session_state["ac_code"] = "0433H"

if "ac_month" not in st.session_state:
    st.session_state["ac_month"] = "10"

if "ac_year" not in st.session_state:
    st.session_state["ac_year"] = "26"

if "purpose_desc" not in st.session_state:
    st.session_state["purpose_desc"] = "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD"

if "include_lgu_suffix" not in st.session_state:
    st.session_state["include_lgu_suffix"] = True

HISTORY_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "canvass_history.txt")

def init_history_db_file():
    """Ensure canvass_history.txt exists with default sample entry if empty."""
    need_init = False
    if not os.path.exists(HISTORY_FILE_PATH):
        need_init = True
    else:
        try:
            with open(HISTORY_FILE_PATH, "r", encoding="utf-8") as f:
                c = f.read().strip()
            if not c or ("Abstract Number:" not in c and "# USER_CLEARED_DATABASE" not in c):
                need_init = True
        except Exception:
            need_init = True

    if need_init:
        try:
            with open(HISTORY_FILE_PATH, "w", encoding="utf-8") as f:
                f.write("# ==============================================================================\n")
                f.write("# ABSTRACT OF CANVASS DATABASE HISTORY\n")
                f.write("# Each record below is saved automatically upon document generation.\n")
                f.write("# ==============================================================================\n\n")
            sample_entry = {
                "record_id": "REC_INITIAL_0443H",
                "ac_number": "AC-0443H-10-26",
                "ac_code": "0443H",
                "ac_month": "10",
                "ac_year": "26",
                "rfq_closing_date": "2026-10-26",
                "opened_date": "October 26, 2026",
                "winning_bidder": "RBK PRINTING SERVICES",
                "winning_location": "MARAGONDON, CAVITE",
                "winning_amount": "1,000.00",
                "purpose_desc": "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD",
                "include_lgu_suffix": True,
                "purpose": "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD IN LGU MARAGONDON, CAVITE.",
                "b1_name": "RBK PRINTING SERVICES",
                "b1_loc": "MARAGONDON, CAVITE",
                "b1_amt": "1,000.00",
                "b2_name": "SECOND BIDDER",
                "b2_loc": "MARAGONDON, CAVITE",
                "b2_amt": "1,200.00",
                "b3_name": "THIRD BIDDER",
                "b3_loc": "MARAGONDON, CAVITE",
                "b3_amt": "1,500.00",
                "time_str": "Initial"
            }
            write_single_entry_to_txt(sample_entry)
        except Exception as e:
            print("Error initializing history file:", e)

def format_entry_to_txt_block(entry: dict) -> str:
    """Format history dictionary into human-readable .txt database format."""
    rec_id = entry.get("record_id") or f"REC_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{entry.get('ac_code', 'AC')}"
    entry["record_id"] = rec_id
    
    rfq_date_val = entry.get("rfq_closing_date", "")
    if isinstance(rfq_date_val, datetime.date):
        rfq_date_str = rfq_date_val.strftime("%B %d, %Y")
    else:
        rfq_date_str = str(rfq_date_val)
        
    b1_name = entry.get("winning_bidder") or entry.get("b1_name") or "—"
    amt = entry.get("winning_amount") or entry.get("b1_amt") or "0.00"
    clean_amt = re.sub(r"[^\d.,]", "", str(amt)).strip()
    
    payload_dict = {}
    for k, v in entry.items():
        if isinstance(v, (datetime.date, datetime.datetime)):
            payload_dict[k] = v.isoformat()
        elif k != "doc_bytes":
            payload_dict[k] = v
            
    payload_json = json.dumps(payload_dict, ensure_ascii=False)
    
    block = [
        "================================================================================",
        f"Record ID: {rec_id}",
        f"Abstract Number: {entry.get('ac_number', '')}",
        f"RFQ Closing Date: {rfq_date_str}",
        f"Bidder 1: {b1_name}",
        f"Amount: {clean_amt}",
        "Extra Information needed for loading into form:",
        f"Purpose: {entry.get('purpose_desc', entry.get('purpose', ''))}",
        f"Include LGU Suffix: {entry.get('include_lgu_suffix', True)}",
        f"Bidder 1 Location: {entry.get('b1_loc', entry.get('winning_location', ''))}",
        f"Bidder 2: {entry.get('b2_name', '')} | {entry.get('b2_loc', '')} | {entry.get('b2_amt', '')}",
        f"Bidder 3: {entry.get('b3_name', '')} | {entry.get('b3_loc', '')} | {entry.get('b3_amt', '')}",
        f"Timestamp: {entry.get('time_str', datetime.datetime.now().strftime('%b %d, %Y %I:%M %p'))}",
        f"Data JSON: {payload_json}",
        "================================================================================\n"
    ]
    return "\n".join(block)

def write_single_entry_to_txt(entry: dict):
    """Append a single record block to canvass_history.txt."""
    block_str = format_entry_to_txt_block(entry)
    with open(HISTORY_FILE_PATH, "a", encoding="utf-8") as f:
        f.write(block_str)

def read_all_history_from_txt() -> list:
    """Read and parse all history entries from canvass_history.txt (newest first)."""
    init_history_db_file()
        
    records = []
    try:
        with open(HISTORY_FILE_PATH, "r", encoding="utf-8") as f:
            content = f.read()
            
        blocks = re.split(r'\n={10,}\n', content)
        for raw_b in blocks:
            raw_b = raw_b.strip()
            if not raw_b or "Abstract Number:" not in raw_b:
                continue
            
            entry = {}
            json_m = re.search(r"Data JSON:\s*(\{.*\})", raw_b)
            if json_m:
                try:
                    entry = json.loads(json_m.group(1))
                except Exception:
                    entry = {}
                    
            for line in raw_b.splitlines():
                line = line.strip()
                if line.startswith("Record ID:") and "record_id" not in entry:
                    entry["record_id"] = line.split(":", 1)[1].strip()
                elif line.startswith("Abstract Number:") and "ac_number" not in entry:
                    entry["ac_number"] = line.split(":", 1)[1].strip()
                elif line.startswith("RFQ Closing Date:") and "rfq_closing_date" not in entry:
                    entry["rfq_closing_date"] = line.split(":", 1)[1].strip()
                elif line.startswith("Bidder 1:") and "winning_bidder" not in entry:
                    entry["winning_bidder"] = line.split(":", 1)[1].strip()
                elif line.startswith("Amount:") and "winning_amount" not in entry:
                    entry["winning_amount"] = line.split(":", 1)[1].strip()
                elif line.startswith("Purpose:") and "purpose_desc" not in entry:
                    entry["purpose_desc"] = line.split(":", 1)[1].strip()
                elif line.startswith("Timestamp:") and "time_str" not in entry:
                    entry["time_str"] = line.split(":", 1)[1].strip()
                    
            if entry.get("ac_number") or entry.get("record_id"):
                if not entry.get("record_id"):
                    entry["record_id"] = f"REC_{len(records)+1}"
                records.append(entry)
    except Exception as e:
        print("Error reading canvass_history.txt:", e)
        
    return list(reversed(records))

def delete_entry_from_txt(record_id: str):
    """Delete a record from canvass_history.txt by record_id."""
    if not os.path.exists(HISTORY_FILE_PATH):
        return
    try:
        all_entries = read_all_history_from_txt()
        remaining = [e for e in all_entries if e.get("record_id") != record_id]
        with open(HISTORY_FILE_PATH, "w", encoding="utf-8") as f:
            f.write("# ==============================================================================\n")
            f.write("# ABSTRACT OF CANVASS DATABASE HISTORY\n")
            f.write("# Each record below is saved automatically upon document generation.\n")
            f.write("# ==============================================================================\n\n")
            for e in reversed(remaining):
                f.write(format_entry_to_txt_block(e))
    except Exception as e:
        print("Error deleting record from canvass_history.txt:", e)

def clear_all_history_txt():
    """Clear all records from canvass_history.txt."""
    try:
        with open(HISTORY_FILE_PATH, "w", encoding="utf-8") as f:
            f.write("# ==============================================================================\n")
            f.write("# ABSTRACT OF CANVASS DATABASE HISTORY\n")
            f.write("# Each record below is saved automatically upon document generation.\n")
            f.write("# ==============================================================================\n")
            f.write("# USER_CLEARED_DATABASE\n\n")
    except Exception as e:
        print("Error clearing canvass_history.txt:", e)

# Always load the persistent history from canvass_history.txt on every page run
st.session_state.abstract_history = read_all_history_from_txt()

def load_sample_data():
    st.session_state["rfq_closing_date"] = datetime.date(2026, 10, 26)
    st.session_state["ac_code"] = "0433H"
    st.session_state["ac_month"] = "10"
    st.session_state["ac_year"] = "26"
    st.session_state["purpose_desc"] = "PURCHASE OF ID LACE FOR GIP"
    st.session_state["include_lgu_suffix"] = True
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
    st.session_state.generated_pdf_bytes = None
    st.session_state.generated_preview_png_bytes = None
    st.session_state.generated_context = None
    st.toast("Sample data loaded.")

def clear_form_data():
    today = datetime.date.today()
    st.session_state["rfq_closing_date"] = today
    st.session_state["ac_code"] = ""
    st.session_state["ac_month"] = f"{today.month:02d}"
    st.session_state["ac_year"] = f"{today.year % 100:02d}"
    st.session_state["purpose_desc"] = ""
    st.session_state["include_lgu_suffix"] = True
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
    st.session_state.generated_pdf_bytes = None
    st.session_state.generated_preview_png_bytes = None
    st.session_state.generated_context = None
    st.toast("Form cleared.")

def load_history_entry(entry):
    """Load a past history entry from database into the form."""
    raw_ac = entry.get("ac_number", "")
    m = re.search(r"AC\s*[-–]\s*([A-Za-z0-9]+)\s*[-–]\s*(\d{1,2})\s*[-–]\s*(\d{2,4})", raw_ac)
    if m:
        st.session_state["ac_code"] = m.group(1)
        st.session_state["ac_month"] = m.group(2).zfill(2)
        yr = m.group(3)
        st.session_state["ac_year"] = yr[-2:] if len(yr) >= 2 else yr
    else:
        st.session_state["ac_code"] = entry.get("ac_code", "")
        st.session_state["ac_month"] = str(entry.get("ac_month", "10")).zfill(2)
        st.session_state["ac_year"] = str(entry.get("ac_year", "26"))
        
    rfq_d = entry.get("rfq_closing_date")
    if isinstance(rfq_d, datetime.date):
        st.session_state["rfq_closing_date"] = rfq_d
    elif isinstance(rfq_d, str):
        try:
            st.session_state["rfq_closing_date"] = datetime.date.fromisoformat(rfq_d)
        except Exception:
            try:
                st.session_state["rfq_closing_date"] = datetime.datetime.strptime(rfq_d, "%B %d, %Y").date()
            except Exception:
                try:
                    m_val = int(st.session_state["ac_month"])
                    y_val = int(st.session_state["ac_year"])
                    y_full = 2000 + y_val if y_val < 100 else y_val
                    st.session_state["rfq_closing_date"] = datetime.date(y_full, m_val, 1)
                except Exception:
                    pass
        
    st.session_state["purpose_desc"] = entry.get("purpose_desc", entry.get("purpose", ""))
    st.session_state["include_lgu_suffix"] = entry.get("include_lgu_suffix", True)
    
    st.session_state["b1_name"] = entry.get("b1_name") or entry.get("winning_bidder") or ""
    st.session_state["b1_loc"] = entry.get("b1_loc") or entry.get("winning_location") or ""
    st.session_state["b1_amt"] = entry.get("b1_amt") or entry.get("winning_amount") or ""
    
    st.session_state.form_data["bidder_1_name"] = st.session_state["b1_name"]
    st.session_state.form_data["bidder_1_location"] = st.session_state["b1_loc"]
    st.session_state.form_data["bidder_1_amount"] = st.session_state["b1_amt"]
    
    if "b2_name" in entry:
        st.session_state["b2_name"] = entry.get("b2_name", "")
        st.session_state["b2_loc"] = entry.get("b2_loc", "")
        st.session_state["b2_amt"] = entry.get("b2_amt", "")
        st.session_state.form_data["bidder_2_name"] = st.session_state["b2_name"]
        st.session_state.form_data["bidder_2_location"] = st.session_state["b2_loc"]
        st.session_state.form_data["bidder_2_amount"] = st.session_state["b2_amt"]
        
    if "b3_name" in entry:
        st.session_state["b3_name"] = entry.get("b3_name", "")
        st.session_state["b3_loc"] = entry.get("b3_loc", "")
        st.session_state["b3_amt"] = entry.get("b3_amt", "")
        st.session_state.form_data["bidder_3_name"] = st.session_state["b3_name"]
        st.session_state.form_data["bidder_3_location"] = st.session_state["b3_loc"]
        st.session_state.form_data["bidder_3_amount"] = st.session_state["b3_amt"]
        
    st.session_state.generated_doc_bytes = None
    st.session_state.generated_pdf_bytes = None
    st.session_state.generated_preview_png_bytes = None
    st.session_state.generated_context = None
    st.toast(f"Loaded {raw_ac or entry.get('ac_code', '')} into form.")

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
# Sidebar: Persistent Database History & Compact Action Controls
# -----------------------------------------------------------------------------
@st.dialog("Clear Database Confirmation")
def clear_db_pin_dialog():
    st.markdown("""
    <div style="font-size: 13.5px; color: #f1f5f9; margin-bottom: 14px; line-height: 1.5;">
        This action will permanently delete all records from <code>canvass_history.txt</code>.<br><br>
        Please enter the <strong>4-digit Master PIN Code</strong> to proceed:
    </div>
    """, unsafe_allow_html=True)
    pin = st.text_input("Master PIN", type="password", max_chars=4, placeholder="••••", key="master_pin_modal_input")
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Confirm Clear", type="primary", use_container_width=True):
            if pin.strip() == "2025":
                clear_all_history_txt()
                st.session_state.abstract_history = []
                st.session_state.history_expanded_ids = set()
                st.toast("Database history cleared successfully.")
                st.rerun()
            else:
                st.error("Incorrect Master PIN Code. Access denied.")
    with col2:
        if st.button("Cancel", use_container_width=True):
            st.rerun()

if "history_expanded_ids" not in st.session_state:
    st.session_state.history_expanded_ids = set()

with st.sidebar:
    st.markdown("<div class='sidebar-heading'>Abstract History (Database)</div>", unsafe_allow_html=True)
    
    col_h_clear, col_h_ref = st.columns([1, 1])
    with col_h_clear:
        if st.button("Clear DB", key="btn_clear_db", use_container_width=True, help="Clear all records from canvass_history.txt (Requires Master PIN)"):
            clear_db_pin_dialog()
    with col_h_ref:
        if st.button("Refresh", key="btn_ref_db", use_container_width=True, help="Reload records from canvass_history.txt"):
            st.session_state.abstract_history = read_all_history_from_txt()
            st.toast("History refreshed from database.")
            st.rerun()
            
    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
    
    history_items = st.session_state.abstract_history
    if not history_items:
        st.caption("No abstracts recorded yet in database.")
    else:
        PAGE_SIZE = 5
        total_records = len(history_items)
        total_pages = max(1, (total_records + PAGE_SIZE - 1) // PAGE_SIZE)
        
        if "history_page" not in st.session_state:
            st.session_state.history_page = 1
        if st.session_state.history_page > total_pages:
            st.session_state.history_page = total_pages
        if st.session_state.history_page < 1:
            st.session_state.history_page = 1
            
        current_page = st.session_state.history_page
        start_idx = (current_page - 1) * PAGE_SIZE
        end_idx = min(start_idx + PAGE_SIZE, total_records)
        page_records = history_items[start_idx:end_idx]
        
        col_cap, col_pg_info = st.columns([1.2, 1.0])
        with col_cap:
            st.caption(f"{total_records} record(s) in DB")
        with col_pg_info:
            st.caption(f"Page {current_page} of {total_pages}")
            
        for idx_offset, item in enumerate(page_records):
            idx = start_idx + idx_offset
            rec_id = item.get("record_id", f"REC_{idx}")
            ac_num = item.get("ac_number", "—")
            winner = item.get("winning_bidder") or item.get("b1_name") or "—"
            raw_amt = item.get("winning_amount") or item.get("b1_amt") or ""
            clean_amt = re.sub(r"[^\d.,]", "", str(raw_amt)).strip()
            amount_display = f"₱ {clean_amt}" if clean_amt else "—"
            time_tag = item.get("time_str", "")
            
            is_open = rec_id in st.session_state.history_expanded_ids
            expanded_cls = " is-expanded" if is_open else ""
            
            # Card header (Always visible & fully clickable): Abstract Number, Date/Time, Bidder 1, Amount
            st.markdown(f"""
            <div class="history-item{expanded_cls}" id="card_{rec_id}" data-card-id="{rec_id}">
                <div class="history-top-row">
                    <div class="history-ac">{ac_num}</div>
                    <div class="history-date-badge">{time_tag}</div>
                </div>
                <div class="history-winner"><strong>Bidder 1:</strong> {winner}</div>
                <div class="history-meta">
                    <span class="history-amount">{amount_display}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Hidden toggle button triggered by clicking the history card directly
            st.markdown("<div class='history-hidden-toggle'>", unsafe_allow_html=True)
            if st.button("tog", key=f"hist_tog_{rec_id}_{idx}"):
                if is_open:
                    st.session_state.history_expanded_ids.discard(rec_id)
                else:
                    st.session_state.history_expanded_ids.add(rec_id)
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Extra Information & Action Controls: ONLY shown when user clicks the card to expand
            if is_open:
                rfq_date_disp = item.get("rfq_closing_date", "")
                if isinstance(rfq_date_disp, datetime.date):
                    rfq_date_disp = rfq_date_disp.strftime("%b %d, %Y")
                purpose_disp = item.get("purpose_desc") or item.get("purpose") or "—"
                b1_loc_disp = item.get("b1_loc") or item.get("winning_location") or "—"
                
                b2_name = item.get("b2_name", "")
                b2_html = f"<div class='history-extra-info-row'><strong>Bidder 2:</strong> {b2_name} | {item.get('b2_loc', '')} | {item.get('b2_amt', '')}</div>" if b2_name else ""
                
                b3_name = item.get("b3_name", "")
                b3_html = f"<div class='history-extra-info-row'><strong>Bidder 3:</strong> {b3_name} | {item.get('b3_loc', '')} | {item.get('b3_amt', '')}</div>" if b3_name else ""
                
                st.markdown(f"""
                <div class="history-extra-info">
                    <div class="history-extra-info-row"><strong>Purpose:</strong> {purpose_disp}</div>
                    <div class="history-extra-info-row"><strong>Bidder 1 Location:</strong> {b1_loc_disp}</div>
                    {b2_html}
                    {b3_html}
                    <div class="history-extra-info-row"><strong>RFQ Closing Date:</strong> {rfq_date_disp}</div>
                </div>
                """, unsafe_allow_html=True)
                
                h_c1, h_c2 = st.columns([1, 1])
                with h_c1:
                    if st.button("Load", key=f"hist_load_{rec_id}_{idx}", use_container_width=True, help="Load this canvass into the form"):
                        load_history_entry(item)
                        st.rerun()
                with h_c2:
                    if st.button("Delete", key=f"hist_del_{rec_id}_{idx}", use_container_width=True, help="Delete this record from database"):
                        delete_entry_from_txt(rec_id)
                        st.session_state.abstract_history = read_all_history_from_txt()
                        st.session_state.history_expanded_ids.discard(rec_id)
                        st.toast(f"Deleted {ac_num} from database.")
                        st.rerun()

        # Pagination controls: 5 by 5 histories
        if total_pages > 1:
            st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
            pg_col1, pg_col2, pg_col3 = st.columns([1.1, 1.4, 1.1])
            with pg_col1:
                if st.button("◀ Prev", key="btn_pg_prev", disabled=(current_page <= 1), use_container_width=True):
                    st.session_state.history_page = max(1, current_page - 1)
                    st.rerun()
            with pg_col2:
                st.markdown(
                    f"<div style='text-align: center; font-size: 11px; font-weight: 600; color: #94a3b8; line-height: 28px;'>Page {current_page} / {total_pages}</div>",
                    unsafe_allow_html=True
                )
            with pg_col3:
                if st.button("Next ▶", key="btn_pg_next", disabled=(current_page >= total_pages), use_container_width=True):
                    st.session_state.history_page = min(total_pages, current_page + 1)
                    st.rerun()
                        
    components.html("""
    <script>
    (function() {
        try {
            const pDoc = window.parent.document;
            if (!pDoc) return;
            if (!pDoc.getElementById('hist-card-delegation-js')) {
                const s = pDoc.createElement('script');
                s.id = 'hist-card-delegation-js';
                s.textContent = `
                document.addEventListener('click', function(e) {
                    const card = e.target.closest('.stSidebar .history-item[data-card-id]');
                    if (!card) return;
                    if (e.target && (e.target.tagName === 'BUTTON' || e.target.closest('button'))) return;
                    const cardId = card.getAttribute('data-card-id') || (card.id || '').replace('card_', '');
                    if (!cardId) return;
                    const btn = document.querySelector('.stSidebar [class*="st-key-hist_tog_' + cardId + '"] button') ||
                                document.querySelector('[class*="st-key-hist_tog_' + cardId + '"] button');
                    if (btn) btn.click();
                }, true);
                `;
                pDoc.head.appendChild(s);
            }
        } catch(e) {}
    })();
    </script>
    """, height=1)

# -----------------------------------------------------------------------------
# Main Header Bar with Logo and Blue-Green Gradient
# -----------------------------------------------------------------------------
logo_img_tag = f'<img src="data:image/png;base64,{logo_b64}" class="app-header-logo" alt="Maragondon Logo">' if logo_b64 else ''

st.markdown(f"""
<div class="app-header-bar">
    <div class="app-header-logo-wrapper">
        {logo_img_tag}
    </div>
    <div>
        <div class="app-header-title">ABSTRACT OF CANVASS GENERATOR</div>
        <div class="app-header-credit">Made by: Dave Tabuso - OJT of Granby Colleges and Science and Technology</div>
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
    
    col_meta1, col_meta2 = st.columns([1, 1])
    
    with col_meta1:
        # Abstract Number: AC - [Code] - [TwoDigits] - [EndOfYear]
        st.markdown('<div class="unified-field-label">Abstract Number</div>', unsafe_allow_html=True)
        ac_c1, ac_c2, ac_c3, ac_c4 = st.columns([1.1, 2.7, 1.6, 1.6])
        with ac_c1:
            st.text_input("Prefix", value="AC-", disabled=True, label_visibility="collapsed")
        with ac_c2:
            ac_code_input = st.text_input(
                "Abstract Code",
                placeholder="0001B",
                label_visibility="collapsed",
                key="ac_code"
            )
        with ac_c3:
            ac_month_input = st.text_input(
                "Month",
                placeholder="10",
                label_visibility="collapsed",
                key="ac_month",
                on_change=sync_rfq_date_from_ac
            )
        with ac_c4:
            ac_year_input = st.text_input(
                "Year",
                placeholder="26",
                label_visibility="collapsed",
                key="ac_year",
                on_change=sync_rfq_date_from_ac
            )

        clean_ac_code = ac_code_input.strip() if ac_code_input.strip() else "0001B"
        clean_ac_month = ac_month_input.strip() if ac_month_input.strip() else "10"
        clean_ac_month = clean_ac_month.zfill(2) if clean_ac_month.isdigit() else clean_ac_month
        clean_ac_year = ac_year_input.strip() if ac_year_input.strip() else "26"
        if len(clean_ac_year) > 2 and clean_ac_year.isdigit():
            clean_ac_year = clean_ac_year[-2:]
        ac_number_input = f"AC-{clean_ac_code}-{clean_ac_month}-{clean_ac_year}"
        st.caption(f"Document Abstract No.: **{ac_number_input}**")

    with col_meta2:
        # RFQ Closing Date (Calendar Picker Method)
        st.markdown('<div class="unified-field-label">RFQ Closing Date</div>', unsafe_allow_html=True)
        rfq_date_input = st.date_input(
            "RFQ Closing Date",
            key="rfq_closing_date",
            format="MM/DD/YYYY",
            label_visibility="collapsed",
            on_change=sync_ac_from_rfq_date
        )
        if isinstance(rfq_date_input, datetime.date):
            formatted_opened_date = rfq_date_input.strftime("%B %d, %Y")
        else:
            formatted_opened_date = "October 26, 2026"

        st.caption(f"Document Date: **{formatted_opened_date}**")

    raw_purpose_input = st.text_area(
        "Purpose / Description",
        value=st.session_state.get("purpose_desc", "TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD").upper(),
        placeholder="e.g. TARPAULIN FOR USE IN UPCOMING GIVE A GIFT FOR INDIGENT PWD, or RENTAL OF VEHICLE, or REIMBURSEMENT...",
        height=60,
        key="purpose_desc"
    )

    opt_col1, opt_col2 = st.columns([1, 1])
    with opt_col1:
        auto_clean = st.checkbox(
            "Auto-clean boilerplate prefix ('FOR: PAYMENT FOR THE')",
            value=st.session_state.get("auto_clean_purpose", True),
            key="auto_clean_purpose"
        )
    with opt_col2:
        include_lgu_suffix = st.checkbox(
            "Include location suffix ('IN LGU MARAGONDON, CAVITE.')",
            value=st.session_state.get("include_lgu_suffix", True),
            key="include_lgu_suffix"
        )

    # Compute final purpose line (Strictly All Upper Case)
    desc_cleaned = raw_purpose_input.strip().upper()
    if auto_clean and desc_cleaned:
        desc_cleaned = clean_purpose_description(desc_cleaned).upper()

    if include_lgu_suffix:
        final_purpose = f"{desc_cleaned} IN LGU MARAGONDON, CAVITE." if desc_cleaned else "IN LGU MARAGONDON, CAVITE."
    else:
        final_purpose = f"{desc_cleaned}." if desc_cleaned and not desc_cleaned.endswith(('.', '!', '?')) else desc_cleaned

    final_purpose = final_purpose.upper()

    st.markdown(
        f'<div class="doc-line-preview-container">Document Line Preview: FOR: PAYMENT FOR THE <span id="doc-line-preview-text" class="doc-line-preview-bold">{final_purpose or "[DESCRIPTION]"}</span></div>',
        unsafe_allow_html=True
    )

    # Real-time reactive preview on every keystroke + auto-capitalization on leaving typing area
    components.html("""
    <script>
    (function() {
        function setupPurposeReactivity() {
            try {
                const parentDoc = window.parent.document;
                if (!parentDoc) return;
                
                const textarea = parentDoc.querySelector('textarea[aria-label="Purpose / Description"]') 
                              || parentDoc.querySelector('div[data-testid="stTextArea"] textarea');
                if (!textarea) return;
                
                function updatePreview() {
                    const liveSpan = parentDoc.getElementById('doc-line-preview-text');
                    if (!liveSpan || !textarea) return;
                    
                    const checkboxes = parentDoc.querySelectorAll('input[type="checkbox"]');
                    let includeLgu = true;
                    let autoClean = true;
                    checkboxes.forEach(cb => {
                        const lbl = cb.closest('label');
                        const labelText = lbl ? lbl.innerText : '';
                        if (labelText.includes('IN LGU MARAGONDON')) {
                            includeLgu = cb.checked;
                        } else if (labelText.includes('FOR: PAYMENT FOR THE')) {
                            autoClean = cb.checked;
                        }
                    });
                    
                    let val = textarea.value.trim().toUpperCase();
                    if (autoClean && val) {
                        val = val.replace(/^(?:FOR:\\s*)?PAYMENT\\s+FOR\\s+THE\\s*|^FOR:\\s*/i, '').trim();
                        val = val.replace(/[\\s,.-]*IN\\s+LGU[- ]*MARAGONDON[,\\s]*CAVITE\\.?$/i, '').trim();
                    }
                    
                    let finalPurpose = '';
                    if (includeLgu) {
                        finalPurpose = val ? (val + ' IN LGU MARAGONDON, CAVITE.') : 'IN LGU MARAGONDON, CAVITE.';
                    } else {
                        finalPurpose = val ? (val.endsWith('.') || val.endsWith('!') || val.endsWith('?') ? val : val + '.') : '';
                    }
                    
                    liveSpan.textContent = finalPurpose ? finalPurpose.toUpperCase() : '[DESCRIPTION]';
                }
                
                if (!textarea.__reactiveBound) {
                    textarea.__reactiveBound = true;
                    textarea.addEventListener('input', updatePreview);
                    textarea.addEventListener('keyup', updatePreview);
                    
                    // Auto-capitalize on leaving the typing area (blur)
                    textarea.addEventListener('blur', function() {
                        const upper = textarea.value.toUpperCase();
                        if (textarea.value !== upper) {
                            const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
                            if (nativeSetter) {
                                nativeSetter.call(textarea, upper);
                            } else {
                                textarea.value = upper;
                            }
                            textarea.dispatchEvent(new Event('input', { bubbles: true }));
                            textarea.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                        updatePreview();
                    });
                }
                
                const checkboxes = parentDoc.querySelectorAll('input[type="checkbox"]');
                checkboxes.forEach(cb => {
                    if (!cb.__previewBound) {
                        cb.__previewBound = true;
                        cb.addEventListener('change', updatePreview);
                    }
                });

                // Auto-capitalize bidder name inputs on blur
                const bidderInputs = parentDoc.querySelectorAll('input[aria-label="Bidder 1 Name"], input[aria-label="Bidder 2 Name"], input[aria-label="Bidder 3 Name"]');
                bidderInputs.forEach(inp => {
                    if (!inp.__upperBound) {
                        inp.__upperBound = true;
                        inp.addEventListener('blur', function() {
                            const upper = inp.value.toUpperCase();
                            if (inp.value !== upper) {
                                const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                                if (nativeSetter) {
                                    nativeSetter.call(inp, upper);
                                } else {
                                    inp.value = upper;
                                }
                                inp.dispatchEvent(new Event('input', { bubbles: true }));
                                inp.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                        });
                    }
                });
                
                updatePreview();
            } catch (e) {
                // Ignore any cross-frame errors
            }
        }
        
        setupPurposeReactivity();
        setTimeout(setupPurposeReactivity, 200);
        setTimeout(setupPurposeReactivity, 600);
        setTimeout(setupPurposeReactivity, 1200);
    })();
    </script>
    """, height=0)

# -----------------------------------------------------------------------------
# Section 2: Bidders Section (Strict 2-Column Mode: Bidders & Amount)
# -----------------------------------------------------------------------------
with st.container():
    st.markdown("<div class='panel-title' style='margin-top: 10px;'>2. Bidders & Bid Amounts (Inside Table in Word Document)</div>", unsafe_allow_html=True)

    # Table Header Row: 2 Columns Only
    h_col1, h_col2 = st.columns([7, 3])
    with h_col1:
        st.markdown("<div class='bidder-col-head'><strong>NAME OF BIDDERS & LOCATION</strong></div>", unsafe_allow_html=True)
    with h_col2:
        st.markdown("<div class='bidder-col-head'><strong>AMOUNT OF BID</strong></div>", unsafe_allow_html=True)
    
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

    # Ensure bidder names are strictly and absolute all uppercase
    b1_name = b1_name.strip().upper()
    b2_name = b2_name.strip().upper()
    b3_name = b3_name.strip().upper()

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
# Section 3: Award Section (Always & Absolute Winner: Bidder 1)
# -----------------------------------------------------------------------------
with st.container():
    st.markdown("<div class='panel-title' style='margin-top: 10px;'>3. Award Section</div>", unsafe_allow_html=True)

    # Winner is always and absolute Bidder 1 automatically
    winning_name = b1_name.strip()
    winning_loc = b1_loc.strip()
    winning_amt = b1_amt_formatted
    winning_num = b1_num

    winner_is_valid = True
    if winning_num is not None:
        for b_num in [b2_num, b3_num]:
            if b_num is not None and winning_num > b_num:
                winner_is_valid = False
                break

    st.markdown(f"""
    <div style="background: rgba(15, 23, 42, 0.85); border: 1px solid rgba(56, 189, 248, 0.35); border-left: 4px solid #10b981; padding: 12px 16px; border-radius: 6px; color: #f8fafc; margin-bottom: 8px;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
            <div>
                <span style="font-size: 11px; text-transform: uppercase; color: #6ee7b7; font-weight: 700; letter-spacing: 0.8px;">Designated Winner (Automatic: Bidder 1)</span><br>
                <span style="font-size: 15px; font-weight: 800; color: #38bdf8;">{winning_name or '[Enter Bidder 1 in Section 2]'}</span>
            </div>
            <div style="text-align: right;">
                <span style="font-size: 11px; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Award Amount</span><br>
                <span style="font-size: 16px; font-weight: 800; color: #34d399;">{winning_amt or '—'}</span>
            </div>
        </div>
        <div style="font-size: 12px; color: #cbd5e1; margin-top: 6px;">
            <strong>Location:</strong> {winning_loc or '—'}
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not winner_is_valid:
        st.markdown(f"""
        <div class="val-box-warning" style="margin-bottom: 8px;">
            <strong>Validation Reminder:</strong> Bidder 1 ({winning_amt}) has a higher quotation than one of the competitors. Since Bidder 1 is the designated awardee, ensure Bidder 1 has the lowest quotation in Section 2.
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
progress_placeholder = st.empty()

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
            # Step 1: Initializing & validating data (18%)
            progress_placeholder.markdown(
                render_generation_progress_card(
                    pct=18,
                    status_text="Validating Municipal Data...",
                    subtitle="Checking bidder rankings, abstract number, and closing dates",
                    current_step=1
                ),
                unsafe_allow_html=True
            )
            time.sleep(0.12)
            
            b1_name_clean = b1_name.strip().upper()
            b2_name_clean = b2_name.strip().upper()
            b3_name_clean = b3_name.strip().upper()
            winning_name_clean = winning_name.strip().upper()

            context = {
                "purpose": final_purpose.upper(),
                "opened_date": formatted_opened_date,
                "bidder_1_name": b1_name_clean,
                "bidder_1_location": b1_loc.strip(),
                "bidder_1_amount": b1_amt_formatted,
                "bidder_2_name": b2_name_clean,
                "bidder_2_location": b2_loc.strip(),
                "bidder_2_amount": b2_amt_formatted,
                "bidder_3_name": b3_name_clean,
                "bidder_3_location": b3_loc.strip(),
                "bidder_3_amount": b3_amt_formatted,
                "winning_bidder_name": winning_name_clean,
                "winning_bidder": winning_name_clean,
                "winning_bidder_location": winning_loc.strip(),
                "ac_number": ac_number_input.strip(),
            }
            
            # Step 2: Compiling Word Template (45%)
            progress_placeholder.markdown(
                render_generation_progress_card(
                    pct=45,
                    status_text="Compiling Word Document (.docx)...",
                    subtitle="Injecting municipal data, winning bidder clause, and pricing tables",
                    current_step=2
                ),
                unsafe_allow_html=True
            )
            
            tpl = DocxTemplate(template_path)
            tpl.render(context, autoescape=True)
            
            out_stream = io.BytesIO()
            tpl.save(out_stream)
            doc_bytes = out_stream.getvalue()
            
            st.session_state.generated_doc_bytes = doc_bytes
            st.session_state.generated_context = context
            st.session_state.generated_ac_number = ac_number_input.strip()
            
            # Step 3: PDF Generation via Word COM (72%)
            progress_placeholder.markdown(
                render_generation_progress_card(
                    pct=72,
                    status_text="Pre-Generating Official PDF via Word Engine...",
                    subtitle="Compiling 1:1 vector print layout and official seal",
                    current_step=3
                ),
                unsafe_allow_html=True
            )
            
            pdf_bytes = generate_pdf_from_docx_bytes(doc_bytes)
            st.session_state.generated_pdf_bytes = pdf_bytes
            
            # Step 4: High-Definition 150-DPI Paper Preview (92%)
            progress_placeholder.markdown(
                render_generation_progress_card(
                    pct=92,
                    status_text="Rasterizing 150-DPI Document Preview...",
                    subtitle="Rendering locked 150-DPI paper preview for instant inspection",
                    current_step=4
                ),
                unsafe_allow_html=True
            )
            
            png_bytes = render_pdf_to_preview_png(pdf_bytes)
            st.session_state.generated_preview_png_bytes = png_bytes
            
            # Step 5: Finished & Ready (100%)
            progress_placeholder.markdown(
                render_generation_progress_card(
                    pct=100,
                    status_text="Finished! Official Document Ready.",
                    subtitle="Abstract of Canvass successfully compiled, rendered, and sealed",
                    current_step=4,
                    finished=True
                ),
                unsafe_allow_html=True
            )
            
            now_str = datetime.datetime.now().strftime("%b %d, %I:%M %p")
            history_entry = {
                "record_id": f"REC_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{clean_ac_code}",
                "ac_number": ac_number_input.strip(),
                "ac_code": clean_ac_code,
                "ac_month": clean_ac_month,
                "ac_year": clean_ac_year,
                "rfq_closing_date": rfq_date_input.strftime("%B %d, %Y") if isinstance(rfq_date_input, datetime.date) else str(rfq_date_input),
                "opened_date": formatted_opened_date,
                "winning_bidder": winning_name.strip(),
                "winning_location": winning_loc.strip(),
                "winning_amount": winning_amt,
                "purpose_desc": raw_purpose_input.strip().upper(),
                "include_lgu_suffix": include_lgu_suffix,
                "purpose": final_purpose.upper(),
                "b1_name": b1_name.strip().upper(),
                "b1_loc": b1_loc.strip(),
                "b1_amt": b1_amt_formatted.strip() if b1_amt_formatted else "",
                "b2_name": b2_name.strip().upper(),
                "b2_loc": b2_loc.strip(),
                "b2_amt": b2_amt_formatted.strip() if b2_amt_formatted else "",
                "b3_name": b3_name.strip().upper(),
                "b3_loc": b3_loc.strip(),
                "b3_amt": b3_amt_formatted.strip() if b3_amt_formatted else "",
                "time_str": now_str
            }
            write_single_entry_to_txt(history_entry)
            st.session_state.abstract_history = read_all_history_from_txt()
            st.session_state.history_page = 1
            
            time.sleep(0.55)
            progress_placeholder.empty()
            st.session_state.abstract_history = read_all_history_from_txt()
            st.session_state.history_page = 1
            st.rerun()
            
        except Exception as e:
            progress_placeholder.empty()
            st.error(f"Failed to generate document: {str(e)}")

# -----------------------------------------------------------------------------
# Section 4: Pre-Generated PDF Document Preview & Direct Print
# -----------------------------------------------------------------------------
if st.session_state.generated_doc_bytes is not None:
    st.markdown("<hr style='margin: 18px 0 12px 0;'>", unsafe_allow_html=True)
    
    ctx = st.session_state.generated_context
    clean_ac = re.sub(r"[^\w\-]", "_", ctx["ac_number"]).strip("_")
    download_filename = f"Abstract_of_Canvass_{clean_ac or 'filled'}.docx"
    
    doc_bytes = st.session_state.generated_doc_bytes
    doc_b64 = base64.b64encode(doc_bytes).decode("utf-8") if doc_bytes else ""
    
    pdf_bytes = st.session_state.get("generated_pdf_bytes", b"")
    pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8") if pdf_bytes else ""
    
    png_bytes = st.session_state.get("generated_preview_png_bytes", b"")
    if not png_bytes and pdf_bytes:
        png_bytes = render_pdf_to_preview_png(pdf_bytes)
        st.session_state.generated_preview_png_bytes = png_bytes
    if not png_bytes and doc_bytes:
        if not pdf_bytes:
            pdf_bytes = generate_pdf_from_docx_bytes(doc_bytes)
            st.session_state.generated_pdf_bytes = pdf_bytes
            pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8") if pdf_bytes else ""
        if pdf_bytes:
            png_bytes = render_pdf_to_preview_png(pdf_bytes)
            st.session_state.generated_preview_png_bytes = png_bytes
    png_b64 = base64.b64encode(png_bytes).decode("utf-8") if png_bytes else ""
    
    ac_num_display = ctx.get("ac_number", "AC-0001B-10-26")
    
    if png_b64:
        paper_card_inner = f'<img id="preview-paper-img" src="data:image/png;base64,{png_b64}" alt="Official Document Preview" class="preview-paper-img" />'
    else:
        paper_card_inner = build_html_document_fallback(ctx, logo_b64)
    
    # -------------------------------------------------------------------------
    # 1. Fullscreen Preview Overlay Modal (Hovers over Main Page)
    # -------------------------------------------------------------------------
    preview_overlay_html = f"""<div id="preview-modal-overlay" class="preview-modal-overlay">
<div class="preview-top-bar">
<div class="preview-title-box">
<img src="data:image/png;base64,{thumb_b64}" class="preview-bar-logo" alt="Maragondon Seal" />
<div class="preview-bar-text">
<div class="preview-bar-heading">OFFICIAL DOCUMENT PREVIEW</div>
<div class="preview-bar-sub">Abstract No.: {ac_num_display}</div>
</div>
</div>
<div class="preview-actions-group">
<a id="overlay-dl-btn" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{doc_b64}" download="{download_filename}" class="btn-overlay-dl">Download Word Document (.docx)</a>
<button type="button" id="overlay-print-btn" class="btn-overlay-print">Print Document</button>
<button type="button" id="overlay-close-btn" class="btn-overlay-close">✕ Close</button>
</div>
</div>
<div class="preview-sheet-wrapper">
<div class="preview-paper-card" id="paper-card-root">
{paper_card_inner}
</div>
</div>
<div id="another-prompt-modal" class="prompt-modal-backdrop">
<div class="prompt-dialog-box">
<img src="data:image/png;base64,{thumb_b64}" class="prompt-seal-logo" alt="Maragondon Seal" />
<div class="prompt-title">Generate Another Abstract of Canvass?</div>
<div class="prompt-desc">Your document action was initiated. Would you like to create another canvass now, or stay on this document?</div>
<div class="prompt-btn-group">
<button type="button" id="prompt-btn-confirm" class="prompt-btn-primary">Yes, Generate Another</button>
<button type="button" id="prompt-btn-cancel" class="prompt-btn-secondary">No, Stay on Current Document</button>
</div>
</div>
</div>
</div>"""
    st.markdown(preview_overlay_html, unsafe_allow_html=True)
    
    # -------------------------------------------------------------------------
    # 2. Main Page In-Place Action Bar & Controls
    # -------------------------------------------------------------------------
    exp_col1, exp_col2, exp_col3 = st.columns([4.6, 3.0, 2.4])
    with exp_col1:
        st.markdown("<div class='panel-title' style='margin: 0;'>Official Document Preview</div>", unsafe_allow_html=True)
        st.markdown(
            """<button type="button" id="btn-reopen-preview" class="btn-reopen-preview">Open Document Preview Overlay</button>""",
            unsafe_allow_html=True
        )
    with exp_col2:
        st.download_button(
            label="Download Word Document (.docx)",
            data=st.session_state.generated_doc_bytes,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    with exp_col3:
        components.html(f"""
        <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        .print-btn {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            height: 38px;
            background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
            color: #ffffff !important;
            font-family: Arial, Helvetica, sans-serif !important;
            font-weight: 700 !important;
            font-size: 14px !important;
            border: 1px solid rgba(52, 211, 153, 0.8) !important;
            border-radius: 8px !important;
            cursor: pointer !important;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35) !important;
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
            outline: none !important;
            -webkit-appearance: none !important;
            appearance: none !important;
        }}
        .print-btn:hover {{
            background: linear-gradient(135deg, #0d9488 0%, #10b981 50%, #34d399 100%) !important;
            border: 1px solid #6ee7b7 !important;
            border-radius: 8px !important;
            outline: none !important;
            box-shadow: 0 6px 20px rgba(16, 185, 129, 0.55) !important;
        }}
        .print-btn:active {{
            border-radius: 8px !important;
            outline: none !important;
            box-shadow: 0 2px 8px rgba(16, 185, 129, 0.4) !important;
        }}
        </style>
        <button type="button" class="print-btn" onclick="triggerDirectPrint()">
            Print Document
        </button>
        <script>
        function triggerDirectPrint() {{
            try {{
                if (window.parent && window.parent.runDirectPrint) {{
                    window.parent.runDirectPrint();
                }}
            }} catch(e) {{
                console.error("Print call error:", e);
            }}
        }}
        </script>
        """, height=42)
        
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # In-page 1:1 paper document view
    in_page_preview = f"""<div style="display: flex; justify-content: center; align-items: flex-start; width: 100%; padding: 10px 0 30px 0;">
<div style="width: 100%; max-width: 820px; background: #ffffff; border-radius: 4px; box-shadow: 0 14px 44px rgba(0, 0, 0, 0.7), 0 2px 10px rgba(0, 0, 0, 0.45); overflow: hidden; border: 1px solid rgba(148, 163, 184, 0.35);">
{paper_card_inner}
</div>
</div>"""
    st.markdown(in_page_preview, unsafe_allow_html=True)

    # Companion script to manage overlay dialog, print, and generate-another prompts
    components.html(f"""
    <script>
    (function() {{
        try {{
            const pDoc = window.parent.document;
            const pWin = window.parent;
            if (!pDoc) return;

            // Direct print helper attached to parent window
            pWin.runDirectPrint = function() {{
                try {{
                    let printFrame = pDoc.getElementById('app-direct-print-frame');
                    if (printFrame) printFrame.remove();
                    printFrame = pDoc.createElement('iframe');
                    printFrame.id = 'app-direct-print-frame';
                    printFrame.style.position = 'fixed';
                    printFrame.style.top = '-10000px';
                    printFrame.style.left = '-10000px';
                    printFrame.style.width = '1000px';
                    printFrame.style.height = '1400px';
                    printFrame.style.border = '0';
                    printFrame.style.visibility = 'hidden';
                    pDoc.body.appendChild(printFrame);

                    const frameDoc = printFrame.contentWindow.document;
                    frameDoc.open();
                    frameDoc.write(`
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Abstract of Canvass</title>
                            <style>
                                @page {{ size: auto; margin: 0mm; }}
                                @media print {{
                                    html, body {{ margin: 0 !important; padding: 0 !important; background: #ffffff !important; }}
                                    img {{ width: 100% !important; max-width: 100% !important; height: auto !important; display: block !important; page-break-inside: avoid !important; }}
                                }}
                                html, body {{ margin: 0; padding: 0; background: #ffffff; }}
                                img {{ width: 100%; height: auto; display: block; }}
                            </style>
                        </head>
                        <body>
                            <div id="print-body-content"></div>
                        </body>
                        </html>
                    `);
                    frameDoc.close();

                    const printBodyContainer = frameDoc.getElementById('print-body-content');
                    const hasPng = "{png_b64}".length > 10;
                    if (hasPng) {{
                        printBodyContainer.innerHTML = '<img id="print-image" src="data:image/png;base64,{png_b64}" />';
                    }} else {{
                        const paperElem = pDoc.getElementById('paper-card-root');
                        printBodyContainer.innerHTML = paperElem ? paperElem.innerHTML : '';
                    }}

                    const pImg = frameDoc.getElementById('print-image');
                    function doPrint() {{
                        try {{
                            printFrame.contentWindow.focus();
                            printFrame.contentWindow.print();
                        }} catch(e) {{
                            console.error('Print trigger error:', e);
                        }}
                        function showPromptOnParent() {{
                            const m = pDoc.getElementById('another-prompt-modal');
                            const ov = pDoc.getElementById('preview-modal-overlay');
                            if (m) m.classList.add('is-open');
                            if (ov) ov.classList.add('prompt-active');
                            try {{
                                pDoc.body.classList.add('prompt-active');
                                pDoc.documentElement.classList.add('prompt-active');
                            }} catch(e) {{}}
                        }}
                        try {{
                            printFrame.contentWindow.addEventListener('afterprint', function() {{
                                setTimeout(showPromptOnParent, 600);
                            }}, {{ once: true }});
                        }} catch(e) {{}}
                        setTimeout(showPromptOnParent, 2500);
                    }}

                    if (pImg) {{
                        if (pImg.complete) setTimeout(doPrint, 150);
                        else pImg.onload = function() {{ setTimeout(doPrint, 150); }};
                    }} else {{
                        setTimeout(doPrint, 200);
                    }}
                }} catch(err) {{
                    console.error('runDirectPrint error:', err);
                }}
            }};

            // Inject persistent head delegation script if not already present
            if (!pDoc.getElementById('preview-overlay-delegation-js')) {{
                const s = pDoc.createElement('script');
                s.id = 'preview-overlay-delegation-js';
                s.textContent = `
                (function() {{
                    document.addEventListener('click', function(e) {{
                        // 1. Reopen preview overlay button
                        const reopen = e.target.closest('#btn-reopen-preview') || e.target.closest('.btn-reopen-preview');
                        if (reopen) {{
                            e.preventDefault();
                            e.stopPropagation();
                            const ov = document.getElementById('preview-modal-overlay');
                            if (ov) {{
                                ov.classList.remove('is-hidden');
                                ov.scrollTop = 0;
                            }}
                            return;
                        }}

                        // 2. Close overlay button
                        const closeBtn = e.target.closest('#overlay-close-btn') || e.target.closest('.btn-overlay-close');
                        if (closeBtn) {{
                            e.preventDefault();
                            e.stopPropagation();
                            const ov = document.getElementById('preview-modal-overlay');
                            if (ov) ov.classList.add('is-hidden');
                            return;
                        }}

                        // 3. Print button in overlay
                        const printBtn = e.target.closest('#overlay-print-btn') || e.target.closest('.btn-overlay-print');
                        if (printBtn) {{
                            e.preventDefault();
                            e.stopPropagation();
                            if (window.runDirectPrint) window.runDirectPrint();
                            return;
                        }}

                        // 4. Download button in overlay
                        const dlBtn = e.target.closest('#overlay-dl-btn');
                        if (dlBtn) {{
                            const ov = document.getElementById('preview-modal-overlay');
                            if (ov) ov.classList.add('prompt-active');
                            try {{
                                document.body.classList.add('prompt-active');
                                document.documentElement.classList.add('prompt-active');
                            }} catch(err) {{}}
                            const m = document.getElementById('another-prompt-modal');
                            if (m) m.classList.add('is-open');
                        }}

                        // 5. Cancel button in "Generate Another?" modal
                        const cancelBtn = e.target.closest('#prompt-btn-cancel');
                        if (cancelBtn) {{
                            e.preventDefault();
                            e.stopPropagation();
                            const m = document.getElementById('another-prompt-modal');
                            if (m) m.classList.remove('is-open');
                            const ov = document.getElementById('preview-modal-overlay');
                            if (ov) ov.classList.remove('prompt-active');
                            try {{
                                document.body.classList.remove('prompt-active');
                                document.documentElement.classList.remove('prompt-active');
                            }} catch(err) {{}}
                            return;
                        }}

                        // 6. Confirm button in "Generate Another?" modal
                        const confirmBtn = e.target.closest('#prompt-btn-confirm');
                        if (confirmBtn) {{
                            e.preventDefault();
                            e.stopPropagation();
                            const m = document.getElementById('another-prompt-modal');
                            if (m) m.classList.remove('is-open');
                            const ov = document.getElementById('preview-modal-overlay');
                            if (ov) {{
                                ov.classList.remove('prompt-active');
                                ov.classList.add('is-hidden');
                            }}
                            try {{
                                document.body.classList.remove('prompt-active');
                                document.documentElement.classList.remove('prompt-active');
                            }} catch(err) {{}}
                            try {{
                                const buttons = Array.from(document.querySelectorAll('button'));
                                const clearBtn = buttons.find(b => (b.innerText || b.textContent || '').includes('Clear Form'));
                                if (clearBtn) clearBtn.click();
                            }} catch(err) {{}}
                            try {{
                                window.scrollTo({{ top: 0, behavior: 'smooth' }});
                            }} catch(err) {{}}
                            return;
                        }}
                    }}, true);

                    // Prevent wheel/touch bubbling when prompt is open
                    const stopScroll = function(e) {{
                        const m = document.getElementById('another-prompt-modal');
                        const ov = document.getElementById('preview-modal-overlay');
                        if ((m && m.classList.contains('is-open')) || (ov && ov.classList.contains('prompt-active'))) {{
                            const withinModal = e.target && e.target.closest && e.target.closest('.prompt-dialog-box');
                            if (!withinModal) {{
                                e.preventDefault();
                                e.stopPropagation();
                                return false;
                            }}
                        }}
                    }};
                    document.addEventListener('wheel', stopScroll, {{ passive: false }});
                    document.addEventListener('touchmove', stopScroll, {{ passive: false }});
                }})();
                `;
                pDoc.head.appendChild(s);
            }}

            // Also attach listener to download button in page
            const dlBtn = pDoc.querySelector('[data-testid="stDownloadButton"] button');
            if (dlBtn && !dlBtn.__promptBound) {{
                dlBtn.__promptBound = true;
                dlBtn.addEventListener('click', function() {{
                    const ov = pDoc.getElementById('preview-modal-overlay');
                    if (ov) ov.classList.add('prompt-active');
                    try {{
                        pDoc.body.classList.add('prompt-active');
                        pDoc.documentElement.classList.add('prompt-active');
                    }} catch(e) {{}}
                    const m = pDoc.getElementById('another-prompt-modal');
                    if (m) m.classList.add('is-open');
                }});
            }}
        }} catch(e) {{
            console.log('Overlay script initialization error:', e);
        }}
    }})();
    </script>
    """, height=1)
