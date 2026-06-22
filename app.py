"""
WMCT WI Dashboard — Streamlit Application
Replicates all functionality from the Excel/VBA dashboard.
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import os
from datetime import datetime
import io
import re

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="WMCT WI Dashboard",
    page_icon="⭐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# DATA STORE  (session_state acts as in-memory DB)
# ─────────────────────────────────────────────
VALID_PHASES = [
    "Ship Planning", "Berth Planning", "Operation Management",
    "Yard Define", "Yard Planning", "Terminal Monitoring",
    "Reefer Operation", "DG Operation", "Gate",
    "Security Management", "Planning", "Operations",
    "General", "CATOS Admin",
]

STATUS_OPTIONS = ["Not Started", "In Progress", "Completed"]

SCORE_OPTIONS = ["", "1", "2", "3", "4", "5"]

DEFAULT_BINOMES = [
    {"name": "Binôme 1", "members": "Hamza & Badr"},
    {"name": "Binôme 2", "members": "Amin & Anas"},
    {"name": "Binôme 3", "members": "Oussam & Ibrahim"},
    {"name": "Binôme 4", "members": "Oussama & Mehdi"},
    {"name": "Binôme 5", "members": "Khalil & Oualid"},
]

DEFAULT_WIS = [
    {"id": "WI-V1-1", "title": "WI: Open Vessel",                          "phase": "Ship Planning",  "confirmed": False, "sop": ""},
    {"id": "WI-V1-2", "title": "WI: Create Scenario Vessel & Voyage",       "phase": "Ship Planning",  "confirmed": False, "sop": ""},
    {"id": "WI-V1-3", "title": "WI: Create Stoppage",                       "phase": "Planning",       "confirmed": True,  "sop": "SOP_Template_CATOS.xlsx"},
    {"id": "WI-V1-4", "title": "WI: ROB to Restow / Restow to ROB",         "phase": "Operations",     "confirmed": False, "sop": ""},
    {"id": "WI-V1-5", "title": "WI: Bay Split & Crane Assignment",           "phase": "Planning",       "confirmed": False, "sop": ""},
    {"id": "WI-V1-6", "title": "WI: Bay Split & Crane Assignment (2)",       "phase": "Planning",       "confirmed": True,  "sop": "WMCT_WI_Dashboard.xlsm"},
    {"id": "WI-V1-7", "title": "WI: Standar work",                          "phase": "Operations",     "confirmed": False, "sop": ""},
]

# progress[(wi_id, binome_name)] = {"status": ..., "score": ..., "notes": ...}
DEFAULT_PROGRESS = {
    ("WI-V1-1","Binôme 1"):{"status":"Completed","score":"5","notes":""},
    ("WI-V1-1","Binôme 2"):{"status":"Completed","score":"4","notes":""},
    ("WI-V1-1","Binôme 3"):{"status":"Completed","score":"3","notes":""},
    ("WI-V1-1","Binôme 4"):{"status":"Completed","score":"5","notes":""},
    ("WI-V1-1","Binôme 5"):{"status":"Not Started","score":"","notes":""},
    ("WI-V1-2","Binôme 1"):{"status":"Completed","score":"4","notes":""},
    ("WI-V1-2","Binôme 2"):{"status":"In Progress","score":"2","notes":""},
    ("WI-V1-2","Binôme 3"):{"status":"Completed","score":"5","notes":""},
    ("WI-V1-2","Binôme 4"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-2","Binôme 5"):{"status":"Completed","score":"3","notes":""},
    ("WI-V1-3","Binôme 1"):{"status":"In Progress","score":"2","notes":""},
    ("WI-V1-3","Binôme 2"):{"status":"Completed","score":"5","notes":""},
    ("WI-V1-3","Binôme 3"):{"status":"Completed","score":"","notes":""},
    ("WI-V1-3","Binôme 4"):{"status":"In Progress","score":"3","notes":""},
    ("WI-V1-3","Binôme 5"):{"status":"Completed","score":"4","notes":""},
    ("WI-V1-4","Binôme 1"):{"status":"Completed","score":"","notes":""},
    ("WI-V1-4","Binôme 2"):{"status":"Not Started","score":"","notes":""},
    ("WI-V1-4","Binôme 3"):{"status":"In Progress","score":"2","notes":""},
    ("WI-V1-4","Binôme 4"):{"status":"In Progress","score":"4","notes":""},
    ("WI-V1-4","Binôme 5"):{"status":"In Progress","score":"3","notes":""},
    ("WI-V1-5","Binôme 1"):{"status":"Not Started","score":"5","notes":""},
    ("WI-V1-5","Binôme 2"):{"status":"Completed","score":"3","notes":""},
    ("WI-V1-5","Binôme 3"):{"status":"Completed","score":"4","notes":""},
    ("WI-V1-5","Binôme 4"):{"status":"Completed","score":"","notes":""},
    ("WI-V1-5","Binôme 5"):{"status":"Not Started","score":"","notes":""},
    ("WI-V1-6","Binôme 1"):{"status":"Completed","score":"","notes":""},
    ("WI-V1-6","Binôme 2"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-6","Binôme 3"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-6","Binôme 4"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-6","Binôme 5"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-7","Binôme 1"):{"status":"Not Started","score":"","notes":""},
    ("WI-V1-7","Binôme 2"):{"status":"In Progress","score":"","notes":""},
    ("WI-V1-7","Binôme 3"):{"status":"Completed","score":"","notes":""},
    ("WI-V1-7","Binôme 4"):{"status":"Not Started","score":"","notes":""},
    ("WI-V1-7","Binôme 5"):{"status":"In Progress","score":"","notes":""},
}

def init_state():
    if "wis" not in st.session_state:
        st.session_state.wis = DEFAULT_WIS.copy()
    if "binomes" not in st.session_state:
        st.session_state.binomes = DEFAULT_BINOMES.copy()
    if "progress" not in st.session_state:
        st.session_state.progress = {str(k): v for k, v in DEFAULT_PROGRESS.items()}

def get_progress(wi_id, binome_name):
    key = str((wi_id, binome_name))
    return st.session_state.progress.get(key, {"status": "Not Started", "score": "", "notes": ""})

def set_progress(wi_id, binome_name, data):
    key = str((wi_id, binome_name))
    st.session_state.progress[key] = data

def wi_is_completed(wi_id):
    """A WI is completed if at least one binôme marked it Completed."""
    for b in st.session_state.binomes:
        p = get_progress(wi_id, b["name"])
        if p["status"] == "Completed":
            return True
    return False

def next_wi_id():
    ids = [w["id"] for w in st.session_state.wis]
    nums = []
    for i in ids:
        try:
            nums.append(int(i.split("-")[-1]))
        except:
            pass
    nxt = max(nums) + 1 if nums else 1
    return f"WI-V1-{nxt}"

def create_new_wi(name, phase, sop_link, confirmed):
    """Shared logic used by both the sidebar page and the floating Quick-Add dialog."""
    if not name or not name.strip():
        return None, "Le nom de la WI ne peut pas être vide."
    new_id = next_wi_id()
    title = f"WI: {name.strip()}"
    new_wi = {
        "id": new_id, "title": title, "phase": phase,
        "confirmed": confirmed, "sop": (sop_link or "").strip(),
    }
    st.session_state.wis.append(new_wi)
    for b in st.session_state.binomes:
        set_progress(new_id, b["name"], {"status": "Not Started", "score": "", "notes": ""})
    return new_id, title

# ─────────────────────────────────────────────
# COMPUTED KPIs
# ─────────────────────────────────────────────
def compute_kpis():
    wis = st.session_state.wis
    total = len(wis)
    completed = sum(1 for w in wis if wi_is_completed(w["id"]))
    in_progress = sum(
        1 for w in wis
        if not wi_is_completed(w["id"]) and any(
            get_progress(w["id"], b["name"])["status"] == "In Progress"
            for b in st.session_state.binomes
        )
    )
    confirmed = sum(1 for w in wis if w["confirmed"])
    not_confirmed = total - confirmed
    pct = completed / total if total else 0
    return {
        "total": total, "completed": completed, "in_progress": in_progress,
        "confirmed": confirmed, "not_confirmed": not_confirmed, "pct": pct,
    }

def compute_binome_stats():
    rows = []
    for b in st.session_state.binomes:
        bn = b["name"]
        bm = b["members"]
        comp = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Completed")
        not_started = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Not Started")
        total = len(st.session_state.wis)
        pct = comp / total if total else 0
        if pct >= 0.7:
            status = "🟢 On Track"
        elif pct >= 0.4:
            status = "🔶 On Track"
        else:
            status = "🔴 Need Support"
        rows.append({"Binôme": bn, "Members": bm, "Completed": comp,
                     "Not Started": not_started, "Completion %": f"{pct:.0%}", "Status": status})
    return rows

def compute_wi_table():
    rows = []
    for w in st.session_state.wis:
        teams_completed = sum(1 for b in st.session_state.binomes if get_progress(w["id"], b["name"])["status"] == "Completed")
        in_prog = sum(1 for b in st.session_state.binomes if get_progress(w["id"], b["name"])["status"] == "In Progress")
        total_b = len(st.session_state.binomes)
        pct = teams_completed / total_b if total_b else 0
        mgr = "✅ Confirmed" if w["confirmed"] else "⏳ Pending"
        sop = w["sop"] if w["sop"] else "—"
        rows.append({
            "WI ID": w["id"], "Title": w["title"], "Phase": w["phase"],
            "Teams ✅": teams_completed, "In Progress 🔄": in_prog,
            "Completion %": f"{pct:.0%}", "Manager": mgr, "SOP": sop,
        })
    return rows

def compute_category_stats():
    cat_map = {}
    for w in st.session_state.wis:
        ph = w["phase"]
        if ph not in cat_map:
            cat_map[ph] = {"total": 0, "completed": 0}
        cat_map[ph]["total"] += 1
        if wi_is_completed(w["id"]):
            cat_map[ph]["completed"] += 1
    rows = []
    for phase in VALID_PHASES:
        d = cat_map.get(phase, {"total": 0, "completed": 0})
        rows.append({
            "Category": phase,
            "Total WIs": d["total"],
            "Completed": d["completed"],
            "Not Completed": d["total"] - d["completed"],
            "Completion %": f"{d['completed']/d['total']:.0%}" if d["total"] else "—",
        })
    return rows

def incomplete_wis():
    out = []
    for w in st.session_state.wis:
        if not wi_is_completed(w["id"]):
            statuses = [get_progress(w["id"], b["name"])["status"] for b in st.session_state.binomes]
            if any(s == "In Progress" for s in statuses):
                overall = "🔶 In Progress"
            else:
                overall = "⏳ Not Started"
            out.append({"WI ID": w["id"], "Title": w["title"], "Phase": w["phase"], "Status": overall})
    return out

# ─────────────────────────────────────────────
# EXPORT
# ─────────────────────────────────────────────
def export_to_excel():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # WI Register
        pd.DataFrame(st.session_state.wis).to_excel(writer, sheet_name="WI_Register", index=False)
        # Progress Data
        prog_rows = []
        for w in st.session_state.wis:
            row = {"WI ID": w["id"], "Title": w["title"], "Phase": w["phase"]}
            for b in st.session_state.binomes:
                p = get_progress(w["id"], b["name"])
                row[f"{b['name']} Status"] = p["status"]
                row[f"{b['name']} Score"] = p["score"]
                row[f"{b['name']} Notes"] = p["notes"]
            prog_rows.append(row)
        pd.DataFrame(prog_rows).to_excel(writer, sheet_name="Progress_Data", index=False)
        # Summary
        pd.DataFrame(compute_wi_table()).to_excel(writer, sheet_name="Dashboard_Summary", index=False)
        pd.DataFrame(compute_category_stats()).to_excel(writer, sheet_name="Insights", index=False)
    return buf.getvalue()

def drive_file_id(link):
    """Extracts the file ID from common Google Drive share-link formats."""
    if not link:
        return None
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", link)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", link)
    if m:
        return m.group(1)
    return None

def render_drive_preview(sop_link, key_prefix):
    """Shows a Google Drive-hosted SOP file inline (embedded preview + open link)."""
    if not sop_link:
        return
    file_id = drive_file_id(sop_link)
    with st.expander("📊 Aperçu du fichier SOP (Google Drive)", expanded=False):
        if file_id:
            embed_url = f"https://drive.google.com/file/d/{file_id}/preview"
            components.iframe(embed_url, height=420)
        else:
            st.caption("Lien détecté mais l'aperçu intégré n'est disponible que pour les liens Google Drive standards.")
        st.link_button("🔗 Ouvrir dans Google Drive", sop_link, key=f"open_drive_{key_prefix}")

def logo_svg(uid="a", size=46):
    """WMCT logo badge — light circle so it stays visible on the dark sidebar/header."""
    return f"""
<svg width="{size}" height="{size}" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="wmctGrad-{uid}" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#2e75b6"/>
      <stop offset="100%" stop-color="#f59c20"/>
    </linearGradient>
  </defs>
  <circle cx="32" cy="32" r="31" fill="#ffffff" stroke="url(#wmctGrad-{uid})" stroke-width="2.5"/>
  <rect x="14" y="36" width="9" height="9" fill="#f59c20" rx="1.5"/>
  <rect x="24" y="36" width="9" height="9" fill="#2e75b6" rx="1.5"/>
  <rect x="34" y="36" width="9" height="9" fill="#00b050" rx="1.5"/>
  <rect x="44" y="36" width="6" height="9" fill="#c00000" rx="1.5"/>
  <path d="M20 32 L20 16 L30 16 L36 24 L36 32" stroke="#0f2b46" stroke-width="2.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="20" cy="16" r="2.4" fill="#f59c20"/>
  <path d="M10 49 H54" stroke="#0f2b46" stroke-width="2" stroke-linecap="round"/>
</svg>
"""

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Poppins', sans-serif; }

  /* Header bar */
  .wmct-header {
    background: linear-gradient(135deg, #0f2b46 0%, #1f4e79 55%, #2e75b6 100%);
    color: white; padding: 22px 28px; border-radius: 14px;
    margin-bottom: 24px; display: flex; align-items: center; gap: 16px;
    box-shadow: 0 4px 18px rgba(15,43,70,.25);
  }
  .wmct-header .logo-box { flex-shrink: 0; }
  .wmct-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; letter-spacing: .3px; }
  .wmct-header p  { margin: 4px 0 0; opacity: .85; font-size: .88rem; }

  /* Sidebar logo */
  .sidebar-logo { display:flex; align-items:center; gap:12px; margin-bottom:6px; padding:4px 0; }
  .sidebar-logo svg { flex-shrink: 0; display:block; }
  .sidebar-logo .brand { font-size:1.0rem; font-weight:700; color:white; line-height:1.25; }
  .sidebar-logo .brand small { display:block; font-weight:400; opacity:.75; font-size:.7rem; }

  /* KPI cards */
  .kpi-grid { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 24px; }
  .kpi-card {
    flex: 1; min-width: 130px;
    background: white; border-radius: 12px;
    padding: 18px 20px; text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,.07);
    border-top: 4px solid #2e75b6;
    transition: transform .15s ease, box-shadow .15s ease;
  }
  .kpi-card:hover { transform: translateY(-3px); box-shadow: 0 6px 16px rgba(0,0,0,.12); }
  .kpi-card .val { font-size: 2rem; font-weight: 700; color: #1f4e79; }
  .kpi-card .lbl { font-size: .78rem; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: .5px; }
  .kpi-card.green  { border-top-color: #00b050; }
  .kpi-card.green .val { color: #00b050; }
  .kpi-card.orange { border-top-color: #f59c20; }
  .kpi-card.orange .val { color: #f59c20; }
  .kpi-card.red    { border-top-color: #c00000; }
  .kpi-card.red .val { color: #c00000; }

  /* Section titles */
  .section-title {
    font-size: 1.1rem; font-weight: 700; color: #1f4e79;
    border-left: 4px solid #2e75b6; padding-left: 10px;
    margin: 28px 0 14px;
  }

  /* Status badges */
  .badge-completed   { background:#e2f0d9; color:#375623; border-radius:4px; padding:2px 8px; font-size:.82rem; font-weight:600; }
  .badge-inprogress  { background:#fff2cc; color:#7d6608; border-radius:4px; padding:2px 8px; font-size:.82rem; font-weight:600; }
  .badge-notstarted  { background:#f2f2f2; color:#444; border-radius:4px; padding:2px 8px; font-size:.82rem; font-weight:600; }

  /* Sidebar */
  section[data-testid="stSidebar"] { background: linear-gradient(180deg,#0f2b46 0%, #1f4e79 100%); }
  section[data-testid="stSidebar"] * { color: white !important; }
  section[data-testid="stSidebar"] .stSelectbox label,
  section[data-testid="stSidebar"] .stTextInput label { color: #cde !important; }
  section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.15); }

  /* Floating "Add New WI" action button — visible on every page */
  div.st-key-fab_container {
    position: fixed; bottom: 22px; right: 26px; z-index: 9999;
    width: auto;
  }
  div.st-key-fab_container button {
    background: linear-gradient(135deg, #f59c20 0%, #e0830a 100%) !important;
    color: white !important; border: none !important;
    border-radius: 50px !important; padding: 14px 26px !important;
    font-weight: 700 !important; font-size: .95rem !important;
    box-shadow: 0 6px 18px rgba(229,140,10,.45) !important;
    transition: transform .15s ease, box-shadow .15s ease !important;
  }
  div.st-key-fab_container button:hover {
    transform: translateY(-2px) scale(1.03);
    box-shadow: 0 8px 22px rgba(229,140,10,.55) !important;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# QUICK-ADD WI DIALOG (opened from the floating button on every page)
# ─────────────────────────────────────────────
@st.dialog("➕ Ajouter une nouvelle WI", width="large")
def add_wi_dialog():
    st.caption("Ce formulaire est accessible depuis n'importe quelle page via le bouton flottant.")
    qa_name = st.text_input("Nom de la WI *", placeholder="ex. Open Vessel", key="qa_name")
    qa_phase = st.selectbox("Phase / Catégorie *", VALID_PHASES, key="qa_phase")
    qa_link = st.text_input(
        "🔗 Lien Google Drive du fichier SOP (optionnel)",
        placeholder="https://drive.google.com/file/d/.../view?usp=sharing",
        key="qa_link",
    )
    qa_confirmed = st.checkbox("✅ Confirmé par le manager", key="qa_confirmed")

    if qa_link.strip():
        render_drive_preview(qa_link.strip(), key_prefix="qa")

    c1, c2 = st.columns(2)
    if c1.button("➕ Créer la WI", type="primary", use_container_width=True, key="qa_submit"):
        new_id, result = create_new_wi(qa_name, qa_phase, qa_link, qa_confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Créée : **{new_id}** — {result}")
            st.rerun()
    if c2.button("Annuler", use_container_width=True, key="qa_cancel"):
        st.rerun()

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
init_state()

# Sidebar navigation
with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-logo">
        {logo_svg('sidebar', 48)}
        <div class="brand">WMCT Dashboard<small>CATOS Training Phase</small></div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    page = st.radio("Navigation", [
        "📊 Dashboard",
        "📋 WI Register",
        "🔄 Progress Data",
        "💡 Insights",
        "➕ Add New WI",
        "👥 Manage Binômes",
        "📤 Export",
    ])
    st.markdown("---")
    if st.button("➕ Add New WI (quick)", use_container_width=True, key="sidebar_quick_add"):
        add_wi_dialog()
    st.markdown("---")
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# ══════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════
if page == "📊 Dashboard":
    st.markdown(f"""
    <div class="wmct-header">
      <div class="logo-box">{logo_svg('header', 46)}</div>
      <div>
        <h1>WMCT — WEST MED CONTAINER TERMINAL</h1>
        <p>CATOS Training Phase  |  WI Progress Dashboard  |  Pre-Go-Live Training</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    kpis = compute_kpis()

    # KPI Cards
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("📋 Total WIs", kpis["total"])
    with c2:
        st.metric("✅ Completed", kpis["completed"])
    with c3:
        st.metric("🔄 In Progress", kpis["in_progress"])
    with c4:
        st.metric("⏳ Not Confirmed", kpis["not_confirmed"])
    with c5:
        st.metric("🔐 Confirmed", kpis["confirmed"])

    # Progress bar
    st.markdown(f"**Overall Completion: {kpis['pct']:.0%}**")
    st.progress(kpis["pct"])

    # ── Binôme Overview ──
    st.markdown('<div class="section-title">BINÔME PROGRESS OVERVIEW</div>', unsafe_allow_html=True)
    binome_rows = compute_binome_stats()
    df_binome = pd.DataFrame(binome_rows)
    st.dataframe(df_binome, use_container_width=True, hide_index=True)

    # Mini progress bars per binôme
    cols = st.columns(len(st.session_state.binomes))
    for i, b in enumerate(st.session_state.binomes):
        bn = b["name"]
        total = len(st.session_state.wis)
        comp = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Completed")
        pct = comp / total if total else 0
        with cols[i]:
            st.caption(f"**{bn}**\n{b['members']}")
            st.progress(pct, text=f"{pct:.0%}")

    # ── WI Table ──
    st.markdown('<div class="section-title">WI STATUS TABLE</div>', unsafe_allow_html=True)
    df_wi = pd.DataFrame(compute_wi_table())
    st.dataframe(df_wi, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# PAGE: WI REGISTER
# ══════════════════════════════════════════════
elif page == "📋 WI Register":
    st.markdown("## 📋 WI Register")
    st.caption("All Work Instructions — edit titles, phases, manager confirmation and the Google Drive SOP link.")

    for i, w in enumerate(st.session_state.wis):
        with st.expander(f"**{w['id']}** — {w['title']}", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                new_title = st.text_input("WI Title", w["title"], key=f"title_{i}")
                new_phase = st.selectbox("Phase / Category", VALID_PHASES,
                                         index=VALID_PHASES.index(w["phase"]) if w["phase"] in VALID_PHASES else 0,
                                         key=f"phase_{i}")
            with col2:
                new_confirmed = st.checkbox("✅ Manager Confirmed", w["confirmed"], key=f"confirmed_{i}")
                new_sop = st.text_input(
                    "🔗 Lien Google Drive du fichier SOP",
                    w["sop"], key=f"sop_{i}",
                    placeholder="https://drive.google.com/file/d/.../view?usp=sharing",
                    help="Mets le fichier SOP sur Google Drive, clique sur 'Partager' → 'Copier le lien', puis colle-le ici.",
                )

            if st.button("💾 Save", key=f"save_wi_{i}"):
                st.session_state.wis[i].update({
                    "title": new_title, "phase": new_phase,
                    "confirmed": new_confirmed, "sop": new_sop.strip(),
                })
                st.success("Saved!")
                st.rerun()

            render_drive_preview(w["sop"], key_prefix=f"reg_{i}")

            if st.button("🗑️ Delete WI", key=f"del_wi_{i}", type="secondary"):
                st.session_state.wis.pop(i)
                st.rerun()


# ══════════════════════════════════════════════
# PAGE: PROGRESS DATA
# ══════════════════════════════════════════════
elif page == "🔄 Progress Data":
    st.markdown("## 🔄 Progress Data")
    st.caption("Update each binôme's status, score (1-5) and notes for every WI.")

    binome_names = [b["name"] for b in st.session_state.binomes]
    selected_binome = st.selectbox("Filter by Binôme (or see all)", ["All"] + binome_names)

    for w in st.session_state.wis:
        st.markdown(f"### {w['id']} — {w['title']}")
        st.caption(f"Phase: {w['phase']}")

        binomes_to_show = st.session_state.binomes if selected_binome == "All" \
            else [b for b in st.session_state.binomes if b["name"] == selected_binome]

        cols = st.columns(len(binomes_to_show))
        for j, b in enumerate(binomes_to_show):
            bn = b["name"]
            p = get_progress(w["id"], bn)
            with cols[j]:
                st.caption(f"**{bn}**\n_{b['members']}_")
                new_status = st.selectbox("Status", STATUS_OPTIONS,
                                          index=STATUS_OPTIONS.index(p["status"]) if p["status"] in STATUS_OPTIONS else 0,
                                          key=f"st_{w['id']}_{bn}")
                new_score  = st.selectbox("Score", SCORE_OPTIONS,
                                          index=SCORE_OPTIONS.index(p["score"]) if p["score"] in SCORE_OPTIONS else 0,
                                          key=f"sc_{w['id']}_{bn}")
                new_notes  = st.text_input("Notes", p["notes"], key=f"no_{w['id']}_{bn}")
                if st.button("Save", key=f"sv_{w['id']}_{bn}"):
                    set_progress(w["id"], bn, {"status": new_status, "score": new_score, "notes": new_notes})
                    st.success("✅")
                    st.rerun()
        st.divider()


# ══════════════════════════════════════════════
# PAGE: INSIGHTS
# ══════════════════════════════════════════════
elif page == "💡 Insights":
    st.markdown("## 💡 Insights & Analytics")
    st.caption("A WI is 'Completed' once at least 1 binôme has marked it Completed.")

    # Category breakdown
    st.markdown('<div class="section-title">WIs BY CATEGORY</div>', unsafe_allow_html=True)
    df_cat = pd.DataFrame(compute_category_stats())
    df_cat_active = df_cat[df_cat["Total WIs"] > 0]
    st.dataframe(df_cat_active, use_container_width=True, hide_index=True)

    # Bar chart
    if not df_cat_active.empty:
        st.markdown("**Category completion chart**")
        chart_data = df_cat_active.set_index("Category")[["Completed", "Not Completed"]]
        st.bar_chart(chart_data, color=["#00b050", "#c00000"])

    # Incomplete WIs
    st.markdown('<div class="section-title">WI À SUIVRE — NON COMPLÈTES</div>', unsafe_allow_html=True)
    inc = incomplete_wis()
    if inc:
        df_inc = pd.DataFrame(inc)
        # Group by phase
        for phase in sorted(df_inc["Phase"].unique()):
            st.markdown(f"**{phase}**")
            sub = df_inc[df_inc["Phase"] == phase][["WI ID", "Title", "Status"]]
            st.dataframe(sub, use_container_width=True, hide_index=True)
        st.info(f"Total WI non complètes : **{len(inc)}**")
    else:
        st.success("✅ Aucune WI incomplète — tout est à jour !")

    # Binôme performance
    st.markdown('<div class="section-title">BINÔME PERFORMANCE</div>', unsafe_allow_html=True)
    for b in st.session_state.binomes:
        bn = b["name"]
        scores = []
        for w in st.session_state.wis:
            p = get_progress(w["id"], bn)
            if p["score"]:
                try:
                    scores.append(int(p["score"]))
                except:
                    pass
        avg_score = sum(scores) / len(scores) if scores else 0
        comp_count = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Completed")
        total = len(st.session_state.wis)
        pct = comp_count / total if total else 0
        c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
        c1.write(f"**{bn}** — {b['members']}")
        c2.metric("Completed", f"{comp_count}/{total}")
        c3.metric("Avg Score", f"{avg_score:.1f}/5" if avg_score else "N/A")
        c4.progress(pct)


# ══════════════════════════════════════════════
# PAGE: ADD NEW WI
# ══════════════════════════════════════════════
elif page == "➕ Add New WI":
    st.markdown("## ➕ Add New Work Instruction")
    st.caption("Astuce : ce formulaire est aussi accessible en un clic depuis n'importe quelle page via le bouton orange flottant en bas à droite.")

    with st.form("add_wi_form"):
        wi_name = st.text_input("WI Name *", placeholder="e.g. Open Vessel")
        phase   = st.selectbox("Phase / Category *", VALID_PHASES)
        sop     = st.text_input(
            "🔗 Lien Google Drive du fichier SOP (optionnel)",
            placeholder="https://drive.google.com/file/d/.../view?usp=sharing",
            help="Mets le fichier SOP sur Google Drive, clique sur 'Partager' → 'Copier le lien', puis colle-le ici.",
        )
        confirmed = st.checkbox("Manager Confirmed")
        submitted = st.form_submit_button("➕ Create WI", type="primary")

    if sop.strip():
        render_drive_preview(sop.strip(), key_prefix="addpage")

    if submitted:
        new_id, result = create_new_wi(wi_name, phase, sop, confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Created: **{new_id}** — {result} ({phase})")


# ══════════════════════════════════════════════
# PAGE: MANAGE BINÔMES
# ══════════════════════════════════════════════
elif page == "👥 Manage Binômes":
    st.markdown("## 👥 Manage Binômes / Groups")

    # List existing
    st.markdown("### Current Binômes")
    for i, b in enumerate(st.session_state.binomes):
        c1, c2, c3 = st.columns([2, 3, 1])
        c1.text_input("Name", b["name"], key=f"bn_{i}", disabled=True)
        new_members = c2.text_input("Members", b["members"], key=f"bm_{i}")
        if c3.button("Save", key=f"sb_{i}"):
            st.session_state.binomes[i]["members"] = new_members
            st.success("Saved!")
            st.rerun()

    st.divider()
    st.markdown("### Add New Binôme")
    with st.form("add_binome_form"):
        nb_name    = st.text_input("Binôme Name *", placeholder="e.g. Binôme 6")
        nb_members = st.text_input("Members *", placeholder="e.g. Ali & Omar")
        sub_b = st.form_submit_button("➕ Add Binôme", type="primary")

    if sub_b:
        if not nb_name.strip() or not nb_members.strip():
            st.error("Both fields are required.")
        else:
            st.session_state.binomes.append({"name": nb_name.strip(), "members": nb_members.strip()})
            for w in st.session_state.wis:
                set_progress(w["id"], nb_name.strip(), {"status": "Not Started", "score": "", "notes": ""})
            st.success(f"Added: {nb_name}")
            st.rerun()


# ══════════════════════════════════════════════
# PAGE: EXPORT
# ══════════════════════════════════════════════
elif page == "📤 Export":
    st.markdown("## 📤 Export Data")

    st.info("Download the full dashboard data as an Excel file.")
    excel_bytes = export_to_excel()
    fname = f"WMCT_WI_Dashboard_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="⬇️ Download Excel Export",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()
    st.markdown("### Export incomplete WIs report")
    inc = incomplete_wis()
    if inc:
        df_inc = pd.DataFrame(inc)
        csv = df_inc.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Incomplete WIs (CSV)", csv, "incomplete_wis.csv", "text/csv")
    else:
        st.success("✅ No incomplete WIs!")

    st.divider()
    st.markdown("### Current data snapshot")
    tab1, tab2, tab3 = st.tabs(["WI Register", "Progress Data", "Category Stats"])
    with tab1:
        st.dataframe(pd.DataFrame(st.session_state.wis), use_container_width=True, hide_index=True)
    with tab2:
        rows = []
        for w in st.session_state.wis:
            for b in st.session_state.binomes:
                p = get_progress(w["id"], b["name"])
                rows.append({"WI": w["id"], "Binôme": b["name"], **p})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    with tab3:
        st.dataframe(pd.DataFrame(compute_category_stats()), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# FLOATING "ADD NEW WI" BUTTON — visible on every page
# ══════════════════════════════════════════════
with st.container(key="fab_container"):
    if st.button("➕ Add New WI", key="fab_add_wi"):
        add_wi_dialog()
