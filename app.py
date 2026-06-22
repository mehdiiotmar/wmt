"""
WMCT WI Dashboard — Streamlit Application
With Excel Import/Export for persistent save (Google Drive workflow)
"""

import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import io

st.set_page_config(
    page_title="WMCT WI Dashboard",
    page_icon="⭐",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

def create_new_wi(name, phase, sop_link, sop_file, confirmed):
    if not name or not name.strip():
        return None, "Le nom de la WI ne peut pas être vide."
    new_id = next_wi_id()
    title = f"WI: {name.strip()}"
    new_wi = {
        "id": new_id, "title": title, "phase": phase,
        "confirmed": confirmed, "sop": sop_link or "",
        "sop_file_name": None, "sop_file_bytes": None,
    }
    if sop_file is not None:
        new_wi["sop_file_name"] = sop_file.name
        new_wi["sop_file_bytes"] = sop_file.getvalue()
    st.session_state.wis.append(new_wi)
    for b in st.session_state.binomes:
        set_progress(new_id, b["name"], {"status": "Not Started", "score": "", "notes": ""})
    return new_id, title

# ─────────────────────────────────────────────
# IMPORT FROM EXCEL
# ─────────────────────────────────────────────
def import_from_excel(uploaded_file):
    """
    Reads the Excel save file and restores wis, binomes, and progress.
    Returns (success: bool, message: str)
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # ── 1. WI Register ──────────────────────────────────
        if "WI_Register" not in sheet_names:
            return False, "Feuille 'WI_Register' introuvable dans ce fichier."
        df_wi = pd.read_excel(xls, sheet_name="WI_Register")
        required_wi_cols = {"id", "title", "phase", "confirmed", "sop"}
        if not required_wi_cols.issubset(set(df_wi.columns)):
            return False, f"Colonnes manquantes dans WI_Register. Attendu: {required_wi_cols}"

        new_wis = []
        for _, row in df_wi.iterrows():
            wi = {
                "id":        str(row["id"]),
                "title":     str(row["title"]),
                "phase":     str(row["phase"]) if str(row["phase"]) in VALID_PHASES else VALID_PHASES[0],
                "confirmed": bool(row["confirmed"]) if pd.notna(row["confirmed"]) else False,
                "sop":       str(row["sop"]) if pd.notna(row["sop"]) else "",
                "sop_file_name": None,
                "sop_file_bytes": None,
            }
            new_wis.append(wi)

        # ── 2. Progress Data ─────────────────────────────────
        if "Progress_Data" not in sheet_names:
            return False, "Feuille 'Progress_Data' introuvable dans ce fichier."
        df_prog = pd.read_excel(xls, sheet_name="Progress_Data")

        # Detect binôme names from column headers: "<binome> Status"
        binome_names = []
        for col in df_prog.columns:
            if col.endswith(" Status"):
                bname = col[: -len(" Status")]
                binome_names.append(bname)

        if not binome_names:
            return False, "Aucun binôme trouvé dans Progress_Data (colonnes '<Binôme> Status' attendues)."

        new_binomes = [{"name": bn, "members": bn} for bn in binome_names]

        # Try to recover member names from Binomes sheet if present
        if "Binomes" in sheet_names:
            df_bin = pd.read_excel(xls, sheet_name="Binomes")
            if {"name", "members"}.issubset(set(df_bin.columns)):
                member_map = {str(r["name"]): str(r["members"]) for _, r in df_bin.iterrows()}
                for b in new_binomes:
                    if b["name"] in member_map:
                        b["members"] = member_map[b["name"]]

        new_progress = {}
        for _, row in df_prog.iterrows():
            wi_id = str(row["WI ID"])
            for bn in binome_names:
                status_col = f"{bn} Status"
                score_col  = f"{bn} Score"
                notes_col  = f"{bn} Notes"
                status = str(row[status_col]) if status_col in df_prog.columns and pd.notna(row.get(status_col)) else "Not Started"
                score  = str(int(row[score_col])) if score_col in df_prog.columns and pd.notna(row.get(score_col)) else ""
                notes  = str(row[notes_col]) if notes_col in df_prog.columns and pd.notna(row.get(notes_col)) else ""
                if status not in STATUS_OPTIONS:
                    status = "Not Started"
                key = str((wi_id, bn))
                new_progress[key] = {"status": status, "score": score, "notes": notes}

        # ── 3. Commit to session state ───────────────────────
        st.session_state.wis      = new_wis
        st.session_state.binomes  = new_binomes
        st.session_state.progress = new_progress

        return True, f"✅ Import réussi — {len(new_wis)} WIs, {len(new_binomes)} binômes chargés."

    except Exception as e:
        return False, f"Erreur lors de l'import : {e}"

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
        if w.get("sop_file_name"):
            sop = f"📊 {w['sop_file_name']}"
        elif w["sop"]:
            sop = w["sop"]
        else:
            sop = "—"
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
        wi_rows = []
        for w in st.session_state.wis:
            wi_rows.append({
                "id": w["id"], "title": w["title"], "phase": w["phase"],
                "confirmed": w["confirmed"], "sop": w["sop"],
            })
        pd.DataFrame(wi_rows).to_excel(writer, sheet_name="WI_Register", index=False)

        # Binomes (so member names survive round-trip)
        pd.DataFrame(st.session_state.binomes).to_excel(writer, sheet_name="Binomes", index=False)

        # Progress Data
        prog_rows = []
        for w in st.session_state.wis:
            row = {"WI ID": w["id"], "Title": w["title"], "Phase": w["phase"]}
            for b in st.session_state.binomes:
                p = get_progress(w["id"], b["name"])
                row[f"{b['name']} Status"] = p["status"]
                row[f"{b['name']} Score"]  = p["score"]
                row[f"{b['name']} Notes"]  = p["notes"]
            prog_rows.append(row)
        pd.DataFrame(prog_rows).to_excel(writer, sheet_name="Progress_Data", index=False)

        # Summary
        pd.DataFrame(compute_wi_table()).to_excel(writer, sheet_name="Dashboard_Summary", index=False)
        pd.DataFrame(compute_category_stats()).to_excel(writer, sheet_name="Insights", index=False)
    return buf.getvalue()

def render_sop_file_preview(wi, key_prefix):
    if wi.get("sop_file_bytes"):
        with st.expander(f"📊 Aperçu SOP — {wi['sop_file_name']}", expanded=False):
            try:
                sop_df = pd.read_excel(io.BytesIO(wi["sop_file_bytes"]))
                st.dataframe(sop_df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.warning(f"Impossible d'afficher l'aperçu : {e}")
            st.download_button(
                "⬇️ Télécharger ce fichier SOP",
                data=wi["sop_file_bytes"],
                file_name=wi["sop_file_name"],
                key=f"dl_sop_{key_prefix}",
            )

LOGO_SVG_RAW = """<svg width="46" height="46" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="wmctGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#2e75b6"/>
      <stop offset="100%" stop-color="#f59c20"/>
    </linearGradient>
  </defs>
  <circle cx="32" cy="32" r="31" fill="#0f2b46" stroke="url(#wmctGrad)" stroke-width="2"/>
  <rect x="14" y="34" width="9" height="9" fill="#f59c20" rx="1"/>
  <rect x="24" y="34" width="9" height="9" fill="#2e75b6" rx="1"/>
  <rect x="34" y="34" width="9" height="9" fill="#00b050" rx="1"/>
  <rect x="44" y="34" width="6" height="9" fill="#c00000" rx="1"/>
  <path d="M20 30 L20 16 L30 16 L36 24 L36 30" stroke="white" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="20" cy="16" r="2.2" fill="#ffd966"/>
  <path d="M10 48 H54" stroke="#cfe2f3" stroke-width="2" stroke-linecap="round"/>
</svg>"""

import base64
_logo_b64 = base64.b64encode(LOGO_SVG_RAW.encode()).decode()
LOGO_IMG_TAG = f'<img src="data:image/svg+xml;base64,{_logo_b64}" width="46" height="46" style="vertical-align:middle;"/>'

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Poppins', sans-serif; }

  .wmct-header {
    background: linear-gradient(135deg, #0f2b46 0%, #1f4e79 55%, #2e75b6 100%);
    color: white; padding: 22px 28px; border-radius: 14px;
    margin-bottom: 24px; display: flex; align-items: center; gap: 16px;
    box-shadow: 0 4px 18px rgba(15,43,70,.25);
  }
  .wmct-header .logo-box { flex-shrink: 0; }
  .wmct-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; letter-spacing: .3px; }
  .wmct-header p  { margin: 4px 0 0; opacity: .85; font-size: .88rem; }

  .sidebar-logo { display:flex; align-items:center; gap:10px; margin-bottom:4px; }
  .sidebar-logo .brand { font-size:1.05rem; font-weight:700; color:white; line-height:1.1; }
  .sidebar-logo .brand small { display:block; font-weight:400; opacity:.7; font-size:.7rem; }

  /* Import banner */
  .import-banner {
    background: linear-gradient(135deg, #1a3a5c 0%, #2e75b6 100%);
    border-radius: 12px; padding: 16px 20px; margin-bottom: 18px;
    border: 1.5px dashed #f59c20;
  }
  .import-banner h3 { color: #f59c20 !important; margin: 0 0 6px; font-size: .95rem; }
  .import-banner p  { color: #cde !important; margin: 0; font-size: .82rem; }

  .kpi-card {
    flex: 1; min-width: 130px;
    background: white; border-radius: 12px;
    padding: 18px 20px; text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,.07);
    border-top: 4px solid #2e75b6;
  }
  .kpi-card .val { font-size: 2rem; font-weight: 700; color: #1f4e79; }
  .kpi-card .lbl { font-size: .78rem; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: .5px; }

  .section-title {
    font-size: 1.1rem; font-weight: 700; color: #1f4e79;
    border-left: 4px solid #2e75b6; padding-left: 10px;
    margin: 28px 0 14px;
  }

  section[data-testid="stSidebar"] { background: linear-gradient(180deg,#0f2b46 0%, #1f4e79 100%); }
  section[data-testid="stSidebar"] * { color: white !important; }
  section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.15); }

  div.st-key-fab_container {
    position: fixed; bottom: 22px; right: 26px; z-index: 9999; width: auto;
  }
  div.st-key-fab_container button {
    background: linear-gradient(135deg, #f59c20 0%, #e0830a 100%) !important;
    color: white !important; border: none !important;
    border-radius: 50px !important; padding: 14px 26px !important;
    font-weight: 700 !important; font-size: .95rem !important;
    box-shadow: 0 6px 18px rgba(229,140,10,.45) !important;
  }
</style>
""", unsafe_allow_html=True)

@st.dialog("➕ Ajouter une nouvelle WI", width="large")
def add_wi_dialog():
    st.caption("Ce formulaire est accessible depuis n'importe quelle page via le bouton flottant.")
    qa_name = st.text_input("Nom de la WI *", placeholder="ex. Open Vessel", key="qa_name")
    qa_phase = st.selectbox("Phase / Catégorie *", VALID_PHASES, key="qa_phase")
    qa_link = st.text_input("Lien SOP (optionnel)", key="qa_link")
    qa_file = st.file_uploader("📎 Fichier Excel SOP", type=["xlsx", "xls"], key="qa_file")
    qa_confirmed = st.checkbox("✅ Confirmé par le manager", key="qa_confirmed")
    c1, c2 = st.columns(2)
    if c1.button("➕ Créer la WI", type="primary", use_container_width=True, key="qa_submit"):
        new_id, result = create_new_wi(qa_name, qa_phase, qa_link, qa_file, qa_confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Créée : **{new_id}** — {result}")
            st.rerun()
    if c2.button("Annuler", use_container_width=True, key="qa_cancel"):
        st.rerun()

init_state()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-logo">
        {LOGO_IMG_TAG}
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
      <div class="logo-box">{LOGO_IMG_TAG}</div>
      <div>
        <h1>WMCT — WEST MED CONTAINER TERMINAL</h1>
        <p>CATOS Training Phase  |  WI Progress Dashboard  |  Pre-Go-Live Training</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # How-to-save callout
    st.info(
        "💡 **Astuce sauvegarde** — Cliquez sur **⬇️ Sauvegarder** dans la barre latérale après chaque session "
        "et gardez le fichier Excel sur Google Drive. La prochaine fois, rechargez-le via **📂 Charger une sauvegarde**.",
        icon=None,
    )

    kpis = compute_kpis()
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.metric("📋 Total WIs", kpis["total"])
    with c2: st.metric("✅ Completed", kpis["completed"])
    with c3: st.metric("🔄 In Progress", kpis["in_progress"])
    with c4: st.metric("⏳ Not Confirmed", kpis["not_confirmed"])
    with c5: st.metric("🔐 Confirmed", kpis["confirmed"])

    st.markdown(f"**Overall Completion: {kpis['pct']:.0%}**")
    st.progress(kpis["pct"])

    st.markdown('<div class="section-title">BINÔME PROGRESS OVERVIEW</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame(compute_binome_stats()), use_container_width=True, hide_index=True)

    cols = st.columns(len(st.session_state.binomes))
    for i, b in enumerate(st.session_state.binomes):
        bn = b["name"]
        total = len(st.session_state.wis)
        comp = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Completed")
        pct = comp / total if total else 0
        with cols[i]:
            st.caption(f"**{bn}**\n{b['members']}")
            st.progress(pct, text=f"{pct:.0%}")

    st.markdown('<div class="section-title">WI STATUS TABLE</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame(compute_wi_table()), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# PAGE: WI REGISTER
# ══════════════════════════════════════════════
elif page == "📋 WI Register":
    st.markdown("## 📋 WI Register")
    st.caption("All Work Instructions — edit titles, phases, manager confirmation and SOP links.")
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
                new_sop = st.text_input("SOP Link (optionnel)", w["sop"], key=f"sop_{i}")
            new_sop_file = st.file_uploader("📎 Fichier Excel SOP", type=["xlsx", "xls"], key=f"sopfile_{i}")
            if st.button("💾 Save", key=f"save_wi_{i}"):
                st.session_state.wis[i].update({
                    "title": new_title, "phase": new_phase,
                    "confirmed": new_confirmed, "sop": new_sop,
                })
                if new_sop_file is not None:
                    st.session_state.wis[i]["sop_file_name"] = new_sop_file.name
                    st.session_state.wis[i]["sop_file_bytes"] = new_sop_file.getvalue()
                st.success("Saved!")
                st.rerun()
            render_sop_file_preview(w, key_prefix=f"reg_{i}")
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
    df_cat = pd.DataFrame(compute_category_stats())
    df_cat_active = df_cat[df_cat["Total WIs"] > 0]
    st.markdown('<div class="section-title">WIs BY CATEGORY</div>', unsafe_allow_html=True)
    st.dataframe(df_cat_active, use_container_width=True, hide_index=True)
    if not df_cat_active.empty:
        st.bar_chart(df_cat_active.set_index("Category")[["Completed", "Not Completed"]], color=["#00b050", "#c00000"])
    st.markdown('<div class="section-title">WI À SUIVRE — NON COMPLÈTES</div>', unsafe_allow_html=True)
    inc = incomplete_wis()
    if inc:
        df_inc = pd.DataFrame(inc)
        for phase in sorted(df_inc["Phase"].unique()):
            st.markdown(f"**{phase}**")
            st.dataframe(df_inc[df_inc["Phase"] == phase][["WI ID", "Title", "Status"]], use_container_width=True, hide_index=True)
        st.info(f"Total WI non complètes : **{len(inc)}**")
    else:
        st.success("✅ Aucune WI incomplète — tout est à jour !")
    st.markdown('<div class="section-title">BINÔME PERFORMANCE</div>', unsafe_allow_html=True)
    for b in st.session_state.binomes:
        bn = b["name"]
        scores = []
        for w in st.session_state.wis:
            p = get_progress(w["id"], bn)
            if p["score"]:
                try: scores.append(int(p["score"]))
                except: pass
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
    with st.form("add_wi_form"):
        wi_name   = st.text_input("WI Name *", placeholder="e.g. Open Vessel")
        phase     = st.selectbox("Phase / Category *", VALID_PHASES)
        sop       = st.text_input("SOP Link (optionnel)")
        sop_file  = st.file_uploader("📎 Fichier Excel SOP", type=["xlsx", "xls"])
        confirmed = st.checkbox("Manager Confirmed")
        submitted = st.form_submit_button("➕ Create WI", type="primary")
    if submitted:
        new_id, result = create_new_wi(wi_name, phase, sop, sop_file, confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Created: **{new_id}** — {result} ({phase})")


# ══════════════════════════════════════════════
# PAGE: MANAGE BINÔMES
# ══════════════════════════════════════════════
elif page == "👥 Manage Binômes":
    st.markdown("## 👥 Manage Binômes / Groups")
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
    st.markdown("## 📤 Export & Sauvegarde")

    # ── SAVE/RESTORE GUIDE ──────────────────────────────
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1a3a5c,#2e75b6);border-radius:12px;padding:18px 22px;border:1.5px dashed #f59c20;margin-bottom:20px;">
      <h3 style="color:#f59c20;margin:0 0 8px;">📖 Comment sauvegarder & restaurer vos données</h3>
      <ol style="color:#cde;margin:0;padding-left:18px;font-size:.88rem;line-height:1.8;">
        <li><strong style="color:white;">Exportez</strong> le fichier Excel ci-dessous après chaque session de travail.</li>
        <li>Enregistrez ce fichier sur <strong style="color:#f59c20;">Google Drive</strong> (ou clé USB / bureau).</li>
        <li>La prochaine fois que vous ouvrez l'application, utilisez <strong style="color:white;">📂 Charger une sauvegarde</strong> dans la barre latérale gauche pour tout restaurer en 1 clic.</li>
      </ol>
    </div>
    """, unsafe_allow_html=True)

    excel_bytes = export_to_excel()
    fname = f"WMCT_Save_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="⬇️ Télécharger la sauvegarde Excel",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.divider()
    st.markdown("### Restaurer depuis un fichier")
    st.caption("Vous pouvez aussi importer ici (même résultat que la barre latérale).")
    restore_file = st.file_uploader("📂 Choisir un fichier de sauvegarde", type=["xlsx", "xls"], key="export_page_import")
    if restore_file is not None:
        if st.button("🔄 Restaurer", type="primary", key="export_page_do_import"):
            ok, msg = import_from_excel(restore_file)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.divider()
    st.markdown("### Export données incomplètes (CSV)")
    inc = incomplete_wis()
    if inc:
        csv = pd.DataFrame(inc).to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ WIs incomplètes (CSV)", csv, "incomplete_wis.csv", "text/csv")
    else:
        st.success("✅ No incomplete WIs!")

    st.divider()
    st.markdown("### Aperçu des données actuelles")
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


# ── FLOATING BUTTON ──────────────────────────────────────────────
with st.container(key="fab_container"):
    if st.button("➕ Add New WI", key="fab_add_wi"):
        add_wi_dialog()
