"""
WMCT WI Dashboard — Streamlit Application
Modern Design + Interactive Visualizations (Plotly)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import os
from datetime import datetime
import io
import base64

st.set_page_config(
    page_title="WMCT WI Dashboard",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CONSTANTS
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

# Color palette
COLORS = {
    "completed": "#10b981",
    "in_progress": "#f59e0b",
    "not_started": "#6366f1",
    "danger": "#ef4444",
    "primary": "#0ea5e9",
    "dark": "#0f172a",
    "surface": "#1e293b",
    "card": "#0f172a",
    "text": "#f1f5f9",
    "muted": "#94a3b8",
}

PHASE_COLORS = {
    "Ship Planning": "#0ea5e9",
    "Berth Planning": "#6366f1",
    "Operation Management": "#f59e0b",
    "Yard Define": "#10b981",
    "Yard Planning": "#14b8a6",
    "Terminal Monitoring": "#ec4899",
    "Reefer Operation": "#8b5cf6",
    "DG Operation": "#ef4444",
    "Gate": "#f97316",
    "Security Management": "#64748b",
    "Planning": "#3b82f6",
    "Operations": "#22c55e",
    "General": "#a855f7",
    "CATOS Admin": "#06b6d4",
}

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

# ─────────────────────────────────────────────
# STATE
# ─────────────────────────────────────────────
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
# IMPORT / EXPORT
# ─────────────────────────────────────────────
def import_from_excel(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        if "WI_Register" not in sheet_names:
            return False, "Feuille 'WI_Register' introuvable."
        df_wi = pd.read_excel(xls, sheet_name="WI_Register")
        required_wi_cols = {"id", "title", "phase", "confirmed", "sop"}
        if not required_wi_cols.issubset(set(df_wi.columns)):
            return False, f"Colonnes manquantes dans WI_Register."
        new_wis = []
        for _, row in df_wi.iterrows():
            wi = {
                "id": str(row["id"]), "title": str(row["title"]),
                "phase": str(row["phase"]) if str(row["phase"]) in VALID_PHASES else VALID_PHASES[0],
                "confirmed": bool(row["confirmed"]) if pd.notna(row["confirmed"]) else False,
                "sop": str(row["sop"]) if pd.notna(row["sop"]) else "",
                "sop_file_name": None, "sop_file_bytes": None,
            }
            new_wis.append(wi)
        if "Progress_Data" not in sheet_names:
            return False, "Feuille 'Progress_Data' introuvable."
        df_prog = pd.read_excel(xls, sheet_name="Progress_Data")
        binome_names = [col[:-7] for col in df_prog.columns if col.endswith(" Status")]
        if not binome_names:
            return False, "Aucun binôme trouvé dans Progress_Data."
        new_binomes = [{"name": bn, "members": bn} for bn in binome_names]
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
                score_col = f"{bn} Score"
                notes_col = f"{bn} Notes"
                status = str(row[status_col]) if status_col in df_prog.columns and pd.notna(row.get(status_col)) else "Not Started"
                score = str(int(row[score_col])) if score_col in df_prog.columns and pd.notna(row.get(score_col)) else ""
                notes = str(row[notes_col]) if notes_col in df_prog.columns and pd.notna(row.get(notes_col)) else ""
                if status not in STATUS_OPTIONS:
                    status = "Not Started"
                new_progress[str((wi_id, bn))] = {"status": status, "score": score, "notes": notes}
        st.session_state.wis = new_wis
        st.session_state.binomes = new_binomes
        st.session_state.progress = new_progress
        return True, f"✅ Import réussi — {len(new_wis)} WIs, {len(new_binomes)} binômes chargés."
    except Exception as e:
        return False, f"Erreur lors de l'import : {e}"

def export_to_excel():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wi_rows = [{"id": w["id"], "title": w["title"], "phase": w["phase"], "confirmed": w["confirmed"], "sop": w["sop"]} for w in st.session_state.wis]
        pd.DataFrame(wi_rows).to_excel(writer, sheet_name="WI_Register", index=False)
        pd.DataFrame(st.session_state.binomes).to_excel(writer, sheet_name="Binomes", index=False)
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
    return buf.getvalue()

# ─────────────────────────────────────────────
# KPI COMPUTATIONS
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
    return {"total": total, "completed": completed, "in_progress": in_progress,
            "confirmed": confirmed, "not_confirmed": not_confirmed, "pct": pct}

def compute_binome_stats():
    rows = []
    for b in st.session_state.binomes:
        bn = b["name"]
        bm = b["members"]
        comp = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Completed")
        in_prog = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "In Progress")
        not_started = sum(1 for w in st.session_state.wis if get_progress(w["id"], bn)["status"] == "Not Started")
        total = len(st.session_state.wis)
        pct = comp / total if total else 0
        scores = []
        for w in st.session_state.wis:
            p = get_progress(w["id"], bn)
            if p["score"]:
                try: scores.append(int(p["score"]))
                except: pass
        avg_score = sum(scores) / len(scores) if scores else 0
        if pct >= 0.7:
            status = "On Track"
        elif pct >= 0.4:
            status = "At Risk"
        else:
            status = "Need Support"
        rows.append({"Binôme": bn, "Members": bm, "Completed": comp, "In Progress": in_prog,
                     "Not Started": not_started, "Total": total, "Pct": pct,
                     "Avg Score": avg_score, "Status": status})
    return rows

def compute_category_stats():
    cat_map = {}
    for w in st.session_state.wis:
        ph = w["phase"]
        if ph not in cat_map:
            cat_map[ph] = {"total": 0, "completed": 0, "in_progress": 0}
        cat_map[ph]["total"] += 1
        if wi_is_completed(w["id"]):
            cat_map[ph]["completed"] += 1
        elif any(get_progress(w["id"], b["name"])["status"] == "In Progress" for b in st.session_state.binomes):
            cat_map[ph]["in_progress"] += 1
    rows = []
    for phase in VALID_PHASES:
        d = cat_map.get(phase, {"total": 0, "completed": 0, "in_progress": 0})
        if d["total"] > 0:
            rows.append({
                "Category": phase,
                "Total WIs": d["total"],
                "Completed": d["completed"],
                "In Progress": d["in_progress"],
                "Not Started": d["total"] - d["completed"] - d["in_progress"],
                "Completion %": d["completed"] / d["total"] if d["total"] else 0,
            })
    return rows

def incomplete_wis():
    out = []
    for w in st.session_state.wis:
        if not wi_is_completed(w["id"]):
            statuses = [get_progress(w["id"], b["name"])["status"] for b in st.session_state.binomes]
            overall = "In Progress" if any(s == "In Progress" for s in statuses) else "Not Started"
            out.append({"WI ID": w["id"], "Title": w["title"], "Phase": w["phase"], "Status": overall})
    return out

# ─────────────────────────────────────────────
# PLOTLY CHART HELPERS
# ─────────────────────────────────────────────
plotly_layout = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#cbd5e1", size=12),
    margin=dict(l=10, r=10, t=40, b=10),
    showlegend=True,
    legend=dict(
        bgcolor="rgba(30,41,59,0.8)",
        bordercolor="rgba(148,163,184,0.2)",
        borderwidth=1,
        font=dict(color="#cbd5e1"),
    ),
)

def chart_overall_donut(kpis):
    labels = ["Completed", "In Progress", "Not Started"]
    in_prog = kpis["in_progress"]
    not_s = kpis["total"] - kpis["completed"] - in_prog
    values = [kpis["completed"], in_prog, max(not_s, 0)]
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.65,
        marker_colors=[COLORS["completed"], COLORS["in_progress"], "#334155"],
        textinfo="label+percent",
        textfont=dict(color="#f1f5f9", size=11),
        hovertemplate="<b>%{label}</b><br>%{value} WIs<br>%{percent}<extra></extra>",
    ))
    fig.add_annotation(
        text=f"<b>{kpis['pct']:.0%}</b><br><span style='font-size:10px'>Overall</span>",
        x=0.5, y=0.5, font=dict(size=18, color="#f1f5f9"), showarrow=False,
    )
    fig.update_layout(**plotly_layout, title=dict(text="Overall Completion", font=dict(color="#94a3b8", size=13)))
    return fig

def chart_binome_radar(binome_stats):
    categories = [b["Binôme"] for b in binome_stats]
    categories_closed = categories + [categories[0]]
    fig = go.Figure()
    # Completion ring
    pcts = [b["Pct"] * 100 for b in binome_stats] + [binome_stats[0]["Pct"] * 100]
    fig.add_trace(go.Scatterpolar(
        r=pcts, theta=categories_closed,
        fill="toself", name="Completion %",
        line_color=COLORS["completed"],
        fillcolor="rgba(16,185,129,0.2)",
        hovertemplate="<b>%{theta}</b><br>Completion: %{r:.1f}%<extra></extra>",
    ))
    # Score ring (normalized to 100)
    scores = [b["Avg Score"] / 5 * 100 for b in binome_stats] + [binome_stats[0]["Avg Score"] / 5 * 100]
    fig.add_trace(go.Scatterpolar(
        r=scores, theta=categories_closed,
        fill="toself", name="Avg Score /5",
        line_color=COLORS["in_progress"],
        fillcolor="rgba(245,158,11,0.15)",
        hovertemplate="<b>%{theta}</b><br>Score (norm): %{r:.1f}%<extra></extra>",
    ))
    fig.update_layout(
        **plotly_layout,
        polar=dict(
            bgcolor="rgba(15,23,42,0.5)",
            radialaxis=dict(visible=True, range=[0, 100], gridcolor="rgba(148,163,184,0.15)", tickfont=dict(color="#94a3b8", size=10)),
            angularaxis=dict(gridcolor="rgba(148,163,184,0.2)", tickfont=dict(color="#e2e8f0", size=11)),
        ),
        title=dict(text="Binôme Performance Radar", font=dict(color="#94a3b8", size=13)),
    )
    return fig

def chart_binome_stacked(binome_stats):
    binomes = [b["Binôme"] for b in binome_stats]
    completed = [b["Completed"] for b in binome_stats]
    in_prog = [b["In Progress"] for b in binome_stats]
    not_started = [b["Not Started"] for b in binome_stats]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Completed", x=binomes, y=completed,
                          marker_color=COLORS["completed"], hovertemplate="%{y} completed<extra></extra>"))
    fig.add_trace(go.Bar(name="In Progress", x=binomes, y=in_prog,
                          marker_color=COLORS["in_progress"], hovertemplate="%{y} in progress<extra></extra>"))
    fig.add_trace(go.Bar(name="Not Started", x=binomes, y=not_started,
                          marker_color="#334155", hovertemplate="%{y} not started<extra></extra>"))
    fig.update_layout(
        **plotly_layout,
        barmode="stack",
        title=dict(text="WI Status by Binôme", font=dict(color="#94a3b8", size=13)),
        xaxis=dict(gridcolor="rgba(148,163,184,0.1)", tickfont=dict(color="#94a3b8")),
        yaxis=dict(gridcolor="rgba(148,163,184,0.1)", tickfont=dict(color="#94a3b8")),
    )
    return fig

def chart_category_horizontal(cat_stats):
    if not cat_stats:
        return None
    cats = [r["Category"] for r in cat_stats]
    completed = [r["Completed"] for r in cat_stats]
    in_prog = [r["In Progress"] for r in cat_stats]
    not_started = [r["Not Started"] for r in cat_stats]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Completed", y=cats, x=completed, orientation="h",
                          marker_color=COLORS["completed"]))
    fig.add_trace(go.Bar(name="In Progress", y=cats, x=in_prog, orientation="h",
                          marker_color=COLORS["in_progress"]))
    fig.add_trace(go.Bar(name="Not Started", y=cats, x=not_started, orientation="h",
                          marker_color="#334155"))
    fig.update_layout(
        **plotly_layout,
        barmode="stack",
        height=max(250, len(cat_stats) * 45),
        title=dict(text="WI Status by Category", font=dict(color="#94a3b8", size=13)),
        xaxis=dict(gridcolor="rgba(148,163,184,0.1)", tickfont=dict(color="#94a3b8")),
        yaxis=dict(gridcolor="rgba(148,163,184,0.0)", tickfont=dict(color="#e2e8f0", size=11)),
    )
    return fig

def chart_heatmap():
    """Status heatmap: WIs vs Binomes"""
    wis = st.session_state.wis
    binomes = st.session_state.binomes
    status_map = {"Not Started": 0, "In Progress": 1, "Completed": 2}
    z = []
    for w in wis:
        row = []
        for b in binomes:
            p = get_progress(w["id"], b["name"])
            row.append(status_map.get(p["status"], 0))
        z.append(row)

    wi_labels = [f"{w['id']}" for w in wis]
    binome_labels = [b["name"] for b in binomes]
    text = []
    for w in wis:
        row = []
        for b in binomes:
            p = get_progress(w["id"], b["name"])
            sc = f" ({p['score']}★)" if p["score"] else ""
            row.append(p["status"] + sc)
        text.append(row)

    fig = go.Figure(go.Heatmap(
        z=z, x=binome_labels, y=wi_labels,
        text=text, hovertemplate="<b>%{y} × %{x}</b><br>%{text}<extra></extra>",
        colorscale=[[0, "#1e293b"], [0.5, "#f59e0b"], [1, "#10b981"]],
        showscale=False,
        texttemplate="",
    ))
    fig.update_layout(
        **plotly_layout,
        height=max(250, len(wis) * 45),
        title=dict(text="Status Heatmap — WI × Binôme", font=dict(color="#94a3b8", size=13)),
        xaxis=dict(side="top", tickfont=dict(color="#e2e8f0", size=11)),
        yaxis=dict(tickfont=dict(color="#94a3b8", size=10), autorange="reversed"),
        margin=dict(l=90, r=10, t=70, b=10),
    )
    return fig

def chart_score_distribution():
    """Box plot of scores per binôme"""
    data = []
    for b in st.session_state.binomes:
        for w in st.session_state.wis:
            p = get_progress(w["id"], b["name"])
            if p["score"]:
                try:
                    data.append({"Binôme": b["name"], "Score": int(p["score"]), "WI": w["id"]})
                except:
                    pass
    if not data:
        return None
    df = pd.DataFrame(data)
    fig = go.Figure()
    for b in st.session_state.binomes:
        bn = b["name"]
        subset = df[df["Binôme"] == bn]["Score"].tolist()
        if subset:
            fig.add_trace(go.Box(
                y=subset, name=bn,
                boxpoints="all", jitter=0.4, pointpos=-1.8,
                marker=dict(size=5, opacity=0.7),
                hovertemplate=f"<b>{bn}</b><br>Score: %{{y}}<extra></extra>",
            ))
    fig.update_layout(
        **plotly_layout,
        title=dict(text="Score Distribution by Binôme", font=dict(color="#94a3b8", size=13)),
        yaxis=dict(range=[0, 5.5], gridcolor="rgba(148,163,184,0.1)", tickfont=dict(color="#94a3b8"), dtick=1),
        xaxis=dict(tickfont=dict(color="#e2e8f0")),
    )
    return fig

def chart_completion_gauge(pct):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=pct * 100,
        number=dict(suffix="%", font=dict(color="#f1f5f9", size=36)),
        gauge=dict(
            axis=dict(range=[0, 100], tickcolor="#94a3b8", tickfont=dict(color="#94a3b8")),
            bar=dict(color=COLORS["completed"], thickness=0.3),
            bgcolor="rgba(30,41,59,0.5)",
            borderwidth=0,
            steps=[
                dict(range=[0, 40], color="rgba(239,68,68,0.15)"),
                dict(range=[40, 70], color="rgba(245,158,11,0.15)"),
                dict(range=[70, 100], color="rgba(16,185,129,0.15)"),
            ],
            threshold=dict(line=dict(color="#10b981", width=3), thickness=0.85, value=70),
        ),
    ))
    fig.update_layout(
        **plotly_layout,
        height=200,
        title=dict(text="Overall Progress", font=dict(color="#94a3b8", size=13)),
        margin=dict(l=20, r=20, t=40, b=10),
    )
    return fig

# ─────────────────────────────────────────────
# LOGO
# ─────────────────────────────────────────────
LOGO_SVG_RAW = """<svg width="46" height="46" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="wmctGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0ea5e9"/>
      <stop offset="100%" stop-color="#6366f1"/>
    </linearGradient>
  </defs>
  <circle cx="32" cy="32" r="31" fill="#0f172a" stroke="url(#wmctGrad)" stroke-width="2"/>
  <rect x="14" y="34" width="9" height="9" fill="#0ea5e9" rx="1"/>
  <rect x="24" y="34" width="9" height="9" fill="#6366f1" rx="1"/>
  <rect x="34" y="34" width="9" height="9" fill="#10b981" rx="1"/>
  <rect x="44" y="34" width="6" height="9" fill="#f59e0b" rx="1"/>
  <path d="M20 30 L20 16 L30 16 L36 24 L36 30" stroke="white" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="20" cy="16" r="2.2" fill="#f59e0b"/>
  <path d="M10 48 H54" stroke="#334155" stroke-width="2" stroke-linecap="round"/>
</svg>"""
_logo_b64 = base64.b64encode(LOGO_SVG_RAW.encode()).decode()
LOGO_IMG_TAG = f'<img src="data:image/svg+xml;base64,{_logo_b64}" width="42" height="42" style="vertical-align:middle;"/>'

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, sans-serif;
    background: #0f172a;
  }

  /* ── Page background ── */
  .stApp { background: #0f172a; }
  .main .block-container { padding: 1.5rem 2rem 4rem; max-width: 1400px; }

  /* ── Sidebar ── */
  section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #020617 0%, #0f172a 100%) !important;
    border-right: 1px solid rgba(148,163,184,0.1);
  }
  section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
  section[data-testid="stSidebar"] hr { border-color: rgba(148,163,184,0.1) !important; }
  section[data-testid="stSidebar"] .stRadio > div { gap: 0.25rem; }
  section[data-testid="stSidebar"] .stRadio label {
    padding: 0.45rem 0.75rem !important;
    border-radius: 8px !important;
    transition: background 0.15s !important;
  }
  section[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(148,163,184,0.1) !important;
  }

  /* ── Header ── */
  .dash-header {
    background: linear-gradient(135deg, #0c1a2e 0%, #0f2b46 50%, #162d4a 100%);
    border: 1px solid rgba(14,165,233,0.2);
    border-radius: 16px;
    padding: 20px 28px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 18px;
    position: relative;
    overflow: hidden;
  }
  .dash-header::before {
    content: '';
    position: absolute;
    top: -50%; right: -10%;
    width: 300px; height: 300px;
    background: radial-gradient(circle, rgba(14,165,233,0.12) 0%, transparent 70%);
    pointer-events: none;
  }
  .dash-header h1 {
    margin: 0; font-size: 1.3rem; font-weight: 700;
    color: #f1f5f9; letter-spacing: 0.3px;
  }
  .dash-header p { margin: 3px 0 0; color: #94a3b8; font-size: 0.82rem; }

  /* ── KPI Cards ── */
  .kpi-grid { display: flex; gap: 14px; margin: 20px 0; flex-wrap: wrap; }
  .kpi-card {
    flex: 1; min-width: 130px;
    background: #1e293b;
    border: 1px solid rgba(148,163,184,0.1);
    border-radius: 14px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
    transition: transform 0.2s, border-color 0.2s;
  }
  .kpi-card:hover { transform: translateY(-2px); border-color: rgba(14,165,233,0.4); }
  .kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    background: var(--accent, #0ea5e9);
    border-radius: 14px 14px 0 0;
  }
  .kpi-card .val {
    font-size: 2rem; font-weight: 700;
    color: var(--accent, #f1f5f9);
    line-height: 1;
  }
  .kpi-card .lbl {
    font-size: 0.72rem; color: #64748b;
    text-transform: uppercase; letter-spacing: 0.8px;
    margin-top: 6px;
  }
  .kpi-card .icon { font-size: 1.2rem; margin-bottom: 6px; opacity: 0.8; }

  /* ── Section titles ── */
  .section-title {
    font-size: 0.72rem;
    font-weight: 600;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin: 32px 0 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .section-title::after {
    content: '';
    flex: 1;
    height: 1px;
    background: rgba(148,163,184,0.1);
  }

  /* ── Status badges ── */
  .badge {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
  }
  .badge-completed { background: rgba(16,185,129,0.15); color: #10b981; }
  .badge-progress  { background: rgba(245,158,11,0.15);  color: #f59e0b; }
  .badge-pending   { background: rgba(99,102,241,0.15); color: #818cf8; }
  .badge-danger    { background: rgba(239,68,68,0.15);  color: #f87171; }

  /* ── Binome cards ── */
  .binome-card {
    background: #1e293b;
    border: 1px solid rgba(148,163,184,0.1);
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 0;
  }
  .binome-name { font-weight: 600; color: #f1f5f9; font-size: 0.9rem; }
  .binome-members { color: #64748b; font-size: 0.78rem; margin: 2px 0 8px; }
  .binome-pct { font-size: 1.4rem; font-weight: 700; color: #0ea5e9; }

  /* ── Progress bar ── */
  .prog-bar-bg {
    height: 6px; background: #334155;
    border-radius: 10px; overflow: hidden; margin-top: 6px;
  }
  .prog-bar-fill {
    height: 100%; border-radius: 10px;
    background: linear-gradient(90deg, #10b981, #0ea5e9);
    transition: width 0.4s ease;
  }

  /* ── Streamlit overrides ── */
  .stMetric { background: #1e293b !important; border-radius: 12px !important; padding: 14px 18px !important; border: 1px solid rgba(148,163,184,0.1) !important; }
  .stMetric label { color: #64748b !important; font-size: 0.72rem !important; text-transform: uppercase !important; letter-spacing: 0.8px !important; }
  .stMetric [data-testid="stMetricValue"] { color: #f1f5f9 !important; font-weight: 700 !important; font-size: 1.7rem !important; }
  div[data-testid="stDataFrame"] { border-radius: 12px !important; overflow: hidden !important; }
  .stDataFrame thead th { background: #1e293b !important; color: #94a3b8 !important; }
  .stDataFrame tbody td { background: #0f172a !important; color: #e2e8f0 !important; }
  .stExpander { background: #1e293b !important; border: 1px solid rgba(148,163,184,0.1) !important; border-radius: 12px !important; }
  .stExpander header { color: #f1f5f9 !important; }
  .stTabs [data-baseweb="tab-list"] { background: #1e293b !important; border-radius: 10px !important; padding: 4px !important; gap: 2px !important; }
  .stTabs [data-baseweb="tab"] { color: #94a3b8 !important; border-radius: 8px !important; }
  .stTabs [aria-selected="true"] { background: #0ea5e9 !important; color: white !important; }
  .stSelectbox > div { background: #1e293b !important; border-color: rgba(148,163,184,0.2) !important; border-radius: 8px !important; color: #f1f5f9 !important; }
  .stTextInput > div > div { background: #1e293b !important; border-color: rgba(148,163,184,0.2) !important; border-radius: 8px !important; color: #f1f5f9 !important; }
  .stButton button {
    background: linear-gradient(135deg, #0ea5e9 0%, #6366f1 100%) !important;
    border: none !important;
    border-radius: 8px !important;
    color: white !important;
    font-weight: 600 !important;
    transition: opacity 0.2s !important;
  }
  .stButton button:hover { opacity: 0.85 !important; }
  .stButton [kind="secondary"] button {
    background: #334155 !important;
    background-image: none !important;
  }
  .stInfo { background: rgba(14,165,233,0.1) !important; border-color: rgba(14,165,233,0.3) !important; border-radius: 10px !important; color: #94a3b8 !important; }
  .stSuccess { background: rgba(16,185,129,0.1) !important; border-color: rgba(16,185,129,0.3) !important; border-radius: 10px !important; color: #94a3b8 !important; }
  .stError { background: rgba(239,68,68,0.1) !important; border-color: rgba(239,68,68,0.3) !important; border-radius: 10px !important; color: #f87171 !important; }
  .stForm { background: #1e293b !important; border-color: rgba(148,163,184,0.1) !important; border-radius: 14px !important; padding: 20px !important; }
  .stDivider { border-color: rgba(148,163,184,0.1) !important; }
  .stCheckbox label { color: #e2e8f0 !important; }
  div[data-testid="stFileUploader"] { background: #1e293b !important; border-color: rgba(148,163,184,0.2) !important; border-radius: 10px !important; }
  .stDownloadButton button {
    background: linear-gradient(135deg, #10b981, #0ea5e9) !important;
    border: none !important;
    border-radius: 8px !important;
    color: white !important;
    font-weight: 600 !important;
  }

  /* ── FAB ── */
  div.st-key-fab_container {
    position: fixed; bottom: 24px; right: 28px; z-index: 9999; width: auto !important;
  }
  div.st-key-fab_container button {
    background: linear-gradient(135deg, #0ea5e9, #6366f1) !important;
    border: none !important;
    border-radius: 50px !important;
    padding: 14px 28px !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    box-shadow: 0 8px 24px rgba(99,102,241,0.4) !important;
    letter-spacing: 0.3px !important;
  }

  /* ── Sidebar brand ── */
  .sidebar-brand { display: flex; align-items: center; gap: 12px; padding: 4px 0 12px; }
  .sidebar-brand .brand-text { line-height: 1.2; }
  .sidebar-brand .brand-name { font-size: 1rem; font-weight: 700; color: #f1f5f9 !important; }
  .sidebar-brand .brand-sub  { font-size: 0.72rem; color: #475569 !important; }

  /* ── WI row in register ── */
  .wi-row {
    background: #1e293b;
    border: 1px solid rgba(148,163,184,0.08);
    border-radius: 12px;
    padding: 14px 18px;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 12px;
    transition: border-color 0.2s;
  }
  .wi-row:hover { border-color: rgba(14,165,233,0.3); }
  .wi-id { font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #64748b; min-width: 80px; }
  .wi-title { font-weight: 500; color: #f1f5f9; flex: 1; font-size: 0.9rem; }
  .wi-phase { font-size: 0.72rem; color: #94a3b8; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# DIALOG
# ─────────────────────────────────────────────
@st.dialog("➕ New Work Instruction", width="large")
def add_wi_dialog():
    qa_name = st.text_input("WI Name *", placeholder="e.g. Open Vessel")
    col1, col2 = st.columns(2)
    with col1:
        qa_phase = st.selectbox("Phase / Category *", VALID_PHASES)
        qa_confirmed = st.checkbox("✅ Confirmed by manager")
    with col2:
        qa_link = st.text_input("SOP Link (optional)")
        qa_file = st.file_uploader("SOP Excel file", type=["xlsx", "xls"])
    c1, c2 = st.columns(2)
    if c1.button("Create WI", type="primary", use_container_width=True):
        new_id, result = create_new_wi(qa_name, qa_phase, qa_link, qa_file, qa_confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Created: **{new_id}** — {result}")
            st.rerun()
    if c2.button("Cancel", use_container_width=True):
        st.rerun()

init_state()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-brand">
        {LOGO_IMG_TAG}
        <div class="brand-text">
            <div class="brand-name">WMCT Dashboard</div>
            <div class="brand-sub">CATOS Training Phase</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    page = st.radio("Navigation", [
        "📊 Dashboard",
        "📋 WI Register",
        "🔄 Progress Data",
        "💡 Insights",
        "➕ Add New WI",
        "👥 Binômes",
        "📤 Export & Save",
    ], label_visibility="collapsed")

    st.divider()

    with st.expander("📂 Load a save file"):
        restore_file = st.file_uploader("Choose Excel backup", type=["xlsx", "xls"], key="sidebar_import", label_visibility="collapsed")
        if restore_file is not None:
            if st.button("🔄 Restore", type="primary", use_container_width=True):
                ok, msg = import_from_excel(restore_file)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    st.divider()

    if st.button("➕ Quick Add WI", use_container_width=True):
        add_wi_dialog()

    st.divider()
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# ══════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════
if page == "📊 Dashboard":
    st.markdown(f"""
    <div class="dash-header">
      <div>{LOGO_IMG_TAG}</div>
      <div>
        <h1>WEST MED CONTAINER TERMINAL</h1>
        <p>CATOS Training Phase &nbsp;·&nbsp; WI Progress Dashboard &nbsp;·&nbsp; Pre-Go-Live Training</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    kpis = compute_kpis()
    binome_stats = compute_binome_stats()

    # ── KPI strip ──
    cols = st.columns(5)
    kpi_data = [
        ("📋", str(kpis["total"]),       "Total WIs",      "#0ea5e9"),
        ("✅", str(kpis["completed"]),    "Completed",      "#10b981"),
        ("🔄", str(kpis["in_progress"]), "In Progress",    "#f59e0b"),
        ("🔐", str(kpis["confirmed"]),   "Confirmed",      "#6366f1"),
        ("⏳", str(kpis["not_confirmed"]),"Not Confirmed",  "#ef4444"),
    ]
    for col, (icon, val, lbl, color) in zip(cols, kpi_data):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--accent:{color}">
                <div class="icon">{icon}</div>
                <div class="val" style="color:{color}">{val}</div>
                <div class="lbl">{lbl}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("")

    # ── Row 1: Gauge + Donut + Radar ──
    col1, col2, col3 = st.columns([1.2, 1.2, 1.6])
    with col1:
        st.plotly_chart(chart_completion_gauge(kpis["pct"]), use_container_width=True, config={"displayModeBar": False})
    with col2:
        st.plotly_chart(chart_overall_donut(kpis), use_container_width=True, config={"displayModeBar": False})
    with col3:
        st.plotly_chart(chart_binome_radar(binome_stats), use_container_width=True, config={"displayModeBar": False})

    # ── Row 2: Stacked bar + Heatmap ──
    st.markdown('<div class="section-title">TEAM PERFORMANCE</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns([1, 1.2])
    with col_a:
        st.plotly_chart(chart_binome_stacked(binome_stats), use_container_width=True, config={"displayModeBar": False})
    with col_b:
        st.plotly_chart(chart_heatmap(), use_container_width=True, config={"displayModeBar": False})

    # ── Binôme mini-cards ──
    st.markdown('<div class="section-title">BINÔME OVERVIEW</div>', unsafe_allow_html=True)
    cols = st.columns(len(binome_stats))
    status_colors = {"On Track": "#10b981", "At Risk": "#f59e0b", "Need Support": "#ef4444"}
    for col, b in zip(cols, binome_stats):
        color = status_colors.get(b["Status"], "#94a3b8")
        pct_int = int(b["Pct"] * 100)
        with col:
            st.markdown(f"""
            <div class="binome-card">
                <div class="binome-name">{b['Binôme']}</div>
                <div class="binome-members">{b['Members']}</div>
                <div class="binome-pct">{pct_int}%</div>
                <div class="prog-bar-bg"><div class="prog-bar-fill" style="width:{pct_int}%"></div></div>
                <div style="margin-top:8px; font-size:0.72rem; color:{color}; font-weight:600">{b['Status']}</div>
                <div style="font-size:0.72rem; color:#64748b; margin-top:2px">Avg score: {b['Avg Score']:.1f}/5</div>
            </div>
            """, unsafe_allow_html=True)

    # ── Score distribution ──
    score_chart = chart_score_distribution()
    if score_chart:
        st.markdown('<div class="section-title">SCORE ANALYSIS</div>', unsafe_allow_html=True)
        st.plotly_chart(score_chart, use_container_width=True, config={"displayModeBar": False})


# ══════════════════════════════════════════════
# PAGE: WI REGISTER
# ══════════════════════════════════════════════
elif page == "📋 WI Register":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">📋 WI Register</h2>', unsafe_allow_html=True)
    st.caption("All Work Instructions — edit titles, phases, manager confirmation and SOP links.")
    st.divider()

    for i, w in enumerate(st.session_state.wis):
        status_label = "✅ Confirmed" if w["confirmed"] else "⏳ Pending"
        phase_color = PHASE_COLORS.get(w["phase"], "#64748b")
        with st.expander(f"**{w['id']}** — {w['title']}", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                new_title = st.text_input("WI Title", w["title"], key=f"title_{i}")
                new_phase = st.selectbox("Phase / Category", VALID_PHASES,
                                         index=VALID_PHASES.index(w["phase"]) if w["phase"] in VALID_PHASES else 0,
                                         key=f"phase_{i}")
            with col2:
                new_confirmed = st.checkbox("✅ Manager Confirmed", w["confirmed"], key=f"confirmed_{i}")
                new_sop = st.text_input("SOP Link (optional)", w["sop"], key=f"sop_{i}")
            new_sop_file = st.file_uploader("📎 Excel SOP file", type=["xlsx", "xls"], key=f"sopfile_{i}")
            c1, c2 = st.columns([1, 4])
            if c1.button("💾 Save", key=f"save_wi_{i}", type="primary"):
                st.session_state.wis[i].update({
                    "title": new_title, "phase": new_phase,
                    "confirmed": new_confirmed, "sop": new_sop,
                })
                if new_sop_file:
                    st.session_state.wis[i]["sop_file_name"] = new_sop_file.name
                    st.session_state.wis[i]["sop_file_bytes"] = new_sop_file.getvalue()
                st.success("Saved!")
                st.rerun()
            if c2.button("🗑️ Delete", key=f"del_wi_{i}"):
                st.session_state.wis.pop(i)
                st.rerun()


# ══════════════════════════════════════════════
# PAGE: PROGRESS DATA
# ══════════════════════════════════════════════
elif page == "🔄 Progress Data":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">🔄 Progress Data</h2>', unsafe_allow_html=True)
    st.caption("Update each binôme's status, score (1-5), and notes for every WI.")
    st.divider()

    binome_names = [b["name"] for b in st.session_state.binomes]
    col_filter, col_sort = st.columns([2, 2])
    with col_filter:
        selected_binome = st.selectbox("Filter by Binôme", ["All"] + binome_names)
    with col_sort:
        selected_phase = st.selectbox("Filter by Phase", ["All"] + VALID_PHASES)

    st.markdown("")

    for w in st.session_state.wis:
        if selected_phase != "All" and w["phase"] != selected_phase:
            continue
        phase_color = PHASE_COLORS.get(w["phase"], "#64748b")
        st.markdown(f"""
        <div style="background:#1e293b;border:1px solid rgba(148,163,184,0.1);border-radius:12px;
                    padding:12px 18px;margin-bottom:14px;border-left:3px solid {phase_color}">
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
                <span style="font-family:'JetBrains Mono',monospace;font-size:0.75rem;color:#64748b">{w['id']}</span>
                <span style="font-weight:600;color:#f1f5f9;font-size:0.92rem">{w['title']}</span>
                <span style="font-size:0.72rem;color:{phase_color};background:rgba(255,255,255,0.05);
                             padding:2px 8px;border-radius:20px">{w['phase']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        binomes_to_show = st.session_state.binomes if selected_binome == "All" \
            else [b for b in st.session_state.binomes if b["name"] == selected_binome]
        cols = st.columns(len(binomes_to_show))
        for j, b in enumerate(binomes_to_show):
            bn = b["name"]
            p = get_progress(w["id"], bn)
            status_color = {"Completed": "#10b981", "In Progress": "#f59e0b", "Not Started": "#334155"}.get(p["status"], "#64748b")
            with cols[j]:
                st.markdown(f"""
                <div style="background:#0f172a;border:1px solid {status_color}33;border-radius:10px;
                            padding:10px 12px;margin-bottom:8px">
                    <div style="font-weight:600;color:#f1f5f9;font-size:0.82rem">{bn}</div>
                    <div style="color:#475569;font-size:0.72rem;margin-bottom:8px">{b['members']}</div>
                </div>
                """, unsafe_allow_html=True)
                new_status = st.selectbox("Status", STATUS_OPTIONS,
                                          index=STATUS_OPTIONS.index(p["status"]) if p["status"] in STATUS_OPTIONS else 0,
                                          key=f"st_{w['id']}_{bn}")
                new_score = st.selectbox("Score", SCORE_OPTIONS,
                                          index=SCORE_OPTIONS.index(p["score"]) if p["score"] in SCORE_OPTIONS else 0,
                                          key=f"sc_{w['id']}_{bn}")
                new_notes = st.text_input("Notes", p["notes"], key=f"no_{w['id']}_{bn}")
                if st.button("Save", key=f"sv_{w['id']}_{bn}"):
                    set_progress(w["id"], bn, {"status": new_status, "score": new_score, "notes": new_notes})
                    st.success("✅")
                    st.rerun()


# ══════════════════════════════════════════════
# PAGE: INSIGHTS
# ══════════════════════════════════════════════
elif page == "💡 Insights":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">💡 Insights & Analytics</h2>', unsafe_allow_html=True)
    st.divider()

    cat_stats = compute_category_stats()
    binome_stats = compute_binome_stats()

    tab1, tab2, tab3 = st.tabs(["📊 Categories", "👥 Binômes", "⚠️ Attention Needed"])

    with tab1:
        if cat_stats:
            col1, col2 = st.columns([1.8, 1])
            with col1:
                st.plotly_chart(chart_category_horizontal(cat_stats), use_container_width=True,
                                config={"displayModeBar": False})
            with col2:
                # Pie of WI distribution by category
                df_cat = pd.DataFrame(cat_stats)
                fig_pie = px.pie(df_cat, names="Category", values="Total WIs",
                                  color="Category",
                                  color_discrete_map={r["Category"]: PHASE_COLORS.get(r["Category"], "#64748b") for r in cat_stats},
                                  hole=0.5)
                fig_pie.update_traces(textfont=dict(color="#f1f5f9", size=10), hovertemplate="<b>%{label}</b><br>%{value} WIs<extra></extra>")
                fig_pie.update_layout(**plotly_layout, title=dict(text="Distribution by Category", font=dict(color="#94a3b8", size=13)), showlegend=False)
                st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

            st.markdown('<div class="section-title">CATEGORY DETAILS</div>', unsafe_allow_html=True)
            df_display = pd.DataFrame(cat_stats)
            df_display["Completion %"] = df_display["Completion %"].apply(lambda x: f"{x:.0%}")
            st.dataframe(df_display, use_container_width=True, hide_index=True)

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(chart_binome_radar(binome_stats), use_container_width=True, config={"displayModeBar": False})
        with col2:
            score_fig = chart_score_distribution()
            if score_fig:
                st.plotly_chart(score_fig, use_container_width=True, config={"displayModeBar": False})
        # Detailed table
        df_b = pd.DataFrame(binome_stats)
        df_b["Completion %"] = df_b["Pct"].apply(lambda x: f"{x:.0%}")
        df_b["Avg Score"] = df_b["Avg Score"].apply(lambda x: f"{x:.1f}" if x else "—")
        cols_show = ["Binôme", "Members", "Completed", "In Progress", "Not Started", "Completion %", "Avg Score", "Status"]
        st.dataframe(df_b[cols_show], use_container_width=True, hide_index=True)

    with tab3:
        inc = incomplete_wis()
        if inc:
            df_inc = pd.DataFrame(inc)
            # Status breakdown donut
            status_counts = df_inc["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            col1, col2 = st.columns([1, 2])
            with col1:
                fig_inc = px.pie(status_counts, names="Status", values="Count",
                                  color="Status",
                                  color_discrete_map={"In Progress": COLORS["in_progress"], "Not Started": "#334155"},
                                  hole=0.55)
                fig_inc.update_traces(textfont=dict(color="#f1f5f9"), hovertemplate="<b>%{label}</b><br>%{value}<extra></extra>")
                fig_inc.update_layout(**plotly_layout, title=dict(text="Incomplete WIs", font=dict(color="#94a3b8", size=13)), showlegend=True, margin=dict(l=0, r=0, t=40, b=0))
                st.plotly_chart(fig_inc, use_container_width=True, config={"displayModeBar": False})
            with col2:
                st.markdown(f"""
                <div style="background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.2);
                            border-radius:12px;padding:16px 20px;margin-bottom:12px">
                    <div style="color:#f87171;font-weight:700;font-size:1.3rem">{len(inc)}</div>
                    <div style="color:#94a3b8;font-size:0.78rem;margin-top:2px">WIs still incomplete</div>
                </div>
                """, unsafe_allow_html=True)
                for phase in sorted(df_inc["Phase"].unique()):
                    phase_wis = df_inc[df_inc["Phase"] == phase]
                    pcolor = PHASE_COLORS.get(phase, "#64748b")
                    st.markdown(f'<div style="color:{pcolor};font-weight:600;font-size:0.8rem;margin:10px 0 4px">{phase}</div>', unsafe_allow_html=True)
                    for _, row in phase_wis.iterrows():
                        badge = "badge-progress" if row["Status"] == "In Progress" else "badge-pending"
                        st.markdown(f"""
                        <div style="display:flex;align-items:center;justify-content:space-between;
                                    padding:6px 10px;background:#1e293b;border-radius:8px;margin-bottom:4px">
                            <span style="font-family:'JetBrains Mono',monospace;color:#64748b;font-size:0.72rem">{row['WI ID']}</span>
                            <span style="color:#e2e8f0;font-size:0.82rem">{row['Title']}</span>
                            <span class="badge {badge}">{row['Status']}</span>
                        </div>
                        """, unsafe_allow_html=True)
        else:
            st.success("✅ All WIs are complete — great work!")


# ══════════════════════════════════════════════
# PAGE: ADD NEW WI
# ══════════════════════════════════════════════
elif page == "➕ Add New WI":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">➕ Add Work Instruction</h2>', unsafe_allow_html=True)
    st.divider()
    with st.form("add_wi_form"):
        col1, col2 = st.columns(2)
        with col1:
            wi_name = st.text_input("WI Name *", placeholder="e.g. Open Vessel")
            phase = st.selectbox("Phase / Category *", VALID_PHASES)
        with col2:
            sop = st.text_input("SOP Link (optional)")
            confirmed = st.checkbox("Manager Confirmed")
        sop_file = st.file_uploader("📎 Excel SOP file", type=["xlsx", "xls"])
        submitted = st.form_submit_button("➕ Create WI", type="primary")
    if submitted:
        new_id, result = create_new_wi(wi_name, phase, sop, sop_file, confirmed)
        if new_id is None:
            st.error(result)
        else:
            st.success(f"✅ Created: **{new_id}** — {result} ({phase})")


# ══════════════════════════════════════════════
# PAGE: BINÔMES
# ══════════════════════════════════════════════
elif page == "👥 Binômes":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">👥 Manage Binômes</h2>', unsafe_allow_html=True)
    st.divider()
    st.markdown("### Current Binômes")
    for i, b in enumerate(st.session_state.binomes):
        with st.expander(f"**{b['name']}** — {b['members']}", expanded=False):
            c1, c2 = st.columns(2)
            c1.text_input("Name", b["name"], key=f"bn_{i}", disabled=True)
            new_members = c2.text_input("Members", b["members"], key=f"bm_{i}")
            if st.button("💾 Save", key=f"sb_{i}"):
                st.session_state.binomes[i]["members"] = new_members
                st.success("Saved!")
                st.rerun()
    st.divider()
    st.markdown("### Add New Binôme")
    with st.form("add_binome_form"):
        col1, col2 = st.columns(2)
        with col1:
            nb_name = st.text_input("Binôme Name *", placeholder="e.g. Binôme 6")
        with col2:
            nb_members = st.text_input("Members *", placeholder="e.g. Ali & Omar")
        sub_b = st.form_submit_button("➕ Add Binôme", type="primary")
    if sub_b:
        if not nb_name.strip() or not nb_members.strip():
            st.error("Both fields are required.")
        else:
            st.session_state.binomes.append({"name": nb_name.strip(), "members": nb_members.strip()})
            for w in st.session_state.wis:
                set_progress(w["id"], nb_name.strip(), {"status": "Not Started", "score": "", "notes": ""})
            st.success(f"✅ Added: {nb_name}")
            st.rerun()


# ══════════════════════════════════════════════
# PAGE: EXPORT & SAVE
# ══════════════════════════════════════════════
elif page == "📤 Export & Save":
    st.markdown('<h2 style="color:#f1f5f9;font-weight:700;margin-bottom:4px">📤 Export & Save</h2>', unsafe_allow_html=True)
    st.divider()

    st.markdown("""
    <div style="background:linear-gradient(135deg,#0c1a2e,#162d4a);border:1px dashed rgba(14,165,233,0.4);
                border-radius:14px;padding:18px 22px;margin-bottom:20px">
        <div style="color:#0ea5e9;font-weight:700;font-size:0.92rem;margin-bottom:8px">💡 How to save & restore your data</div>
        <ol style="color:#94a3b8;margin:0;padding-left:18px;font-size:0.84rem;line-height:1.9">
            <li><strong style="color:#f1f5f9">Export</strong> the Excel file below after each session.</li>
            <li>Store it on <strong style="color:#0ea5e9">Google Drive</strong> or USB.</li>
            <li>Next time, use <strong style="color:#f1f5f9">📂 Load a save file</strong> in the sidebar to restore everything in one click.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

    excel_bytes = export_to_excel()
    fname = f"WMCT_Save_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="⬇️ Download Excel Backup",
        data=excel_bytes, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Restore from file")
        restore_file = st.file_uploader("Choose backup file", type=["xlsx", "xls"], key="export_page_import")
        if restore_file:
            if st.button("🔄 Restore", type="primary"):
                ok, msg = import_from_excel(restore_file)
                st.success(msg) if ok else st.error(msg)
                if ok: st.rerun()
    with col2:
        inc = incomplete_wis()
        st.markdown("### Export incomplete WIs (CSV)")
        if inc:
            csv = pd.DataFrame(inc).to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Incomplete WIs CSV", csv, "incomplete_wis.csv", "text/csv", use_container_width=True)
        else:
            st.success("✅ No incomplete WIs!")

    st.divider()
    st.markdown("### Current data preview")
    tab1, tab2, tab3 = st.tabs(["WI Register", "Progress Data", "Category Stats"])
    with tab1:
        st.dataframe(pd.DataFrame(st.session_state.wis), use_container_width=True, hide_index=True)
    with tab2:
        rows = [{"WI": w["id"], "Binôme": b["name"], **get_progress(w["id"], b["name"])}
                for w in st.session_state.wis for b in st.session_state.binomes]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    with tab3:
        df_cs = pd.DataFrame(compute_category_stats())
        if not df_cs.empty:
            df_cs["Completion %"] = df_cs["Completion %"].apply(lambda x: f"{x:.0%}")
        st.dataframe(df_cs, use_container_width=True, hide_index=True)


# ── FLOATING BUTTON ────────────────────────────
with st.container(key="fab_container"):
    if st.button("➕ Add New WI", key="fab_add_wi"):
        add_wi_dialog()
