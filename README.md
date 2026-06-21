# ⭐ WMCT WI Dashboard

**West Med Container Terminal — CATOS Training Phase**  
Work Instructions Progress Tracker

---

## 🚀 Deploy on Streamlit Cloud (free)

1. **Push this folder to GitHub** (public or private repo)
2. Go to **[share.streamlit.io](https://share.streamlit.io)**
3. Click **"New app"** → connect your GitHub repo
4. Set **Main file path** to `app.py`
5. Click **Deploy** — done ✅

---

## 💻 Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 📁 Project structure

```
wmct_project/
├── app.py                  ← Main application
├── requirements.txt        ← Python dependencies
├── README.md               ← This file
└── .streamlit/
    └── config.toml         ← Theme & server settings
```

---

## 📋 Features

| Feature | Description |
|---|---|
| 📊 Dashboard | KPIs, binôme progress bars, WI status table |
| 📋 WI Register | View / edit / delete Work Instructions |
| 🔄 Progress Data | Update status, score (1–5), notes per binôme |
| 💡 Insights | Category stats, charts, incomplete WIs report |
| ➕ Add New WI | Create WIs with phase & SOP link |
| 👥 Manage Binômes | Add/edit binômes and member names |
| 📤 Export | Download Excel (.xlsx) or CSV reports |

---

## 🔑 Statuses

- **Not Started** — binôme has not begun this WI
- **In Progress** — binôme is working on it
- **Completed** — binôme finished (counts toward WI completion)

A WI is considered globally **Completed** once at least **1 binôme** marks it Completed.

---

*Generated from WMCT_WI_Dashboard_UPDATED.xlsm*
