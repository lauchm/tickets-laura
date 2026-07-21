import streamlit as st
import pandas as pd
from datetime import date
import gspread
from google.oauth2.service_account import Credentials
import json

st.set_page_config(
    page_title="Mis Tickets · COHIMAR",
    page_icon="🗂️",
    layout="wide"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

AREAS   = ["ERP", "IA", "CRM", "Infraestructura", "Seguridad", "SEO / Web", "Extra"]
ESTADOS = ["Pendiente", "En curso", "Resuelto"]

AREA_COLORS = {
    "ERP":            "#3730A3",
    "IA":             "#166534",
    "CRM":            "#9A3412",
    "Infraestructura":"#5B21B6",
    "Seguridad":      "#9F1239",
    "SEO / Web":      "#134E4A",
    "Externo":          "#4A5568",
    "Otros":          "#4A5568",
}

ESTADO_EMOJI = {"Pendiente": "⏳", "En curso": "🔄", "Resuelto": "✅"}
NAVY = "#1B2F5E"
RED  = "#C8001E"

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  .cohimar-header {{ background:{NAVY}; border-radius:10px 10px 0 0; padding:1.1rem 1.5rem 0; margin-bottom:0; }}
  .cohimar-header-top {{ display:flex; align-items:center; justify-content:space-between; margin-bottom:0.9rem; }}
  .cohimar-title {{ color:rgba(255,255,255,0.65); font-size:12px; letter-spacing:0.06em; text-transform:uppercase; font-weight:500; }}
  .hdr-lines {{ display:flex; height:3px; }}
  .l-red {{ background:{RED}; flex:2; height:3px; }}
  .l-wh  {{ background:rgba(255,255,255,0.85); flex:0.5; height:3px; }}
  .l-nv  {{ background:#2A4580; flex:4; height:3px; }}
  .stats-bar {{ background:#2A4580; border-left:4px solid {RED}; display:flex; border-radius:0 0 10px 10px; margin-bottom:1.5rem; }}
  .stat-box {{ flex:1; padding:0.75rem 1rem; text-align:center; border-right:1px solid rgba(255,255,255,0.08); }}
  .stat-box:last-child {{ border-right:none; }}
  .stat-n {{ font-size:24px; font-weight:700; color:#fff; line-height:1; }}
  .stat-n.red {{ color:#FF6B6B; }}
  .stat-l {{ font-size:10px; color:rgba(255,255,255,0.5); text-transform:uppercase; letter-spacing:0.05em; margin-top:3px; }}
  .ticket-card {{ background:#fff; border:1.5px solid #D8DCE6; border-radius:8px; padding:0.85rem 1rem; margin-bottom:10px; }}
  .ticket-card:hover {{ border-color:{NAVY}; }}
  .ticket-title {{ font-size:14px; font-weight:600; color:#1A202C; margin-bottom:6px; }}
  .ticket-meta {{ display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin-bottom:5px; }}
  .area-badge {{ padding:2px 8px; border-radius:20px; font-size:11px; font-weight:700; color:#fff; display:inline-block; }}
  .ticket-date {{ font-size:11px; color:#8892A4; }}
  .ticket-desc {{ font-size:12px; color:#4A5568; line-height:1.5; margin-top:4px; }}
  .ruta-box {{ background:#F7F8FA; border:1px solid #D8DCE6; border-radius:5px; padding:4px 8px; font-size:11px; font-family:monospace; color:{NAVY}; margin-top:6px; word-break:break-all; }}
  .col-header-pend {{ border-top:3px solid #FBBF24; border-radius:8px 8px 0 0; padding:8px 12px; background:#fff; border:1px solid #EEF0F4; font-size:12px; font-weight:700; color:#92400E; text-transform:uppercase; letter-spacing:0.04em; }}
  .col-header-curso {{ border-top:3px solid #3B82F6; border-radius:8px 8px 0 0; padding:8px 12px; background:#fff; border:1px solid #EEF0F4; font-size:12px; font-weight:700; color:#1E3A8A; text-transform:uppercase; letter-spacing:0.04em; }}
  .col-header-done  {{ border-top:3px solid #10B981; border-radius:8px 8px 0 0; padding:8px 12px; background:#fff; border:1px solid #EEF0F4; font-size:12px; font-weight:700; color:#065F46; text-transform:uppercase; letter-spacing:0.04em; }}
  .stButton > button {{ background:{NAVY} !important; color:#fff !important; border:none !important; border-radius:6px !important; font-weight:600 !important; font-size:13px !important; }}
  .stButton > button:hover {{ background:#2A4580 !important; }}
  #MainMenu, footer, header {{ visibility:hidden; }}
</style>
""", unsafe_allow_html=True)


@st.cache_resource
def get_sheet():
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open(st.secrets["SHEET_NAME"])
    try:
        ws = sh.worksheet("Tickets")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Tickets", rows=1000, cols=7)
        ws.append_row(["ID", "Fecha", "Título", "Área", "Descripción", "Ruta", "Estado"])
    return ws


def load_tickets():
    ws = get_sheet()
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=["ID","Fecha","Título","Área","Descripción","Ruta","Estado"])
    return pd.DataFrame(data)


def add_ticket(titulo, area, descripcion, ruta, estado, fecha):
    ws = get_sheet()
    ticket_id = int(pd.Timestamp.now().timestamp() * 1000)
    ws.append_row([ticket_id, str(fecha), titulo, area, descripcion, ruta, estado])
    st.cache_resource.clear()


def update_estado(ticket_id, nuevo_estado):
    ws = get_sheet()
    cell = ws.find(str(ticket_id))
    if cell:
        ws.update_cell(cell.row, 7, nuevo_estado)
    st.cache_resource.clear()


def delete_ticket(ticket_id):
    ws = get_sheet()
    cell = ws.find(str(ticket_id))
    if cell:
        ws.delete_rows(cell.row)
    st.cache_resource.clear()


# HEADER
st.markdown(f"""
<div class="cohimar-header">
  <div class="cohimar-header-top">
    <span style="color:white;font-size:20px;font-weight:800;letter-spacing:0.05em;">COHIMAR</span>
    <span class="cohimar-title">Gestión de Tickets · 2026</span>
  </div>
  <div class="hdr-lines">
    <div class="l-red"></div><div class="l-wh"></div><div class="l-nv"></div>
  </div>
</div>
""", unsafe_allow_html=True)

try:
    df = load_tickets()
except Exception as e:
    st.error(f"Error conectando con Google Sheets: {e}")
    st.stop()

# STATS
n_pend  = len(df[df["Estado"] == "Pendiente"]) if not df.empty else 0
n_curso = len(df[df["Estado"] == "En curso"])  if not df.empty else 0
n_done  = len(df[df["Estado"] == "Resuelto"])  if not df.empty else 0
n_total = len(df)

st.markdown(f"""
<div class="stats-bar">
  <div class="stat-box"><div class="stat-n">{n_pend}</div><div class="stat-l">⏳ Pendiente</div></div>
  <div class="stat-box"><div class="stat-n">{n_curso}</div><div class="stat-l">🔄 En curso</div></div>
  <div class="stat-box"><div class="stat-n">{n_done}</div><div class="stat-l">✅ Resueltos</div></div>
  <div class="stat-box"><div class="stat-n red">{n_total}</div><div class="stat-l">Total tickets</div></div>
</div>
""", unsafe_allow_html=True)

# FORMULARIO
with st.expander("➕  Nuevo ticket", expanded=False):
    with st.form("nuevo_ticket", clear_on_submit=True):
        col1, col2 = st.columns([3, 1])
        with col1:
            titulo = st.text_input("Título *", placeholder="Ej: Configurar nuevo PC en administración")
        with col2:
            fecha = st.date_input("Fecha *", value=date.today())
        col3, col4 = st.columns([1, 1])
        with col3:
            area = st.selectbox("Área *", AREAS)
        with col4:
            estado = st.selectbox("Estado inicial", ESTADOS)
        descripcion = st.text_area("Descripción", placeholder="Qué hiciste, cómo lo resolviste, resultado...", height=90)
        ruta = st.text_input("📁 Ruta del documento (opcional)", placeholder=r"\\servidor\Documentacion\archivo.docx")
        submitted = st.form_submit_button("✓ Crear ticket", use_container_width=True)
        if submitted:
            if not titulo.strip():
                st.error("El título es obligatorio.")
            else:
                add_ticket(titulo.strip(), area, descripcion.strip(), ruta.strip(), estado, fecha)
                st.success("✅ Ticket creado.")
                st.rerun()

# FILTROS
if not df.empty:
    st.markdown("---")
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        f_area = st.multiselect("Filtrar por área", AREAS)
    with fc2:
        f_estado = st.multiselect("Filtrar por estado", ESTADOS)
    with fc3:
        f_buscar = st.text_input("🔍 Buscar")
    df_fil = df.copy()
    if f_area:   df_fil = df_fil[df_fil["Área"].isin(f_area)]
    if f_estado: df_fil = df_fil[df_fil["Estado"].isin(f_estado)]
    if f_buscar:
        mask = (df_fil["Título"].str.contains(f_buscar, case=False, na=False) |
                df_fil["Descripción"].str.contains(f_buscar, case=False, na=False))
        df_fil = df_fil[mask]
else:
    df_fil = df.copy()

# KANBAN
st.markdown("---")
col_pend, col_curso, col_done = st.columns(3)

def render_column(col, estado_label, header_class, df_data):
    with col:
        tickets_col = df_data[df_data["Estado"] == estado_label] if not df_data.empty else pd.DataFrame()
        cnt = len(tickets_col)
        st.markdown(
            f'<div class="{header_class}">{ESTADO_EMOJI[estado_label]} {estado_label} &nbsp;'
            f'<span style="background:#EEF0F4;color:#4A5568;padding:1px 8px;border-radius:20px;font-size:11px;">{cnt}</span></div>',
            unsafe_allow_html=True
        )
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        if tickets_col.empty:
            st.markdown('<div style="text-align:center;color:#8892A4;font-size:13px;padding:1.5rem 0;">Sin tickets</div>', unsafe_allow_html=True)
        else:
            for _, row in tickets_col.iterrows():
                tid = row["ID"]
                color = AREA_COLORS.get(row["Área"], "#4A5568")
                ruta_html = f'<div class="ruta-box">📁 {row["Ruta"]}</div>' if row.get("Ruta") else ""
                desc_html = f'<div class="ticket-desc">{row["Descripción"]}</div>' if row.get("Descripción") else ""
                st.markdown(f"""
                <div class="ticket-card">
                  <div class="ticket-title">{row["Título"]}</div>
                  <div class="ticket-meta">
                    <span class="area-badge" style="background:{color};">{row["Área"]}</span>
                    <span class="ticket-date">{row.get("Fecha","")}</span>
                  </div>
                  {desc_html}{ruta_html}
                </div>""", unsafe_allow_html=True)

                btns = st.columns(3)
                if estado_label == "Pendiente":
                    with btns[0]:
                        if st.button("▶ Iniciar", key=f"ini_{tid}"):
                            update_estado(tid, "En curso"); st.rerun()
                elif estado_label == "En curso":
                    with btns[0]:
                        if st.button("✓ Resolver", key=f"res_{tid}"):
                            update_estado(tid, "Resuelto"); st.rerun()
                    with btns[1]:
                        if st.button("← Pend.", key=f"pen_{tid}"):
                            update_estado(tid, "Pendiente"); st.rerun()
                elif estado_label == "Resuelto":
                    with btns[0]:
                        if st.button("↩ Reabrir", key=f"rea_{tid}"):
                            update_estado(tid, "En curso"); st.rerun()
                with btns[2]:
                    if st.button("🗑", key=f"del_{tid}"):
                        delete_ticket(tid); st.rerun()

render_column(col_pend,  "Pendiente", "col-header-pend",  df_fil)
render_column(col_curso, "En curso",  "col-header-curso",  df_fil)
render_column(col_done,  "Resuelto",  "col-header-done",   df_fil)

# EXPORTAR
if not df.empty:
    st.markdown("---")
    ec1, ec2 = st.columns([3,1])
    with ec1:
        st.caption(f"📋 {n_total} tickets registrados · {n_done} resueltos")
    with ec2:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("↓ Exportar CSV", data=csv,
                           file_name="tickets_laura_cohimar_2026.csv",
                           mime="text/csv", use_container_width=True)
