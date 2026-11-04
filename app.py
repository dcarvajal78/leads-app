import streamlit as st
import anthropic
import base64
import json
import io
from datetime import date, datetime
from supabase import create_client, Client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Leads Clientes Nuevos",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

ESTADOS = ["Pendiente contacto", "Correo enviado", "Antecedentes recibidos", "Pendiente de creación", "Cliente creado"]
ESTADO_COLOR = {
    "Pendiente contacto": "🟠",
    "Correo enviado": "🔵",
    "Antecedentes recibidos": "🟣",
    "Pendiente de creación": "🟡",
    "Cliente creado": "🟢",
}

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
  .main .block-container { padding-top: 1.5rem; max-width: 1200px; }
  h1 { color: #1F3864; font-size: 1.5rem !important; }
  .metric-card { background: white; border-radius: 10px; padding: 16px 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid; }
  .stTabs [data-baseweb="tab-list"] { gap: 4px; }
  .stTabs [data-baseweb="tab"] { border-radius: 6px 6px 0 0; font-weight: 600; }
  div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
  .badge { display:inline-block; padding:2px 10px; border-radius:99px; font-size:12px; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── Clientes ──────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def get_anthropic():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── DB helpers ────────────────────────────────────────────────
def load_leads():
    sb = get_supabase()
    res = sb.table("leads").select("*").order("id", desc=True).execute()
    return res.data or []

def create_lead(data: dict):
    sb = get_supabase()
    data["fecha"] = data.get("fecha") or date.today().isoformat()
    data["estado"] = data.get("estado", "Pendiente contacto")
    return sb.table("leads").insert(data).execute()

def update_lead(lead_id: int, data: dict):
    sb = get_supabase()
    return sb.table("leads").update(data).eq("id", lead_id).execute()

def delete_lead(lead_id: int):
    sb = get_supabase()
    return sb.table("leads").delete().eq("id", lead_id).execute()

def avanzar_estado(lead: dict):
    idx = ESTADOS.index(lead["estado"]) if lead["estado"] in ESTADOS else -1
    if idx >= len(ESTADOS) - 1:
        return
    nuevo = ESTADOS[idx + 1]
    updates = {"estado": nuevo}
    fecha_map = {
        "Correo enviado": "fecha_correo",
        "Antecedentes recibidos": "fecha_antecedentes",
        "Cliente creado": "fecha_creado",
    }
    if nuevo in fecha_map:
        updates[fecha_map[nuevo]] = date.today().isoformat()
    update_lead(lead["id"], updates)

def retroceder_estado(lead: dict):
    idx = ESTADOS.index(lead["estado"]) if lead["estado"] in ESTADOS else 0
    if idx <= 0:
        return
    prev = ESTADOS[idx - 1]
    updates = {"estado": prev}
    # Limpiar fecha del estado que se abandona
    fecha_map = {
        "Correo enviado": "fecha_correo",
        "Antecedentes recibidos": "fecha_antecedentes",
        "Cliente creado": "fecha_creado",
    }
    if lead["estado"] in fecha_map:
        updates[fecha_map[lead["estado"]]] = None
    update_lead(lead["id"], updates)

# ── Extracción IA ─────────────────────────────────────────────
def extraer_datos_imagen(img_bytes: bytes, media_type: str) -> dict:
    b64 = base64.standard_b64encode(img_bytes).decode()
    client = get_anthropic()
    prompt = """Extrae los datos del formulario de contacto en esta imagen y devuélvelos como JSON con estas claves exactas (usa null si no aparece el campo):
{
  "nombre": "",
  "rut_cliente": "",
  "rut_empresa": "",
  "razon_social": "",
  "giro": "",
  "region": "",
  "comuna": "",
  "direccion": "",
  "telefono": "",
  "mail": "",
  "motivo": ""
}
Responde SOLO con el JSON, sin texto adicional."""
    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    texto = response.content[0].text.strip()
    if texto.startswith("```"):
        texto = "\n".join(texto.split("\n")[1:])
        if "```" in texto:
            texto = texto[:texto.index("```")]
    return json.loads(texto)

# ── Export Excel ──────────────────────────────────────────────
def generar_excel(leads: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"
    headers = ["ID","Fecha","Nombre","RUT Cliente","RUT Empresa","Razón Social","Giro",
               "Región","Comuna","Dirección","Teléfono","Mail","Motivo","Estado",
               "Correo Enviado","Antecedentes","Cliente Creado","Observaciones"]
    keys = ["id","fecha","nombre","rut_cliente","rut_empresa","razon_social","giro",
            "region","comuna","direccion","telefono","mail","motivo","estado",
            "fecha_correo","fecha_antecedentes","fecha_creado","observaciones"]
    thin = Side(style="thin", color="B8C4E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", start_color="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    for ri, lead in enumerate(reversed(leads), 2):
        for ci, key in enumerate(keys, 1):
            c = ws.cell(row=ri, column=ci, value=lead.get(key))
            c.font = Font(name="Arial", size=10)
            c.border = border
            c.alignment = Alignment(vertical="center")
    widths = [6,12,22,14,14,28,16,16,14,26,12,28,16,22,16,16,16,28]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── Correo ────────────────────────────────────────────────────
def generar_correo(lead: dict) -> dict:
    nombre = lead.get("nombre", "")
    razon = lead.get("razon_social") or nombre
    mail = lead.get("mail", "")
    cuerpo = f"""Estimado/a {nombre},

Junto con saludarle, acusamos recibo de su solicitud de registro como cliente. Para continuar con el proceso de activación de su cuenta, necesitamos que nos haga llegar los siguientes antecedentes:

1. Cédula de identidad del representante legal (ambos lados)
2. Iniciación de actividades emitida por el SII
3. Carpeta tributaria electrónica del SII (sii.cl → Mi SII → Carpeta Tributaria)
4. Datos bancarios: banco, tipo de cuenta, N° de cuenta, RUT titular y correo para confirmación de transferencia

Por favor, envíe los documentos respondiendo este correo.

Una vez recibidos y validados, procederemos a la creación de su cuenta y le notificaremos a la brevedad.

Saludos cordiales,
[TU NOMBRE] | Autoplanet"""
    return {
        "asunto": f"Solicitud de Antecedentes — {razon}",
        "para": mail,
        "cuerpo": cuerpo
    }

# ══════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════

st.markdown("## 📋 Leads Clientes Nuevos — Autoplanet")

tab_panel, tab_nuevo = st.tabs(["📊 Panel de seguimiento", "📷 Nuevo lead desde foto"])

# ── TAB 1: Panel ──────────────────────────────────────────────
with tab_panel:
    leads = load_leads()

    # Stats
    cnt = {e: sum(1 for l in leads if l.get("estado") == e) for e in ESTADOS}
    c0, c1, c2, c3, c4, c5 = st.columns(6)
    c0.metric("Total Leads", len(leads))
    c1.metric("🟠 Pendiente contacto", cnt["Pendiente contacto"])
    c2.metric("🔵 Correo enviado", cnt["Correo enviado"])
    c3.metric("🟣 Antecedentes recibidos", cnt["Antecedentes recibidos"])
    c4.metric("🟡 Pendiente de creación", cnt["Pendiente de creación"])
    c5.metric("🟢 Cliente creado", cnt["Cliente creado"])

    st.divider()

    # Filtros + export
    col_b, col_f, col_exp = st.columns([3, 2, 1])
    buscar = col_b.text_input("🔍 Buscar", placeholder="Nombre, RUT, razón social…", label_visibility="collapsed")
    filtro_e = col_f.selectbox("Estado", ["Todos"] + ESTADOS, label_visibility="collapsed")
    excel_bytes = generar_excel(leads)
    col_exp.download_button("⬇ Excel", data=excel_bytes,
                            file_name=f"leads_{date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)

    # Filtrar
    filtrados = leads
    if buscar:
        q = buscar.lower()
        filtrados = [l for l in filtrados if q in " ".join([
            str(l.get("nombre","")), str(l.get("rut_cliente","")),
            str(l.get("razon_social","")), str(l.get("mail",""))
        ]).lower()]
    if filtro_e != "Todos":
        filtrados = [l for l in filtrados if l.get("estado") == filtro_e]

    if not filtrados:
        st.info("Sin leads que coincidan con el filtro.")
    else:
        for lead in filtrados:
            with st.container(border=True):
                col1, col2, col3, col4 = st.columns([3, 2, 2, 3])
                with col1:
                    st.markdown(f"**{lead.get('nombre','')}**  \n"
                                f"<small>{lead.get('razon_social','')}</small>", unsafe_allow_html=True)
                    st.caption(f"RUT: {lead.get('rut_cliente','')} | {lead.get('comuna','')} | {lead.get('telefono','')}")
                with col2:
                    estado = lead.get("estado", "")
                    st.markdown(f"{ESTADO_COLOR.get(estado,'')} **{estado}**")
                    st.caption(f"Fecha: {lead.get('fecha','')}")
                with col3:
                    mail = lead.get("mail","")
                    if mail:
                        st.markdown(f"📧 [{mail}](mailto:{mail})")
                    st.caption(lead.get("giro",""))
                with col4:
                    ba, bb, bc, bd, be = st.columns(5)
                    if ba.button("◀", key=f"ret_{lead['id']}", use_container_width=True, help="Retroceder estado"):
                        retroceder_estado(lead)
                        st.rerun()

                    if bb.button("▶", key=f"av_{lead['id']}", use_container_width=True, help="Avanzar estado"):
                        avanzar_estado(lead)
                        st.rerun()

                    correo_data = generar_correo(lead)
                    if bc.button("📧", key=f"co_{lead['id']}", use_container_width=True, help="Generar correo"):
                        st.session_state[f"show_correo_{lead['id']}"] = True
                        if lead.get("estado") == "Pendiente contacto":
                            avanzar_estado(lead)

                    rut = lead.get("rut_cliente") or lead.get("rut_empresa") or ""
                    sii_url = f"https://www2.sii.cl/stc/noauthz"
                    bd.link_button("🔎", sii_url, use_container_width=True, help=f"Consultar SII (RUT: {rut})")

                    if be.button("🗑", key=f"del_{lead['id']}", use_container_width=True, help="Eliminar"):
                        delete_lead(lead["id"])
                        st.rerun()

                    # Modal correo
                    if st.session_state.get(f"show_correo_{lead['id']}"):
                        with st.expander("📧 Correo generado", expanded=True):
                            st.text_input("Para:", value=correo_data["para"], key=f"cp_{lead['id']}")
                            st.text_input("Asunto:", value=correo_data["asunto"], key=f"ca_{lead['id']}")
                            cuerpo_txt = st.text_area("Cuerpo:", value=correo_data["cuerpo"],
                                                       height=200, key=f"cb_{lead['id']}")
                            st.code(cuerpo_txt, language=None)
                            if st.button("Cerrar", key=f"cc_{lead['id']}"):
                                st.session_state[f"show_correo_{lead['id']}"] = False
                                st.rerun()

# ── TAB 2: Nuevo lead ─────────────────────────────────────────
with tab_nuevo:
    st.markdown("### Subir foto del formulario")

    uploaded = st.file_uploader(
        "Arrastra o selecciona la captura del formulario",
        type=["jpg", "jpeg", "png", "webp"],
        help="La IA extrae los datos automáticamente"
    )

    if uploaded:
        col_img, col_form = st.columns([1, 1])
        with col_img:
            st.image(uploaded, caption="Formulario recibido", use_container_width=True)

        with col_form:
            # Extraer datos si es una imagen nueva
            key_img = f"extracted_{uploaded.name}_{uploaded.size}"
            if key_img not in st.session_state:
                with st.spinner("Extrayendo datos con IA…"):
                    try:
                        img_bytes = uploaded.read()
                        mt = uploaded.type or "image/jpeg"
                        datos = extraer_datos_imagen(img_bytes, mt)
                        st.session_state[key_img] = datos
                        st.success("Datos extraídos correctamente")
                    except Exception as e:
                        st.error(f"Error al procesar la imagen: {e}")
                        st.session_state[key_img] = {}

            datos = st.session_state.get(key_img, {})

            st.markdown("#### Revisa y corrige los datos")
            with st.form("form_lead"):
                c1, c2 = st.columns(2)
                fecha = c1.date_input("Fecha recepción", value=date.today())
                motivo = c2.text_input("Motivo contacto", value=datos.get("motivo") or "")
                nombre = st.text_input("Nombre completo *", value=datos.get("nombre") or "")
                c3, c4 = st.columns(2)
                rut_cli = c3.text_input("RUT Cliente", value=datos.get("rut_cliente") or "")
                rut_emp = c4.text_input("RUT Empresa", value=datos.get("rut_empresa") or "")
                razon = st.text_input("Razón Social", value=datos.get("razon_social") or "")
                c5, c6 = st.columns(2)
                giro = c5.text_input("Giro", value=datos.get("giro") or "")
                region = c6.text_input("Región", value=datos.get("region") or "")
                c7, c8 = st.columns(2)
                comuna = c7.text_input("Comuna", value=datos.get("comuna") or "")
                telefono = c8.text_input("Teléfono", value=datos.get("telefono") or "")
                direccion = st.text_input("Dirección", value=datos.get("direccion") or "")
                mail = st.text_input("Mail", value=datos.get("mail") or "")
                obs = st.text_area("Observaciones", height=80)

                submitted = st.form_submit_button("💾 Guardar lead", use_container_width=True, type="primary")
                if submitted:
                    if not nombre.strip():
                        st.error("El nombre es requerido")
                    else:
                        create_lead({
                            "fecha": fecha.isoformat(),
                            "nombre": nombre.strip(),
                            "rut_cliente": rut_cli.strip(),
                            "rut_empresa": rut_emp.strip(),
                            "razon_social": razon.strip(),
                            "giro": giro.strip(),
                            "region": region.strip(),
                            "comuna": comuna.strip(),
                            "direccion": direccion.strip(),
                            "telefono": telefono.strip(),
                            "mail": mail.strip(),
                            "motivo": motivo.strip(),
                            "observaciones": obs.strip(),
                        })
                        del st.session_state[key_img]
                        st.success("✅ Lead guardado correctamente")
                        st.balloons()
    else:
        st.markdown("""
        <div style="border:2px dashed #CBD5E0;border-radius:10px;padding:60px;text-align:center;background:white">
            <div style="font-size:48px">📷</div>
            <p style="color:#718096;margin-top:12px;font-size:15px">
                Sube una captura del formulario de contacto<br>
                <strong>y la IA extraerá los datos automáticamente</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
