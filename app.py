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

# ââ Config ââââââââââââââââââââââââââââââââââââââââââââââââââââ
st.set_page_config(
    page_title="Leads Clientes Nuevos",
    page_icon="ð",
    layout="wide",
    initial_sidebar_state="collapsed"
)

ESTADOS = ["Pendiente contacto", "Correo enviado", "Antecedentes recibidos", "Pendiente de creaciÃ³n", "Cliente creado"]
ESTADO_COLOR = {
    "Pendiente contacto": "ð ",
    "Correo enviado": "ðµ",
    "Antecedentes recibidos": "ð£",
    "Pendiente de creaciÃ³n": "ð¡",
    "Cliente creado": "ð¢",
}

# ââ CSS âââââââââââââââââââââââââââââââââââââââââââââââââââââââ
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

# ââ Clientes ââââââââââââââââââââââââââââââââââââââââââââââââââ
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def get_anthropic():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ââ DB helpers ââââââââââââââââââââââââââââââââââââââââââââââââ
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
    fecha_map = {
        "Correo enviado": "fecha_correo",
        "Antecedentes recibidos": "fecha_antecedentes",
        "Cliente creado": "fecha_creado",
    }
    if lead["estado"] in fecha_map:
        updates[fecha_map[lead["estado"]]] = None
    update_lead(lead["id"], updates)

# ââ ExtracciÃ³n IA ââââââââââââââââââââââââââââââââââââââââââââââ
def extraer_datos_imagen(img_bytes: bytes, media_type: str) -> dict:
    b64 = base64.standard_b64encode(img_bytes).decode()
    client = get_anthropic()
    prompt = """Extrae los datos del formulario de contacto en esta imagen y devuÃ©lvelos como JSON con estas claves exactas (usa null si no aparece el campo):
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

# ââ Export Excel âââââââââââââââââââââââââââââââââââââââââââââ
def generar_excel(leads: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"
    headers = ["ID","Fecha","Nombre","RUT Cliente","RUT Empresa","RazÃ³n Social","Giro",
               "RegiÃ³n","Comuna","DirecciÃ³n","TelÃ©fono","Mail","Motivo","Estado",
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

# ââ Correo ââââââââââââââââââââââââââââââââââââââââââââââââââââ
def generar_correo(lead: dict) -> dict:
    nombre = lead.get("nombre", "")
    razon = lead.get("razon_social") or nombre
    mail = lead.get("mail", "")
    cuerpo = f"""Estimado/a {nombre},

Junto con saludarle, acusamos recibo de su solicitud de registro como cliente. Para continuar con el proceso de activaciÃ³n de su cuenta, necesitamos que nos haga llegar los siguientes antecedentes:

1. CÃ©dula de identidad del representante legal (ambos lados)
2. IniciaciÃ³n de actividades emitida por el SII
3. Carpeta tributaria electrÃ³nica del SII (sii.cl â Mi SII â Carpeta Tributaria)
4. Datos bancarios: banco, tipo de cuenta, NÂ° de cuenta, RUT titular y correo para confirmaciÃ³n de transferencia

Por favor, envÃ­e los documentos respondiendo este correo.

Una vez recibidos y validados, procederemos a la creaciÃ³n de su cuenta y le notificaremos a la brevedad.

Saludos cordiales,
[TU NOMBRE] | Autoplanet"""
    return {
        "asunto": f"Solicitud de Antecedentes â {razon}",
        "para": mail,
        "cuerpo": cuerpo
    }

# ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
#  UI
# ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

st.markdown("## ð Leads Clientes Nuevos â Autoplanet")

tab_panel, tab_nuevo = st.tabs(["ð Panel de seguimiento", "ð· Nuevo lead desde foto"])

# ââ TAB 1: Panel ââââââââââââââââââââââââââââââââââââââââââââââ
with tab_panel:
    leads = load_leads()

    # Stats
    cnt = {e: sum(1 for l in leads if l.get("estado") == e) for e in ESTADOS}
    c0, c1, c2, c3, c4, c5 = st.columns(6)
    c0.metric("Total Leads", len(leads))
    c1.metric("ð  Pendiente contacto", cnt["Pendiente contacto"])
    c2.metric("ðµ Correo enviado", cnt["Correo enviado"])
    c3.metric("ð£ Antecedentes recibidos", cnt["Antecedentes recibidos"])
    c4.metric("ð¡ Pendiente de creaciÃ³n", cnt["Pendiente de creaciÃ³n"])
    c5.metric("ð¢ Cliente creado", cnt["Cliente creado"])

    st.divider()

    # Filtros + export
    col_b, col_f, col_exp = st.columns([3, 2, 1])
    buscar = col_b.text_input("ð Buscar", placeholder="Nombre, RUT, razÃ³n socialâ¦", label_visibility="collapsed")
    filtro_e = col_f.selectbox("Estado", ["Todos"] + ESTADOS, label_visibility="collapsed")
    excel_bytes = generar_excel(leads)
    col_exp.download_button("{¬ Excel", data=excel_bytes,
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
                        st.markdown(f"ð§ [{mail}](mailto:{mail})")
                    st.caption(lead.get("giro",""))
                with col4:
                    ba, bb, bc, bd, be = st.columns(5)
                    if ba.button("â", key=f"ret_{lead['id']}", use_container_width=True, help="Retroceder estado"):
                        retroceder_estado(lead)
                        st.rerun()

                    if bb.button("â¶", key=f"av_{lead['id']}", use_container_width=True, help="Avanzar estado"):
                        avanzar_estado(lead)
                        st.rerun()

                    correo_data = generar_correo(lead)
                    if bc.button("{Â~N°­äõ½}í±l¥uô°ÕÍ}½¹Ñ¥¹É}Ý¥Ñ õQÉÕ°¡±Àô¹ÉÈ½ÉÉ¼¤è(ÍÐ¹ÍÍÍ¥½¹}ÍÑÑmÍ¡½Ý}½ÉÉ½}í±l¥uôtôQÉÕ(¥±¹Ð ÍÑ¼¤ôôA¹¥¹Ñ½¹ÑÑ¼è(Ù¹éÉ}ÍÑ¼¡±¤((ÉÕÐô±¹Ð ÉÕÑ}±¥¹Ñ¤½È±¹Ð ÉÕÑ}µÁÉÍ¤½È(Í¥¥}ÕÉ°ô¡ÑÑÁÌè¼½ÝÝÜÈ¹Í¥¤¹°½ÍÑ½¹½ÕÑ¡è(¹±¥¹­}ÕÑÑ½¸ Â~R8°Í¥¥}ÕÉ°°ÕÍ}½¹Ñ¥¹É}Ý¥Ñ õQÉÕ°¡±Àõ
½¹ÍÕ±ÑÈM%$¡IUPèíÉÕÑô¤¤((¥¹ÕÑÑ½¸ {Â~^D°­äõ±}í±l¥uô°ÕÍ}½¹Ñ¥¹É}Ý¥Ñ õQÉÕ°¡±Àô±¥µ¥¹È¤è(±Ñ}±¡±l¥t¤(ÍÐ¹ÉÉÕ¸ ¤((5½°½ÉÉ¼(¥ÍÐ¹ÍÍÍ¥½¹}ÍÑÑ¹Ð¡Í¡½Ý}½ÉÉ½}í±l¥uô¤è(Ý¥Ñ ÍÐ¹áÁ¹È Â~N
½ÉÉ¼¹É¼°áÁ¹õQÉÕ¤è(ÍÐ¹ÑáÑ}¥¹ÁÕÐ AÉè°Ù±Õõ½ÉÉ½}ÑlÁÉt°­äõÁ}í±l¥uô¤(ÍÐ¹ÑáÑ}¥¹ÁÕÐ ÍÕ¹Ñ¼è°Ù±Õõ½ÉÉ½}ÑlÍÕ¹Ñ¼t°­äõ}í±l¥uô¤(ÕÉÁ½}ÑáÐôÍÐ¹ÑáÑ}É 
ÕÉÁ¼è°Ù±Õõ½ÉÉ½}ÑlÕÉÁ¼t°(¡¥¡ÐôÈÀÀ°­äõ}í±l¥uô¤(ÍÐ¹½¡ÕÉÁ½}ÑáÐ°±¹Õõ9½¹¤(¥ÍÐ¹ÕÑÑ½¸ 
ÉÉÈ°­äõ}í±l¥uô¤è(ÍÐ¹ÍÍÍ¥½¹}ÍÑÑmÍ¡½Ý}½ÉÉ½}í±l¥uôtô±Í(ÍÐ¹ÉÉÕ¸ ¤((RR QÈè9ÕÙ¼±RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR )Ý¥Ñ Ñ}¹ÕÙ¼è(ÍÐ¹µÉ­½Ý¸ MÕ¥È½Ñ¼°½ÉµÕ±É¥¼¤((ÕÁ±½ôÍÐ¹¥±}ÕÁ±½È (ÉÉÍÑÉ¼Í±¥½¹±ÁÑÕÉ°½ÉµÕ±É¥¼°(ÑåÁõl©Á°©Á°Á¹°ÝÀt°(¡±Àô1%áÑÉ±½ÌÑ½ÌÕÑ½·Ñ¥µ¹Ñ(¤((¥ÕÁ±½è(½±}¥µ°½±}½É´ôÍÐ¹½±Õµ¹Ì¡lÄ°Åt¤(Ý¥Ñ ½±}¥µè(ÍÐ¹¥µ¡ÕÁ±½°ÁÑ¥½¸ô½ÉµÕ±É¥¼É¥¥¼°ÕÍ}½¹Ñ¥¹É}Ý¥Ñ õQÉÕ¤((Ý¥Ñ ½±}½É´è(áÑÉÈÑ½ÌÍ¤ÌÕ¹¥µ¸¹ÕÙ(­å}¥µôáÑÉÑ}íÕÁ±½¹¹µõ}íÕÁ±½¹Í¥éô(¥­å}¥µ¹½Ð¥¸ÍÐ¹ÍÍÍ¥½¹}ÍÑÑè(Ý¥Ñ ÍÐ¹ÍÁ¥¹¹È áÑÉå¹¼Ñ½Ì½¸%¤è(ÑÉäè(¥µ}åÑÌôÕÁ±½¹É ¤(µÐôÕÁ±½¹ÑåÁ½È¥µ½©Á(Ñ½ÌôáÑÉÉ}Ñ½Í}¥µ¸¡¥µ}åÑÌ°µÐ¤(ÍÐ¹ÍÍÍ¥½¹}ÍÑÑm­å}¥µtôÑ½Ì(ÍÐ¹ÍÕÍÌ Ñ½ÌáÑÉµ½Ì½ÉÉÑµ¹Ñ¤(áÁÐáÁÑ¥½¸Ìè(ÍÐ¹ÉÉ½È¡ÉÉ½È°ÁÉ½ÍÈ±¥µ¸èíô¤(ÍÐ¹ÍÍÍ¥½¹}ÍÑÑm­å}¥µtôíô((Ñ½ÌôÍÐ¹ÍÍÍ¥½¹}ÍÑÑ¹Ð¡­å}¥µ°íô¤((ÍÐ¹µÉ­½Ý¸ IÙ¥Íä½ÉÉ¥±½ÌÑ½Ì¤(Ý¥Ñ ÍÐ¹½É´ ½Éµ}±¤è(Ä°ÈôÍÐ¹½±Õµ¹Ì È¤(¡ôÄ¹Ñ}¥¹ÁÕÐ ¡ÉÁ§Í¸°Ù±ÕõÑ¹Ñ½ä ¤¤(µ½Ñ¥Ù¼ôÈ¹ÑáÑ}¥¹ÁÕÐ 5½Ñ¥Ù¼½¹ÑÑ¼°Ù±ÕõÑ½Ì¹Ð µ½Ñ¥Ù¼¤½È¤(¹½µÉôÍÐ¹ÑáÑ}¥¹ÁÕÐ 9½µÉ½µÁ±Ñ¼¨°Ù±ÕõÑ½Ì¹Ð ¹½µÉ¤½È¤(Ì°ÐôÍÐ¹½±Õµ¹Ì È¤(ÉÕÑ}±¤ôÌ¹ÑáÑ}¥¹ÁÕÐ IUP
±¥¹Ñ°Ù±ÕõÑ½Ì¹Ð ÉÕÑ}±¥¹Ñ¤½È¤(ÉÕÑ}µÀôÐ¹ÑáÑ}¥¹ÁÕÐ IUPµÁÉÍ°Ù±ÕõÑ½Ì¹Ð ÉÕÑ}µÁÉÍ¤½È¤(Éé½¸ôÍÐ¹ÑáÑ}¥¹ÁÕÐ IëÍ¸M½¥°°Ù±ÕõÑ½Ì¹Ð Éé½¹}Í½¥°¤½È¤(Ô°ØôÍÐ¹½±Õµ¹Ì È¤(¥É¼ôÔ¹ÑáÑ}¥¹ÁÕÐ ¥É¼°Ù±ÕõÑ½Ì¹Ð ¥É¼¤½È¤(É¥½¸ôØ¹ÑáÑ}¥¹ÁÕÐ I§Í¸°Ù±ÕõÑ½Ì¹Ð É¥½¸¤½È¤(Ü°àôÍÐ¹½±Õµ¹Ì È¤(½µÕ¹ôÜ¹ÑáÑ}¥¹ÁÕÐ 
½µÕ¹°Ù±ÕõÑ½Ì¹Ð ½µÕ¹¤½È¤(Ñ±½¹¼ôà¹ÑáÑ}¥¹ÁÕÐ Q³¥½¹¼°Ù±ÕõÑ½Ì¹Ð Ñ±½¹¼¤½È¤(¥É¥½¸ôÍÐ¹ÑáÑ}¥¹ÁÕÐ ¥É§Í¸°Ù±ÕõÑ½Ì¹Ð ¥É¥½¸¤½È¤(µ¥°ôÍÐ¹ÑáÑ}¥¹ÁÕÐ 5¥°°Ù±ÕõÑ½Ì¹Ð µ¥°¤½È¤(½ÌôÍÐ¹ÑáÑ}É =ÍÉÙ¥½¹Ì°¡¥¡ÐôàÀ¤((ÍÕµ¥ÑÑôÍÐ¹½Éµ}ÍÕµ¥Ñ}ÕÑÑ½¸ {Â~Jû"Guardar lead", use_container_width=True, type="primary")
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
                        st.success("ð Lead guardado correctamente")
                        st.balloons()
    else:
        st.markdown("""
        <div style="border:2px dashed #CBD5E0;border-radius:10px;padding:60px;text-align:center;background:white">
            <div style="font-size:48px">ð·</div>
            <p style="color:#718096;margin-top:12px;font-size:15px">
                Sube una captura del formulario de contacto<br>
                <strong>y la IA extraerÃ¡ los datos automÃ¡ticamente</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
