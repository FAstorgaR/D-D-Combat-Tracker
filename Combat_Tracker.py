import streamlit as st
import pandas as pd
import random
import re
import json
import os
import uuid

SAVE_FOLDER = "encuentros"
AUTOSAVE_FILE = "_autosave.json"

if not os.path.exists(SAVE_FOLDER):
    os.makedirs(SAVE_FOLDER)

st.set_page_config(layout="wide", page_title="D&D Combat Tracker")

# --------------------------------------------------
# ESTILO
# --------------------------------------------------

st.markdown("""
<style>

.tooltip {
position: relative;
display: inline-block;
cursor: pointer;
}

.tooltip .tooltiptext {
visibility: hidden;
width: 420px;
background-color: #111;
color: #fff;
text-align: left;
border-radius: 6px;
padding: 10px;
position: absolute;
z-index: 1;
top: 110%;
left: 0;
font-size:12px;
}

.tooltip:hover .tooltiptext {
visibility: visible;
}

.barra-turno {
position: sticky;
top: 0;
z-index: 999;
background: linear-gradient(180deg, #0e1117 85%, transparent);
padding: 12px 14px;
margin-bottom: 16px;
border: 2px solid #00ff88;
border-radius: 10px;
}

.actor-card {
border-radius: 10px;
padding: 12px 14px;
margin-bottom: 14px;
background-color: #161b22;
border: 1px solid #30363d;
}

.card-jugador {
border-left: 5px solid #4a9eff;
}

.card-enemigo {
border-left: 5px solid #e85d5d;
}

.turno_activo {
border: 2px solid #00ff88 !important;
background-color: #0d1f16 !important;
box-shadow: 0 0 12px rgba(0, 255, 136, 0.25);
}

.card-caido {
opacity: 0.6;
filter: grayscale(35%);
}

.nombre_turno {
color: #00ff88;
font-weight: bold;
}

.sello-caido {
color: #ff4444;
font-weight: bold;
font-size: 0.9rem;
letter-spacing: 0.5px;
}

.hp-bar-wrap {
width: 100%;
height: 14px;
background-color: #2a2a2a;
border-radius: 7px;
overflow: hidden;
margin: 8px 0;
}

.hp-bar-fill {
height: 100%;
border-radius: 7px;
transition: width 0.2s ease;
}

.chip-condicion {
display: inline-block;
background-color: #2d333b;
color: #e6edf3;
border: 1px solid #444;
border-radius: 12px;
padding: 2px 10px;
margin: 2px 4px 2px 0;
font-size: 12px;
}

.stats-table {
width: 100%;
border-collapse: collapse;
margin: 6px 0 10px 0;
font-size: 11px;
}

.stats-table th,
.stats-table td {
border: 1px solid #333;
padding: 3px 6px;
text-align: center;
}

.stats-table th {
background-color: #1a1a1a;
}

.info-box {
border: 1px solid #444;
border-radius: 6px;
padding: 8px 10px;
margin: 6px 0;
background-color: #1a1a1a;
font-size: 13px;
}

.info-box b {
display: block;
margin-bottom: 4px;
color: #00ff88;
}

</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# CONDICIONES
# --------------------------------------------------

CONDICIONES = [
    "N/A",
    "Apresado (Grappled)",
    "Asustado (Frightened)",
    "Aturdido (Stunned)",
    "Cegado (Blinded)",
    "Derribado (Prone)",
    "Encantado (Charmed)",
    "Ensordecido (Deafened)",
    "Envenenado (Poisoned)",
    "Incapacitado (Incapacitated)",
    "Inconsciente (Unconscious)",
    "Invisible",
    "Paralizado (Paralyzed)",
    "Petrificado (Petrified)",
    "Restringido (Restrained)",
    "Exhaustion",
    "Otros"
]

# --------------------------------------------------
# BESTIARIO
# --------------------------------------------------

archivo_excel = "Bestiario v2.xlsx"

hojas = {
    "CR <1": "CR menor a 1",
    "CR 1-2": "CR 1 y 2",
    "CR 3": "CR3",
    "CR 4": "CR4",
    "CR 5": "CR5",
    "CR 6": "CR6",
    "CR 8": "CR8"
}

ATRIBUTOS = ["FUE", "DES", "CON", "INT", "SAB", "CAR"]
SUBATRIBUTOS = ["P", "M", "S"]
CAMPOS_INFO = ["Velocidad", "Iniciativa", "Habilidades", "Resistencias e Inmunidades", "Sentidos"]


def normalizar_columnas(df):

    nuevas = []
    ultimo_atributo = ""

    for col in df.columns:

        if isinstance(col, tuple):
            n0 = col[0]
            n1 = col[1] if len(col) > 1 else ""
        else:
            nuevas.append((str(col).strip(), ""))
            continue

        n0s = "" if pd.isna(n0) else str(n0).strip()
        if n0s and not n0s.startswith("Unnamed"):
            ultimo_atributo = n0s
        elif ultimo_atributo:
            n0s = ultimo_atributo

        n1s = "" if pd.isna(n1) else str(n1).strip()
        nuevas.append((n0s, n1s))

    df.columns = pd.MultiIndex.from_tuples(nuevas)
    return df


def buscar_columna(df, nombre):

    for col in df.columns:
        if str(col[0]).strip() == nombre:
            return col

    return None


def formatear_valor_stat(valor):

    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"

    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return "—"

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return texto


def extraer_stats_fila(fila):

    stats = {}

    for atributo in ATRIBUTOS:
        stats[atributo] = {
            sub: formatear_valor_stat(fila.get((atributo, sub), ""))
            for sub in SUBATRIBUTOS
        }

    return stats


def tiene_stats(stats):

    if not stats:
        return False

    return any(
        stats.get(atributo, {}).get(sub, "—") != "—"
        for atributo in ATRIBUTOS
        for sub in SUBATRIBUTOS
    )


def formatear_stats_html(stats):

    if not tiene_stats(stats):
        return ""

    filas = "".join(
        f"<tr><th>{atributo}</th>"
        f"<td>{stats[atributo]['P']}</td>"
        f"<td>{stats[atributo]['M']}</td>"
        f"<td>{stats[atributo]['S']}</td></tr>"
        for atributo in ATRIBUTOS
    )

    return f"""
<b>Atributos</b>
<table class="stats-table">
<tr><th></th><th>P</th><th>M</th><th>S</th></tr>
{filas}
</table>
"""


def formatear_texto_info(valor):

    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"

    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return "—"

    return texto.replace("\n", "<br>")


def extraer_info_fila(datos):

    return {
        campo: formatear_texto_info(datos.get(campo, ""))
        for campo in CAMPOS_INFO
    }


def tiene_info(info):

    if not info:
        return False

    return any(info.get(campo, "—") != "—" for campo in CAMPOS_INFO)


def formatear_info_html(info):

    if not tiene_info(info):
        return ""

    cajas = "".join(
        f"<div class='info-box'><b>{campo}</b>{info.get(campo, '—')}</div>"
        for campo in CAMPOS_INFO
        if info.get(campo, "—") != "—"
    )

    return cajas


def parsear_mod_iniciativa(texto):

    if not texto or texto == "—":
        return 0

    texto = str(texto).strip().replace("−", "-")
    match = re.search(r"([+\-]?\d+)", texto)

    if match:
        return int(match.group(1))

    return 0


def calcular_init_total(actor):

    tiro = int(actor.get("Init_tiro", 0))
    mod = int(actor.get("Init_mod", 0))

    if actor.get("aplicar_mod_init", True):
        return tiro + mod

    return tiro


def actualizar_init_actor(actor):

    actor["Init"] = calcular_init_total(actor)


def texto_init_desglose(actor):

    tiro = int(actor.get("Init_tiro", 0))
    mod = int(actor.get("Init_mod", 0))
    total = actor.get("Init", tiro)

    if actor.get("aplicar_mod_init", True) and mod != 0:
        return f"Init {total} ({tiro}{mod:+d})"

    return f"Init {total}"


def nuevo_uid():

    return str(uuid.uuid4())[:8]


def asegurar_uid(actor):

    if not actor.get("uid"):
        actor["uid"] = nuevo_uid()

    return actor["uid"]


def asegurar_uids_encuentro(encuentro):

    for actor in encuentro:
        asegurar_uid(actor)


@st.cache_data
def cargar_bestiario():

    bestiario = {}

    for cr, hoja in hojas.items():

        df = pd.read_excel(archivo_excel, sheet_name=hoja, header=[0, 1])
        df = normalizar_columnas(df)

        legend_col = None

        for col in df.columns:
            if "legend" in str(col[0]).lower():
                legend_col = col

        columnas = [
            "Nombre",
            "CA",
            "PG",
            "Velocidad",
            "Iniciativa",
            "Habilidades",
            "Resistencias e Inmunidades",
            "Sentidos",
            "Rasgos",
            "Acciones",
            "Acciones adicionales",
            "Reacciones"
        ]

        salida = pd.DataFrame()

        for col in columnas:
            origen = buscar_columna(df, col)
            salida[col] = df[origen] if origen is not None else ""

        if legend_col is not None:
            salida["Acciones legendarias"] = df[legend_col]
        else:
            salida["Acciones legendarias"] = ""

        salida["Stats"] = df.apply(extraer_stats_fila, axis=1)
        salida = salida.dropna(subset=["Nombre"])

        bestiario[cr] = salida

    return bestiario


# --------------------------------------------------
# FUNCIONES
# --------------------------------------------------

def tirar_dados(expresion):

    patron = r"(\d+)d(\d+)(\+\d+)?"
    match = re.search(patron, expresion)

    if not match:
        return int(expresion.split()[0])

    num = int(match.group(1))
    dado = int(match.group(2))

    bonus = match.group(3)
    bonus = int(bonus[1:]) if bonus else 0

    return sum(random.randint(1, dado) for _ in range(num)) + bonus


def aplicar_dano(actor, cantidad):

    if actor["temp_hp"] > 0:

        if cantidad <= actor["temp_hp"]:
            actor["temp_hp"] -= cantidad
            return

        cantidad -= actor["temp_hp"]
        actor["temp_hp"] = 0

    actor["HP"] -= cantidad

    if actor["HP"] < 0:
        actor["HP"] = 0


def curar(actor, cantidad):

    actor["HP"] += cantidad

    if actor["HP"] > actor["HP_max"]:
        actor["HP"] = actor["HP_max"]


def color_hp(actor):

    if actor["HP"] <= 0:
        return "#ff4444"

    ratio = actor["HP"] / actor["HP_max"]

    if ratio <= 0.25:
        return "#ff8c00"
    if ratio <= 0.50:
        return "#ffd700"

    return "#00ff88"


def barra_hp_html(actor):

    ratio = 0 if actor["HP_max"] <= 0 else min(actor["HP"] / actor["HP_max"], 1)
    pct = int(ratio * 100)
    color = color_hp(actor)

    return f"""
<div class="hp-bar-wrap">
<div class="hp-bar-fill" style="width:{pct}%; background-color:{color};"></div>
</div>
"""


def decrementar_condiciones(actor):

    nuevas = []

    for condicion in actor.get("Condiciones", []):
        condicion["duracion"] -= 1
        if condicion["duracion"] > 0:
            nuevas.append(condicion)

    actor["Condiciones"] = nuevas


def avanzar_turno():

    if not st.session_state.encuentro:
        return

    actor_actual = st.session_state.encuentro[st.session_state.turno]
    decrementar_condiciones(actor_actual)

    st.session_state.turno += 1

    if st.session_state.turno >= len(st.session_state.encuentro):
        st.session_state.turno = 0
        st.session_state.ronda += 1


def firma_iniciativas(encuentro):

    return tuple(sorted(int(a.get("Init", 0)) for a in encuentro))


def ordenar_por_iniciativa():

    encuentro = st.session_state.encuentro

    if not encuentro:
        return

    turno = min(st.session_state.turno, len(encuentro) - 1)
    activo = encuentro[turno]

    encuentro.sort(
        key=lambda a: (-int(a.get("Init", 0)), a.get("Nombre", ""))
    )

    for i, actor in enumerate(encuentro):
        if actor is activo:
            st.session_state.turno = i
            return


def actualizar_orden_iniciativa(force=False):

    encuentro = st.session_state.encuentro

    if not encuentro:
        return

    firma = firma_iniciativas(encuentro)

    if firma != st.session_state.get("firma_inits"):
        st.session_state.orden_manual = False
        st.session_state.firma_inits = firma

    if force or not st.session_state.get("orden_manual", False):
        ordenar_por_iniciativa()
        st.session_state.orden_manual = False


def datos_encuentro():

    return {
        "encuentro": st.session_state.encuentro,
        "turno": st.session_state.turno,
        "ronda": st.session_state.ronda,
        "sugerencias": st.session_state.sugerencias
    }


def autosave():

    ruta = os.path.join(SAVE_FOLDER, AUTOSAVE_FILE)

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos_encuentro(), f, ensure_ascii=False)


def guardar_encuentro(nombre):

    if not nombre.strip():
        st.warning("Escribe un nombre para guardar el encuentro.")
        return

    ruta = os.path.join(SAVE_FOLDER, f"{nombre.strip()}.json")

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos_encuentro(), f, ensure_ascii=False)

    autosave()
    st.success(f"Encuentro guardado como {nombre.strip()}")


def listar_encuentros():

    archivos = []

    for f in os.listdir(SAVE_FOLDER):
        if f.endswith(".json") and f != AUTOSAVE_FILE:
            archivos.append(f.replace(".json", ""))

    return sorted(archivos)


def cargar_datos(data):

    st.session_state.encuentro = data.get("encuentro", [])
    st.session_state.turno = data.get("turno", 0)
    st.session_state.ronda = data.get("ronda", 1)
    st.session_state.sugerencias = data.get("sugerencias", {})

    for actor in st.session_state.encuentro:
        actor.setdefault("tipo", "enemigo")
        actor.setdefault("Stats", {})
        actor.setdefault("Info", {})
        actor.setdefault("Init_tiro", actor.get("Init", 0))
        actor.setdefault("Init_mod", parsear_mod_iniciativa(actor.get("Info", {}).get("Iniciativa", "")))
        actor.setdefault("aplicar_mod_init", True)
        asegurar_uid(actor)
        actualizar_init_actor(actor)


def cargar_encuentro(nombre):

    ruta = os.path.join(SAVE_FOLDER, f"{nombre}.json")

    with open(ruta, encoding="utf-8") as f:
        cargar_datos(json.load(f))

    autosave()


def cargar_autosave():

    ruta = os.path.join(SAVE_FOLDER, AUTOSAVE_FILE)

    if not os.path.exists(ruta):
        return False

    with open(ruta, encoding="utf-8") as f:
        cargar_datos(json.load(f))

    return True


def sugerir_accion(actor):

    acciones = actor["Acciones"]

    if not acciones:
        return "Sin acciones registradas"

    lista = acciones.split("<br>")
    lista = [a.strip() for a in lista if a.strip()]

    if not lista:
        return "Sin acciones registradas"

    return random.choice(lista)


def clase_tarjeta(actor, turno):

    clases = ["actor-card"]

    if actor.get("tipo") == "jugador":
        clases.append("card-jugador")
    else:
        clases.append("card-enemigo")

    if turno:
        clases.append("turno_activo")

    if actor["HP"] <= 0:
        clases.append("card-caido")

    return " ".join(clases)


def render_sidebar_actor(actor, titulo):

    st.sidebar.markdown(f"### {titulo}")
    st.sidebar.markdown(f"**{actor['Nombre']}**")

    tipo = "Jugador" if actor.get("tipo") == "jugador" else "Enemigo"
    st.sidebar.caption(f"{tipo} | {texto_init_desglose(actor)}")

    stats = actor.get("Stats", {})
    info = actor.get("Info", {})

    st.sidebar.markdown(barra_hp_html(actor), unsafe_allow_html=True)
    st.sidebar.markdown(
        f"CA: {actor['CA']} | HP: {actor['HP']}/{actor['HP_max']} | Temp: {actor['temp_hp']}",
        unsafe_allow_html=True
    )

    if actor["HP"] <= 0:
        st.sidebar.error("INCONSCIENTE / CAÍDO")

    if actor.get("Condiciones"):
        chips = " ".join(
            f"<span class='chip-condicion'>{c['nombre']} ({c['duracion']})</span>"
            for c in actor["Condiciones"]
        )
        st.sidebar.markdown(chips, unsafe_allow_html=True)

    if tiene_stats(stats):
        st.sidebar.markdown("#### Atributos")
        st.sidebar.markdown(formatear_stats_html(stats), unsafe_allow_html=True)

    if tiene_info(info):
        st.sidebar.markdown("#### Ficha rápida")
        st.sidebar.markdown(formatear_info_html(info), unsafe_allow_html=True)

    st.sidebar.markdown("#### Rasgos")
    st.sidebar.markdown(actor["Rasgos"] or "—", unsafe_allow_html=True)

    st.sidebar.markdown("#### Acciones")
    st.sidebar.markdown(actor["Acciones"] or "—", unsafe_allow_html=True)

    if actor.get("Bonus"):
        st.sidebar.markdown("#### Bonus")
        st.sidebar.markdown(actor["Bonus"], unsafe_allow_html=True)

    if actor.get("Reacciones"):
        st.sidebar.markdown("#### Reacciones")
        st.sidebar.markdown(actor["Reacciones"], unsafe_allow_html=True)

    if actor.get("Legendarias"):
        st.sidebar.markdown("#### Legendarias")
        st.sidebar.markdown(actor["Legendarias"], unsafe_allow_html=True)


def render_actor_card(i, actor, minimal):

    uid = asegurar_uid(actor)
    turno = i == st.session_state.turno
    caido = actor["HP"] <= 0
    stats = actor.get("Stats", {})
    info = actor.get("Info", {})
    stats_html = formatear_stats_html(stats)
    info_html = formatear_info_html(info)
    clase = clase_tarjeta(actor, turno)

    st.markdown(f"<div class='{clase}'>", unsafe_allow_html=True)

    nombre_html = (
        f"<span class='nombre_turno'>{actor['Nombre']}</span>"
        if turno else actor["Nombre"]
    )

    tipo_badge = "🛡️" if actor.get("tipo") == "jugador" else "💀"

    header = st.columns([5, 1])

    with header[0]:
        if not minimal:
            tooltip = f"""
{stats_html}
{info_html}
<b>Rasgos</b><br>{actor['Rasgos']}<br><br>
<b>Acciones</b><br>{actor['Acciones']}<br><br>
<b>Bonus</b><br>{actor['Bonus']}<br><br>
<b>Reacciones</b><br>{actor['Reacciones']}<br><br>
<b>Legendarias</b><br>{actor['Legendarias']}
"""
            st.markdown(f"""
<div class="tooltip">
<h3>{tipo_badge} {nombre_html}</h3>
<span class="tooltiptext">{tooltip}</span>
</div>
""", unsafe_allow_html=True)
        else:
            st.markdown(f"### {tipo_badge} {nombre_html}", unsafe_allow_html=True)

    with header[1]:
        if not minimal and st.button("📖", key=f"view_{uid}", help="Ver statblock"):
            st.session_state.selected_index = i
            st.rerun()

    if caido:
        st.markdown("<span class='sello-caido'>⛔ INCONSCIENTE / CAÍDO</span>", unsafe_allow_html=True)

    st.markdown(barra_hp_html(actor), unsafe_allow_html=True)

    hp_color = color_hp(actor)
    linea_combate = (
        f"CA: {actor['CA']} | "
        f"<span style='color:{hp_color}'>HP: {actor['HP']}/{actor['HP_max']}</span> | "
        f"Temp: {actor['temp_hp']}"
    )

    if info.get("Velocidad", "—") != "—":
        linea_combate += f" | Vel: {info['Velocidad']}"

    if info.get("Iniciativa", "—") != "—":
        linea_combate += f" | Iniciativa (ficha): {info['Iniciativa']}"

    st.markdown(linea_combate, unsafe_allow_html=True)

    init_cols = st.columns([2, 2, 2, 2, 4])

    with init_cols[0]:
        actor["Init_tiro"] = st.number_input(
            "Tiro d20 (mesa)",
            0, 50,
            int(actor.get("Init_tiro", 0)),
            key=f"init_tiro_{uid}",
            help="Resultado del d20 tirado en la mesa"
        )

    with init_cols[1]:
        mod = int(actor.get("Init_mod", 0))
        if mod != 0 or actor.get("tipo") == "jugador":
            actor["Init_mod"] = st.number_input(
                "Mod. Init",
                -10, 15,
                mod,
                key=f"init_mod_{uid}",
                help="Modificador de iniciativa (ficha o manual)"
            )
        elif info.get("Iniciativa", "—") != "—":
            st.caption(f"Mod. ficha: {mod:+d}")

    with init_cols[2]:
        if int(actor.get("Init_mod", 0)) != 0:
            actor["aplicar_mod_init"] = st.checkbox(
                "Sumar mod.",
                value=actor.get("aplicar_mod_init", True),
                key=f"init_aplica_{uid}",
                help="Sumar el modificador de iniciativa al tiro de mesa"
            )
        else:
            actor["aplicar_mod_init"] = False

    actualizar_init_actor(actor)

    with init_cols[3]:
        st.markdown(f"**{texto_init_desglose(actor)}**")

    with init_cols[4]:
        flecha1, flecha2 = st.columns(2)
        with flecha1:
            if not minimal and st.button("⬆", key=f"up_{uid}", disabled=i == 0):
                st.session_state.encuentro[i], st.session_state.encuentro[i - 1] = \
                    st.session_state.encuentro[i - 1], st.session_state.encuentro[i]
                if st.session_state.turno == i:
                    st.session_state.turno = i - 1
                elif st.session_state.turno == i - 1:
                    st.session_state.turno = i
                st.session_state.orden_manual = True
                autosave()
                st.rerun()

        with flecha2:
            if not minimal and st.button("⬇", key=f"down_{uid}", disabled=i >= len(st.session_state.encuentro) - 1):
                st.session_state.encuentro[i], st.session_state.encuentro[i + 1] = \
                    st.session_state.encuentro[i + 1], st.session_state.encuentro[i]
                if st.session_state.turno == i:
                    st.session_state.turno = i + 1
                elif st.session_state.turno == i + 1:
                    st.session_state.turno = i
                st.session_state.orden_manual = True
                autosave()
                st.rerun()

    if actor.get("Condiciones"):
        chips = " ".join(
            f"<span class='chip-condicion'>{c['nombre']} ({c['duracion']}r)</span>"
            for c in actor["Condiciones"]
        )
        st.markdown(chips, unsafe_allow_html=True)

    if not minimal:
        if tiene_stats(stats):
            with st.expander("📊 Atributos (P / M / S)", expanded=False):
                st.markdown(stats_html, unsafe_allow_html=True)

        if tiene_info(info):
            with st.expander("📋 Velocidad / Iniciativa / Habilidades / Resistencias / Sentidos", expanded=False):
                st.markdown(info_html, unsafe_allow_html=True)

    daño_cols = st.columns([2, 1, 1, 1, 1, 2, 2])

    with daño_cols[0]:
        valor = st.number_input("Cantidad", 0, 999, key=f"valor_{uid}", disabled=caido)

    with daño_cols[1]:
        if st.button("+5", key=f"d5_{uid}", disabled=caido):
            aplicar_dano(actor, 5)
            autosave()
            st.rerun()

    with daño_cols[2]:
        if st.button("+10", key=f"d10_{uid}", disabled=caido):
            aplicar_dano(actor, 10)
            autosave()
            st.rerun()

    with daño_cols[3]:
        if st.button("+15", key=f"d15_{uid}", disabled=caido):
            aplicar_dano(actor, 15)
            autosave()
            st.rerun()

    with daño_cols[4]:
        medio = max(1, actor["HP"] // 2)
        if st.button("½ HP", key=f"dhalf_{uid}", disabled=caido):
            aplicar_dano(actor, medio)
            autosave()
            st.rerun()

    with daño_cols[5]:
        d1, d2, d3 = st.columns(3)
        if d1.button("Daño", key=f"d_{uid}", disabled=caido):
            aplicar_dano(actor, valor)
            autosave()
            st.rerun()
        if d2.button("Curar", key=f"c_{uid}", disabled=caido):
            curar(actor, valor)
            autosave()
            st.rerun()
        if d3.button("Temp", key=f"t_{uid}"):
            actor["temp_hp"] += valor
            autosave()
            st.rerun()

    with daño_cols[6]:
        if st.button("➡ Fin turno", key=f"next_{uid}", disabled=not turno):
            avanzar_turno()
            autosave()
            st.rerun()

    if not minimal:
        ctrl = st.columns(6)

        cond = ctrl[1].selectbox("Condición", CONDICIONES, key=f"cond_{uid}")
        duracion = ctrl[1].number_input("Duración (rondas)", 1, 20, 1, key=f"dur_{uid}")

        if cond == "Exhaustion":
            nivel = st.slider("Exhaustion", 1, 6, key=f"ex_{uid}")
            cond = f"Exhaustion {nivel}"

        if ctrl[1].button("Agregar", key=f"add_{uid}"):
            if cond != "N/A":
                actor["Condiciones"].append({
                    "nombre": cond,
                    "duracion": duracion
                })
                autosave()
                st.rerun()

        if actor["Condiciones"]:
            for j, c in enumerate(actor["Condiciones"]):
                c1, c2 = st.columns([6, 1])
                c1.write(f"{c['nombre']} ({c['duracion']} rondas)")
                if c2.button("❌", key=f"cond_del_{uid}_{j}"):
                    actor["Condiciones"].pop(j)
                    autosave()
                    st.rerun()

        if ctrl[2].button("➕ Duplicar", key=f"dup_{uid}"):
            nuevo = actor.copy()
            nuevo["uid"] = nuevo_uid()
            nuevo["Nombre"] = actor["Nombre"] + " copia"
            st.session_state.encuentro.insert(i + 1, nuevo)
            autosave()
            st.rerun()

        if ctrl[4].button("❌ Eliminar", key=f"del_{uid}"):
            st.session_state.encuentro.pop(i)
            if st.session_state.turno >= len(st.session_state.encuentro):
                st.session_state.turno = max(0, len(st.session_state.encuentro) - 1)
            autosave()
            st.rerun()

        if ctrl[5].button("🧠 Sugerir", key=f"sugerir_{uid}"):
            st.session_state.sugerencias[uid] = sugerir_accion(actor)
            autosave()
            st.rerun()

        if uid in st.session_state.sugerencias:
            st.info(f"💡 {st.session_state.sugerencias[uid]}")

    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

if "encuentro" not in st.session_state:
    st.session_state.encuentro = []

if "turno" not in st.session_state:
    st.session_state.turno = 0

if "ronda" not in st.session_state:
    st.session_state.ronda = 1

if "selected_index" not in st.session_state:
    st.session_state.selected_index = None

if "sugerencias" not in st.session_state:
    st.session_state.sugerencias = {}

if "orden_manual" not in st.session_state:
    st.session_state.orden_manual = False

if "firma_inits" not in st.session_state:
    st.session_state.firma_inits = ()

if "autosave_loaded" not in st.session_state:
    st.session_state.autosave_loaded = cargar_autosave()

bestiario = cargar_bestiario()

# --------------------------------------------------
# SIDEBAR
# --------------------------------------------------

st.sidebar.title("Referencia")

if st.session_state.encuentro:
    turno_idx = min(st.session_state.turno, len(st.session_state.encuentro) - 1)
    render_sidebar_actor(
        st.session_state.encuentro[turno_idx],
        f"Turno activo — Ronda {st.session_state.ronda}"
    )

    if (
        st.session_state.selected_index is not None
        and st.session_state.selected_index < len(st.session_state.encuentro)
        and st.session_state.selected_index != turno_idx
    ):
        st.sidebar.divider()
        render_sidebar_actor(
            st.session_state.encuentro[st.session_state.selected_index],
            "Statblock seleccionado"
        )
else:
    st.sidebar.info("Agrega participantes en la pestaña Preparar.")

# --------------------------------------------------
# TITULO
# --------------------------------------------------

st.title("🐉 D&D Combat Tracker")
st.caption("V4 — Mesa real | Iniciativa manual | Autoguardado activo")

tab_preparar, tab_combate = st.tabs(["🛠️ Preparar", "⚔️ Combate"])

# --------------------------------------------------
# PREPARAR
# --------------------------------------------------

with tab_preparar:

    st.subheader("Guardar / Cargar encuentro")

    save_col, load_col, auto_col = st.columns(3)

    with save_col:
        nombre_save = st.text_input("Nombre encuentro", key="nombre_save")
        if st.button("💾 Guardar encuentro", use_container_width=True):
            guardar_encuentro(nombre_save)

    with load_col:
        archivos = listar_encuentros()
        archivo_cargar = st.selectbox("Encuentros guardados", archivos if archivos else ["—"])
        if st.button("📂 Cargar encuentro", use_container_width=True, disabled=not archivos):
            cargar_encuentro(archivo_cargar)
            st.rerun()

    with auto_col:
        if st.session_state.autosave_loaded:
            st.success("Autoguardado restaurado")
        if st.button("🔄 Recuperar autoguardado", use_container_width=True):
            if cargar_autosave():
                st.session_state.autosave_loaded = True
                st.rerun()
            else:
                st.warning("No hay autoguardado disponible.")

    st.divider()

    prep1, prep2 = st.columns(2)

    with prep1:
        st.subheader("Agregar enemigo")

        cr = st.selectbox("CR", list(bestiario.keys()), key="cr_select")
        busqueda = st.text_input("Buscar monstruo", key="busqueda")
        df_cr = bestiario[cr]

        if busqueda:
            df_cr = df_cr[df_cr["Nombre"].str.contains(busqueda, case=False)]

        nombre = st.selectbox("Monstruo", df_cr["Nombre"], key="monstruo_select")
        cantidad = st.number_input("Cantidad", 1, 20, 1, key="cantidad")

        if not df_cr.empty:
            preview = df_cr[df_cr["Nombre"] == nombre].iloc[0]
            preview_info = extraer_info_fila(preview)
            preview_mod = parsear_mod_iniciativa(preview_info.get("Iniciativa", ""))
            if preview_mod != 0 or preview_info.get("Iniciativa", "—") != "—":
                st.caption(f"Mod. iniciativa (ficha): **{preview_mod:+d}** ({preview_info.get('Iniciativa', '—')})")

        init_tiro_enemigo = st.number_input(
            "Tiro d20 (mesa) al agregar",
            0, 50, 0,
            key="init_tiro_enemigo",
            help="Resultado del d20 en la mesa. El mod. de ficha se suma si está activo abajo."
        )
        aplicar_mod_enemigo = st.checkbox(
            "Sumar mod. de iniciativa de la ficha",
            value=True,
            key="aplicar_mod_enemigo",
            help="Suma el modificador del bestiario al tiro de mesa"
        )

        if st.button("➕ Agregar enemigo", use_container_width=True):

            datos = df_cr[df_cr["Nombre"] == nombre].iloc[0]
            hp_text = str(datos["PG"])
            info_datos = extraer_info_fila(datos)
            init_mod = parsear_mod_iniciativa(info_datos.get("Iniciativa", ""))

            for n in range(cantidad):

                hp_valor = tirar_dados(hp_text)
                init_tiro = int(init_tiro_enemigo)

                enemigo = {
                    "uid": nuevo_uid(),
                    "Nombre": f"{datos['Nombre']} {n + 1}",
                    "tipo": "enemigo",
                    "CA": int(datos["CA"]),
                    "HP": hp_valor,
                    "HP_max": hp_valor,
                    "temp_hp": 0,
                    "Init_tiro": init_tiro,
                    "Init_mod": init_mod,
                    "aplicar_mod_init": aplicar_mod_enemigo and init_mod != 0,
                    "Condiciones": [],
                    "Rasgos": str(datos["Rasgos"]).replace("\n", "<br>"),
                    "Acciones": str(datos["Acciones"]).replace("\n", "<br>"),
                    "Bonus": str(datos["Acciones adicionales"]).replace("\n", "<br>"),
                    "Reacciones": str(datos["Reacciones"]).replace("\n", "<br>"),
                    "Legendarias": str(datos["Acciones legendarias"]).replace("\n", "<br>"),
                    "Stats": datos["Stats"] if isinstance(datos["Stats"], dict) else extraer_stats_fila(datos),
                    "Info": info_datos
                }
                actualizar_init_actor(enemigo)

                st.session_state.encuentro.append(enemigo)

            st.session_state.orden_manual = False
            autosave()
            st.success(f"Se agregaron {cantidad} enemigo(s).")
            st.rerun()

    with prep2:
        st.subheader("Agregar jugador")

        nombre_jugador = st.text_input("Nombre jugador", key="nombre_jugador")
        ca = st.number_input("CA jugador", 0, 30, 10, key="ca_jugador")
        hp = st.number_input("HP jugador", 0, 500, 10, key="hp_jugador")
        init_tiro_jugador = st.number_input(
            "Tiro d20 (mesa) al agregar",
            0, 50, 0,
            key="init_tiro_jugador",
            help="Resultado del d20 en la mesa"
        )
        init_mod_jugador = st.number_input(
            "Mod. iniciativa",
            -10, 15, 0,
            key="init_mod_jugador",
            help="Modificador de iniciativa del personaje"
        )
        aplicar_mod_jugador = st.checkbox(
            "Sumar mod. al tiro",
            value=True,
            key="aplicar_mod_jugador"
        )

        if st.button("➕ Agregar jugador", use_container_width=True):

            init_tiro = int(init_tiro_jugador)
            init_mod = int(init_mod_jugador)

            jugador = {
                "uid": nuevo_uid(),
                "Nombre": nombre_jugador,
                "tipo": "jugador",
                "CA": ca,
                "HP": hp,
                "HP_max": hp,
                "temp_hp": 0,
                "Init_tiro": init_tiro,
                "Init_mod": init_mod,
                "aplicar_mod_init": aplicar_mod_jugador and init_mod != 0,
                "Condiciones": [],
                "Rasgos": "",
                "Acciones": "",
                "Bonus": "",
                "Reacciones": "",
                "Legendarias": "",
                "Stats": {},
                "Info": {}
            }
            actualizar_init_actor(jugador)

            st.session_state.encuentro.append(jugador)
            st.session_state.orden_manual = False
            autosave()
            st.success(f"Jugador {nombre_jugador} agregado.")
            st.rerun()

    st.divider()
    st.subheader("Participantes actuales")

    if st.session_state.encuentro:
        st.write(f"**{len(st.session_state.encuentro)}** participantes en el encuentro.")
        st.caption("En Combate se ordenan por Init total (mayor primero). Usa ⬆ / ⬇ para empates o ajustes manuales.")

        for i, actor in enumerate(st.session_state.encuentro):
            uid = asegurar_uid(actor)
            tipo = "🛡️" if actor.get("tipo") == "jugador" else "💀"
            fila, btn = st.columns([6, 1])

            with fila:
                st.write(
                    f"{tipo} **{actor['Nombre']}** — "
                    f"{texto_init_desglose(actor)} — "
                    f"HP: {actor['HP']}/{actor['HP_max']}"
                )

            with btn:
                if st.button("❌", key=f"prep_del_{uid}", help="Eliminar participante"):
                    st.session_state.encuentro.pop(i)
                    if st.session_state.turno >= len(st.session_state.encuentro):
                        st.session_state.turno = max(0, len(st.session_state.encuentro) - 1)
                    st.session_state.orden_manual = False
                    autosave()
                    st.rerun()
    else:
        st.info("Aún no hay participantes. Agrega enemigos o jugadores arriba.")

    st.caption(f"Bestiario cargado: {sum(len(df) for df in bestiario.values())} criaturas")

# --------------------------------------------------
# COMBATE
# --------------------------------------------------

with tab_combate:

    vista_min = st.toggle("Vista minimalista (proyectar en pantalla)", value=False)

    if not st.session_state.encuentro:
        st.info("No hay participantes. Ve a la pestaña Preparar para armar el encuentro.")
    else:
        actualizar_orden_iniciativa()

        activo = st.session_state.encuentro[st.session_state.turno]

        orden_hint = "orden manual" if st.session_state.orden_manual else "ordenado por Init ↓"
        st.markdown(f"""
<div class="barra-turno">
<b>Ronda {st.session_state.ronda}</b> —
Turno activo: <span style="color:#00ff88">{activo['Nombre']}</span>
({texto_init_desglose(activo)}) · <span style="font-size:0.85em">{orden_hint}</span>
</div>
""", unsafe_allow_html=True)

        barra = st.columns([2, 2, 2, 2, 2])

        if barra[0].button("⏭️ Siguiente turno", use_container_width=True, type="primary"):
            avanzar_turno()
            autosave()
            st.rerun()

        if barra[1].button("🔢 Ordenar por Init", use_container_width=True):
            st.session_state.orden_manual = False
            actualizar_orden_iniciativa(force=True)
            autosave()
            st.rerun()

        if barra[2].button("🔁 Reiniciar ronda", use_container_width=True):
            st.session_state.turno = 0
            st.session_state.ronda = 1
            autosave()
            st.rerun()

        if barra[3].button("💾 Autoguardar ahora", use_container_width=True):
            autosave()
            st.toast("Encuentro autoguardado")

        if barra[4].button("🗑️ Vaciar encuentro", use_container_width=True):
            st.session_state.encuentro = []
            st.session_state.turno = 0
            st.session_state.ronda = 1
            st.session_state.sugerencias = {}
            autosave()
            st.rerun()

        st.markdown("""
<script>
function scrollToActive(){
const active = window.parent.document.querySelector('.turno_activo');
if(active){
active.scrollIntoView({behavior:"smooth", block:"center"});
}
}
setTimeout(scrollToActive, 200);
</script>
""", unsafe_allow_html=True)

        for i, actor in enumerate(st.session_state.encuentro):
            render_actor_card(i, actor, vista_min)
