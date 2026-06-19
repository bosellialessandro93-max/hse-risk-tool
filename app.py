import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from datetime import datetime


# =========================
# CONFIG STREAMLIT
# =========================
st.set_page_config(
    page_title="HSE Risk Platform",
    page_icon="🦺",
    layout="wide"
)


# =========================
# FUNZIONI
# =========================
def parse_nc_data(nc_data):
    nc_types = []
    nc_themes = []

    if not nc_data.strip():
        return nc_types, nc_themes

    rows = nc_data.split("|")

    for r in rows:
        parts = r.split(",")

        if len(parts) != 2:
            continue

        theme = parts[0].strip().lower()
        severity = parts[1].strip().lower()

        if severity not in ["lieve", "media", "grave", "critica"]:
            continue

        nc_themes.append(theme)
        nc_types.append(severity)

    return nc_types, nc_themes


def calculate_risk(data):
    weights = {
        "lieve": 1,
        "media": 2,
        "grave": 4,
        "critica": 6
    }

    severity = sum(weights.get(x.lower(), 0) for x in data["nc_types"])
    nc_score = min(data["num_nc"] * 2 + severity, 100)

    stop_event = data["stopworks"] * 8
    stop_culture = min(data["stopworks"] * 2, 10)
    stop_score = max(0, min(stop_event, 100) - stop_culture)

    crit_score = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 15 +
        data["crit_ontime"] * 5,
        100
    )

    inspection_score = max(
        0,
        60 - data["inspections"] * 3 + data["num_nc"] * 4
    )

    complexity = data["app"] + data["sub"] * 1.5
    complexity_score = min(complexity * 5, 100)

    activity_score = min((data["activities"] ** 1.5) * 4, 100)

    nc_density = data["num_nc"] / (data["inspections"] + 1)

    penalty_density = 30 if nc_density > 1 else 15 if nc_density > 0.5 else 0
    penalty_control = 30 if data["inspections"] < 3 and data["num_nc"] >= 3 else 0
    interference_penalty = 15 if data["activities"] >= 3 else 0

    electrical_risk = 0
    excavation_risk = 0

    if data["activity_type"] == "elettrici":
        electrical_risk += 10

        if data["awareness"] < 3:
            electrical_risk += 10

        if "elettrico" in data["nc_themes"]:
            electrical_risk += 15

    if data["activity_type"] == "scavi":
        excavation_risk += 10

        if "scavi" in data["nc_themes"]:
            excavation_risk += 20

        if data["activities"] >= 3:
            excavation_risk += 10

    fase_penalty = 0

    if data["fase"] == "commissioning":
        fase_penalty += 15

        if data["activity_type"] == "elettrici" and data["awareness"] < 5:
            fase_penalty += 15

    complexity_penalty = 10 if data["sub"] > 5 else 0

    awareness_ratio = data["awareness"] / (
        data["activities"] + data["inspections"] + 1
    )

    awareness_bonus = 20 if awareness_ratio > 1 else 10 if awareness_ratio > 0.5 else 0

    risk = (
        nc_score * 0.15 +
        stop_score * 0.10 +
        crit_score * 0.15 +
        inspection_score * 0.15 +
        complexity_score * 0.10 +
        activity_score * 0.10 +
        penalty_density +
        penalty_control +
        interference_penalty +
        electrical_risk +
        excavation_risk +
        fase_penalty +
        complexity_penalty -
        awareness_bonus
    )

    risk = max(0, min(100, risk))

    if risk <= 20:
        level = "Basso"
    elif risk <= 40:
        level = "Medio basso"
    elif risk <= 60:
        level = "Medio"
    elif risk <= 80:
        level = "Alto"
    else:
        level = "Critico"

    details = {
        "NC": nc_score,
        "Stop Work": stop_score,
        "Criticità": crit_score,
        "Ispezioni": inspection_score,
        "Complessità": complexity_score,
        "Attività": activity_score,
        "Penalità densità NC": penalty_density,
        "Penalità controllo": penalty_control,
        "Penalità interferenze": interference_penalty,
        "Rischio elettrico": electrical_risk,
        "Rischio scavi": excavation_risk,
        "Penalità fase": fase_penalty,
        "Penalità complessità": complexity_penalty,
        "Bonus awareness": -awareness_bonus
    }

    return risk, level, details


def generate_actions(data):
    actions = []

    nc = data["num_nc"]
    inspections = data["inspections"]
    awareness = data["awareness"]
    activities = data["activities"]
    activity_type = data["activity_type"]

    nc_density = nc / (inspections + 1)

    if nc > inspections:
        actions.append((
            "MUST HAVE",
            f"🔴 Sono state rilevate {nc} NC a fronte di sole {inspections} ispezioni. "
            "Il sistema appare più reattivo che preventivo."
        ))

    if nc_density > 1:
        actions.append((
            "MUST HAVE",
            f"🔴 La densità NC è molto elevata: {round(nc_density, 2)}. "
            "Serve aumentare il presidio operativo."
        ))

    if inspections < 3 and nc >= 3:
        actions.append((
            "MUST HAVE",
            "🔴 Poche ispezioni rispetto alle NC rilevate. Pianificare controlli HSE più frequenti."
        ))

    if inspections > 20 and nc < 3:
        actions.append((
            "NICE TO HAVE",
            f"🟢 Le {inspections} ispezioni hanno generato solo {nc} NC. "
            "Il sistema di controllo sembra efficace."
        ))

    if awareness < inspections:
        actions.append((
            "MUST HAVE",
            f"🟡 Le sensibilizzazioni ({awareness}) sono inferiori alle ispezioni ({inspections}). "
            "Rafforzare cultura sicurezza e comunicazione operativa."
        ))

    if awareness > 30:
        actions.append((
            "NICE TO HAVE",
            f"🟢 Elevato numero di sensibilizzazioni: {awareness}. "
            "Elemento positivo per la riduzione del rischio comportamentale."
        ))

    if activities >= 3:
        actions.append((
            "MUST HAVE",
            f"🔴 Sono presenti {activities} attività contemporanee. "
            "Gestire interferenze, viabilità, aree segregate e coordinamento."
        ))

    if activity_type == "elettrici":
        actions.append((
            "MUST HAVE",
            "⚡ Attività elettriche presenti. Verificare procedure, autorizzazioni, sezionamenti e DPI."
        ))

    if activity_type == "scavi":
        actions.append((
            "MUST HAVE",
            "🚧 Attività di scavo presenti. Verificare sottoservizi, fronti di scavo, accessi e stabilità."
        ))

    if "elettrico" in data["nc_themes"]:
        actions.append((
            "MUST HAVE",
            "🔴 Rilevate NC su rischio elettrico. Necessario intervento immediato."
        ))

    if "scavi" in data["nc_themes"]:
        actions.append((
            "MUST HAVE",
            "🔴 Rilevate NC su scavi. Verificare condizioni operative prima della prosecuzione."
        ))

    if data["fase"] == "commissioning":
        actions.append((
            "MUST HAVE",
            "🟠 Fase di commissioning: aumentare controllo su prove, energie, autorizzazioni e interferenze."
        ))

    if len(actions) < 3:
        actions.append((
            "IDEA",
            "Mantenere controllo periodico, tracciamento NC e miglioramento continuo."
        ))

    return actions


def generate_ppt(nome, risk, level, df, actions, data):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}\nData: {datetime.now().strftime('%d/%m/%Y')}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risultato valutazione"
    slide.placeholders[1].text = (
        f"Risk Index: {round(risk)} / 100\n"
        f"Livello rischio: {level}\n\n"
        f"Fase: {data['fase']}\n"
        f"Tipologia attività: {data['activity_type']}"
    )

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Driver rischio"

    table = slide.shapes.add_table(
        len(df) + 1,
        2,
        Inches(1),
        Inches(1.4),
        Inches(7.5),
        Inches(4)
    ).table

    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"], 1))

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni suggerite"

    text = ""
    for lvl, action in actions:
        text += f"{lvl} - {action}\n\n"

    slide.placeholders[1].text = text

    return prs


def generate_excel(nome, risk, level, df, actions, data):
    buffer = BytesIO()

    summary = pd.DataFrame([{
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Cantiere": nome,
        "Risk Index": round(risk, 1),
        "Livello": level,
        "Fase": data["fase"],
        "Tipologia attività": data["activity_type"],
        "Ispezioni": data["inspections"],
        "NC": data["num_nc"],
        "Stop Work": data["stopworks"],
        "Criticità aperte": data["crit_open"],
        "Criticità in tempo": data["crit_ontime"],
        "Criticità in ritardo": data["crit_late"],
        "Appaltatori": data["app"],
        "Subappaltatori": data["sub"],
        "Sensibilizzazioni": data["awareness"],
        "Attività contemporanee": data["activities"]
    }])

    actions_df = pd.DataFrame(actions, columns=["Priorità", "Azione"])

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Sintesi", index=False)
        df.to_excel(writer, sheet_name="Driver rischio", index=False)
        actions_df.to_excel(writer, sheet_name="Azioni", index=False)

    buffer.seek(0)
    return buffer


# =========================
# UI
# =========================
st.title("🦺 HSE Risk Platform")
st.caption("Valutazione rischio HSE per cantieri, attività critiche e interferenze operative.")

with st.sidebar:
    st.header("📋 Input valutazione")

    nome = st.text_input("Nome cantiere", value="Cantiere Demo")

    fase = st.selectbox(
        "Fase",
        ["cantierizzazione", "costruzione", "commissioning", "punch list"]
    )

    st.divider()

    inspections = st.number_input("Ispezioni", 0, 1000, 10)
    num_nc = st.number_input("Numero NC", 0, 100, 0)

    nc_data = st.text_area(
        "NC",
        placeholder="Esempio: elettrico,grave | scavi,media | viabilità,lieve"
    )

    stop = st.number_input("Stop Work", 0, 50, 0)

    st.divider()

    crit_open = st.number_input("Criticità aperte", 0, 100, 0)
    crit_ontime = st.number_input("Criticità risolte in tempo", 0, 100, 0)
    crit_late = st.number_input("Criticità risolte in ritardo", 0, 100, 0)

    st.divider()

    app = st.number_input("Appaltatori", 0, 50, 0)
    sub = st.number_input("Subappaltatori", 0, 50, 0)

    awareness = st.number_input("Sensibilizzazioni", 0, 1000, 0)

    activity_type = st.selectbox(
        "Tipologia attività",
        ["scavi", "elettrici", "altro"]
    )

    activities = st.number_input("Numero attività contemporanee", 0, 20, 0)

    calcola = st.button("Calcola rischio", type="primary")


# =========================
# CALCOLO
# =========================
if calcola:
    nc_themes, nc_types = parse_nc_data(nc_data)

    data = {
        "inspections": inspections,
        "num_nc": num_nc,
        "nc_types": nc_types,
        "nc_themes": nc_themes,
        "stopworks": stop,
        "crit_open": crit_open,
        "crit_ontime": crit_ontime,
        "crit_late": crit_late,
        "app": app,
        "sub": sub,
        "awareness": awareness,
        "fase": fase,
        "activity_type": activity_type,
        "activities": activities
    }

    risk, level, details = calculate_risk(data)
    actions = generate_actions(data)

    df = pd.DataFrame({
        "Componente": list(details.keys()),
        "Valore": list(details.values())
    })

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Risk Index", f"{round(risk)} / 100")
    col2.metric("Livello", level)
    col3.metric("NC", num_nc)
    col4.metric("Ispezioni", inspections)

    st.divider()

    left, right = st.columns([1.2, 1])

    with left:
        st.subheader("📊 Driver rischio")
        st.bar_chart(df.set_index("Componente"))

    with right:
        st.subheader("📌 Sintesi valutazione")

        st.write(f"""
        Il cantiere **{nome}** ha un Risk Index pari a **{round(risk)} / 100**.

        Il livello di rischio è classificato come **{level}**.

        I principali elementi considerati sono:
        - NC rilevate: **{num_nc}**
        - Ispezioni effettuate: **{inspections}**
        - Stop Work: **{stop}**
        - Attività contemporanee: **{activities}**
        - Fase: **{fase}**
        - Tipologia attività: **{activity_type}**
        """)

        if nc_data and len(nc_types) == 0:
            st.warning(
                "Attenzione: le NC inserite non rispettano il formato richiesto. "
                "Usa ad esempio: elettrico,grave | scavi,media"
            )

    st.divider()

    st.subheader("🎯 Azioni suggerite")

    actions_df = pd.DataFrame(actions, columns=["Priorità", "Azione"])
    st.dataframe(actions_df, use_container_width=True)

    st.divider()

    st.subheader("📤 Export report")

    ppt = generate_ppt(nome, risk, level, df, actions, data)
    ppt_buffer = BytesIO()
    ppt.save(ppt_buffer)
    ppt_buffer.seek(0)

    excel_buffer = generate_excel(nome, risk, level, df, actions, data)

    safe_nome = nome.replace(" ", "_").replace("/", "_")

    col_a, col_b = st.columns(2)

    with col_a:
        st.download_button(
            label="📥 Scarica report PowerPoint",
            data=ppt_buffer,
            file_name=f"HSE_Risk_Report_{safe_nome}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    with col_b:
        st.download_button(
            label="📥 Scarica report Excel",
            data=excel_buffer,
            file_name=f"HSE_Risk_Report_{safe_nome}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Compila i dati nella sidebar e clicca su **Calcola rischio**.")
