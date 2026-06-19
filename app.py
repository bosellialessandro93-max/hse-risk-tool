import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO


# ==========================================
# CALCOLO RISCHIO
# ==========================================
def calculate_risk(data):

    weights = {"lieve": 1, "media": 2, "grave": 4, "critica": 6}

    severity = sum(weights.get(x.lower(), 0) for x in data["nc_types"])
    ncScore = min(data["num_nc"] * 2 + severity, 100)

    stop_event = data["stopworks"] * 8
    stop_culture = min(data["stopworks"] * 2, 10)
    stopScore = min(stop_event, 100) - stop_culture

    critScore = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 15 +
        data["crit_ontime"] * 5,
        100
    )

    inspectionScore = max(10, 50 - data["inspections"] * 2 + data["num_nc"] * 1.5)

    complexity = data["app"] + data["sub"] * 1.5
    complexityScore = min(complexity * 5, 100)

    # ATTIVITÀ (NUOVO)
    activityScore = min(
        data["activities"] * 5 +
        (5 if data["activity_type"] == "elettrici" else 3),
        100
    )

    awareness_bonus = min(data["awareness"] * 0.1, 15)

    fase_weights = {
        "cantierizzazione": 20,
        "costruzione": 40,
        "commissioning": 60,
        "punch list": 30
    }

    faseScore = fase_weights.get(data["fase"], 30)

    risk = (
        ncScore * 0.20 +
        stopScore * 0.15 +
        critScore * 0.20 +
        inspectionScore * 0.10 +
        complexityScore * 0.10 +
        faseScore * 0.10 +
        activityScore * 0.15
    ) - awareness_bonus

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

    return risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore


# ==========================================
# AZIONI INTELLIGENTI
# ==========================================
def generate_actions(data):

    actions = []

    nc = data["num_nc"]
    inspections = data["inspections"]
    awareness = data["awareness"]
    activities = data["activities"]
    activity_type = data["activity_type"]

    crit_open = data["crit_open"]
    crit_late = data["crit_late"]
    crit_ontime = data["crit_ontime"]

    total_crit = crit_open + crit_late + crit_ontime
    late_ratio = (crit_late / total_crit) if total_crit > 0 else 0

    # ✅ ISPEZIONI vs NC
    if inspections > 20 and nc < 3:
        actions.append(("NICE TO HAVE",
            f"🟢 Elevato numero di ispezioni ({inspections}) e poche NC ({nc}) → sistema efficace."))

    if inspections < 5 and nc > 3:
        actions.append(("MUST HAVE",
            f"🔴 Poche ispezioni ({inspections}) e molte NC ({nc}) → rischio elevato e bassa prevenzione."))

    if nc > inspections:
        actions.append(("MUST HAVE",
            f"🔴 NC ({nc}) superiori alle ispezioni ({inspections}) → sistema reattivo."))

    # ✅ AWARENESS
    if awareness > 30:
        actions.append(("NICE TO HAVE",
            f"🟢 Buon livello di sensibilizzazione ({awareness}) → riduzione del rischio."))

    if awareness < inspections:
        actions.append(("MUST HAVE",
            f"🟡 Sensibilizzazioni ({awareness}) inferiori alle ispezioni ({inspections}) → aumentare."))

    # ✅ CRITICITÀ
    if late_ratio > 0.3:
        actions.append(("MUST HAVE",
            f"🔴 {round(late_ratio*100)}% criticità chiuse in ritardo → esposizione prolungata."))

    if crit_open > 10:
        actions.append(("MUST HAVE",
            f"🔴 {crit_open} criticità aperte → necessario piano chiusura."))

    if crit_ontime > crit_late:
        actions.append(("NICE TO HAVE",
            "🟢 Buona gestione delle criticità (chiuse nei tempi)."))

    # ✅ ATTIVITÀ
    if activities >= 3:
        actions.append(("MUST HAVE",
            f"🔴 {activities} attività simultanee → alto rischio interferenze."))

        if activity_type == "elettrici":
            actions.append(("MUST HAVE",
                "⚡ Sensibilizzazione rischio elettrico consigliata (LOTO, PES/PAV)."))

        if activity_type == "scavi":
            actions.append(("MUST HAVE",
                "🚧 Sensibilizzazione scavi: seppellimento e sottoservizi."))

    if len(actions) < 3:
        actions.append(("IDEA", "👉 Mantenere controllo e miglioramento continuo."))

    return actions


# ==========================================
# PPT
# ==========================================
def generate_ppt(nome, risk, level, df, actions, data):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risultato"
    slide.placeholders[1].text = f"Risk Index: {round(risk)} - {level}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Attività"
    slide.placeholders[1].text = f"{data['activity_type']} - {data['activities']} attività"

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Driver rischio"

    table = slide.shapes.add_table(len(df) + 1, 2, Inches(1), Inches(1.5), Inches(6), Inches(3)).table
    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"]))

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni"

    text = ""
    for lvl, a in actions:
        text += f"{lvl} - {a}\n"

    slide.placeholders[1].text = text

    return prs


# ==========================================
# UI
# ==========================================
st.title("🦺 HSE Risk Platform")

nome = st.text_input("Nome cantiere")

fase = st.selectbox("Fase", ["cantierizzazione", "costruzione", "commissioning", "punch list"])

inspections = st.number_input("Ispezioni", 0, 1000, 10)
num_nc = st.number_input("Numero NC", 0, 100, 0)

nc_data = st.text_area("NC (es: quota,grave | elettrico,media)")

stop = st.number_input("Stop Work", 0, 50, 0)

crit_open = st.number_input("Criticità aperte", 0, 100, 0)
crit_ontime = st.number_input("Criticità risolte in tempo", 0, 100, 0)
crit_late = st.number_input("Criticità risolte in ritardo", 0, 100, 0)

app = st.number_input("Appaltatori", 0, 50, 0)
sub = st.number_input("Subappaltatori", 0, 50, 0)

awareness = st.number_input("Sensibilizzazioni", 0, 1000, 0)

# ✅ NUOVE FEATURES
activity_type = st.selectbox("Tipologia attività", ["scavi", "elettrici", "altro"])
activities = st.number_input("Numero attività contemporanee", 0, 20, 0)


# ==========================================
# CALCOLO
# ==========================================
if st.button("Calcola"):

    nc_types = []

    if nc_data:
        for r in nc_data.split("|"):
            parts = r.split(",")
            if len(parts) == 2:
                nc_types.append(parts[1].strip())

    data = {
        "inspections": inspections,
        "num_nc": num_nc,
        "nc_types": nc_types,
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

    result = calculate_risk(data)
    risk, level = result[0], result[1]

    actions = generate_actions(data)

    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità", "Attività"],
        "Valore": list(result[2:])
    })

    st.bar_chart(df.set_index("Componente"))

    st.subheader("🎯 Azioni suggerite")
    for lvl, a in actions:
        st.write(f"{lvl} - {a}")

    prs = generate_ppt(nome, risk, level, df, actions, data)
    buffer = BytesIO()
    prs.save(buffer)

    st.download_button("📥 Scarica PPT", buffer.getvalue(), "HSE_Report.pptx")
