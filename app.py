import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

# ==========================================
# CALCOLO RISCHIO AVANZATO
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

    inspectionScore = max(0, 60 - data["inspections"] * 3 + data["num_nc"] * 4)

    complexity = data["app"] + data["sub"] * 1.5
    complexityScore = min(complexity * 5, 100)

    activityScore = min((data["activities"] ** 1.5) * 4, 100)

    # ==============================
    # PENALITA'
    # ==============================

    nc_density = data["num_nc"] / (data["inspections"] + 1)

    if nc_density > 1:
        penalty_density = 30
    elif nc_density > 0.5:
        penalty_density = 15
    else:
        penalty_density = 0

    if data["inspections"] < 3 and data["num_nc"] >= 3:
        penalty_control = 30
    else:
        penalty_control = 0

    if data["activities"] >= 3:
        interference_penalty = 15
    else:
        interference_penalty = 0

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

    if data["sub"] > 5:
        complexity_penalty = 10
    else:
        complexity_penalty = 0

    # ==============================
    # BONUS
    # ==============================
    awareness_ratio = data["awareness"] / (data["activities"] + data["inspections"] + 1)

    if awareness_ratio > 1:
        awareness_bonus = 20
    elif awareness_ratio > 0.5:
        awareness_bonus = 10
    else:
        awareness_bonus = 0

    # ==============================
    # RISCHIO FINALE
    # ==============================
    risk = (
        ncScore * 0.15 +
        stopScore * 0.10 +
        critScore * 0.15 +
        inspectionScore * 0.15 +
        complexityScore * 0.10 +
        activityScore * 0.10
        + penalty_density
        + penalty_control
        + interference_penalty
        + electrical_risk
        + excavation_risk
        + fase_penalty
        + complexity_penalty
        - awareness_bonus
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

    return risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore


# ==========================================
# AZIONI
# ==========================================
def generate_actions(data):

    actions = []

    if data["num_nc"] > data["inspections"]:
        actions.append(("MUST HAVE", "Sistema reattivo: troppe NC rispetto alle ispezioni"))

    if data["awareness"] < data["inspections"]:
        actions.append(("MUST HAVE", "Aumentare sensibilizzazione per migliorare comportamento"))

    if data["activity_type"] == "elettrici":
        actions.append(("MUST HAVE", "Sensibilizzazione rischio elettrico"))

    if data["activity_type"] == "scavi":
        actions.append(("MUST HAVE", "Sensibilizzazione rischio scavi"))

    if data["activities"] >= 3:
        actions.append(("MUST HAVE", "Gestire interferenze tra attività simultanee"))

    return actions


# ==========================================
# UI
# ==========================================
st.title("🦺 HSE Risk Platform")

nome = st.text_input("Nome cantiere")

fase = st.selectbox("Fase", ["cantierizzazione", "costruzione", "commissioning", "punch list"])

inspections = st.number_input("Ispezioni", 0, 1000, 10)
num_nc = st.number_input("Numero NC", 0, 100, 0)

nc_data = st.text_area("NC (es: elettrico,grave | scavi,media)")

stop = st.number_input("Stop Work", 0, 50, 0)

crit_open = st.number_input("Criticità aperte", 0, 100, 0)
crit_ontime = st.number_input("Criticità risolte in tempo", 0, 100, 0)
crit_late = st.number_input("Criticità risolte in ritardo", 0, 100, 0)

app = st.number_input("Appaltatori", 0, 50, 0)
sub = st.number_input("Subappaltatori", 0, 50, 0)

awareness = st.number_input("Sensibilizzazioni", 0, 1000, 0)

activity_type = st.selectbox("Tipologia attività", ["scavi", "elettrici", "altro"])
activities = st.number_input("Numero attività contemporanee", 0, 20, 0)

# ==========================================
# CALCOLO
# ==========================================
if st.button("Calcola"):

    nc_types = []
    nc_themes = []

    if nc_data:
        for r in nc_data.split("|"):
            parts = r.split(",")
            if len(parts) == 2:
                nc_themes.append(parts[0].strip())
                nc_types.append(parts[1].strip())

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

    risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore = calculate_risk(data)

    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità", "Attività"],
        "Valore": [ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    actions = generate_actions(data)

    st.subheader("🎯 Azioni suggerite")
    for lvl, a in actions:
        st.write(f"{lvl} - {a}")
