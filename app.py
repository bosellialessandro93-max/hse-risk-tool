import streamlit as st

def calculate_risk(data):

    weights = {
        "lieve": 1,
        "media": 2,
        "grave": 4,
        "critica": 6
    }

    severity = sum(weights.get(x.strip().lower(), 0) for x in data["nc_types"])
    ncScore = min((data["num_nc"] * 2 + severity), 100)

    stopWorkScore = min(data["stopworks"] * 20, 100)
    criticalityScore = min(data["criticalities"] * 10, 100)
    inspectionScore = max(10, 50 - data["inspections"] * 2)

    complexity = data["appaltatori"] + (data["subappaltatori"] * 1.5)
    complexityScore = min(complexity * 5, 100)

    awarenessMitigation = min(data["awareness"] * 0.1, 15)

    riskScore = (
        ncScore * 0.30 +
        stopWorkScore * 0.25 +
        criticalityScore * 0.20 +
        inspectionScore * 0.10 +
        complexityScore * 0.15
    ) - awarenessMitigation

    riskScore = max(0, min(100, riskScore))

    if riskScore <= 20:
        level = "Basso"
    elif riskScore <= 40:
        level = "Medio basso"
    elif riskScore <= 60:
        level = "Medio"
    elif riskScore <= 80:
        level = "Alto"
    else:
        level = "Critico"

    return riskScore, level


def suggest_actions(level):
    actions_map = {
        "Basso": ["Monitoraggio continuo", "Sensibilizzazioni"],
        "Medio basso": ["Verifica NC", "Formazione mirata"],
        "Medio": ["Analisi cause", "Aumentare ispezioni", "Briefing appaltatori"],
        "Alto": ["Audit HSE", "Piano correttivo"],
        "Critico": ["Escalation immediata", "Stop attività"]
    }
    return actions_map[level]


st.title("🦺 HSE Risk Tool")

inspections = st.number_input("Numero ispezioni", 0, 1000, 10)
num_nc = st.number_input("Numero NC", 0, 100, 0)
nc_input = st.text_input("Tipologia NC (es: media,media,lieve)")
stopworks = st.number_input("Stop Work", 0, 50, 0)
criticalities = st.number_input("Criticità aperte", 0, 100, 0)
appaltatori = st.number_input("Appaltatori", 0, 50, 0)
subappaltatori = st.number_input("Subappaltatori", 0, 50, 0)
awareness = st.number_input("Persone sensibilizzate", 0, 1000, 0)

if st.button("Calcola rischio"):

    nc_types = [x.strip() for x in nc_input.split(",") if x.strip() != ""]

    data = {
        "inspections": inspections,
        "num_nc": num_nc,
        "nc_types": nc_types,
        "stopworks": stopworks,
        "criticalities": criticalities,
        "appaltatori": appaltatori,
        "subappaltatori": subappaltatori,
        "awareness": awareness
    }

    risk, level = calculate_risk(data)
    actions = suggest_actions(level)

    st.subheader("📊 Risultato")
    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    st.subheader("🎯 Azioni")
    for a in actions:
        st.write(f"- {a}")
