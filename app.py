import streamlit as st
import pandas as pd

# ==========================================
# CALCOLO RISCHIO
# ==========================================

def calculate_risk(data):

    weights = {
        "lieve": 1,
        "media": 2,
        "grave": 4,
        "critica": 6
    }

    severity = sum(weights.get(x.strip().lower(), 0) for x in data["nc_types"])
    ncScore = min((data["num_nc"] * 2 + severity), 100)

    # Stop Work bilanciato (rischio + cultura)
    stop_risk = min(data["stopworks"] * 10, 100)
    stop_bonus = min(data["stopworks"] * 2, 10)
    stopWorkScore = stop_risk - stop_bonus

    criticalityScore = min(data["criticalities"] * 10, 100)
    inspectionScore = max(10, 50 - data["inspections"] * 2)

    complexity = data["appaltatori"] + (data["subappaltatori"] * 1.5)
    complexityScore = min(complexity * 5, 100)

    awarenessMitigation = min(data["awareness"] * 0.1, 15)

    # fase cantiere
    fase_weights = {
        "cantierizzazione": 20,
        "costruzione": 40,
        "commissioning": 60,
        "punch list": 30
    }

    faseScore = fase_weights.get(data["fase"], 30)

    riskScore = (
        ncScore * 0.25 +
        stopWorkScore * 0.20 +
        criticalityScore * 0.20 +
        inspectionScore * 0.10 +
        complexityScore * 0.15 +
        faseScore * 0.10
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

    return riskScore, level, ncScore, stopWorkScore, criticalityScore, inspectionScore, complexityScore


# ==========================================
# AZIONI
# ==========================================

def suggest_actions(level):
    actions_map = {
        "Basso": ["Monitoraggio continuo", "Sensibilizzazioni"],
        "Medio basso": ["Verifica NC", "Formazione mirata"],
        "Medio": ["Analisi cause", "Aumentare ispezioni", "Briefing appaltatori"],
        "Alto": ["Audit HSE", "Piano correttivo", "Revisione subappalti"],
        "Critico": ["Escalation immediata", "Stop attività", "Indagine eventi"]
    }
    return actions_map[level]


# ==========================================
# UI
# ==========================================

st.title("🦺 HSE Risk Tool PRO")

sito = st.selectbox("Seleziona sito", ["La Spezia", "Altro sito", "Test"])

fase = st.selectbox(
    "Fase cantiere",
    ["cantierizzazione", "costruzione", "commissioning", "punch list"]
)

inspections = st.number_input("Numero ispezioni", 0, 1000, 10)
num_nc = st.number_input("Numero NC", 0, 100, 0)

nc_data = st.text_area(
    "NC (formato: tema,tipo | es: quota,grave | elettrico,media)"
)

stopworks = st.number_input("Stop Work", 0, 50, 0)
criticalities = st.number_input("Criticità aperte", 0, 100, 0)

appaltatori = st.number_input("Appaltatori", 0, 50, 0)
subappaltatori = st.number_input("Subappaltatori", 0, 50, 0)

awareness = st.number_input("Persone sensibilizzate", 0, 1000, 0)

uploaded_file = st.file_uploader("Carica file Excel criticità", type=["xlsx"])


# ==========================================
# CALCOLO
# ==========================================

if st.button("Calcola rischio"):

    nc_types = []
    nc_themes = []

    if nc_data:
        rows = nc_data.split("|")
        for r in rows:
            parts = r.split(",")
            if len(parts) == 2:
                nc_themes.append(parts[0].strip())
                nc_types.append(parts[1].strip())

    # Excel parsing
    if uploaded_file is not None:
        df_excel = pd.read_excel(uploaded_file)

        if "Stato" in df_excel.columns:
            crit_open = df_excel[
                df_excel["Stato"].str.lower() == "aperto"
            ]
            criticalities = len(crit_open)

    data = {
        "inspections": inspections,
        "num_nc": num_nc,
        "nc_types": nc_types,
        "stopworks": stopworks,
        "criticalities": criticalities,
        "appaltatori": appaltatori,
        "subappaltatori": subappaltatori,
        "awareness": awareness,
        "fase": fase
    }

    risk, level, ncScore, stopWorkScore, criticalityScore, inspectionScore, complexityScore = calculate_risk(data)

    actions = suggest_actions(level)

    # ==========================================
    # OUTPUT
    # ==========================================

    st.subheader("📊 Risultato")

    st.metric("Risk Index", round(risk))
    st.metric("Livello rischio", level)

    # ==========================================
    # DASHBOARD
    # ==========================================

    st.subheader("📈 Dashboard")

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità"],
        "Valore": [ncScore, stopWorkScore, criticalityScore, inspectionScore, complexityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    # ==========================================
    # DRIVER
    # ==========================================

    st.subheader("🚨 Driver principali")

    drivers = []

    if stopworks > 2:
        drivers.append("Stop Work elevati")

    if criticalities > 5:
        drivers.append("Criticità aperte elevate")

    if num_nc > 3:
        drivers.append("Numero NC significativo")

    for d in drivers:
        st.write("-", d)

    # ==========================================
    # AZIONI
    # ==========================================

    st.subheader("🎯 Azioni suggerite")

    for a in actions:
        st.write("-", a)

    # ==========================================
    # SPIEGAZIONE
    # ==========================================

    st.subheader("🧠 Spiegazione")

    explanation = f"Il rischio è {level}. "

    if stopworks > 2:
        explanation += "Stop Work indicano attività preventiva. "

    if criticalities > 5:
        explanation += "Presenza di backlog criticità. "

    if num_nc > 0:
        explanation += "Le NC contribuiscono al rischio. "

    if inspections > 20:
        explanation += "Buon controllo operativo. "

    st.write(explanation)

    # ==========================================
    # WARNING
    # ==========================================

    if stopworks >= num_nc and stopworks > 0:
        st.warning("⚠️ Molti Stop Work → verificare coerenza con NC")

    if criticalities > 5:
        st.warning("⚠️ Backlog criticità elevato")

    # ==========================================
    # STORE MULTI SITO
    # ==========================================

    st.subheader("🌍 Dati sito")

    df_store = pd.DataFrame({
        "Sito": [sito],
        "Rischio": [round(risk)]
    })

    st.write(df_store)

    # ==========================================
    # INSIGHT AI (rule-based)
    # ==========================================

    st.subheader("🤖 Insight automatici")

    if "quota" in " ".join(nc_themes):
        st.write("- Rafforzare formazione lavori in quota")

    if stopworks > 3 and criticalities > 5:
        st.write("- Verificare conversione Stop Work → azioni reali")

    if criticalities > 5:
        st.write("- Ridurre backlog criticità")
