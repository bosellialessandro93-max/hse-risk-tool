import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

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

    stop_risk = min(data["stopworks"] * 10, 100)
    stop_bonus = min(data["stopworks"] * 2, 10)
    stopWorkScore = stop_risk - stop_bonus

    criticalityScore = min(data["criticalities"] * 10, 100)
    inspectionScore = max(10, 50 - data["inspections"] * 2)

    complexity = data["appaltatori"] + (data["subappaltatori"] * 1.5)
    complexityScore = min(complexity * 5, 100)

    awarenessMitigation = min(data["awareness"] * 0.1, 15)

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

    return riskScore, level, ncScore, stopWorkScore, criticalityScore


# ==========================================
# AZIONI INTELLIGENTI
# ==========================================

def generate_actions(data, level, ncScore, stopScore, critScore):

    actions = []

    # NC
    if ncScore > 40:
        actions.append("Analizzare le principali Non Conformità e definire azioni correttive")
    
    # Stop Work
    if data["stopworks"] > 2:
        actions.append("Verificare l'efficacia delle Stop Work e la loro conversione in azioni strutturate")

    # Criticità
    if data["criticalities"] > 5:
        actions.append("Ridurre backlog criticità aperte con piano di chiusura prioritizzato")

    # Fase commissioning
    if data["fase"] == "commissioning":
        actions.append("Rafforzare controlli operativi in fase di commissioning")

    # Complessità
    if data["appaltatori"] + data["subappaltatori"] > 5:
        actions.append("Rafforzare coordinamento appaltatori e subappaltatori")

    # livello alto
    if level in ["Alto", "Critico"]:
        actions.append("Attivare audit HSE straordinario")

    return actions


# ==========================================
# GENERAZIONE PPT
# ==========================================

def generate_ppt(nome, risk, level, df, actions):

    prs = Presentation()

    # slide titolo
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}"

    # slide risultato
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risultato"
    slide.placeholders[1].text = f"Risk Index: {round(risk)}\nLivello: {level}"

    # slide dashboard
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Componenti rischio"

    rows = len(df) + 1
    table = slide.shapes.add_table(rows, 2, Inches(1), Inches(1.5), Inches(6), Inches(3)).table

    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"]))

    # slide azioni
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni suggerite"

    content = slide.placeholders[1]
    content.text = "\n".join(actions)

    return prs


# ==========================================
# UI
# ==========================================

st.title("🦺 HSE Risk Tool PRO")

nome_cantiere = st.text_input("Nome cantiere")

fase = st.selectbox(
    "Fase cantiere",
    ["cantierizzazione", "costruzione", "commissioning", "punch list"]
)

inspections = st.number_input("Numero ispezioni", 0, 1000, 10)
num_nc = st.number_input("Numero NC", 0, 100, 0)

nc_data = st.text_area("NC (formato: tema,tipo | es: quota,grave | elettrico,media)")

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

    if nc_data:
        rows = nc_data.split("|")
        for r in rows:
            parts = r.split(",")
            if len(parts) == 2:
                nc_types.append(parts[1].strip())

    if uploaded_file:
        df_excel = pd.read_excel(uploaded_file)
        if "Stato" in df_excel.columns:
            criticalities = len(df_excel[df_excel["Stato"].str.lower() == "aperto"])

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

    risk, level, ncScore, stopScore, critScore = calculate_risk(data)

    actions = generate_actions(data, level, ncScore, stopScore, critScore)

    st.subheader("📊 Risultato")
    st.metric("Risk Index", round(risk))
    st.metric("Livello rischio", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità"],
        "Valore": [ncScore, stopScore, critScore]
    })

    st.subheader("📈 Dashboard")
    st.bar_chart(df.set_index("Componente"))

    st.subheader("🎯 Azioni suggerite")
    for a in actions:
        st.write("-", a)

    # PPT
    prs = generate_ppt(nome_cantiere, risk, level, df, actions)

    buffer = BytesIO()
    prs.save(buffer)

    st.download_button(
        "📥 Scarica report PPT",
        buffer.getvalue(),
        "HSE_Risk_Report.pptx"
    )
``
