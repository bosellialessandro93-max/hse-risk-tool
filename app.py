import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

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

    return riskScore, level, ncScore, stopWorkScore, criticalityScore, inspectionScore, complexityScore


# ==========================================
# PPT REPORT
# ==========================================

def generate_ppt(nome, risk, level, df):

    prs = Presentation()

    # Slide 1 - titolo
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "HSE Risk Report"
    subtitle.text = f"Cantiere: {nome}"

    # Slide 2 - risultato
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    slide.shapes.title.text = "Risultato"
    content = slide.placeholders[1]

    content.text = f"Risk Index: {round(risk)}\nLivello: {level}"

    # Slide 3 - dashboard
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)

    slide.shapes.title.text = "Componenti rischio"

    left = Inches(1)
    top = Inches(1.5)

    rows = len(df) + 1
    cols = 2

    table = slide.shapes.add_table(rows, cols, left, top, Inches(6), Inches(3)).table

    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"]))

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

    if uploaded_file is not None:
        df_excel = pd.read_excel(uploaded_file)

        if "Stato" in df_excel.columns:
            crit_open = df_excel[df_excel["Stato"].str.lower() == "aperto"]
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

    st.subheader("📊 Risultato")
    st.metric("Risk Index", round(risk))
    st.metric("Livello rischio", level)

    st.subheader("📈 Dashboard")

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità"],
        "Valore": [ncScore, stopWorkScore, criticalityScore, inspectionScore, complexityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    # ==========================================
    # DOWNLOAD PPT
    # ==========================================

    prs = generate_ppt(nome_cantiere, risk, level, df)
    
    ppt_bytes = None
    from io import BytesIO
    buf = BytesIO()
    prs.save(buf)
    ppt_bytes = buf.getvalue()

    st.download_button(
        label="📥 Scarica Report PPT",
        data=ppt_bytes,
        file_name="HSE_Risk_Report.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
