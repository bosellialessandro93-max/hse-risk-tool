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

    stop_risk = min(data["stopworks"] * 10, 100)
    stop_bonus = min(data["stopworks"] * 2, 10)
    stopScore = stop_risk - stop_bonus

    critScore = min(data["criticalities"] * 10, 100)

    inspectionScore = max(10, 50 - data["inspections"] * 2)

    complexity = data["app"] + data["sub"] * 1.5
    complexityScore = min(complexity * 5, 100)

    awareness = min(data["awareness"] * 0.1, 15)

    fase_weights = {
        "cantierizzazione": 20,
        "costruzione": 40,
        "commissioning": 60,
        "punch list": 30
    }

    faseScore = fase_weights.get(data["fase"], 30)

    risk = (
        ncScore * 0.25 +
        stopScore * 0.20 +
        critScore * 0.20 +
        inspectionScore * 0.10 +
        complexityScore * 0.15 +
        faseScore * 0.10
    ) - awareness

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

    return risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore


# ==========================================
# AZIONI INTELLIGENTI (STABILI)
# ==========================================
def generate_actions(data, level, nc_types, nc_themes, stopScore, critScore):

    actions = []

    themes_text = " ".join(nc_themes).lower()

    # CRITICITÀ
    if data["criticalities"] > 5:
        actions.append("🔴 HIGH → backlog criticità elevato: attivare piano di chiusura")

        if data["stopworks"] > 0:
            actions.append("🔴 Stop Work NON convertite in azioni efficaci")

    # NC TEMATICHE
    if "elettrico" in themes_text:
        actions.append("🟡 Rafforzare controlli rischio elettrico")

    if "quota" in themes_text:
        actions.append("🟡 Migliorare sicurezza lavori in quota")

    if "dpi" in themes_text:
        actions.append("🟡 Migliorare comportamento utilizzo DPI")

    # STOP WORK
    if data["stopworks"] > data["num_nc"]:
        actions.append("🟡 Disallineamento Stop Work / NC")

    if data["stopworks"] > 2:
        actions.append("🟢 Cultura preventiva attiva")

    # FASE
    if data["fase"] == "commissioning":
        actions.append("🔴 Fase commissioning → presidio HSE rafforzato")

    # ISPEZIONI
    if data["inspections"] < 5:
        actions.append("🟡 Incrementare attività ispettiva")

    if data["inspections"] > 30:
        actions.append("🟢 Buon livello di controllo")

    return actions


# ==========================================
# PPT
# ==========================================
def generate_ppt(nome, risk, level, df, actions):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risultato"
    slide.placeholders[1].text = f"Risk Index: {round(risk)} - {level}"

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Driver rischio"

    table = slide.shapes.add_table(len(df) + 1, 2, Inches(1), Inches(1.5), Inches(6), Inches(3)).table

    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"]))

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni suggerite"
    slide.placeholders[1].text = "\n".join(actions)

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
crit = st.number_input("Criticità", 0, 100, 0)

app = st.number_input("Appaltatori", 0, 50, 0)
sub = st.number_input("Subappaltatori", 0, 50, 0)

awareness = st.number_input("Sensibilizzazioni", 0, 1000, 0)

file = st.file_uploader("Excel criticità", type=["xlsx"])


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

    if file is not None:
        df_excel = pd.read_excel(file)
        if "Stato" in df_excel.columns:
            crit = len(df_excel[df_excel["Stato"].str.lower() == "aperto"])

    data = {
        "inspections": inspections,
        "num_nc": num_nc,
        "nc_types": nc_types,
        "stopworks": stop,
        "criticalities": crit,
        "app": app,
        "sub": sub,
        "awareness": awareness,
        "fase": fase
    }

    risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore = calculate_risk(data)

    actions = generate_actions(data, level, nc_types, nc_themes, stopScore, critScore)

    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità"],
        "Valore": [ncScore, stopScore, critScore, inspectionScore, complexityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    st.subheader("🎯 Azioni suggerite")
    for a in actions:
        st.write(a)

    prs = generate_ppt(nome, risk, level, df, actions)
    buffer = BytesIO()
    prs.save(buffer)

    st.download_button("📥 Scarica PPT", buffer.getvalue(), "HSE_Report.pptx")
