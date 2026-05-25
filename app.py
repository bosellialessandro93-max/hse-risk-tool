import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO


# ==========================================
# CALCOLO RISCHIO (MIGLIORATO)
# ==========================================
def calculate_risk(data):

    weights = {"lieve": 1, "media": 2, "grave": 4, "critica": 6}

    severity = sum(weights.get(x.lower(), 0) for x in data["nc_types"])
    ncScore = min(data["num_nc"] * 2 + severity, 100)

    # STOP WORK (evento + cultura separati)
    stop_event = data["stopworks"] * 8
    stop_culture = min(data["stopworks"] * 2, 10)
    stopScore = min(stop_event, 100) - stop_culture

    # CRITICITÀ PESATE
    critScore = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 15 +
        data["crit_ontime"] * 5,
        100
    )

    # ISPEZIONI DINAMICHE
    inspectionScore = max(10, 50 - data["inspections"] * 2 + data["num_nc"] * 1.5)

    # COMPLESSITÀ
    complexity = data["app"] + data["sub"] * 1.5
    complexityScore = min(complexity * 5, 100)

    awareness_bonus = min(data["awareness"] * 0.1, 15)

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

    return risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore


# ==========================================
# AZIONI INTELLIGENTI (LOGICA AVANZATA)
# ==========================================
def generate_actions(data):

    actions = []

    nc = data["num_nc"]
    stop = data["stopworks"]
    inspections = data["inspections"]

    crit_open = data["crit_open"]
    crit_late = data["crit_late"]
    crit_ontime = data["crit_ontime"]

    awareness = data["awareness"]

    total_crit = crit_open + crit_late + crit_ontime
    late_ratio = (crit_late / total_crit) if total_crit > 0 else 0

    # =====================
    # CORRELAZIONE PRINCIPALE
    # =====================
    if nc > 5 and stop > 3:
        actions.append(("MUST HAVE", "🔴 Eventi elevati → aumento probabilità infortunio"))
        actions.append(("MUST HAVE", "Aumentare ispezioni e sensibilizzazioni"))

    elif inspections > 20 and nc < 3 and stop < 2:
        actions.append(("MUST HAVE", "🟢 Sistema sotto controllo"))
        actions.append(("NICE TO HAVE", "Mantenere attuale livello di monitoraggio"))

    # =====================
    # CRITICITÀ PERFORMANCE
    # =====================
    if late_ratio > 0.3:
        actions.append(("MUST HAVE", "🔴 Alta % criticità risolte in ritardo"))
        actions.append(("MUST HAVE", "Rafforzare processo chiusura azioni"))

    if crit_open > 10:
        actions.append(("MUST HAVE", "🔴 Troppe criticità aperte"))
        actions.append(("MUST HAVE", "Piano straordinario di chiusura"))

    if crit_ontime > crit_late:
        actions.append(("NICE TO HAVE", "🟢 Buona capacità risoluzione criticità"))
    
    # =====================
    # CONTROLLO
    # =====================
    if inspections < 5 and nc > 3:
        actions.append(("MUST HAVE", "🔴 Controllo insufficiente"))
        actions.append(("MUST HAVE", "Aumentare presenza HSE"))

    if inspections > 25 and nc > 8:
        actions.append(("NICE TO HAVE", "🟡 Controlli numerosi ma poco efficaci"))

    # =====================
    # AWARENESS
    # =====================
    if awareness < inspections:
        actions.append(("NICE TO HAVE", "Sensibilizzazione da aumentare"))

    if awareness > 50 and nc < 3:
        actions.append(("IDEA", "Alta awareness sta riducendo eventi"))

    # =====================
    # GARANTIRE MINIMO 3 OUTPUT
    # =====================
    if len(actions) < 3:
        if nc == 0:
            actions.append(("IDEA", "Nessuna NC: verificare reale capacità di rilevazione"))
        if inspections > 0:
            actions.append(("IDEA", "Continuare monitoraggio operativo"))
        actions.append(("IDEA", "Rafforzare cultura preventiva"))

    return actions


# ==========================================
# PPT (MIGLIORATO)
# ==========================================
def generate_ppt(nome, risk, level, df, actions, data):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risultato"
    slide.placeholders[1].text = f"Risk Index: {round(risk)} - {level}"

    # KPI CRITICITÀ
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Gestione criticità"
    slide.placeholders[1].text = (
        f"Aperte: {data['crit_open']}\n"
        f"Risolte in tempo: {data['crit_ontime']}\n"
        f"Risolte in ritardo: {data['crit_late']}"
    )

    # DRIVER RISCHIO
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Driver rischio"

    table = slide.shapes.add_table(len(df) + 1, 2, Inches(1), Inches(1.5), Inches(6), Inches(3)).table
    table.cell(0, 0).text = "Componente"
    table.cell(0, 1).text = "Valore"

    for i, row in df.iterrows():
        table.cell(i + 1, 0).text = str(row["Componente"])
        table.cell(i + 1, 1).text = str(round(row["Valore"]))

    # AZIONI
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni suggerite"

    text = ""
    for level_tag, action in actions:
        text += f"{level_tag} - {action}\n"

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

# NUOVE CRITICITÀ
crit_open = st.number_input("Criticità aperte", 0, 100, 0)
crit_ontime = st.number_input("Criticità risolte in tempo", 0, 100, 0)
crit_late = st.number_input("Criticità risolte in ritardo", 0, 100, 0)

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

    # LETTURA EXCEL MIGLIORATA
    if file is not None:
        df_excel = pd.read_excel(file)

        if "Stato" in df_excel.columns and "Scadenza" in df_excel.columns:

            crit_open = len(df_excel[df_excel["Stato"].str.lower() == "aperto"])

            df_closed = df_excel[df_excel["Stato"].str.lower() == "chiuso"]

            if "Data chiusura" in df_excel.columns:
                df_closed["in_tempo"] = df_closed["Data chiusura"] <= df_closed["Scadenza"]

                crit_ontime = len(df_closed[df_closed["in_tempo"] == True])
                crit_late = len(df_closed[df_closed["in_tempo"] == False])

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
        "fase": fase
    }

    risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore = calculate_risk(data)

    actions = generate_actions(data)

    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità"],
        "Valore": [ncScore, stopScore, critScore, inspectionScore, complexityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    st.subheader("🎯 Azioni suggerite")
    for lvl, a in actions:
        st.write(f"{lvl} - {a}")

    prs = generate_ppt(nome, risk, level, df, actions, data)
    buffer = BytesIO()
    prs.save(buffer)

    st.download_button("📥 Scarica PPT", buffer.getvalue(), "HSE_Report.pptx")
