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

    # NUOVO: attività e interferenze
    activityScore = min(data["activities"] * 5 + (5 if data["activity_type"] == "elettrici" else 3), 100)

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

    # Livello
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
# AZIONI INTELLIGENTI (SUPER ARGOMENTATE)
# ==========================================
def generate_actions(data):

    actions = []

    nc = data["num_nc"]
    stop = data["stopworks"]
    inspections = data["inspections"]
    awareness = data["awareness"]
    activities = data["activities"]
    activity_type = data["activity_type"]

    crit_open = data["crit_open"]
    crit_late = data["crit_late"]
    crit_ontime = data["crit_ontime"]

    total_crit = crit_open + crit_late + crit_ontime
    late_ratio = (crit_late / total_crit) if total_crit > 0 else 0

    # ======================================
    # CORRELAZIONE ISPEZIONI / NC
    # ======================================
    if inspections > 20 and nc < 3:
        actions.append(("NICE TO HAVE",
            f"🟢 Il numero elevato di ispezioni ({inspections}) rispetto alle NC ({nc}) "
            "indica un sistema di controllo efficace che sta prevenendo eventi."))

    if inspections < 5 and nc > 3:
        actions.append(("MUST HAVE",
            f"🔴 Sono state registrate {nc} NC a fronte di sole {inspections} ispezioni. "
            "Il basso livello di controllo aumenta il rischio di eventi non intercettati."))

    if nc > inspections:
        actions.append(("MUST HAVE",
            f"🔴 Il numero di NC ({nc}) supera quello delle ispezioni ({inspections}). "
            "Questo indica un sistema reattivo e non preventivo."))

    # ======================================
    # AWARENESS
    # ======================================
    if awareness > 30:
        actions.append(("NICE TO HAVE",
            f"🟢 Sono state effettuate {awareness} sensibilizzazioni: "
            "questo contribuisce alla riduzione del rischio migliorando i comportamenti."))

    if awareness < inspections:
        actions.append(("MUST HAVE",
            f"🟡 Le sensibilizzazioni ({awareness}) sono inferiori alle ispezioni ({inspections}). "
            "È necessario rafforzare la cultura della sicurezza."))

    # ======================================
    # CRITICITÀ
    # ======================================
    if late_ratio > 0.3:
        actions.append(("MUST HAVE",
            f"🔴 Il {round(late_ratio*100)}% delle criticità viene chiuso in ritardo. "
            "Questo aumenta la permanenza dell’esposizione al rischio."))

    if crit_open > 10:
        actions.append(("MUST HAVE",
            f"🔴 Sono presenti {crit_open} criticità aperte. "
            "È necessaria una riduzione immediata dello stock aperto."))

    if crit_ontime > crit_late:
        actions.append(("NICE TO HAVE",
            "🟢 Buona capacità di gestione delle criticità, la maggior parte viene chiusa nei tempi."))

    # ======================================
    # ATTIVITÀ E INTERFERENZE
    # ======================================
    if activities >= 3:
        actions.append(("MUST HAVE",
            f"🔴 Sono presenti {activities} attività in contemporanea ({activity_type}). "
            "Questo aumenta il rischio interferenziale."))

        if activity_type == "elettrici":
            actions.append(("MUST HAVE",
                "⚡ Attività elettriche in corso: si raccomanda sensibilizzazione su rischio elettrico "
                "e verifica PES/PAV e procedure LOTO."))

        if activity_type == "scavi":
            actions.append(("MUST HAVE",
                "🚧 Attività di scavo in corso: rafforzare sensibilizzazione su seppellimento, "
                "sottoservizi e stabilità fronti."))

    # ======================================
    # STOP WORK
    # ======================================
    if stop > 3:
        actions.append(("MUST HAVE",
            f"🔴 Elevato numero di Stop Work ({stop}). "
            "Indica criticità operative ricorrenti."))

    # ======================================
    # GARANTIRE MINIMO OUTPUT
    # ======================================
    if len(actions) < 3:
        actions.append(("IDEA",
            "👉 Mantenere monitoraggio continuo e rafforzare approccio preventivo."))

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

    # NUOVO: attività
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Attività in corso"
    slide.placeholders[1].text = (
        f"Tipologia: {data['activity_type']}\n"
        f"Numero attività: {data['activities']}"
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

crit_open = st.number_input("Criticità aperte", 0, 100, 0)
crit_ontime = st.number_input("Criticità risolte in tempo", 0, 100, 0)
crit_late = st.number_input("Criticità risolte in ritardo", 0, 100, 0)

app = st.number_input("Appaltatori", 0, 50, 0)
sub = st.number_input("Subappaltatori", 0, 50, 0)

awareness = st.number_input("Sensibilizzazioni", 0, 1000, 0)

# ✅ NUOVE SEZIONI
activity_type = st.selectbox("Tipologia attività", ["scavi", "elettrici", "altro"])
activities = st.number_input("Numero attività in corso", 0, 20, 0)

file = st.file_uploader("Excel criticità", type=["xlsx"])


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

    if file is not None:
        df_excel = pd.read_excel(file)

        if "Stato" in df_excel.columns:
            crit_open = len(df_excel[df_excel["Stato"].str.lower() == "aperto"])

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

    risk, level, ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore = calculate_risk(data)

    actions = generate_actions(data)

    st.metric("Risk Index", round(risk))
    st.metric("Livello", level)

    df = pd.DataFrame({
        "Componente": ["NC", "Stop Work", "Criticità", "Ispezioni", "Complessità", "Attività"],
        "Valore": [ncScore, stopScore, critScore, inspectionScore, complexityScore, activityScore]
    })

    st.bar_chart(df.set_index("Componente"))

    st.subheader("🎯 Azioni suggerite")
    for lvl, a in actions:
        st.write(f"{lvl} - {a}")

    prs = generate_ppt(nome, risk, level, df, actions, data)
    buffer = BytesIO()
    prs.save(buffer)

    st.download_button("📥 Scarica PPT", buffer.getvalue(), "HSE_Report.pptx")
``
