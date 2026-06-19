import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from datetime import datetime


# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="HSE Risk Platform",
    page_icon="🦺",
    layout="wide"
)


# =========================
# DIZIONARI RISCHIO
# =========================
NC_SEVERITY_WEIGHTS = {
    "lieve": 2,
    "media": 5,
    "grave": 10,
    "critica": 18
}

NC_THEME_WEIGHTS = {
    "elettrico": 16,
    "scavi": 16,
    "quota": 14,
    "sollevamento": 14,
    "spazi confinati": 18,
    "incendio": 14,
    "viabilità": 10,
    "dpi": 7,
    "ordine e pulizia": 6,
    "documentale": 5,
    "ambientale": 8,
    "altro": 6
}

ACTIVITY_WEIGHTS = {
    "elettrici": 18,
    "scavi": 18,
    "lavori in quota": 16,
    "sollevamenti": 16,
    "spazi confinati": 20,
    "hot work": 15,
    "movimentazione mezzi": 12,
    "civili": 8,
    "meccanici": 10,
    "altro": 6
}

AWARENESS_BONUS = {
    "rischio elettrico": 8,
    "scavi e sottoservizi": 8,
    "lavori in quota": 7,
    "sollevamenti": 7,
    "spazi confinati": 9,
    "hot work / incendio": 7,
    "viabilità e mezzi": 6,
    "dpi": 4,
    "ordine e pulizia": 4,
    "toolbox generale": 3,
    "altro": 2
}

LINKS_ACTIVITY_NC = {
    "elettrici": ["elettrico"],
    "scavi": ["scavi"],
    "lavori in quota": ["quota"],
    "sollevamenti": ["sollevamento"],
    "spazi confinati": ["spazi confinati"],
    "hot work": ["incendio"],
    "movimentazione mezzi": ["viabilità"]
}

LINKS_AWARENESS_ACTIVITY = {
    "rischio elettrico": ["elettrici"],
    "scavi e sottoservizi": ["scavi"],
    "lavori in quota": ["lavori in quota"],
    "sollevamenti": ["sollevamenti"],
    "spazi confinati": ["spazi confinati"],
    "hot work / incendio": ["hot work"],
    "viabilità e mezzi": ["movimentazione mezzi"],
    "dpi": ["civili", "meccanici", "elettrici", "scavi"],
    "ordine e pulizia": ["civili", "meccanici"]
}


# =========================
# UTILITY
# =========================
def risk_level(risk):
    if risk <= 20:
        return "Basso"
    elif risk <= 40:
        return "Medio basso"
    elif risk <= 60:
        return "Medio"
    elif risk <= 80:
        return "Alto"
    return "Critico"


def normalize_list(values):
    return [v.strip().lower() for v in values if v.strip()]


def calculate_risk(data):
    details = []
    explanations = []

    inspections = data["inspections"]
    num_nc = data["num_nc"]
    nc_items = data["nc_items"]
    activities = data["activities"]
    awareness_types = data["awareness_types"]
    awareness_count = data["awareness_count"]

    # -------------------------
    # NC SCORE
    # -------------------------
    nc_score = 0

    for item in nc_items:
        theme = item["theme"]
        severity = item["severity"]

        theme_score = NC_THEME_WEIGHTS.get(theme, 6)
        severity_score = NC_SEVERITY_WEIGHTS.get(severity, 5)
        item_score = theme_score + severity_score

        nc_score += item_score

        explanations.append(
            f"NC su '{theme}' con gravità '{severity}': +{item_score} punti."
        )

    nc_score = min(nc_score, 100)

    # -------------------------
    # DENSITÀ NC
    # -------------------------
    nc_density = num_nc / max(inspections, 1)

    if nc_density >= 1:
        density_penalty = 25
        explanations.append("Densità NC molto alta rispetto alle ispezioni: +25 punti.")
    elif nc_density >= 0.5:
        density_penalty = 15
        explanations.append("Densità NC significativa rispetto alle ispezioni: +15 punti.")
    elif nc_density >= 0.25:
        density_penalty = 8
        explanations.append("Densità NC moderata rispetto alle ispezioni: +8 punti.")
    else:
        density_penalty = 0

    # -------------------------
    # CONTROLLO
    # -------------------------
    if inspections == 0 and num_nc > 0:
        control_score = 30
        explanations.append("NC presenti senza ispezioni registrate: +30 punti.")
    elif inspections < 3 and num_nc >= 3:
        control_score = 25
        explanations.append("Poche ispezioni rispetto alle NC rilevate: +25 punti.")
    elif inspections >= 20 and num_nc <= 2:
        control_score = -10
        explanations.append("Molte ispezioni con poche NC: -10 punti.")
    elif inspections >= 10 and num_nc <= 2:
        control_score = -5
        explanations.append("Buon presidio ispettivo con poche NC: -5 punti.")
    else:
        control_score = 0

    # -------------------------
    # STOP WORK
    # -------------------------
    stop_score = min(data["stopworks"] * 10, 40)

    if data["stopworks"] > 0:
        explanations.append(
            f"{data['stopworks']} Stop Work registrati: +{stop_score} punti."
        )

    # -------------------------
    # CRITICITÀ
    # -------------------------
    crit_score = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 10 +
        data["crit_ontime"] * 3,
        100
    )

    if data["crit_open"] > 0:
        explanations.append(
            f"{data['crit_open']} criticità aperte: incremento rischio."
        )

    if data["crit_late"] > 0:
        explanations.append(
            f"{data['crit_late']} criticità risolte in ritardo: peggiora la capacità di chiusura."
        )

    # -------------------------
    # COMPLESSITÀ ORGANIZZATIVA
    # -------------------------
    complexity_score = min(
        data["app"] * 5 +
        data["sub"] * 8,
        70
    )

    if data["sub"] >= 5:
        complexity_score += 15
        explanations.append("Numero elevato di subappaltatori: +15 punti.")

    complexity_score = min(complexity_score, 100)

    # -------------------------
    # ATTIVITÀ
    # -------------------------
    activity_score = 0

    for act in activities:
        score = ACTIVITY_WEIGHTS.get(act, 6)
        activity_score += score
        explanations.append(f"Attività '{act}': +{score} punti.")

    if len(activities) >= 3:
        activity_score += 20
        explanations.append("Tre o più attività contemporanee: +20 punti per interferenza.")

    if len(activities) >= 5:
        activity_score += 15
        explanations.append("Cinque o più attività contemporanee: +15 punti aggiuntivi.")

    activity_score = min(activity_score, 100)

    # -------------------------
    # CORRELAZIONE ATTIVITÀ / NC
    # -------------------------
    correlation_penalty = 0

    nc_themes = [item["theme"] for item in nc_items]

    for act in activities:
        expected_nc = LINKS_ACTIVITY_NC.get(act, [])

        for theme in expected_nc:
            if theme in nc_themes:
                correlation_penalty += 15
                explanations.append(
                    f"Correlazione critica: attività '{act}' con NC su '{theme}': +15 punti."
                )

    correlation_penalty = min(correlation_penalty, 45)

    # -------------------------
    # FASE
    # -------------------------
    fase_score = 0

    if data["fase"] == "commissioning":
        fase_score += 18
        explanations.append("Fase commissioning: +18 punti per presenza energie/prove/interferenze.")
    elif data["fase"] == "punch list":
        fase_score += 10
        explanations.append("Fase punch list: +10 punti per attività residue e discontinuità operative.")
    elif data["fase"] == "cantierizzazione":
        fase_score += 8
        explanations.append("Fase cantierizzazione: +8 punti per allestimenti e avvio attività.")

    # -------------------------
    # BONUS SENSIBILIZZAZIONI
    # Non correlate alle ispezioni.
    # Bonus basato su quantità, qualità e coerenza con attività/NC.
    # -------------------------
    awareness_bonus = 0

    for aw in awareness_types:
        bonus = AWARENESS_BONUS.get(aw, 2)
        awareness_bonus += bonus
        explanations.append(f"Sensibilizzazione '{aw}': -{bonus} punti.")

    if awareness_count >= 5:
        awareness_bonus += 5
        explanations.append("Numero adeguato di sensibilizzazioni: -5 punti.")

    if awareness_count >= 10:
        awareness_bonus += 5
        explanations.append("Buona continuità nelle sensibilizzazioni: -5 punti.")

    # Bonus coerenza attività/sensibilizzazione
    for aw in awareness_types:
        covered_activities = LINKS_AWARENESS_ACTIVITY.get(aw, [])

        for act in activities:
            if act in covered_activities:
                awareness_bonus += 5
                explanations.append(
                    f"Sensibilizzazione coerente con attività '{act}': -5 punti."
                )

    # Bonus coerenza NC/sensibilizzazione
    for aw in awareness_types:
        if aw == "rischio elettrico" and "elettrico" in nc_themes:
            awareness_bonus += 4
            explanations.append("Sensibilizzazione mirata su NC elettriche: -4 punti.")

        if aw == "scavi e sottoservizi" and "scavi" in nc_themes:
            awareness_bonus += 4
            explanations.append("Sensibilizzazione mirata su NC scavi: -4 punti.")

        if aw == "lavori in quota" and "quota" in nc_themes:
            awareness_bonus += 4
            explanations.append("Sensibilizzazione mirata su NC lavori in quota: -4 punti.")

        if aw == "spazi confinati" and "spazi confinati" in nc_themes:
            awareness_bonus += 5
            explanations.append("Sensibilizzazione mirata su spazi confinati: -5 punti.")

    awareness_bonus = min(awareness_bonus, 35)

    # -------------------------
    # CALCOLO FINALE
    # -------------------------
    risk = (
        nc_score * 0.20 +
        crit_score * 0.15 +
        activity_score * 0.20 +
        complexity_score * 0.10 +
        stop_score * 0.10 +
        density_penalty +
        control_score +
        correlation_penalty +
        fase_score -
        awareness_bonus
    )

    risk = max(0, min(100, risk))
    level = risk_level(risk)

    details.append(("NC", nc_score))
    details.append(("Criticità", crit_score))
    details.append(("Attività", activity_score))
    details.append(("Complessità", complexity_score))
    details.append(("Stop Work", stop_score))
    details.append(("Densità NC", density_penalty))
    details.append(("Controllo", control_score))
    details.append(("Correlazioni critiche", correlation_penalty))
    details.append(("Fase", fase_score))
    details.append(("Bonus sensibilizzazioni", -awareness_bonus))

    return risk, level, details, explanations


def generate_actions(data, risk, level):
    actions = []

    nc_themes = [item["theme"] for item in data["nc_items"]]
    severe_nc = [
        item for item in data["nc_items"]
        if item["severity"] in ["grave", "critica"]
    ]

    if risk >= 80:
        actions.append((
            "MUST HAVE",
            "🔴 Rischio critico: convocare coordinamento operativo immediato e validare piano azioni prima della prosecuzione."
        ))

    if severe_nc:
        actions.append((
            "MUST HAVE",
            f"🔴 Presenti {len(severe_nc)} NC gravi/critiche. Definire responsabile, scadenza e verifica chiusura."
        ))

    for act in data["activities"]:
        if act == "elettrici":
            actions.append((
                "MUST HAVE",
                "⚡ Attività elettriche: verificare sezionamenti, autorizzazioni, messa in sicurezza, DPI e competenze PES/PAV/PEI."
            ))

        if act == "scavi":
            actions.append((
                "MUST HAVE",
                "🚧 Scavi: verificare sottoservizi, stabilità fronti, accessi, delimitazioni e presenza di acqua/materiale instabile."
            ))

        if act == "lavori in quota":
            actions.append((
                "MUST HAVE",
                "🪜 Lavori in quota: verificare parapetti, ancoraggi, PLE, imbracature, accessi e piano di emergenza."
            ))

        if act == "sollevamenti":
            actions.append((
                "MUST HAVE",
                "🏗️ Sollevamenti: verificare piano di sollevamento, portate, accessori, interferenze e area segregata."
            ))

        if act == "spazi confinati":
            actions.append((
                "MUST HAVE",
                "☠️ Spazi confinati: autorizzazione, monitoraggio atmosfera, recupero emergenza, presidio esterno e permesso di lavoro."
            ))

        if act == "hot work":
            actions.append((
                "MUST HAVE",
                "🔥 Hot work: verificare permesso, estintori, rimozione combustibili, fire watch e controllo post-attività."
            ))

    if len(data["activities"]) >= 3:
        actions.append((
            "MUST HAVE",
            "🔴 Interferenze elevate: predisporre matrice interferenziale giornaliera e briefing pre-job tra imprese."
        ))

    if "elettrico" in nc_themes:
        actions.append((
            "MUST HAVE",
            "🔴 NC elettriche: bloccare attività non conformi fino a ripristino condizioni sicure."
        ))

    if "scavi" in nc_themes:
        actions.append((
            "MUST HAVE",
            "🔴 NC su scavi: rivalutare rischio sezione di scavo e sottoservizi prima della prosecuzione."
        ))

    if "quota" in nc_themes:
        actions.append((
            "MUST HAVE",
            "🔴 NC lavori in quota: verifica immediata protezioni collettive e sistemi anticaduta."
        ))

    if data["awareness_count"] == 0:
        actions.append((
            "MUST HAVE",
            "🟡 Nessuna sensibilizzazione registrata: pianificare toolbox mirati prima delle attività critiche."
        ))

    missing_awareness = []

    if "elettrici" in data["activities"] and "rischio elettrico" not in data["awareness_types"]:
        missing_awareness.append("rischio elettrico")

    if "scavi" in data["activities"] and "scavi e sottoservizi" not in data["awareness_types"]:
        missing_awareness.append("scavi e sottoservizi")

    if "lavori in quota" in data["activities"] and "lavori in quota" not in data["awareness_types"]:
        missing_awareness.append("lavori in quota")

    if "sollevamenti" in data["activities"] and "sollevamenti" not in data["awareness_types"]:
        missing_awareness.append("sollevamenti")

    if missing_awareness:
        actions.append((
            "MUST HAVE",
            "🟠 Mancano sensibilizzazioni mirate su: " + ", ".join(missing_awareness) + "."
        ))

    if data["crit_open"] > 0:
        actions.append((
            "MUST HAVE",
            f"🟠 Criticità aperte: {data['crit_open']}. Definire priorità, owner e data di chiusura."
        ))

    if data["crit_late"] > 0:
        actions.append((
            "MUST HAVE",
            f"🟠 Criticità chiuse in ritardo: {data['crit_late']}. Analizzare cause del ritardo e capacità di follow-up."
        ))

    if level in ["Basso", "Medio basso"]:
        actions.append((
            "NICE TO HAVE",
            "🟢 Mantenere trend positivo con ispezioni mirate, tracciamento NC e toolbox brevi sulle attività del giorno."
        ))

    if len(actions) < 4:
        actions.append((
            "IDEA",
            "Mantenere presidio operativo e aggiornare la valutazione rischio in caso di nuove attività o nuove imprese."
        ))

    return actions


def generate_ppt(nome, risk, level, driver_df, actions_df, explanations, data):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = (
        f"Cantiere: {nome}\n"
        f"Data: {datetime.now().strftime('%d/%m/%Y')}"
    )

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Sintesi rischio"
    slide.placeholders[1].text = (
        f"Risk Index: {round(risk)} / 100\n"
        f"Livello: {level}\n\n"
        f"Fase: {data['fase']}\n"
        f"Attività: {', '.join(data['activities']) if data['activities'] else 'Nessuna'}\n"
        f"Sensibilizzazioni: {', '.join(data['awareness_types']) if data['awareness_types'] else 'Nessuna'}"
    )

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Driver rischio"

    table = slide.shapes.add_table(
        len(driver_df) + 1,
        2,
        Inches(0.8),
        Inches(1.4),
        Inches(8),
        Inches(4.5)
    ).table

    table.cell(0, 0).text = "Driver"
    table.cell(0, 1).text = "Valore"

    for i, row in driver_df.iterrows():
        table.cell(i + 1, 0).text = str(row["Driver"])
        table.cell(i + 1, 1).text = str(round(row["Valore"], 1))

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Azioni prioritarie"

    text = ""
    for _, row in actions_df.head(8).iterrows():
        text += f"{row['Priorità']} - {row['Azione']}\n\n"

    slide.placeholders[1].text = text

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Motivazione calcolo"

    text = ""
    for e in explanations[:10]:
        text += f"• {e}\n"

    slide.placeholders[1].text = text

    return prs


def generate_excel(nome, risk, level, driver_df, actions_df, explanations_df, data):
    buffer = BytesIO()

    summary = pd.DataFrame([{
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Cantiere": nome,
        "Risk Index": round(risk, 1),
        "Livello": level,
        "Fase": data["fase"],
        "Ispezioni": data["inspections"],
        "NC": data["num_nc"],
        "Stop Work": data["stopworks"],
        "Criticità aperte": data["crit_open"],
        "Criticità in tempo": data["crit_ontime"],
        "Criticità in ritardo": data["crit_late"],
        "Appaltatori": data["app"],
        "Subappaltatori": data["sub"],
        "Sensibilizzazioni numero": data["awareness_count"],
        "Tipologie attività": ", ".join(data["activities"]),
        "Tipologie sensibilizzazioni": ", ".join(data["awareness_types"])
    }])

    nc_df = pd.DataFrame(data["nc_items"])

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Sintesi", index=False)
        driver_df.to_excel(writer, sheet_name="Driver rischio", index=False)
        actions_df.to_excel(writer, sheet_name="Azioni", index=False)
        explanations_df.to_excel(writer, sheet_name="Motivazione", index=False)
        nc_df.to_excel(writer, sheet_name="NC dettaglio", index=False)

    buffer.seek(0)
    return buffer


# =========================
# UI
# =========================
st.title("🦺 HSE Risk Platform PRO")
st.caption("Valutazione rischio HSE con correlazione tra attività, NC, sensibilizzazioni e output operativo.")

with st.sidebar:
    st.header("📋 Dati cantiere")

    nome = st.text_input("Nome cantiere", value="Cantiere Demo")

    fase = st.selectbox(
        "Fase",
        ["cantierizzazione", "costruzione", "commissioning", "punch list"]
    )

    st.divider()

    st.header("🔍 Controlli e NC")

    inspections = st.number_input("Numero ispezioni", 0, 1000, 10)
    num_nc = st.number_input("Numero totale NC", 0, 300, 0)

    st.write("Dettaglio NC")

    nc_items = []

    nc_count_detail = st.number_input(
        "Quante NC vuoi dettagliare?",
        0,
        20,
        min(num_nc, 3)
    )

    for i in range(nc_count_detail):
        col1, col2 = st.columns(2)

        with col1:
            theme = st.selectbox(
                f"Tema NC {i + 1}",
                list(NC_THEME_WEIGHTS.keys()),
                key=f"nc_theme_{i}"
            )

        with col2:
            severity = st.selectbox(
                f"Gravità NC {i + 1}",
                list(NC_SEVERITY_WEIGHTS.keys()),
                key=f"nc_severity_{i}"
            )

        nc_items.append({
            "theme": theme,
            "severity": severity
        })

    st.divider()

    st.header("⛔ Stop Work e criticità")

    stopworks = st.number_input("Stop Work", 0, 100, 0)
    crit_open = st.number_input("Criticità aperte", 0, 200, 0)
    crit_ontime = st.number_input("Criticità risolte in tempo", 0, 200, 0)
    crit_late = st.number_input("Criticità risolte in ritardo", 0, 200, 0)

    st.divider()

    st.header("🏗️ Organizzazione")

    app = st.number_input("Appaltatori", 0, 100, 1)
    sub = st.number_input("Subappaltatori", 0, 100, 0)

    st.divider()

    st.header("⚙️ Attività in corso")

    activities = st.multiselect(
        "Tipologie attività presenti",
        list(ACTIVITY_WEIGHTS.keys()),
        default=["civili"]
    )

    st.divider()

    st.header("🧠 Sensibilizzazioni")

    awareness_count = st.number_input(
        "Numero sensibilizzazioni effettuate",
        0,
        1000,
        0
    )

    awareness_types = st.multiselect(
        "Tipologie sensibilizzazioni effettuate",
        list(AWARENESS_BONUS.keys())
    )

    st.divider()

    calcola = st.button("Calcola rischio", type="primary")


# =========================
# MAIN
# =========================
if not calcola:
    st.info("Compila i dati nella sidebar e clicca su **Calcola rischio**.")
    st.stop()


data = {
    "nome": nome,
    "fase": fase,
    "inspections": inspections,
    "num_nc": num_nc,
    "nc_items": nc_items,
    "stopworks": stopworks,
    "crit_open": crit_open,
    "crit_ontime": crit_ontime,
    "crit_late": crit_late,
    "app": app,
    "sub": sub,
    "activities": normalize_list(activities),
    "awareness_count": awareness_count,
    "awareness_types": normalize_list(awareness_types)
}

risk, level, details, explanations = calculate_risk(data)
actions = generate_actions(data, risk, level)

driver_df = pd.DataFrame(details, columns=["Driver", "Valore"])
actions_df = pd.DataFrame(actions, columns=["Priorità", "Azione"])
explanations_df = pd.DataFrame(explanations, columns=["Motivazione"])

# =========================
# DASHBOARD
# =========================
col1, col2, col3, col4 = st.columns(4)

col1.metric("Risk Index", f"{round(risk)} / 100")
col2.metric("Livello", level)
col3.metric("NC", num_nc)
col4.metric("Attività", len(activities))

st.divider()

left, right = st.columns([1.2, 1])

with left:
    st.subheader("📊 Driver rischio")
    st.bar_chart(driver_df.set_index("Driver"))

with right:
    st.subheader("📌 Lettura del risultato")

    st.write(f"""
    Il cantiere **{nome}** ha un indice rischio pari a **{round(risk)} / 100**.

    Il livello è **{level}**.

    Il calcolo considera:
    - gravità e tema delle NC;
    - attività effettivamente in corso;
    - coerenza tra attività e NC;
    - criticità aperte o chiuse in ritardo;
    - complessità organizzativa;
    - fase del cantiere;
    - bonus autonomo per sensibilizzazioni mirate.
    """)

    if awareness_count > 0 and not awareness_types:
        st.warning(
            "Hai inserito un numero di sensibilizzazioni ma nessuna tipologia. "
            "Per valorizzarle meglio, seleziona almeno una tipologia."
        )

    if num_nc > len(nc_items):
        st.warning(
            f"Hai indicato {num_nc} NC totali ma ne hai dettagliate solo {len(nc_items)}. "
            "Il rischio sarà più preciso se dettagli più NC."
        )

st.divider()

# =========================
# DETTAGLIO
# =========================
tab1, tab2, tab3, tab4 = st.tabs([
    "🎯 Azioni",
    "🧠 Motivazione",
    "📋 NC",
    "📤 Export"
])

with tab1:
    st.subheader("Azioni suggerite")
    st.dataframe(actions_df, use_container_width=True)

with tab2:
    st.subheader("Perché il sistema ha calcolato questo rischio")
    st.dataframe(explanations_df, use_container_width=True)

with tab3:
    st.subheader("Dettaglio NC inserite")

    if nc_items:
        st.dataframe(pd.DataFrame(nc_items), use_container_width=True)
    else:
        st.info("Nessuna NC dettagliata.")

with tab4:
    st.subheader("Download report")

    ppt = generate_ppt(
        nome,
        risk,
        level,
        driver_df,
        actions_df,
        explanations,
        data
    )

    ppt_buffer = BytesIO()
    ppt.save(ppt_buffer)
    ppt_buffer.seek(0)

    excel_buffer = generate_excel(
        nome,
        risk,
        level,
        driver_df,
        actions_df,
        explanations_df,
        data
    )

    safe_nome = nome.replace(" ", "_").replace("/", "_")

    col_a, col_b = st.columns(2)

    with col_a:
        st.download_button(
            label="📥 Scarica PowerPoint",
            data=ppt_buffer,
            file_name=f"HSE_Risk_Report_{safe_nome}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    with col_b:
        st.download_button(
            label="📥 Scarica Excel",
            data=excel_buffer,
            file_name=f"HSE_Risk_Report_{safe_nome}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
