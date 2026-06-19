import os
import json
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


st.set_page_config(
    page_title="HSE Risk Platform AI",
    page_icon="🦺",
    layout="wide"
)


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


def get_api_key():
    key = None

    try:
        key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        key = None

    if not key:
        key = os.getenv("OPENAI_API_KEY")

    return key


def has_ai():
    return OpenAI is not None and bool(get_api_key())


def call_ai(prompt, max_tokens=1200):
    if not has_ai():
        return None

    try:
        client = OpenAI(api_key=get_api_key())

        response = client.responses.create(
            model="gpt-5.5",
            instructions=(
                "Sei un HSE Manager senior. Scrivi in italiano, tono professionale, "
                "pratico, diretto. Non inventare dati. Basa l'analisi solo sugli input forniti. "
                "Le indicazioni sono supporto decisionale e non sostituiscono obblighi normativi, "
                "valutazioni specialistiche o procedure aziendali."
            ),
            input=prompt,
            max_output_tokens=max_tokens
        )

        return response.output_text

    except Exception as e:
        return f"AI non disponibile: {e}"


def risk_level(risk):
    if risk <= 20:
        return "Basso"
    if risk <= 40:
        return "Medio basso"
    if risk <= 60:
        return "Medio"
    if risk <= 80:
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

    nc_score = 0

    for item in nc_items:
        theme = item["theme"]
        severity = item["severity"]
        theme_score = NC_THEME_WEIGHTS.get(theme, 6)
        severity_score = NC_SEVERITY_WEIGHTS.get(severity, 5)
        item_score = theme_score + severity_score
        nc_score += item_score
        explanations.append(f"NC su '{theme}' con gravità '{severity}': +{item_score} punti.")

    nc_score = min(nc_score, 100)

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

    stop_score = min(data["stopworks"] * 10, 40)

    if data["stopworks"] > 0:
        explanations.append(f"{data['stopworks']} Stop Work registrati: +{stop_score} punti.")

    crit_score = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 10 +
        data["crit_ontime"] * 3,
        100
    )

    if data["crit_open"] > 0:
        explanations.append(f"{data['crit_open']} criticità aperte: incremento rischio.")

    if data["crit_late"] > 0:
        explanations.append(f"{data['crit_late']} criticità risolte in ritardo: peggiora il follow-up.")

    complexity_score = min(data["app"] * 5 + data["sub"] * 8, 70)

    if data["sub"] >= 5:
        complexity_score += 15
        explanations.append("Numero elevato di subappaltatori: +15 punti.")

    complexity_score = min(complexity_score, 100)

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

    correlation_penalty = 0
    nc_themes = [item["theme"] for item in nc_items]

    for act in activities:
        for theme in LINKS_ACTIVITY_NC.get(act, []):
            if theme in nc_themes:
                correlation_penalty += 15
                explanations.append(f"Correlazione critica: attività '{act}' con NC su '{theme}': +15 punti.")

    correlation_penalty = min(correlation_penalty, 45)

    fase_score = 0

    if data["fase"] == "commissioning":
        fase_score += 18
        explanations.append("Fase commissioning: +18 punti.")
    elif data["fase"] == "punch list":
        fase_score += 10
        explanations.append("Fase punch list: +10 punti.")
    elif data["fase"] == "cantierizzazione":
        fase_score += 8
        explanations.append("Fase cantierizzazione: +8 punti.")

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

    for aw in awareness_types:
        covered_activities = LINKS_AWARENESS_ACTIVITY.get(aw, [])

        for act in activities:
            if act in covered_activities:
                awareness_bonus += 5
                explanations.append(f"Sensibilizzazione coerente con attività '{act}': -5 punti.")

    awareness_bonus = min(awareness_bonus, 35)

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
    severe_nc = [item for item in data["nc_items"] if item["severity"] in ["grave", "critica"]]

    if risk >= 80:
        actions.append(("MUST HAVE", "🔴 Rischio critico: coordinamento operativo immediato prima della prosecuzione."))

    if severe_nc:
        actions.append(("MUST HAVE", f"🔴 Presenti {len(severe_nc)} NC gravi/critiche. Definire owner, scadenza e verifica chiusura."))

    for act in data["activities"]:
        if act == "elettrici":
            actions.append(("MUST HAVE", "⚡ Verificare sezionamenti, autorizzazioni, LOTO, DPI e competenze PES/PAV/PEI."))
        if act == "scavi":
            actions.append(("MUST HAVE", "🚧 Verificare sottoservizi, fronti, accessi, delimitazioni e stabilità scavo."))
        if act == "lavori in quota":
            actions.append(("MUST HAVE", "🪜 Verificare parapetti, ancoraggi, PLE, imbracature e piano emergenza."))
        if act == "sollevamenti":
            actions.append(("MUST HAVE", "🏗️ Verificare piano sollevamento, portate, accessori, interferenze e area segregata."))
        if act == "spazi confinati":
            actions.append(("MUST HAVE", "☠️ Verificare permesso, atmosfera, recupero emergenza, presidio esterno e procedura."))
        if act == "hot work":
            actions.append(("MUST HAVE", "🔥 Verificare permesso hot work, estintori, rimozione combustibili e fire watch."))

    if len(data["activities"]) >= 3:
        actions.append(("MUST HAVE", "🔴 Predisporre matrice interferenziale giornaliera e briefing pre-job tra imprese."))

    if "elettrico" in nc_themes:
        actions.append(("MUST HAVE", "🔴 NC elettriche: sospendere attività non conformi fino a ripristino condizioni sicure."))

    if "scavi" in nc_themes:
        actions.append(("MUST HAVE", "🔴 NC scavi: rivalutare condizioni dello scavo prima della prosecuzione."))

    if data["awareness_count"] == 0:
        actions.append(("MUST HAVE", "🟡 Nessuna sensibilizzazione registrata: pianificare toolbox mirati prima delle attività critiche."))

    missing = []

    if "elettrici" in data["activities"] and "rischio elettrico" not in data["awareness_types"]:
        missing.append("rischio elettrico")
    if "scavi" in data["activities"] and "scavi e sottoservizi" not in data["awareness_types"]:
        missing.append("scavi e sottoservizi")
    if "lavori in quota" in data["activities"] and "lavori in quota" not in data["awareness_types"]:
        missing.append("lavori in quota")
    if "sollevamenti" in data["activities"] and "sollevamenti" not in data["awareness_types"]:
        missing.append("sollevamenti")

    if missing:
        actions.append(("MUST HAVE", "🟠 Mancano sensibilizzazioni mirate su: " + ", ".join(missing) + "."))

    if data["crit_open"] > 0:
        actions.append(("MUST HAVE", f"🟠 Criticità aperte: {data['crit_open']}. Definire priorità, owner e data chiusura."))

    if len(actions) < 4:
        actions.append(("IDEA", "Mantenere presidio operativo e aggiornare la valutazione in caso di nuove attività."))

    return actions


def compact_context(data, risk, level, driver_df, actions_df, explanations_df):
    return json.dumps({
        "cantiere": data["nome"],
        "fase": data["fase"],
        "risk_index": round(risk, 1),
        "livello": level,
        "ispezioni": data["inspections"],
        "nc_totali": data["num_nc"],
        "nc_dettaglio": data["nc_items"],
        "stop_work": data["stopworks"],
        "criticita_aperte": data["crit_open"],
        "criticita_tempo": data["crit_ontime"],
        "criticita_ritardo": data["crit_late"],
        "appaltatori": data["app"],
        "subappaltatori": data["sub"],
        "attivita": data["activities"],
        "sensibilizzazioni_numero": data["awareness_count"],
        "sensibilizzazioni_tipologia": data["awareness_types"],
        "driver": driver_df.to_dict(orient="records"),
        "azioni_base": actions_df.to_dict(orient="records"),
        "motivazioni_calcolo": explanations_df.to_dict(orient="records")
    }, ensure_ascii=False, indent=2)


def fallback_executive_summary(data, risk, level):
    activities = ", ".join(data["activities"]) if data["activities"] else "nessuna attività critica selezionata"
    return (
        f"Il cantiere {data['nome']} presenta un Risk Index pari a {round(risk)}/100, "
        f"classificato come {level}. Il rischio è influenzato da fase {data['fase']}, "
        f"attività in corso ({activities}), {data['num_nc']} NC totali e "
        f"{data['crit_open']} criticità aperte. Le sensibilizzazioni incidono come bonus autonomo "
        f"e risultano più efficaci quando sono coerenti con le attività effettivamente presenti."
    )


def generate_ai_outputs(data, risk, level, driver_df, actions_df, explanations_df):
    context = compact_context(data, risk, level, driver_df, actions_df, explanations_df)

    outputs = {}

    prompts = {
        "Executive Summary": f"""
Genera un Executive Summary HSE manageriale in massimo 180 parole.
Deve spiegare perché il rischio è {level}, quali sono i driver principali,
quali elementi mitigano il rischio e quale decisione operativa suggerisci.

DATI:
{context}
""",
        "Root Cause Analysis": f"""
Esegui una Root Cause Analysis HSE.
Restituisci:
1. Cause probabili
2. Debolezze organizzative
3. Segnali precursori
4. Controlli da verificare
5. Conclusione operativa

DATI:
{context}
""",
        "Toolbox Talk": f"""
Genera un Toolbox Talk di 5 minuti per gli operatori.
Struttura:
- Titolo
- Obiettivo
- Rischi principali
- Regole operative
- Comportamenti vietati
- 5 domande finali per verificare comprensione

DATI:
{context}
""",
        "Action Plan AI": f"""
Genera un piano azioni HSE operativo in tabella markdown con colonne:
Azione | Priorità | Owner suggerito | Scadenza | Evidenza richiesta.
Le scadenze devono essere realistiche: immediata, 24h, 48h, 7 giorni.

DATI:
{context}
""",
        "Come ridurre il rischio": f"""
Suggerisci come ridurre il Risk Index sotto la soglia inferiore più vicina.
Esempio: se rischio 75, obiettivo sotto 60; se 58, obiettivo sotto 40.
Indica massimo 7 leve concrete e l'impatto atteso qualitativo.

DATI:
{context}
"""
    }

    for name, prompt in prompts.items():
        ai_text = call_ai(prompt)

        if ai_text:
            outputs[name] = ai_text
        else:
            if name == "Executive Summary":
                outputs[name] = fallback_executive_summary(data, risk, level)
            else:
                outputs[name] = "AI non configurata. Aggiungi OPENAI_API_KEY per generare questa sezione."

    return outputs


def answer_hse_advisor(question, data, risk, level, driver_df, actions_df, explanations_df, ai_outputs):
    context = compact_context(data, risk, level, driver_df, actions_df, explanations_df)

    prompt = f"""
Rispondi alla domanda dell'utente come HSE Advisor.
Usa solo i dati del report.
Sii concreto, operativo e sintetico.

DOMANDA:
{question}

DATI REPORT:
{context}

SEZIONI AI GIÀ GENERATE:
{json.dumps(ai_outputs, ensure_ascii=False, indent=2)}
"""

    ai_text = call_ai(prompt, max_tokens=900)

    if ai_text:
        return ai_text

    return (
        "AI non configurata. Posso comunque dirti che il rischio va letto partendo da: "
        "NC gravi/critiche, attività contemporanee, correlazioni attività-NC, criticità aperte "
        "e presenza o assenza di sensibilizzazioni mirate."
    )


def generate_ppt(nome, risk, level, driver_df, actions_df, explanations_df, ai_outputs, data):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report AI"
    slide.placeholders[1].text = f"Cantiere: {nome}\nData: {datetime.now().strftime('%d/%m/%Y')}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Executive Summary"
    slide.placeholders[1].text = ai_outputs.get("Executive Summary", "")[:1500]

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
    slide.shapes.title.text = "Action Plan AI"
    slide.placeholders[1].text = ai_outputs.get("Action Plan AI", "")[:1800]

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Root Cause Analysis"
    slide.placeholders[1].text = ai_outputs.get("Root Cause Analysis", "")[:1800]

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Toolbox Talk"
    slide.placeholders[1].text = ai_outputs.get("Toolbox Talk", "")[:1800]

    return prs


def generate_excel(nome, risk, level, driver_df, actions_df, explanations_df, ai_outputs, data):
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

    ai_df = pd.DataFrame(
        [{"Sezione": k, "Testo": v} for k, v in ai_outputs.items()]
    )

    nc_df = pd.DataFrame(data["nc_items"])

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Sintesi", index=False)
        driver_df.to_excel(writer, sheet_name="Driver rischio", index=False)
        actions_df.to_excel(writer, sheet_name="Azioni base", index=False)
        explanations_df.to_excel(writer, sheet_name="Motivazione", index=False)
        ai_df.to_excel(writer, sheet_name="AI Output", index=False)
        nc_df.to_excel(writer, sheet_name="NC dettaglio", index=False)

    buffer.seek(0)
    return buffer


st.title("🦺 HSE Risk Platform AI")
st.caption("Risk assessment HSE con Executive Summary, Root Cause Analysis, Toolbox Talk, Action Plan AI e HSE Advisor.")

if has_ai():
    st.success("AI attiva: OPENAI_API_KEY rilevata.")
else:
    st.warning("AI non configurata. Il tool funziona comunque, ma le sezioni AI useranno fallback o messaggi placeholder.")

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
    generate_ai_now = st.checkbox("Genera sezioni AI", value=True)

    calcola = st.button("Calcola rischio", type="primary")


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

if generate_ai_now:
    with st.spinner("Generazione contenuti AI..."):
        ai_outputs = generate_ai_outputs(data, risk, level, driver_df, actions_df, explanations_df)
else:
    ai_outputs = {
        "Executive Summary": fallback_executive_summary(data, risk, level),
        "Root Cause Analysis": "Generazione AI disattivata.",
        "Toolbox Talk": "Generazione AI disattivata.",
        "Action Plan AI": "Generazione AI disattivata.",
        "Come ridurre il rischio": "Generazione AI disattivata."
    }


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
    st.subheader("🤖 Executive Summary AI")
    st.write(ai_outputs.get("Executive Summary", ""))

    if awareness_count > 0 and not awareness_types:
        st.warning("Hai inserito sensibilizzazioni senza tipologia: seleziona le tipologie per valorizzarle meglio.")

    if num_nc > len(nc_items):
        st.warning(f"Hai indicato {num_nc} NC totali ma ne hai dettagliate solo {len(nc_items)}.")


st.divider()

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🎯 Azioni base",
    "🤖 Action Plan AI",
    "🔍 Root Cause",
    "🧰 Toolbox Talk",
    "📉 Riduci rischio",
    "💬 HSE Advisor",
    "📤 Export"
])

with tab1:
    st.subheader("Azioni base generate dal motore rischio")
    st.dataframe(actions_df, use_container_width=True)

    st.subheader("Motivazione tecnica del calcolo")
    st.dataframe(explanations_df, use_container_width=True)

with tab2:
    st.subheader("Action Plan AI")
    st.markdown(ai_outputs.get("Action Plan AI", ""))

with tab3:
    st.subheader("Root Cause Analysis AI")
    st.markdown(ai_outputs.get("Root Cause Analysis", ""))

with tab4:
    st.subheader("Toolbox Talk AI")
    st.markdown(ai_outputs.get("Toolbox Talk", ""))

with tab5:
    st.subheader("Come ridurre il rischio")
    st.markdown(ai_outputs.get("Come ridurre il rischio", ""))

with tab6:
    st.subheader("Chiedi all'HSE Advisor")

    question = st.text_input(
        "Fai una domanda sul report",
        placeholder="Esempio: cosa devo fare nelle prossime 24 ore?"
    )

    if st.button("Chiedi all'AI"):
        if question.strip():
            with st.spinner("Analisi HSE Advisor..."):
                answer = answer_hse_advisor(
                    question,
                    data,
                    risk,
                    level,
                    driver_df,
                    actions_df,
                    explanations_df,
                    ai_outputs
                )

            st.markdown(answer)
        else:
            st.warning("Scrivi una domanda.")

with tab7:
    st.subheader("Download report")

    ppt = generate_ppt(
        nome,
        risk,
        level,
        driver_df,
        actions_df,
        explanations_df,
        ai_outputs,
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
        ai_outputs,
        data
    )

    safe_nome = nome.replace(" ", "_").replace("/", "_")

    col_a, col_b = st.columns(2)

    with col_a:
        st.download_button(
            label="📥 Scarica PowerPoint AI",
            data=ppt_buffer,
            file_name=f"HSE_Risk_Report_AI_{safe_nome}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    with col_b:
        st.download_button(
            label="📥 Scarica Excel AI",
            data=excel_buffer,
            file_name=f"HSE_Risk_Report_AI_{safe_nome}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
