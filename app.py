from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches


st.set_page_config(
    page_title="HSE Risk Platform AI Locale",
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


def normalize_list(values):
    return [v.strip().lower() for v in values if v.strip()]


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

        explanations.append(
            f"NC su '{theme}' con gravità '{severity}': +{item_score} punti."
        )

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
        explanations.append(
            f"{data['stopworks']} Stop Work registrati: +{stop_score} punti."
        )

    crit_score = min(
        data["crit_open"] * 12 +
        data["crit_late"] * 10 +
        data["crit_ontime"] * 3,
        100
    )

    if data["crit_open"] > 0:
        explanations.append(f"{data['crit_open']} criticità aperte: incremento rischio.")

    if data["crit_late"] > 0:
        explanations.append(
            f"{data['crit_late']} criticità risolte in ritardo: peggiora il follow-up."
        )

    complexity_score = min(
        data["app"] * 5 +
        data["sub"] * 8,
        70
    )

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
        expected_nc = LINKS_ACTIVITY_NC.get(act, [])

        for theme in expected_nc:
            if theme in nc_themes:
                correlation_penalty += 15
                explanations.append(
                    f"Correlazione critica: attività '{act}' con NC su '{theme}': +15 punti."
                )

    correlation_penalty = min(correlation_penalty, 45)

    fase_score = 0

    if data["fase"] == "commissioning":
        fase_score += 18
        explanations.append("Fase commissioning: +18 punti per prove, energie e interferenze.")
    elif data["fase"] == "punch list":
        fase_score += 10
        explanations.append("Fase punch list: +10 punti per attività residue e discontinuità operative.")
    elif data["fase"] == "cantierizzazione":
        fase_score += 8
        explanations.append("Fase cantierizzazione: +8 punti per allestimenti e avvio attività.")

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
                explanations.append(
                    f"Sensibilizzazione coerente con attività '{act}': -5 punti."
                )

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


def generate_base_actions(data, risk, level):
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
                "⚡ Verificare sezionamenti, LOTO, autorizzazioni, DPI e competenze PES/PAV/PEI."
            ))

        if act == "scavi":
            actions.append((
                "MUST HAVE",
                "🚧 Verificare sottoservizi, stabilità fronti, accessi, delimitazioni e presenza acqua/materiale instabile."
            ))

        if act == "lavori in quota":
            actions.append((
                "MUST HAVE",
                "🪜 Verificare parapetti, ancoraggi, PLE, imbracature, accessi e piano di emergenza."
            ))

        if act == "sollevamenti":
            actions.append((
                "MUST HAVE",
                "🏗️ Verificare piano di sollevamento, portate, accessori, interferenze e area segregata."
            ))

        if act == "spazi confinati":
            actions.append((
                "MUST HAVE",
                "☠️ Verificare autorizzazione, monitoraggio atmosfera, recupero emergenza e presidio esterno."
            ))

        if act == "hot work":
            actions.append((
                "MUST HAVE",
                "🔥 Verificare permesso hot work, estintori, rimozione combustibili, fire watch e controllo post-attività."
            ))

    if len(data["activities"]) >= 3:
        actions.append((
            "MUST HAVE",
            "🔴 Predisporre matrice interferenziale giornaliera e briefing pre-job tra imprese."
        ))

    if "elettrico" in nc_themes:
        actions.append((
            "MUST HAVE",
            "🔴 NC elettriche: sospendere attività non conformi fino a ripristino condizioni sicure."
        ))

    if "scavi" in nc_themes:
        actions.append((
            "MUST HAVE",
            "🔴 NC su scavi: rivalutare condizioni dello scavo e sottoservizi prima della prosecuzione."
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


def local_executive_summary(data, risk, level, driver_df):
    top_drivers = driver_df.sort_values("Valore", ascending=False).head(3)
    top_text = ", ".join(
        f"{row['Driver']} ({round(row['Valore'], 1)})"
        for _, row in top_drivers.iterrows()
    )

    activities = ", ".join(data["activities"]) if data["activities"] else "nessuna attività critica selezionata"
    awareness = ", ".join(data["awareness_types"]) if data["awareness_types"] else "nessuna sensibilizzazione mirata"

    if level in ["Alto", "Critico"]:
        decision = (
            "Si raccomanda un coordinamento operativo immediato, con verifica delle attività critiche, "
            "chiusura delle NC prioritarie e briefing pre-job prima della prosecuzione."
        )
    elif level == "Medio":
        decision = (
            "Si raccomanda di rafforzare il presidio operativo, chiudere le criticità aperte "
            "e programmare sensibilizzazioni mirate sulle attività più esposte."
        )
    else:
        decision = (
            "Il livello è sotto controllo, ma va mantenuto il monitoraggio periodico e l’aggiornamento "
            "della valutazione in caso di variazioni operative."
        )

    return f"""
Il cantiere **{data['nome']}** presenta un Risk Index pari a **{round(risk)} / 100**, classificato come **{level}**.

I principali driver che influenzano il rischio sono: **{top_text}**.

Le attività in corso sono: **{activities}**.  
Le sensibilizzazioni registrate sono: **{awareness}**.

{decision}
"""


def local_root_cause(data, risk, level):
    nc_themes = [item["theme"] for item in data["nc_items"]]
    severe_nc = [item for item in data["nc_items"] if item["severity"] in ["grave", "critica"]]

    causes = []

    if severe_nc:
        causes.append("presenza di NC gravi o critiche non ancora compensate da azioni correttive robuste")

    if len(data["activities"]) >= 3:
        causes.append("elevato rischio interferenziale dovuto alla contemporaneità delle attività")

    if data["crit_open"] > 0:
        causes.append("presenza di criticità aperte che aumentano l’esposizione residua")

    if data["crit_late"] > 0:
        causes.append("ritardi nella chiusura delle criticità, indice di follow-up non pienamente efficace")

    if data["sub"] >= 5:
        causes.append("complessità organizzativa elevata per numero significativo di subappaltatori")

    if data["awareness_count"] == 0:
        causes.append("assenza di sensibilizzazioni registrate")

    if not causes:
        causes.append("nessuna causa dominante evidente; il rischio deriva dalla combinazione dei fattori inseriti")

    controls = []

    if "elettrico" in nc_themes or "elettrici" in data["activities"]:
        controls.append("verifica LOTO, sezionamenti, autorizzazioni elettriche e competenze PES/PAV/PEI")

    if "scavi" in nc_themes or "scavi" in data["activities"]:
        controls.append("verifica sottoservizi, stabilità fronti, accessi e segregazione scavi")

    if "quota" in nc_themes or "lavori in quota" in data["activities"]:
        controls.append("verifica protezioni collettive, ancoraggi, PLE e DPI anticaduta")

    if "sollevamenti" in data["activities"]:
        controls.append("verifica piano di sollevamento, portate, accessori e area segregata")

    if not controls:
        controls.append("verifica controlli operativi standard, briefing pre-job e tracciamento NC")

    return f"""
### Cause probabili
- {chr(10).join(["- " + c for c in causes])}

### Debolezze organizzative possibili
- Pianificazione operativa non pienamente allineata ai rischi reali di campo.
- Follow-up delle criticità da rafforzare.
- Presidio delle interferenze da formalizzare meglio quando sono presenti più attività.

### Controlli da verificare
- {chr(10).join(["- " + c for c in controls])}

### Conclusione operativa
Il livello **{level}** richiede un piano azioni proporzionato al rischio, con priorità alle attività critiche, alle NC gravi/critiche e alle criticità ancora aperte.
"""


def local_toolbox_talk(data):
    activities = data["activities"]
    nc_themes = [item["theme"] for item in data["nc_items"]]

    title_parts = []

    if activities:
        title_parts.extend(activities[:2])

    if nc_themes:
        title_parts.extend(nc_themes[:2])

    title = " / ".join(title_parts) if title_parts else "Sicurezza operativa giornaliera"

    rules = [
        "Non iniziare attività senza autorizzazione e briefing pre-job.",
        "Segnalare immediatamente condizioni non sicure o interferenze.",
        "Mantenere l’area ordinata, accessibile e segregata dove necessario.",
        "Usare DPI coerenti con il rischio specifico dell’attività.",
        "Fermare l’attività in caso di dubbio o condizione non controllata."
    ]

    if "elettrici" in activities or "elettrico" in nc_themes:
        rules.append("Per attività elettriche: verificare sezionamento, assenza tensione e procedura LOTO.")

    if "scavi" in activities or "scavi" in nc_themes:
        rules.append("Per scavi: verificare sottoservizi, stabilità fronti e accessi sicuri.")

    if "lavori in quota" in activities or "quota" in nc_themes:
        rules.append("Per lavori in quota: verificare protezioni anticaduta prima di accedere all’area.")

    if "sollevamenti" in activities:
        rules.append("Per sollevamenti: nessuno deve sostare sotto carichi sospesi o in area non autorizzata.")

    return f"""
### Titolo
**Toolbox Talk - {title}**

### Obiettivo
Allineare la squadra sui rischi principali della giornata e sulle condizioni minime per lavorare in sicurezza.

### Rischi principali
- Attività in corso: **{", ".join(activities) if activities else "non specificate"}**
- NC rilevate: **{", ".join(nc_themes) if nc_themes else "nessuna NC dettagliata"}**
- Fase del cantiere: **{data["fase"]}**

### Regole operative
{chr(10).join(["- " + r for r in rules])}

### Comportamenti vietati
- Improvvisare lavorazioni fuori procedura.
- Rimuovere protezioni o segregazioni senza autorizzazione.
- Operare in aree interferenti senza coordinamento.
- Proseguire il lavoro dopo una condizione di pericolo non risolta.

### Domande finali alla squadra
1. Qual è il rischio principale dell’attività di oggi?
2. Quali autorizzazioni servono prima di iniziare?
3. Quali DPI sono obbligatori?
4. Chi avviso se cambia la condizione di lavoro?
5. Quando devo fermare l’attività?
"""


def local_action_plan(data, risk, level):
    rows = []

    if risk >= 80:
        rows.append(["Coordinamento HSE immediato", "Critica", "Site Manager / HSE", "Immediata", "Verbale briefing e piano azioni"])

    if data["crit_open"] > 0:
        rows.append(["Chiudere criticità aperte prioritarie", "Alta", "HSE / Responsabile impresa", "24-48h", "Registro criticità aggiornato"])

    if data["crit_late"] > 0:
        rows.append(["Analizzare ritardi chiusura criticità", "Media", "HSE Manager", "7 giorni", "Analisi cause e azioni correttive"])

    if len(data["activities"]) >= 3:
        rows.append(["Creare matrice interferenze giornaliera", "Alta", "Construction Manager", "24h", "Matrice interferenze firmata"])

    if "elettrici" in data["activities"]:
        rows.append(["Verifica attività elettriche e LOTO", "Alta", "Preposto / HSE", "24h", "Checklist elettrica e autorizzazioni"])

    if "scavi" in data["activities"]:
        rows.append(["Verifica sicurezza scavi", "Alta", "Preposto / HSE", "24h", "Checklist scavi e verifica sottoservizi"])

    if "lavori in quota" in data["activities"]:
        rows.append(["Verifica lavori in quota", "Alta", "Preposto / HSE", "24h", "Checklist quota / PLE / ancoraggi"])

    if data["awareness_count"] == 0:
        rows.append(["Eseguire toolbox mirato", "Alta", "HSE / Preposto", "Prima dell’attività", "Registro presenze toolbox"])

    missing = []

    if "elettrici" in data["activities"] and "rischio elettrico" not in data["awareness_types"]:
        missing.append("rischio elettrico")

    if "scavi" in data["activities"] and "scavi e sottoservizi" not in data["awareness_types"]:
        missing.append("scavi e sottoservizi")

    if "lavori in quota" in data["activities"] and "lavori in quota" not in data["awareness_types"]:
        missing.append("lavori in quota")

    if missing:
        rows.append([
            "Eseguire sensibilizzazioni mirate: " + ", ".join(missing),
            "Media",
            "HSE",
            "48h",
            "Registro sensibilizzazioni"
        ])

    if not rows:
        rows.append(["Mantenere monitoraggio HSE periodico", "Media", "HSE", "7 giorni", "Report ispezione aggiornato"])

    df = pd.DataFrame(rows, columns=["Azione", "Priorità", "Owner suggerito", "Scadenza", "Evidenza richiesta"])
    return df


def local_reduce_risk(data, risk):
    target = 80 if risk > 80 else 60 if risk > 60 else 40 if risk > 40 else 20

    levers = []

    if data["crit_open"] > 0:
        levers.append("Chiudere le criticità aperte più rilevanti può ridurre in modo significativo il rischio residuo.")

    if len(data["activities"]) >= 3:
        levers.append("Ridurre o separare temporalmente le attività contemporanee abbassa il rischio interferenziale.")

    if data["awareness_count"] == 0:
        levers.append("Eseguire toolbox mirati sulle attività critiche genera mitigazione immediata.")

    if data["num_nc"] > 0:
        levers.append("Chiudere NC gravi/critiche con verifica di efficacia riduce il driver NC e le correlazioni critiche.")

    if data["sub"] >= 5:
        levers.append("Rafforzare coordinamento subappaltatori e responsabilità operative riduce la complessità.")

    if "elettrici" in data["activities"]:
        levers.append("Verificare LOTO, autorizzazioni e competenze elettriche riduce il rischio su attività elettriche.")

    if "scavi" in data["activities"]:
        levers.append("Verificare sottoservizi, fronti e segregazioni riduce il rischio operativo sugli scavi.")

    if not levers:
        levers.append("Mantenere controlli periodici, sensibilizzazioni mirate e aggiornamento del rischio al cambio attività.")

    return f"""
### Obiettivo suggerito
Portare il Risk Index sotto **{target}**.

### Leve principali
{chr(10).join(["- " + l for l in levers])}

### Strategia consigliata
Agire prima sui fattori che generano rischio diretto: **NC gravi/critiche, criticità aperte, attività contemporanee e assenza di sensibilizzazioni mirate**.
"""


def local_hse_advisor(question, data, risk, level, actions_df):
    q = question.lower()

    if "24" in q or "domani" in q or "subito" in q:
        return """
Nelle prossime 24 ore darei priorità a:
1. Verifica delle attività critiche in corso.
2. Chiusura o messa in sicurezza delle NC gravi/critiche.
3. Briefing pre-job con imprese coinvolte.
4. Verifica criticità aperte.
5. Sensibilizzazione mirata sulle attività più rischiose.
"""

    if "principale" in q or "rischio" in q:
        top_actions = actions_df.head(3)["Azione"].tolist()
        return "Il rischio principale emerge dalla combinazione tra attività critiche, NC e criticità aperte. Azioni prioritarie:\n\n" + "\n".join(["- " + a for a in top_actions])

    if "ridurre" in q or "abbassare" in q:
        return local_reduce_risk(data, risk)

    if "fermare" in q or "stop" in q:
        if level in ["Alto", "Critico"]:
            return "Con livello Alto/Critico valuterei stop o sospensione selettiva delle attività non controllate, soprattutto se correlate a NC gravi, scavi, elettrico, quota, sollevamenti o spazi confinati."
        return "Non emerge automaticamente una necessità di fermare le attività, ma va applicato Stop Work se compaiono condizioni non sicure."

    return """
Risposta HSE Advisor locale:
valuta prima NC gravi/critiche, criticità aperte, attività contemporanee, coerenza delle sensibilizzazioni e fase del cantiere.
Le azioni devono essere proporzionate al livello di rischio e tracciate con owner, scadenza ed evidenza.
"""


def generate_ppt(nome, risk, level, driver_df, actions_df, action_plan_df, summary, root_cause, toolbox, data):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "HSE Risk Report"
    slide.placeholders[1].text = f"Cantiere: {nome}\nData: {datetime.now().strftime('%d/%m/%Y')}"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Executive Summary"
    slide.placeholders[1].text = summary.replace("**", "")[:1600]

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
    slide.shapes.title.text = "Action Plan"
    text = ""
    for _, row in action_plan_df.head(7).iterrows():
        text += f"{row['Priorità']} - {row['Azione']} - {row['Scadenza']}\n"
    slide.placeholders[1].text = text[:1800]

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Root Cause Analysis"
    slide.placeholders[1].text = root_cause.replace("###", "").replace("**", "")[:1800]

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Toolbox Talk"
    slide.placeholders[1].text = toolbox.replace("###", "").replace("**", "")[:1800]

    return prs


def generate_excel(nome, risk, level, driver_df, actions_df, action_plan_df, explanations_df, summary, root_cause, toolbox, reduce_risk, data):
    buffer = BytesIO()

    summary_df = pd.DataFrame([{
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

    local_outputs_df = pd.DataFrame([
        {"Sezione": "Executive Summary", "Testo": summary},
        {"Sezione": "Root Cause Analysis", "Testo": root_cause},
        {"Sezione": "Toolbox Talk", "Testo": toolbox},
        {"Sezione": "Come ridurre il rischio", "Testo": reduce_risk}
    ])

    nc_df = pd.DataFrame(data["nc_items"])

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Sintesi", index=False)
        driver_df.to_excel(writer, sheet_name="Driver rischio", index=False)
        actions_df.to_excel(writer, sheet_name="Azioni base", index=False)
        action_plan_df.to_excel(writer, sheet_name="Action Plan", index=False)
        explanations_df.to_excel(writer, sheet_name="Motivazione", index=False)
        local_outputs_df.to_excel(writer, sheet_name="Output locali", index=False)
        nc_df.to_excel(writer, sheet_name="NC dettaglio", index=False)

    buffer.seek(0)
    return buffer


st.title("🦺 HSE Risk Platform AI Locale")
st.caption("Risk assessment HSE con Executive Summary, Root Cause, Toolbox Talk, Action Plan e Advisor locale gratuito.")

st.success("Modalità gratuita attiva: nessuna API key richiesta, nessun costo OpenAI.")

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
base_actions = generate_base_actions(data, risk, level)

driver_df = pd.DataFrame(details, columns=["Driver", "Valore"])
actions_df = pd.DataFrame(base_actions, columns=["Priorità", "Azione"])
explanations_df = pd.DataFrame(explanations, columns=["Motivazione"])

summary = local_executive_summary(data, risk, level, driver_df)
root_cause = local_root_cause(data, risk, level)
toolbox = local_toolbox_talk(data)
action_plan_df = local_action_plan(data, risk, level)
reduce_risk = local_reduce_risk(data, risk)


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
    st.subheader("🤖 Executive Summary locale")
    st.markdown(summary)

    if awareness_count > 0 and not awareness_types:
        st.warning("Hai inserito sensibilizzazioni senza tipologia: seleziona le tipologie per valorizzarle meglio.")

    if num_nc > len(nc_items):
        st.warning(f"Hai indicato {num_nc} NC totali ma ne hai dettagliate solo {len(nc_items)}.")


st.divider()

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🎯 Azioni base",
    "📋 Action Plan",
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
    st.subheader("Action Plan operativo")
    st.dataframe(action_plan_df, use_container_width=True)

with tab3:
    st.subheader("Root Cause Analysis locale")
    st.markdown(root_cause)

with tab4:
    st.subheader("Toolbox Talk locale")
    st.markdown(toolbox)

with tab5:
    st.subheader("Come ridurre il rischio")
    st.markdown(reduce_risk)

with tab6:
    st.subheader("Chiedi all'HSE Advisor locale")

    question = st.text_input(
        "Fai una domanda sul report",
        placeholder="Esempio: cosa devo fare nelle prossime 24 ore?"
    )

    if st.button("Chiedi all'Advisor"):
        if question.strip():
            answer = local_hse_advisor(question, data, risk, level, actions_df)
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
        action_plan_df,
        summary,
        root_cause,
        toolbox,
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
        action_plan_df,
        explanations_df,
        summary,
        root_cause,
        toolbox,
        reduce_risk,
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
