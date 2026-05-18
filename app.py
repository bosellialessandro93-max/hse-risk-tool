# ... (lascio invariata tutta la parte iniziale già funzionante)

def generate_actions(data, level, nc_types, nc_themes, stopScore, critScore):

    actions = []

    # Criticità
    if data["criticalities"] > 5:
        actions.append("🔴 HIGH → backlog criticità elevato: attivare piano chiusura")
        if data["stopworks"] > 0:
            actions.append("🔴 Stop Work non convertite in azioni efficaci")

    # NC TEMATICHE
    text_nc = " ".join(nc_themes).lower()

    if "elettrico" in text_nc:
        actions.append("🟡 Intervento su rischio elettrico")

    if "quota" in text_nc:
        actions.append("🟡 Rafforzare sicurezza lavori in quota")

    if "dpi" in text_nc:
        actions.append("🟡 Migliorare utilizzo DPI")

    # Stop work pattern
    if data["stopworks"] > data["num_nc"]:
        actions.append("🟡 Disallineamento Stop Work / NC")

    # Fase
    if data["fase"] == "commissioning":
        actions.append("🔴 Commissioning → rischio operativo alto")

    # Ispezioni
    if data["inspections"] < 5:
        actions.append("🟡 Incrementare ispezioni")

    if data["inspections"] > 30:
        actions.append("🟢 Buon livello di controllo")

    return actions
