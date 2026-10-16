"""HSE Risk Platform 2.0 — modello gestionale trasparente, non predittivo.
Avvio: streamlit run app.py. Tutti gli input sono riferiti al mese selezionato.
Nessuna chiamata a modelli AI, nessuna API a pagamento.
"""
from __future__ import annotations
from io import BytesIO
from datetime import datetime, date, timezone
from pathlib import Path
from copy import deepcopy
import calendar
import hmac
import json
import math
import os
import re
import sqlite3
import unicodedata
import uuid
import textwrap
import pandas as pd
import altair as alt
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt as PptPt
from pptx.dml.color import RGBColor
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt, Inches as DocxInches

MODEL_VERSION = '2.0.0'
UI_VERSION = '2.0.2'
SCHEMA_VERSION = 2
WEIGHTS = {'NC formali': 30, 'Criticità informali': 8, 'Attività': 22,
           'Interferenze': 15, 'Organizzazione': 10, 'Presidio e controlli': 15}
THEMES = ['elettrico', 'scavi', 'quota', 'sollevamento', 'spazi confinati',
          'incendio', 'viabilità', 'dpi', 'ordine e pulizia', 'documentale', 'ambientale', 'altro']
THEME_FACTOR = dict(zip(THEMES, [1.2, 1.2, 1.15, 1.15, 1.25, 1.15, 1.05, .85, .7, .65, 1., .85]))
SEVERITY = {'bassa': 1., 'media': 2.5, 'alta': 5.}
STATUS = {'Aperta': 1., 'Chiusa nei tempi': .2, 'Chiusa in ritardo': .4}
ACTIVITY = {'elettrici': 80, 'scavi': 75, 'lavori in quota': 85, 'sollevamenti': 80,
            'spazi confinati': 95, 'hot work': 75, 'movimentazione mezzi': 55,
            'civili': 35, 'meccanici': 45, 'altro': 40}
ACTIVITY_THEME = {'elettrici': 'elettrico', 'scavi': 'scavi', 'lavori in quota': 'quota',
                  'sollevamenti': 'sollevamento', 'spazi confinati': 'spazi confinati',
                  'hot work': 'incendio', 'movimentazione mezzi': 'viabilità',
                  'civili': 'ordine e pulizia', 'meccanici': 'dpi', 'altro': 'altro'}
PHASE_THEMES = {'cantierizzazione': ['scavi', 'viabilità', 'ordine e pulizia'],
                'costruzione': ['quota', 'sollevamento', 'dpi'],
                'commissioning': ['elettrico', 'incendio', 'dpi'],
                'punch list': ['elettrico', 'quota', 'ordine e pulizia']}
COLORS = ['#16826b', '#4b8b9e', '#b38212', '#d46b22', '#b83246']


def bounded(x, low=0., high=100.):
    return max(low, min(high, float(x)))


def risk_level(x):
    return ['Basso', 'Medio basso', 'Medio', 'Alto', 'Critico'][min(4, max(0, math.ceil(x / 20) - 1))]


def site_key(name):
    return ' '.join(unicodedata.normalize('NFKC', name).casefold().split())


def safe_filename(name):
    return re.sub(r'[^\w.-]+', '_', name, flags=re.UNICODE).strip('_')[:90] or 'cantiere'


def default_data():
    today = date.today().replace(day=1)
    last_month = date(today.year - (today.month == 1), 12 if today.month == 1 else today.month - 1, 1)
    return dict(nome='', technology='Wind', fase='costruzione', month=last_month.isoformat()[:7],
                app=1, sub=0, company_hse=1, hse_coverage=80., inspections=0,
                inspection_target=4, activities=['civili'], interference_level=3,
                nc_items=[], crit_items=[], awareness=[], stop_positive=0, stop_negative=0,
                imminent_danger=False, hours=None, workers=None, days=None, notes='')


def validate_data(d):
    errors = []
    if not isinstance(d, dict):
        return ['Valutazione non valida.']
    if not isinstance(d.get('nome'), str) or not d['nome'].strip() or len(d['nome']) > 120:
        errors.append('Inserire un nome cantiere (massimo 120 caratteri).')
    try:
        month = datetime.strptime(d.get('month', ''), '%Y-%m')
        if month.strftime('%Y-%m') != d['month'] or month.date() > date.today().replace(day=1):
            errors.append('Il mese deve essere valido e non futuro.')
    except (TypeError, ValueError):
        errors.append('Mese non valido: usare AAAA-MM.')
    if d.get('technology') not in ['Wind', 'Solar', 'BESS'] or d.get('fase') not in PHASE_THEMES:
        errors.append('Tecnologia o fase non valida.')
    ranges = {'app': (0, 200), 'sub': (0, 200), 'company_hse': (0, 200),
              'hse_coverage': (0, 100), 'inspections': (0, 2000), 'inspection_target': (1, 1000),
              'interference_level': (1, 10), 'stop_positive': (0, 1000), 'stop_negative': (0, 1000)}
    for key, (lo, hi) in ranges.items():
        v = d.get(key)
        if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or not lo <= v <= hi:
            errors.append(f'Valore non valido: {key}.')
        elif key not in ['hse_coverage'] and int(v) != v:
            errors.append(f'{key}: usare un numero intero.')
    for key, hi in [('hours', 10000000), ('workers', 100000), ('days', 31)]:
        v = d.get(key)
        if v is not None and (isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or not 0 < v <= hi):
            errors.append(f'{key}: inserire un valore positivo oppure lasciare il dato non disponibile.')
    if d.get('days') and isinstance(d['days'], (int, float)):
        try:
            if d['days'] > calendar.monthrange(month.year, month.month)[1]:
                errors.append('Giorni di attività superiori ai giorni del mese.')
        except UnboundLocalError:
            pass
    activities = d.get('activities', [])
    if not isinstance(activities, list) or not activities or any(a not in ACTIVITY for a in activities):
        errors.append('Selezionare almeno una attività valida.')
    elif len(activities) != len(set(activities)):
        errors.append('Attività duplicate.')
    for kind in ['nc_items', 'crit_items']:
        rows = d.get(kind)
        if not isinstance(rows, list) or len(rows) > 1000:
            errors.append(f'Elenco {kind} non valido (massimo 1000 righe).')
            continue
        for i, row in enumerate(rows):
            if not isinstance(row, dict):
                errors.append(f'{kind}, riga {i+1}: formato errato.')
                continue
            if row.get('theme') not in THEMES or row.get('severity') not in SEVERITY or row.get('status') not in STATUS:
                errors.append(f'{kind}, riga {i+1}: completare tema, gravità e stato.')
            for key in ['late', 'repeat']:
                if not isinstance(row.get(key), bool):
                    errors.append(f'{kind}, riga {i+1}: {key} non valido.')
            if row.get('late') and row.get('status') != 'Aperta':
                errors.append(f'{kind}, riga {i+1}: scaduta si usa solo per una anomalia aperta.')
            if not isinstance(row.get('note', ''), str) or len(row.get('note', '')) > 1000:
                errors.append(f'{kind}, riga {i+1}: nota troppo lunga.')
    rows = d.get('awareness')
    if not isinstance(rows, list) or len(rows) > 500:
        errors.append('Elenco sensibilizzazioni non valido.')
    else:
        topics = []
        for i, row in enumerate(rows):
            if not isinstance(row, dict):
                errors.append('Sensibilizzazione non valida.')
                continue
            topics.append(row.get('theme'))
            count, reach = row.get('count'), row.get('coverage')
            if row.get('theme') not in THEMES or isinstance(count, bool) or not isinstance(count, (int, float)) or not math.isfinite(count) or count < 1 or count > 1000 or int(count) != count:
                errors.append(f'Sensibilizzazione {i+1}: tema o numero incontri non valido.')
            if reach is not None and (isinstance(reach, bool) or not isinstance(reach, (int, float)) or not math.isfinite(reach) or not 0 <= reach <= 100):
                errors.append(f'Sensibilizzazione {i+1}: copertura deve essere tra 0 e 100 oppure vuota.')
        if len(topics) != len(set(topics)):
            errors.append('Usare una sola riga per tema di sensibilizzazione, sommando gli incontri.')
    if not isinstance(d.get('imminent_danger'), bool):
        errors.append('Indicazione pericolo attuale non valida.')
    if not isinstance(d.get('notes', ''), str) or len(d.get('notes', '')) > 5000:
        errors.append('Note generali non valide (massimo 5000 caratteri).')
    if not errors and d['app'] + d['sub'] == 0:
        errors.append('Indicare almeno una impresa operativa.')
    return list(dict.fromkeys(errors))


def anomaly_load(rows):
    return sum(SEVERITY[r['severity']] * THEME_FACTOR[r['theme']] * STATUS[r['status']]
               * (1.25 if r['late'] else 1) * (1.2 if r['repeat'] else 1) for r in rows)


def topic_needs(d):
    needs, reasons = {}, {}
    for activity in d['activities']:
        theme = ACTIVITY_THEME[activity]
        needs[theme] = max(needs.get(theme, 0), ACTIVITY[activity] / 100)
        reasons.setdefault(theme, []).append('attività: ' + activity)
    for field, label, factor in [('nc_items', 'NC formale', 1.), ('crit_items', 'criticità informale', .5)]:
        for row in d[field]:
            if row['status'] == 'Aperta':
                theme = row['theme']
                needs[theme] = needs.get(theme, 0) + SEVERITY[row['severity']] / 5 * factor
                reasons.setdefault(theme, []).append(f"{label} aperta, gravità {row['severity']}")
    for theme in PHASE_THEMES[d['fase']]:
        needs[theme] = max(needs.get(theme, 0), .2)
        reasons.setdefault(theme, []).append('coerenza con fase ' + d['fase'])
    return needs, reasons


def calculate_risk(d):
    """Pesi ipotetici di policy, versionati. Nessuna stima statistica degli infortuni.
    Indice base sempre confrontabile; contesto opzionale separato.
    """
    nc_load, crit_load = anomaly_load(d['nc_items']), anomaly_load(d['crit_items'])
    nc = 100 * (1 - math.exp(-nc_load / 10))
    crit = 100 * (1 - math.exp(-crit_load / 12))
    activity = min(100, max(ACTIVITY[a] for a in d['activities']) + 3 * (len(d['activities']) - 1))
    interference = (d['interference_level'] - 1) / 9 * 100
    org = 100 * (1 - math.exp(-(max(0, d['app'] - 1) + 1.2 * d['sub']) / 8))
    # Nessun requisito universale "un HSE per impresa": si valuta la copertura reale dichiarata.
    inspection_deficit = 100 * (1 - min(1, d['inspections'] / d['inspection_target']))
    controls = .65 * (100 - d['hse_coverage']) + .35 * inspection_deficit
    components = dict(zip(WEIGHTS, [nc, crit, activity, interference, org, controls]))
    drivers = [{'Area': key, 'Indice area': round(value, 3), 'Peso massimo': WEIGHTS[key],
                'Contributo': value * WEIGHTS[key] / 100} for key, value in components.items()]
    positive = min(2., d['stop_positive'] * .5)
    negative = min(6., d['stop_negative'] * 1.5)
    needs, reasons = topic_needs(d)
    numerator = 0.
    awareness_details = []
    for row in d['awareness']:
        theme = row['theme']
        coverage = row['coverage'] / 100 if row['coverage'] is not None else .5
        continuity = min(1., .7 + .1 * (row['count'] - 1))
        effectiveness = coverage * continuity
        contribution = needs.get(theme, 0) * effectiveness
        numerator += contribution
        awareness_details.append({'Tema': theme, 'Incontri': row['count'],
            'Copertura': f"{row['coverage']:.0f}%" if row['coverage'] is not None else 'Non nota: fattore 50%',
            'Pertinenza': round(needs.get(theme, 0), 3), 'Efficacia convenzionale': round(effectiveness, 3),
            'Motivo': '; '.join(reasons.get(theme, ['Nessun collegamento agli input: nessun bonus']))})
    awareness = 8 * numerator / sum(needs.values()) if needs else 0.
    raw = sum(x['Contributo'] for x in drivers) + negative - positive - awareness
    high_open = sum(r['severity'] == 'alta' and r['status'] == 'Aperta' for r in d['nc_items'])
    floor = 81. if d['imminent_danger'] or high_open >= 2 else 61. if high_open == 1 else 0.
    floor_reason = ('Pericolo grave e attuale dichiarato' if d['imminent_danger'] else
                    'Almeno due NC formali alte aperte' if high_open >= 2 else
                    'Una NC formale alta aperta' if high_open else 'Nessuna soglia minima attiva')
    score = max(floor, bounded(raw))
    # Esposizione solo nell'indice contestualizzato: ore > giorni > lavoratori.
    contextual_target, context_basis = None, 'Nessun dato di esposizione disponibile'
    if d.get('hours'):
        contextual_target, context_basis = max(1., d['hours'] / 2000), '1 ispezione ogni 2.000 ore lavorate'
    elif d.get('days'):
        contextual_target = max(1., d['days'] / 5 * max(1., (d.get('workers') or 50) / 50))
        context_basis = '1 ispezione ogni 5 giorni operativi e ogni 50 lavoratori medi (minimo fattore 1)'
    elif d.get('workers'):
        contextual_target, context_basis = max(1., 4 * d['workers'] / 50), '4 ispezioni per 50 lavoratori medi nel mese'
    contextual = None
    if contextual_target:
        context_controls = .65 * (100 - d['hse_coverage']) + .35 * 100 * (1 - min(1., d['inspections'] / contextual_target))
        contextual = max(floor, bounded(raw + (context_controls - controls) * .15))
    saturation = bounded(raw) - raw
    drivers.extend([{'Area': 'Stop Work negativi', 'Indice area': None, 'Peso massimo': 6, 'Contributo': negative},
                    {'Area': 'Stop Work positivi', 'Indice area': None, 'Peso massimo': -2, 'Contributo': -positive},
                    {'Area': 'Sensibilizzazioni pertinenti', 'Indice area': None, 'Peso massimo': -8, 'Contributo': -awareness},
                    {'Area': 'Limite scala 0–100', 'Indice area': None, 'Peso massimo': None, 'Contributo': saturation},
                    {'Area': 'Soglia minima di tutela', 'Indice area': None, 'Peso massimo': None, 'Contributo': score - bounded(raw)}])
    explanations = [
        f'NC formali: carico {nc_load:.2f}; indice area = 100 × (1 − exp(−carico/10)); massimo 30 punti.',
        f'Criticità informali: carico {crit_load:.2f}; indice area = 100 × (1 − exp(−carico/12)); massimo 8 punti.',
        'Carico anomalia = gravità (1/2,5/5) × tema (0,65–1,25) × stato (aperta 1; chiusa nei tempi 0,2; chiusa in ritardo 0,4) × scaduta (1,25) × recidiva (1,2).',
        'I temi modificano la priorità convenzionale; la gravità deve essere assegnata in base al potenziale danno del caso concreto.',
        'Attività: massima pericolosità convenzionale + 3 per ogni tipologia aggiuntiva, massimo 100; peso 22%.',
        'Interferenze: (livello − 1)/9 × 100, peso 15%. Non viene riapplicata la pericolosità delle attività.',
        'Organizzazione: 100 × (1 − exp(−(appaltatori−1 + 1,2×subappaltatori)/8)), peso 10%.',
        f'Presidio: 65% deficit copertura HSE + 35% deficit ispezioni rispetto al target dichiarato ({d["inspection_target"]}); peso 15%.',
        'Nessun malus sul rapporto NC/ispezioni: rilevare più anomalie non rende meno efficace il controllo. Il rapporto resta descrittivo.',
        f'Sensibilizzazioni: −{awareness:.2f}/8. Copertura ponderata dei temi necessari; continuità da 0,7 a 1; copertura non nota usa fattore prudenziale 0,5.',
        f'Stop Work: positivi −{positive:.2f}/2; negativi +{negative:.2f}/6. Ogni evento va classificato una sola volta; anomalia sottostante registrata una sola volta.',
        f'Soglia minima: {floor:.0f}. {floor_reason}. Le criticità informali da sole non attivano una soglia.',
        'I pesi più alti delle NC formali sono una scelta gestionale richiesta, non una dimostrazione di maggiore pericolosità fisica.',
        'Pesi, soglie e target sono ipotesi iniziali di calibrazione; non derivano da dati storici né da un modello predittivo validato.'
    ]
    missing = []
    if not d['inspections']:
        missing.append('Nessuna ispezione: assenza di NC non dimostra assenza di anomalie.')
    if any(d.get(k) is None for k in ['hours', 'workers', 'days']):
        missing.append('Dati di esposizione parziali o assenti: usare l’indice base per la graduatoria comune.')
    if any(r['coverage'] is None for r in d['awareness']):
        missing.append('Copertura delle sensibilizzazioni non sempre nota.')
    return dict(version=MODEL_VERSION, score=round(score, 2), level=risk_level(score),
                raw=round(raw, 3), floor=floor, floor_reason=floor_reason,
                drivers=drivers, awareness_details=awareness_details, awareness_bonus=round(awareness, 3),
                needs=needs, reasons=reasons, context_score=round(contextual, 2) if contextual is not None else None,
                context_target=round(contextual_target, 2) if contextual_target else None, context_basis=context_basis,
                explanations=explanations, missing=missing)


TOOLBOX_CONTENT = {
 'elettrico': ('Contatti elettrici, arco elettrico e riattivazione inattesa',
  ['Verificare stato dell’impianto, confini dell’area e responsabili operativi prima di iniziare.',
   'Applicare la procedura di messa in sicurezza prevista e verificare le condizioni per la consegna del lavoro.',
   'Controllare utensili, DPI previsti e separazione dalle parti che restano in servizio.'],
  'Chi conferma le condizioni di sicurezza e cosa fate se lo stato dell’impianto cambia?'),
 'scavi': ('Intercettazione sottoservizi, cedimenti e accessi allo scavo',
  ['Confrontare rilievi, elaborati e condizioni reali; chiarire i sottoservizi prima dello scavo.',
   'Verificare stabilità, protezioni dei bordi, accessi e distanza di mezzi/materiali secondo il piano operativo.',
   'Fermare la lavorazione in caso di sottoservizio inatteso o variazione delle condizioni.'],
  'Come riconoscete un sottoservizio non censito e chi autorizza la ripresa?'),
 'quota': ('Caduta dall’alto e caduta di materiali',
  ['Verificare accessi e protezioni collettive previste; controllare integrità e corretto utilizzo.',
   'Dove previsti, verificare sistemi anticaduta, compatibilità e piano di recupero.',
   'Delimitare le aree sottostanti e controllare vento e altre condizioni meteo.'],
  'Come viene recuperata una persona sospesa e chi attiva il piano?'),
 'sollevamento': ('Carichi sospesi, interferenze e stabilità del mezzo',
  ['Condividere piano di sollevamento, ruoli, segnali e confini dell’area.',
   'Verificare accessori, portata nella configurazione reale, appoggi e condizioni meteo.',
   'Impedire presenza di persone sotto carichi e nell’area di movimento.'],
  'Qual è il segnale di arresto e chi lo può impartire?'),
 'spazi confinati': ('Atmosfere pericolose, accesso e soccorso',
  ['Identificare e valutare l’ambiente prima dell’accesso; applicare la procedura dedicata.',
   'Verificare isolamento, monitoraggi, ventilazione, comunicazioni e presidio esterno previsti.',
   'Controllare organizzazione e mezzi di recupero; non improvvisare ingressi di soccorso.'],
  'Quali condizioni impediscono l’accesso e come viene attivato il recupero?'),
 'incendio': ('Innesco, propagazione e gestione emergenze',
  ['Verificare materiali combustibili, fonti d’innesco e autorizzazioni operative previste.',
   'Identificare allarmi, vie di esodo, addetti e presidi; per BESS seguire il piano emergenze specifico.',
   'Rispettare esclusioni e sorveglianza prevista per hot work; segnalare anomalie immediatamente.'],
  'Dove vi radunate e quali informazioni date ai soccorsi?'),
 'viabilità': ('Investimento, retromarcia e interazione mezzi-pedoni',
  ['Condividere percorsi, ingressi, sensi di marcia e punti di attraversamento.',
   'Verificare separazione pedoni/mezzi, visibilità e assistenza alle manovre ove prevista.',
   'Adeguare la velocità alle condizioni reali e fermarsi se il percorso non è libero.'],
  'Chi verifica che l’area di manovra sia libera?'),
 'dpi': ('Scelta, compatibilità e uso dei DPI',
  ['Associare i DPI al rischio e alla lavorazione effettiva, secondo la valutazione applicabile.',
   'Controllare integrità, taglia, compatibilità e addestramento richiesto.',
   'Segnalare immediatamente DPI mancanti, usurati o non idonei.'],
  'Quale difetto vi porta a sostituire il DPI prima di iniziare?'),
 'ordine e pulizia': ('Inciampo, accessibilità e gestione materiali',
  ['Mantenere sgombri passaggi, vie di esodo e accesso ai presidi.',
   'Organizzare deposito materiali, cavi e rifiuti nelle aree previste.',
   'Rimuovere o mettere in sicurezza sversamenti e ostacoli prima di proseguire.'],
  'Chi gestisce il ripristino se trovate un passaggio ostruito?'),
 'documentale': ('Coerenza fra pianificazione e lavoro reale',
  ['Verificare disponibilità e aggiornamento dei documenti applicabili alla lavorazione.',
   'Confrontare personale, mezzi e sequenza operativa con quanto autorizzato.',
   'Gestire le variazioni prima dell’avvio e chiarire i ruoli del controllo operativo.'],
  'Quale variazione richiede una nuova verifica prima di lavorare?'),
 'ambientale': ('Sversamenti, rifiuti e protezione del suolo',
  ['Identificare sostanze, punti di travaso e aree sensibili.',
   'Verificare contenimenti, kit di intervento e segregazione dei rifiuti.',
   'In caso di perdita fermare la fonte se sicuro, delimitare e attivare il flusso di segnalazione.'],
  'Quali materiali sono disponibili per contenere una perdita?'),
 'altro': ('Riconoscimento del pericolo e gestione delle variazioni',
  ['Descrivere la sequenza di lavoro e identificare i pericoli specifici.',
   'Confermare protezioni, ruoli e comunicazioni prima dell’avvio.',
   'Interrompere in caso di condizioni diverse da quelle valutate.'],
  'Quale cambiamento rende necessario rivalutare il lavoro?')}


def toolbox_text(theme, d, result):
    title, points, question = TOOLBOX_CONTENT[theme]
    reason = '; '.join(result['reasons'].get(theme, ['Tema scelto dall’utilizzatore']))
    return (f"Sensibilizzazione — {theme.upper()}\nCantiere: {d['nome']} | {d['technology']} | Mese: {d['month']}\n"
            f"Durata proposta: 5–10 minuti. Destinatari: lavoratori e preposti interessati.\n\n"
            f"Perché questo tema: {reason}.\n\nObiettivo: {title}.\n\n"
            'Apertura del preposto: «Prima di iniziare ripercorriamo i rischi di questa attività e le misure previste sul nostro sito».\n\n'
            + '\n'.join(f'{i+1}. {p}' for i, p in enumerate(points))
            + f'\n\nVerifica di comprensione: {question}\n'
            'Chiusura: ogni lavoratore indica un pericolo e la misura da applicare; il preposto verifica sul campo.\n'
            'Registrare data, argomento, partecipanti unici, osservazioni e azioni emerse.\n'
            'Adattare il testo a PSC/POS, procedure e condizioni reali; il toolbox non sostituisce formazione e addestramento.')


def action_plan(d, r):
    rows = []
    def add(priority, area, action, owner, when, evidence):
        rows.append(dict(Priorità=priority, Area=area, Azione=action, Responsabile=owner,
                         Termine=when, Evidenza=evidence, Stato='Da pianificare'))
    if d['imminent_danger']:
        add('P1', 'Pericolo attuale', 'Mettere in sicurezza la situazione e attivare i responsabili secondo le procedure di sito; verificare le condizioni per la ripresa.',
            'Responsabili operativi competenti', 'Immediato', 'Verifica delle misure prima della ripresa')
    for field, label in [('nc_items', 'NC formale'), ('crit_items', 'Criticità informale')]:
        for i, item in enumerate(d[field]):
            if item['status'] == 'Aperta':
                add('P1' if item['severity'] == 'alta' or item['late'] else 'P2', item['theme'],
                    f"{label} {i+1}: definire e attuare la correzione specifica. {item.get('note', '')}".strip(),
                    'Responsabile impresa / preposto', 'Prima della lavorazione esposta' if item['severity'] == 'alta' else 'Entro 7 giorni (proposta)',
                    'Chiusura documentata e verifica di efficacia')
    for driver in sorted(r['drivers'][:6], key=lambda x: x['Contributo'], reverse=True):
        area = driver['Area']
        if driver['Contributo'] < 2:
            continue
        suggestions = {'Attività': ('Rivedere sequenze, mezzi e protezioni delle lavorazioni a maggior potenziale; verificare sul campo le misure.', 'Referenti operativi / preposti'),
            'Interferenze': ('Separare nel tempo o nello spazio le attività interferenti; aggiornare layout e coordinamento.', 'Referenti imprese / coordinamento di cantiere'),
            'Organizzazione': ('Chiarire responsabilità, interfacce fra imprese e supervisione delle aree operative.', 'Direzione cantiere'),
            'Presidio e controlli': ('Rivedere copertura reale delle lavorazioni e piano ispezioni; verificarne qualità e attuazione.', 'Referente HSE / direzione cantiere')}
        if area in suggestions:
            action, owner = suggestions[area]
            add('P2', area, action, owner, 'Entro 7 giorni (proposta)', 'Pianificazione condivisa e riscontro in campo')
    awareness_map = {x['theme']: x for x in d['awareness']}
    for theme, need in sorted(r['needs'].items(), key=lambda x: x[1], reverse=True):
        existing = awareness_map.get(theme)
        if need >= .5 and (not existing or existing['coverage'] is None or existing['coverage'] < 80):
            add('P2', 'Sensibilizzazione: ' + theme,
                'Organizzare un toolbox su ' + theme + '; collegamento: ' + '; '.join(r['reasons'][theme]),
                'Preposto con supporto HSE', 'Prima della prossima attività pertinente', 'Partecipanti unici e verifica comprensione')
    if not d['inspections']:
        add('P2', 'Conoscenza del sito', 'Eseguire una prima verifica strutturata; il numero di anomalie zero non equivale a controllo efficace.',
            'Referente HSE', 'Entro 7 giorni (proposta)', 'Report ispezione')
    return rows


def executive_summary(d, r):
    top = sorted(r['drivers'][:6], key=lambda x: x['Contributo'], reverse=True)[:3]
    nc_counts, crit_counts = anomaly_counts(d['nc_items']), anomaly_counts(d['crit_items'])
    text = (f"{d['nome']} — {d['technology']}, {d['month']}: indice base {r['score']:.1f}/100 ({r['level']}). "
            'I contributi principali sono ' + ', '.join(f"{x['Area']} ({x['Contributo']:.1f} punti)" for x in top) + '. ')
    if r['floor']:
        text += f"È attiva la soglia minima {r['floor']:.0f}: {r['floor_reason']}. "
    text += (f"NC formali: {nc_counts['total']} totali, {nc_counts['open']} aperte, {nc_counts['closed']} chiuse. "
             f"Criticità informali: {crit_counts['total']} totali, {crit_counts['open']} aperte, {crit_counts['closed']} chiuse. ")
    text += f"Le sensibilizzazioni pertinenti riducono l’indice di {r['awareness_bonus']:.1f} punti. "
    if r['missing']:
        text += 'Qualità del dato: ' + ' '.join(r['missing'])
    return text


def simulations(d, r):
    """Ricalcolo del medesimo modello: non previsione della riduzione reale del rischio."""
    rows = []
    changes = [
      ('Chiusura verificata delle NC aperte', lambda x: [item.update(status='Chiusa nei tempi', late=False) for item in x['nc_items'] if item['status'] == 'Aperta']),
      ('Chiusura verificata delle criticità aperte', lambda x: [item.update(status='Chiusa nei tempi', late=False) for item in x['crit_items'] if item['status'] == 'Aperta']),
      ('Copertura HSE al 100%', lambda x: x.update(hse_coverage=100)),
      ('Target ispettivo raggiunto', lambda x: x.update(inspections=max(x['inspections'], x['inspection_target']))),
      ('Interferenza ridotta di 2 livelli', lambda x: x.update(interference_level=max(1, x['interference_level'] - 2))),
      ('Toolbox pertinenti: 4 incontri per tema, copertura 100%', lambda x: x.update(awareness=[{'theme': t, 'count': 4, 'coverage': 100.} for t in topic_needs(x)[0]]))]
    for label, apply in changes:
        sample = deepcopy(d)
        apply(sample)
        result = calculate_risk(sample)
        rows.append({'Scenario indipendente': label, 'Indice simulato': result['score'],
                     'Riduzione punti': round(r['score'] - result['score'], 2),
                     'Soglia residua': result['floor']})
    return rows


def local_advisor(question, d, r):
    """Assistente deterministico: non è un LLM e non risponde a quesiti normativi."""
    q = question.casefold()
    if any(t in q for t in ['toolbox', 'sensibil', 'formaz']):
        themes = sorted(r['needs'], key=r['needs'].get, reverse=True)[:3]
        return 'Temi prioritari: ' + ', '.join(themes) + '.\n\n' + '\n\n'.join(
            f"**{t}**: " + '; '.join(r['reasons'][t]) for t in themes)
    if any(t in q for t in ['ridur', 'interven', 'miglior', 'azion']):
        return '\n\n'.join(f"**{x['Area']}** — {x['Azione']}" for x in action_plan(d, r)[:8]) or 'Nessuna azione specifica generata: mantenere e verificare le misure previste.'
    if any(t in q for t in ['hse', 'presidio', 'ispez']):
        return (f"Copertura HSE dichiarata: {d['hse_coverage']:.0f}%. Ispezioni: {d['inspections']} / target {d['inspection_target']}. "
                'Il numero di HSE non determina automaticamente adeguatezza: valutare presenza, competenze, turni e aree effettivamente seguite.')
    return executive_summary(d, r) + '\n\nSono un assistente a regole: posso spiegare indice, azioni, presidio e temi dei toolbox; non interpreto liberamente domande normative.'


def setting(key, default=''):
    try:
        return st.secrets.get(key, os.environ.get('HSE_' + key.upper(), default))
    except Exception:
        return os.environ.get('HSE_' + key.upper(), default)


class HistoryStore:
    """Sessione senza configurazione; SQLite locale o PostgreSQL opzionali.
    Il database condiviso è accessibile solo dopo il gate applicativo in main().
    Non si salvano risultati fidati da import: vengono ricalcolati dai dati validati.
    """
    def __init__(self, mode='session', url='', path='hse_history.sqlite3'):
        self.mode, self.url, self.path = mode, url, path
        if mode != 'session':
            with self.connect() as conn:
                conn.execute('CREATE TABLE IF NOT EXISTS hse_evaluations (id TEXT PRIMARY KEY, payload TEXT NOT NULL)')

    def connect(self):
        if self.mode == 'postgres':
            import psycopg
            return psycopg.connect(self.url, connect_timeout=10)
        return sqlite3.connect(self.path, timeout=20)

    def all(self):
        if self.mode == 'session':
            return deepcopy(st.session_state.setdefault('history_records', []))
        with self.connect() as conn:
            return [json.loads(row[0]) for row in conn.execute('SELECT payload FROM hse_evaluations').fetchall()]

    def save_many(self, records):
        if self.mode == 'session':
            by_id = {x['id']: x for x in self.all()}
            for record in records:
                by_id.setdefault(record['id'], record)
            st.session_state.history_records = list(by_id.values())
            return
        placeholder = '%s' if self.mode == 'postgres' else '?'
        sql = f'INSERT INTO hse_evaluations (id, payload) VALUES ({placeholder}, {placeholder}) ON CONFLICT(id) DO NOTHING'
        with self.connect() as conn:
            for record in records:
                conn.execute(sql, (record['id'], json.dumps(record, ensure_ascii=False, allow_nan=False)))


def new_record(d, r):
    return dict(id=str(uuid.uuid4()), saved_at=datetime.now(timezone.utc).isoformat(),
                model_version=MODEL_VERSION, data=deepcopy(d), result=deepcopy(r))


def backup_bytes(records):
    return json.dumps(dict(schema_version=SCHEMA_VERSION, records=records), ensure_ascii=False, indent=2, allow_nan=False).encode('utf-8')


def parse_backup(raw):
    if len(raw) > 20 * 1024 * 1024:
        raise ValueError('Backup troppo grande: massimo 20 MB.')
    content = json.loads(raw)
    if not isinstance(content, dict) or content.get('schema_version') != SCHEMA_VERSION or not isinstance(content.get('records'), list):
        raise ValueError('Formato backup non riconosciuto (schema 2 richiesto).')
    if len(content['records']) > 10000:
        raise ValueError('Massimo 10.000 valutazioni per importazione.')
    records = []
    seen = set()
    for i, old in enumerate(content['records']):
        if not isinstance(old, dict):
            raise ValueError(f'Record {i+1} non valido.')
        try:
            ident = str(uuid.UUID(old['id']))
            timestamp = datetime.fromisoformat(old['saved_at'])
            if timestamp.tzinfo is None:
                timestamp = timestamp.replace(tzinfo=timezone.utc)
            timestamp = timestamp.astimezone(timezone.utc).isoformat()
        except (KeyError, ValueError, TypeError, AttributeError):
            raise ValueError(f'Identificativo/data non valido al record {i+1}.')
        if ident in seen:
            raise ValueError('Identificativi duplicati nel backup.')
        seen.add(ident)
        errors = validate_data(old.get('data'))
        if errors:
            raise ValueError(f'Record {i+1}: ' + '; '.join(errors[:5]))
        if old.get('model_version') != MODEL_VERSION:
            raise ValueError(f'Record {i+1}: versione modello non supportata; importare con la versione originale.')
        d = old['data']
        records.append(dict(id=ident, saved_at=timestamp, model_version=MODEL_VERSION,
                            data=d, result=calculate_risk(d)))
    return records


def history_frame(records, latest_only=True):
    rows = [dict(id=x['id'], Cantiere=x['data']['nome'], site=site_key(x['data']['nome']),
                 Tecnologia=x['data']['technology'], Mese=x['data']['month'],
                 Indice=x['result']['score'], Livello=x['result']['level'],
                 Contestualizzato=x['result']['context_score'], Salvato=x['saved_at'], Versione=x['model_version']) for x in records]
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    frame = frame.sort_values(['Mese', 'Salvato', 'id'])
    if latest_only:
        frame = frame.drop_duplicates(['site', 'Mese', 'Versione'], keep='last')
    return frame


def word_report(d, r, theme=None):
    doc = Document()
    doc.add_heading('HSE Risk Platform | ' + ('Toolbox Talk' if theme else 'Report mensile'), 0)
    doc.add_paragraph(f"{d['nome']} · {d['technology']} · {d['month']} · Modello {MODEL_VERSION}")
    if theme:
        for line in toolbox_text(theme, d, r).splitlines():
            doc.add_paragraph(line)
    else:
        doc.add_paragraph(executive_summary(d, r))
        for field, title in [('nc_items', 'NC formali'), ('crit_items', 'Criticità informali')]:
            doc.add_heading(title, 1)
            counts = anomaly_counts(d[field])
            doc.add_paragraph(f"Totale {counts['total']} | Aperte {counts['open']} | Chiuse {counts['closed']}")
            for event in d[field]:
                doc.add_paragraph(f"{event['theme']} | {event['severity']} | {event['status']} | {event.get('note', '')}")
        doc.add_heading('Motivazioni e metodo', 1)
        for line in r['explanations']:
            doc.add_paragraph(line)
        doc.add_heading('Azioni proposte', 1)
        for row in action_plan(d, r):
            doc.add_paragraph(f"{row['Priorità']} | {row['Area']} — {row['Azione']} | {row['Responsabile']} | {row['Termine']}")
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def excel_report(d, r, records):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        data_tables = {
          'Sintesi': pd.DataFrame([{'Cantiere': d['nome'], 'Tecnologia': d['technology'], 'Mese': d['month'],
             'Indice base': r['score'], 'Livello': r['level'], 'Indice contestualizzato': r['context_score'], 'Versione modello': MODEL_VERSION, 'NC totali': len(d['nc_items']), 'NC aperte': anomaly_counts(d['nc_items'])['open'], 'NC chiuse': anomaly_counts(d['nc_items'])['closed']}]),
          'Contributi': pd.DataFrame(r['drivers']), 'Action Plan': pd.DataFrame(action_plan(d, r)),
          'NC formali': pd.DataFrame(d['nc_items']), 'Criticità informali': pd.DataFrame(d['crit_items']),
          'Sensibilizzazioni': pd.DataFrame(r['awareness_details']),
          'Metodo': pd.DataFrame({'Spiegazione': r['explanations']}),
          'Scenari': pd.DataFrame(simulations(d, r)), 'Storico': history_frame(records),
          'Input': pd.DataFrame([{'Campo': k, 'Valore': json.dumps(v, ensure_ascii=False)} for k, v in d.items()])}
        from openpyxl.styles import PatternFill, Font, Alignment
        for title, frame in data_tables.items():
            frame.to_excel(writer, sheet_name=title, index=False)
            sheet = writer.sheets[title]
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.fill = PatternFill('solid', fgColor='123047')
                cell.font = Font(color='FFFFFF', bold=True)
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    # Non eseguire testo utente come formule Excel.
                    if isinstance(cell.value, str) and cell.value.startswith(('=', '+', '-', '@')):
                        cell.data_type = 's'
                    cell.alignment = Alignment(vertical='top', wrap_text=True)
            for col in sheet.columns:
                sheet.column_dimensions[col[0].column_letter].width = min(65, max(15, max(len(str(c.value or '')) for c in col[:60]) + 2))
            sheet.sheet_view.showGridLines = False
        from openpyxl.chart import BarChart, Reference
        sheet = writer.sheets['Contributi']
        chart = BarChart()
        chart.title = 'Contributi all’indice base'
        chart.add_data(Reference(sheet, min_col=4, min_row=1, max_row=sheet.max_row), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=sheet.max_row))
        sheet.add_chart(chart, 'F2')
    return buffer.getvalue()


def ppt_report(d, r):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    pages = [('HSE Risk Platform', f"{d['nome']} | {d['technology']} | {d['month']}\nIndice {r['score']:.1f}/100 — {r['level']}\nModello gestionale {MODEL_VERSION}"),
             ('Sintesi del mese', executive_summary(d, r)),
             ('Contributi al punteggio', '\n'.join(f"{x['Area']}: {x['Contributo']:+.1f} punti" for x in r['drivers'])),
             ('Priorità operative', '\n'.join(f"{x['Priorità']} · {x['Area']}: {x['Azione']}" for x in action_plan(d, r)[:5])),
             ('Metodo e limiti', '\n'.join(r['explanations'][-3:]) + '\nIndice base per confronto; contestualizzato separato.\nNessuna validazione predittiva disponibile.')]
    for theme in sorted(r['needs'], key=r['needs'].get, reverse=True)[:3]:
        title, points, question = TOOLBOX_CONTENT[theme]
        pages.append(('Toolbox | ' + theme, title + '\n\n' + '\n'.join(points) + '\n\n' + question))
    for field, title in [('nc_items', 'Registro NC formali'), ('crit_items', 'Registro criticità informali')]:
        counts = anomaly_counts(d[field])
        body = f"Totale {counts['total']} · Aperte {counts['open']} · Chiuse {counts['closed']}\n" + '\n'.join(f"{e['theme']} | {e['severity']} | {e['status']}: {e.get('note', '')}" for e in d[field])
        pages.append((title, body))
    paginated = []
    for title, body in pages:
        lines = []
        for paragraph in body.splitlines():
            lines.extend(textwrap.wrap(paragraph, width=102) or [''])
        for start in range(0, len(lines), 15):
            paginated.append((title + (' · continua' if start else ''), '\n'.join(lines[start:start+15])))
    for title, body in paginated:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string('F3F6FA')
        for top, height, text, size, color in [(0.55, 1., title, 30, '123047'), (1.7, 5., body, 18, '254559')]:
            box = slide.shapes.add_textbox(Inches(.7), Inches(top), Inches(11.9), Inches(height))
            tf = box.text_frame
            tf.word_wrap = True
            for i, line in enumerate(text.splitlines()):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = line
                p.font.size = PptPt(size)
                p.font.color.rgb = RGBColor.from_string(color)
                p.space_after = PptPt(10)
    buffer = BytesIO()
    prs.save(buffer)
    return buffer.getvalue()


CHECKLIST_LIBRARY = {
    "Controlli generali pre-attività": [
        "L'attività è inserita nel programma lavori / sinottico aggiornato?",
        "È stato eseguito il coordinamento con le imprese coinvolte e interferenti?",
        "Sono stati individuati il preposto e i referenti operativi dell'attività?",
        "La documentazione operativa è disponibile, aggiornata e coerente con l'attività?",
        "Sono state verificate formazione, informazione, addestramento e idoneità dei lavoratori?",
        "Sono disponibili attrezzature, mezzi e DPI idonei e verificati?",
        "L'area di lavoro è stata delimitata, segnalata e resa accessibile in sicurezza?",
        "Sono state valutate le condizioni ambientali e meteorologiche?",
        "Sono state definite le modalità di comunicazione e gestione delle emergenze?",
        "È stato eseguito il briefing / pre-job check prima dell'avvio?",
    ],
    "Sollevamenti": [
        "È stato predisposto e condiviso il piano di sollevamento?",
        "Sono stati verificati peso, baricentro e punti di presa del carico?",
        "La gru o l'apparecchio di sollevamento è idoneo e correttamente posizionato?",
        "Sono disponibili verifiche, controlli e documentazione dell'attrezzatura?",
        "Gruista, imbracatore e segnalatore possiedono formazione e abilitazioni richieste?",
        "Gli accessori di sollevamento sono identificati, integri e adeguati alla portata?",
        "L'area di manovra e la zona sottostante sono segregate?",
        "Sono state verificate interferenze con linee elettriche, strutture e altre attività?",
        "Sono definite modalità di comunicazione univoche tra gli operatori?",
        "Sono rispettati i limiti meteo e di vento previsti?",
    ],
    "Lavori elettrici": [
        "Sono definiti ruoli, responsabilità e autorizzazioni per il lavoro elettrico?",
        "Sono verificate le competenze PES, PAV, PEI o equivalenti richieste?",
        "È disponibile il permesso / piano di lavoro applicabile?",
        "Sono state eseguite sezionamento, blocco, identificazione e prevenzione delle richiusure?",
        "È stata verificata l'assenza di tensione con strumento idoneo?",
        "Sono state applicate eventuali terre e cortocircuiti previste?",
        "Sono definite e rispettate le distanze di sicurezza dalle parti attive?",
        "DPI, attrezzi isolati e strumenti sono idonei e verificati?",
        "Sono gestite le interferenze con impianti o stalli adiacenti in esercizio?",
        "Sono definite le procedure di emergenza e soccorso elettrico?",
    ],
    "Scavi": [
        "Sono stati individuati e tracciati i sottoservizi presenti?",
        "È disponibile l'autorizzazione allo scavo e la documentazione prevista?",
        "Sono state valutate stabilità dei fronti, profondità e caratteristiche del terreno?",
        "Sono previste armature, pendenze o sistemi di protezione adeguati?",
        "Sono presenti accessi e uscite sicure dallo scavo?",
        "Materiali e mezzi sono mantenuti a distanza sicura dal ciglio?",
        "Lo scavo è delimitato e segnalato anche nelle ore notturne?",
        "Sono gestite acqua, cedimenti, atmosfere pericolose e condizioni meteo?",
        "Sono impediti accessi e soste nel raggio operativo dei mezzi?",
        "È prevista una verifica prima di ogni turno e dopo eventi che possano alterare lo scavo?",
    ],
    "Lavori in quota": [
        "È stata privilegiata una protezione collettiva rispetto ai DPI anticaduta?",
        "Ponteggi, parapetti, trabattelli o opere provvisionali sono conformi e verificati?",
        "Sono individuati e verificati i punti di ancoraggio?",
        "Imbracature, cordini e dispositivi anticaduta sono integri e idonei?",
        "I lavoratori hanno formazione e addestramento specifici?",
        "Sono gestiti caduta materiali e segregazione dell'area sottostante?",
        "Scale e accessi sono utilizzati esclusivamente nelle condizioni consentite?",
        "Sono considerate condizioni meteo e vento?",
        "È disponibile un piano di recupero e soccorso dell'operatore sospeso?",
        "Sono state verificate interferenze con impianti elettrici e altre lavorazioni?",
    ],
    "PLE": [
        "La PLE è idonea alla quota, allo sbraccio e alle condizioni del terreno?",
        "Sono disponibili verifiche periodiche, controlli e manuale d'uso?",
        "L'operatore è abilitato e autorizzato all'utilizzo?",
        "È stata eseguita la verifica pre-uso?",
        "Stabilizzatori e piastre sono correttamente posizionati?",
        "Sono rispettati i limiti di portata e numero di occupanti?",
        "Sono gestite distanze da linee elettriche e ostacoli?",
        "L'area sottostante e il raggio di manovra sono segregati?",
        "È previsto l'utilizzo dei DPI anticaduta indicati dal costruttore?",
        "È noto il sistema di discesa di emergenza e presente personale addestrato a terra?",
    ],
    "Hot Work": [
        "È stato emesso il permesso di lavoro a caldo?",
        "Sono stati rimossi o protetti i materiali combustibili?",
        "Sono disponibili estintori idonei e facilmente raggiungibili?",
        "È stato nominato il fire watch quando necessario?",
        "Bombole, tubazioni e riduttori sono integri e correttamente posizionati?",
        "È garantita adeguata ventilazione?",
        "Sono controllate scintille, scorie e propagazione del calore verso aree adiacenti?",
        "Sono gestite atmosfere potenzialmente esplosive o infiammabili?",
        "I lavoratori utilizzano DPI idonei a calore, radiazioni e fumi?",
        "È previsto il controllo dell'area al termine dell'attività?",
    ],
    "Spazi confinati": [
        "L'ambiente è stato classificato e autorizzato come spazio confinato / sospetto di inquinamento?",
        "È disponibile una procedura specifica e un permesso di ingresso?",
        "Sono stati isolati impianti, energie e possibili ingressi di sostanze?",
        "Sono state eseguite misure atmosferiche prima e durante l'accesso?",
        "Sono garantite ventilazione e condizioni respirabili?",
        "Gli addetti possiedono formazione e addestramento specifici?",
        "È presente un sorvegliante esterno dedicato?",
        "Sono disponibili comunicazioni efficaci tra interno ed esterno?",
        "È predisposto un piano di emergenza con attrezzature di recupero?",
        "È impedito l'accesso a personale non autorizzato?",
    ],
    "Movimentazione mezzi": [
        "È definito e segnalato il piano di viabilità del cantiere?",
        "Sono separati, ove possibile, percorsi pedonali e carrabili?",
        "Mezzi e operatori sono autorizzati e in possesso dei requisiti previsti?",
        "Sono stati eseguiti i controlli pre-uso dei mezzi?",
        "Sono funzionanti segnalatori acustici, luminosi e dispositivi di visibilità?",
        "Sono individuati punti ciechi e necessità di moviere?",
        "Velocità e sensi di marcia sono definiti e rispettati?",
        "Le aree di carico, scarico e manovra sono segregate?",
        "Sono gestite interferenze con scavi, sollevamenti e lavorazioni a terra?",
        "Le condizioni del fondo e delle piste sono adeguate?",
    ],
    "Attività civili": [
        "Le aree di lavoro sono ordinate, pulite e prive di ostacoli?",
        "Sono verificati accessi, percorsi e protezioni dei bordi?",
        "Attrezzature manuali ed elettriche sono idonee e controllate?",
        "Sono gestite polveri, rumore e proiezione di materiali?",
        "Sono previste modalità sicure per taglio, perforazione e demolizioni minori?",
        "Sono stoccati correttamente materiali e prodotti?",
        "Sono controllate le interferenze con impianti esistenti e sottoservizi?",
        "Sono disponibili i DPI specifici per le attività?",
    ],
    "Attività meccaniche": [
        "Macchine e componenti sono stabilizzati prima delle lavorazioni?",
        "Sono isolate e scaricate le energie residue?",
        "Attrezzi e attrezzature sono adatti e in buono stato?",
        "Sono gestiti schiacciamento, cesoiamento e intrappolamento?",
        "Sono previsti sistemi sicuri per movimentazione e posizionamento dei componenti?",
        "Sono controllati fluidi in pressione e temperature residue?",
        "Sono definite le modalità di prova e rimessa in servizio?",
        "L'area è segregata durante prove, avviamenti e movimentazioni?",
    ],
}

ACTIVITY_TO_CHECKLIST = {
    "elettrici": "Lavori elettrici",
    "scavi": "Scavi",
    "lavori in quota": "Lavori in quota",
    "sollevamenti": "Sollevamenti",
    "spazi confinati": "Spazi confinati",
    "hot work": "Hot Work",
    "movimentazione mezzi": "Movimentazione mezzi",
    "civili": "Attività civili",
    "meccanici": "Attività meccaniche",
}


def generate_checklist_word(data, selected_sections):
    document = Document()
    section = document.sections[0]
    section.top_margin = DocxInches(0.65)
    section.bottom_margin = DocxInches(0.65)
    section.left_margin = DocxInches(0.7)
    section.right_margin = DocxInches(0.7)

    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("CHECKLIST HSE PRE-ATTIVITÀ")
    run.bold = True
    run.font.size = Pt(16)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("Verifica preventiva da compilare a cura dell'HSE di cantiere").italic = True

    info_table = document.add_table(rows=3, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.style = "Table Grid"
    info_values = [
        ("Cantiere", data.get("nome", ""), "Data", datetime.now().strftime("%d/%m/%Y")),
        ("Fase", data.get("fase", ""), "Impresa", ""),
        ("Compilata da", "", "Attività / area", ""),
    ]
    for row, values in zip(info_table.rows, info_values):
        for index, value in enumerate(values):
            row.cells[index].text = str(value)
            row.cells[index].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if index in (0, 2):
                row.cells[index].paragraphs[0].runs[0].bold = True

    document.add_paragraph()
    legend = document.add_paragraph()
    legend.add_run("Modalità di compilazione: ").bold = True
    legend.add_run("barrare SÌ, NO oppure N.A. e riportare nelle note le azioni necessarie.")

    for section_name in selected_sections:
        questions = CHECKLIST_LIBRARY.get(section_name, [])
        if not questions:
            continue

        heading = document.add_paragraph()
        heading.paragraph_format.space_before = Pt(10)
        heading.paragraph_format.space_after = Pt(4)
        run = heading.add_run(section_name.upper())
        run.bold = True
        run.font.size = Pt(11)

        table = document.add_table(rows=1, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        headers = ["N.", "Verifica", "SÌ", "NO", "N.A.", "Note / azioni"]
        for index, header in enumerate(headers):
            cell = table.rows[0].cells[index]
            cell.text = header
            cell.paragraphs[0].runs[0].bold = True
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        for number, question in enumerate(questions, start=1):
            cells = table.add_row().cells
            cells[0].text = str(number)
            cells[1].text = question
            cells[2].text = "\u2610"
            cells[3].text = "\u2610"
            cells[4].text = "\u2610"
            cells[5].text = ""
            for cell in cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        widths = [0.35, 4.55, 0.45, 0.45, 0.5, 1.45]
        for row in table.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = DocxInches(width)

    document.add_paragraph()
    notes = document.add_paragraph()
    notes.add_run("NOTE GENERALI / AZIONI DA INTRAPRENDERE").bold = True
    for _ in range(4):
        document.add_paragraph("________________________________________________________________________________")

    signatures = document.add_table(rows=2, cols=3)
    signatures.alignment = WD_TABLE_ALIGNMENT.CENTER
    signatures.style = "Table Grid"
    for index, label in enumerate(["HSE di cantiere", "Preposto", "Responsabile impresa"]):
        signatures.cell(0, index).text = label
        signatures.cell(0, index).paragraphs[0].runs[0].bold = True
        signatures.cell(1, index).text = "\nFirma: __________________________\n"

    footer = document.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Checklist HSE generata tramite HSE Risk Platform")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer



CSS = '''
<style>
:root{--ink:#123047;--accent:#087f8c}
.stApp{background:#f3f6fa}
.block-container{max-width:1450px;padding-top:2rem;padding-bottom:3rem}
h1,h2,h3{color:#123047;letter-spacing:-.025em}
[data-testid="stSidebar"]{background:#e9eff5;border-right:1px solid #d4dfe9}
[data-testid="stMetric"]{background:#fff;border:1px solid #d8e3eb;padding:18px;border-radius:12px}
[data-testid="stMetricLabel"]{color:#385268}
[data-testid="stMetricValue"]{color:#123047}
.stButton>button[kind="primary"]{background:#087f8c;border:0;border-radius:8px}
[data-testid="stVerticalBlockBorderWrapper"]{border-radius:12px}
[data-testid="stCaptionContainer"]{color:#415c70}
@media(max-width:700px){.block-container{padding:1rem}.stApp h1{font-size:1.9rem}}
</style>'''


def clean_rows(frame, awareness=False):
    rows = []
    for raw in frame.to_dict('records'):
        if pd.isna(raw.get('theme')):
            # Una nuova riga vuota non è un evento; una riga compilata in parte va segnalata.
            if all(pd.isna(v) or v in ['', False] for v in raw.values()):
                continue
        if awareness:
            rows.append(dict(theme=raw.get('theme'), count=raw.get('count'),
                             coverage=None if pd.isna(raw.get('coverage')) else float(raw['coverage'])))
        else:
            rows.append(dict(theme=raw.get('theme'), severity=raw.get('severity'), status=raw.get('status'),
                             late=False if pd.isna(raw.get('late')) else bool(raw['late']),
                             repeat=False if pd.isna(raw.get('repeat')) else bool(raw['repeat']),
                             note='' if pd.isna(raw.get('note')) else str(raw['note'])))
    return rows


def anomaly_editor(label, rows, key):
    st.markdown('**' + label + '**')
    frame = pd.DataFrame(rows, columns=['theme', 'severity', 'status', 'late', 'repeat', 'note'])
    config = {
       'theme': st.column_config.SelectboxColumn('Tema', options=THEMES, required=True),
       'severity': st.column_config.SelectboxColumn('Gravità', options=list(SEVERITY), required=True),
       'status': st.column_config.SelectboxColumn('Stato a fine mese', options=list(STATUS), required=True),
       'late': st.column_config.CheckboxColumn('Aperta scaduta', default=False),
       'repeat': st.column_config.CheckboxColumn('Recidiva', default=False),
       'note': st.column_config.TextColumn('Descrizione / riferimento')}
    return clean_rows(st.data_editor(frame, num_rows='dynamic', column_config=config, hide_index=True,
                                    width='stretch', key=key))


def render_form(records):
    if 'draft' not in st.session_state:
        st.session_state.draft = default_data()
    d = st.session_state.draft
    rev = st.session_state.get('draft_revision', 0)
    st.title('\U0001f4dd Valutazione mensile')
    st.caption('WIND / SOLAR / BESS · Input riferiti esclusivamente al mese selezionato')
    st.info('Registrare ogni anomalia una sola volta: NC formale oppure criticità informale. '
            'Per gli eventi precedenti riportare solo quelli ancora aperti a fine mese; non ricontare le chiusure dei mesi precedenti.')
    names = sorted({x['data']['nome'] for x in records}, key=str.casefold)
    if names:
        with st.expander('Usa il nome di un cantiere già presente'):
            chosen = st.selectbox('Cantieri registrati', names)
            if st.button('Usa questo nome'):
                st.session_state.draft = {**d, 'nome': chosen}
                st.session_state.draft_revision = rev + 1
                st.rerun()
    with st.form(f'evaluation_{rev}'):
        st.subheader('\U0001f4cd 1 · Cantiere e periodo')
        a, b, c = st.columns([2, 1, 1])
        name = a.text_input('Nome cantiere', d['nome'], max_chars=120, help='Usare lo stesso nome ogni mese per collegare lo storico.')
        technology = b.selectbox('Tecnologia', ['Wind', 'Solar', 'BESS'], index=['Wind', 'Solar', 'BESS'].index(d['technology']))
        month = c.date_input('Mese di riferimento', date.fromisoformat(d['month'] + '-01'), max_value=date.today(), help='Il giorno viene ignorato: il periodo è sempre mensile.')
        phase = st.selectbox('Fase operativa prevalente', list(PHASE_THEMES), index=list(PHASE_THEMES).index(d['fase']))
        st.caption('Se selezioni il mese corrente, la valutazione è parziale e va confrontata con cautela.')
        st.subheader('\U0001f3d7\ufe0f 2 · Lavorazioni e interferenze')
        activities = st.multiselect('Attività effettivamente svolte nel mese', list(ACTIVITY), default=d['activities'])
        interference = st.slider('Interferenza fra attività', 1, 10, int(d['interference_level']),
          help='1: attività separate. 3: contatti occasionali gestiti. 5: aree/risorse condivise. 7: sovrapposizioni frequenti. 10: interferenze continue senza separazione efficace.')
        st.subheader('\U0001f50d 3 · Anomalie e follow-up')
        st.caption('Dopo aver modificato una cella, premi Invio o clicca fuori dalla cella prima di calcolare. Aperte = presenti a fine mese; chiuse = risolte durante il mese. Gravità alta = potenziale danno grave, da motivare nella descrizione. '
                   'Criticità informali: peso inferiore, nessuna soglia automatica. Se formalizzate, rimuoverle dalla tabella informale.')
        nc = anomaly_editor('\U0001f4dd NC formali · massimo 30 punti', d['nc_items'], f'nc_{rev}')
        crit = anomaly_editor('\U0001f4cc Criticità informali · massimo 8 punti', d['crit_items'], f'crit_{rev}')
        danger = st.checkbox('È ancora presente un pericolo grave e attuale a fine valutazione', value=d['imminent_danger'],
           help='Attiva un allarme dedicato e un indice minimo 81. Non dipende dal canale formale/informale di registrazione.')
        st.subheader('\U0001f477 4 · Organizzazione e presidio')
        a, b, c = st.columns(3)
        app = a.number_input('Imprese appaltatrici operative', 0, 200, int(d['app']))
        sub = b.number_input('Imprese subappaltatrici operative', 0, 200, int(d['sub']))
        hse = c.number_input('Professionisti HSE delle imprese', 0, 200, int(d['company_hse']), help='Dato descrittivo; non equivale a copertura reale.')
        coverage = st.slider('Copertura HSE effettiva delle lavorazioni (%)', 0, 100, int(d['hse_coverage']),
          help='Quota delle lavorazioni/turni/aree per cui è stato assicurato il presidio HSE pianificato. Stima dichiarata, non numero di persone e non criterio normativo universale.')
        a, b = st.columns(2)
        inspections = a.number_input('Ispezioni strutturate eseguite nel mese', 0, 2000, int(d['inspections']))
        target = b.number_input('Target mensile ispezioni · politica comune di confronto', 1, 1000, int(d['inspection_target']),
          help='Default 4 è una ipotesi gestionale. Per graduatorie omogenee usare lo stesso target; il confronto esclude target diversi.')
        st.subheader('\u26d4 5 · Stop Work')
        st.caption('Positivo: interruzione preventiva tempestiva e gestita. Negativo: stop imposto per carenza o mancato rispetto di misure. '
                   'Classificare ciascun evento in una sola categoria; registrare una sola volta anche l’anomalia sottostante. La classificazione resta responsabilità dell’utilizzatore.')
        a, b = st.columns(2)
        positive = a.number_input('Stop Work positivi · massimo −2', 0, 1000, int(d['stop_positive']))
        negative = b.number_input('Stop Work negativi · massimo +6', 0, 1000, int(d['stop_negative']))
        st.subheader('\U0001f9e0 6 · Sensibilizzazioni mirate')
        st.caption('Una riga per tema. Copertura = % di lavoratori destinatari unici raggiunti sul tema nel mese (non somma delle presenze). '
                   'Lasciare vuota se sconosciuta: fattore prudenziale 50%. Senza collegamento ad attività, anomalie aperte o fase non c’è bonus.')
        awareness_frame = pd.DataFrame(d['awareness'], columns=['theme', 'count', 'coverage'])
        awareness = clean_rows(st.data_editor(awareness_frame, num_rows='dynamic', hide_index=True, width='stretch',
            key=f'awareness_{rev}', column_config={
                'theme': st.column_config.SelectboxColumn('Tema', options=THEMES, required=True),
                'count': st.column_config.NumberColumn('Incontri sul tema', min_value=1, max_value=1000, step=1, required=True),
                'coverage': st.column_config.NumberColumn('Destinatari unici raggiunti (%)', min_value=0, max_value=100, step=1)}), awareness=True)
        with st.expander('\U0001f4d0 7 · Dati di esposizione FACOLTATIVI'):
            st.write('Puoi lasciare tutto a zero: significa dato non disponibile. Questi dati producono un indice contestualizzato separato, '
                     'senza modificare l’indice base usato nelle graduatorie. I rapporti adottati sono ipotesi gestionali da calibrare.')
            a, b, c = st.columns(3)
            hours = a.number_input('Ore lavorate complessive del mese', 0., 10000000., float(d['hours'] or 0), step=100.)
            workers = b.number_input('Numero medio lavoratori nel mese', 0, 100000, int(d['workers'] or 0))
            days = c.number_input('Giorni di attività nel mese', 0, 31, int(d['days'] or 0))
        notes = st.text_area('Note della valutazione', d['notes'], max_chars=5000)
        submitted = st.form_submit_button('Calcola la valutazione', icon='\U0001f9ee', type='primary', width='stretch')
    if submitted:
        candidate = dict(nome=name.strip(), technology=technology, fase=phase, month=month.strftime('%Y-%m'),
            app=app, sub=sub, company_hse=hse, hse_coverage=coverage, inspections=inspections,
            inspection_target=target, activities=activities, interference_level=interference,
            nc_items=nc, crit_items=crit, awareness=awareness, stop_positive=positive, stop_negative=negative,
            imminent_danger=danger, hours=hours or None, workers=workers or None, days=days or None, notes=notes)
        errors = validate_data(candidate)
        if errors:
            for error in errors:
                st.error(error)
        else:
            st.session_state.draft = candidate
            st.session_state.pop('trial_discarded', None)
            st.session_state.pop('exports', None)
            st.session_state.pop('advisor_reply', None)
            st.session_state.current_report = new_record(candidate, calculate_risk(candidate))
            st.success('Prova calcolata, non salvata. Apri “Risultati e azioni”; al termine scegli Salva e inserisci oppure Non salvare.')


def driver_chart(r):
    frame = pd.DataFrame(r['drivers'])
    return alt.Chart(frame).mark_bar(cornerRadiusEnd=4).encode(
        x=alt.X('Contributo:Q', title='Punti effettivi sull’indice base'),
        y=alt.Y('Area:N', sort=None, title=None),
        color=alt.condition(alt.datum.Contributo < 0, alt.value('#16826b'), alt.value('#225873')),
        tooltip=['Area:N', alt.Tooltip('Contributo:Q', format='.2f')]).properties(height=390)


def anomaly_counts(rows):
    opened = sum(x['status'] == 'Aperta' for x in rows)
    return {'total': len(rows), 'open': opened, 'closed': len(rows) - opened}


def anomaly_ledger(rows):
    return [dict(Numero=i + 1, Tema=x['theme'], Gravità=x['severity'], Stato=x['status'],
                 Scaduta=x['late'], Recidiva=x['repeat'], Descrizione=x.get('note', ''),
                 **{'Fattore stato': STATUS[x['status']], 'Carico grezzo': round(anomaly_load([x]), 4)})
            for i, x in enumerate(rows)]


REPORT_SECTIONS = ['\U0001f4ca Sintesi del mese', '\U0001f3af Azioni e alert', '\U0001f4cb Action Plan',
                   '\U0001f5c2\ufe0f Registro NC e criticità', '\U0001f50d Cause e motivazioni', '\U0001f9f0 Toolbox Talk',
                   '\U0001f4c9 Come ridurre il rischio', '\U0001f9ee Dettaglio del calcolo', '\U0001f4ac HSE Advisor',
                   '\u2b50 HSE Scorecard', '\U0001f4d1 HSE Toolkit', '\U0001f4e4 Export']


def render_report(store, records):
    st.title('\U0001f4ca Risultati e azioni')
    record = st.session_state.get('current_report')
    if not record:
        if st.session_state.pop('trial_discarded', False):
            st.success('Prova scartata: dati e output rimossi dalla sessione, senza inserimenti nello storico. Puoi iniziare una nuova valutazione.')
        st.info('Calcola una valutazione oppure aprine una dallo storico.')
        return
    d, r = record['data'], record['result']
    st.subheader(f"{d['nome']} · {d['technology']} · {d['month']}")
    st.caption(f"Interfaccia {UI_VERSION} · modello {MODEL_VERSION} · snapshot calcolato: dopo modifiche al modulo premi nuovamente Calcola.")
    a, b = st.columns(2)
    a.metric('\U0001f4c8 Indice base', f"{r['score']:.1f} / 100", r['level'], delta_color='off')
    b.metric('\U0001f50e Indice contestualizzato', '—' if r['context_score'] is None else f"{r['context_score']:.1f} / 100")
    nc, crit = anomaly_counts(d['nc_items']), anomaly_counts(d['crit_items'])
    a, b, c = st.columns(3)
    a.metric('\U0001f4dd NC formali totali', nc['total'])
    b.metric('\U0001f534 NC formali aperte', nc['open'])
    c.metric('\u2705 NC formali chiuse', nc['closed'])
    a, b, c = st.columns(3)
    a.metric('\U0001f4cc Criticità informali totali', crit['total'])
    b.metric('\U0001f7e0 Criticità informali aperte', crit['open'])
    c.metric('\u2611\ufe0f Criticità informali chiuse', crit['closed'])
    if nc['closed']:
        st.info(f"Sono registrate {nc['total']} NC formali: {nc['open']} aperte e {nc['closed']} chiuse. "
                'Le NC chiuse compaiono nel totale e nel registro; mantengono un peso ridotto nel calcolo e non generano nuove azioni di chiusura.')
    st.caption('\U0001f7e2 0–20 Basso · >20–40 Medio basso · >40–60 Medio · >60–80 Alto · >80–100 Critico. '
               'Indice gestionale sperimentale, non probabilità di infortunio.')
    if r['floor']:
        st.error(f"\U0001f6a8 Soglia minima attiva: {r['floor_reason']} → almeno {r['floor']:.0f}/100. I bonus non la possono compensare.")
    if d['month'] == date.today().strftime('%Y-%m'):
        st.warning('\U0001f4c5 Mese corrente: dati parziali, esclusi dalla graduatoria consolidata.')
    is_saved = record['id'] in {x['id'] for x in records}
    if is_saved:
        st.success('\U0001f4be Valutazione già inserita nello storico.')
    else:
        st.info('\U0001f9ea MODALITÀ PROVA · input e output temporanei. In fondo alla pagina scegli Salva e inserisci oppure Non salvare.')
    st.divider()
    section = st.selectbox('\U0001f9ed Apri una sezione del report', REPORT_SECTIONS, key='report_section')
    st.subheader(section)
    st.caption('12 sezioni dedicate: sintesi, alert, piano azioni, registro, motivazioni, toolbox, scenari, calcolo, advisor, scorecard, checklist ed export.')
    actions = action_plan(d, r)
    if section == REPORT_SECTIONS[0]:
        st.write(executive_summary(d, r))
        st.altair_chart(driver_chart(r), width='stretch')
        st.caption('Contributi effettivi, incluse le rettifiche per limite di scala e soglia minima.')
        nc_driver = next(x for x in r['drivers'] if x['Area'] == 'NC formali')
        crit_driver = next(x for x in r['drivers'] if x['Area'] == 'Criticità informali')
        st.write(f"\U0001f4dd NC formali: {nc['total']} registrate, contributo {nc_driver['Contributo']:.2f} punti su 30. "
                 f"\U0001f4cc Criticità informali: {crit['total']} registrate, contributo {crit_driver['Contributo']:.2f} punti su 8.")
    elif section == REPORT_SECTIONS[1]:
        if actions:
            for item in actions:
                text = f"{item['Priorità']} · {item['Area']} — {item['Azione']}"
                (st.warning if item['Priorità'] == 'P1' else st.info)(text)
        else:
            st.success('\u2705 Nessun alert specifico generato. Mantenere e verificare le misure previste.')
        st.caption('Azioni correttive sulle anomalie aperte. Le chiuse restano nel registro e nel calcolo con peso residuo.')
    elif section == REPORT_SECTIONS[2]:
        if actions:
            st.dataframe(pd.DataFrame(actions), hide_index=True, width='stretch')
        else:
            st.info('Nessuna azione specifica proposta.')
        st.caption('Responsabili e scadenze proposti sono da confermare. Il piano esportabile non è un registro persistente delle chiusure.')
    elif section == REPORT_SECTIONS[3]:
        for field, title in [('nc_items', '\U0001f4dd Registro completo delle NC formali'), ('crit_items', '\U0001f4cc Registro completo delle criticità informali')]:
            st.markdown('### ' + title)
            if d[field]:
                st.dataframe(pd.DataFrame(anomaly_ledger(d[field])), hide_index=True, width='stretch')
                counts = anomaly_counts(d[field])
                st.caption(f"Totale {counts['total']} · aperte {counts['open']} · chiuse {counts['closed']}. Nessun filtro di stato applicato.")
                for item in anomaly_ledger(d[field]):
                    with st.expander(f"{item['Numero']} · {item['Tema']} · {item['Stato']}"):
                        st.write(item['Descrizione'] or 'Nessuna descrizione inserita.')
                        st.write(f"Carico grezzo {item['Carico grezzo']:.4f}; fattore stato {item['Fattore stato']:.1f}.")
            else:
                st.info('Nessuna anomalia registrata in questo elenco.')
        st.caption('Chiusa nei tempi: fattore 0,2; chiusa in ritardo: 0,4; aperta: 1. '
                   'Il carico grezzo entra nella funzione non lineare: non equivale a punti finali sommabili per riga.')
    elif section == REPORT_SECTIONS[4]:
        st.write(executive_summary(d, r))
        for item in sorted(r['drivers'][:6], key=lambda x: x['Contributo'], reverse=True):
            st.markdown(f"#### {item['Area']} · {item['Contributo']:.2f} punti")
            st.write(f"Indice area {item['Indice area']:.2f}/100 · peso massimo {item['Peso massimo']} punti.")
            if item['Area'] in ['NC formali', 'Criticità informali']:
                field = 'nc_items' if item['Area'] == 'NC formali' else 'crit_items'
                for event in d[field]:
                    st.write(f"• {event['theme']} / {event['severity']} / {event['status']}: {event.get('note', '')}")
        st.caption('Le motivazioni spiegano i dati e il punteggio. Le cause profonde richiedono verifica sul campo: il tool non le dimostra automaticamente.')
    elif section == REPORT_SECTIONS[5]:
        if r['awareness_details']:
            frame = pd.DataFrame(r['awareness_details'])
            frame['Bonus punti'] = [8 * item['Pertinenza'] * item['Efficacia convenzionale'] / sum(r['needs'].values()) for item in r['awareness_details']]
            st.dataframe(frame, hide_index=True, width='stretch')
        else:
            st.info('Nessuna sensibilizzazione registrata nel mese.')
        themes = sorted(THEMES, key=lambda t: r['needs'].get(t, 0), reverse=True)
        theme = st.selectbox('\U0001f9f0 Genera un toolbox a tema · priorità in cima', themes)
        st.text(toolbox_text(theme, d, r))
        st.download_button('\U0001f4e5 Scarica toolbox Word', word_report(d, r, theme), f"Toolbox_{safe_filename(theme)}_{d['month']}.docx")
    elif section == REPORT_SECTIONS[6]:
        st.dataframe(pd.DataFrame(simulations(d, r)), hide_index=True, width='stretch')
        st.caption('Scenari indipendenti e non sommabili: simulano il punteggio, non la riduzione reale del rischio. '
                   'Chiusura significa misure attuate e verificate; il pericolo attuale resta attivo finché non viene rimosso e rivalutato.')
    elif section == REPORT_SECTIONS[7]:
        st.code('Base = clamp(Σ indice_area × peso / 100 + SW_neg − SW_pos − bonus_toolbox, 0, 100)\nFinale = max(Base, soglia_minima)')
        st.dataframe(pd.DataFrame(r['drivers']), hide_index=True, width='stretch')
        for text in r['explanations']:
            st.write('• ' + text)
        st.markdown('### \U0001f50e Dati facoltativi e qualità del dato')
        st.write(r['context_basis'])
        if r['context_target'] is not None:
            st.write(f"Target contestualizzato: {r['context_target']:.2f} ispezioni; cambia solo il deficit ispettivo nell’indice separato.")
        for text in r['missing']:
            st.warning(text)
        st.write('NC/ispezioni (descrittivo): ' + (f"{len(d['nc_items']) / d['inspections']:.2f}" if d['inspections'] else 'non valutabile'))
        if d['hours']:
            st.write(f"NC formali / 10.000 ore: {len(d['nc_items']) / d['hours'] * 10000:.2f}; include arretrato aperto, non è un tasso di nuove NC.")
    elif section == REPORT_SECTIONS[8]:
        st.caption('Assistente locale a regole, senza API AI. Spiega indice, azioni, presidio e temi dei toolbox.')
        q = st.text_input('\U0001f4ac Domanda sul report', max_chars=500)
        if st.button('\U0001f4ac Spiega il report'):
            st.session_state.advisor_reply = (record['id'], local_advisor(q, d, r))
        reply = st.session_state.get('advisor_reply')
        if reply and reply[0] == record['id']:
            st.markdown(reply[1])
    elif section == REPORT_SECTIONS[9]:
        st.caption('Indici area 0–100: più alto significa maggiore attenzione. La scorecard usa gli stessi fattori della formula.')
        for row in r['drivers'][:6]:
            st.write(f"{row['Area']}: {row['Indice area']:.1f}/100 · contributo {row['Contributo']:.2f} punti")
            st.progress(bounded(row['Indice area']) / 100)
    elif section == REPORT_SECTIONS[10]:
        defaults = list(dict.fromkeys(['Controlli generali pre-attività'] + [ACTIVITY_TO_CHECKLIST[a] for a in d['activities'] if a in ACTIVITY_TO_CHECKLIST]))
        sections = st.multiselect('\U0001f4d1 Capitoli della checklist', list(CHECKLIST_LIBRARY), default=defaults)
        if sections:
            st.download_button('\U0001f4e5 Scarica checklist Word', generate_checklist_word(d, sections).getvalue(), f"Checklist_{safe_filename(d['nome'])}.docx")
    elif section == REPORT_SECTIONS[11]:
        st.caption('Report temporanei fino al download volontario. Il backup della singola valutazione è disponibile dopo Salva e inserisci.')
        filename = safe_filename(d['nome']) + '_' + d['month']
        if st.button('Prepara i report completi', icon='\U0001f4e6'):
            with st.spinner('Preparazione report…'):
                st.session_state.exports = dict(id=record['id'], excel=excel_report(d, r, records), ppt=ppt_report(d, r), word=word_report(d, r))
        exports = st.session_state.get('exports')
        if exports and exports['id'] == record['id']:
            st.download_button('\U0001f4ca Excel completo', exports['excel'], filename + '.xlsx')
            st.download_button('\U0001f4fd\ufe0f PowerPoint', exports['ppt'], filename + '.pptx')
            st.download_button('\U0001f4c4 Report Word', exports['word'], filename + '.docx')
    st.divider()
    st.subheader('\U0001f4be Concludi la valutazione')
    if is_saved:
        st.success('Valutazione già nello storico. Per una nuova prova torna al modulo e ricalcola.')
        st.download_button('\U0001f4e5 Backup della valutazione salvata', backup_bytes([record]), f"{safe_filename(d['nome'])}_{d['month']}.json", 'application/json')
    else:
        st.caption('Salva e inserisci conserva input e output nello storico. Non salvare scarta la prova e svuota il modulo. '
                   'Lo storico precedente resta disponibile; eventuali file già scaricati restano sul computer.')
        save_col, discard_col = st.columns(2)
        if save_col.button('Salva e inserisci', type='primary', width='stretch', icon='\U0001f4be'):
            try:
                store.save_many([record])
                st.rerun()
            except Exception:
                st.error('Salvataggio non riuscito: prova ancora temporanea. Riprova oppure scegli Non salvare.')
        if discard_col.button('Non salvare', width='stretch', icon='\U0001f9ea'):
            discard_trial()
            st.rerun()

def discard_trial():
    """Scarta solo la prova corrente, mai i record già inseriti nello storico."""
    revision = st.session_state.get('draft_revision', 0) + 1
    for key in list(st.session_state):
        if key in {'draft', 'current_report', 'exports', 'advisor_reply'} or key.startswith(('nc_', 'crit_', 'awareness_')):
            del st.session_state[key]
    st.session_state.draft_revision = revision
    st.session_state.trial_discarded = True


def render_history(records):
    st.title('\U0001f5d3\ufe0f Storico per cantiere')
    frame = history_frame(records)
    if frame.empty:
        st.info('Salva la prima valutazione per iniziare lo storico.')
        return
    sites = sorted(frame['Cantiere'].unique())
    site = st.selectbox('Cantiere', sites)
    selected = frame[frame['site'] == site_key(site)].copy()
    selected['Periodo'] = pd.to_datetime(selected['Mese'])
    st.caption('Ogni salvataggio è una revisione conservata. I grafici usano l’ultima revisione per cantiere, mese e versione modello. '
               'I mesi mancanti interrompono la linea: non sono valutazioni pari a zero.')
    target_values = {x['data']['inspection_target'] for x in records if site_key(x['data']['nome']) == site_key(site)}
    if len(target_values) > 1:
        st.warning('Il target ispettivo è cambiato nello storico: parte del trend può dipendere da questo cambiamento di criterio.')
    # Reindicizzazione per non collegare mesi mancanti e dividerli in segmenti.
    series = selected.set_index('Periodo').reindex(pd.date_range(selected['Periodo'].min(), selected['Periodo'].max(), freq='MS')).rename_axis('Periodo').reset_index()
    series['segmento'] = series['Indice'].isna().cumsum()
    st.altair_chart(alt.Chart(series.dropna(subset=['Indice'])).mark_line(point=True).encode(
        x=alt.X('Periodo:T', title='Mese'), y=alt.Y('Indice:Q', scale=alt.Scale(domain=[0,100])),
        detail='segmento:N', tooltip=['Periodo:T', 'Indice:Q', 'Livello:N']), width='stretch')
    st.dataframe(selected.drop(columns=['site','Periodo','id']), hide_index=True, width='stretch')
    revisions = [x for x in records if site_key(x['data']['nome']) == site_key(site)]
    revisions.sort(key=lambda x: (x['data']['month'], x['saved_at']), reverse=True)
    lookup = {x['id']: x for x in revisions}
    ident = st.selectbox('Apri una valutazione o revisione', list(lookup), format_func=lambda i:
        f"{lookup[i]['data']['month']} · {lookup[i]['result']['score']:.1f} · salvata {lookup[i]['saved_at'][:19]} · {i[:8]}")
    a, b = st.columns(2)
    if a.button('Apri nei risultati'):
        st.session_state.current_report = deepcopy(lookup[ident])
        st.success('Valutazione caricata: apri “Risultati e azioni”.')
    if b.button('Usa come base per una nuova valutazione'):
        st.session_state.draft = deepcopy(lookup[ident]['data'])
        st.session_state.draft_revision = st.session_state.get('draft_revision', 0) + 1
        st.success('Modulo preparato. Apri “Valutazione mensile” e aggiorna mese e dati prima del calcolo.')


def render_comparison(records):
    st.title('\U0001f4ca Confronto cantieri')
    frame = history_frame(records)
    if frame.empty:
        st.info('Servono valutazioni salvate nello storico.')
        return
    months = sorted([m for m in frame['Mese'].unique() if m < date.today().strftime('%Y-%m')], reverse=True)
    if not months:
        st.info('La graduatoria consolidata richiede almeno un mese concluso.')
        return
    month = st.selectbox('Mese concluso · stesso periodo per tutti', months)
    types = st.multiselect('Tecnologie', ['Wind','Solar','BESS'], default=['Wind','Solar','BESS'])
    lookup = {x['id']: x for x in records}
    frame['Target'] = frame['id'].map(lambda i: lookup[i]['data']['inspection_target'])
    target = st.selectbox('Politica comune · target ispezioni', sorted(frame.loc[frame['Mese'] == month, 'Target'].unique()))
    selected = frame[(frame['Mese'] == month) & frame['Tecnologia'].isin(types) & (frame['Target'] == target) & (frame['Versione'] == MODEL_VERSION)].sort_values('Indice', ascending=False)
    st.caption('Indice base; stesso mese, stessa versione modello e stesso target ispezioni. '
               'La graduatoria dipende anche da completezza e uniformità delle registrazioni; l’indice contestualizzato non viene mescolato al base.')
    if selected.empty:
        st.info('Nessuna valutazione con questi filtri.')
        return
    st.altair_chart(alt.Chart(selected).mark_bar(cornerRadiusEnd=4, color='#225873').encode(
        x=alt.X('Indice:Q', scale=alt.Scale(domain=[0,100])), y=alt.Y('Cantiere:N', sort='-x'),
        tooltip=['Cantiere','Tecnologia','Indice','Livello']), width='stretch')
    st.dataframe(selected[['Cantiere','Tecnologia','Mese','Indice','Livello','Target']], hide_index=True, width='stretch')
    counts = selected.groupby('Livello').size().reset_index(name='Cantieri')
    st.subheader('Distribuzione dei cantieri per fascia')
    st.altair_chart(alt.Chart(counts).mark_arc(innerRadius=65).encode(
        theta='Cantieri:Q', color=alt.Color('Livello:N', scale=alt.Scale(domain=['Basso','Medio basso','Medio','Alto','Critico'], range=COLORS)),
        tooltip=['Livello','Cantieri']), width='stretch')


def render_archive(store, records):
    st.title('\U0001f5c4\ufe0f Archivio e backup')
    if store.mode == 'session':
        st.warning('Modalità gratuita senza configurazione: dati disponibili solo nella sessione corrente. '
                   'Scarica il backup JSON prima di chiudere o ricaricare il browser; importalo alla prossima apertura. Non è uno storico cloud automatico.')
    elif store.mode == 'sqlite':
        st.warning('Archivio SQLite locale: persiste su un disco stabile, ma su Streamlit Community Cloud può essere perso al riavvio o al redeploy. Scarica regolarmente il backup.')
    else:
        st.success('Archivio PostgreSQL configurato: storico condiviso fra gli utenti autorizzati dell’app. La conservazione dipende dal servizio database configurato.')
    st.metric('Valutazioni e revisioni archiviate', len(records))
    st.download_button('Scarica backup completo JSON', backup_bytes(records), f"HSE_backup_{date.today().isoformat()}.json", 'application/json')
    uploaded = st.file_uploader('Ripristina o unisci un backup JSON · massimo 20 MB', type=['json'])
    if uploaded is not None and st.button('Verifica e importa il backup'):
        try:
            incoming = parse_backup(uploaded.getvalue())
            existing = {x['id']: x for x in records}
            for item in incoming:
                if item['id'] in existing and (item['data'] != existing[item['id']]['data'] or item['saved_at'] != existing[item['id']]['saved_at']):
                    raise ValueError('Conflitto: un identificativo esistente ha dati differenti. Nessun record importato.')
            store.save_many(incoming)
            st.success(f"Importazione completata: {sum(x['id'] not in existing for x in incoming)} nuove valutazioni. I duplicati identici sono ignorati.")
            st.rerun()
        except (ValueError, TypeError, KeyError, OverflowError):
            st.error('Backup non valido o in conflitto. Nessun dato importato: controllare formato, versione modello e campi obbligatori.')
        except Exception:
            st.error('Archivio non raggiungibile: importazione non completata. Riprova quando il collegamento è disponibile.')
    with st.expander('Come scegliere la conservazione dei dati'):
        st.write('Nessun servizio a pagamento è necessario per usare il tool e i backup portabili. '
                 'Uno storico automatico condiviso su Streamlit Cloud richiede un database esterno: il codice supporta PostgreSQL. '
                 'Un eventuale piano gratuito va scelto e configurato separatamente, verificando limiti e condizioni del fornitore. '
                 'SQLite è adatto a un PC/server con disco persistente. Le istruzioni sono nel README del pacchetto.')
        st.write('L’archivio configurato è condiviso da chi conosce la password dell’app; non ci sono account personali o separazione per impresa. '
                 'Usare dati di cantiere essenziali, senza informazioni personali non necessarie.')


def main():
    st.set_page_config(page_title='HSE Risk Platform', page_icon='\U0001f9ba', layout='wide')
    st.markdown(CSS, unsafe_allow_html=True)
    mode = str(setting('storage_mode', 'session')).lower()
    password = str(setting('app_password', ''))
    if mode not in ['session', 'sqlite', 'postgres']:
        st.error('Configurazione archivio non valida.')
        st.stop()
    if mode != 'session' and not password:
        st.error('Per usare un archivio condiviso configura app_password nei Secrets. In alternativa usa storage_mode="session".')
        st.stop()
    if password:
        if not st.session_state.get('authenticated'):
            st.title('HSE Risk Platform')
            with st.form('access'):
                entered = st.text_input('Password di accesso', type='password')
                go = st.form_submit_button('Accedi')
            if go:
                if hmac.compare_digest(entered.encode(), password.encode()):
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error('Password non corretta.')
            st.stop()
    try:
        store = HistoryStore(mode, str(setting('database_url')), str(setting('sqlite_path', 'hse_history.sqlite3')))
        records = store.all()
    except Exception:
        st.error('Archivio non disponibile. Controllare configurazione e collegamento. Nessun passaggio automatico a un archivio vuoto.')
        st.stop()
    with st.sidebar:
        st.markdown('## HSE Risk Platform')
        st.caption('WIND · SOLAR · BESS')
        page = st.radio('\U0001f9ed Navigazione', ['Valutazione mensile', 'Risultati e azioni', 'Storico cantiere', 'Confronto cantieri', 'Archivio e backup'], format_func=lambda label: {'Valutazione mensile':'\U0001f4dd Valutazione mensile', 'Risultati e azioni':'\U0001f3af Risultati e azioni', 'Storico cantiere':'\U0001f5d3\ufe0f Storico cantiere', 'Confronto cantieri':'\U0001f4ca Confronto cantieri', 'Archivio e backup':'\U0001f5c4\ufe0f Archivio e backup'}[label])
        st.divider()
        st.caption(f'Interfaccia {UI_VERSION} · modello {MODEL_VERSION} · pesi sperimentali')
        st.caption('Periodo: un mese. Nessuna API AI. Output generati da regole esplicite.')
        st.caption('Archivio: ' + {'session':'sessione + backup portabile','sqlite':'SQLite locale','postgres':'PostgreSQL condiviso'}[mode])
        if password and st.button('Esci'):
            st.session_state.clear()
            st.rerun()
    if page == 'Valutazione mensile':
        render_form(records)
    elif page == 'Risultati e azioni':
        render_report(store, records)
    elif page == 'Storico cantiere':
        render_history(records)
    elif page == 'Confronto cantieri':
        render_comparison(records)
    else:
        render_archive(store, records)


if __name__ == '__main__':
    main()

