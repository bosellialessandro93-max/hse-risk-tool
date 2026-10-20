"""HSE Risk Platform 3.3 — modello gestionale trasparente, non predittivo.
Avvio: streamlit run app.py. Tutti gli input HSE Enel sono riferiti alla settimana ISO selezionata.
Nessuna chiamata a modelli AI, nessuna API a pagamento.
"""
from __future__ import annotations
from io import BytesIO, StringIO
from datetime import datetime, date, timezone, timedelta
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
import csv
import html
import hashlib
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

MODEL_VERSION = '2.1.0'
LEGACY_MODEL_VERSION = '2.0.0'
UI_VERSION = '3.3.0'
# Quattro ispezioni/mese diventano un riferimento convenzionale di una ispezione/settimana.
DEFAULT_INSPECTION_TARGET = 1
NC_FORMAL_HELP = ('Inserire solo le NC registrate in HSEQ4ALL nel periodo di analisi selezionato. '
                  'Indicare lo stato a fine periodo e registrare ogni anomalia una sola volta, in questo elenco oppure in quello informale.')
NC_INFORMAL_HELP = ('Inserire solo le NC non registrate in HSEQ4ALL e rilevate nel periodo di analisi selezionato. '
                    'Indicare lo stato a fine periodo e registrare ogni anomalia una sola volta, in questo elenco oppure in quello formale.')
STOP_POSITIVE_HELP = ('Interruzione preventiva tempestiva e gestita. Ogni Stop Work positivo riduce il Risk Index di 0,5 punti, '
                      'fino a una riduzione complessiva di 2 punti. Non compensa le soglie minime di tutela. '
                      'Classificare ogni evento una sola volta.')
STOP_NEGATIVE_HELP = ('Stop imposto per carenze o mancato rispetto delle misure. Ogni Stop Work negativo aumenta il Risk Index '
                      'di 1,5 punti, fino a un incremento complessivo di 6 punti. Classificare ogni evento una sola volta.')
SCHEMA_VERSION = 2
# Il precedente 10% di Organizzazione confluisce interamente nelle NC formali.
WEIGHTS = {'NC formali': 40, 'Criticità informali': 8, 'Attività': 22,
           'Interferenze': 15, 'Presidio e controlli': 15}
LEGACY_WEIGHTS = {'NC formali': 30, 'Criticità informali': 8, 'Attività': 22,
                  'Interferenze': 15, 'Organizzazione': 10, 'Presidio e controlli': 15}
assert sum(WEIGHTS.values()) == 100
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


def week_monday(value):
    """Il periodo di valutazione è una settimana ISO, da lunedì a domenica."""
    value = date.fromisoformat(value) if isinstance(value, str) else value
    return value - timedelta(days=value.weekday())


def period_key(d):
    """Chiave compatibile con record settimanali e vecchi record mensili."""
    return d.get('week_start') or d.get('month', '')


def is_weekly_assessment(d):
    return bool(d.get('week_start'))


def period_label(d):
    if is_weekly_assessment(d):
        monday = date.fromisoformat(d['week_start'])
        sunday = monday + timedelta(days=6)
        year, week, _ = monday.isocalendar()
        return f'Settimana {week:02d}/{year} · {monday:%d/%m/%Y}–{sunday:%d/%m/%Y}'
    return f'Mese {d.get("month", "non disponibile")}'


def analysis_period(d):
    if is_weekly_assessment(d):
        return period_label(d)
    dt = datetime.strptime(d['month'], '%Y-%m')
    names = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
             'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre']
    last = calendar.monthrange(dt.year, dt.month)[1]
    return f"{names[dt.month-1]} {dt.year} (01/{dt.month:02d}/{dt.year} - {last:02d}/{dt.month:02d}/{dt.year})"


def next_weekly_draft(source=None):
    """Duplica soltanto il contesto del cantiere, non le registrazioni di un'altra settimana."""
    d = default_data() if source is None else deepcopy(source)
    d.pop('month', None)
    d['week_start'] = week_monday(date.today()).isoformat()
    for field in ('nc_items','crit_items','awareness'):
        d[field] = []
    d.update(inspections=0, inspection_target=DEFAULT_INSPECTION_TARGET,
             stop_positive=0, stop_negative=0, imminent_danger=False,
             hours=None, workers=None, days=None, notes='', input_scope='week_registrations_v1')
    return d


def assessment_drivers(r):
    return [x for x in r['drivers'] if x.get('Indice area') is not None]


def assessment_calculator(d):
    return calculate_risk(d) if is_weekly_assessment(d) else calculate_legacy_risk(d)


def detection_details(d):
    target = int(d.get('inspection_target', DEFAULT_INSPECTION_TARGET))
    coverage = min(1., d['inspections'] / target)
    deficit = 100 * (1 - coverage)
    return dict(target=target, coverage=coverage * 100, deficit=deficit,
                points=deficit * .35 * .15)


def detection_help(target=DEFAULT_INSPECTION_TARGET):
    return (f'Inserire solo le ispezioni registrate in HSEQ4ALL nel periodo di analisi. '
            f'Detection convenzionale = min(ispezioni / {target}, 1) × 100. '
            'Deficit di detection = 100 − detection. Il deficit pesa il 35% nell’area Presidio e controlli, '
            'che pesa il 15% sul Risk Index: al massimo 5,25 punti. '
            f'Con 0 ispezioni il contributo è +5,25 punti; con almeno {target} è 0. '
            'Il riferimento è gestionale, non un requisito normativo né una probabilità statistica di rilevazione. '
            'Il rapporto NC/ispezioni è solo descrittivo: rilevare più NC non penalizza l’efficacia delle ispezioni. '
            'Le esposizioni facoltative cambiano soltanto l’indice contestualizzato separato.')


def scorecard_stars(area_index):
    """Più stelle = minore indice di attenzione dell’area, con arrotondamento a 20 punti."""
    filled = max(0, min(5, int((100 - bounded(area_index)) / 20 + .5)))
    return filled, '★' * filled + '☆' * (5 - filled)


def default_data():
    return dict(nome='', technology='Wind', fase='costruzione', week_start=(week_monday(date.today()) - timedelta(days=7)).isoformat(),
                app=1, sub=0, company_hse=1, hse_coverage=80., inspections=0,
                inspection_target=DEFAULT_INSPECTION_TARGET, activities=['civili'], interference_level=3,
                nc_items=[], crit_items=[], awareness=[], stop_positive=0, stop_negative=0,
                imminent_danger=False, hours=None, workers=None, days=None, notes='')


def validate_data(d):
    errors = []
    if not isinstance(d, dict):
        return ['Valutazione non valida.']
    if not isinstance(d.get('nome'), str) or not d['nome'].strip() or len(d['nome']) > 120:
        errors.append('Inserire un nome cantiere (massimo 120 caratteri).')
    if d.get('week_start') and d.get('month'):
        errors.append('Indicare un solo periodo: settimana o mese storico.')
    month = None
    if is_weekly_assessment(d):
        try:
            monday = date.fromisoformat(d['week_start'])
            if monday.isoformat() != d['week_start'] or monday.weekday() != 0 or monday > date.today():
                errors.append('La settimana deve iniziare di lunedì e non può essere futura.')
        except (TypeError, ValueError, AttributeError):
            errors.append('Settimana non valida: indicare il lunedì in formato AAAA-MM-GG.')
    else:
        # Solo per i record storici del vecchio modello; le nuove valutazioni richiedono week_start.
        try:
            month = datetime.strptime(d.get('month', ''), '%Y-%m')
            if month.strftime('%Y-%m') != d['month'] or month.date() > date.today().replace(day=1):
                errors.append('Il mese storico deve essere valido e non futuro.')
        except (TypeError, ValueError, KeyError):
            errors.append('Periodo mancante: indicare una settimana ISO valida.')
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
    for key, hi in [('hours', 10000000), ('workers', 100000), ('days', 7 if is_weekly_assessment(d) else 31)]:
        v = d.get(key)
        if v is not None and (isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or not 0 < v <= hi):
            errors.append(f'{key}: inserire un valore positivo oppure lasciare il dato non disponibile.')
    if d.get('days') and isinstance(d['days'], (int, float)) and month is not None:
        if d['days'] > calendar.monthrange(month.year, month.month)[1]:
            errors.append('Giorni di attività superiori ai giorni del mese storico.')
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


def calculate_legacy_risk(d):
    """Riproduzione del modello mensile 2.0.0, solo per storico/backup preesistenti.

    Pesi ipotetici di policy, versionati. Nessuna stima statistica degli infortuni.
    Indice base sempre confrontabile; contesto opzionale separato.
    """
    nc_load, crit_load = anomaly_load(d['nc_items']), anomaly_load(d['crit_items'])
    nc = 100 * (1 - math.exp(-nc_load / 10))
    crit = 100 * (1 - math.exp(-crit_load / 12))
    activity = min(100, max(ACTIVITY[a] for a in d['activities']) + 3 * (len(d['activities']) - 1))
    interference = (d['interference_level'] - 1) / 9 * 100
    org = 100 * (1 - math.exp(-(max(0, d['app'] - 1) + 1.2 * d['sub']) / 8))
    # Appaltatori e subappaltatori restano input descrittivi: non generano punti.
    # Nessun requisito universale "un HSE per impresa": si valuta la copertura reale dichiarata.
    inspection_deficit = 100 * (1 - min(1, d['inspections'] / d['inspection_target']))
    controls = .65 * (100 - d['hse_coverage']) + .35 * inspection_deficit
    components = dict(zip(LEGACY_WEIGHTS, [nc, crit, activity, interference, org, controls]))
    drivers = [{'Area': key, 'Indice area': round(value, 3), 'Peso massimo': LEGACY_WEIGHTS[key],
                'Contributo': value * LEGACY_WEIGHTS[key] / 100} for key, value in components.items()]
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
        f'Presidio: 65% deficit copertura HSE + 35% deficit ispezioni rispetto al riferimento gestionale ({d["inspection_target"]}); peso 15%.',
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
    return dict(version=LEGACY_MODEL_VERSION, score=round(score, 2), level=risk_level(score),
                raw=round(raw, 3), floor=floor, floor_reason=floor_reason,
                drivers=drivers, awareness_details=awareness_details, awareness_bonus=round(awareness, 3),
                needs=needs, reasons=reasons, context_score=round(contextual, 2) if contextual is not None else None,
                context_target=round(contextual_target, 2) if contextual_target else None, context_basis=context_basis,
                explanations=explanations, missing=missing)


def calculate_risk(d):
    """Pesi ipotetici di policy, versionati. Nessuna stima statistica degli infortuni.
    Indice base sempre confrontabile; contesto opzionale separato.
    """
    nc_load, crit_load = anomaly_load(d['nc_items']), anomaly_load(d['crit_items'])
    nc = 100 * (1 - math.exp(-nc_load / 10))
    crit = 100 * (1 - math.exp(-crit_load / 12))
    activity = min(100, max(ACTIVITY[a] for a in d['activities']) + 3 * (len(d['activities']) - 1))
    interference = (d['interference_level'] - 1) / 9 * 100
    # Nessun requisito universale "un HSE per impresa": si valuta la copertura reale dichiarata.
    inspection_deficit = 100 * (1 - min(1, d['inspections'] / d['inspection_target']))
    controls = .65 * (100 - d['hse_coverage']) + .35 * inspection_deficit
    components = dict(zip(WEIGHTS, [nc, crit, activity, interference, controls]))
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
        contextual_target, context_basis = max(1., d['hours'] / 500), '1 ispezione ogni 500 ore lavorate nella settimana (parametro gestionale)'
    elif d.get('days'):
        contextual_target = max(1., d['days'] / 5 * max(1., (d.get('workers') or 50) / 50))
        context_basis = '1 ispezione ogni 5 giorni operativi e ogni 50 lavoratori medi (minimo fattore 1)'
    elif d.get('workers'):
        contextual_target, context_basis = max(1., d['workers'] / 50), '1 ispezione per 50 lavoratori medi nella settimana'
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
        f'NC formali: carico {nc_load:.2f}; indice area = 100 × (1 − exp(−carico/10)); massimo 40 punti.',
        f'Criticità informali: carico {crit_load:.2f}; indice area = 100 × (1 − exp(−carico/12)); massimo 8 punti.',
        'Carico anomalia = gravità (1/2,5/5) × tema (0,65–1,25) × stato (aperta 1; chiusa nei tempi 0,2; chiusa in ritardo 0,4) × scaduta (1,25) × recidiva (1,2).',
        'I temi modificano la priorità convenzionale; la gravità deve essere assegnata in base al potenziale danno del caso concreto.',
        'Attività: massima pericolosità convenzionale + 3 per ogni tipologia aggiuntiva, massimo 100; peso 22%.',
        'Interferenze: (livello − 1)/9 × 100, peso 15%. Non viene riapplicata la pericolosità delle attività.',
        'Numero di appaltatori e subappaltatori: dato descrittivo, nessun contributo automatico al Risk Index.',
        f'Presidio: 65% deficit copertura HSE + 35% deficit ispezioni rispetto al riferimento gestionale ({d["inspection_target"]}); peso 15%.',
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
    return (f"Sensibilizzazione — {theme.upper()}\nCantiere: {d['nome']} | {d['technology']} | Periodo: {period_label(d)}\n"
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
    for driver in sorted(assessment_drivers(r), key=lambda x: x['Contributo'], reverse=True):
        area = driver['Area']
        if driver['Contributo'] < 2:
            continue
        suggestions = {'Attività': ('Rivedere sequenze, mezzi e protezioni delle lavorazioni a maggior potenziale; verificare sul campo le misure.', 'Referenti operativi / preposti'),
            'Interferenze': ('Separare nel tempo o nello spazio le attività interferenti; aggiornare layout e coordinamento.', 'Referenti imprese / coordinamento di cantiere'),
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
    top = sorted(assessment_drivers(r), key=lambda x: x['Contributo'], reverse=True)[:3]
    nc_counts, crit_counts = anomaly_counts(d['nc_items']), anomaly_counts(d['crit_items'])
    text = (f"Il cantiere di {d['nome']}, tecnologia {d['technology']}, ha generato per il periodo di analisi "
            f"{analysis_period(d)} un Risk Index pari a {r['score']:.1f}/100, fascia {r['level']}. "
            f"Il cantiere si trova nella fase di {d['fase']}, con le seguenti lavorazioni: "
            + ', '.join(d['activities']) + '. ')
    text += (f"Sono state inserite {nc_counts['total']} NC formali ({nc_counts['open']} aperte e {nc_counts['closed']} chiuse) "
             f"e {crit_counts['total']} criticità informali ({crit_counts['open']} aperte e {crit_counts['closed']} chiuse). ")
    themes = sorted({x['theme'] for x in d['nc_items'] + d['crit_items']})
    if themes:
        text += 'Le segnalazioni riguardano: ' + ', '.join(themes) + '. '
    text += (f"Le ispezioni registrate in HSEQ4ALL nel periodo sono {d['inspections']}. "
             f"Operano {d['app']} imprese appaltatrici e {d['sub']} subappaltatrici, "
             f"con una copertura HSE effettiva dichiarata del {d['hse_coverage']:.0f}%. ")
    text += 'I contributi principali al punteggio sono ' + ', '.join(f"{x['Area']} ({x['Contributo']:.1f} punti)" for x in top) + '. '
    if r['floor']:
        text += f"È attiva la soglia minima {r['floor']:.0f}: {r['floor_reason']}. "
    text += (f"Sono dichiarati {d['stop_positive']} Stop Work positivi e {d['stop_negative']} negativi. "
             f"Le sensibilizzazioni pertinenti riducono l’indice di {r['awareness_bonus']:.1f} punti.")
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
        result = assessment_calculator(sample)
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
                conn.execute('CREATE TABLE IF NOT EXISTS hse_weekly_plans (id TEXT PRIMARY KEY, plan_key TEXT NOT NULL, revision INTEGER NOT NULL, payload TEXT NOT NULL, UNIQUE(plan_key,revision))')

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


    def all_weekly(self):
        if self.mode=='session':
            return deepcopy(st.session_state.setdefault('weekly_records',[]))
        with self.connect() as conn:
            return [json.loads(x[0]) for x in conn.execute('SELECT payload FROM hse_weekly_plans').fetchall()]

    def save_weekly(self,record):
        self.import_weekly([record])

    def import_weekly(self,records):
        existing=self.all_weekly()
        by_id={x['id']:x for x in existing}
        by_revision={(x['plan_key'],x['revision']):x for x in existing}
        fresh=[]
        for record in records:
            key=(record['plan_key'],record['revision'])
            if record['id'] in by_id:
                if by_id[record['id']]!=record:raise ValueError('Conflitto identificativo.')
                continue
            if key in by_revision:raise ValueError('Conflitto revisione.')
            by_id[record['id']]=record;by_revision[key]=record;fresh.append(record)
        if self.mode=='session':
            st.session_state.weekly_records=existing+deepcopy(fresh)
            return
        token='%s' if self.mode=='postgres' else '?'
        sql='INSERT INTO hse_weekly_plans (id,plan_key,revision,payload) VALUES ('+','.join([token]*4)+')'
        with self.connect() as conn:
            for record in fresh:
                conn.execute(sql,(record['id'],record['plan_key'],record['revision'],json.dumps(record,ensure_ascii=False,allow_nan=False)))


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
    records, seen = [], set()
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
        d = old.get('data')
        version = old.get('model_version')
        if version not in (MODEL_VERSION, LEGACY_MODEL_VERSION):
            raise ValueError(f'Record {i+1}: versione modello non supportata.')
        if not isinstance(d, dict) or (version == MODEL_VERSION) != is_weekly_assessment(d):
            raise ValueError(f'Record {i+1}: periodo incompatibile con la versione modello.')
        errors = validate_data(d)
        if errors:
            raise ValueError(f'Record {i+1}: ' + '; '.join(errors[:5]))
        # Non fidarsi di risultati calcolati all'esterno; ricalcolare col modello originale.
        records.append(dict(id=ident, saved_at=timestamp, model_version=version,
                            data=d, result=assessment_calculator(d)))
    return records


def history_frame(records, latest_only=True):
    rows = []
    for x in records:
        d = x['data']
        weekly = is_weekly_assessment(d)
        key = period_key(d)
        rows.append(dict(id=x['id'], Cantiere=d['nome'], site=site_key(d['nome']),
                         Tecnologia=d['technology'], Periodo=key,
                         Frequenza='Settimanale' if weekly else 'Mensile storico',
                         Indice=x['result']['score'], Livello=x['result']['level'],
                         Contestualizzato=x['result']['context_score'],
                         Salvato=x['saved_at'], Versione=x['model_version']))
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    frame = frame.sort_values(['Periodo', 'Salvato', 'id'])
    if latest_only:
        frame = frame.drop_duplicates(['site', 'Periodo', 'Versione'], keep='last')
    return frame


def word_report(d, r, theme=None):
    doc = Document()
    doc.add_heading('HSE Risk Platform | ' + ('Toolbox Talk' if theme else 'Report valutazione'), 0)
    doc.add_paragraph(f"{d['nome']} · {d['technology']} · {period_label(d)} · Modello {r['version']}")
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
          'Sintesi': pd.DataFrame([{'Cantiere': d['nome'], 'Tecnologia': d['technology'], 'Periodo': period_label(d),
             'Indice base': r['score'], 'Livello': r['level'], 'Indice contestualizzato': r['context_score'], 'Versione modello': r['version'], 'NC totali': len(d['nc_items']), 'NC aperte': anomaly_counts(d['nc_items'])['open'], 'NC chiuse': anomaly_counts(d['nc_items'])['closed']}]),
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


def _visual_font(size, bold=False):
    from PIL import ImageFont
    candidates = (['DejaVuSans-Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
                   '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf', 'C:/Windows/Fonts/arialbd.ttf'] if bold else
                  ['DejaVuSans.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
                   '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf', 'C:/Windows/Fonts/arial.ttf'])
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            pass
    return ImageFont.load_default(size=size)


def output_visuals(d, r):
    """Immagini PNG dei risultati calcolati. Stesse grandezze dei grafici e delle scorecard del tool."""
    from PIL import Image, ImageDraw
    ink, muted, teal, pale = '#16354B', '#536B7D', '#087F8C', '#EAF1F6'
    result = {}

    def base(title, subtitle):
        im = Image.new('RGB', (1600, 900), 'white')
        draw = ImageDraw.Draw(im)
        draw.text((64, 38), title, font=_visual_font(46, True), fill=ink)
        draw.text((66, 107), subtitle, font=_visual_font(24), fill=muted)
        return im, draw

    def save(name, im):
        out = BytesIO()
        im.save(out, format='PNG')
        result[name] = out.getvalue()

    im, draw = base('Risk Index', f"{period_label(d)} - indice gestionale su scala 0-100")
    score = bounded(r['score'])
    color = COLORS[min(4, max(0, math.ceil(score / 20) - 1))]
    draw.arc((100, 185, 690, 775), 0, 360, fill=pale, width=54)
    if score > 0:
        draw.arc((100, 185, 690, 775), -90, -90 + score * 3.6, fill=color, width=54)
    draw.text((395, 430), f"{score:.1f}", anchor='mm', font=_visual_font(112, True), fill=color)
    draw.text((395, 535), 'SU 100', anchor='mm', font=_visual_font(31), fill=muted)
    draw.text((810, 235), r['level'], font=_visual_font(64, True), fill=color)
    for i, t in enumerate(['Valore alto = maggiore attenzione HSE', 'Indice base per confronti omogenei', 'Nessuna probabilità di infortunio stimata']):
        draw.text((815, 355 + i * 65), t, font=_visual_font(28), fill=ink)
    context = 'Dati facoltativi non inseriti' if r['context_score'] is None else f"Indice contestualizzato: {r['context_score']:.1f}/100"
    draw.text((815, 590), context, font=_visual_font(27, True), fill=teal)
    if r['floor']:
        draw.text((815, 650), f"Soglia minima attiva: {r['floor']:.0f}", font=_visual_font(28, True), fill='#B83246')
    draw.text((66, 840), 'Basso 0-20   Medio basso >20-40   Medio >40-60   Alto >60-80   Critico >80-100', font=_visual_font(24), fill=muted)
    save('risk', im)

    im, draw = base('Contributi al Risk Index', 'Punti effettivi, incluse riduzioni e soglie minime')
    rows = r['drivers']
    lo = min(-1., min(x['Contributo'] for x in rows)); hi = max(1., max(x['Contributo'] for x in rows))
    left, right = 535, 1410
    project = lambda value: left + (value - lo) / (hi - lo) * (right - left)
    zero = project(0)
    draw.line((zero, 173, zero, 830), fill='#CCD9E2', width=2)
    for i, row in enumerate(rows):
        y = 182 + i * 58
        draw.text((64, y+4), row['Area'], font=_visual_font(25), fill=ink)
        value = row['Contributo']; x = project(value)
        if abs(value) > .00001:
            draw.rounded_rectangle((min(zero, x), y, max(zero, x), y+33), radius=5, fill='#16826B' if value < 0 else '#225873')
        else:
            draw.ellipse((zero-3, y+13, zero+3, y+19), fill='#9DAFBD')
        draw.text((1460, y+3), f'{value:+.2f}', font=_visual_font(25, True), fill=ink)
    save('drivers', im)

    im, draw = base('Segnalazioni del periodo', 'NC formali e criticità informali - conteggi per stato')
    series = [anomaly_counts(d['nc_items']), anomaly_counts(d['crit_items'])]
    limit = max(1, *(x[k] for x in series for k in ['open', 'closed']))
    for i, (title, counts) in enumerate(zip(['NC formali', 'Criticità informali'], series)):
        x = 90 + i * 785
        draw.text((x, 202), title, font=_visual_font(36, True), fill=ink)
        draw.text((x, 267), f"Totale {counts['total']}", font=_visual_font(31), fill=muted)
        for j, (key, label, col) in enumerate([('open', 'Aperte', '#D46B22'), ('closed', 'Chiuse', '#16826B')]):
            y = 385 + j * 170
            draw.text((x, y-50), f"{label}: {counts[key]}", font=_visual_font(31, True), fill=col)
            draw.rounded_rectangle((x, y, x+590, y+44), radius=7, fill=pale)
            if counts[key]:
                draw.rounded_rectangle((x, y, x+590*counts[key]/limit, y+44), radius=7, fill=col)
    draw.text((90, 820), 'Le chiuse mantengono un contributo ridotto. Il registro conserva tutti i dettagli.', font=_visual_font(25), fill=muted)
    save('anomalies', im)

    im, draw = base('HSE Scorecard', 'Più stelle = minore indice di attenzione dell’area. Scala da 0 a 5 stelle.')
    def star(cx, cy, filled):
        points = [(cx + (23 if i % 2 == 0 else 10) * math.cos(-math.pi/2 + i*math.pi/5),
                   cy + (23 if i % 2 == 0 else 10) * math.sin(-math.pi/2 + i*math.pi/5)) for i in range(10)]
        draw.polygon(points, fill='#B38212' if filled else '#E2E8ED')
    for i, row in enumerate(assessment_drivers(r)):
        x, y = 65 + (i % 2)*790, 190 + (i//2)*220
        draw.text((x, y), row['Area'], font=_visual_font(31, True), fill=ink)
        count, _ = scorecard_stars(row['Indice area'])
        for j in range(5):
            star(x+24+j*59, y+80, j<count)
        draw.text((x, y+125), f"Indice area {row['Indice area']:.1f}/100   Contributo {row['Contributo']:.2f} pt", font=_visual_font(25), fill=muted)
    save('scorecard', im)
    return result


def ppt_report(d, r):
    """Export 16:9, con riepilogo dinamico e immagini PNG dei risultati, senza servizi esterni."""
    from pptx.enum.shapes import MSO_SHAPE
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    visuals = output_visuals(d, r)
    navy, teal, muted = '16354B', '087F8C', '536B7D'

    def text(slide, value, x, y, w, h, size=20, bold=False, color=navy):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(.02)
        tf.margin_top = tf.margin_bottom = Inches(.02)
        for i, line in enumerate(value.split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line; p.font.name = 'Aptos'; p.font.size = PptPt(size)
            p.font.bold = bold; p.font.color.rgb = RGBColor.from_string(color)
            p.space_after = PptPt(3); p.line_spacing = 1.12
        return box

    def slide(title, subtitle=''):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor.from_string('FFFFFF')
        text(s, title, .65, .4, 12.0, .7, 32, True)
        if subtitle:
            text(s, subtitle, .68, 1.15, 12., .45, 14, color=muted)
        text(s, f"HSE Risk Platform {UI_VERSION}   Modello {r['version']}   {period_key(d)}", .68, 7.1, 11., .22, 10, color=muted)
        text(s, str(len(prs.slides)), 12.2, 7.1, .4, .22, 10, color=muted)
        return s

    def picture_slide(title, key, subtitle=''):
        s = slide(title, subtitle)
        s.shapes.add_picture(BytesIO(visuals[key]), Inches(2.026), Inches(1.7), width=Inches(9.28))
        return s

    def wrap_measured(value, size=19, inches=11.9):
        # Misura conservativa in punti; protegge anche da parole/identificativi senza spazi.
        font = _visual_font(size * 4)
        limit = (inches * 72 - 8) * 4 * .93
        lines = []
        for paragraph in str(value).splitlines() or ['']:
            line = ''
            for word in paragraph.split():
                candidate = (line + ' ' + word).strip()
                if font.getlength(candidate) <= limit:
                    line = candidate
                    continue
                if line:
                    lines.append(line); line = ''
                for char in word:
                    if line and font.getlength(line + char) > limit:
                        lines.append(line); line = char
                    else:
                        line += char
            lines.append(line)
        return lines

    def paragraphs_slides(title, paragraphs, max_lines=13):
        # Mantiene intere le azioni brevi; divide solo i paragrafi troppo lunghi per una slide.
        pages, current = [], []
        for paragraph in paragraphs:
            block = wrap_measured(paragraph)
            if current and len(current) + 1 + len(block) > max_lines:
                pages.append(current); current = []
            while len(block) > max_lines:
                pages.append(block[:max_lines]); block = block[max_lines:]
            if block:
                if current:
                    current.append('')
                current.extend(block)
        if current:
            pages.append(current)
        for i, lines in enumerate(pages or [['Nessuna voce presente.']]):
            s = slide(title + (' - continua' if i else ''))
            text(s, '\n'.join(lines), .7, 1.65, 11.9, 5.25, 19)

    # La copertina contiene il riepilogo richiesto; il nome completo resta nel testo.
    cover = slide('Sintesi del cantiere', f"{d['technology']}   {analysis_period(d)}")
    summary = executive_summary(d, r).split('I contributi principali al punteggio sono ')[0].strip()
    summary_lines = wrap_measured(summary, size=18)
    text(cover, '\n'.join(summary_lines), .7, 1.85, 11.9, 4.85,
         18 if len(summary_lines) <= 13 else 16)
    # Le eventuali informazioni di qualità del dato sono riportate nelle pagine di lettura.
    picture_slide('Risk Index del periodo', 'risk')
    picture_slide('Analisi dei contributi', 'drivers')
    picture_slide('NC formali e criticità informali', 'anomalies')
    picture_slide('Scorecard delle aree', 'scorecard')
    detect = detection_details(d)
    paragraphs_slides('Ispezioni e detection', [
        f"Ispezioni registrate in HSEQ4ALL: {d['inspections']} nel periodo {period_key(d)}.",
        f"Riferimento gestionale {detect['target']}. Detection convenzionale {detect['coverage']:.0f}%. "
        f"Deficit {detect['deficit']:.0f}%. Contributo del deficit ispettivo: +{detect['points']:.2f} punti.",
        'La detection è un indicatore convenzionale di copertura del riferimento ispettivo. Il rapporto NC/ispezioni resta descrittivo.',
        *r['missing']])
    actions = action_plan(d, r)
    paragraphs_slides('Priorità operative', [f"{x['Priorità']} - {x['Area']}: {x['Azione']}\n"
        f"Responsabile proposto: {x['Responsabile']}. Termine proposto: {x['Termine']}." for x in actions]
        or ['Nessuna azione specifica generata. Mantenere il presidio e verificare le misure.'])
    for field, title in [('nc_items', 'Registro NC formali'), ('crit_items', 'Registro criticità informali')]:
        counts = anomaly_counts(d[field])
        paragraphs_slides(title, [f"Totale {counts['total']}. Aperte {counts['open']}. Chiuse {counts['closed']}."] +
            [f"{i}. {e['theme']} / gravità {e['severity']} / {e['status']}: {e.get('note', '')}" for i,e in enumerate(d[field],1)])
    for theme in sorted(r['needs'], key=r['needs'].get, reverse=True)[:3]:
        title, points, question = TOOLBOX_CONTENT[theme]
        paragraphs_slides('Toolbox su ' + theme, [title, *points, question])
    paragraphs_slides('Metodo e lettura del risultato', [
        'Il Risk Index è un indice gestionale di priorità HSE su scala 0-100. Un valore maggiore richiede maggiore attenzione.',
        'Pesi, soglie e target sono convenzioni gestionali da calibrare. Il risultato non è una probabilità di infortunio.',
        f"Soglia minima: {r['floor']:.0f}. {r['floor_reason']}.",
        'Indice base per confronti omogenei. Indice contestualizzato separato per le esposizioni facoltative.'])
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
.hero{padding:26px 0 30px}.hero p{font-size:1.2rem;color:#415c70}.eyebrow{font-size:.8rem;letter-spacing:.18em;color:#087f8c;font-weight:700}
[data-testid="stVerticalBlockBorderWrapper"]{background:#fff;border-color:#d8e3eb!important;box-shadow:0 3px 14px #12304708}
[data-testid="stButton"] button{min-height:44px;font-weight:600}
[data-testid="stPills"] button{border-radius:8px}
.week-grid{display:grid;grid-template-columns:repeat(7,minmax(130px,1fr));gap:9px;margin:20px 0;overflow-x:auto}
.week-day{background:#eaf0f5;border:1px solid #d7e1ea;border-radius:10px;padding:12px;min-width:130px}
.week-day h4{font-size:1rem;margin:0 0 12px;color:#123047}
.event{background:#fff;border-left:4px solid #087f8c;border-radius:6px;padding:10px;margin-bottom:10px;overflow-wrap:anywhere;font-size:.88rem}
.event span{display:inline-block;margin-top:6px;color:#225873}.event small,.muted{color:#415c70;font-size:.8rem}
.risk-overview{display:grid;grid-template-columns:220px 1fr;gap:30px;background:#fff;border:1px solid #dbe5ec;border-radius:20px;padding:28px;margin:18px 0 24px;box-shadow:0 10px 30px #12304709}
.risk-number{font-size:4.4rem;line-height:1;font-weight:800;letter-spacing:-.06em;color:var(--risk-color)}
.risk-number small{font-size:1.2rem;letter-spacing:0;color:#607384;font-weight:500}
.risk-label{font-size:.78rem;letter-spacing:.13em;font-weight:700;color:#526879;margin-bottom:14px}
.risk-level{color:var(--risk-color);font-size:1.1rem;font-weight:750;margin-top:12px}
.risk-explanation h3{margin:0 0 8px;font-size:1.35rem}.risk-explanation p{color:#40596c;line-height:1.6;margin:0 0 10px}
.risk-spectrum{height:10px;position:relative;border-radius:8px;margin:20px 6px 9px;background:linear-gradient(to right,#16826b 0% 20%,#4b8b9e 20% 40%,#b38212 40% 60%,#d46b22 60% 80%,#b83246 80% 100%)}
.risk-marker{position:absolute;top:-4px;width:5px;height:18px;background:#123047;border:1px solid white;border-radius:3px;transform:translateX(-50%)}
.risk-scale{display:flex;justify-content:space-between;color:#536b7d;font-size:.75rem}
.score-card{padding:22px;background:white;border:1px solid #dbe5ec;border-radius:14px;margin:0 0 16px;border-top:4px solid var(--area-color)}
.score-card h4{margin:0 0 12px;font-size:1.02rem;color:#123047}.score-card .stars{font-size:1.8rem;color:#aa7704;letter-spacing:4px;line-height:1.4}.score-card p{color:#496174;font-size:.88rem;margin:8px 0 0}
.score-card strong{font-size:1.4rem;color:#123047}.report-heading{margin-top:16px}
@media(max-width:700px){.risk-overview{grid-template-columns:1fr;padding:20px;gap:18px}.risk-number{font-size:3.8rem}}
@media(max-width:850px){.week-grid{grid-template-columns:repeat(2,minmax(140px,1fr))}}
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


def add_editor_card(state_key, awareness=False):
    row = {'theme': None, 'count': 1, 'coverage': None} if awareness else {
        'theme': None, 'severity': None, 'status': 'Aperta', 'late': False, 'repeat': False, 'note': ''}
    st.session_state[state_key].append({'uid': uuid.uuid4().hex, **row})


def remove_editor_card(state_key, uid):
    st.session_state[state_key] = [x for x in st.session_state[state_key] if x['uid'] != uid]


def card_editor(label, rows, key, awareness=False):
    state_key = key + '_cards'
    if state_key not in st.session_state:
        st.session_state[state_key] = [{'uid': uuid.uuid4().hex, **deepcopy(x)} for x in rows]
    section_help = None if awareness else NC_FORMAL_HELP if key.startswith('nc_') else NC_INFORMAL_HELP
    st.subheader(label, help=section_help)
    entries = st.session_state[state_key]
    if not entries:
        st.info('Nessuna voce inserita. Usa il pulsante Aggiungi qui sotto per creare la prima scheda.')
    result = []
    for index, item in enumerate(entries):
        uid = item['uid']; prefix = key + '_' + uid
        with st.container(border=True):
            title, remove = st.columns([5, 1])
            title.markdown(f"**{'\U0001f9e0 Sensibilizzazione' if awareness else '\U0001f4dd Segnalazione'} {index + 1}**")
            remove.button('Rimuovi', key=prefix+'_remove', icon='\U0001f5d1\ufe0f', on_click=remove_editor_card, args=(state_key,uid))
            theme = st.selectbox('Argomento' if awareness else 'Tema della segnalazione', THEMES,
                index=THEMES.index(item['theme']) if item.get('theme') in THEMES else None,
                placeholder='Seleziona un tema', key=prefix+'_theme')
            if awareness:
                a,b = st.columns(2)
                count = a.number_input('Incontri effettuati nella settimana',1,1000,int(item.get('count',1)),key=prefix+'_count')
                known = b.checkbox('Conosco la copertura dei destinatari',value=item.get('coverage') is not None,key=prefix+'_known')
                coverage = st.slider('Destinatari unici raggiunti (%)',0,100,int(item.get('coverage') or 0),key=prefix+'_coverage') if known else None
                row = dict(theme=theme,count=count,coverage=coverage)
            else:
                a,b = st.columns(2)
                severity = a.selectbox('Gravità',list(SEVERITY),index=list(SEVERITY).index(item['severity']) if item.get('severity') in SEVERITY else None,
                    placeholder='Seleziona la gravità',key=prefix+'_severity')
                status = b.selectbox('Stato a fine settimana',list(STATUS),index=list(STATUS).index(item.get('status','Aperta')),key=prefix+'_status')
                late = st.checkbox('Aperta e oltre la scadenza',value=bool(item.get('late')) if status=='Aperta' else False,
                    disabled=status!='Aperta',key=prefix+'_late') if status=='Aperta' else False
                repeat = st.checkbox('Segnalazione ricorrente',value=bool(item.get('repeat')),key=prefix+'_repeat')
                note = st.text_area('Descrizione e riferimento',value=item.get('note',''),max_chars=1000,key=prefix+'_note')
                row = dict(theme=theme,severity=severity,status=status,late=late,repeat=repeat,note=note)
            item.update(row)
            result.append(row)
    st.button('Aggiungi sensibilizzazione' if awareness else 'Aggiungi NC formale' if key.startswith('nc_') else 'Aggiungi criticità',
        icon='\u2795',key=key+'_add',on_click=add_editor_card,args=(state_key,awareness),type='secondary',width='stretch')
    return result


def anomaly_editor(label, rows, key):
    return card_editor(label,rows,key)

def render_form(records):
    if 'draft' not in st.session_state:
        st.session_state.draft = default_data()
    if 'week_start' not in st.session_state.draft:
        st.session_state.draft = next_weekly_draft(st.session_state.draft)
    d = st.session_state.draft
    rev = st.session_state.get('draft_revision', 0)
    st.title('\U0001f4dd Valutazione settimanale')
    st.caption('WIND / SOLAR / BESS · Input riferiti esclusivamente alla settimana selezionata')
    st.info('Inserisci i dati del solo periodo selezionato. NC formali: registrate in HSEQ4ALL nel periodo. '
            'Criticità informali: rilevate nel periodo e non registrate in HSEQ4ALL. Ogni anomalia va inserita una sola volta.')
    names = sorted({x['data']['nome'] for x in records}, key=str.casefold)
    if names:
        with st.expander('Usa il nome di un cantiere già presente'):
            chosen = st.selectbox('Cantieri registrati', names)
            if st.button('Usa questo nome'):
                st.session_state.draft = {**d, 'nome': chosen}
                st.session_state.draft_revision = rev + 1
                st.rerun()
    with st.container(border=True):
        st.subheader('\U0001f4cd 1 · Cantiere e periodo')
        a, b, c = st.columns([2, 1, 1])
        name = a.text_input('Nome cantiere', d['nome'], max_chars=120, help='Usare sempre lo stesso nome per collegare le settimane nello storico.')
        technology = b.selectbox('Tecnologia', ['Wind', 'Solar', 'BESS'], index=['Wind', 'Solar', 'BESS'].index(d['technology']))
        chosen_day = c.date_input('Settimana di riferimento (seleziona un giorno)',
            date.fromisoformat(d['week_start']), max_value=date.today(),
            help='Qualunque giorno selezionato viene ricondotto al lunedì. La settimana termina sempre domenica.')
        monday = week_monday(chosen_day)
        sunday = monday + timedelta(days=6)
        iso_year, iso_week, _ = monday.isocalendar()
        st.info(f'Settimana ISO {iso_week:02d}/{iso_year}: {monday:%d/%m/%Y} – {sunday:%d/%m/%Y} (lunedì–domenica).')
        phase = st.selectbox('Fase operativa prevalente', list(PHASE_THEMES), index=list(PHASE_THEMES).index(d['fase']))
        if sunday >= date.today():
            st.caption('Settimana in corso: dati parziali, esclusi dalla graduatoria delle settimane concluse.')
        st.subheader('\U0001f3d7\ufe0f 2 · Lavorazioni e interferenze')
        activities = st.pills('Seleziona le lavorazioni della settimana', list(ACTIVITY), default=d['activities'], selection_mode='multi')
        interference = st.slider('Interferenza fra attività', 1, 10, int(d['interference_level']),
          help='1: attività separate. 3: contatti occasionali gestiti. 5: aree/risorse condivise. 7: sovrapposizioni frequenti. 10: interferenze continue senza separazione efficace.')
        st.subheader('\U0001f50d 3 · Anomalie e follow-up')
        st.caption('Usa Aggiungi per ogni segnalazione e compila la scheda. Aperte = presenti a fine settimana; chiuse = risolte durante la settimana. Gravità alta = potenziale danno grave, da motivare nella descrizione. '
                   'Criticità informali: peso inferiore, nessuna soglia automatica. Se formalizzate, rimuoverle dalla tabella informale.')
        nc = anomaly_editor('\U0001f4dd NC formali', d['nc_items'], f'nc_{rev}')
        crit = anomaly_editor('\U0001f4cc Criticità informali', d['crit_items'], f'crit_{rev}')
        danger = st.checkbox('È ancora presente un pericolo grave e attuale a fine valutazione', value=d['imminent_danger'],
           help='Attiva un allarme dedicato e un indice minimo 81. Non dipende dal canale formale/informale di registrazione.')
        st.subheader('\U0001f477 4 · Organizzazione e presidio')
        a, b, c = st.columns(3)
        app = a.number_input('Imprese appaltatrici operative', 0, 200, int(d['app']))
        sub = b.number_input('Imprese subappaltatrici operative', 0, 200, int(d['sub']))
        hse = c.number_input('Professionisti HSE delle imprese', 0, 200, int(d['company_hse']), help='Dato descrittivo; non equivale a copertura reale.')
        coverage = st.slider('Copertura HSE effettiva delle lavorazioni (%)', 0, 100, int(d['hse_coverage']),
          help='Quota delle lavorazioni/turni/aree per cui è stato assicurato il presidio HSE pianificato. Stima dichiarata, non numero di persone e non criterio normativo universale.')
        target = DEFAULT_INSPECTION_TARGET
        inspections = st.number_input('Numero di ispezioni registrate in HSEQ4ALL', 0, 2000, int(d['inspections']),
                                      help=detection_help(target))
        st.subheader('\u26d4 5 · Stop Work')
        st.caption('Positivo: interruzione preventiva tempestiva e gestita. Negativo: stop imposto per carenza o mancato rispetto di misure. '
                   'Classificare ciascun evento in una sola categoria; registrare una sola volta anche l’anomalia sottostante. La classificazione resta responsabilità dell’utilizzatore.')
        a, b = st.columns(2)
        positive = a.number_input('Stop Work positivi', 0, 1000, int(d['stop_positive']), help=STOP_POSITIVE_HELP)
        negative = b.number_input('Stop Work negativi', 0, 1000, int(d['stop_negative']), help=STOP_NEGATIVE_HELP)
        st.subheader('\U0001f9e0 6 · Sensibilizzazioni mirate')
        st.caption('Una scheda per tema. Copertura = % di lavoratori destinatari unici raggiunti sul tema nella settimana (non somma delle presenze). '
                   'Lasciare vuota se sconosciuta: fattore prudenziale 50%. Senza collegamento ad attività, anomalie aperte o fase non c’è bonus.')
        awareness = card_editor('\U0001f9e0 Sensibilizzazioni effettuate', d['awareness'], f'awareness_{rev}', awareness=True)
        with st.expander('\U0001f4d0 7 · Dati di esposizione FACOLTATIVI'):
            st.write('Puoi lasciare tutto a zero: significa dato non disponibile. Questi dati producono un indice contestualizzato separato, '
                     'senza modificare l’indice base usato nelle graduatorie. I rapporti adottati sono ipotesi gestionali da calibrare.')
            a, b, c = st.columns(3)
            hours = a.number_input('Ore lavorate complessive della settimana', 0., 10000000., float(d['hours'] or 0), step=100.)
            workers = b.number_input('Numero medio lavoratori nella settimana', 0, 100000, int(d['workers'] or 0))
            days = c.number_input('Giorni di attività nella settimana', 0, 7, int(d['days'] or 0))
        notes = st.text_area('Note della valutazione', d['notes'], max_chars=5000)
        submitted = st.button('Calcola la valutazione', icon='\U0001f9ee', type='primary', width='stretch')
    if submitted:
        candidate = dict(nome=name.strip(), technology=technology, fase=phase, week_start=monday.isoformat(),
            app=app, sub=sub, company_hse=hse, hse_coverage=coverage, inspections=inspections,
            inspection_target=target, activities=activities, interference_level=interference,
            nc_items=nc, crit_items=crit, awareness=awareness, stop_positive=positive, stop_negative=negative,
            imminent_danger=danger, hours=hours or None, workers=workers or None, days=days or None, notes=notes,
            input_scope='week_registrations_v1')
        errors = validate_data(candidate)
        if errors:
            for error in errors:
                st.error(error)
        else:
            st.session_state.draft = candidate
            st.session_state.pop('trial_discarded', None)
            st.session_state.pop('exports', None)
            st.session_state.pop('advisor_reply', None)
            st.session_state.pop('report_section', None)
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


REPORT_SECTIONS = ['\U0001f4ca Sintesi del periodo', '\U0001f3af Azioni e alert', '\U0001f4cb Action Plan',
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
    st.subheader(f"{d['nome']} · {d['technology']} · {period_label(d)}")
    st.caption(f"Interfaccia {UI_VERSION} · modello {record['model_version']} · snapshot calcolato: dopo modifiche al modulo premi nuovamente Calcola.")
    color = COLORS[min(4, max(0, math.ceil(r['score'] / 20) - 1))]
    interpretation = {'Basso':'Mantenere il presidio e verificare le condizioni operative.',
                      'Medio basso':'Concentrare il follow-up sulle aree che contribuiscono maggiormente.',
                      'Medio':'Pianificare gli interventi sulle priorità e verificarne l’attuazione.',
                      'Alto':'Dare priorità alle condizioni aperte e rafforzare il presidio operativo.',
                      'Critico':'Gestire subito le condizioni prioritarie con i responsabili competenti.'}[r['level']]
    st.markdown(f'''<section class="risk-overview" style="--risk-color:{color}">
      <div><div class="risk-label">RISK INDEX · INDICE BASE</div>
      <div class="risk-number">{r['score']:.1f}<small> / 100</small></div>
      <div class="risk-level">{html.escape(r['level'])}</div></div>
      <div class="risk-explanation"><h3>Come leggere il risultato</h3>
      <p>Il Risk Index sintetizza la priorità gestionale HSE del cantiere nel periodo analizzato.
      Un valore più alto indica maggiore attenzione richiesta. Non è una percentuale di probabilità di infortunio.</p>
      <p><strong>{html.escape(interpretation)}</strong></p>
      <div class="risk-spectrum"><span class="risk-marker" style="left:{bounded(r['score'])}%"></span></div>
      <div class="risk-scale"><span>0 · Basso</span><span>40 · Medio</span><span>80 · Critico</span><span>100</span></div>
      </div></section>''', unsafe_allow_html=True)
    st.subheader('\U0001f3d7\ufe0f Il cantiere nel periodo di analisi')
    st.write(executive_summary(d, r))
    if r['context_score'] is not None:
        st.info(f"Indice contestualizzato: {r['context_score']:.1f}/100. Usa i dati di esposizione facoltativi. "
                'Per confrontare tutti i cantieri usa il Risk Index base.')
    else:
        st.caption('Indice contestualizzato non disponibile: non sono stati inseriti dati di esposizione. Il Risk Index base è calcolato.')
    st.subheader('\U0001f4cb Segnalazioni del periodo')
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
    if is_weekly_assessment(d) and date.fromisoformat(d['week_start']) + timedelta(days=6) >= date.today():
        st.warning('\U0001f4c5 Settimana in corso: dati parziali, esclusi dalla graduatoria consolidata.')
    is_saved = record['id'] in {x['id'] for x in records}
    if is_saved:
        st.success('\U0001f4be Valutazione già inserita nello storico.')
    else:
        st.info('\U0001f9ea MODALITÀ PROVA · input e output temporanei. In fondo alla pagina scegli Salva e inserisci oppure Non salvare.')
    st.divider()
    section = st.pills('\U0001f9ed Apri una sezione del report', REPORT_SECTIONS, selection_mode='single', default=REPORT_SECTIONS[0], key='report_section') or REPORT_SECTIONS[0]
    st.subheader(section)
    st.caption('12 sezioni dedicate: sintesi, alert, piano azioni, registro, motivazioni, toolbox, scenari, calcolo, advisor, scorecard, checklist ed export.')
    actions = action_plan(d, r)
    if section == REPORT_SECTIONS[0]:
        st.write('Le barre mostrano quanto ogni area contribuisce al Risk Index. I valori negativi rappresentano riduzioni. '
                 'La soglia minima, se attiva, può mantenere elevato il risultato anche in presenza di bonus.')
        st.altair_chart(driver_chart(r), width='stretch')
        st.caption('Contributi effettivi, incluse le rettifiche per limite di scala e soglia minima.')
        nc_driver = next(x for x in r['drivers'] if x['Area'] == 'NC formali')
        crit_driver = next(x for x in r['drivers'] if x['Area'] == 'Criticità informali')
        st.write(f"\U0001f4dd NC formali: {nc['total']} registrate, contributo {nc_driver['Contributo']:.2f} punti su {nc_driver['Peso massimo']}. "
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
        for item in sorted(assessment_drivers(r), key=lambda x: x['Contributo'], reverse=True):
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
            st.info('Nessuna sensibilizzazione registrata nel periodo.')
        themes = sorted(THEMES, key=lambda t: r['needs'].get(t, 0), reverse=True)
        theme = st.selectbox('\U0001f9f0 Genera un toolbox a tema · priorità in cima', themes)
        st.text(toolbox_text(theme, d, r))
        st.download_button('\U0001f4e5 Scarica toolbox Word', word_report(d, r, theme), f"Toolbox_{safe_filename(theme)}_{period_key(d)}.docx")
    elif section == REPORT_SECTIONS[6]:
        st.dataframe(pd.DataFrame(simulations(d, r)), hide_index=True, width='stretch')
        st.caption('Scenari indipendenti e non sommabili: simulano il punteggio, non la riduzione reale del rischio. '
                   'Chiusura significa misure attuate e verificate; il pericolo attuale resta attivo finché non viene rimosso e rivalutato.')
    elif section == REPORT_SECTIONS[7]:
        st.code('Base = clamp(Σ indice_area × peso / 100 + SW_neg − SW_pos − bonus_toolbox, 0, 100)\nFinale = max(Base, soglia_minima)')
        st.dataframe(pd.DataFrame(r['drivers']), hide_index=True, width='stretch')
        for text in r['explanations']:
            st.write('• ' + text)
        detect = detection_details(d)
        st.markdown('### \U0001f50e Detection delle ispezioni')
        st.write(f"{d['inspections']} ispezioni HSEQ4ALL nel periodo · riferimento {detect['target']} · "
                 f"copertura convenzionale {detect['coverage']:.0f}% · deficit {detect['deficit']:.0f}% · "
                 f"contributo del solo deficit ispettivo +{detect['points']:.2f} punti.")
        st.caption(detection_help(detect['target']))
        st.markdown('### \U0001f50e Dati facoltativi e qualità del dato')
        st.write(r['context_basis'])
        if r['context_target'] is not None:
            st.write(f"Target contestualizzato: {r['context_target']:.2f} ispezioni; cambia solo il deficit ispettivo nell’indice separato.")
        for text in r['missing']:
            st.warning(text)
        st.write('NC/ispezioni (descrittivo): ' + (f"{len(d['nc_items']) / d['inspections']:.2f}" if d['inspections'] else 'non valutabile'))
        if d['hours']:
            basis = ('NC registrate nel periodo' if d.get('input_scope') in ('month_registrations_v1','week_registrations_v1')
                     else 'valutazione precedente: può includere arretrato aperto')
            st.write(f"NC formali / 10.000 ore: {len(d['nc_items']) / d['hours'] * 10000:.2f}; {basis}. Indicatore descrittivo.")
    elif section == REPORT_SECTIONS[8]:
        st.caption('Assistente locale a regole, senza API AI. Spiega indice, azioni, presidio e temi dei toolbox.')
        q = st.text_input('\U0001f4ac Domanda sul report', max_chars=500)
        if st.button('\U0001f4ac Spiega il report'):
            st.session_state.advisor_reply = (record['id'], local_advisor(q, d, r))
        reply = st.session_state.get('advisor_reply')
        if reply and reply[0] == record['id']:
            st.markdown(reply[1])
    elif section == REPORT_SECTIONS[9]:
        st.caption('Le stelle sintetizzano il controllo convenzionale dell’area: più stelle = minore indice di attenzione. '
                   'Da 0 a 5 stelle, arrotondate rispetto a (100 − indice area) / 20. Il Risk Index segue la scala opposta: più alto = maggiore attenzione.')
        cols = st.columns(2)
        for i, row in enumerate(assessment_drivers(r)):
            filled, stars = scorecard_stars(row['Indice area'])
            area_color = COLORS[min(4, max(0, math.ceil(row['Indice area'] / 20) - 1))]
            with cols[i % 2]:
                st.markdown(f'''<article class="score-card" style="--area-color:{area_color}">
                  <h4>{html.escape(row['Area'])}</h4><div class="stars" aria-label="{filled} stelle su 5">{stars}</div>
                  <p>Indice di attenzione <strong>{row['Indice area']:.1f}/100</strong></p>
                  <p>Contributo al Risk Index <b>{row['Contributo']:.2f} punti</b> · peso {row['Peso massimo']}%</p>
                  </article>''', unsafe_allow_html=True)
    elif section == REPORT_SECTIONS[10]:
        defaults = list(dict.fromkeys(['Controlli generali pre-attività'] + [ACTIVITY_TO_CHECKLIST[a] for a in d['activities'] if a in ACTIVITY_TO_CHECKLIST]))
        sections = st.multiselect('\U0001f4d1 Capitoli della checklist', list(CHECKLIST_LIBRARY), default=defaults)
        if sections:
            st.download_button('\U0001f4e5 Scarica checklist Word', generate_checklist_word(d, sections).getvalue(), f"Checklist_{safe_filename(d['nome'])}.docx")
    elif section == REPORT_SECTIONS[11]:
        st.caption('Report temporanei fino al download volontario. Il backup della singola valutazione è disponibile dopo Salva e inserisci.')
        filename = safe_filename(d['nome']) + '_' + period_key(d)
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
        st.download_button('\U0001f4e5 Backup della valutazione salvata', backup_bytes([record]), f"{safe_filename(d['nome'])}_{period_key(d)}.json", 'application/json')
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
    site_frame = frame[frame['site'] == site_key(site)].copy()
    selected = site_frame[site_frame['Frequenza'] == 'Settimanale'].copy()
    if not selected.empty:
        selected['Data settimana'] = pd.to_datetime(selected['Periodo'])
        selected = selected.sort_values('Data settimana')
        st.caption('Trend settimanale: un punto per ogni settimana effettivamente valutata (ultima revisione). '
                   'Se manca una o più settimane, la linea collega direttamente le valutazioni precedente e successiva; '
                   'nessun punteggio o peso è assegnato alle settimane senza valutazione.')
        week_dates = selected['Data settimana'].dt.date.tolist()
        missing_weeks = []
        for before, after in zip(week_dates, week_dates[1:]):
            gap = (after - before).days // 7 - 1
            if gap > 0:
                missing_weeks.append(f'{before:%d/%m/%Y} → {after:%d/%m/%Y}: {gap} settimana/e senza valutazione')
        if missing_weeks:
            st.caption('Intervalli privi di rilevazione (nessuna imputazione): ' + '; '.join(missing_weeks[:8]) + ('; ...' if len(missing_weeks) > 8 else ''))
        target_values = {x['data']['inspection_target'] for x in records
                         if site_key(x['data']['nome']) == site_key(site) and is_weekly_assessment(x['data'])}
        if len(target_values) > 1:
            st.warning('Il target ispettivo è cambiato nello storico settimanale: parte del trend può dipendere da questo cambiamento di criterio.')
        scopes = {x['data'].get('input_scope','precedente') for x in records
                  if site_key(x['data']['nome']) == site_key(site) and is_weekly_assessment(x['data'])}
        if len(scopes) > 1:
            st.warning('Perimetri delle registrazioni non uniformi: interpretare il trend con cautela.')
        # NON reindicizzare e NON interpolare i valori. I punti esistono solo per le settimane salvate.
        st.altair_chart(alt.Chart(selected).mark_line(point=True).encode(
            x=alt.X('Data settimana:T', title='Settimana (lunedì)', axis=alt.Axis(format='%d/%m/%Y')),
            y=alt.Y('Indice:Q', title='Risk Index', scale=alt.Scale(domain=[0,100])),
            tooltip=[alt.Tooltip('Data settimana:T', title='Settimana', format='%d/%m/%Y'),
                     alt.Tooltip('Indice:Q', format='.2f'), 'Livello:N']), width='stretch')
        st.dataframe(selected.drop(columns=['site','id','Data settimana']), hide_index=True, width='stretch')
    else:
        st.info('Nessuna valutazione settimanale registrata per questo cantiere.')
    legacy = site_frame[site_frame['Frequenza'] == 'Mensile storico']
    if not legacy.empty:
        with st.expander(f'Archivio mensile precedente (modello {LEGACY_MODEL_VERSION}) · {len(legacy)} periodi'):
            st.caption('Le valutazioni mensili precedenti sono conservate e consultabili, ma non unite al trend settimanale: '
                       'periodi e formule diversi non sono confrontabili direttamente.')
            st.dataframe(legacy.drop(columns=['site','id']), hide_index=True, width='stretch')
    revisions = [x for x in records if site_key(x['data']['nome']) == site_key(site)]
    revisions.sort(key=lambda x: (period_key(x['data']), x['saved_at']), reverse=True)
    lookup = {x['id']: x for x in revisions}
    ident = st.selectbox('Apri una valutazione o revisione', list(lookup), format_func=lambda i:
        f"{period_label(lookup[i]['data'])} · {lookup[i]['result']['score']:.1f} · "
        f"modello {lookup[i]['model_version']} · salvata {lookup[i]['saved_at'][:19]} · {i[:8]}")
    a, b = st.columns(2)
    if a.button('Apri nei risultati'):
        st.session_state.current_report = deepcopy(lookup[ident])
        st.success('Valutazione caricata: apri “Risultati e azioni”.')
    if b.button('Usa come base per una nuova settimana'):
        st.session_state.draft = next_weekly_draft(lookup[ident]['data'])
        st.session_state.draft_revision = st.session_state.get('draft_revision', 0) + 1
        st.success('Modulo preparato per la settimana in corso. Le segnalazioni e i dati del periodo precedente non sono stati copiati automaticamente.')


def render_comparison(records):
    st.title('\U0001f4ca Confronto cantieri')
    frame = history_frame(records)
    if frame.empty:
        st.info('Servono valutazioni salvate nello storico.')
        return
    completed = (frame['Frequenza'] == 'Settimanale') & (frame['Versione'] == MODEL_VERSION)
    week_end = pd.to_datetime(frame.loc[completed, 'Periodo'], errors='coerce') + pd.Timedelta(days=6)
    weeks = sorted(frame.loc[completed].loc[week_end.dt.date < date.today(), 'Periodo'].unique(), reverse=True)
    if not weeks:
        st.info('La graduatoria consolidata richiede almeno una settimana conclusa e salvata nel nuovo modello.')
        return
    week = st.selectbox('Settimana conclusa · stesso periodo per tutti', weeks,
                        format_func=lambda value: period_label({'week_start': value}))
    types = st.multiselect('Tecnologie', ['Wind','Solar','BESS'], default=['Wind','Solar','BESS'])
    lookup = {x['id']: x for x in records}
    frame['Target'] = frame['id'].map(lambda i: lookup[i]['data']['inspection_target'])
    targets = sorted(frame.loc[(frame['Periodo'] == week) & completed,'Target'].unique())
    target = st.selectbox('Politica comune · target ispezioni settimanali', targets)
    selected = frame[(frame['Periodo'] == week) & completed & frame['Tecnologia'].isin(types)
                     & (frame['Target'] == target)].sort_values('Indice', ascending=False)
    st.caption('Indice base; stessa settimana ISO conclusa, stessa versione modello e stesso target ispettivo. '
               'Le settimane senza valutazione non entrano nel confronto e non ricevono alcun valore implicito.')
    if selected.empty:
        st.info('Nessuna valutazione con questi filtri.')
        return
    scopes = {lookup[i]['data'].get('input_scope','precedente') for i in selected['id']}
    if len(scopes) > 1:
        st.warning('Perimetri delle registrazioni non uniformi: verificare l’omogeneità degli input.')
    st.altair_chart(alt.Chart(selected).mark_bar(cornerRadiusEnd=4, color='#225873').encode(
        x=alt.X('Indice:Q', scale=alt.Scale(domain=[0,100])), y=alt.Y('Cantiere:N', sort='-x'),
        tooltip=['Cantiere','Tecnologia','Indice','Livello']), width='stretch')
    st.dataframe(selected[['Cantiere','Tecnologia','Periodo','Indice','Livello','Target']], hide_index=True, width='stretch')
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



# =========================================================
# PORTALE IMPRESA E SINOTTICO SETTIMANALE
# Struttura ricreata dal format; nessuna riga del PDF originale è inclusa.
# =========================================================
WEEKLY_SCHEMA = 2
DAYS_IT = ['Lun','Mar','Mer','Gio','Ven','Sab','Dom']
YES_NO = ['Da verificare','No','Sì']


def default_weekly():
    start = date.today() - timedelta(days=date.today().weekday())
    return dict(site='',technology='Wind',prime='',executor='',supervisor='',prime_lead='',executor_lead='',
                start=start.isoformat(),end=(start+timedelta(days=6)).isoformat(),activities=[],daily_workforce={})


def blank_weekly_activity():
    return dict(id=uuid.uuid4().hex,area='',description='',types=[],roadblock='Da verificare',night='Da verificare',
                confined='Da verificare',equipment='',hotwork='Da verificare',interference='',
                prime='',executor='',preposto='',contact='',pos_ref='',notes='',days={})


def plan_dates(plan, slots=False):
    start = date.fromisoformat(plan['start'])
    count = 8 if slots else (date.fromisoformat(plan['end'])-start).days+1
    return [(start+timedelta(days=i)).isoformat() for i in range(count)]


def weekly_questions(plan, require_pos=True):
    """Domande esplicite per ogni campo mancante; nessuna integrazione inventata."""
    questions=[]
    if not isinstance(plan,dict):
        return ['Sinottico non valido.']
    for key,label in [('site','Qual è il nome del cantiere?'),('prime','Qual è l’impresa affidataria?'),
       ('executor','Qual è l’impresa esecutrice?'),('supervisor','Chi è il supervisore Enel?'),
       ('prime_lead','Chi è il capo cantiere dell’affidataria?'),('executor_lead','Chi è il capo cantiere dell’esecutrice?')]:
        if not isinstance(plan.get(key),str) or not plan[key].strip() or len(plan[key])>160:
            questions.append(label)
    if plan.get('technology') not in ['Wind','Solar','BESS']:
        questions.append('Qual è la tecnologia del cantiere?')
    try:
        start,end=date.fromisoformat(plan['start']),date.fromisoformat(plan['end'])
        if start.weekday()!=0 or not 6 <= (end-start).days <=7:
            questions.append('Impostare un lunedì e un orizzonte di 7 o 8 giorni consecutivi.')
        dates=plan_dates(plan)
    except (KeyError,ValueError,TypeError):
        return questions+['Quali sono le date di inizio e fine della settimana?']
    activities=plan.get('activities')
    if not isinstance(activities,list) or not 1 <=len(activities)<=60:
        return questions+['Aggiungi da 1 a 60 attività pianificate.']
    seen=set()
    for i,row in enumerate(activities,1):
        if not isinstance(row,dict):
            questions.append(f'Attività {i}: formato non valido.');continue
        if not isinstance(row.get('id'),str) or not re.fullmatch(r'[a-f0-9]{32}',row['id']) or row['id'] in seen:
            questions.append(f'Attività {i}: identificativo non valido o duplicato.')
        seen.add(row.get('id',''))
        labels={'area':'In quale area si svolge?', 'description':'Qual è la descrizione della lavorazione?',
                'equipment':'Quali mezzi sono previsti? Scrivi Non previsti se assenti.',
                'interference':'Quali ulteriori rischi/interferenze? Scrivi Non rilevati se assenti.',
                'prime':'Qual è l’affidataria della lavorazione?', 'executor':'Qual è l’esecutrice della lavorazione?',
                'preposto':'Chi è il preposto della lavorazione?', 'contact':'Qual è il contatto reperibile del preposto?'}
        if require_pos:
            value=row.get('pos_ref','')
            if not isinstance(value,str) or not value.strip() or len(value)>600:
                questions.append(f'Attività {i}: qual è il riferimento POS? Inserisci un testo libero non vuoto (massimo 600 caratteri).')
        elif 'pos_ref' in row and (not isinstance(row['pos_ref'],str) or len(row['pos_ref'])>600):
            questions.append(f'Attività {i}: riferimento POS non valido.')
        for field,label in labels.items():
            if not isinstance(row.get(field),str) or not row[field].strip() or len(row[field])>(1200 if field in ['description','interference','equipment'] else 160):
                questions.append(f'Attività {i}: {label}')
        if not isinstance(row.get('notes',''),str) or len(row.get('notes',''))>1200:
            questions.append(f'Attività {i}: note troppo lunghe.')
        for field,label in [('night','È previsto orario notturno?'),('confined','È previsto accesso a spazio confinato?'),('hotwork','Sono previsti lavori a caldo?')]:
            if row.get(field) not in ['Sì','No']:
                questions.append(f'Attività {i}: {label}')
        if not isinstance(row.get('roadblock'),str) or not row['roadblock'].strip() or row['roadblock']=='Da verificare' or len(row['roadblock'])>300:
            questions.append(f'Attività {i}: quali interruzioni di aree/strade? Scrivi Non previste se assenti.')
        types=row.get('types')
        if not isinstance(types,list) or not types or any(x not in ACTIVITY for x in types) or len(types)!=len(set(types)):
            questions.append(f'Attività {i}: quali tipologie di lavoro servono per la checklist?')
        days=row.get('days')
        if not isinstance(days,dict) or not days:
            questions.append(f'Attività {i}: in quali giorni si lavora e con quante persone?')
        else:
            for day,count in days.items():
                if day not in dates or isinstance(count,bool) or not isinstance(count,int) or not 1<=count<=10000:
                    questions.append(f'Attività {i}: giorno o numero persone non valido.')
    totals=plan.get('daily_workforce',{})
    if not isinstance(totals,dict):
        questions.append('Completare la forza lavoro giornaliera verificata.');return questions
    for day in dates:
        assignments=[row.get('days',{}).get(day,0) for row in activities if isinstance(row,dict) and isinstance(row.get('days'),dict)]
        if any(not isinstance(n,int) or isinstance(n,bool) or n<0 for n in assignments):
            continue
        lo,hi=max(assignments or [0]),sum(assignments)
        n=totals.get(day)
        if isinstance(n,bool) or not isinstance(n,int) or not lo<=n<=hi:
            questions.append(f'{day}: quante persone uniche sono previste? Valore coerente fra {lo} e {hi}, senza duplicare persone su più attività.')
    return list(dict.fromkeys(questions))


def plan_identity(plan):
    return hashlib.sha256((site_key(plan['site'])+'|'+site_key(plan['executor'])+'|'+plan['start']).encode()).hexdigest()


def latest_weekly(records):
    latest={}
    for record in sorted(records,key=lambda r:(r['revision'],r['published_at'],r['id'])):
        latest[record['plan_key']]=record
    return list(latest.values())


def publish_weekly(store, plan, base_id=None):
    errors=weekly_questions(plan)
    if errors:
        raise ValueError('Completare il sinottico: '+'; '.join(errors[:4]))
    ident=plan_identity(plan)
    latest=next((x for x in latest_weekly(store.all_weekly()) if x['plan_key']==ident),None)
    if (latest and latest['id']!=base_id) or (not latest and base_id):
        raise ValueError('Esiste una revisione diversa. Carica l’ultima dal calendario prima di pubblicare una modifica.')
    if latest and latest['plan']==plan:
        return latest
    record=dict(id=str(uuid.uuid4()),schema=WEEKLY_SCHEMA,plan_key=ident,
                revision=latest['revision']+1 if latest else 0,
                published_at=datetime.now(timezone.utc).isoformat(),plan=deepcopy(plan))
    store.save_weekly(record)
    return record


def weekly_backup(records):
    return json.dumps({'weekly_schema':WEEKLY_SCHEMA,'plans':records},ensure_ascii=False,indent=2,allow_nan=False).encode()


def parse_weekly_backup(raw):
    if len(raw)>20*1024*1024:
        raise ValueError('Backup troppo grande.')
    payload=json.loads(raw)
    if not isinstance(payload,dict) or payload.get('weekly_schema') not in [1,WEEKLY_SCHEMA] or not isinstance(payload.get('plans'),list) or len(payload['plans'])>10000:
        raise ValueError('Formato archivio settimanale non valido.')
    results=[];seen=set()
    for row in payload['plans']:
        if not isinstance(row,dict) or row.get('schema') not in [1,WEEKLY_SCHEMA]:
            raise ValueError('Revisione non valida.')
        uuid.UUID(row['id']);datetime.fromisoformat(row['published_at'])
        if weekly_questions(row['plan'],require_pos=row['schema']>=2) or row.get('plan_key')!=plan_identity(row['plan']):
            raise ValueError('Dati pianificazione incompleti o non coerenti.')
        if not isinstance(row.get('revision'),int) or isinstance(row['revision'],bool) or row['revision']<0 or (row['plan_key'],row['revision']) in seen:
            raise ValueError('Numero revisione non valido o duplicato.')
        seen.add((row['plan_key'],row['revision']));results.append(row)
    return results


def checklist_for_plan(plan):
    types={t for row in plan['activities'] for t in row.get('types',[]) if t in ACTIVITY}
    if any(x.get('hotwork')=='Sì' for x in plan['activities']):types.add('hot work')
    if any(x.get('confined')=='Sì' for x in plan['activities']):types.add('spazi confinati')
    sections=['Controlli generali pre-attività']+[ACTIVITY_TO_CHECKLIST[x] for x in ACTIVITY if x in types and x in ACTIVITY_TO_CHECKLIST]
    return list(dict.fromkeys(sections))


def parse_planning_text(text, start):
    """Import guidato senza LLM: righe delimitate o campi etichettati; valori mancanti restano mancanti.
    Colonne esempio: area;attivita;tipi;giorni;persone;preposto;contatto;mezzi;notturno;confinato;caldo;interferenze;interruzioni;affidataria;esecutrice;riferimento_pos;note
    Giorni: date ISO separate da virgole. Tipi: chiavi ACTIVITY separate da virgole.
    """
    if len(text)>100000:raise ValueError('Testo troppo lungo (massimo 100.000 caratteri).')
    aliases={'area':'area','area di lavoro':'area','attivita':'description','attività':'description','descrizione':'description',
             'tipi':'types','giorni':'days_raw','persone':'people_raw','preposto':'preposto','preposto impresa':'preposto',
             'contatto':'contact','mezzi':'equipment','mezzi d’opera':'equipment','notturno':'night','confinato':'confined',
             'caldo':'hotwork','interferenze':'interference','interruzioni':'roadblock','affidataria':'prime','esecutrice':'executor','note':'notes','riferimento pos':'pos_ref','riferimento_pos':'pos_ref','pos':'pos_ref','pos_ref':'pos_ref'}
    lines=text.strip().splitlines();rawrows=[]
    if not lines:return [],['Testo vuoto.']
    delimiter=max([';','|','\t'],key=lambda ch:lines[0].count(ch))
    headers=next(csv.reader([lines[0]],delimiter=delimiter))
    if sum(site_key(h) in aliases for h in headers)>=2:
        for row in csv.DictReader(StringIO(text.strip()),delimiter=delimiter):
            rawrows.append({aliases[site_key(k)]:v.strip() for k,v in row.items() if k and site_key(k) in aliases and isinstance(v,str)})
    else:
        for block in re.split(r'\n\s*\n',text):
            row={}
            for line in block.splitlines():
                if ':' in line:
                    key,value=line.split(':',1)
                    if site_key(key) in aliases:row[aliases[site_key(key)]]=value.strip()
            if row:rawrows.append(row)
    rows=[];notes=[]
    for source in rawrows[:60]:
        row=blank_weekly_activity()
        for key,value in source.items():
            if key in row and key not in ['days','types']:
                row[key]=value
        row['types']=[x.strip().lower() for x in source.get('types','').split(',') if x.strip().lower() in ACTIVITY]
        for key in ['night','confined','hotwork']:
            value=site_key(row[key])
            row[key]='Sì' if value in ['si','sì','yes'] else 'No' if value=='no' else 'Da verificare'
        count=source.get('people_raw','').strip()
        for day in source.get('days_raw','').split(','):
            day=day.strip();daily_count=count
            if '=' in day:day,daily_count=day.split('=',1)
            try:
                date.fromisoformat(day)
                if daily_count.isdigit() and int(daily_count)>0:row['days'][day]=int(daily_count)
            except ValueError:pass
        rows.append(row)
    if not rows:notes.append('La disposizione delle colonne non è riconosciuta: nessuna attività importata automaticamente. Usa il testo estratto come riferimento e aggiungi le schede; il tool chiederà tutti i campi mancanti.')
    notes.append('Riconoscimento assistito: verificare tutte le schede. Date, nomi e risposte non riconosciuti non vengono inventati.')
    return rows,notes


def extract_pdf_grid(pdf, start):
    """Riconosce la griglia del format negli upload successivi, non usa il PDF di riferimento."""
    fields=['area','attivita','tipi','giorni','persone','preposto','contatto','mezzi','notturno','confinato','caldo','interferenze','interruzioni','affidataria','esecutrice','riferimento_pos','note']
    found=[];startdate=date.fromisoformat(start)
    for page in pdf.pages:
        for table in page.extract_tables():
            header_index=next((i for i,row in enumerate(table) if any('area di lavoro' in site_key(x or '') for x in row) and any('attivit' in site_key(x or '') for x in row)),None)
            if header_index is None:continue
            header_indices=[j for j,x in enumerate(table[header_index]) if x and x.strip()]
            if len(header_indices) not in [20,21]:continue
            has_pos=len(header_indices)==21
            header=[table[header_index][j] for j in header_indices]
            dates={}
            for j in range(10,18):
                match=re.search(r'(\d{1,2})/(\d{1,2})',re.sub(r'\s+','',header[j] or ''))
                if match:
                    for year in [startdate.year-1,startdate.year,startdate.year+1]:
                        try:
                            day=date(year,int(match[2]),int(match[1]))
                            if 0<=(day-startdate).days<=7:dates[j]=day.isoformat()
                        except ValueError:pass
            for cells in table[header_index+1:]:
                if len(cells)<=max(header_indices):continue
                values=[(cells[j] or '').strip() for j in header_indices]
                if not values[1] or values[0].upper().startswith(('TOT','GUIDA')):continue
                preposto_parts=values[18].splitlines()
                days=[]
                for j,day in dates.items():
                    count=values[j].strip()
                    if count.isdigit() and int(count)>0:days.append(day+'='+count)
                found.append([values[0],values[1],'',','.join(days),'',preposto_parts[0] if preposto_parts else '',
                              ' '.join(preposto_parts[1:]),values[5],values[3],values[4],values[6],values[7],values[2],values[8],values[9],(' '.join(values[19].split()) if ' '.join(values[19].split())!='Da integrare (piano precedente)' else '') if has_pos else '',values[20] if has_pos else values[19]])
    if not found:return ''
    output=StringIO();writer=csv.writer(output,delimiter=';');writer.writerow(fields);writer.writerows(found)
    return output.getvalue()


def extract_planning_upload(data, filename, start=None):
    if len(data)>15*1024*1024:raise ValueError('Massimo 15 MB per documento.')
    if filename.lower().endswith('.pdf'):
        import pdfplumber
        with pdfplumber.open(BytesIO(data)) as pdf:
            if len(pdf.pages)>5:raise ValueError('Massimo 5 pagine per importazione.')
            grid=extract_pdf_grid(pdf,start) if start else ''
            if grid:return grid,'Griglia PDF riconosciuta: verifica date, nomi e risposte; completa le tipologie per la checklist e i campi mancanti.'
            return '\n\n'.join(page.extract_text() or '' for page in pdf.pages), 'Testo PDF estratto; se la griglia non è riconosciuta occorre completare le schede. Per PDF scansionati carica un’immagine per usare OCR.'
    from PIL import Image
    import pytesseract
    image=Image.open(BytesIO(data))
    if image.width*image.height>20000000:raise ValueError('Immagine troppo grande: massimo 20 milioni di pixel.')
    if image.width<2200:image=image.resize((int(image.width*2),int(image.height*2)))
    available=pytesseract.get_languages(config='')
    language='+'.join(x for x in ['ita','eng'] if x in available)
    if not language:raise ValueError('Lingue OCR italiano/inglese non installate.')
    text=pytesseract.image_to_string(image.convert('RGB'),lang=language,config='--psm 6',timeout=40)
    return text,'OCR locale: date, nomi e colonne devono essere confermati. Le immagini sorgente non sono archiviate.'


def plan_table_rows(plan):
    dates=plan_dates(plan,slots=True)
    headers=['Area di lavoro','Attività','Interruz. area/strada','Orario notturno','Spazio confinato',"Mezzi d’opera",'Lavori a caldo','Ulteriori rischi / interferenze','Affidataria','Esecutrice']
    headers += [DAYS_IT[date.fromisoformat(day).weekday()]+'\n'+date.fromisoformat(day).strftime('%d/%m') for day in dates]
    headers += ['Preposto impresa / contatto','Riferimento POS','Note / misure di coordinamento']
    rows=[headers]
    for row in plan['activities']:
        rows.append([row['area'],row['description'],row['roadblock'],row['night'],row['confined'],row['equipment'],row['hotwork'],
                     row['interference'],row['prime'],row['executor']]+[str(row['days'].get(day,'')) for day in dates]+
                    [row['preposto']+'\n'+row['contact'],row.get('pos_ref','').strip() or 'Da integrare (piano precedente)',row['notes']])
    rows.append(['TOT FORZA LAVORO']+['']*9+[str(plan['daily_workforce'].get(day,'')) for day in dates]+['','',''])
    return rows


def sinottico_pdf(plan, revision=0, draft=True):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    buffer=BytesIO();width,height=landscape(A3);usable=width-72
    doc=SimpleDocTemplate(buffer,pagesize=(width,height),leftMargin=30,rightMargin=30,topMargin=28,bottomMargin=30)
    normal=ParagraphStyle('cell',fontName='Helvetica',fontSize=7.2,leading=9,spaceAfter=0)
    heading=ParagraphStyle('head',parent=normal,fontName='Helvetica-Bold',alignment=TA_CENTER,fontSize=7)
    def p(value,style=normal):return Paragraph(html.escape(str(value)).replace('\n','<br/>'),style)
    title=Table([[p('SINOTTICO DI COORDINAMENTO ATTIVITÀ',heading),p('DAL '+plan['start'],heading),p('AL '+plan['end'],heading),p(f'Rev. {revision:02d}'+(' · BOZZA' if draft else ''),heading)]],colWidths=[usable*.45,usable*.2,usable*.2,usable*.15],hAlign='LEFT')
    title.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#ffbf00')),('GRID',(0,0),(-1,-1),.6,colors.grey),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    metadata=Table([[p('Nome impresa affidataria: '+plan['prime']),p('Nome impresa esecutrice: '+plan['executor']),p('Supervisore Enel: '+plan['supervisor'])],
                    [p('Firma C.C. affidataria: __________________\nReferente: '+plan['prime_lead']),p('Firma C.C. esecutrice: __________________\nReferente: '+plan['executor_lead']),p('Firma supervisore Enel: __________________')]],colWidths=[usable/3]*3,hAlign='LEFT')
    metadata.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.5,colors.grey),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),9),('BOTTOMPADDING',(0,0),(-1,-1),9)]))
    ratios=[6,16,4,3,4,8,4,7,5,5]+[2]*8+[7,7,8]
    rows=plan_table_rows(plan);table=Table([[p(v,heading if i==0 else normal) for v in row] for i,row in enumerate(rows)],colWidths=[usable*x/100 for x in ratios],repeatRows=1,hAlign='LEFT')
    styles=[('LEFTPADDING',(0,0),(-1,-1),1.5),('RIGHTPADDING',(0,0),(-1,-1),1.5),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#999999')),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#eeeeee')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('SPAN',(0,-1),(9,-1)),('BACKGROUND',(10,-1),(17,-1),colors.HexColor('#ffbf00'))]
    for j,day in enumerate(plan_dates(plan,slots=True),10):
        if day>plan['end']:styles.append(('BACKGROUND',(j,0),(j,-2),colors.HexColor('#dddddd')))
        elif date.fromisoformat(day).weekday()==6:styles.append(('BACKGROUND',(j,0),(j,-2),colors.HexColor('#ffe2e2')))
        for i,row in enumerate(plan['activities'],1):
            if day in row['days']:styles.append(('BACKGROUND',(j,i),(j,i),colors.HexColor('#bbd7ed')))
    table.setStyle(TableStyle(styles))
    guide=('GUIDA ALLA COMPILAZIONE DEL SINOTTICO\n'
      'Indicare area, descrizione della lavorazione, interruzioni, orario notturno, spazio confinato, mezzi, lavori a caldo, ulteriori rischi e interferenze, affidataria, esecutrice, preposto reperibile e riferimento POS per ogni attività. '
      'Le celle giornaliere riportano le persone assegnate alla lavorazione; una cella vuota significa attività non prevista. '
      'Il totale giornaliero rappresenta le persone uniche dichiarate e verificate dall’impresa, evitando duplicazioni fra attività. '
      'La colonna oltre il periodo è fuori pianificazione. Le firme restano da apporre: la pubblicazione nel tool non equivale ad approvazione del piano o autorizzazione al lavoro.')
    story=[p('Cantiere: '+plan['site']+' · '+plan['technology']),Spacer(1,8),title,metadata,table,Spacer(1,10),p(guide)]
    def footer(canvas,document):
        canvas.setFont('Helvetica',7);canvas.drawString(30,15,'HSE Risk Platform · '+('BOZZA' if draft else 'PIANIFICAZIONE INSERITA · firme da acquisire'))
        canvas.drawRightString(width-30,15,str(document.page))
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    return buffer.getvalue()


def sinottico_word(plan, revision=0, draft=True):
    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc=Document();sec=doc.sections[0];sec.orientation=WD_ORIENT.LANDSCAPE
    sec.page_width=DocxInches(16.54);sec.page_height=DocxInches(11.69)
    sec.left_margin=sec.right_margin=DocxInches(.4)
    doc.styles['Normal'].font.name='Arial';doc.styles['Normal'].font.size=Pt(7)
    doc.add_paragraph(f"Cantiere: {plan['site']} | {plan['technology']}")
    band=doc.add_table(rows=1,cols=4);band.style='Table Grid'
    for j,value in enumerate(['SINOTTICO DI COORDINAMENTO ATTIVITÀ','DAL '+plan['start'],'AL '+plan['end'],f'Rev. {revision:02d}'+(' · BOZZA' if draft else '')]):
        cell=band.cell(0,j);cell.text=value
        shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'FFBF00');cell._tc.get_or_add_tcPr().append(shade)
        for run in cell.paragraphs[0].runs:run.bold=True
    meta=doc.add_table(rows=2,cols=3);meta.style='Table Grid'
    for i,text in enumerate(['Nome impresa affidataria: '+plan['prime'],'Nome impresa esecutrice: '+plan['executor'],'Supervisore Enel: '+plan['supervisor']]):meta.cell(0,i).text=text
    for i,text in enumerate(['Firma C.C. affidataria: __________\nReferente: '+plan['prime_lead'],'Firma C.C. esecutrice: __________\nReferente: '+plan['executor_lead'],'Firma supervisore Enel: __________']):meta.cell(1,i).text=text
    rows=plan_table_rows(plan);table=doc.add_table(rows=0,cols=21);table.style='Table Grid';table.autofit=False
    ratios=[6,16,4,3,4,8,4,7,5,5]+[2]*8+[7,7,8]
    for j,ratio in enumerate(ratios):table.columns[j].width=DocxInches(15.74*ratio/100)
    for i,values in enumerate(rows):
        cells=table.add_row().cells
        for j,value in enumerate(values):
            cells[j].text=str(value);cells[j].width=DocxInches(15.74*ratios[j]/100)
            if i in [0,len(rows)-1]:
                shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'FFBF00' if i==len(rows)-1 else 'EEEEEE');cells[j]._tc.get_or_add_tcPr().append(shade)
        if i==0:
            repeat=OxmlElement('w:tblHeader');table.rows[0]._tr.get_or_add_trPr().append(repeat)
    total_cell=table.rows[-1].cells[0].merge(table.rows[-1].cells[9])
    total_cell.text='TOT FORZA LAVORO'
    for row in table.rows:
        for cell in row.cells:
            margins=OxmlElement('w:tcMar')
            for side in ['top','left','bottom','right']:
                item=OxmlElement('w:'+side);item.set(qn('w:w'),'40');item.set(qn('w:type'),'dxa');margins.append(item)
            cell._tc.get_or_add_tcPr().append(margins)
    doc.add_paragraph('GUIDA ALLA COMPILAZIONE DEL SINOTTICO')
    doc.add_paragraph('Riferimento POS: testo libero obbligatorio per ogni attività; indicare il documento e le parti pertinenti. Le celle giorno indicano persone assegnate. I totali sono persone uniche dichiarate e verificate. Firme da acquisire: inserimento nel tool non equivale ad approvazione o autorizzazione al lavoro.')
    out=BytesIO();doc.save(out);return out.getvalue()


def calendar_events(records, site=None, company=None):
    events=[]
    for record in latest_weekly(records):
        plan=record['plan']
        if site and site_key(plan['site'])!=site_key(site):continue
        if company and site_key(plan['executor'])!=site_key(company):continue
        for row in plan['activities']:
            for day,count in row['days'].items():
                events.append(dict(day=day,site=plan['site'],company=plan['executor'],description=row['description'],
                                   preposto=row['preposto'],area=row['area'],people=count,task_id=row['id'],record_id=record['id']))
    return sorted(events,key=lambda x:(x['day'],x['site'],x['company'],x['description']))


def calendar_ics(events):
    def escape(value):return str(value).replace('\\','\\\\').replace('\n','\\n').replace(';','\\;').replace(',','\\,').replace('\r','')
    lines=['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//HSE Risk Platform//Calendario//IT','CALSCALE:GREGORIAN']
    for event in events:
        day=date.fromisoformat(event['day'])
        lines+=['BEGIN:VEVENT','UID:'+event['record_id']+'-'+event['task_id']+'-'+day.strftime('%Y%m%d')+'@hse-risk',
                'DTSTAMP:'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ'),'DTSTART;VALUE=DATE:'+day.strftime('%Y%m%d'),
                'DTEND;VALUE=DATE:'+(day+timedelta(days=1)).strftime('%Y%m%d'),'SUMMARY:'+escape(event['description']),
                'DESCRIPTION:'+escape('Preposto: '+event['preposto']+'\nImpresa: '+event['company']),
                'LOCATION:'+escape(event['site']+' / '+event['area']),'END:VEVENT']
    lines.append('END:VCALENDAR')
    folded=[]
    for line in lines:
        current=''
        for char in line:
            if len((current+char).encode('utf-8'))>73:
                folded.append(current);current=' '+char
            else:current+=char
        folded.append(current)
    return ('\r\n'.join(folded)+'\r\n').encode('utf-8')


def add_weekly_activity():
    st.session_state.weekly_draft['activities'].append(blank_weekly_activity())
    st.session_state.pop('weekly_ready',None)


def remove_weekly_activity(uid):
    st.session_state.weekly_draft['activities']=[x for x in st.session_state.weekly_draft['activities'] if x['id']!=uid]
    st.session_state.pop('weekly_ready',None)


def reset_weekly():
    for key in list(st.session_state):
        if key!='weekly_records' and key.startswith(('weekly_', 'wk_','import_')):del st.session_state[key]


def load_weekly_record(record):
    reset_weekly()
    st.session_state.weekly_draft=deepcopy(record['plan'])
    st.session_state.weekly_base=record['id']
    st.session_state.weekly_base_key=record['plan_key']
    st.session_state.weekly_revision=record['revision']
    st.session_state.weekly_ui=uuid.uuid4().hex


def render_company_planning(store):
    if 'weekly_draft' not in st.session_state:st.session_state.weekly_draft=default_weekly()
    if 'weekly_ui' not in st.session_state:st.session_state.weekly_ui=uuid.uuid4().hex
    plan=st.session_state.weekly_draft;ui=st.session_state.weekly_ui
    st.title('\U0001f3d7\ufe0f Pianificazione settimanale impresa')
    st.caption('Aggiungi le attività, ottieni le verifiche preliminari e completa il sinottico prima di inserirlo nel calendario comune.')
    if store.mode=='session':st.warning('Archivio di sessione: il calendario è comune ai due percorsi in questa sessione. Per colleghi su browser diversi serve PostgreSQL condiviso oppure un server con SQLite persistente. Esporta il backup prima di chiudere.')
    with st.container(border=True):
        st.subheader('\U0001f4cd 1 · Cantiere e settimana')
        a,b=st.columns([2,1])
        plan['site']=a.text_input('Cantiere',plan['site'],max_chars=160,key='wk_site_'+ui)
        plan['technology']=b.selectbox('Tecnologia cantiere',['Wind','Solar','BESS'],index=['Wind','Solar','BESS'].index(plan['technology']),key='wk_tech_'+ui)
        a,b=st.columns(2)
        chosen=a.date_input('Inizio settimana · lunedì',date.fromisoformat(plan['start']),key='wk_start_'+ui)
        horizon=b.radio('Orizzonte pianificazione',[7,8],index=0 if (date.fromisoformat(plan['end'])-date.fromisoformat(plan['start'])).days==6 else 1,horizontal=True,key='wk_horizon_'+ui)
        plan['start']=chosen.isoformat();plan['end']=(chosen+timedelta(days=horizon-1)).isoformat()
        if chosen.weekday()!=0:st.warning('Scegli un lunedì per allineare sinottico e calendario settimanale.')
        a,b=st.columns(2)
        plan['prime']=a.text_input('Impresa affidataria',plan['prime'],max_chars=160,key='wk_prime_'+ui)
        plan['executor']=b.text_input('Impresa esecutrice',plan['executor'],max_chars=160,key='wk_executor_'+ui)
        a,b,c=st.columns(3)
        plan['prime_lead']=a.text_input('C.C. affidataria · nominativo',plan['prime_lead'],max_chars=160,key='wk_primelead_'+ui)
        plan['executor_lead']=b.text_input('C.C. esecutrice · nominativo',plan['executor_lead'],max_chars=160,key='wk_executorlead_'+ui)
        plan['supervisor']=c.text_input('Supervisore Enel · nominativo',plan['supervisor'],max_chars=160,key='wk_supervisor_'+ui)
    with st.expander('\U0001f4f7 2 · Importa un sinottico da immagine, PDF o testo'):
        st.caption('Riconoscimento locale e gratuito, senza servizi AI esterni. I risultati sono proposte da verificare; le informazioni illeggibili restano vuote e generano domande.')
        uploaded=st.file_uploader('Documento da leggere · PNG JPG WEBP o PDF',type=['png','jpg','jpeg','webp','pdf'],key='import_file_'+ui)
        if uploaded is not None and st.button('Estrai il testo del documento',key='import_ocr_'+ui):
            try:
                with st.spinner('Lettura del documento…'):
                    extracted,warning=extract_planning_upload(uploaded.getvalue(),uploaded.name,plan['start'])
                st.session_state['import_text_'+ui]=extracted;st.info(warning)
            except Exception:
                st.warning('Lettura automatica non disponibile o non riuscita. Per le immagini occorrono Tesseract e le lingue installate; puoi incollare il testo oppure compilare le schede senza OCR.')
        st.caption('Testo strutturato: una riga per attività e intestazioni separate da punto e virgola. In alternativa blocchi con Area:, Attività:, Preposto:, ecc. I giorni devono essere date AAAA-MM-GG e persone un numero intero.')
        st.code('area;attivita;tipi;giorni;persone;preposto;contatto;mezzi;notturno;confinato;caldo;interferenze;interruzioni;affidataria;esecutrice;riferimento_pos;note',language='text')
        text=st.text_area('Testo da verificare e importare',height=200,key='import_text_'+ui,max_chars=100000)
        if st.button('Prepara le schede dal testo',key='import_parse_'+ui):
            try:
                rows,notes=parse_planning_text(text,plan['start']);st.session_state.weekly_import=rows
                for note in notes:st.info(note)
            except Exception:st.error('Testo non interpretabile: verifica separatori e intestazioni, oppure usa le schede manuali.')
        imported=st.session_state.get('weekly_import',[])
        if imported:
            st.dataframe(pd.DataFrame([{'Area':x['area'],'Attività':x['description'],'Preposto':x['preposto'],'Riferimento POS':x.get('pos_ref',''),'Giorni riconosciuti':', '.join(x['days'])} for x in imported]),hide_index=True,width='stretch')
            if st.button('Conferma e aggiungi queste schede alla bozza',key='import_accept_'+ui):
                if len(plan['activities'])+len(imported)>60:st.error('Massimo 60 attività per sinottico.')
                else:
                    plan['activities'].extend(deepcopy(imported));st.session_state.pop('weekly_import',None);st.session_state.pop('weekly_ready',None);st.rerun()
    st.subheader('\U0001f5c2\ufe0f 3 · Attività della settimana')
    st.caption('Una scheda per lavorazione e preposto. Se cambiano il preposto o le condizioni operative crea una scheda distinta. I pulsanti dei giorni indicano quando è prevista la lavorazione.')
    if not plan['activities']:st.info('Il sinottico è vuoto. Premi Aggiungi attività: nessuna riga del modello allegato è stata precaricata.')
    for index,row in enumerate(plan['activities'],1):
        prefix='wk_task_'+ui+'_'+row['id']
        with st.expander(f"\U0001f3d7\ufe0f Attività {index} · {row['area'] or 'area da indicare'} · {row['description'][:60] or 'da compilare'}", expanded=index==len(plan['activities'])):
            title,remove=st.columns([5,1]);title.markdown(f"### \U0001f3d7\ufe0f Attività {index} · {row['area'] or 'area da indicare'}")
            remove.button('Rimuovi',icon='\U0001f5d1\ufe0f',key=prefix+'_remove',on_click=remove_weekly_activity,args=(row['id'],))
            a,b=st.columns([1,3])
            row['area']=a.text_input('Area di lavoro',row['area'],max_chars=160,key=prefix+'_area')
            row['description']=b.text_area('Descrizione della lavorazione',row['description'],max_chars=1200,key=prefix+'_description')
            row['types']=st.pills('Tipologie per la checklist',list(ACTIVITY),default=row['types'],selection_mode='multi',key=prefix+'_types')
            a,b,c=st.columns(3)
            row['night']=a.radio('Lavoro notturno',YES_NO,index=YES_NO.index(row['night']),key=prefix+'_night')
            row['confined']=b.radio('Spazio confinato',YES_NO,index=YES_NO.index(row['confined']),key=prefix+'_confined')
            row['hotwork']=c.radio('Lavori a caldo',YES_NO,index=YES_NO.index(row['hotwork']),key=prefix+'_hotwork')
            a,b=st.columns(2)
            row['equipment']=a.text_area('Mezzi d’opera · oppure Non previsti',row['equipment'],max_chars=1200,key=prefix+'_equipment')
            row['interference']=b.text_area('Ulteriori rischi / interferenze · oppure Non rilevati',row['interference'],max_chars=1200,key=prefix+'_interference')
            row['roadblock']=st.text_input('Interruzioni area / strada · oppure Non previste',row['roadblock'] if row['roadblock']!='Da verificare' else '',max_chars=300,key=prefix+'_road')
            a,b=st.columns(2)
            row['prime']=a.text_input('Affidataria della lavorazione',row['prime'] or plan['prime'],max_chars=160,key=prefix+'_prime')
            row['executor']=b.text_input('Esecutrice della lavorazione',row['executor'] or plan['executor'],max_chars=160,key=prefix+'_executor')
            a,b=st.columns(2)
            row['preposto']=a.text_input('Preposto incaricato',row['preposto'],max_chars=160,key=prefix+'_preposto')
            row['contact']=b.text_input('Contatto reperibile del preposto',row['contact'],max_chars=160,key=prefix+'_contact')
            row['pos_ref']=st.text_area('Riferimento POS · obbligatorio',row.get('pos_ref',''),max_chars=600,key=prefix+'_pos_ref',
                help='Testo libero: indica il POS pertinente alla lavorazione, ad esempio titolo/impresa, revisione e paragrafo o scheda. Il tool non verifica il documento citato.')
            row['notes']=st.text_area('Note e misure di coordinamento',row['notes'],max_chars=1200,key=prefix+'_notes')
            st.markdown('**\U0001f4c5 Giorni e persone assegnate**')
            selected={}
            for offset,day in enumerate(plan_dates(plan)):
                a,b=st.columns([3,1]);dt=date.fromisoformat(day)
                enabled=a.checkbox(DAYS_IT[dt.weekday()]+' '+dt.strftime('%d/%m'),value=day in row['days'],key=prefix+'_day_'+day)
                if enabled:selected[day]=b.number_input('Persone',1,10000,int(row['days'].get(day,1)),key=prefix+'_people_'+day)
            row['days']=selected
    st.button('Aggiungi attività',icon='\u2795',width='stretch',on_click=add_weekly_activity,disabled=len(plan['activities'])>=60)
    sections=checklist_for_plan(plan)
    with st.container(border=True):
        st.subheader('\U0001f4d1 4 · Verifiche preliminari automatiche')
        st.caption('La checklist segue le tipologie selezionate e aggiunge i capitoli per lavori a caldo o spazi confinati quando dichiarati. Non richiede la pubblicazione del sinottico.')
        for name in sections:st.write('\u2713 '+name)
        if plan['activities'] and plan['site'].strip():
            document=generate_checklist_word({'nome':plan['site'],'fase':'Pianificazione settimanale '+plan['start']},sections)
            st.download_button('Scarica checklist preliminare Word',document.getvalue(),f"Checklist_{safe_filename(plan['site'])}_{plan['start']}.docx",icon='\U0001f4e5')
    st.subheader('\U0001f477 5 · Forza lavoro giornaliera')
    st.caption('Indica le persone uniche previste: la stessa persona impegnata su più attività non va contata due volte. Il totale deve essere coerente con le assegnazioni alle schede.')
    totals={}
    for day in plan_dates(plan):
        assignments=[x['days'].get(day,0) for x in plan['activities']];lo=max(assignments or [0]);hi=sum(assignments)
        value=plan['daily_workforce'].get(day,hi);value=min(hi,max(lo,value))
        totals[day]=st.number_input(f"{DAYS_IT[date.fromisoformat(day).weekday()]} {day} · persone uniche (assegnazioni {hi})",lo,hi,int(value),key='wk_total_'+ui+'_'+day+'_'+str(lo)+'_'+str(hi))
    plan['daily_workforce']=totals
    st.subheader('\U0001f9fe 6 · Sinottico')
    problems=weekly_questions(plan)
    if problems:
        st.warning(f'Servono {len(problems)} chiarimenti prima di completare il sinottico.')
        for question in problems:st.write('\u2753 '+question)
    else:st.success('Tutti i campi obbligatori sono compilati. Verifica la bozza prima dell’inserimento.')
    if st.button('Verifica e prepara il sinottico',icon='\U0001f50e',type='primary'):
        if problems:st.error('Completa i campi indicati nelle domande.')
        else:
            try:
                rev=st.session_state.get('weekly_revision',-1)+1
                st.session_state.weekly_ready=dict(fingerprint=hashlib.sha256(json.dumps(plan,sort_keys=True).encode()).hexdigest(),
                    pdf=sinottico_pdf(plan,rev,True),word=sinottico_word(plan,rev,True))
            except Exception:st.error('Generazione non riuscita. Riduci testi troppo lunghi e verifica i dati; nessuna pianificazione inserita.')
    ready=st.session_state.get('weekly_ready')
    fingerprint=hashlib.sha256(json.dumps(plan,sort_keys=True).encode()).hexdigest()
    if ready and ready['fingerprint']==fingerprint and not problems:
        a,b=st.columns(2)
        a.download_button('Bozza sinottico PDF A3',ready['pdf'],'Sinottico_BOZZA.pdf',icon='\U0001f4c4')
        b.download_button('Bozza sinottico Word',ready['word'],'Sinottico_BOZZA.docx',icon='\U0001f4c4')
        confirm=st.checkbox('Confermo dati, date, preposti e totali giornalieri; inserisco il sinottico nel calendario condiviso.',key='wk_confirm_'+fingerprint)
        if st.button('Conferma e inserisci',icon='\U0001f4be',type='primary',disabled=not confirm):
            try:
                base=st.session_state.get('weekly_base') if st.session_state.get('weekly_base_key')==plan_identity(plan) else None
                inserted=publish_weekly(store,deepcopy(plan),base)
                st.session_state.weekly_base=inserted['id'];st.session_state.weekly_base_key=inserted['plan_key'];st.session_state.weekly_revision=inserted['revision']
                st.session_state.weekly_inserted=inserted
                st.session_state['calendar_week']=date.fromisoformat(plan['start'])
                st.success('Sinottico inserito: visibile nel calendario HSE Enel e HSE impresa. Nessuna firma o approvazione è stata simulata.')
            except Exception as exc:
                if isinstance(exc,ValueError):st.error(str(exc))
                else:st.error('Inserimento non riuscito o revisione concorrente. Ricarica il calendario prima di riprovare. La bozza resta disponibile.')
    elif ready:st.info('La bozza è cambiata: premi di nuovo Verifica e prepara il sinottico prima di confermare.')
    st.caption('Non salvare scarta solo la bozza e i documenti temporanei. Le pianificazioni già inserite restano nel calendario e nell’archivio.')
    if st.button('Non salvare la pianificazione',icon='\U0001f9ea'):
        reset_weekly();st.rerun()


def render_weekly_calendar(store):
    st.title('\U0001f5d3\ufe0f Calendario condiviso delle attività')
    st.caption('Mostra soltanto i sinottici confermati, nell’ultima revisione inserita. Le bozze non compaiono. Descrizione attività e preposto sono sempre visibili.')
    if store.mode=='session':st.warning('Calendario di questa sessione: per sincronizzare altri browser configura un archivio comune. Puoi anche trasferire i dati con il backup settimanale.')
    st.button('Aggiorna calendario',icon='\U0001f504')
    records=store.all_weekly();latest=latest_weekly(records)
    if not latest:
        st.info('Nessuna pianificazione confermata. Il percorso HSE impresa consente di preparare il primo sinottico.')
        return
    a,b=st.columns(2)
    sites=sorted({x['plan']['site'] for x in latest});site=a.selectbox('Cantiere del calendario',sites)
    companies=sorted({x['plan']['executor'] for x in latest if site_key(x['plan']['site'])==site_key(site)})
    company=b.selectbox('Impresa del calendario',['Tutte']+companies)
    chosen=st.date_input('Settimana da visualizzare',date.today(),key='calendar_week')
    start=chosen-timedelta(days=chosen.weekday());days=[(start+timedelta(days=i)).isoformat() for i in range(7)]
    events=calendar_events(records,site,None if company=='Tutte' else company)
    weekevents=[e for e in events if e['day'] in days]
    parts=['<div class="week-grid">']
    for day in days:
        dt=date.fromisoformat(day);items=[e for e in weekevents if e['day']==day]
        parts.append('<section class="week-day"><h4>'+DAYS_IT[dt.weekday()]+' '+dt.strftime('%d/%m')+'</h4>')
        if not items:parts.append('<p class="muted">Nessuna attività inserita</p>')
        for event in items:
            parts.append('<article class="event"><strong>'+html.escape(event['description'])+'</strong><br><span>\U0001f477 '+html.escape(event['preposto'])+'</span><br><small>'+html.escape(event['company']+' · '+event['area'])+'</small></article>')
        parts.append('</section>')
    parts.append('</div>');st.markdown(''.join(parts),unsafe_allow_html=True)
    st.download_button('Esporta questa settimana nel calendario ICS',calendar_ics(weekevents),f'Attivita_{start.isoformat()}.ics','text/calendar',icon='\U0001f4c5')
    st.subheader('\U0001f4c4 Sinottici inseriti e revisioni')
    candidates=[r for r in latest if site_key(r['plan']['site'])==site_key(site) and (company=='Tutte' or site_key(r['plan']['executor'])==site_key(company)) and r['plan']['start']<=days[-1] and r['plan']['end']>=days[0]]
    for record in candidates:
        p=record['plan'];uid=record['id']
        with st.expander(f"{p['executor']} · {p['start']} / {p['end']} · Rev. {record['revision']:02d}"):
            if any(not str(x.get('pos_ref','')).strip() for x in p['activities']):
                st.warning('Pianificazione precedente senza riferimento POS: lo storico è conservato. Caricala per aggiornare e integra il riferimento POS di ogni attività prima della nuova conferma.')
            st.caption('Firme da acquisire. Inserimento nel calendario non equivale ad approvazione o autorizzazione al lavoro.')
            a,b=st.columns(2)
            a.download_button('Sinottico PDF A3',sinottico_pdf(p,record['revision'],False),f'Sinottico_{safe_filename(p["executor"])}_{p["start"]}.pdf',key='cal_pdf_'+uid)
            b.download_button('Sinottico Word',sinottico_word(p,record['revision'],False),f'Sinottico_{safe_filename(p["executor"])}_{p["start"]}.docx',key='cal_doc_'+uid)
            if st.session_state.get('portal_role')=='HSE impresa' and st.button('Carica per aggiornare questa pianificazione',key='cal_load_'+uid):
                load_weekly_record(record);st.success('Pianificazione caricata: apri Attività e sinottico per modificare e inserire una nuova revisione.')
    st.caption(f'{len(records)} revisioni conservate. Nel calendario è visibile solo l’ultima revisione di ciascuna impresa, cantiere e settimana.')


def render_weekly_archive(store):
    st.subheader('\U0001f5c4\ufe0f Archivio dei sinottici')
    records=store.all_weekly();st.metric('Revisioni settimanali archiviate',len(records))
    st.caption('I backup precedenti restano importabili. I riferimenti POS mancanti si completano caricando la pianificazione per una nuova revisione; lo storico originale non viene modificato.')
    st.download_button('Scarica backup settimanale completo',weekly_backup(records),'HSE_sinottici_backup.json','application/json')
    uploaded=st.file_uploader('Importa un backup settimanale',type=['json'],key='weekly_backup_file')
    if uploaded is not None and st.button('Verifica e importa i sinottici'):
        try:
            incoming=parse_weekly_backup(uploaded.getvalue());store.import_weekly(incoming)
            st.success('Backup importato. Le pianificazioni confermate sono ora disponibili nel calendario.');st.rerun()
        except Exception:st.error('Backup non valido, revisione in conflitto o archivio non disponibile. Nessun import parziale eseguito.')
    if store.mode=='session':st.warning('Scarica il backup prima di chiudere o ricaricare il browser. Il salvataggio di sessione non è un archivio cloud permanente.')


def choose_portal():
    if st.session_state.get('portal_role') in ['HSE Enel','HSE impresa']:return True
    st.markdown('<div class="hero"><span class="eyebrow">WIND · SOLAR · BESS</span><h1>HSE Risk Platform</h1><p>Valutazioni, pianificazione e coordinamento operativo.</p></div>',unsafe_allow_html=True)
    a,b=st.columns(2)
    with a:
        with st.container(border=True):
            st.markdown('## \U0001f9ba HSE Enel');st.write('Valuta i cantieri, analizza le priorità e consulta il calendario delle imprese.')
            if st.button('HSE Enel',type='primary',width='stretch'):st.session_state.portal_role='HSE Enel';st.rerun()
    with b:
        with st.container(border=True):
            st.markdown('## \U0001f3d7\ufe0f HSE impresa');st.write('Pianifica la settimana, prepara le verifiche preliminari e inserisci il sinottico.')
            if st.button('HSE impresa',type='primary',width='stretch'):st.session_state.portal_role='HSE impresa';st.rerun()
    st.caption('Scegli il percorso operativo. Nessuna attività viene inserita nel calendario senza conferma esplicita.')
    return False

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
    if not choose_portal():
        return
    role=st.session_state.portal_role
    with st.sidebar:
        st.markdown('## HSE Risk Platform')
        st.caption('\U0001f9ba '+role)
        if st.button('Cambia percorso',icon='↔\ufe0f'):
            st.session_state.pop('portal_role',None);st.rerun()
        if role=='HSE impresa':
            company_page=st.radio('\U0001f9ed Percorso impresa',['Attività e sinottico','Calendario condiviso','Archivio settimanale'])
        st.caption('WIND · SOLAR · BESS')
        if role=='HSE Enel':
            page = st.radio('\U0001f9ed Navigazione', ['Valutazione settimanale', 'Risultati e azioni', 'Storico cantiere', 'Confronto cantieri', 'Archivio e backup', 'Calendario condiviso'], format_func=lambda label: {'Valutazione settimanale':'\U0001f4dd Valutazione settimanale', 'Risultati e azioni':'\U0001f3af Risultati e azioni', 'Storico cantiere':'\U0001f5d3\ufe0f Storico cantiere', 'Confronto cantieri':'\U0001f4ca Confronto cantieri', 'Archivio e backup':'\U0001f5c4\ufe0f Archivio e backup', 'Calendario condiviso':'\U0001f5d3\ufe0f Calendario condiviso'}[label])
        st.divider()
        st.caption(f'Interfaccia {UI_VERSION} · modello {MODEL_VERSION} · pesi sperimentali')
        st.caption('Periodo: una settimana ISO (lunedì–domenica). Nessuna API AI. Output generati da regole esplicite.')
        st.caption('Archivio: ' + {'session':'sessione + backup portabile','sqlite':'SQLite locale','postgres':'PostgreSQL condiviso'}[mode])
        if password and st.button('Esci'):
            st.session_state.clear()
            st.rerun()
    if role=='HSE impresa':
        if company_page=='Attività e sinottico':render_company_planning(store)
        elif company_page=='Calendario condiviso':render_weekly_calendar(store)
        else:render_weekly_archive(store)
        return
    if page == 'Calendario condiviso':
        render_weekly_calendar(store)
    elif page == 'Valutazione settimanale':
        render_form(records)
    elif page == 'Risultati e azioni':
        render_report(store, records)
    elif page == 'Storico cantiere':
        render_history(records)
    elif page == 'Confronto cantieri':
        render_comparison(records)
    else:
        render_archive(store, records)
        render_weekly_archive(store)


if __name__ == '__main__':
    main()
