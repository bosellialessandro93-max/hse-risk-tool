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
