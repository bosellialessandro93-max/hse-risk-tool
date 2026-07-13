from __future__ import annotations

from html import escape
from typing import Final
from urllib.parse import quote

import streamlit as st


# ============================================================
# CONFIGURAZIONE
# ============================================================

APP_NAME: Final = "HSE Document Hub"
APP_VERSION: Final = "1.1"

SHAREPOINT_SITE_URL: Final = (
    "https://enelcom.sharepoint.com/sites/"
    "ArchivioDocumentaleHSECantieri"
)

SHAREPOINT_SITE_PATH: Final = (
    "/sites/ArchivioDocumentaleHSECantieri"
)

DOCUMENT_LIBRARY: Final = "Shared Documents"
ROOT_FOLDER: Final = "Archivio Documentale HSE Cantieri"


# Devono corrispondere esattamente alle cartelle SharePoint.
SITI: Final = [
    "Udine BESS",
    "Fusina BESS",
    "Pontestura",
    "La Spezia",
    "Acate",
    "Vizzini",
]


# Devono corrispondere esattamente alle sottocartelle SharePoint.
CATEGORIE: Final = [
    "Pre Job Check",
    "Checklist Mezzi",
    "Checklist Attrezzature",
    "Sensibilizzazioni",
    "Induction",
    "Test",
    "Procedure",
    "Costi della sicurezza",
    "Documentazione generica",
]


CATEGORY_ICONS: Final = {
    "Pre Job Check": "🧾",
    "Checklist Mezzi": "🚚",
    "Checklist Attrezzature": "🛠️",
    "Sensibilizzazioni": "📣",
    "Induction": "🎓",
    "Test": "🧪",
    "Procedure": "📘",
    "Costi della sicurezza": "💶",
    "Documentazione generica": "📂",
}


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def safe(value: object) -> str:
    return escape(str(value or ""))


def render_html(content: str) -> None:
    """
    Visualizza HTML reale senza elaborazione Markdown.
    In questo modo i tag <div> non vengono mai mostrati come testo.
    """
    st.html(content)


def configure_page() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🛡️",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def initialize_state() -> None:
    defaults = {
        "role": None,
        "page": "home",
        "selected_site": SITI[0],
        "selected_category": CATEGORIE[0],
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def go_to(page: str, role: str | None = None) -> None:
    st.session_state.page = page

    if role is not None:
        st.session_state.role = role


# ============================================================
# STILE
# ============================================================

def apply_styles() -> None:
    st.html(
        """
        <style>
        :root {
            --navy: #12304a;
            --blue: #176b87;
            --cyan: #24a3b6;
            --ice: #eef6f8;
            --line: #dbe6ea;
            --muted: #647783;
            --shadow: 0 10px 30px rgba(18, 48, 74, 0.08);
        }

        .stApp {
            background:
                radial-gradient(
                    circle at 100% 0%,
                    rgba(36, 163, 182, 0.11),
                    transparent 28rem
                ),
                linear-gradient(
                    180deg,
                    #f8fbfc 0%,
                    #f3f7f8 100%
                );
        }

        .block-container {
            max-width: 1500px;
            padding-top: 1.6rem;
            padding-bottom: 4rem;
        }

        h1, h2, h3 {
            color: var(--navy);
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(
                    180deg,
                    #102d46 0%,
                    #174f68 100%
                );
        }

        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] label {
            color: white !important;
        }

        [data-testid="stSidebar"] .stButton > button {
            background: rgba(255,255,255,0.12);
            color: white;
            border: 1px solid rgba(255,255,255,0.22);
        }

        [data-testid="stSidebar"] .stLinkButton > a {
            background: rgba(255,255,255,0.12);
            color: white !important;
            border: 1px solid rgba(255,255,255,0.22);
        }

        [data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background: white !important;
            color: #16232d !important;
        }

        [data-testid="stSidebar"] div[data-baseweb="select"] span {
            color: #16232d !important;
        }

        .hero {
            padding: 2.2rem 2.4rem;
            border-radius: 24px;
            background:
                linear-gradient(
                    120deg,
                    rgba(18,48,74,0.98),
                    rgba(23,107,135,0.94)
                );
            box-shadow: var(--shadow);
            color: white;
            position: relative;
            overflow: hidden;
            margin-bottom: 1.4rem;
        }

        .hero::after {
            content: "";
            width: 300px;
            height: 300px;
            border-radius: 50%;
            position: absolute;
            right: -80px;
            top: -110px;
            background: rgba(255,255,255,0.08);
        }

        .hero-kicker {
            text-transform: uppercase;
            font-size: 0.76rem;
            letter-spacing: 0.14em;
            font-weight: 700;
            opacity: 0.82;
            position: relative;
            z-index: 1;
        }

        .hero-title {
            font-size: clamp(2rem, 4vw, 3.35rem);
            line-height: 1.02;
            font-weight: 800;
            margin: 0.45rem 0 0.55rem;
            position: relative;
            z-index: 1;
        }

        .hero-copy {
            max-width: 850px;
            opacity: 0.90;
            font-size: 1.03rem;
            position: relative;
            z-index: 1;
        }

        .metric-card,
        .site-card,
        .category-card,
        .panel,
        .access-card,
        .process-step {
            background: rgba(255,255,255,0.96);
            border: 1px solid var(--line);
            box-shadow: var(--shadow);
            border-radius: 18px;
        }

        .metric-card {
            min-height: 138px;
            padding: 1.25rem 1.35rem;
        }

        .metric-label {
            color: var(--muted);
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
        }

        .metric-value {
            color: var(--navy);
            font-size: 2rem;
            font-weight: 800;
            margin-top: 0.45rem;
        }

        .metric-foot {
            color: var(--muted);
            font-size: 0.82rem;
            margin-top: 0.25rem;
        }

        .section-title {
            color: var(--navy);
            font-size: 1.35rem;
            font-weight: 800;
            margin: 1.6rem 0 0.85rem;
        }

        .site-card,
        .category-card {
            padding: 1.2rem 1.25rem;
            min-height: 170px;
            margin-bottom: 0.65rem;
        }

        .site-name,
        .category-name {
            color: var(--navy);
            font-size: 1.12rem;
            font-weight: 800;
        }

        .card-icon {
            font-size: 1.7rem;
            margin-bottom: 0.7rem;
        }

        .card-number {
            color: var(--navy);
            font-size: 1.85rem;
            font-weight: 800;
            margin-top: 0.8rem;
        }

        .card-caption,
        .muted {
            color: var(--muted);
            font-size: 0.86rem;
        }

        .access-card {
            border-radius: 22px;
            padding: 1.8rem;
            min-height: 250px;
        }

        .access-icon {
            font-size: 2.2rem;
        }

        .access-title {
            color: var(--navy);
            font-size: 1.45rem;
            font-weight: 800;
            margin: 0.8rem 0 0.45rem;
        }

        .access-copy {
            color: var(--muted);
            min-height: 72px;
        }

        .success-banner,
        .small-banner,
        .warning-banner {
            padding: 0.9rem 1rem;
            border-radius: 14px;
            margin-bottom: 1rem;
        }

        .success-banner {
            background: #eaf7f0;
            border: 1px solid #c9e9d6;
            color: #195e3a;
        }

        .small-banner {
            background: #eef6f8;
            border: 1px solid #dbe9ed;
            color: #365463;
        }

        .warning-banner {
            background: #fff8e5;
            border: 1px solid #f1dfaa;
            color: #705715;
        }

        .panel {
            padding: 1.4rem;
            margin-top: 1rem;
        }

        .process-step {
            padding: 1rem 1.1rem;
            min-height: 130px;
            margin-bottom: 0.8rem;
        }

        .process-number {
            color: var(--cyan);
            font-size: 1.45rem;
            font-weight: 900;
        }

        .process-title {
            color: var(--navy);
            font-size: 1rem;
            font-weight: 800;
            margin-top: 0.25rem;
        }

        .process-copy {
            color: var(--muted);
            font-size: 0.85rem;
            margin-top: 0.3rem;
        }

        .stButton > button,
        .stLinkButton > a {
            border-radius: 11px;
            font-weight: 750;
            min-height: 42px;
        }

        .stButton > button[kind="primary"] {
            background:
                linear-gradient(
                    90deg,
                    #176b87,
                    #1b829e
                );
            border: none;
        }

        .footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--line);
            color: var(--muted);
            font-size: 0.8rem;
            text-align: center;
        }
        </style>
        """
    )


# ============================================================
# LINK SHAREPOINT
# ============================================================

def sharepoint_folder_path(
    site: str | None = None,
    category: str | None = None,
) -> str:
    parts = [
        SHAREPOINT_SITE_PATH,
        DOCUMENT_LIBRARY,
        ROOT_FOLDER,
    ]

    if site:
        parts.append(site)

    if category:
        parts.append(category)

    return "/" + "/".join(
        part.strip("/")
        for part in parts
    ).lstrip("/")


def sharepoint_folder_url(
    site: str | None = None,
    category: str | None = None,
) -> str:
    folder_path = sharepoint_folder_path(
        site=site,
        category=category,
    )

    encoded_path = quote(
        folder_path,
        safe="",
    )

    return (
        f"{SHAREPOINT_SITE_URL}/"
        f"{quote(DOCUMENT_LIBRARY, safe='')}/"
        f"Forms/AllItems.aspx"
        f"?id={encoded_path}"
    )


# ============================================================
# COMPONENTI
# ============================================================

def render_header(title: str, subtitle: str) -> None:
    render_html(
        f"""
        <div class="hero">
            <div class="hero-kicker">Portale HSE cantieri</div>
            <div class="hero-title">{safe(title)}</div>
            <div class="hero-copy">{safe(subtitle)}</div>
        </div>
        """
    )


def metric_card(label: str, value: object, foot: str) -> None:
    render_html(
        f"""
        <div class="metric-card">
            <div class="metric-label">{safe(label)}</div>
            <div class="metric-value">{safe(value)}</div>
            <div class="metric-foot">{safe(foot)}</div>
        </div>
        """
    )


def render_footer() -> None:
    render_html(
        f"""
        <div class="footer">
            HSE Document Hub · Archivio documentale cantieri
            · Versione {APP_VERSION}
        </div>
        """
    )


# ============================================================
# SIDEBAR
# ============================================================

def render_sidebar() -> None:
    if st.session_state.page == "home":
        return

    role_label = (
        "Accesso Enel"
        if st.session_state.role == "enel"
        else "Accesso Imprese"
    )

    with st.sidebar:
        st.markdown("## 🛡️ HSE Document Hub")
        st.caption(role_label)
        st.divider()

        if st.session_state.role == "enel":
            if st.button("▦ Dashboard", width="stretch"):
                go_to("enel_dashboard")
                st.rerun()

            if st.button("⌕ Ricerca documenti", width="stretch"):
                go_to("search")
                st.rerun()

            if st.button("▤ Gestione archivio", width="stretch"):
                go_to("documents")
                st.rerun()

        else:
            if st.button("⬆ Carica documenti", width="stretch"):
                go_to("company_upload")
                st.rerun()

            if st.button("▤ Consulta archivio", width="stretch"):
                go_to("company_documents")
                st.rerun()

        st.divider()
        st.caption("Navigazione rapida")

        selected_site = st.selectbox(
            "Cantiere",
            SITI,
            index=SITI.index(
                st.session_state.selected_site
            ),
            key="sidebar_site",
        )

        st.session_state.selected_site = selected_site

        if st.button(
            "Apri cantiere nel portale",
            width="stretch",
        ):
            go_to("site")
            st.rerun()

        st.link_button(
            "Apri cantiere su SharePoint",
            sharepoint_folder_url(site=selected_site),
            width="stretch",
        )

        st.divider()

        if st.button(
            "← Esci dall'area",
            width="stretch",
        ):
            st.session_state.role = None
            st.session_state.page = "home"
            st.rerun()


# ============================================================
# HOME
# ============================================================

def render_home() -> None:
    render_header(
        APP_NAME,
        (
            "Un unico punto di accesso per raggiungere, consultare "
            "e archiviare la documentazione HSE dei cantieri."
        ),
    )

    render_html(
        """
        <div class="success-banner">
            Il portale indirizza gli utenti verso la cartella
            SharePoint corretta. I documenti restano conservati
            esclusivamente nell'archivio aziendale.
        </div>
        """
    )

    col1, col2 = st.columns(2, gap="large")

    with col1:
        render_html(
            """
            <div class="access-card">
                <div class="access-icon">🏢</div>
                <div class="access-title">Accesso Enel</div>
                <div class="access-copy">
                    Dashboard dei cantieri, ricerca guidata,
                    consultazione delle categorie e accesso
                    alle cartelle SharePoint.
                </div>
            </div>
            """
        )

        if st.button(
            "Entra nell'area Enel",
            type="primary",
            width="stretch",
        ):
            go_to("enel_dashboard", "enel")
            st.rerun()

    with col2:
        render_html(
            """
            <div class="access-card">
                <div class="access-icon">👷</div>
                <div class="access-title">Accesso Imprese</div>
                <div class="access-copy">
                    Percorso guidato per selezionare il cantiere,
                    scegliere la categoria e aprire la cartella
                    corretta per il caricamento.
                </div>
            </div>
            """
        )

        if st.button(
            "Entra nell'area Imprese",
            width="stretch",
        ):
            go_to("company_upload", "company")
            st.rerun()


# ============================================================
# DASHBOARD ENEL
# ============================================================

def render_enel_dashboard() -> None:
    render_header(
        "Dashboard documentale",
        (
            "Accesso centralizzato all'archivio SharePoint, "
            "ai cantieri e alle relative categorie documentali."
        ),
    )

    columns = st.columns(4)

    metrics = [
        (
            "Cantieri configurati",
            len(SITI),
            "Cartelle principali disponibili",
        ),
        (
            "Categorie per cantiere",
            len(CATEGORIE),
            "Struttura documentale standard",
        ),
        (
            "Archivio ufficiale",
            "SharePoint",
            "Conservazione aziendale",
        ),
        (
            "Stato collegamento",
            "Operativo",
            "Apertura diretta delle cartelle",
        ),
    ]

    for column, metric in zip(columns, metrics):
        with column:
            metric_card(*metric)

    render_html(
        '<div class="section-title">Accesso generale</div>'
    )

    col1, col2 = st.columns(2)

    with col1:
        st.link_button(
            "Apri archivio generale SharePoint",
            sharepoint_folder_url(),
            type="primary",
            width="stretch",
        )

    with col2:
        if st.button(
            "Apri ricerca guidata",
            width="stretch",
        ):
            go_to("search")
            st.rerun()

    render_html(
        '<div class="section-title">Cantieri</div>'
    )

    rows = [
        SITI[index:index + 3]
        for index in range(0, len(SITI), 3)
    ]

    for row in rows:
        cards = st.columns(3)

        for card, site in zip(cards, row):
            with card:
                render_html(
                    f"""
                    <div class="site-card">
                        <div class="card-icon">🏗️</div>
                        <div class="site-name">{safe(site)}</div>
                        <div class="card-number">{len(CATEGORIE)}</div>
                        <div class="card-caption">
                            categorie documentali configurate
                        </div>
                    </div>
                    """
                )

                if st.button(
                    "Apri nel portale",
                    key=f"portal_{site}",
                    width="stretch",
                ):
                    st.session_state.selected_site = site
                    go_to("site")
                    st.rerun()

                st.link_button(
                    "Apri su SharePoint",
                    sharepoint_folder_url(site=site),
                    width="stretch",
                )


# ============================================================
# PAGINA CANTIERE
# ============================================================

def render_site() -> None:
    site = st.session_state.selected_site

    render_header(
        site,
        (
            "Consulta le categorie documentali e accedi "
            "direttamente alle relative cartelle SharePoint."
        ),
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        site = st.selectbox(
            "Seleziona il cantiere",
            SITI,
            index=SITI.index(site),
            key="site_page_selector",
        )

        st.session_state.selected_site = site

    with col2:
        st.write("")
        st.write("")

        st.link_button(
            "Apri cantiere su SharePoint",
            sharepoint_folder_url(site=site),
            type="primary",
            width="stretch",
        )

    rows = [
        CATEGORIE[index:index + 3]
        for index in range(0, len(CATEGORIE), 3)
    ]

    for row in rows:
        cards = st.columns(3)

        for card, category in zip(cards, row):
            with card:
                render_html(
                    f"""
                    <div class="category-card">
                        <div class="card-icon">
                            {CATEGORY_ICONS[category]}
                        </div>
                        <div class="category-name">
                            {safe(category)}
                        </div>
                        <div class="card-number">📁</div>
                        <div class="card-caption">
                            Cartella documentale SharePoint
                        </div>
                    </div>
                    """
                )

                if st.button(
                    "Seleziona categoria",
                    key=f"select_{site}_{category}",
                    width="stretch",
                ):
                    st.session_state.selected_category = category
                    go_to("documents")
                    st.rerun()

                st.link_button(
                    "Apri cartella",
                    sharepoint_folder_url(
                        site=site,
                        category=category,
                    ),
                    width="stretch",
                )


# ============================================================
# GESTIONE ARCHIVIO
# ============================================================

def render_documents() -> None:
    render_header(
        "Gestione archivio",
        (
            "Seleziona il cantiere e la categoria per raggiungere "
            "la cartella documentale corretta."
        ),
    )

    col1, col2 = st.columns(2)

    with col1:
        site = st.selectbox(
            "Cantiere",
            SITI,
            index=SITI.index(
                st.session_state.selected_site
            ),
            key="documents_site",
        )

    with col2:
        category = st.selectbox(
            "Categoria",
            CATEGORIE,
            index=CATEGORIE.index(
                st.session_state.selected_category
            ),
            key="documents_category",
        )

    st.session_state.selected_site = site
    st.session_state.selected_category = category

    render_html(
        f"""
        <div class="panel">
            <div class="category-name">
                {CATEGORY_ICONS[category]} Destinazione selezionata
            </div>
            <div style="margin-top:0.8rem;">
                <strong>Cantiere:</strong> {safe(site)}
            </div>
            <div style="margin-top:0.35rem;">
                <strong>Categoria:</strong> {safe(category)}
            </div>
            <div class="muted" style="margin-top:0.8rem;">
                Il pulsante apre direttamente la cartella
                SharePoint corrispondente.
            </div>
        </div>
        """
    )

    st.write("")

    st.link_button(
        "Apri la cartella documentale",
        sharepoint_folder_url(
            site=site,
            category=category,
        ),
        type="primary",
        width="stretch",
    )


# ============================================================
# RICERCA
# ============================================================

def render_search() -> None:
    render_header(
        "Ricerca documenti",
        (
            "Individua il cantiere e la categoria prima "
            "di aprire l'archivio SharePoint."
        ),
    )

    query = st.text_input(
        "Filtra cantieri e categorie",
        placeholder=(
            "Esempio: Udine, checklist, procedure, mezzi..."
        ),
    )

    normalized_query = query.strip().lower()

    matching_sites = [
        site
        for site in SITI
        if not normalized_query
        or normalized_query in site.lower()
    ]

    matching_categories = [
        category
        for category in CATEGORIE
        if not normalized_query
        or normalized_query in category.lower()
    ]

    col1, col2 = st.columns(2)

    with col1:
        render_html(
            '<div class="section-title">Cantieri</div>'
        )

        for site in matching_sites:
            st.link_button(
                f"🏗️ {site}",
                sharepoint_folder_url(site=site),
                width="stretch",
            )

    with col2:
        render_html(
            '<div class="section-title">Categorie trovate</div>'
        )

        for category in matching_categories:
            st.write(
                f"{CATEGORY_ICONS[category]} **{category}**"
            )

    render_html(
        '<div class="section-title">Ricerca mirata</div>'
    )

    col3, col4 = st.columns(2)

    with col3:
        selected_site = st.selectbox(
            "Cantiere da consultare",
            SITI,
            key="search_site",
        )

    with col4:
        selected_category = st.selectbox(
            "Categoria da consultare",
            CATEGORIE,
            key="search_category",
        )

    st.link_button(
        "Apri la cartella selezionata",
        sharepoint_folder_url(
            site=selected_site,
            category=selected_category,
        ),
        type="primary",
        width="stretch",
    )


# ============================================================
# AREA IMPRESE
# ============================================================

def render_company_upload() -> None:
    render_header(
        "Caricamento documenti",
        (
            "Seleziona il cantiere e la categoria. "
            "Il portale aprirà la cartella SharePoint corretta."
        ),
    )

    render_html(
        """
        <div class="success-banner">
            I file non vengono memorizzati su Streamlit.
            Il caricamento avviene direttamente su SharePoint.
        </div>
        """
    )

    col1, col2 = st.columns(2)

    with col1:
        site = st.selectbox(
            "1. Seleziona il cantiere",
            SITI,
            index=SITI.index(
                st.session_state.selected_site
            ),
            key="company_upload_site",
        )

    with col2:
        category = st.selectbox(
            "2. Seleziona la categoria",
            CATEGORIE,
            index=CATEGORIE.index(
                st.session_state.selected_category
            ),
            key="company_upload_category",
        )

    st.session_state.selected_site = site
    st.session_state.selected_category = category

    render_html(
        '<div class="section-title">Procedura di caricamento</div>'
    )

    columns = st.columns(4)

    steps = [
        (
            "01",
            "Apri la cartella",
            "Usa il pulsante presente in fondo alla pagina.",
        ),
        (
            "02",
            "Accedi con Microsoft",
            "Utilizza il tuo account aziendale.",
        ),
        (
            "03",
            "Clicca Carica",
            "In SharePoint scegli Carica e poi File.",
        ),
        (
            "04",
            "Seleziona i documenti",
            "I file saranno conservati nella cartella scelta.",
        ),
    ]

    for column, step in zip(columns, steps):
        number, title, description = step

        with column:
            render_html(
                f"""
                <div class="process-step">
                    <div class="process-number">{safe(number)}</div>
                    <div class="process-title">{safe(title)}</div>
                    <div class="process-copy">{safe(description)}</div>
                </div>
                """
            )

    render_html(
        f"""
        <div class="panel">
            <div class="category-name">
                Destinazione del caricamento
            </div>
            <div style="margin-top:0.8rem;">
                <strong>Cantiere:</strong> {safe(site)}
            </div>
            <div style="margin-top:0.35rem;">
                <strong>Categoria:</strong> {safe(category)}
            </div>
        </div>
        """
    )

    st.write("")

    st.link_button(
        "Apri SharePoint e carica i documenti",
        sharepoint_folder_url(
            site=site,
            category=category,
        ),
        type="primary",
        width="stretch",
    )


def render_company_documents() -> None:
    render_header(
        "Consulta archivio",
        (
            "Accedi ai documenti già caricati selezionando "
            "il cantiere e la categoria."
        ),
    )

    col1, col2 = st.columns(2)

    with col1:
        site = st.selectbox(
            "Cantiere",
            SITI,
            key="company_documents_site",
        )

    with col2:
        category_options = [
            "Tutte le categorie",
            *CATEGORIE,
        ]

        category = st.selectbox(
            "Categoria",
            category_options,
            key="company_documents_category",
        )

    if category == "Tutte le categorie":
        target_url = sharepoint_folder_url(site=site)
        destination = f"{site} / Tutte le categorie"
    else:
        target_url = sharepoint_folder_url(
            site=site,
            category=category,
        )
        destination = f"{site} / {category}"

    render_html(
        f"""
        <div class="panel">
            <div class="category-name">
                Archivio selezionato
            </div>
            <div style="margin-top:0.8rem;">
                {safe(destination)}
            </div>
        </div>
        """
    )

    st.write("")

    st.link_button(
        "Apri i documenti su SharePoint",
        target_url,
        type="primary",
        width="stretch",
    )


# ============================================================
# AVVIO
# ============================================================

def main() -> None:
    configure_page()
    apply_styles()
    initialize_state()
    render_sidebar()

    page = st.session_state.page

    if page == "home":
        render_home()

    elif page == "enel_dashboard":
        render_enel_dashboard()

    elif page == "site":
        render_site()

    elif page == "documents":
        render_documents()

    elif page == "search":
        render_search()

    elif page == "company_upload":
        render_company_upload()

    elif page == "company_documents":
        render_company_documents()

    else:
        st.session_state.page = "home"
        st.session_state.role = None
        st.rerun()

    render_footer()


if __name__ == "__main__":
    main()
