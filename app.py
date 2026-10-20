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
