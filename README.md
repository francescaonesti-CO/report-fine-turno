# Report Turno Polizia Locale - versione con salvataggio automatico

Questa versione integra:

- Report operatore con PDF istituzionale e logo
- Salvataggio automatico della bozza operatore sul dispositivo
- Pulsante **Nuovo turno / cancella bozza**
- Esportazione file dati JSON per ogni report
- Dashboard ufficiale con aggregazione JSON
- Export Excel avanzato
- Pulsante **Report ufficiale**
- Generazione PDF **Report Ufficiale di Turno** in modalità quasi automatica

## Salvataggio automatico

Il report operatore viene salvato automaticamente nel browser del dispositivo usato dall'operatore.

Questo significa che l'operatore può:

1. iniziare il report durante il turno;
2. chiudere l'app;
3. riaprirla più tardi;
4. ritrovare i dati già inseriti.

A fine turno deve comunque scaricare:

- PDF del report;
- file dati JSON.

Quando si inizia un nuovo servizio, usare il pulsante **Nuovo turno / cancella bozza**.

## Flusso consigliato

1. Gli operatori compilano il report durante il turno.
2. A fine turno scaricano PDF + JSON.
3. L'ufficiale apre **Report ufficiale** o **Dashboard ufficiale**.
4. Carica i JSON ricevuti dagli operatori.
5. L'app genera sintesi, eventi rilevanti e tabella violazioni.
6. L'ufficiale integra briefing, personale assente/in ritardo, anomalie, attività ispettive, esiti e note.
7. Clicca **Genera PDF Report Ufficiale**.

## Deploy su Vercel

- Framework: Vite
- Install Command: `npm install`
- Build Command: `npm run build`
- Output Directory: `dist`

Caricare su GitHub tutti i file della cartella estratta, non lo ZIP.
Branch di sviluppo per integrazione Supabase.
