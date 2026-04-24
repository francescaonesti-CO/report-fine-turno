# Report Turno Polizia Locale - versione con Report Ufficiale UDT

Questa versione integra:

- Report operatore con PDF istituzionale e logo
- Esportazione file dati JSON per ogni report
- Dashboard ufficiale con aggregazione JSON
- Export Excel avanzato
- Nuovo pulsante **Report ufficiale**
- Generazione PDF **Report Ufficiale di Turno** in modalità quasi automatica

## Flusso consigliato

1. Gli operatori compilano il report e scaricano PDF + JSON.
2. L'ufficiale apre **Report ufficiale** o **Dashboard ufficiale**.
3. Carica i JSON ricevuti dagli operatori.
4. L'app genera sintesi, eventi rilevanti e tabella violazioni.
5. L'ufficiale integra briefing, personale assente/in ritardo, anomalie, attività ispettive, esiti e note.
6. Clicca **Genera PDF Report Ufficiale**.

## Deploy su Vercel

- Framework: Vite
- Install Command: `npm install`
- Build Command: `npm run build`
- Output Directory: `dist`

Caricare su GitHub tutti i file della cartella estratta, non lo ZIP.
