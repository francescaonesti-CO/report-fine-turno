# Report Turno Polizia Locale

Web app React/Vite per la compilazione del report di servizio e dashboard aggregata per l'ufficiale.

## Funzioni incluse

### Report operatore
- Compilazione dati turno
- Scelta orario ordinario/straordinario
- Reparto/servizio
- Operatori, veicoli, km
- Interventi dinamici
- Servizio scuole fino a 3 scuole
- Atti redatti
- Violazioni e provvedimenti
- PDF scaricabile
- Email precompilata livello 1
- Esportazione file dati JSON per dashboard

### Dashboard ufficiale
- Caricamento multiplo file JSON ricevuti dagli operatori
- Aggregazione dati giornalieri o multi-turno
- Totale report ricevuti
- Totale interventi
- Totale violazioni/provvedimenti
- Totale km
- Distribuzione per tipologia intervento
- Distribuzione per origine intervento
- Distribuzione per reparto
- Note e criticità
- PDF aggregato per Comandante
- Esportazione tabella CSV

## Installazione locale

```bash
npm install
npm run dev
```

## Pubblicazione su Vercel

1. Caricare tutti i file del progetto in un repository GitHub.
2. Importare il repository in Vercel.
3. Usare le impostazioni automatiche Vite:
   - Build Command: `npm run build`
   - Output Directory: `dist`
   - Install Command: `npm install`
4. Deploy.

## Uso operativo consigliato

1. L'operatore compila il report.
2. Scarica PDF e file dati JSON.
3. Invia entrambi via mail all'ufficiale di coordinamento.
4. L'ufficiale apre la Dashboard ufficiale.
5. Carica tutti i file JSON ricevuti.
6. Genera il PDF aggregato da inviare al Comandante.

## Export Excel avanzato

Nella Dashboard ufficiale è disponibile il pulsante **Esporta Excel avanzato**.
Il file `.xlsx` generato contiene:

- **Dashboard**: KPI principali e sintesi operativa.
- **Grafici**: viste grafiche a barre orizzontali compatibili con Excel.
- **Interventi**: una riga per ogni intervento, pronta per filtri e tabelle pivot.
- **Report**: riepilogo per ciascun report caricato.
- **Riepilogo**: dati aggregati per tipologia, origine, reparto, fascia oraria e contatori.
- **Criticità**: interventi evidenziati automaticamente come rilevanti.
- **Atti_Violazioni**: dettaglio dei contatori per report.
- **README**: guida rapida interna al file Excel.

L'analisi temporale include: mattino, pomeriggio, sera e notte.
