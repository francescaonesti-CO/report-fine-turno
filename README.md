# Report Turno Polizia Locale

Web app React/Vite per compilare il report giornaliero degli operatori di Polizia Locale.

## Funzioni incluse

- Selezione turno con orari predefiniti
- Ordinario / straordinario
- Reparto con campo "Altri servizi"
- Operatori e matricole
- Veicoli e chilometraggio con calcolo automatico
- Interventi dinamici con origine: Centrale Operativa, UDT, Di iniziativa, Altro
- TSO e ASO separati
- Servizio scuole con massimo 3 scuole
- Atti redatti
- Violazioni, fermi e sequestri
- Anteprima report
- PDF scaricabile
- Email precompilata con mailto

## Installazione locale

```bash
npm install
npm run dev
```

## Deploy su Vercel

1. Carica questa cartella su GitHub.
2. Vai su Vercel > Add New Project.
3. Importa la repository.
4. Framework: Vite.
5. Build command: `npm run build`.
6. Output directory: `dist`.
7. Deploy.

## Nota sull'invio email

Questa versione usa il metodo sicuro B1:
- scarica il PDF;
- apre l'email precompilata;
- l'operatore allega manualmente il PDF.

Per l'invio automatico completo con allegato serve una successiva integrazione con backend o servizio email.
