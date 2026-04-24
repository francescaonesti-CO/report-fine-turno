# Report Turno Polizia Locale - versione PDF agenti migliorato

Web app React/Vite per:
- compilare il report di servizio degli operatori;
- generare PDF istituzionale con logo e layout migliorato;
- esportare file dati JSON per la dashboard;
- usare la dashboard ufficiale con report aggregato ed export Excel avanzato.

## Novità di questa versione
- PDF agenti con header Comune di Monza e logo Polizia Locale.
- Dicitura: Settore Polizia Locale, Protezione Civile.
- Layout più accattivante e leggibile: riepilogo iniziale, sezioni ordinate, schede intervento, footer istituzionale.
- Mantenute dashboard, PDF aggregato ed export Excel avanzato.

## Deploy su Vercel
1. Scompattare lo ZIP.
2. Caricare tutti i file nella root del repository GitHub.
3. Importare il repository su Vercel.
4. Impostazioni:
   - Framework: Vite
   - Build command: npm run build
   - Output directory: dist
   - Install command: npm install

## Attenzione
Il file `public/POLIZIA.png` deve restare nel progetto: viene usato per il logo nei PDF.
