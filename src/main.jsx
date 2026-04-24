import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import jsPDF from 'jspdf';
import * as XLSX from 'xlsx';
import './style.css';

const TURNI = [
  '06.00-13.00', '06.30-13.30', '07.00-14.00', '08.00-15.00', '09.00-16.00',
  '11.00-18.00', '12.00-19.00', '12.30-19.30', '13.00-20.00', '16.59-23.59',
  '00.00-07.00', 'Altro orario'
];

const REPARTI = [
  'Radiomobile', 'Quartieri', 'Tutela del consumatore e delle imprese', 'Tutela ambiente e paesaggio',
  'Polizia stradale', 'Intervento rapido motociclisti', 'Centrale operativa', 'N.O.S.T.',
  'Corso di formazione', 'Altri servizi'
];

const TIPI_INTERVENTO = [
  'Sinistro stradale', 'TSO', 'ASO', 'Posto di controllo', 'Viabilità', 'Servizio scuole',
  'Controllo commerciale / annonaria', 'Controllo edilizio', 'Controllo parchi / aree verdi',
  'Sicurezza urbana', 'Intervento per animali', 'Abbandono rifiuti', 'Rumori / disturbo quiete',
  'Supporto ad altro ente', 'Notifica / accertamento', 'Altro'
];

const ORIGINI = ['Centrale Operativa', 'UDT', 'Di iniziativa', 'Altro'];

const emptyOperatore = () => ({ nome: '', matricola: '', qualifica: '' });
const emptyVeicolo = () => ({ sigla: '', kmInizio: '', kmFine: '' });
const emptyScuola = () => ({ nome: '', momento: '', orario: '', criticita: '' });
const emptyIntervento = () => ({
  tipo: 'Sinistro stradale', origine: 'Centrale Operativa', origineAltro: '', oraInizio: '', oraFine: '', luogo: '', descrizione: '', esito: '', note: '',
  conFeriti: 'Senza feriti', veicoliCoinvolti: '', rilievi: 'No', personeControllate: '', veicoliControllati: '', verbaliElevati: '', fermiSequestri: '',
  motivoViabilita: '', strade: '', scuole: [emptyScuola()]
});

const emptyCounters = () => ({
  relazioni: 0, annotazioni: 0, verbaliCds: 0, verbaliRegolamenti: 0, sequestriAmministrativi: 0,
  fermiAmministrativi: 0, sequestriPenali: 0, cnr: 0, altriAttiNumero: 0, altriAttiDescrizione: '',
  preavvisiCds: 0, vdcCds: 0, regPolizia: 0, regEdilizio: 0, regBenessereAnimali: 0,
  annonaria: 0, altreNorme: 0, altreNormeDescrizione: '', fermi: 0, sequestri: 0
});

const LABELS = {
  relazioni: 'Relazioni di servizio', annotazioni: 'Annotazioni di servizio', verbaliCds: 'Verbali CdS', verbaliRegolamenti: 'Verbali regolamenti',
  sequestriAmministrativi: 'Sequestri amministrativi', fermiAmministrativi: 'Fermi amministrativi', sequestriPenali: 'Sequestri penali', cnr: 'C.N.R.',
  altriAttiNumero: 'Altri atti', preavvisiCds: 'Preavvisi CdS', vdcCds: 'VdC CdS', regPolizia: 'Regolamento Polizia',
  regEdilizio: 'Regolamento Edilizio', regBenessereAnimali: 'Regolamento Benessere Animali', annonaria: 'Annonaria / commercio',
  altreNorme: 'Altre norme', fermi: 'Fermi', sequestri: 'Sequestri'
};

function today() { return new Date().toISOString().slice(0, 10); }
function n(value) { return Number(value || 0); }
function km(v) { return Math.max(0, n(v.kmFine) - n(v.kmInizio)); }
function turnoLabel(report) { return report.turno === 'Altro orario' ? `${report.altroTurnoInizio || '?'}-${report.altroTurnoFine || '?'}` : report.turno; }
function repartoLabel(report) { return report.reparto === 'Altri servizi' ? `${report.reparto}: ${report.altroServizio || '-'}` : report.reparto; }
function sanitizeFileName(s) { return String(s || 'report').replace(/[^a-z0-9._-]+/gi, '-').replace(/-+/g, '-'); }

function Field({ label, children }) { return <label className="field"><span>{label}</span>{children}</label>; }
function Input({ value, onChange, type = 'text', placeholder = '' }) { return <input type={type} value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} />; }
function Textarea({ value, onChange, placeholder = '' }) { return <textarea value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} />; }
function Select({ value, onChange, children }) { return <select value={value} onChange={e => onChange(e.target.value)}>{children}</select>; }
function Counter({ label, value, onChange }) {
  const valueNumber = n(value);
  return <div className="counter"><span>{label}</span><button type="button" onClick={() => onChange(Math.max(0, valueNumber - 1))}>−</button><input type="number" min="0" value={valueNumber} onChange={e => onChange(n(e.target.value))} /><button type="button" onClick={() => onChange(valueNumber + 1)}>+</button></div>;
}

function baseReport() {
  return {
    schemaVersion: 2,
    data: today(), turno: '06.00-13.00', altroTurnoInizio: '', altroTurnoFine: '', orarioTipo: 'Ordinario',
    reparto: 'Radiomobile', altroServizio: '', destinatario: '',
    operatori: [emptyOperatore(), emptyOperatore()], veicoli: [emptyVeicolo()], interventi: [emptyIntervento()],
    counters: emptyCounters(), noteUdt: '', dichiarazione: false, createdAt: new Date().toISOString()
  };
}


const emptyAttivitaIspettiva = () => ({ tipo: '', reparto: '', luogo: '', orario: '', esito: '', violazioni: '', note: '' });
function baseOfficialReport() {
  return {
    data: today(), turno: '1° turno', ufficiale: '', qualifica: '', briefing: '', assenti: '', ritardi: '', noteGenerali: '',
    eventiManuali: '', anomalie: '', attivitaIspettive: [emptyAttivitaIspettiva()], esiti: '', comunicazioneEq: '', notaComandante: ''
  };
}

function App() {
  const [mode, setMode] = useState('operatore');
  const [report, setReport] = useState(baseReport());
  const [importedReports, setImportedReports] = useState([]);
  const [officialReport, setOfficialReport] = useState(baseOfficialReport());

  return <main>
    <img id="pdfLogo" src="/POLIZIA.png" alt="Logo Polizia Locale" style={{ display: 'none' }} />
    <header className="hero">
      <div>
        <p className="eyebrow">Polizia Locale</p>
        <h1>Report Turno</h1>
        <p>Compilazione operatori, dashboard aggregata e report ufficiale UDT.</p>
      </div>
      <nav className="tabs">
        <button className={mode === 'operatore' ? 'active' : ''} onClick={() => setMode('operatore')}>Report operatore</button>
        <button className={mode === 'dashboard' ? 'active' : ''} onClick={() => setMode('dashboard')}>Dashboard ufficiale</button>
        <button className={mode === 'ufficiale' ? 'active' : ''} onClick={() => setMode('ufficiale')}>Report ufficiale</button>
      </nav>
    </header>
    {mode === 'operatore' && <OperatorReport report={report} setReport={setReport} />}
    {mode === 'dashboard' && <Dashboard reports={importedReports} setReports={setImportedReports} />}
    {mode === 'ufficiale' && <OfficialReport reports={importedReports} setReports={setImportedReports} official={officialReport} setOfficial={setOfficialReport} />}
  </main>;
}

function OperatorReport({ report, setReport }) {
  const update = (patch) => setReport(prev => ({ ...prev, ...patch }));
  const updateArray = (key, index, patch) => setReport(prev => ({ ...prev, [key]: prev[key].map((x, i) => i === index ? { ...x, ...patch } : x) }));
  const addArray = (key, item) => setReport(prev => ({ ...prev, [key]: [...prev[key], item] }));
  const removeArray = (key, index) => setReport(prev => ({ ...prev, [key]: prev[key].filter((_, i) => i !== index) }));

  const totalKm = useMemo(() => report.veicoli.reduce((sum, v) => sum + km(v), 0), [report.veicoli]);
  const totaleViolazioni = useMemo(() => getTotaleViolazioni(report), [report]);
  const text = useMemo(() => reportText(report), [report, totalKm, totaleViolazioni]);

  function generatePdf() {
    const doc = buildServicePdf(report);
    doc.save(`report-turno-${sanitizeFileName(report.data)}-${sanitizeFileName(turnoLabel(report))}.pdf`);
  }

  function exportJson() {
    const payload = { ...report, schemaVersion: 2, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `dati-report-${sanitizeFileName(report.data)}-${sanitizeFileName(turnoLabel(report))}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function sendMail() {
    const subject = encodeURIComponent(`Report turno Polizia Locale - ${report.data} - ${turnoLabel(report)}`);
    const body = encodeURIComponent(`Si trasmette il report del turno di servizio.\n\nAllegare il PDF scaricato dall'app e, per la dashboard dell'ufficiale, anche il file dati JSON.\n\n${text.slice(0, 1200)}${text.length > 1200 ? '\n\n[Report completo in allegato PDF]' : ''}`);
    window.location.href = `mailto:${encodeURIComponent(report.destinatario)}?subject=${subject}&body=${body}`;
  }

  return <>
    <section className="card notice">
      <h2>Flusso operativo</h2>
      <p>Al termine del turno l'operatore scarica <strong>PDF</strong> e <strong>file dati JSON</strong>, poi invia entrambi all'ufficiale. L'ufficiale carica i JSON nella dashboard e genera il report aggregato per il Comandante.</p>
    </section>

    <section className="card">
      <h2>1. Dati turno</h2>
      <div className="grid">
        <Field label="Data servizio"><Input type="date" value={report.data} onChange={v => update({ data: v })} /></Field>
        <Field label="Turno"><Select value={report.turno} onChange={v => update({ turno: v })}>{TURNI.map(t => <option key={t}>{t}</option>)}</Select></Field>
        <Field label="Tipologia orario"><Select value={report.orarioTipo} onChange={v => update({ orarioTipo: v })}><option>Ordinario</option><option>Straordinario</option></Select></Field>
        <Field label="Reparto"><Select value={report.reparto} onChange={v => update({ reparto: v })}>{REPARTI.map(r => <option key={r}>{r}</option>)}</Select></Field>
      </div>
      {report.turno === 'Altro orario' && <div className="grid two"><Field label="Ora inizio"><Input value={report.altroTurnoInizio} onChange={v => update({ altroTurnoInizio: v })} placeholder="es. 10.00" /></Field><Field label="Ora fine"><Input value={report.altroTurnoFine} onChange={v => update({ altroTurnoFine: v })} placeholder="es. 17.00" /></Field></div>}
      {report.reparto === 'Altri servizi' && <Field label="Specificare altro servizio"><Input value={report.altroServizio} onChange={v => update({ altroServizio: v })} /></Field>}
    </section>

    <section className="card">
      <h2>2. Operatori</h2>
      {report.operatori.map((op, idx) => <div className="rowCard" key={idx}><div className="grid three"><Field label="Nome e cognome"><Input value={op.nome} onChange={v => updateArray('operatori', idx, { nome: v })} /></Field><Field label="Matricola"><Input value={op.matricola} onChange={v => updateArray('operatori', idx, { matricola: v })} /></Field><Field label="Qualifica"><Input value={op.qualifica} onChange={v => updateArray('operatori', idx, { qualifica: v })} /></Field></div><button className="ghost" onClick={() => removeArray('operatori', idx)}>Rimuovi</button></div>)}
      <button onClick={() => addArray('operatori', emptyOperatore())}>+ Aggiungi operatore</button>
    </section>

    <section className="card">
      <h2>3. Veicoli e chilometraggio</h2>
      {report.veicoli.map((v, idx) => <div className="rowCard" key={idx}><div className="grid four"><Field label="Veicolo / sigla"><Input value={v.sigla} onChange={x => updateArray('veicoli', idx, { sigla: x })} /></Field><Field label="Km inizio"><Input type="number" value={v.kmInizio} onChange={x => updateArray('veicoli', idx, { kmInizio: x })} /></Field><Field label="Km fine"><Input type="number" value={v.kmFine} onChange={x => updateArray('veicoli', idx, { kmFine: x })} /></Field><Field label="Km percorsi"><input readOnly value={km(v)} /></Field></div><button className="ghost" onClick={() => removeArray('veicoli', idx)}>Rimuovi</button></div>)}
      <button onClick={() => addArray('veicoli', emptyVeicolo())}>+ Aggiungi veicolo</button>
    </section>

    <section className="card">
      <h2>4. Interventi effettuati</h2>
      {report.interventi.map((i, idx) => <Intervento key={idx} i={i} idx={idx} updateIntervento={(patch) => updateArray('interventi', idx, patch)} remove={() => removeArray('interventi', idx)} />)}
      <button onClick={() => addArray('interventi', emptyIntervento())}>+ Aggiungi intervento</button>
    </section>

    <section className="card">
      <h2>5. Atti redatti</h2>
      <div className="counterGrid">
        {['relazioni','annotazioni','verbaliCds','verbaliRegolamenti','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr'].map(key => <Counter key={key} label={LABELS[key]} value={report.counters[key]} onChange={v => update({ counters: { ...report.counters, [key]: v } })} />)}
      </div>
      <div className="grid two"><Counter label="Altri atti" value={report.counters.altriAttiNumero} onChange={v => update({ counters: { ...report.counters, altriAttiNumero: v } })} /><Field label="Descrizione altri atti"><Input value={report.counters.altriAttiDescrizione} onChange={v => update({ counters: { ...report.counters, altriAttiDescrizione: v } })} /></Field></div>
    </section>

    <section className="card">
      <h2>6. Violazioni e provvedimenti</h2>
      <div className="counterGrid">
        {['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme','fermi','sequestri'].map(key => <Counter key={key} label={LABELS[key]} value={report.counters[key]} onChange={v => update({ counters: { ...report.counters, [key]: v } })} />)}
      </div>
      <Field label="Specificare altre norme"><Input value={report.counters.altreNormeDescrizione} onChange={v => update({ counters: { ...report.counters, altreNormeDescrizione: v } })} /></Field>
      <div className="totalBox">Totale violazioni / provvedimenti: <strong>{totaleViolazioni}</strong></div>
    </section>

    <section className="card">
      <h2>7. Note e invio</h2>
      <Field label="Note per UDT / Ufficiale di coordinamento"><Textarea value={report.noteUdt} onChange={v => update({ noteUdt: v })} /></Field>
      <Field label="Email ufficiale destinatario"><Input value={report.destinatario} onChange={v => update({ destinatario: v })} placeholder="es. ufficiale@comune.monza.it" /></Field>
      <label className="check"><input type="checkbox" checked={report.dichiarazione} onChange={e => update({ dichiarazione: e.target.checked })} /> Confermo la dichiarazione finale degli operatori.</label>
      <div className="actions"><button onClick={generatePdf}>Scarica PDF</button><button onClick={exportJson}>Scarica file dati JSON</button><button className="primary" onClick={sendMail}>Invia email precompilata</button></div>
    </section>

    <section className="card preview">
      <h2>Anteprima report</h2>
      <pre>{text}</pre>
    </section>
  </>;
}

function Intervento({ i, idx, updateIntervento, remove }) {
  const updateScuola = (sidx, patch) => updateIntervento({ scuole: i.scuole.map((s, n) => n === sidx ? { ...s, ...patch } : s) });
  const addScuola = () => { if (i.scuole.length < 3) updateIntervento({ scuole: [...i.scuole, emptyScuola()] }); };
  const removeScuola = (sidx) => updateIntervento({ scuole: i.scuole.filter((_, n) => n !== sidx) });
  return <div className="intervento"><div className="interventoHead"><h3>Intervento {idx + 1}</h3><button className="ghost" onClick={remove}>Rimuovi</button></div>
    <div className="grid three"><Field label="Tipo intervento"><Select value={i.tipo} onChange={v => updateIntervento({ tipo: v })}>{TIPI_INTERVENTO.map(t => <option key={t}>{t}</option>)}</Select></Field><Field label="Origine"><Select value={i.origine} onChange={v => updateIntervento({ origine: v })}>{ORIGINI.map(o => <option key={o}>{o}</option>)}</Select></Field><Field label="Luogo"><Input value={i.luogo} onChange={v => updateIntervento({ luogo: v })} /></Field></div>
    {i.origine === 'Altro' && <Field label="Specificare da chi è arrivata la disposizione"><Input value={i.origineAltro} onChange={v => updateIntervento({ origineAltro: v })} /></Field>}
    <div className="grid two"><Field label="Ora inizio"><Input value={i.oraInizio} onChange={v => updateIntervento({ oraInizio: v })} placeholder="es. 08.15" /></Field><Field label="Ora fine"><Input value={i.oraFine} onChange={v => updateIntervento({ oraFine: v })} placeholder="es. 09.00" /></Field></div>
    <Field label="Descrizione"><Textarea value={i.descrizione} onChange={v => updateIntervento({ descrizione: v })} /></Field>
    {i.tipo === 'Sinistro stradale' && <div className="grid three"><Field label="Feriti"><Select value={i.conFeriti} onChange={v => updateIntervento({ conFeriti: v })}><option>Con feriti</option><option>Senza feriti</option></Select></Field><Field label="Veicoli coinvolti"><Input type="number" value={i.veicoliCoinvolti} onChange={v => updateIntervento({ veicoliCoinvolti: v })} /></Field><Field label="Rilievi effettuati"><Select value={i.rilievi} onChange={v => updateIntervento({ rilievi: v })}><option>Sì</option><option>No</option></Select></Field></div>}
    {i.tipo === 'Posto di controllo' && <div className="grid four"><Field label="Veicoli controllati"><Input type="number" value={i.veicoliControllati} onChange={v => updateIntervento({ veicoliControllati: v })} /></Field><Field label="Persone controllate"><Input type="number" value={i.personeControllate} onChange={v => updateIntervento({ personeControllate: v })} /></Field><Field label="Verbali elevati"><Input type="number" value={i.verbaliElevati} onChange={v => updateIntervento({ verbaliElevati: v })} /></Field><Field label="Fermi / sequestri"><Input type="number" value={i.fermiSequestri} onChange={v => updateIntervento({ fermiSequestri: v })} /></Field></div>}
    {i.tipo === 'Viabilità' && <div className="grid two"><Field label="Motivo viabilità"><Input value={i.motivoViabilita} onChange={v => updateIntervento({ motivoViabilita: v })} placeholder="incidente, cantiere, evento..." /></Field><Field label="Strade interessate"><Input value={i.strade} onChange={v => updateIntervento({ strade: v })} /></Field></div>}
    {i.tipo === 'Servizio scuole' && <div className="schoolBox"><h4>Scuole presidiate</h4>{i.scuole.map((s, sidx) => <div className="rowCard" key={sidx}><div className="grid four"><Field label={`Scuola ${sidx + 1}`}><Input value={s.nome} onChange={v => updateScuola(sidx, { nome: v })} /></Field><Field label="Ingresso / uscita"><Select value={s.momento} onChange={v => updateScuola(sidx, { momento: v })}><option value="">Seleziona</option><option>Ingresso</option><option>Uscita</option><option>Ingresso e uscita</option></Select></Field><Field label="Orario"><Input value={s.orario} onChange={v => updateScuola(sidx, { orario: v })} /></Field><Field label="Criticità"><Input value={s.criticita} onChange={v => updateScuola(sidx, { criticita: v })} /></Field></div>{i.scuole.length > 1 && <button className="ghost" onClick={() => removeScuola(sidx)}>Rimuovi scuola</button>}</div>)}{i.scuole.length < 3 && <button onClick={addScuola}>+ Aggiungi scuola</button>}</div>}
    <div className="grid two"><Field label="Esito"><Input value={i.esito} onChange={v => updateIntervento({ esito: v })} /></Field><Field label="Note"><Input value={i.note} onChange={v => updateIntervento({ note: v })} /></Field></div>
  </div>;
}

function Dashboard({ reports, setReports }) {
  const [commanderNotes, setCommanderNotes] = useState('');
  const aggregate = useMemo(() => aggregateReports(reports), [reports]);
  const text = useMemo(() => commanderReportText(aggregate, reports, commanderNotes), [aggregate, reports, commanderNotes]);

  async function importFiles(e) {
    const files = Array.from(e.target.files || []);
    const parsed = [];
    for (const file of files) {
      try {
        const raw = await file.text();
        const data = JSON.parse(raw);
        if (data && data.interventi && data.counters) parsed.push(data);
      } catch (err) {
        alert(`File non leggibile: ${file.name}`);
      }
    }
    setReports(prev => [...prev, ...parsed]);
    e.target.value = '';
  }

  function generateCommanderPdf() {
    const doc = buildCommanderPdf(aggregate, reports, commanderNotes);
    doc.save(`report-aggregato-comandante-${sanitizeFileName(aggregate.dateLabel)}.pdf`);
  }

  function exportCsv() {
    const rows = [['Data','Turno','Orario','Reparto','Operatori','Interventi','Violazioni/Provvedimenti','Km','Note']];
    reports.forEach(r => rows.push([r.data, turnoLabel(r), r.orarioTipo, repartoLabel(r), operatorNames(r).join('; '), r.interventi?.length || 0, getTotaleViolazioni(r), getKmTotali(r), r.noteUdt || '']));
    const csv = rows.map(row => row.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(';')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `tabella-report-${sanitizeFileName(aggregate.dateLabel)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return <>
    <section className="card notice">
      <h2>Dashboard ufficiale</h2>
      <p>Carica i file <strong>JSON</strong> ricevuti dagli operatori. La dashboard aggrega automaticamente interventi, reparti, origini, violazioni, atti, chilometri e note rilevanti.</p>
      <div className="actions"><label className="fileButton">Carica file JSON<input type="file" accept="application/json,.json" multiple onChange={importFiles} /></label><button className="ghost" onClick={() => setReports([])}>Svuota dashboard</button></div>
    </section>

    <section className="metrics">
      <Metric label="Report caricati" value={reports.length} />
      <Metric label="Interventi totali" value={aggregate.totalInterventi} />
      <Metric label="Violazioni / provvedimenti" value={aggregate.totaleViolazioni} />
      <Metric label="Km percorsi" value={aggregate.kmTotali} />
    </section>

    <section className="card">
      <h2>Quadro riepilogativo</h2>
      {reports.length === 0 ? <p className="muted">Nessun report caricato.</p> : <div className="tableWrap"><table><thead><tr><th>Data</th><th>Turno</th><th>Reparto</th><th>Operatori</th><th>Interventi</th><th>Violazioni</th><th>Km</th></tr></thead><tbody>{reports.map((r, idx) => <tr key={idx}><td>{r.data}</td><td>{turnoLabel(r)}<br/><small>{r.orarioTipo}</small></td><td>{repartoLabel(r)}</td><td>{operatorNames(r).join(', ') || '-'}</td><td>{r.interventi?.length || 0}</td><td>{getTotaleViolazioni(r)}</td><td>{getKmTotali(r)}</td></tr>)}</tbody></table></div>}
    </section>

    <section className="card split">
      <Distribution title="Interventi per tipologia" data={aggregate.byTipo} />
      <Distribution title="Origine interventi" data={aggregate.byOrigine} />
      <Distribution title="Reparti" data={aggregate.byReparto} />
      <Distribution title="Atti e violazioni" data={aggregate.counters} labels={LABELS} />
    </section>

    <section className="card">
      <h2>Note rilevanti e criticità</h2>
      {aggregate.notes.length === 0 ? <p className="muted">Nessuna nota inserita nei report caricati.</p> : <ul className="noteList">{aggregate.notes.map((note, idx) => <li key={idx}><strong>{note.data} - {note.turno} - {note.reparto}</strong><br/>{note.testo}</li>)}</ul>}
      <Field label="Note dell'ufficiale per il Comandante"><Textarea value={commanderNotes} onChange={setCommanderNotes} placeholder="Inserire valutazioni, criticità da attenzionare, esigenze operative, proposte..." /></Field>
      <div className="actions"><button onClick={generateCommanderPdf}>Genera PDF aggregato</button><button onClick={() => exportExcelAvanzato(reports, aggregate, commanderNotes)}>Esporta Excel avanzato</button><button onClick={exportCsv}>Esporta tabella CSV</button></div>
    </section>

    <section className="card preview">
      <h2>Anteprima report aggregato</h2>
      <pre>{text}</pre>
    </section>
  </>;
}

function OfficialReport({ reports, setReports, official, setOfficial }) {
  const aggregate = useMemo(() => aggregateReports(reports), [reports]);
  const autoSintesi = useMemo(() => officialSynthesis(aggregate, reports), [aggregate, reports]);
  const autoEventi = useMemo(() => officialEventsText(reports), [reports]);
  const preview = useMemo(() => officialReportText(aggregate, reports, official, autoSintesi, autoEventi), [aggregate, reports, official, autoSintesi, autoEventi]);
  const update = (patch) => setOfficial(prev => ({ ...prev, ...patch }));
  const updateAttivita = (idx, patch) => setOfficial(prev => ({ ...prev, attivitaIspettive: prev.attivitaIspettive.map((x, i) => i === idx ? { ...x, ...patch } : x) }));
  const addAttivita = () => setOfficial(prev => ({ ...prev, attivitaIspettive: [...prev.attivitaIspettive, emptyAttivitaIspettiva()] }));
  const removeAttivita = (idx) => setOfficial(prev => ({ ...prev, attivitaIspettive: prev.attivitaIspettive.filter((_, i) => i !== idx) }));

  async function importFiles(e) {
    const files = Array.from(e.target.files || []);
    const parsed = [];
    for (const file of files) {
      try {
        const raw = await file.text();
        const data = JSON.parse(raw);
        if (data && data.interventi && data.counters) parsed.push(data);
      } catch (err) {
        alert(`File non leggibile: ${file.name}`);
      }
    }
    setReports(prev => [...prev, ...parsed]);
    e.target.value = '';
  }

  function generateOfficialPdf() {
    const doc = buildOfficialShiftPdf(aggregate, reports, official, autoSintesi, autoEventi);
    doc.save(`report-ufficiale-turno-${sanitizeFileName(official.data || aggregate.dateLabel)}-${sanitizeFileName(official.turno)}.pdf`);
  }

  return <>
    <section className="card notice">
      <h2>Report ufficiale UDT</h2>
      <p>Modalità quasi automatica: carica i JSON degli operatori, verifica la sintesi generata, integra briefing, personale, attività ispettive, anomalie e note per il Comandante.</p>
      <div className="actions"><label className="fileButton">Carica JSON operatori<input type="file" accept="application/json,.json" multiple onChange={importFiles} /></label><button className="ghost" onClick={() => setReports([])}>Svuota dati caricati</button></div>
    </section>
    <section className="metrics"><Metric label="Report operatori" value={reports.length} /><Metric label="Interventi" value={aggregate.totalInterventi} /><Metric label="Violazioni / provvedimenti" value={aggregate.totaleViolazioni} /><Metric label="Km" value={aggregate.kmTotali} /></section>
    <section className="card"><h2>1. Dati report ufficiale</h2><div className="grid four"><Field label="Data"><Input type="date" value={official.data} onChange={v => update({ data: v })} /></Field><Field label="Turno"><Input value={official.turno} onChange={v => update({ turno: v })} placeholder="es. 1° turno" /></Field><Field label="Ufficiale di turno"><Input value={official.ufficiale} onChange={v => update({ ufficiale: v })} /></Field><Field label="Qualifica"><Input value={official.qualifica} onChange={v => update({ qualifica: v })} placeholder="es. Commissario Capo" /></Field></div></section>
    <section className="card"><h2>2. Sintesi automatica</h2><p className="muted">Questa sintesi nasce dai report operatori caricati. Nel PDF viene riportata come quadro iniziale.</p><pre className="miniPreview">{autoSintesi}</pre><Field label="Integrazioni dell'ufficiale alla sintesi"><Textarea value={official.eventiManuali} onChange={v => update({ eventiManuali: v })} placeholder="Inserire eventuali elementi aggiuntivi non presenti nei report operatori..." /></Field></section>
    <section className="card"><h2>3. Briefing, personale e note</h2><div className="grid two"><Field label="Briefing operativo"><Input value={official.briefing} onChange={v => update({ briefing: v })} placeholder="es. 06.45" /></Field><Field label="Note generali"><Input value={official.noteGenerali} onChange={v => update({ noteGenerali: v })} placeholder="es. Con il personale a disposizione coperte 11 scuole" /></Field></div><div className="grid two"><Field label="A.P.L. assenti"><Textarea value={official.assenti} onChange={v => update({ assenti: v })} /></Field><Field label="A.P.L. in ritardo"><Textarea value={official.ritardi} onChange={v => update({ ritardi: v })} /></Field></div></section>
    <section className="card"><h2>4. Eventi degni di rilievo</h2><p className="muted">Eventi rilevanti individuati automaticamente: sinistri con feriti, TSO/ASO, interventi con parole chiave critiche o lunga durata.</p><pre className="miniPreview">{autoEventi || 'Nessun evento rilevante automatico rilevato.'}</pre></section>
    <section className="card"><h2>5. Anomalie e attività ispettive</h2><Field label="Anomalie riscontrate durante il turno"><Textarea value={official.anomalie} onChange={v => update({ anomalie: v })} /></Field><h3>Attività ispettive</h3>{official.attivitaIspettive.map((a, idx) => <div className="rowCard" key={idx}><div className="grid four"><Field label="Tipo attività"><Input value={a.tipo} onChange={v => updateAttivita(idx, { tipo: v })} placeholder="es. annonaria, ambiente..." /></Field><Field label="Reparto / pattuglia"><Input value={a.reparto} onChange={v => updateAttivita(idx, { reparto: v })} /></Field><Field label="Luogo"><Input value={a.luogo} onChange={v => updateAttivita(idx, { luogo: v })} /></Field><Field label="Orario"><Input value={a.orario} onChange={v => updateAttivita(idx, { orario: v })} /></Field></div><div className="grid three"><Field label="Esito"><Input value={a.esito} onChange={v => updateAttivita(idx, { esito: v })} /></Field><Field label="Violazioni collegate"><Input value={a.violazioni} onChange={v => updateAttivita(idx, { violazioni: v })} /></Field><Field label="Note"><Input value={a.note} onChange={v => updateAttivita(idx, { note: v })} /></Field></div><button className="ghost" onClick={() => removeAttivita(idx)}>Rimuovi attività</button></div>)}<button onClick={addAttivita}>+ Aggiungi attività ispettiva</button></section>
    <section className="card"><h2>6. Esiti e comunicazioni</h2><Field label="Esiti"><Textarea value={official.esiti} onChange={v => update({ esiti: v })} /></Field><div className="grid two"><Field label="Comunicazione all'E.Q. di turno"><Textarea value={official.comunicazioneEq} onChange={v => update({ comunicazioneEq: v })} /></Field><Field label="Nota per il Comandante"><Textarea value={official.notaComandante} onChange={v => update({ notaComandante: v })} /></Field></div><div className="actions"><button className="primary" onClick={generateOfficialPdf}>Genera PDF Report Ufficiale</button></div></section>
    <section className="card preview"><h2>Anteprima report ufficiale</h2><pre>{preview}</pre></section>
  </>;
}

function Metric({ label, value }) { return <div className="metric"><strong>{value}</strong><span>{label}</span></div>; }
function Distribution({ title, data, labels = {} }) {
  const entries = Object.entries(data || {}).filter(([, value]) => n(value) > 0).sort((a, b) => n(b[1]) - n(a[1]));
  return <div><h3>{title}</h3>{entries.length === 0 ? <p className="muted">Nessun dato.</p> : <ul className="distList">{entries.map(([key, value]) => <li key={key}><span>{labels[key] || key}</span><strong>{value}</strong></li>)}</ul>}</div>;
}

function makePdf(title, subtitle = '') {
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  doc.setProperties({ title, subject: 'Report Polizia Locale', author: 'Polizia Locale' });
  addHeader(doc, title, subtitle);
  return doc;
}

function addHeader(doc, title, subtitle = '') {
  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, 210, 34, 'F');
  try {
    const img = document.getElementById('pdfLogo');
    if (img && img.complete) doc.addImage(img, 'PNG', 12, 6, 22, 22);
  } catch (e) {}
  doc.setDrawColor(12, 47, 97);
  doc.setLineWidth(0.35);
  doc.line(40, 7, 40, 28);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(12, 47, 97);
  doc.text('COMUNE DI MONZA', 45, 13);
  doc.setFontSize(10.5);
  doc.text('Settore Polizia Locale, Protezione Civile', 45, 20);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.8);
  doc.setTextColor(40, 48, 60);
  doc.text('Via Marsala 13 | 20900 Monza', 145, 11);
  doc.text('Tel. 039.2816313', 145, 17);
  doc.text('polizialocale@comune.monza.it', 145, 23);
  doc.setDrawColor(12, 47, 97);
  doc.setLineWidth(0.6);
  doc.line(12, 32, 198, 32);
  doc.setFillColor(12, 47, 97);
  doc.roundedRect(12, 38, 126, 10, 1.8, 1.8, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text(title, 75, 44.5, { align: 'center' });
  doc.setDrawColor(180, 196, 224);
  doc.setLineWidth(1.2);
  doc.line(141, 39, 136, 47);
  doc.line(145, 39, 140, 47);
  doc.line(149, 39, 144, 47);
  if (subtitle) {
    doc.setTextColor(55, 65, 81);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text(subtitle, 198, 44.5, { align: 'right' });
  }
  doc.setTextColor(0, 0, 0);
}

function addFooter(doc) {
  const pages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pages; i++) {
    doc.setPage(i);
    doc.setDrawColor(12, 47, 97);
    doc.setLineWidth(0.35);
    doc.line(12, 284, 198, 284);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.4);
    doc.setTextColor(12, 47, 97);
    doc.text('Settore Polizia Locale, Protezione Civile', 12, 289);
    doc.text('Via Marsala 13 | 20900 Monza', 12, 293);
    doc.text('Tel. 039.2816313', 86, 291);
    doc.text('polizialocale@comune.monza.it', 127, 291);
    doc.setFillColor(12, 47, 97);
    doc.roundedRect(178, 287, 20, 7, 1.2, 1.2, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.text(`Pag. ${i} di ${pages}`, 188, 291.6, { align: 'center' });
    doc.setTextColor(0, 0, 0);
  }
}

function ensureSpace(doc, y, needed = 18, title = '', subtitle = '') {
  if (y + needed <= 276) return y;
  doc.addPage();
  addHeader(doc, title, subtitle);
  return 54;
}

function section(doc, label, y, pdfTitle = '', subtitle = '') {
  y = ensureSpace(doc, y, 13, pdfTitle, subtitle);
  doc.setFillColor(245, 248, 252);
  doc.setDrawColor(214, 222, 232);
  doc.roundedRect(12, y, 186, 9, 1.4, 1.4, 'FD');
  doc.setFillColor(12, 47, 97);
  doc.rect(12, y, 3.2, 9, 'F');
  doc.setTextColor(12, 47, 97);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.text(label.toUpperCase(), 18, y + 6);
  doc.setTextColor(0, 0, 0);
  return y + 12;
}

function box(doc, y, h, pdfTitle = '', subtitle = '') {
  y = ensureSpace(doc, y, h + 4, pdfTitle, subtitle);
  doc.setDrawColor(222, 226, 230);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(12, y, 186, h, 1.5, 1.5, 'FD');
  return y;
}

function kvGrid(doc, items, y, columns = 2, pdfTitle = '', subtitle = '') {
  const rowH = 11;
  const rows = Math.ceil(items.length / columns);
  y = box(doc, y, rows * rowH + 4, pdfTitle, subtitle) + 7;
  const colW = 186 / columns;
  items.forEach((item, idx) => {
    const col = idx % columns;
    const row = Math.floor(idx / columns);
    const x = 16 + col * colW;
    const yy = y + row * rowH;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(80, 86, 94);
    doc.text(item.label, x, yy);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(0, 0, 0);
    doc.text(String(item.value || '-'), x, yy + 4.5, { maxWidth: colW - 9 });
  });
  return y + rows * rowH + 2;
}

function simpleTable(doc, headers, rows, y, widths, pdfTitle = '', subtitle = '') {
  const headerH = 8;
  const lineH = 5;
  y = ensureSpace(doc, y, headerH + 8, pdfTitle, subtitle);
  let x = 12;
  doc.setFillColor(235, 239, 244);
  doc.rect(12, y, 186, headerH, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7.8);
  headers.forEach((h, i) => { doc.text(h, x + 2, y + 5.2, { maxWidth: widths[i] - 4 }); x += widths[i]; });
  y += headerH;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.6);
  rows.forEach(row => {
    const cellLines = row.map((cell, i) => doc.splitTextToSize(String(cell || '-'), widths[i] - 4));
    const h = Math.max(7, Math.max(...cellLines.map(lines => lines.length)) * lineH + 2);
    y = ensureSpace(doc, y, h + 2, pdfTitle, subtitle);
    doc.setDrawColor(232, 235, 238);
    doc.rect(12, y, 186, h);
    let xx = 12;
    cellLines.forEach((lines, i) => {
      doc.text(lines, xx + 2, y + 5, { maxWidth: widths[i] - 4 });
      xx += widths[i];
      if (i < widths.length - 1) doc.line(xx, y, xx, y + h);
    });
    y += h;
  });
  return y + 4;
}

function paragraph(doc, text, y, pdfTitle = '', subtitle = '', maxWidth = 178) {
  const lines = doc.splitTextToSize(String(text || '-'), maxWidth);
  y = ensureSpace(doc, y, lines.length * 5 + 8, pdfTitle, subtitle);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.text(lines, 16, y + 5);
  return y + lines.length * 5 + 8;
}


function serviceSummaryBox(doc, report, y, pdfTitle = '', subtitle = '') {
  y = ensureSpace(doc, y, 30, pdfTitle, subtitle);
  const interventi = (report.interventi || []).length;
  const violazioni = getTotaleViolazioni(report);
  const atti = ['relazioni','annotazioni','verbaliCds','verbaliRegolamenti','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].reduce((s, k) => s + n((report.counters || {})[k]), 0);
  const criticita = (report.interventi || []).filter(isInterventoCritico).length + (testoCritico(report.noteUdt) ? 1 : 0);
  doc.setFillColor(247, 250, 252);
  doc.setDrawColor(214, 222, 232);
  doc.roundedRect(12, y, 186, 29, 1.8, 1.8, 'FD');
  const items = [
    ['Interventi', interventi],
    ['Violazioni / provv.', violazioni],
    ['Atti redatti', atti],
    ['Criticità', criticita],
  ];
  items.forEach((item, idx) => {
    const x = 20 + idx * 44;
    if (idx > 0) {
      doc.setDrawColor(225, 231, 239);
      doc.line(x - 8, y + 6, x - 8, y + 23);
    }
    doc.setTextColor(12, 47, 97);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(String(item[1]), x, y + 13);
    doc.setFontSize(7.6);
    doc.setTextColor(75, 85, 99);
    doc.text(item[0].toUpperCase(), x, y + 21);
  });
  doc.setTextColor(0, 0, 0);
  return y + 35;
}

function serviceInterventionCard(doc, i, idx, y, pdfTitle = '', subtitle = '') {
  const critic = isInterventoCritico(i);
  y = ensureSpace(doc, y, 30, pdfTitle, subtitle);
  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(214, 222, 232);
  doc.roundedRect(12, y, 186, 0.1, 1, 1, 'S');
  const startY = y;
  doc.setFillColor(critic ? 254 : 245, critic ? 242 : 248, critic ? 242 : 252);
  doc.roundedRect(12, y, 186, 10, 1.4, 1.4, 'F');
  doc.setFillColor(12, 47, 97);
  doc.roundedRect(15, y + 2.1, 18, 5.8, 1.2, 1.2, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7.5);
  doc.text(`${i.oraInizio || '--'}-${i.oraFine || '--'}`, 24, y + 6.1, { align: 'center' });
  doc.setTextColor(12, 47, 97);
  doc.setFontSize(9.2);
  doc.text(`${idx + 1}. ${i.tipo || 'Intervento'}`, 38, y + 6.3);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(65, 75, 90);
  const origine = i.origine === 'Altro' ? `Altro: ${i.origineAltro || '-'}` : (i.origine || '-');
  doc.text(`Origine: ${origine} | Luogo: ${i.luogo || '-'}`, 98, y + 6.3, { maxWidth: 96 });
  y += 13;
  const scuole = i.tipo === 'Servizio scuole' ? (i.scuole || []).filter(s => s.nome || s.momento || s.orario || s.criticita).map((s, pos) => `Scuola ${pos + 1}: ${s.nome || '-'} (${s.momento || '-'} ${s.orario || '-'}) Criticità: ${s.criticita || '-'}`).join('\n') : '';
  const dettagli = extraDetails(i).replace(/\n/g, ' ').trim();
  const body = `Descrizione: ${i.descrizione || '-'}\nEsito: ${i.esito || '-'}${dettagli ? '\n' + dettagli : ''}${scuole ? '\n' + scuole : ''}\nNote: ${i.note || '-'}`;
  const lines = doc.splitTextToSize(body, 176);
  const h = Math.max(14, lines.length * 4.4 + 7);
  y = ensureSpace(doc, startY, 13 + h, pdfTitle, subtitle) + 13;
  doc.setTextColor(15, 23, 42);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.2);
  doc.text(lines, 18, y + 3.5);
  return y + h + 4;
}

function buildServicePdf(report) {
  const title = 'REPORT DI SERVIZIO';
  const subtitle = `${report.data} | Turno ${turnoLabel(report)} | ${report.orarioTipo}`;
  const doc = makePdf(title, subtitle);
  let y = 54;

  y = serviceSummaryBox(doc, report, y, title, subtitle);

  y = section(doc, 'Dati generali del turno', y, title, subtitle);
  y = kvGrid(doc, [
    { label: 'Data servizio', value: report.data },
    { label: 'Turno', value: turnoLabel(report) },
    { label: 'Tipologia orario', value: report.orarioTipo },
    { label: 'Reparto / servizio', value: repartoLabel(report) },
  ], y, 2, title, subtitle) + 2;

  y = section(doc, 'Operatori', y, title, subtitle);
  const operatorRows = (report.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).map(o => [o.nome, o.matricola, o.qualifica]);
  y = simpleTable(doc, ['Nominativo', 'Matricola', 'Qualifica'], operatorRows.length ? operatorRows : [['-', '-', '-']], y, [90, 40, 56], title, subtitle);

  y = section(doc, 'Veicoli e chilometraggio', y, title, subtitle);
  const vehicleRows = (report.veicoli || []).map(v => [v.sigla || '-', v.kmInizio || '-', v.kmFine || '-', km(v)]);
  y = simpleTable(doc, ['Veicolo', 'Km inizio', 'Km fine', 'Km percorsi'], vehicleRows.length ? vehicleRows : [['-', '-', '-', '-']], y, [72, 38, 38, 38], title, subtitle);
  y = kvGrid(doc, [{ label: 'Totale km percorsi', value: getKmTotali(report) }], y, 1, title, subtitle) + 2;

  y = section(doc, 'Interventi effettuati', y, title, subtitle);
  (report.interventi || []).forEach((i, idx) => {
    y = serviceInterventionCard(doc, i, idx, y, title, subtitle);
  });

  y = section(doc, 'Atti redatti', y, title, subtitle);
  const c = report.counters || emptyCounters();
  y = simpleTable(doc, ['Tipologia', 'N.'], [
    ['Relazioni di servizio', c.relazioni], ['Annotazioni di servizio', c.annotazioni], ['Verbali CdS', c.verbaliCds], ['Verbali regolamenti', c.verbaliRegolamenti],
    ['Sequestri amministrativi', c.sequestriAmministrativi], ['Fermi amministrativi', c.fermiAmministrativi], ['Sequestri penali', c.sequestriPenali], ['C.N.R.', c.cnr], [`Altri atti ${c.altriAttiDescrizione || ''}`, c.altriAttiNumero]
  ], y, [150, 36], title, subtitle);

  y = section(doc, 'Violazioni e provvedimenti', y, title, subtitle);
  y = simpleTable(doc, ['Tipologia', 'N.'], [
    ['Preavvisi CdS', c.preavvisiCds], ['VdC CdS', c.vdcCds], ['Regolamento Polizia', c.regPolizia], ['Regolamento Edilizio', c.regEdilizio],
    ['Regolamento Benessere Animali', c.regBenessereAnimali], ['Annonaria / commercio', c.annonaria], [`Altre norme ${c.altreNormeDescrizione || ''}`, c.altreNorme], ['Fermi', c.fermi], ['Sequestri', c.sequestri], ['TOTALE', getTotaleViolazioni(report)]
  ], y, [150, 36], title, subtitle);

  y = section(doc, 'Note per UDT / Ufficiale di coordinamento', y, title, subtitle);
  y = paragraph(doc, report.noteUdt || '-', y, title, subtitle);

  y = section(doc, 'Dichiarazione e firme', y, title, subtitle);
  y = paragraph(doc, 'Gli operatori dichiarano che quanto riportato nel presente report corrisponde fedelmente alle attività effettivamente svolte e riscontrate durante il turno di servizio, consapevoli delle proprie responsabilità amministrative e penali anche in considerazione dell’art. 328 C.P.', y, title, subtitle);
  y = ensureSpace(doc, y, 24, title, subtitle);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  const names = operatorNames(report);
  const signRows = names.length ? names : ['Operatore 1', 'Operatore 2', 'Operatore 3'];
  signRows.slice(0, 6).forEach((name, idx) => {
    const yy = y + idx * 9;
    doc.text(name, 16, yy);
    doc.line(82, yy + 1, 190, yy + 1);
  });

  addFooter(doc);
  return doc;
}

function buildCommanderPdf(aggregate, reports, commanderNotes) {
  const title = 'REPORT AGGREGATO PER IL COMANDANTE';
  const subtitle = `Periodo/Data: ${aggregate.dateLabel}`;
  const doc = makePdf(title, subtitle);
  let y = 54;

  y = section(doc, 'Sintesi operativa', y, title, subtitle);
  y = kvGrid(doc, [
    { label: 'Report ricevuti', value: reports.length },
    { label: 'Interventi totali', value: aggregate.totalInterventi },
    { label: 'Violazioni / provvedimenti', value: aggregate.totaleViolazioni },
    { label: 'Km totali percorsi', value: aggregate.kmTotali },
  ], y, 4, title, subtitle) + 2;

  y = section(doc, 'Interventi per tipologia', y, title, subtitle);
  y = simpleTable(doc, ['Tipologia', 'Totale'], objectRows(aggregate.byTipo), y, [150, 36], title, subtitle);

  y = section(doc, 'Origine degli interventi', y, title, subtitle);
  y = simpleTable(doc, ['Origine', 'Totale'], objectRows(aggregate.byOrigine), y, [150, 36], title, subtitle);

  y = section(doc, 'Reparti / servizi rendicontati', y, title, subtitle);
  y = simpleTable(doc, ['Reparto / servizio', 'Report'], objectRows(aggregate.byReparto), y, [150, 36], title, subtitle);

  y = section(doc, 'Atti, violazioni e provvedimenti', y, title, subtitle);
  y = simpleTable(doc, ['Voce', 'Totale'], objectRows(aggregate.counters, LABELS), y, [150, 36], title, subtitle);

  y = section(doc, 'Dettaglio report ricevuti', y, title, subtitle);
  const detailRows = reports.map(r => [r.data, turnoLabel(r), r.orarioTipo, repartoLabel(r), operatorNames(r).join(', ') || '-', (r.interventi || []).length, getTotaleViolazioni(r), getKmTotali(r)]);
  y = simpleTable(doc, ['Data', 'Turno', 'Tipo', 'Reparto', 'Operatori', 'Int.', 'Viol.', 'Km'], detailRows.length ? detailRows : [['-', '-', '-', '-', '-', '-', '-', '-']], y, [22, 27, 23, 36, 43, 11, 12, 12], title, subtitle);

  y = section(doc, 'Note rilevanti / criticità', y, title, subtitle);
  const notes = aggregate.notes.length ? aggregate.notes.map(n => `${n.data} | ${n.turno} | ${n.reparto}: ${n.testo}`).join('\n\n') : 'Nessuna nota rilevante.';
  y = paragraph(doc, notes, y, title, subtitle);

  y = section(doc, 'Note dell\'ufficiale per il Comandante', y, title, subtitle);
  y = paragraph(doc, commanderNotes || '-', y, title, subtitle);

  addFooter(doc);
  return doc;
}

function objectRows(obj, labels = {}) {
  const rows = Object.entries(obj || {}).filter(([, value]) => n(value) > 0).sort((a, b) => n(b[1]) - n(a[1])).map(([key, value]) => [labels[key] || key, value]);
  return rows.length ? rows : [['Nessun dato', '-']];
}


function parseTimeToMinutes(value) {
  if (!value) return null;
  const clean = String(value).trim().replace('.', ':');
  const match = clean.match(/^(\d{1,2})(?::(\d{1,2}))?$/);
  if (!match) return null;
  const h = Number(match[1]);
  const m = Number(match[2] || 0);
  if (Number.isNaN(h) || Number.isNaN(m)) return null;
  return (h * 60 + m) % 1440;
}
function durataMinuti(oraInizio, oraFine) {
  const start = parseTimeToMinutes(oraInizio);
  const end = parseTimeToMinutes(oraFine);
  if (start === null || end === null) return '';
  return end >= start ? end - start : (1440 - start) + end;
}
function fasciaOraria(oraInizio) {
  const minutes = parseTimeToMinutes(oraInizio);
  if (minutes === null) return 'Non indicata';
  const h = Math.floor(minutes / 60);
  if (h >= 6 && h < 12) return 'Mattino';
  if (h >= 12 && h < 18) return 'Pomeriggio';
  if (h >= 18 && h < 22) return 'Sera';
  return 'Notte';
}
function testoCritico(text = '') {
  const value = String(text).toLowerCase();
  return ['critic', 'pericol', 'ferit', 'aggression', 'grave', 'emergenza', 'rischio', 'tso', 'aso', 'sinistro'].some(k => value.includes(k));
}
function isInterventoCritico(intervento) {
  return intervento?.conFeriti === 'Con feriti' || ['TSO', 'ASO'].includes(intervento?.tipo) || testoCritico(`${intervento?.tipo || ''} ${intervento?.descrizione || ''} ${intervento?.esito || ''} ${intervento?.note || ''}`);
}
function addCount(target, key, amount = 1) {
  const label = key || 'Non indicato';
  target[label] = n(target[label]) + n(amount || 0);
}
function orderedObjectRows(obj, labels = {}) {
  return Object.entries(obj || {}).filter(([, value]) => n(value) > 0).sort((a, b) => n(b[1]) - n(a[1])).map(([key, value]) => ({ Voce: labels[key] || key, Totale: n(value) }));
}
function makeBar(value, max) {
  const blocks = Math.round((n(value) / Math.max(1, n(max))) * 20);
  return '█'.repeat(blocks);
}
function buildExcelData(reports, aggregate, commanderNotes) {
  const interventiRows = [];
  const reportRows = [];
  const countersRows = [];
  const byFascia = { Mattino: 0, Pomeriggio: 0, Sera: 0, Notte: 0, 'Non indicata': 0 };
  let interventoPiuLungo = { durata: -1, label: '-' };
  const criticitaRows = [];
  reports.forEach((r, reportIndex) => {
    const operatori = operatorNames(r).join('; ');
    reportRows.push({ Data: r.data || '', Turno: turnoLabel(r), 'Tipologia orario': r.orarioTipo || '', 'Reparto / servizio': repartoLabel(r), Operatori: operatori, 'Numero operatori': (r.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).length, 'Interventi totali': (r.interventi || []).length, 'Violazioni / provvedimenti': getTotaleViolazioni(r), 'Km percorsi': getKmTotali(r), 'Note UDT': r.noteUdt || '' });
    Object.entries(r.counters || {}).forEach(([key, value]) => { if (typeof value === 'number' && n(value) > 0) countersRows.push({ Data: r.data || '', Turno: turnoLabel(r), Reparto: repartoLabel(r), Voce: LABELS[key] || key, Totale: n(value) }); });
    (r.interventi || []).forEach((i, idx) => {
      const durata = durataMinuti(i.oraInizio, i.oraFine);
      const fascia = fasciaOraria(i.oraInizio);
      addCount(byFascia, fascia);
      const origine = i.origine === 'Altro' ? `Altro: ${i.origineAltro || '-'}` : (i.origine || 'Non indicata');
      const scuole = i.tipo === 'Servizio scuole' ? (i.scuole || []).filter(s => s.nome || s.momento || s.orario || s.criticita).map((s, pos) => `Scuola ${pos + 1}: ${s.nome || '-'} (${s.momento || '-'} ${s.orario || '-'}) Criticità: ${s.criticita || '-'}`).join(' | ') : '';
      const criticita = isInterventoCritico(i) ? 'Sì' : 'No';
      if (durata !== '' && durata > interventoPiuLungo.durata) interventoPiuLungo = { durata, label: `${r.data || ''} ${turnoLabel(r)} - ${i.tipo || '-'} (${durata} min)` };
      const row = { Data: r.data || '', Turno: turnoLabel(r), 'Tipologia orario': r.orarioTipo || '', 'Reparto / servizio': repartoLabel(r), Operatori: operatori, 'N. report': reportIndex + 1, 'N. intervento': idx + 1, 'Tipo intervento': i.tipo || '', Origine: origine, 'Ora inizio': i.oraInizio || '', 'Ora fine': i.oraFine || '', 'Durata minuti': durata, 'Fascia oraria': fascia, Luogo: i.luogo || '', Descrizione: i.descrizione || '', Esito: i.esito || '', Note: i.note || '', Criticità: criticita, Feriti: i.conFeriti || '', 'Veicoli coinvolti': i.veicoliCoinvolti || '', 'Rilievi effettuati': i.rilievi || '', 'Veicoli controllati': i.veicoliControllati || '', 'Persone controllate': i.personeControllate || '', 'Verbali elevati': i.verbaliElevati || '', 'Fermi / sequestri intervento': i.fermiSequestri || '', 'Motivo viabilità': i.motivoViabilita || '', 'Strade interessate': i.strade || '', 'Scuole presidiate': scuole };
      interventiRows.push(row);
      if (criticita === 'Sì') criticitaRows.push({ Data: row.Data, Turno: row.Turno, Reparto: row['Reparto / servizio'], 'Tipo intervento': row['Tipo intervento'], Orario: `${row['Ora inizio']} - ${row['Ora fine']}`, Luogo: row.Luogo, Motivo: `${row.Descrizione} ${row.Note}`.trim() });
    });
  });
  const maxTipo = Math.max(1, ...Object.values(aggregate.byTipo || {}).map(n));
  const maxOrigine = Math.max(1, ...Object.values(aggregate.byOrigine || {}).map(n));
  const maxReparto = Math.max(1, ...Object.values(aggregate.byReparto || {}).map(n));
  const maxFascia = Math.max(1, ...Object.values(byFascia).map(n));
  const topTipo = orderedObjectRows(aggregate.byTipo)[0]?.Voce || '-';
  const topReparto = orderedObjectRows(aggregate.byReparto)[0]?.Voce || '-';
  const topFascia = orderedObjectRows(byFascia)[0]?.Voce || '-';
  const dashboardRows = [['REPORT AGGREGATO POLIZIA LOCALE - DASHBOARD EXCEL', ''], ['Periodo / data', aggregate.dateLabel], ['Report ricevuti', reports.length], ['Interventi totali', aggregate.totalInterventi], ['Violazioni / provvedimenti', aggregate.totaleViolazioni], ['Km totali percorsi', aggregate.kmTotali], ['Tipologia intervento prevalente', topTipo], ['Reparto più impegnato', topReparto], ['Fascia oraria più intensa', topFascia], ['Intervento più lungo', interventoPiuLungo.label], ['Note ufficiale', commanderNotes || '-']];
  const graficiRows = [['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byTipo).map(r => ['Tipologia interventi', r.Voce, r.Totale, makeBar(r.Totale, maxTipo)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byOrigine).map(r => ['Origine interventi', r.Voce, r.Totale, makeBar(r.Totale, maxOrigine)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byReparto).map(r => ['Reparti / servizi', r.Voce, r.Totale, makeBar(r.Totale, maxReparto)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(byFascia).map(r => ['Fasce orarie', r.Voce, r.Totale, makeBar(r.Totale, maxFascia)])];
  const riepilogoRows = [{ Sezione: 'KPI', Voce: 'Report ricevuti', Totale: reports.length }, { Sezione: 'KPI', Voce: 'Interventi totali', Totale: aggregate.totalInterventi }, { Sezione: 'KPI', Voce: 'Violazioni / provvedimenti', Totale: aggregate.totaleViolazioni }, { Sezione: 'KPI', Voce: 'Km totali percorsi', Totale: aggregate.kmTotali }, ...orderedObjectRows(aggregate.byTipo).map(r => ({ Sezione: 'Tipologia interventi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.byOrigine).map(r => ({ Sezione: 'Origine interventi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.byReparto).map(r => ({ Sezione: 'Reparti / servizi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(byFascia).map(r => ({ Sezione: 'Fasce orarie', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.counters, LABELS).map(r => ({ Sezione: 'Atti e violazioni', Voce: r.Voce, Totale: r.Totale }))];
  const readmeRows = [['Foglio', 'Contenuto'], ['Dashboard', 'KPI principali e sintesi operativa pronta per lettura comando.'], ['Grafici', 'Tabelle visuali con barre orizzontali compatibili con Excel, utili per stampa e analisi rapida.'], ['Interventi', 'Dataset dettagliato: una riga per ogni intervento, pronto per filtri e tabelle pivot.'], ['Report', 'Una riga per ogni report caricato.'], ['Riepilogo', 'Dati aggregati per categoria.'], ['Criticità', 'Interventi evidenziati automaticamente come rilevanti.'], ['Atti_Violazioni', 'Dettaglio contatori per report.']];
  return { dashboardRows, graficiRows, interventiRows, reportRows, riepilogoRows, criticitaRows, countersRows, readmeRows };
}
function applyWorksheetLayout(ws, widths = []) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
}
function exportExcelAvanzato(reports, aggregate, commanderNotes) {
  if (!reports.length) { alert('Carica almeno un file JSON prima di esportare Excel.'); return; }
  const data = buildExcelData(reports, aggregate, commanderNotes);
  const wb = XLSX.utils.book_new();
  const wsDashboard = XLSX.utils.aoa_to_sheet(data.dashboardRows); applyWorksheetLayout(wsDashboard, [38, 90]); XLSX.utils.book_append_sheet(wb, wsDashboard, 'Dashboard');
  const wsGrafici = XLSX.utils.aoa_to_sheet(data.graficiRows); applyWorksheetLayout(wsGrafici, [26, 48, 12, 28]); XLSX.utils.book_append_sheet(wb, wsGrafici, 'Grafici');
  const wsInterventi = XLSX.utils.json_to_sheet(data.interventiRows.length ? data.interventiRows : [{ Messaggio: 'Nessun intervento presente nei report caricati' }]); applyWorksheetLayout(wsInterventi, [12,16,18,32,36,10,12,28,24,12,12,14,16,30,48,36,36,12,18,16,16,18,18,18,18,18,30,32,48]); XLSX.utils.book_append_sheet(wb, wsInterventi, 'Interventi');
  const wsReport = XLSX.utils.json_to_sheet(data.reportRows.length ? data.reportRows : [{ Messaggio: 'Nessun report caricato' }]); applyWorksheetLayout(wsReport, [12,16,18,32,42,16,16,24,14,50]); XLSX.utils.book_append_sheet(wb, wsReport, 'Report');
  const wsRiepilogo = XLSX.utils.json_to_sheet(data.riepilogoRows); applyWorksheetLayout(wsRiepilogo, [26,48,14]); XLSX.utils.book_append_sheet(wb, wsRiepilogo, 'Riepilogo');
  const wsCriticita = XLSX.utils.json_to_sheet(data.criticitaRows.length ? data.criticitaRows : [{ Messaggio: 'Nessuna criticità automatica rilevata' }]); applyWorksheetLayout(wsCriticita, [12,16,30,28,18,30,80]); XLSX.utils.book_append_sheet(wb, wsCriticita, 'Criticità');
  const wsCounters = XLSX.utils.json_to_sheet(data.countersRows.length ? data.countersRows : [{ Messaggio: 'Nessun atto o violazione conteggiato' }]); applyWorksheetLayout(wsCounters, [12,16,32,40,12]); XLSX.utils.book_append_sheet(wb, wsCounters, 'Atti_Violazioni');
  const wsReadme = XLSX.utils.aoa_to_sheet(data.readmeRows); applyWorksheetLayout(wsReadme, [24,100]); XLSX.utils.book_append_sheet(wb, wsReadme, 'README');
  XLSX.writeFile(wb, `report-aggregato-analisi-${sanitizeFileName(aggregate.dateLabel)}.xlsx`, { compression: true });
}
function getKmTotali(report) { return (report.veicoli || []).reduce((sum, v) => sum + km(v), 0); }
function getTotaleViolazioni(report) {
  const c = report.counters || {};
  return ['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme','fermi','sequestri'].reduce((s, k) => s + n(c[k]), 0);
}
function operatorNames(report) { return (report.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).map(o => `${o.nome || 'Operatore'}${o.matricola ? ` mtr. ${o.matricola}` : ''}`); }
function extraDetails(i) {
  if (i.tipo === 'Sinistro stradale') return `   Dettagli: ${i.conFeriti}; veicoli coinvolti ${i.veicoliCoinvolti || '-'}; rilievi ${i.rilievi}\n`;
  if (i.tipo === 'Posto di controllo') return `   Controlli: veicoli ${i.veicoliControllati || '0'}; persone ${i.personeControllate || '0'}; verbali ${i.verbaliElevati || '0'}; fermi/sequestri ${i.fermiSequestri || '0'}\n`;
  if (i.tipo === 'Viabilità') return `   Motivo: ${i.motivoViabilita || '-'}; strade interessate: ${i.strade || '-'}\n`;
  return '';
}

function reportText(report) {
  const ops = operatorNames(report).map(x => `- ${x}`).join('\n') || '- Non indicati';
  const mezzi = (report.veicoli || []).map(v => `- ${v.sigla || 'Veicolo'} | Km inizio ${v.kmInizio || '-'} | Km fine ${v.kmFine || '-'} | Km percorsi ${km(v)}`).join('\n');
  const interventi = (report.interventi || []).map((i, idx) => {
    const scuole = i.tipo === 'Servizio scuole' ? (i.scuole || []).map((s, pos) => `   Scuola ${pos+1}: ${s.nome || '-'} | ${s.momento || '-'} | ${s.orario || '-'} | Criticità: ${s.criticita || '-'}`).join('\n') : '';
    return `${idx + 1}. ${i.tipo} | ${i.origine}${i.origine === 'Altro' ? ': ' + (i.origineAltro || '-') : ''}\n   Orario: ${i.oraInizio || '-'} - ${i.oraFine || '-'} | Luogo: ${i.luogo || '-'}\n   Descrizione: ${i.descrizione || '-'}\n   Esito: ${i.esito || '-'}\n${extraDetails(i)}${scuole ? '\n' + scuole : ''}\n   Note: ${i.note || '-'}`;
  }).join('\n\n') || '- Nessun intervento inserito';
  const c = report.counters || emptyCounters();
  return `REPORT DI SERVIZIO - POLIZIA LOCALE\n\nDATA: ${report.data}\nTURNO: ${turnoLabel(report)} (${report.orarioTipo})\nREPARTO: ${repartoLabel(report)}\n\nOPERATORI\n${ops}\n\nVEICOLI\n${mezzi || '- Non indicati'}\nTotale km percorsi: ${getKmTotali(report)}\n\nINTERVENTI EFFETTUATI\n${interventi}\n\nATTI REDATTI\nRelazioni: ${c.relazioni}\nAnnotazioni: ${c.annotazioni}\nVerbali CdS: ${c.verbaliCds}\nVerbali regolamenti: ${c.verbaliRegolamenti}\nSequestri amministrativi: ${c.sequestriAmministrativi}\nFermi amministrativi: ${c.fermiAmministrativi}\nSequestri penali: ${c.sequestriPenali}\nCNR: ${c.cnr}\nAltri atti: ${c.altriAttiNumero} ${c.altriAttiDescrizione || ''}\n\nVIOLAZIONI / PROVVEDIMENTI\nPreavvisi CdS: ${c.preavvisiCds}\nVdC CdS: ${c.vdcCds}\nRegolamento Polizia: ${c.regPolizia}\nRegolamento Edilizio: ${c.regEdilizio}\nRegolamento Benessere Animali: ${c.regBenessereAnimali}\nAnnonaria / commercio: ${c.annonaria}\nAltre norme: ${c.altreNorme} ${c.altreNormeDescrizione || ''}\nFermi: ${c.fermi}\nSequestri: ${c.sequestri}\nTOTALE: ${getTotaleViolazioni(report)}\n\nNOTE PER UDT / UFFICIALE DI COORDINAMENTO\n${report.noteUdt || '-'}\n\nDICHIARAZIONE\nGli operatori dichiarano che quanto riportato corrisponde fedelmente alle attività effettivamente svolte e riscontrate durante il turno di servizio.\nConferma dichiarazione: ${report.dichiarazione ? 'SI' : 'NO'}\n`;
}

function aggregateReports(reports) {
  const aggregate = { totalInterventi: 0, totaleViolazioni: 0, kmTotali: 0, byTipo: {}, byOrigine: {}, byReparto: {}, counters: {}, notes: [], dates: new Set() };
  reports.forEach(r => {
    if (r.data) aggregate.dates.add(r.data);
    aggregate.totalInterventi += (r.interventi || []).length;
    aggregate.totaleViolazioni += getTotaleViolazioni(r);
    aggregate.kmTotali += getKmTotali(r);
    aggregate.byReparto[repartoLabel(r)] = n(aggregate.byReparto[repartoLabel(r)]) + 1;
    (r.interventi || []).forEach(i => {
      aggregate.byTipo[i.tipo || 'Non indicato'] = n(aggregate.byTipo[i.tipo || 'Non indicato']) + 1;
      const origine = i.origine === 'Altro' ? `Altro: ${i.origineAltro || '-'}` : (i.origine || 'Non indicata');
      aggregate.byOrigine[origine] = n(aggregate.byOrigine[origine]) + 1;
    });
    Object.entries(r.counters || {}).forEach(([key, value]) => { if (typeof value === 'number') aggregate.counters[key] = n(aggregate.counters[key]) + n(value); });
    if (r.noteUdt) aggregate.notes.push({ data: r.data, turno: turnoLabel(r), reparto: repartoLabel(r), testo: r.noteUdt });
  });
  const dates = Array.from(aggregate.dates).sort();
  aggregate.dateLabel = dates.length === 0 ? today() : dates.length === 1 ? dates[0] : `${dates[0]}_${dates[dates.length - 1]}`;
  return aggregate;
}

function listEntries(obj, labels = {}) {
  const entries = Object.entries(obj || {}).filter(([, value]) => n(value) > 0).sort((a, b) => n(b[1]) - n(a[1]));
  return entries.length ? entries.map(([k, v]) => `- ${labels[k] || k}: ${v}`).join('\n') : '- Nessun dato';
}

function commanderReportText(aggregate, reports, commanderNotes) {
  const details = reports.map((r, idx) => `${idx + 1}. ${r.data} | ${turnoLabel(r)} | ${r.orarioTipo} | ${repartoLabel(r)} | Operatori: ${operatorNames(r).join(', ') || '-'} | Interventi: ${(r.interventi || []).length} | Violazioni/Provvedimenti: ${getTotaleViolazioni(r)} | Km: ${getKmTotali(r)}`).join('\n') || '- Nessun report caricato';
  const notes = aggregate.notes.length ? aggregate.notes.map(n => `- ${n.data} ${n.turno} ${n.reparto}: ${n.testo}`).join('\n') : '- Nessuna nota rilevante';
  return `REPORT AGGREGATO PER IL COMANDANTE\n\nPERIODO / DATA: ${aggregate.dateLabel}\nREPORT RICEVUTI: ${reports.length}\nINTERVENTI TOTALI: ${aggregate.totalInterventi}\nVIOLAZIONI / PROVVEDIMENTI TOTALI: ${aggregate.totaleViolazioni}\nKM TOTALI PERCORSI: ${aggregate.kmTotali}\n\nINTERVENTI PER TIPOLOGIA\n${listEntries(aggregate.byTipo)}\n\nORIGINE INTERVENTI\n${listEntries(aggregate.byOrigine)}\n\nREPARTI / SERVIZI RENDICONTATI\n${listEntries(aggregate.byReparto)}\n\nATTI E VIOLAZIONI\n${listEntries(aggregate.counters, LABELS)}\n\nDETTAGLIO REPORT RICEVUTI\n${details}\n\nNOTE RILEVANTI / CRITICITÀ\n${notes}\n\nNOTE DELL'UFFICIALE PER IL COMANDANTE\n${commanderNotes || '-'}\n`;
}

function relevantInterventions(reports) {
  const out = [];
  reports.forEach(r => (r.interventi || []).forEach(i => {
    const durata = durataMinuti(i.oraInizio, i.oraFine);
    if (isInterventoCritico(i) || n(durata) >= 90) out.push({ report: r, intervento: i, durata });
  }));
  return out.sort((a, b) => (parseTimeToMinutes(a.intervento.oraInizio) ?? 9999) - (parseTimeToMinutes(b.intervento.oraInizio) ?? 9999));
}
function officialSynthesis(aggregate, reports) {
  const topTipo = Object.entries(aggregate.byTipo || {}).sort((a, b) => n(b[1]) - n(a[1]))[0]?.[0] || 'attività operative ordinarie';
  const all = reports.flatMap(r => r.interventi || []);
  const conFeriti = all.filter(i => i.tipo === 'Sinistro stradale' && i.conFeriti === 'Con feriti').length;
  const tso = all.filter(i => i.tipo === 'TSO').length;
  const aso = all.filter(i => i.tipo === 'ASO').length;
  return `Nel turno indicato sono stati acquisiti ${reports.length} report degli operatori, per complessivi ${aggregate.totalInterventi} interventi rendicontati. L'attività prevalente risulta: ${topTipo}. Sono state registrate ${aggregate.totaleViolazioni} violazioni/provvedimenti e ${aggregate.kmTotali} km complessivi. Si segnalano ${conFeriti} sinistri con feriti, ${tso} TSO e ${aso} ASO.`;
}
function officialEventsText(reports) {
  return relevantInterventions(reports).map(({ report, intervento, durata }) => {
    const op = operatorNames(report).join(', ') || repartoLabel(report);
    const extra = durata ? ` Durata indicativa: ${durata} minuti.` : '';
    return `- Ore ${intervento.oraInizio || '--'}: ${intervento.tipo || 'intervento'} in ${intervento.luogo || 'luogo non indicato'}, pattuglia/reparto ${op}. ${intervento.descrizione || ''} Esito: ${intervento.esito || '-'}${extra}`;
  }).join('\n');
}
function officialReportText(aggregate, reports, official, autoSintesi, autoEventi) {
  const attivita = (official.attivitaIspettive || []).filter(a => a.tipo || a.reparto || a.luogo || a.esito || a.note).map((a, idx) => `${idx + 1}. ${a.tipo || '-'} | ${a.reparto || '-'} | ${a.luogo || '-'} | ${a.orario || '-'} | Esito: ${a.esito || '-'} | Violazioni: ${a.violazioni || '-'} | Note: ${a.note || '-'}`).join('\n') || '- Nessuna attività ispettiva indicata';
  return `REPORT UFFICIALE DI TURNO\n\nDATA E TURNO\n${official.data || aggregate.dateLabel} ${official.turno || ''}\n\nBRIEFING OPERATIVO\n${official.briefing || '-'}\n\nA.P.L. ASSENTI\n${official.assenti || '-'}\n\nA.P.L. IN RITARDO\n${official.ritardi || '-'}\n\nNOTE\n${official.noteGenerali || '-'}\n\nSINTESI OPERATIVA\n${autoSintesi}\n${official.eventiManuali ? '\nIntegrazioni: ' + official.eventiManuali : ''}\n\nEVENTI DEGNI DI RILIEVO\n${autoEventi || '- Nessun evento rilevante automatico rilevato'}\n\nANOMALIE RISCONTRATE DURANTE IL TURNO\n${official.anomalie || '-'}\n\nATTIVITÀ ISPETTIVE\n${attivita}\n\nESITI\n${official.esiti || '-'}\n\nCOMUNICAZIONE ALL'E.Q. DI TURNO\n${official.comunicazioneEq || '-'}\n\nNOTA PER IL COMANDANTE\n${official.notaComandante || '-'}\n\nVIOLAZIONI RISCONTRATE\nTotale violazioni/provvedimenti: ${aggregate.totaleViolazioni}\n\nFIRMA\n${official.qualifica || ''}\n${official.ufficiale || ''}`;
}
function buildOfficialShiftPdf(aggregate, reports, official, autoSintesi, autoEventi) {
  const title = 'REPORT UFFICIALE DI TURNO';
  const subtitle = `${official.data || aggregate.dateLabel} | ${official.turno || ''}`;
  const doc = makePdf(title, subtitle);
  let y = 54;
  y = section(doc, 'Data e turno', y, title, subtitle);
  y = kvGrid(doc, [{ label: 'Data', value: official.data || aggregate.dateLabel }, { label: 'Turno', value: official.turno || '-' }, { label: 'Ufficiale di turno', value: official.ufficiale || '-' }, { label: 'Qualifica', value: official.qualifica || '-' }], y, 2, title, subtitle) + 2;
  y = section(doc, 'Riepilogo rapido', y, title, subtitle);
  y = kvGrid(doc, [{ label: 'Report operatori', value: reports.length }, { label: 'Interventi totali', value: aggregate.totalInterventi }, { label: 'Violazioni / provvedimenti', value: aggregate.totaleViolazioni }, { label: 'Km complessivi', value: aggregate.kmTotali }], y, 4, title, subtitle) + 2;
  y = section(doc, 'Briefing operativo', y, title, subtitle); y = paragraph(doc, official.briefing || '-', y, title, subtitle);
  y = section(doc, 'Personale', y, title, subtitle);
  y = kvGrid(doc, [{ label: 'A.P.L. assenti', value: official.assenti || '-' }, { label: 'A.P.L. in ritardo', value: official.ritardi || '-' }, { label: 'Note', value: official.noteGenerali || '-' }], y, 1, title, subtitle) + 2;
  y = section(doc, 'Sintesi operativa', y, title, subtitle); y = paragraph(doc, `${autoSintesi}${official.eventiManuali ? '\n\nIntegrazioni: ' + official.eventiManuali : ''}`, y, title, subtitle);
  y = section(doc, 'Eventi degni di rilievo', y, title, subtitle); y = paragraph(doc, autoEventi || 'Nessun evento rilevante automatico rilevato.', y, title, subtitle);
  y = section(doc, 'Anomalie riscontrate durante il turno', y, title, subtitle); y = paragraph(doc, official.anomalie || '-', y, title, subtitle);
  y = section(doc, 'Attività ispettive', y, title, subtitle);
  const attivitaRows = (official.attivitaIspettive || []).filter(a => a.tipo || a.reparto || a.luogo || a.esito || a.note).map(a => [a.tipo || '-', a.reparto || '-', a.luogo || '-', a.orario || '-', a.esito || '-', a.violazioni || '-', a.note || '-']);
  y = simpleTable(doc, ['Tipo', 'Reparto', 'Luogo', 'Orario', 'Esito', 'Viol.', 'Note'], attivitaRows.length ? attivitaRows : [['-', '-', '-', '-', '-', '-', '-']], y, [28, 30, 28, 18, 30, 18, 34], title, subtitle);
  y = section(doc, 'Esiti', y, title, subtitle); y = paragraph(doc, official.esiti || '-', y, title, subtitle);
  y = section(doc, "Comunicazione all'E.Q. di turno", y, title, subtitle); y = paragraph(doc, official.comunicazioneEq || '-', y, title, subtitle);
  y = section(doc, 'Nota per il Comandante', y, title, subtitle); y = paragraph(doc, official.notaComandante || '-', y, title, subtitle);
  y = section(doc, 'Violazioni riscontrate', y, title, subtitle);
  const violationRows = reports.map(r => [operatorNames(r).join(', ') || '-', repartoLabel(r), n((r.counters || {}).preavvisiCds), n((r.counters || {}).vdcCds), getTotaleViolazioni(r)]);
  y = simpleTable(doc, ['AA.PP.LL.', 'Zona/Reparto', 'Preavvisi CdS', 'VdC CdS/Altro', 'Totale'], violationRows.length ? violationRows : [['-', '-', '-', '-', '-']], y, [58, 48, 26, 28, 26], title, subtitle);
  y = ensureSpace(doc, y, 22, title, subtitle); doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.text('FIRMA:', 16, y + 4); doc.text(official.qualifica || '-', 16, y + 12); doc.text(official.ufficiale || '-', 16, y + 18); doc.line(80, y + 18, 190, y + 18);
  addFooter(doc); return doc;
}

createRoot(document.getElementById('root')).render(<App />);
