import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import jsPDF from 'jspdf';
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

function today() {
  return new Date().toISOString().slice(0, 10);
}

function Field({ label, children }) {
  return <label className="field"><span>{label}</span>{children}</label>;
}

function Input({ value, onChange, type = 'text', placeholder = '' }) {
  return <input type={type} value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} />;
}

function Textarea({ value, onChange, placeholder = '' }) {
  return <textarea value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} />;
}

function Select({ value, onChange, children }) {
  return <select value={value} onChange={e => onChange(e.target.value)}>{children}</select>;
}

function Counter({ label, value, onChange }) {
  const n = Number(value || 0);
  return (
    <div className="counter">
      <span>{label}</span>
      <button type="button" onClick={() => onChange(Math.max(0, n - 1))}>−</button>
      <input type="number" min="0" value={n} onChange={e => onChange(Number(e.target.value || 0))} />
      <button type="button" onClick={() => onChange(n + 1)}>+</button>
    </div>
  );
}

function App() {
  const [report, setReport] = useState({
    data: today(), turno: '06.00-13.00', altroTurnoInizio: '', altroTurnoFine: '', orarioTipo: 'Ordinario',
    reparto: 'Radiomobile', altroServizio: '', destinatario: '',
    operatori: [emptyOperatore(), emptyOperatore()], veicoli: [emptyVeicolo()], interventi: [emptyIntervento()],
    counters: emptyCounters(), noteUdt: '', dichiarazione: false
  });

  const update = (patch) => setReport(prev => ({ ...prev, ...patch }));
  const updateArray = (key, index, patch) => setReport(prev => ({ ...prev, [key]: prev[key].map((x, i) => i === index ? { ...x, ...patch } : x) }));
  const addArray = (key, item) => setReport(prev => ({ ...prev, [key]: [...prev[key], item] }));
  const removeArray = (key, index) => setReport(prev => ({ ...prev, [key]: prev[key].filter((_, i) => i !== index) }));

  const totalKm = useMemo(() => report.veicoli.reduce((sum, v) => sum + Math.max(0, Number(v.kmFine || 0) - Number(v.kmInizio || 0)), 0), [report.veicoli]);
  const totaleViolazioni = useMemo(() => {
    const c = report.counters;
    return ['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme','fermi','sequestri']
      .reduce((s, k) => s + Number(c[k] || 0), 0);
  }, [report.counters]);

  function reportText() {
    const turno = report.turno === 'Altro orario' ? `${report.altroTurnoInizio}-${report.altroTurnoFine}` : report.turno;
    const reparto = report.reparto === 'Altri servizi' ? `${report.reparto}: ${report.altroServizio}` : report.reparto;
    const ops = report.operatori.filter(o => o.nome || o.matricola || o.qualifica).map(o => `- ${o.nome} ${o.qualifica ? '(' + o.qualifica + ')' : ''} mtr. ${o.matricola}`).join('\n') || '- Non indicati';
    const mezzi = report.veicoli.map(v => `- ${v.sigla || 'Veicolo'} | Km inizio ${v.kmInizio || '-'} | Km fine ${v.kmFine || '-'} | Km percorsi ${Math.max(0, Number(v.kmFine || 0) - Number(v.kmInizio || 0))}`).join('\n');
    const interventi = report.interventi.map((i, idx) => {
      const scuole = i.tipo === 'Servizio scuole' ? i.scuole.map((s, n) => `   Scuola ${n+1}: ${s.nome || '-'} | ${s.momento || '-'} | ${s.orario || '-'} | Criticità: ${s.criticita || '-'}`).join('\n') : '';
      const dettagli = extraDetails(i);
      return `${idx + 1}. ${i.tipo} | ${i.origine}${i.origine === 'Altro' ? ': ' + i.origineAltro : ''}\n   Orario: ${i.oraInizio || '-'} - ${i.oraFine || '-'} | Luogo: ${i.luogo || '-'}\n   Descrizione: ${i.descrizione || '-'}\n   Esito: ${i.esito || '-'}\n${dettagli}${scuole ? '\n' + scuole : ''}\n   Note: ${i.note || '-'}`;
    }).join('\n\n');
    const c = report.counters;
    return `REPORT DI SERVIZIO - POLIZIA LOCALE\n\nDATA: ${report.data}\nTURNO: ${turno} (${report.orarioTipo})\nREPARTO: ${reparto}\n\nOPERATORI\n${ops}\n\nVEICOLI\n${mezzi}\nTotale km percorsi: ${totalKm}\n\nINTERVENTI EFFETTUATI\n${interventi}\n\nATTI REDATTI\nRelazioni: ${c.relazioni}\nAnnotazioni: ${c.annotazioni}\nVerbali CdS: ${c.verbaliCds}\nVerbali regolamenti: ${c.verbaliRegolamenti}\nSequestri amministrativi: ${c.sequestriAmministrativi}\nFermi amministrativi: ${c.fermiAmministrativi}\nSequestri penali: ${c.sequestriPenali}\nCNR: ${c.cnr}\nAltri atti: ${c.altriAttiNumero} ${c.altriAttiDescrizione}\n\nVIOLAZIONI / PROVVEDIMENTI\nPreavvisi CdS: ${c.preavvisiCds}\nVdC CdS: ${c.vdcCds}\nRegolamento Polizia: ${c.regPolizia}\nRegolamento Edilizio: ${c.regEdilizio}\nRegolamento Benessere Animali: ${c.regBenessereAnimali}\nAnnonaria / commercio: ${c.annonaria}\nAltre norme: ${c.altreNorme} ${c.altreNormeDescrizione}\nFermi: ${c.fermi}\nSequestri: ${c.sequestri}\nTOTALE: ${totaleViolazioni}\n\nNOTE PER UDT / UFFICIALE DI COORDINAMENTO\n${report.noteUdt || '-'}\n\nDICHIARAZIONE\nGli operatori dichiarano che quanto riportato corrisponde fedelmente alle attività effettivamente svolte e riscontrate durante il turno di servizio.\nConferma dichiarazione: ${report.dichiarazione ? 'SI' : 'NO'}\n`;
  }

  function extraDetails(i) {
    if (i.tipo === 'Sinistro stradale') return `   Dettagli: ${i.conFeriti}; veicoli coinvolti ${i.veicoliCoinvolti || '-'}; rilievi ${i.rilievi}\n`;
    if (i.tipo === 'Posto di controllo') return `   Controlli: veicoli ${i.veicoliControllati || '0'}; persone ${i.personeControllate || '0'}; verbali ${i.verbaliElevati || '0'}; fermi/sequestri ${i.fermiSequestri || '0'}\n`;
    if (i.tipo === 'Viabilità') return `   Motivo: ${i.motivoViabilita || '-'}; strade interessate: ${i.strade || '-'}\n`;
    return '';
  }

  function generatePdf() {
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(15);
    doc.text('REPORT DI SERVIZIO', 105, 15, { align: 'center' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    const lines = doc.splitTextToSize(reportText(), 180);
    let y = 28;
    lines.forEach(line => {
      if (y > 280) { doc.addPage(); y = 15; }
      doc.text(line, 15, y);
      y += 5;
    });
    doc.save(`report-turno-${report.data}.pdf`);
  }

  function sendMail() {
    const subject = encodeURIComponent(`Report turno Polizia Locale - ${report.data} - ${report.turno}`);
    const body = encodeURIComponent(`Si trasmette il report del turno di servizio.\n\nNota: allegare il PDF scaricato dall'app.\n\n${reportText().slice(0, 1200)}${reportText().length > 1200 ? '\n\n[Report completo in allegato PDF]' : ''}`);
    window.location.href = `mailto:${encodeURIComponent(report.destinatario)}?subject=${subject}&body=${body}`;
  }

  return (
    <main>
      <header className="hero">
        <div>
          <p className="eyebrow">Polizia Locale</p>
          <h1>Report Turno</h1>
          <p>Compilazione guidata, PDF e invio email a fine servizio.</p>
        </div>
      </header>

      <section className="card">
        <h2>1. Dati turno</h2>
        <div className="grid">
          <Field label="Data servizio"><Input type="date" value={report.data} onChange={v => update({ data: v })} /></Field>
          <Field label="Turno"><Select value={report.turno} onChange={v => update({ turno: v })}>{TURNI.map(t => <option key={t}>{t}</option>)}</Select></Field>
          <Field label="Tipologia orario"><Select value={report.orarioTipo} onChange={v => update({ orarioTipo: v })}><option>Ordinario</option><option>Straordinario</option></Select></Field>
          <Field label="Reparto"><Select value={report.reparto} onChange={v => update({ reparto: v })}>{REPARTI.map(r => <option key={r}>{r}</option>)}</Select></Field>
        </div>
        {report.turno === 'Altro orario' && <div className="grid two"><Field label="Ora inizio"><Input value={report.altroTurnoInizio} onChange={v => update({ altroTurnoInizio: v })} placeholder="es. 10.00" /></Field><Field label="Ora fine"><Input value={report.altroTurnoFine} onChange={v => update({ altroTurnoFine: v })} placeholder="es. 17.00" /></Field></div>}
        {report.reparto === 'Altri servizi' && <Field label="Specificare servizio"><Input value={report.altroServizio} onChange={v => update({ altroServizio: v })} /></Field>}
      </section>

      <section className="card">
        <h2>2. Operatori</h2>
        {report.operatori.map((o, idx) => <div className="rowCard" key={idx}><div className="grid three"><Field label="Nome e cognome"><Input value={o.nome} onChange={v => updateArray('operatori', idx, { nome: v })} /></Field><Field label="Matricola"><Input value={o.matricola} onChange={v => updateArray('operatori', idx, { matricola: v })} /></Field><Field label="Qualifica"><Input value={o.qualifica} onChange={v => updateArray('operatori', idx, { qualifica: v })} /></Field></div><button className="ghost" onClick={() => removeArray('operatori', idx)}>Rimuovi</button></div>)}
        <button onClick={() => addArray('operatori', emptyOperatore())}>+ Aggiungi operatore</button>
      </section>

      <section className="card">
        <h2>3. Veicoli e chilometraggio</h2>
        {report.veicoli.map((v, idx) => <div className="rowCard" key={idx}><div className="grid four"><Field label="Veicolo"><Input value={v.sigla} onChange={x => updateArray('veicoli', idx, { sigla: x })} /></Field><Field label="Km inizio"><Input type="number" value={v.kmInizio} onChange={x => updateArray('veicoli', idx, { kmInizio: x })} /></Field><Field label="Km fine"><Input type="number" value={v.kmFine} onChange={x => updateArray('veicoli', idx, { kmFine: x })} /></Field><Field label="Km percorsi"><input readOnly value={Math.max(0, Number(v.kmFine || 0) - Number(v.kmInizio || 0))} /></Field></div><button className="ghost" onClick={() => removeArray('veicoli', idx)}>Rimuovi</button></div>)}
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
          <Counter label="Relazioni di servizio" value={report.counters.relazioni} onChange={v => update({ counters: { ...report.counters, relazioni: v } })} />
          <Counter label="Annotazioni di servizio" value={report.counters.annotazioni} onChange={v => update({ counters: { ...report.counters, annotazioni: v } })} />
          <Counter label="Verbali CdS" value={report.counters.verbaliCds} onChange={v => update({ counters: { ...report.counters, verbaliCds: v } })} />
          <Counter label="Verbali regolamenti" value={report.counters.verbaliRegolamenti} onChange={v => update({ counters: { ...report.counters, verbaliRegolamenti: v } })} />
          <Counter label="Sequestri amministrativi" value={report.counters.sequestriAmministrativi} onChange={v => update({ counters: { ...report.counters, sequestriAmministrativi: v } })} />
          <Counter label="Fermi amministrativi" value={report.counters.fermiAmministrativi} onChange={v => update({ counters: { ...report.counters, fermiAmministrativi: v } })} />
          <Counter label="Sequestri penali" value={report.counters.sequestriPenali} onChange={v => update({ counters: { ...report.counters, sequestriPenali: v } })} />
          <Counter label="C.N.R." value={report.counters.cnr} onChange={v => update({ counters: { ...report.counters, cnr: v } })} />
        </div>
        <div className="grid two"><Counter label="Altri atti" value={report.counters.altriAttiNumero} onChange={v => update({ counters: { ...report.counters, altriAttiNumero: v } })} /><Field label="Descrizione altri atti"><Input value={report.counters.altriAttiDescrizione} onChange={v => update({ counters: { ...report.counters, altriAttiDescrizione: v } })} /></Field></div>
      </section>

      <section className="card">
        <h2>6. Violazioni e provvedimenti</h2>
        <div className="counterGrid">
          <Counter label="Preavvisi CdS" value={report.counters.preavvisiCds} onChange={v => update({ counters: { ...report.counters, preavvisiCds: v } })} />
          <Counter label="VdC CdS" value={report.counters.vdcCds} onChange={v => update({ counters: { ...report.counters, vdcCds: v } })} />
          <Counter label="Reg. Polizia" value={report.counters.regPolizia} onChange={v => update({ counters: { ...report.counters, regPolizia: v } })} />
          <Counter label="Reg. Edilizio" value={report.counters.regEdilizio} onChange={v => update({ counters: { ...report.counters, regEdilizio: v } })} />
          <Counter label="Reg. Benessere Animali" value={report.counters.regBenessereAnimali} onChange={v => update({ counters: { ...report.counters, regBenessereAnimali: v } })} />
          <Counter label="Annonaria / commercio" value={report.counters.annonaria} onChange={v => update({ counters: { ...report.counters, annonaria: v } })} />
          <Counter label="Altre norme" value={report.counters.altreNorme} onChange={v => update({ counters: { ...report.counters, altreNorme: v } })} />
          <Counter label="Fermi" value={report.counters.fermi} onChange={v => update({ counters: { ...report.counters, fermi: v } })} />
          <Counter label="Sequestri" value={report.counters.sequestri} onChange={v => update({ counters: { ...report.counters, sequestri: v } })} />
        </div>
        <Field label="Specificare altre norme"><Input value={report.counters.altreNormeDescrizione} onChange={v => update({ counters: { ...report.counters, altreNormeDescrizione: v } })} /></Field>
        <div className="totalBox">Totale violazioni / provvedimenti: <strong>{totaleViolazioni}</strong></div>
      </section>

      <section className="card">
        <h2>7. Note e invio</h2>
        <Field label="Note per UDT / Ufficiale di coordinamento"><Textarea value={report.noteUdt} onChange={v => update({ noteUdt: v })} /></Field>
        <Field label="Email ufficiale destinatario"><Input value={report.destinatario} onChange={v => update({ destinatario: v })} placeholder="es. ufficiale@comune.monza.it" /></Field>
        <label className="check"><input type="checkbox" checked={report.dichiarazione} onChange={e => update({ dichiarazione: e.target.checked })} /> Confermo la dichiarazione finale degli operatori.</label>
        <div className="actions"><button onClick={generatePdf}>Scarica PDF</button><button className="primary" onClick={sendMail}>Invia email precompilata</button></div>
      </section>

      <section className="card preview">
        <h2>Anteprima report</h2>
        <pre>{reportText()}</pre>
      </section>
    </main>
  );
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

createRoot(document.getElementById('root')).render(<App />);
