import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import jsPDF from 'jspdf';
import * as XLSX from 'xlsx';
import './style.css';
import { supabase } from './lib/supabaseClient';

const MONZA_COMMAND_ID = 'ae6f07c1-404f-41a1-9be7-9ff0bc83c325';

function normalizeTime(value) {
  if (!value) return null;
  const clean = String(value).replace('.', ':').trim();
  const parts = clean.split(':');
  if (parts.length < 2) return null;
  return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:00`;
}

function getShiftTimes(report) {
  if (report.turno === 'Altro orario') {
    return {
      start_time: normalizeTime(report.altroTurnoInizio),
      end_time: normalizeTime(report.altroTurnoFine),
    };
  }

  const [start, end] = String(report.turno || '').split('-');

  return {
    start_time: normalizeTime(start),
    end_time: normalizeTime(end),
  };
}
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

const REPARTI_CON_ZONA = ['Radiomobile', 'Quartieri', 'Intervento rapido motociclisti'];
function richiedeZonaServizio(reparto) { return REPARTI_CON_ZONA.includes(reparto); }

const TIPI_INTERVENTO = [
  'Sinistro stradale', 'Codice della strada','TSO', 'ASO', 'Posto di controllo', 'Viabilità', 'Servizio scuole',
  'Controllo commerciale / annonaria', 'Controllo edilizio', 'Controllo parchi / aree verdi',
  'Sicurezza urbana', 'Intervento per animali', 'Abbandono rifiuti', 'Rumori / disturbo quiete',
  'Supporto ad altro ente', 'Notifica / accertamento', 'Altro'
];

const DETTAGLI_CODICE_STRADA = [
  'Controllo soste',
  'Buca su carreggiata',
  'Veicolo sospetto',
  'Segnaletica danneggiata',
  'Guasto semaforo'
];

const ORIGINI = ['Centrale Operativa', 'UDT', 'Di iniziativa', 'Altro'];


const PERSONALE = [
  ['DONGIOVANNI','GIOVANNI','11267','Dirigente'], ['GALLI','FRANCESCA','6164','Commissario capo coord.'], ['MORA','GIORGIO','5181','Commissario capo coord.'], ['ESPOSITO','GERARDO','2675','Commissario capo coord.'], ['IENGO','FERDINANDO','5182','Commissario capo coord.'], ['DI GIOVANNANTONIO','LEONILDA','5723','Commissario capo coord.'], ['ANGLANI','PAOLO','3093','Commissario Capo'], ['DI SIMONE','CINZIA ANTONELLA','4878','Commissario'], ['BARNABA','ANTONIO','6160','Vice Commissario'], ['FABRIS','ALESSANDRO','6743','Vice Commissario'], ['MAGGI','GIUSEPPANTONIO','7499','Vice Commissario'], ['NEOLA','SALVATORE','10331','Vice Commissario'], ['MILITELLO','LORETO','7701','Vice Commissario'], ['RANGO','MONIA','9969','Vice Commissario'], ['GALASSO','PIERLUIGI','11471','Vice Commissario'], ['LONGO','GIOVANNA','11012','Vice Commissario'], ['LOCATI','EMANUELA','3104','Specialista di Vigilanza'], ['FERRARO','SERGIO','10244','Sovrintendente esperto'], ['ADAMO','MASSIMO','5024','Sovrintendente esperto'], ['VISCIONE','ADELE ROSA','4818','Sovrintendente scelto'], ['PESCE','LAURA','5627','Sovrintendente scelto'], ['CERIELLO','LORENZO','5673','Sovrintendente scelto'], ['GIORGILLI','GUIDO','5680','Sovrintendente scelto'], ['ROTELLI','CLAUDIO','6179','Sovrintendente'], ['LAURIOLA','GIOVANNI','6169','Sovrintendente'], ['SCARAMOZZINO','MARIA ASSUNTA','6304','Sovrintendente'], ['REPICI','VINCENZO','6631','Sovrintendente'], ['CELLAMARE','LEONORA','6334','Sovrintendente'], ['MINACAPILLI','SALVATORE','6337','Sovrintendente'], ['MARZOLI','SABRINA','7004','Sovrintendente'], ['MONTALBANO','CATERINA','6648','Sovrintendente'], ['CAPUTO','MAURIZIO','6901','Sovrintendente'], ['GRANILLO','GIUSEPPE','6788','Sovrintendente'], ['MORELLI','PAOLO','6628','Sovrintendente'], ['SEPE','ANTONIO','6867','Sovrintendente'], ['SCINO','ROBERTO','6869','Sovrintendente'], ['LEMBO','ANTONINO','7005','Assistente esperto'], ['D ALCONZO','LUIGI','7009','Assistente esperto'], ['SCRENCI','SALVATORE','7016','Assistente esperto'], ['PALAZZO','MASSIMILIANO','7017','Assistente esperto'], ['COLOMBO','CHIARA','7491','Assistente esperto'], ['VESCERA','MARCO','7143','Assistente esperto'], ['INVERNIZZI','ANDREA PIETRO','7495','Assistente esperto'], ['PIEMONTESE','PAOLO','7703','Assistente esperto'], ['MAIORANO','GERARDO','7700','Assistente esperto'], ['CALO','DAVIDE','7704','Vice Commissario'], ['GIUGNO','CLAUDIO','7494','Assistente esperto'], ['CARAMELLA','GIUSEPPE','10324','Assistente esperto'], ['AGNELLO','LORENZO','10107','Assistente esperto'], ['SUOZZO','MARIO','10965','Assistente Scelto'], ['LAMONICA','PAOLA','9818','Assistente Scelto'], ['SCALISE','MARCO','10820','Assistente Scelto'], ['DELL ERBA','DOMENICO','10461','Assistente Scelto'], ['MARCONI','INES BARBARA','5663','Assistente Scelto'], ['ONESTI','FRANCESCA','9654','Assistente Scelto'], ['LAZZATI','ALESSANDRO','9657','Assistente Scelto'], ['IANDOLO','FABIO','9659','Assistente Scelto'], ['TRENTO','LORENA','9668','Assistente Scelto'], ['GOFFO','RAUL','9819','Assistente Scelto'], ['SALA','ERIKA','10117','Assistente Scelto'], ['BANCOLINI','SIMONE','11150','Assistente'], ['DASSI','BARBARA','10010','Assistente Scelto'], ['SCIBILIA','LOREDANA','10459','Assistente'], ['SCARPIELLO','RAFFAELE','10249','Assistente'], ['SPATARI','MARIA PAOLA','10251','Assistente'], ['AQUINO','PATRIZIA','10333','Assistente'], ['SAMMARCO','MARIA ORSOLA','10457','Assistente'], ['PEDRAZZI','ALICE','10685','Agente Scelto'], ['MOTTA','CRISTIAN','10686','Agente Scelto'], ['SPINA','GIUSEPPE','10687','Agente Scelto'], ['TOSI','FABIO','11151','Agente Scelto'], ['MUSCIACCHIO','ANTONIO','10774','Agente Scelto'], ['ZERRI','EMANUELE','11144','Agente Scelto'], ['BOLIGNANI','MIRKO','10909','Agente Scelto'], ['SCIBELLI','FILIPPO','10897','Agente Scelto'], ['CANTORE','PIERLUIGI','9996','Agente Scelto'], ['CUSCUNA','ALESSANDRO','10896','Agente Scelto'], ['ALABISO','ROBERTO','11056','Agente Scelto'], ['GIORDANO','ANNAMARIA','11010','Agente Scelto'], ['CASCIANA','CONCETTA CHIARA','11053','Agente Scelto'], ['CAROTENUTO','ILARIA','11004','Agente Scelto'], ['ABBRESCIA','MARGHERITA','11002','Agente Scelto'], ['PAOLELLA','GABRIELE','11014','Agente Scelto'], ['FUSCO','LUIGI','11009','Agente Scelto'], ['DELL UTRI','MARCO','11007','Agente Scelto'], ['BERGNA','EMANUELE','11246','Agente'], ['SGUEGLIA','FABIO','11179','Agente'], ['BARBIERI','ENRICO','11248','Agente'], ['LATTUADA','SERGIO IVAN','11052','Agente'], ['COPPOLA','DANIELE','11369','Agente'], ['PAPPALARDO','MARIAGRAZIA','11371','Agente'], ['RANIERI','CRISTINA MARTINA','11245','Agente'], ['PATANE','VALENTINA CARMELA','10910','Agente'], ['ELEFANTE COSTANZA','CHRISTIAN','11250','Agente'], ['TOMARCHIO','ROSARIO','11252','Agente'], ['CONDELLO','ANDREA','11255','Agente'], ['IOVINELLA','MARCO','11256','Agente'], ['CALABRESE','FRANCESCA','11257','Agente'], ['SALIERNO','DESIDERIA','11365','Agente'], ['LAPORTA','DAMIANO','11253','Agente'], ['BARZAGHI','MATTEO','11268','Agente'], ['PEPE','DONATO','11285','Agente'], ['GIAMBELLUCA','SALVATORE','11286','Agente'], ['ANGIULI','VITO','11296','Agente'], ['CORATTO','DAVIDE','11352','Agente'], ['TALERICO','CONCETTA','11351','Agente'], ['RAIOLA','ALFONSO','11364','Agente'], ['D ANGELO','EMANUELA','11366','Agente'], ['FILETTI','MATTIA','11367','Agente'], ['PORTALURI','RICCARDO','11368','Agente'], ['OLTOLINI','MATTIA','11404','Agente']
].map(([cognome, nome, matricola, qualifica]) => ({ cognome, nome, matricola, qualifica }));
const QUALIFICHE_UFFICIALI = ['Dirigente', 'Commissario capo coord.', 'Commissario Capo', 'Commissario', 'Vice Commissario', 'Specialista di Vigilanza'];
function isUfficiale(persona) { return persona && QUALIFICHE_UFFICIALI.includes(persona.qualifica); }
function fullNamePersona(p) { return p.cognome + ' ' + p.nome; }
function findPersonaByMatricola(matricola) { return PERSONALE.find(p => p.matricola === String(matricola || '').trim()); }

const emptyOperatore = () => ({ nome: '', matricola: '', qualifica: '' });
const emptyVeicolo = () => ({ sigla: '', kmInizio: '', kmFine: '', carburante: 'No', importoCarburante: '', oraPrelievoCard: '', oraRestituzioneCard: '', anomaliaVeicolo: '' });
const emptyScuola = () => ({ nome: '', momento: '', orario: '', criticita: '' });
const emptyDocumentoRitirato = () => ({ tipo: '', quantita: '', note: '' });
const emptyVerbaleDistinta = () => ({ numero: '', tipo: '', norma: '', importo: '', operatore: '', note: '' });
const emptyIntervento = () => ({
  tipo: 'Sinistro stradale', origine: 'Centrale Operativa', origineAltro: '', oraInizio: '', oraFine: '', luogo: '', descrizione: '', esito: '', note: '',
  conFeriti: 'Senza feriti', veicoliCoinvolti: '', rilievi: 'No', personeControllate: '', veicoliControllati: '', verbaliElevati: '', fermiSequestri: '',
  motivoViabilita: '', strade: '', scuole: [emptyScuola()],
  cdsDettaglio: '', cdsRimozione: 'No', cdsMotivazione: '', cdsRipristino: 'No', cdsFeriti: 'No', cdsVeicoliCoinvolti: '',
  cdsVerificaEffettuata: 'No', cdsSegnalazione: 'No', cdsPericolo: 'No', cdsInterventoRichiesto: 'No', cdsStatoSemaforo: 'spento'
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
    schemaVersion: 3,
    data: today(), turno: '06.00-13.00', altroTurnoInizio: '', altroTurnoFine: '', orarioTipo: 'Ordinario',
    reparto: 'Radiomobile', altroServizio: '', zonaServizio: '', destinatario: '',
    operatori: [emptyOperatore(), emptyOperatore()], veicoli: [emptyVeicolo()], interventi: [emptyIntervento()],
    counters: emptyCounters(), documentiRitirati: [], distintaVerbali: [], noteUdt: '', dichiarazione: false, createdAt: new Date().toISOString()
  };
}


const emptyAttivitaIspettiva = () => ({ tipo: '', reparto: '', luogo: '', orario: '', esito: '', violazioni: '', note: '' });
function baseOfficialReport() {
  return {
    data: today(), turno: '1° turno', ufficiale: '', qualifica: '', briefing: '', assenti: '', ritardi: '', noteGenerali: '',
    eventiManuali: '', anomalie: '', attivitaIspettive: [emptyAttivitaIspettiva()], esiti: '', comunicazioneEq: '', notaComandante: ''
  };
}


const REPORT_DRAFT_KEY = 'reportTurnoPoliziaLocale_draft_operatore_v1';

function loadReportDraft() {
  if (typeof window === 'undefined') return baseReport();
  try {
    const saved = window.localStorage.getItem(REPORT_DRAFT_KEY);
    if (!saved) return baseReport();
    const parsed = JSON.parse(saved);
    const fresh = baseReport();
    return {
      ...fresh,
      ...parsed,
      counters: { ...emptyCounters(), ...(parsed.counters || {}) },
      operatori: Array.isArray(parsed.operatori) && parsed.operatori.length ? parsed.operatori : fresh.operatori,
      veicoli: Array.isArray(parsed.veicoli) && parsed.veicoli.length ? parsed.veicoli : fresh.veicoli,
      interventi: Array.isArray(parsed.interventi) && parsed.interventi.length ? parsed.interventi : fresh.interventi,
      documentiRitirati: Array.isArray(parsed.documentiRitirati) ? parsed.documentiRitirati : [],
      distintaVerbali: Array.isArray(parsed.distintaVerbali) ? parsed.distintaVerbali : []
    };
  } catch (e) {
    console.warn('Bozza report non leggibile, avvio nuovo report.', e);
    return baseReport();
  }
}

function saveReportDraft(report) {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(REPORT_DRAFT_KEY, JSON.stringify({ ...report, draftSavedAt: new Date().toISOString() }));
  } catch (e) {
    console.warn('Salvataggio automatico non riuscito.', e);
  }
}

function clearReportDraft() {
  if (typeof window === 'undefined') return;
  window.localStorage.removeItem(REPORT_DRAFT_KEY);
}

function LoginScreen({ onLogin }) {
  const [matricola, setMatricola] = useState('');
  const [error, setError] = useState('');
  function submit(e) {
    e.preventDefault();
    const persona = findPersonaByMatricola(matricola);
    if (!persona) { setError('Matricola non riconosciuta. Verificare il numero inserito.'); return; }
    onLogin({ persona, ruolo: matricola.trim() === '9654' ? 'admin' : (isUfficiale(persona) ? 'ufficiale' : 'operatore') });
  }
  return <main><img id="pdfLogo" src="/POLIZIA.png" alt="Logo Polizia Locale" style={{ display: 'none' }} />
    <section className="loginCard"><div className="loginBrand"><img src="/POLIZIA.png" alt="Polizia Locale" /><div><p className="eyebrow">Comune di Monza</p><h1>Report Turno</h1><p>Accesso riservato al personale di Polizia Locale.</p></div></div>
    <form onSubmit={submit} className="loginForm"><Field label="Matricola"><Input value={matricola} onChange={setMatricola} placeholder="Inserisci la tua matricola" /></Field>{error && <p className="errorText">{error}</p>}<button className="primary" type="submit">Accedi</button></form><p className="muted">Protezione base tramite matricola personale. Per l'uso istituzionale definitivo sarà necessaria validazione IT/DPO.</p></section></main>;
}

function App() {
  const [auth, setAuth] = useState(() => { try { return JSON.parse(window.localStorage.getItem('reportPL_auth') || 'null'); } catch { return null; } });
  const [mode, setMode] = useState(() => (auth?.ruolo === 'ufficiale' || auth?.ruolo === 'admin') ? 'dashboard' : 'operatore');
  const [report, setReport] = useState(loadReportDraft);
  const [lastSaved, setLastSaved] = useState('');
  const [importedReports, setImportedReports] = useState([]);
  const [officialReport, setOfficialReport] = useState(baseOfficialReport());
  useEffect(() => { if (!auth?.persona) return; const persona = auth.persona; const nominativo = fullNamePersona(persona); setReport(prev => { const ops = Array.isArray(prev.operatori) ? [...prev.operatori] : []; const has = ops.some(o => String(o.matricola) === persona.matricola); const blank = ops.findIndex(o => !o.nome && !o.matricola && !o.qualifica); if (!has) { const op = { nome: nominativo, matricola: persona.matricola, qualifica: persona.qualifica }; if (blank >= 0) ops[blank] = op; else ops.unshift(op); } return { ...prev, operatori: ops.length ? ops : [{ nome: nominativo, matricola: persona.matricola, qualifica: persona.qualifica }] }; }); if (auth?.ruolo === 'admin' || isUfficiale(persona)) setOfficialReport(prev => ({ ...prev, ufficiale: prev.ufficiale || nominativo, qualifica: prev.qualifica || persona.qualifica })); }, [auth]);
  useEffect(() => { saveReportDraft(report); setLastSaved(new Date().toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' })); }, [report]);
  function handleLogin(nextAuth) { window.localStorage.setItem('reportPL_auth', JSON.stringify(nextAuth)); setAuth(nextAuth); setMode((nextAuth.ruolo === 'ufficiale' || nextAuth.ruolo === 'admin') ? 'dashboard' : 'operatore'); }
  function logout() { window.localStorage.removeItem('reportPL_auth'); setAuth(null); setMode('operatore'); }
  function resetOperatorReport() { const ok = window.confirm('Vuoi iniziare un nuovo report? La bozza salvata su questo dispositivo verrà cancellata. Prima di procedere, scarica PDF e JSON se il turno è concluso.'); if (!ok) return; clearReportDraft(); const fresh = baseReport(); if (auth?.persona) fresh.operatori = [{ nome: fullNamePersona(auth.persona), matricola: auth.persona.matricola, qualifica: auth.persona.qualifica }]; setReport(fresh); }
  if (!auth) return <LoginScreen onLogin={handleLogin} />;
  const ufficiale = auth.ruolo === 'ufficiale' || auth.ruolo === 'admin';
  return <main><img id="pdfLogo" src="/POLIZIA.png" alt="Logo Polizia Locale" style={{ display: 'none' }} /><header className="hero"><div><p className="eyebrow">Polizia Locale</p><h1>Report Turno</h1><p>Accesso: <strong>{fullNamePersona(auth.persona)}</strong> — {auth.persona.qualifica}</p></div><nav className="tabs"><button className={mode === 'operatore' ? 'active' : ''} onClick={() => setMode('operatore')}>Report operatore</button>{ufficiale && <button className={mode === 'dashboard' ? 'active' : ''} onClick={() => setMode('dashboard')}>Dashboard ufficiale</button>}{ufficiale && <button className={mode === 'ufficiale' ? 'active' : ''} onClick={() => setMode('ufficiale')}>Report ufficiale</button>}<button className="ghost" onClick={logout}>Esci</button></nav></header>{mode === 'operatore' && <OperatorReport report={report} setReport={setReport} lastSaved={lastSaved} resetReport={resetOperatorReport} />}{mode === 'dashboard' && ufficiale && <Dashboard reports={importedReports} setReports={setImportedReports} />}{mode === 'ufficiale' && ufficiale && <OfficialReport reports={importedReports} setReports={setImportedReports} official={officialReport} setOfficial={setOfficialReport} />}</main>;
}

function OperatorReport({ report, setReport, lastSaved, resetReport }) {
  const [dbSaving, setDbSaving] = useState(false);
  const update = (patch) => setReport(prev => ({ ...prev, ...patch }));
  const updateArray = (key, index, patch) => setReport(prev => ({ ...prev, [key]: prev[key].map((x, i) => i === index ? { ...x, ...patch } : x) }));
  const addArray = (key, item) => setReport(prev => ({ ...prev, [key]: [...prev[key], item] }));
  const removeArray = (key, index) => setReport(prev => ({ ...prev, [key]: prev[key].filter((_, i) => i !== index) }));

  const totalKm = useMemo(() => report.veicoli.reduce((sum, v) => sum + km(v), 0), [report.veicoli]);
  const totaleViolazioni = useMemo(() => getTotaleViolazioni(report), [report]);
  const text = useMemo(() => reportText(report), [report, totalKm, totaleViolazioni]);

  function generatePdf() {
    printServiceReport(report);
  }

  function exportJson() {
    const payload = { ...report, schemaVersion: 3, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `dati-report-${sanitizeFileName(report.data)}-${sanitizeFileName(turnoLabel(report))}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }
  async function saveToDatabase() {
  try {
    setDbSaving(true);

    const times = getShiftTimes(report);

    const { data: savedReport, error: reportError } = await supabase
      .from('reports')
      .insert([
        {
          command_id: MONZA_COMMAND_ID,
          service_date: report.data,
          start_time: times.start_time,
          end_time: times.end_time,
          status: 'inviato',
          notes: JSON.stringify(report),
        },
      ])
      .select()
      .single();

    if (reportError) {
      console.error(reportError);
      alert('Errore durante il salvataggio del report.');
      return;
    }

    const interventionsToInsert = (report.interventi || []).map(i => ({
      report_id: savedReport.id,
      intervention_time: normalizeTime(i.oraInizio),
      location: i.luogo || null,
      description: `${i.tipo || ''} - ${i.descrizione || ''}`.trim(),
      outcome: i.esito || null,
      notes: JSON.stringify(i),
    }));

    if (interventionsToInsert.length > 0) {
      const { error: interventionsError } = await supabase
        .from('interventions')
        .insert(interventionsToInsert);

      if (interventionsError) {
        console.error(interventionsError);
        alert('Report salvato, ma errore nel salvataggio degli interventi.');
        return;
      }
    }

    alert('Report salvato correttamente nel database.');
  } catch (err) {
    console.error(err);
    alert('Errore imprevisto durante il salvataggio.');
  } finally {
    setDbSaving(false);
  }
}

  function sendMail() {
    const subject = encodeURIComponent(`Report turno Polizia Locale - ${report.data} - ${turnoLabel(report)}`);
    const body = encodeURIComponent(`Si trasmette il report del turno di servizio.\n\nAllegare il PDF scaricato dall'app e, per la dashboard dell'ufficiale, anche il file dati JSON.\n\n${text.slice(0, 1200)}${text.length > 1200 ? '\n\n[Report completo in allegato PDF]' : ''}`);
    window.location.href = `mailto:${encodeURIComponent(report.destinatario)}?subject=${subject}&body=${body}`;
  }

  return <>
    <section className="card notice">
      <h2>Flusso operativo</h2>
      <p>Il report viene <strong>salvato automaticamente su questo dispositivo</strong> mentre viene compilato. L'operatore può inserirlo durante il turno, chiudere l'app e ritrovarlo alla riapertura.</p>
      <p className="muted">Ultimo salvataggio automatico: <strong>{lastSaved || 'in corso'}</strong></p>
      <div className="actions"><button className="ghost" onClick={resetReport}>Nuovo turno / cancella bozza</button></div>
      <p>Al termine del turno l'operatore scarica <strong>PDF</strong> e <strong>file dati JSON</strong>, poi invia entrambi all'ufficiale. L'ufficiale carica i JSON nella dashboard e genera il report aggregato per il Comandante.</p>
    </section>

    <section className="card">
      <h2>1. Dati turno</h2>
      <div className="grid">
        <Field label="Data servizio"><Input type="date" value={report.data} onChange={v => update({ data: v })} /></Field>
        <Field label="Turno"><Select value={report.turno} onChange={v => update({ turno: v })}>{TURNI.map(t => <option key={t}>{t}</option>)}</Select></Field>
        <Field label="Tipologia orario"><Select value={report.orarioTipo} onChange={v => update({ orarioTipo: v })}><option>Ordinario</option><option>Straordinario</option><option>Conto terzi</option></Select></Field>
        <Field label="Reparto"><Select value={report.reparto} onChange={v => update({ reparto: v })}>{REPARTI.map(r => <option key={r}>{r}</option>)}</Select></Field>
      </div>
      {report.turno === 'Altro orario' && <div className="grid two"><Field label="Ora inizio"><Input value={report.altroTurnoInizio} onChange={v => update({ altroTurnoInizio: v })} placeholder="es. 10.00" /></Field><Field label="Ora fine"><Input value={report.altroTurnoFine} onChange={v => update({ altroTurnoFine: v })} placeholder="es. 17.00" /></Field></div>}
      {report.reparto === 'Altri servizi' && <Field label="Specificare altro servizio"><Input value={report.altroServizio} onChange={v => update({ altroServizio: v })} /></Field>}
      {richiedeZonaServizio(report.reparto) && <Field label="Zona di servizio"><Input value={report.zonaServizio || ''} onChange={v => update({ zonaServizio: v })} placeholder="es. Zona A, Presidio città, Zona B/C" /></Field>}
    </section>

    <section className="card">
      <h2>2. Operatori</h2>
      {report.operatori.map((op, idx) => <div className="rowCard" key={idx}><div className="grid three"><Field label="Nome e cognome"><Input value={op.nome} onChange={v => updateArray('operatori', idx, { nome: v })} /></Field><Field label="Matricola"><Input value={op.matricola} onChange={v => updateArray('operatori', idx, { matricola: v })} /></Field><Field label="Qualifica"><Input value={op.qualifica} onChange={v => updateArray('operatori', idx, { qualifica: v })} /></Field></div><button className="ghost" onClick={() => removeArray('operatori', idx)}>Rimuovi</button></div>)}
      <button onClick={() => addArray('operatori', emptyOperatore())}>+ Aggiungi operatore</button>
    </section>

    <section className="card">
      <h2>3. Veicoli e chilometraggio</h2>
      {report.veicoli.map((v, idx) => <div className="rowCard" key={idx}>
        <div className="grid four">
          <Field label="Veicolo / sigla"><Input value={v.sigla} onChange={x => updateArray('veicoli', idx, { sigla: x })} /></Field>
          <Field label="Km inizio"><Input type="number" value={v.kmInizio} onChange={x => updateArray('veicoli', idx, { kmInizio: x })} /></Field>
          <Field label="Km fine"><Input type="number" value={v.kmFine} onChange={x => updateArray('veicoli', idx, { kmFine: x })} /></Field>
          <Field label="Km percorsi"><input readOnly value={km(v)} /></Field>
        </div>
        <div className="grid four">
          <Field label="Effettuato carburante"><Select value={v.carburante || 'No'} onChange={x => updateArray('veicoli', idx, { carburante: x })}><option>No</option><option>Sì</option></Select></Field>
          <Field label="Importo carburante"><Input value={v.importoCarburante || ''} onChange={x => updateArray('veicoli', idx, { importoCarburante: x })} placeholder="es. 50,00 €" /></Field>
          <Field label="Prelevata card C.O. alle ore"><Input value={v.oraPrelievoCard || ''} onChange={x => updateArray('veicoli', idx, { oraPrelievoCard: x })} placeholder="es. 08.10" /></Field>
          <Field label="Restituzione card alle ore"><Input value={v.oraRestituzioneCard || ''} onChange={x => updateArray('veicoli', idx, { oraRestituzioneCard: x })} placeholder="es. 12.55" /></Field>
        </div>
        <Field label="Segnalazione anomalie o danni veicolo"><Textarea value={v.anomaliaVeicolo || ''} onChange={x => updateArray('veicoli', idx, { anomaliaVeicolo: x })} placeholder="Descrivere eventuali anomalie, danni, malfunzionamenti o necessità di manutenzione." /></Field>
        <button className="ghost" onClick={() => removeArray('veicoli', idx)}>Rimuovi</button>
      </div>)}
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
        {['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr'].map(key => <Counter key={key} label={LABELS[key]} value={report.counters[key]} onChange={v => update({ counters: { ...report.counters, [key]: v } })} />)}
      </div>
      <div className="grid two"><Counter label="Altri atti" value={report.counters.altriAttiNumero} onChange={v => update({ counters: { ...report.counters, altriAttiNumero: v } })} /><Field label="Descrizione altri atti"><Input value={report.counters.altriAttiDescrizione} onChange={v => update({ counters: { ...report.counters, altriAttiDescrizione: v } })} /></Field></div>
    </section>

    <section className="card">
      <h2>6. Violazioni</h2>
      <div className="counterGrid">
        {['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme'].map(key => <Counter key={key} label={LABELS[key]} value={report.counters[key]} onChange={v => update({ counters: { ...report.counters, [key]: v } })} />)}
      </div>
      <Field label="Specificare altre norme"><Input value={report.counters.altreNormeDescrizione} onChange={v => update({ counters: { ...report.counters, altreNormeDescrizione: v } })} /></Field>
      <div className="totalBox">Totale violazioni: <strong>{totaleViolazioni}</strong></div>
    </section>

    <section className="card">
      <h2>7. Documenti ritirati</h2>
      <p className="muted">Compila questa sezione solo se nel turno sono stati ritirati documenti.</p>
      {(report.documentiRitirati || []).length === 0 && <p className="muted">Nessun documento ritirato inserito.</p>}
      {(report.documentiRitirati || []).map((doc, idx) => <div className="rowCard" key={idx}>
        <div className="grid three">
          <Field label="Tipo documento"><Select value={doc.tipo} onChange={v => updateArray('documentiRitirati', idx, { tipo: v })}><option value="">Seleziona</option><option>Patente</option><option>Carta di circolazione</option><option>Documento assicurativo</option><option>Autorizzazione / licenza</option><option>Altro</option></Select></Field>
          <Field label="Quantità"><Input type="number" value={doc.quantita} onChange={v => updateArray('documentiRitirati', idx, { quantita: v })} /></Field>
          <Field label="Note"><Input value={doc.note} onChange={v => updateArray('documentiRitirati', idx, { note: v })} /></Field>
        </div>
        <button className="ghost" onClick={() => removeArray('documentiRitirati', idx)}>Rimuovi documento</button>
      </div>)}
      <button onClick={() => addArray('documentiRitirati', emptyDocumentoRitirato())}>+ Aggiungi documento ritirato</button>
    </section>

    <section className="card">
      <h2>8. Distinta verbali</h2>
      <p className="muted">La distinta viene generata automaticamente dai numeri inseriti nella sezione “Violazioni”. Non serve compilare i singoli verbali uno per uno.</p>
      <div className="totalBox">Totale distinta: <strong>{totaleViolazioni}</strong></div>
      <div className="actions"><button onClick={() => buildVerbaliPdf(report).save(`distinta-verbali-${sanitizeFileName(report.data)}-${sanitizeFileName(turnoLabel(report))}.pdf`)}>Scarica PDF distinta verbali</button></div>
    </section>

    <section className="card">
      <h2>9. Note e invio</h2>
      <Field label="Note per UDT / Ufficiale di coordinamento"><Textarea value={report.noteUdt} onChange={v => update({ noteUdt: v })} /></Field>
      <Field label="Email ufficiale destinatario"><Input value={report.destinatario} onChange={v => update({ destinatario: v })} placeholder="es. ufficiale@comune.monza.it" /></Field>
<div className="actions">
  <button onClick={generatePdf}>Apri report stampabile</button>
  <button onClick={exportJson}>Scarica file dati JSON</button>
  <button onClick={saveToDatabase} disabled={dbSaving}>
    {dbSaving ? 'Salvataggio...' : 'Salva su database'}
  </button>
  <button className="primary" onClick={sendMail}>Invia email precompilata</button>
</div>
    </section>

  </>;
}

function Intervento({ i, idx, updateIntervento, remove }) {
  const updateScuola = (sidx, patch) => updateIntervento({ scuole: i.scuole.map((s, n) => n === sidx ? { ...s, ...patch } : s) });
  const addScuola = () => { if (i.scuole.length < 3) updateIntervento({ scuole: [...i.scuole, emptyScuola()] }); };
  const removeScuola = (sidx) => updateIntervento({ scuole: i.scuole.filter((_, n) => n !== sidx) });
  const resetCds = {
    cdsDettaglio: '', cdsRimozione: 'No', cdsMotivazione: '', cdsRipristino: 'No', cdsFeriti: 'No', cdsVeicoliCoinvolti: '',
    cdsVerificaEffettuata: 'No', cdsSegnalazione: 'No', cdsPericolo: 'No', cdsInterventoRichiesto: 'No', cdsStatoSemaforo: 'spento'
  };
  return <div className="intervento">
    <div className="interventoHead"><h3>Intervento {idx + 1}</h3><button className="ghost" onClick={remove}>Rimuovi</button></div>
    <div className="grid three">
      <Field label="Tipo intervento"><Select value={i.tipo} onChange={v => updateIntervento({ tipo: v, ...(v !== 'Codice della strada' ? resetCds : {}) })}>{TIPI_INTERVENTO.map(t => <option key={t}>{t}</option>)}</Select></Field>
      <Field label="Origine"><Select value={i.origine} onChange={v => updateIntervento({ origine: v })}>{ORIGINI.map(o => <option key={o}>{o}</option>)}</Select></Field>
      <Field label="Luogo"><Input value={i.luogo} onChange={v => updateIntervento({ luogo: v })} /></Field>
    </div>
    {i.tipo === 'Codice della strada' && <div className="schoolBox"><h4>Dettaglio Codice della strada</h4>
      <div className="grid three">
        <Field label="Dettaglio intervento"><Select value={i.cdsDettaglio || ''} onChange={v => updateIntervento({ cdsDettaglio: v })}><option value="">Seleziona</option>{DETTAGLI_CODICE_STRADA.map(d => <option key={d}>{d}</option>)}</Select></Field>
        {i.cdsDettaglio === 'Controllo soste' && <><Field label="Rimozione veicolo"><Select value={i.cdsRimozione || 'No'} onChange={v => updateIntervento({ cdsRimozione: v })}><option>No</option><option>Sì</option></Select></Field><Field label="Motivazione"><Input value={i.cdsMotivazione || ''} onChange={v => updateIntervento({ cdsMotivazione: v })} placeholder="es. intralcio, passo carrabile, area mercato..." /></Field></>}
        {i.cdsDettaglio === 'Buca su carreggiata' && <Field label="Richiesto intervento per ripristino"><Select value={i.cdsRipristino || 'No'} onChange={v => updateIntervento({ cdsRipristino: v })}><option>No</option><option>Sì</option></Select></Field>}
        {i.cdsDettaglio === 'Sinistro stradale' && <><Field label="Feriti"><Select value={i.cdsFeriti || 'No'} onChange={v => updateIntervento({ cdsFeriti: v })}><option>No</option><option>Sì</option></Select></Field><Field label="Veicoli coinvolti"><Input type="number" value={i.cdsVeicoliCoinvolti || ''} onChange={v => updateIntervento({ cdsVeicoliCoinvolti: v })} /></Field></>}
        {i.cdsDettaglio === 'Veicolo sospetto' && <><Field label="Verifica effettuata"><Select value={i.cdsVerificaEffettuata || 'No'} onChange={v => updateIntervento({ cdsVerificaEffettuata: v })}><option>No</option><option>Sì</option></Select></Field><Field label="Segnalazione"><Select value={i.cdsSegnalazione || 'No'} onChange={v => updateIntervento({ cdsSegnalazione: v })}><option>No</option><option>Sì</option></Select></Field></>}
        {i.cdsDettaglio === 'Segnaletica danneggiata' && <><Field label="Pericolo"><Select value={i.cdsPericolo || 'No'} onChange={v => updateIntervento({ cdsPericolo: v })}><option>No</option><option>Sì</option></Select></Field><Field label="Intervento richiesto"><Select value={i.cdsInterventoRichiesto || 'No'} onChange={v => updateIntervento({ cdsInterventoRichiesto: v })}><option>No</option><option>Sì</option></Select></Field></>}
        {i.cdsDettaglio === 'Guasto semaforo' && <><Field label="Stato"><Select value={i.cdsStatoSemaforo || 'spento'} onChange={v => updateIntervento({ cdsStatoSemaforo: v })}><option>spento</option><option>lampeggiante</option></Select></Field><Field label="Intervento richiesto"><Select value={i.cdsInterventoRichiesto || 'No'} onChange={v => updateIntervento({ cdsInterventoRichiesto: v })}><option>No</option><option>Sì</option></Select></Field></>}
      </div>
    </div>}
    {i.origine === 'Altro' && <Field label="Specificare da chi è arrivata la disposizione"><Input value={i.origineAltro} onChange={v => updateIntervento({ origineAltro: v })} /></Field>}
    <div className="grid two"><Field label="Ora inizio"><Input value={i.oraInizio} onChange={v => updateIntervento({ oraInizio: v })} placeholder="es. 08.15" /></Field><Field label="Ora fine"><Input value={i.oraFine} onChange={v => updateIntervento({ oraFine: v })} placeholder="es. 09.00" /></Field></div>
    <Field label={i.tipo === 'Altro' ? 'Descrizione intervento' : 'Descrizione'}><Textarea value={i.descrizione} onChange={v => updateIntervento({ descrizione: v })} /></Field>
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
    const rows = [['Data','Turno','Orario','Reparto','Zona di servizio','Operatori','Interventi','Violazioni','Km','Note']];
    reports.forEach(r => rows.push([r.data, turnoLabel(r), r.orarioTipo, repartoLabel(r), r.zonaServizio || '', operatorNames(r).join('; '), r.interventi?.length || 0, getTotaleViolazioni(r), getKmTotali(r), r.noteUdt || '']));
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
      <Metric label="Violazioni" value={aggregate.totaleViolazioni} />
      <Metric label="Km percorsi" value={aggregate.kmTotali} />
    </section>

    <section className="card">
      <h2>Quadro riepilogativo</h2>
      {reports.length === 0 ? <p className="muted">Nessun report caricato.</p> : <div className="tableWrap"><table><thead><tr><th>Data</th><th>Turno</th><th>Reparto</th><th>Zona</th><th>Operatori</th><th>Interventi</th><th>Violazioni</th><th>Km</th></tr></thead><tbody>{reports.map((r, idx) => <tr key={idx}><td>{r.data}</td><td>{turnoLabel(r)}<br/><small>{r.orarioTipo}</small></td><td>{repartoLabel(r)}</td><td>{r.zonaServizio || '-'}</td><td>{operatorNames(r).join(', ') || '-'}</td><td>{r.interventi?.length || 0}</td><td>{getTotaleViolazioni(r)}</td><td>{getKmTotali(r)}</td></tr>)}</tbody></table></div>}
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

  </>;
}

function OfficialReport({ reports, setReports, official, setOfficial }) {
  const aggregate = useMemo(() => aggregateReports(reports), [reports]);
  const autoSintesi = useMemo(() => officialSynthesis(aggregate, reports), [aggregate, reports]);
  const autoEventi = useMemo(() => officialEventsText(reports), [reports]);
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
    printOfficialReport(aggregate, reports, official, autoSintesi, autoEventi);
  }

  return <>
    <section className="card notice">
      <h2>Report ufficiale UDT</h2>
      <p>Modalità quasi automatica: carica i JSON degli operatori, verifica la sintesi generata, integra briefing, personale, attività ispettive, anomalie e note per il Comandante.</p>
      <div className="actions"><label className="fileButton">Carica JSON operatori<input type="file" accept="application/json,.json" multiple onChange={importFiles} /></label><button className="ghost" onClick={() => setReports([])}>Svuota dati caricati</button></div>
    </section>
    <section className="metrics"><Metric label="Report operatori" value={reports.length} /><Metric label="Interventi" value={aggregate.totalInterventi} /><Metric label="Violazioni" value={aggregate.totaleViolazioni} /><Metric label="Km" value={aggregate.kmTotali} /></section>
    <section className="card"><h2>1. Dati report ufficiale</h2><div className="grid four"><Field label="Data"><Input type="date" value={official.data} onChange={v => update({ data: v })} /></Field><Field label="Turno"><Input value={official.turno} onChange={v => update({ turno: v })} placeholder="es. 1° turno" /></Field><Field label="Ufficiale di turno"><Input value={official.ufficiale} onChange={v => update({ ufficiale: v })} /></Field><Field label="Qualifica"><Input value={official.qualifica} onChange={v => update({ qualifica: v })} placeholder="es. Commissario Capo" /></Field></div></section>
    <section className="card"><h2>2. Sintesi automatica</h2><p className="muted">Questa sintesi nasce dai report operatori caricati. Nel PDF viene riportata come quadro iniziale.</p><pre className="miniPreview">{autoSintesi}</pre><Field label="Integrazioni dell'ufficiale alla sintesi"><Textarea value={official.eventiManuali} onChange={v => update({ eventiManuali: v })} placeholder="Inserire eventuali elementi aggiuntivi non presenti nei report operatori..." /></Field></section>
    <section className="card"><h2>3. Briefing, personale e note</h2><div className="grid two"><Field label="Briefing operativo"><Input value={official.briefing} onChange={v => update({ briefing: v })} placeholder="es. 06.45" /></Field><Field label="Note generali"><Input value={official.noteGenerali} onChange={v => update({ noteGenerali: v })} placeholder="es. Con il personale a disposizione coperte 11 scuole" /></Field></div><div className="grid two"><Field label="A.P.L. assenti"><Textarea value={official.assenti} onChange={v => update({ assenti: v })} /></Field><Field label="A.P.L. in ritardo"><Textarea value={official.ritardi} onChange={v => update({ ritardi: v })} /></Field></div></section>
    <section className="card"><h2>4. Eventi degni di rilievo</h2><p className="muted">Eventi rilevanti individuati automaticamente: sinistri con feriti, TSO/ASO, interventi con parole chiave critiche o lunga durata.</p><pre className="miniPreview">{autoEventi || 'Nessun evento rilevante automatico rilevato.'}</pre></section>
    <section className="card"><h2>5. Anomalie e attività ispettive</h2><Field label="Anomalie riscontrate durante il turno"><Textarea value={official.anomalie} onChange={v => update({ anomalie: v })} /></Field><h3>Attività ispettive</h3>{official.attivitaIspettive.map((a, idx) => <div className="rowCard" key={idx}><div className="grid four"><Field label="Tipo attività"><Input value={a.tipo} onChange={v => updateAttivita(idx, { tipo: v })} placeholder="es. annonaria, ambiente..." /></Field><Field label="Reparto / pattuglia"><Input value={a.reparto} onChange={v => updateAttivita(idx, { reparto: v })} /></Field><Field label="Luogo"><Input value={a.luogo} onChange={v => updateAttivita(idx, { luogo: v })} /></Field><Field label="Orario"><Input value={a.orario} onChange={v => updateAttivita(idx, { orario: v })} /></Field></div><div className="grid three"><Field label="Esito"><Input value={a.esito} onChange={v => updateAttivita(idx, { esito: v })} /></Field><Field label="Violazioni collegate"><Input value={a.violazioni} onChange={v => updateAttivita(idx, { violazioni: v })} /></Field><Field label="Note"><Input value={a.note} onChange={v => updateAttivita(idx, { note: v })} /></Field></div><button className="ghost" onClick={() => removeAttivita(idx)}>Rimuovi attività</button></div>)}<button onClick={addAttivita}>+ Aggiungi attività ispettiva</button></section>
    <section className="card"><h2>6. Esiti e comunicazioni</h2><Field label="Esiti"><Textarea value={official.esiti} onChange={v => update({ esiti: v })} /></Field><div className="grid two"><Field label="Comunicazione all'E.Q. di turno"><Textarea value={official.comunicazioneEq} onChange={v => update({ comunicazioneEq: v })} /></Field><Field label="Nota per il Comandante"><Textarea value={official.notaComandante} onChange={v => update({ notaComandante: v })} /></Field></div><div className="actions"><button className="primary" onClick={generateOfficialPdf}>Apri report ufficiale stampabile</button></div></section>
  </>;
}

function Metric({ label, value }) { return <div className="metric"><strong>{value}</strong><span>{label}</span></div>; }
function Distribution({ title, data, labels = {} }) {
  const entries = Object.entries(data || {}).filter(([, value]) => n(value) > 0).sort((a, b) => n(b[1]) - n(a[1]));
  return <div><h3>{title}</h3>{entries.length === 0 ? <p className="muted">Nessun dato.</p> : <ul className="distList">{entries.map(([key, value]) => <li key={key}><span>{labels[key] || key}</span><strong>{value}</strong></li>)}</ul>}</div>;
}


// ===== REPORT PROFESSIONALI STAMPABILI HTML/CSS =====
// Questa sezione sostituisce il PDF disegnato a coordinate: apre una pagina HTML stampabile,
// stabile, allineata e facilmente salvabile in PDF dal browser.
function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[ch]));
}
function safeText(value, fallback = '-') { const s = String(value ?? '').trim(); return s || fallback; }
function splitList(text) { return String(text || '').split(/\n|;/).map(x => x.trim()).filter(Boolean); }
function statusList(label, count, names, cls) {
  const items = Array.isArray(names) ? names : splitList(names);
  return `<div class="status-row"><span>${esc(label)}</span><strong class="${cls}">${esc(count)}</strong></div>${items.length ? `<ul class="compact-list ${cls}">${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}`;
}
function logoHtml() { return `<img class="brand-logo" src="/POLIZIA.png" alt="Polizia Locale Monza" />`; }
function iconSvg(type) {
  const common = 'width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"';
  const icons = {
    car: `<svg ${common}><path d="M3 13l2-5a3 3 0 0 1 3-2h8a3 3 0 0 1 3 2l2 5"/><path d="M5 13h14v5H5z"/><circle cx="7.5" cy="18" r="1.5"/><circle cx="16.5" cy="18" r="1.5"/></svg>`,
    doc: `<svg ${common}><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v5h5"/><path d="M9 13h6"/><path d="M9 17h6"/></svg>`,
    warn: `<svg ${common}><path d="M12 3l10 18H2z"/><path d="M12 9v5"/><path d="M12 18h.01"/></svg>`,
    clip: `<svg ${common}><path d="M9 12l6-6a3 3 0 0 1 4 4l-8 8a5 5 0 0 1-7-7l8-8"/></svg>`,
    users: `<svg ${common}><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    clipboard: `<svg ${common}><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4a3 3 0 0 1 6 0"/><path d="M9 11h6"/><path d="M9 15h6"/></svg>`,
    table: `<svg ${common}><path d="M3 4h18v16H3z"/><path d="M3 10h18"/><path d="M9 4v16"/><path d="M15 4v16"/></svg>`,
    fuel: `<svg ${common}><path d="M4 3h10v18H4z"/><path d="M14 8h2l3 3v7a2 2 0 0 0 4 0v-6"/><path d="M7 7h4"/></svg>`,
    check: `<svg ${common}><path d="M20 6L9 17l-5-5"/></svg>`,
    user: `<svg ${common}><circle cx="12" cy="7" r="4"/><path d="M5 21v-2a7 7 0 0 1 14 0v2"/></svg>`
  };
  return icons[type] || icons.doc;
}
function printShell(title, pagesHtml) {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>
    :root{--blue:#0d2b57;--line:#cfd8e3;--soft:#f6f8fb;--text:#111827;--muted:#64748b;--orange:#ea7a1a;--red:#dc2626;--green:#087a47;}
    *{box-sizing:border-box} body{margin:0;background:#e5e7eb;color:var(--text);font-family:Arial,Helvetica,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    .toolbar{position:sticky;top:0;z-index:10;padding:12px;text-align:center;background:#111827;color:#fff;box-shadow:0 2px 10px #0002}.toolbar button{background:#fff;border:0;border-radius:8px;padding:10px 18px;font-weight:700;cursor:pointer}.toolbar span{margin-left:14px;color:#d1d5db;font-size:13px}
    .page{width:297mm;height:210mm;margin:14px auto;background:white;position:relative;padding:8mm 9mm 9mm;overflow:hidden;box-shadow:0 3px 18px #0002;page-break-after:always;}
    .header{height:26mm;display:grid;grid-template-columns:25mm 1fr 48mm;gap:8mm;align-items:start;border-bottom:1.6px solid var(--blue);padding-bottom:4mm}.brand-logo{width:22mm;height:22mm;object-fit:contain}.brand h1{font-size:16pt;line-height:1.05;margin:2mm 0 0;color:var(--blue);font-weight:800}.brand h2{font-size:13pt;line-height:1.1;margin:1mm 0;color:var(--blue);font-weight:800}.brand p,.contacts p{margin:1.3mm 0 0;font-size:7.5pt;font-weight:700}.contacts{text-align:left;font-size:7.3pt;line-height:1.25;color:#111827;padding-top:2mm}.titlebar{margin:5mm 0 4mm;height:9mm;background:var(--blue);color:#fff;display:grid;grid-template-columns:1fr auto;align-items:center;padding:0 6mm;font-weight:800;font-size:12pt;letter-spacing:.2px}.titlebar .subtitle{font-size:9pt;font-weight:700}
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:3mm;margin-bottom:4mm}.kpi{height:25mm;border:1px solid var(--line);border-radius:2mm;display:grid;grid-template-columns:13mm 1fr;grid-template-rows:9mm 1fr;align-items:center;padding:3mm;background:#fff}.kpi .ico{grid-row:1/3;color:var(--blue);display:flex;align-items:center;justify-content:center}.kpi .label{font-size:7.2pt;font-weight:800;color:var(--blue);text-transform:uppercase;text-align:center}.kpi .num{font-size:22pt;line-height:1;font-weight:900;color:var(--blue);text-align:center}.kpi .sub{font-size:6.8pt;color:#111;text-align:center;margin-top:1mm}.grid-2{display:grid;grid-template-columns:1fr 1.2fr;gap:5mm}.grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:5mm}.panel{border:1px solid var(--line);border-radius:1.2mm;background:#fff;overflow:hidden;margin-bottom:4mm}.panel-title{height:7mm;background:var(--blue);color:#fff;font-size:8.2pt;font-weight:800;text-transform:uppercase;display:flex;align-items:center;gap:2mm;padding:0 4mm}.panel-title svg{width:14px;height:14px}.panel-body{padding:4mm;font-size:8.3pt;line-height:1.42}.panel-body.tight{padding:3mm}.status-row{display:grid;grid-template-columns:1fr 16mm;align-items:center;border-bottom:1px solid #e5e7eb;padding:1.6mm 0;font-size:8.2pt}.status-row strong{text-align:center;font-size:10.5pt}.green{color:var(--green)!important}.orange{color:var(--orange)!important}.red{color:var(--red)!important}.blue{color:var(--blue)!important}.compact-list{margin:1mm 0 2mm 5mm;padding:0;font-size:7.3pt}.compact-list li{margin:0.7mm 0}.eventbox{border:1.4px solid #f4a261;background:#fff8f1;border-radius:1.2mm;margin-top:4mm;min-height:18mm;display:grid;grid-template-columns:16mm 1fr;align-items:center;padding:3mm;color:#111827}.eventbox .eventico{color:var(--orange);display:flex;justify-content:center}.eventbox h3{margin:0 0 1.5mm;color:#d95f02;font-size:9pt;text-transform:uppercase}.eventbox p{margin:0;font-size:8.3pt}.table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:7.3pt}.table th{background:var(--blue);color:#fff;padding:2mm 1.4mm;text-align:center;font-size:7pt}.table td{border:1px solid #d7dee8;padding:1.8mm 1.3mm;vertical-align:top;overflow:hidden;word-wrap:break-word}.table tbody tr:nth-child(even) td{background:#f8fafc}.table .num{text-align:center;font-weight:700}.table .total td{background:#0d2b57!important;color:#fff;font-weight:800}.small-list{display:grid;gap:1.5mm}.small-row{display:grid;grid-template-columns:1fr 14mm;gap:3mm}.small-row strong{text-align:right}.footer{position:absolute;left:9mm;right:9mm;bottom:5mm;height:8mm;border-top:1px solid #cbd5e1;display:grid;grid-template-columns:1fr 1fr auto;align-items:center;font-size:6.8pt;color:var(--blue);font-weight:700}.signature{font-family:Georgia,serif;font-style:italic;font-size:13pt;text-align:center;margin-top:4mm}.detail-grid{display:grid;grid-template-columns:1.1fr .9fr;gap:5mm}.detail-grid-2{display:grid;grid-template-columns:.92fr 1.08fr;gap:5mm}.bullet-list{margin:0;padding-left:4mm}.bullet-list li{margin:1.4mm 0}.center{text-align:center}.right{text-align:right}.muted{color:var(--muted)}
    @media print{body{background:white}.toolbar{display:none}.page{margin:0;box-shadow:none;page-break-after:always}@page{size:A4 landscape;margin:0}}
  </style></head><body><div class="toolbar"><button onclick="window.print()">Stampa / Salva in PDF</button><span>Imposta orientamento: Orizzontale. Disattiva intestazioni/piè di pagina del browser.</span></div>${pagesHtml}</body></html>`;
}
function footerHtml(page, total=2) { return `<div class="footer"><span>Settore Polizia Locale, Protezione Civile</span><span>Via Marsala 13 | 20900 Monza &nbsp;&nbsp; Tel. 039 28161 &nbsp;&nbsp; polizialocale@comune.monza.it</span><span>Pag. ${page} di ${total}</span></div>`; }
function headerHtml(title, subtitle) { return `<div class="header">${logoHtml()}<div class="brand"><h1>COMUNE DI MONZA</h1><h2>Polizia Locale</h2><p>Settore Polizia Locale e Protezione Civile</p></div><div class="contacts"><p>Via Marsala 13</p><p>20900 Monza</p><p>Tel. 039 28161</p><p>polizialocale@comune.monza.it</p></div></div><div class="titlebar"><span>${esc(title)}</span><span class="subtitle">${esc(subtitle)}</span></div>`; }
function kpiBox(icon, label, value) { return `<div class="kpi"><div class="ico">${iconSvg(icon)}</div><div class="label">${esc(label)}</div><div><div class="num">${esc(value)}</div><div class="sub">Totali</div></div></div>`; }
function panel(title, icon, body, extraClass='') { return `<section class="panel ${extraClass}"><div class="panel-title">${iconSvg(icon)}<span>${esc(title)}</span></div><div class="panel-body">${body}</div></section>`; }
function openPrintWindow(html) { const w = window.open('', '_blank'); if (!w) { alert('Popup bloccato: consenti le finestre popup per stampare il report.'); return; } w.document.open(); w.document.write(html); w.document.close(); setTimeout(() => { try { w.focus(); } catch(e) {} }, 300); }
function printServiceReport(report) { openPrintWindow(buildServicePrintHtml(report)); }
function printOfficialReport(aggregate, reports, official, autoSintesi, autoEventi) { openPrintWindow(buildOfficialPrintHtml(aggregate, reports, official, autoSintesi, autoEventi)); }
function buildServicePrintHtml(report) {
  const c = report.counters || emptyCounters();
  const interventions = report.interventi || [];
  const atti = ['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].reduce((s,k)=>s+n(c[k]),0);
  const eventi = interventions.filter(isInterventoCritico).length;
  const subtitle = `${formatDateIT(report.data)} | ${turnoLabel(report)} | ${report.orarioTipo || '-'}`;
  const vehiclesUsed = (report.veicoli||[]).filter(v=>v.sigla||v.kmInizio||v.kmFine).length || '-';
  const vehicleBody = `<div class="small-row"><span>Veicoli impiegati</span><strong>${esc(vehiclesUsed)}</strong></div><div class="small-row"><span>Totale km percorsi</span><strong>${esc(getKmTotali(report))} km</strong></div>`;
  const carburanteBody = (report.veicoli||[]).some(v=>v.carburante==='Sì') ? (report.veicoli||[]).map(v=>`<div class="small-row"><span>${esc(v.sigla||'Veicolo')}</span><strong>${esc(v.importoCarburante||'-')}</strong></div>`).join('') : '<p>Nessun rifornimento indicato.</p>';
  const anomalie = (report.veicoli||[]).map(v=>v.anomaliaVeicolo).filter(Boolean).join('; ') || 'Nessuna anomalia segnalata.';
  const operators = operatorNames(report).join('<br>') || '-';
  const dichiarazioneFinale = `
  <div style="margin-top:20px; font-size:12px;">
    Gli operatori dichiarano che quanto riportato nel presente report corrisponde alle attività effettivamente svolte e riscontrate durante il turno di servizio, consapevoli della proprie responsabilità amministrative e penali anche in considerazione dell'art. 328 C.P.
  </div>
`;
  const interventiHtml = interventions.length ? `<ul class="bullet-list">${interventions.slice(0,8).map(i=>`<li><strong>${esc(i.tipo)}</strong>${i.oraInizio?` — ${esc(i.oraInizio)}`:''}<br><span class="muted">${esc(i.luogo || '')}</span> ${esc(i.descrizione || i.esito || '')}</li>`).join('')}</ul>` : '<p>Nessun intervento inserito.</p>';
  const violazioniRows = [['Codice della Strada', n(c.vdcCds)+n(c.preavvisiCds)], ['Regolamenti comunali', n(c.regPolizia)+n(c.regEdilizio)+n(c.regBenessereAnimali)], ['Annonaria / commercio', n(c.annonaria)], ['Altro', n(c.altreNorme)], ['TOTALE', getTotaleViolazioni(report)]];
  const violazioniTable = `<table class="table"><thead><tr><th>Tipo violazione</th><th style="width:22mm">Nr.</th></tr></thead><tbody>${violazioniRows.map((r,idx)=>`<tr class="${idx===violazioniRows.length-1?'total':''}"><td>${esc(r[0])}</td><td class="num">${esc(r[1])}</td></tr>`).join('')}</tbody></table>`;
  const attiBody = `<div class="small-list">${[['Relazioni',c.relazioni],['Annotazioni',c.annotazioni],['Fermi amm.',c.fermiAmministrativi],['Sequestri amm.',c.sequestriAmministrativi],['Sequestri penali',c.sequestriPenali],['C.N.R.',c.cnr],['Altri atti',c.altriAttiNumero]].map(([l,v])=>`<div class="small-row"><span>${esc(l)}</span><strong>${esc(n(v))}</strong></div>`).join('')}</div>`;
  const docs = report.documentiRitirati || [];
  const docsBody = docs.length ? `<ul class="bullet-list">${docs.map(d=>`<li>${esc(d.tipo || 'Documento')} — ${esc(d.quantita || 1)} ${d.note?`<br><span class="muted">${esc(d.note)}</span>`:''}</li>`).join('')}</ul>` : '<p>Nessun documento ritirato.</p>';
  const page1 = `<section class="page">${headerHtml('REPORT DI SERVIZIO', subtitle)}<div class="kpis">${kpiBox('car','Interventi',interventions.length)}${kpiBox('doc','Violazioni',getTotaleViolazioni(report))}${kpiBox('clip','Atti redatti',atti)}${kpiBox('warn','Eventi',eventi)}</div><div class="grid-3">${panel('Veicoli','car',vehicleBody)}${panel('Carburante','fuel',carburanteBody)}${panel('Anomalie veicolo','warn',`<p>${esc(anomalie)}</p>`)}</div><div class="grid-2">${panel('Note di servizio','clipboard',`<p>${esc(report.noteUdt || '-')}</p>`)}${panel('Operatori','users',`<p><strong>Reparto:</strong> ${esc(repartoLabel(report))}</p><p>${operators}</p>`)}</div>${footerHtml(1)}</section>`;
  const page2 = `<section class="page">${headerHtml('REPORT DI SERVIZIO - DETTAGLIO', subtitle)}<div class="detail-grid">${panel('Interventi effettuati','car',interventiHtml)}${panel('Violazioni contestate','table',violazioniTable,'tight-panel')}</div><div class="detail-grid-2">${panel('Atti redatti','clip',attiBody)}${panel('Osservazioni','clipboard',`<p>${esc(report.osservazioni || report.noteUdt || 'Nessuna osservazione particolare da segnalare.')}</p>`)}</div><div class="detail-grid-2">${panel('Documenti ritirati','doc',docsBody)}${panel('Firma agente','user',`<p><strong>${esc(operatorNames(report)[0] || '-')}</strong></p><div class="signature">Firma</div>`)}</div>${footerHtml(2)}</section>`;
  return printShell('Report di servizio', page1 + page2 + dichiarazioneFinale);
}
function buildOfficialPrintHtml(aggregate, reports, official, autoSintesi, autoEventi) {
  const date = formatDateIT(official.data || aggregate.dateLabel);
  const subtitle = `${date} | ${official.turno || '-'}`;
  const attiObj = attiObjectFromReports(reports);
  const attiTot = totalAttiFromReports(reports);
  const eventiCount = reports.reduce((s,r)=>s+(r.interventi||[]).filter(isInterventoCritico).length,0) + countTextItems(official.eventiManuali);
  const assentiList = splitList(official.assenti);
  const ritardiList = splitList(official.ritardi);
  const presenti = reports.length ? new Set(reports.flatMap(r=>operatorNames(r))).size : '-';
  const personaleBody = `${statusList('Presenti', presenti, [], 'green')}${statusList('Ritardo', ritardiList.length, ritardiList, 'orange')}${statusList('Assenti', assentiList.length, assentiList, 'red')}<div class="status-row"><span>Totale</span><strong class="blue">${esc(presenti === '-' ? '-' : Number(presenti)+ritardiList.length+assentiList.length)}</strong></div>`;
  const eventiText = [autoEventi, official.eventiManuali, official.anomalie].filter(Boolean).join('\n') || 'Nessun evento rilevante automatico rilevato.';
  const page1 = `<section class="page">${headerHtml('REPORT UFFICIALE DI TURNO', subtitle)}<div class="kpis">${kpiBox('car','Interventi',aggregate.totalInterventi)}${kpiBox('doc','Verbali',aggregate.totaleViolazioni)}${kpiBox('warn','Eventi',eventiCount)}${kpiBox('clip','Atti redatti',attiTot)}</div><div class="grid-2">${panel('Personale','users',personaleBody)}${panel('Briefing operativo','clipboard',`<p>${esc(official.briefing || '-')}</p>`)}</div><div class="eventbox"><div class="eventico">${iconSvg('warn')}</div><div><h3>Eventi / anomalie degne di rilievo</h3><p>${esc(eventiText)}</p></div></div>${footerHtml(1)}</section>`;
  const rows = buildViolationRows(reports);
  const tableRows = rows.length ? rows : [['-','-',0,0,0,0,0,0]];
  const violTable = `<table class="table"><thead><tr><th style="width:42mm">Pattuglia</th><th style="width:38mm">Reparto</th><th>Prev.</th><th>C.d.S.</th><th>Urbana</th><th>Annonaria</th><th>Altre</th><th>Tot.</th></tr></thead><tbody>${tableRows.map((r,idx)=>`<tr class="${idx===tableRows.length-1 && rows.length?'total':''}">${r.map((c,i)=>`<td class="${i>=2?'num':''}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const attiBody = `<div class="small-list">${[['Fermi amministrativi',attiObj.fermiAmministrativi],['Sequestri amministrativi',attiObj.sequestriAmministrativi],['Sequestri penali',attiObj.sequestriPenali],['Notizie di reato',attiObj.cnr],['TOTALE',attiTot]].map(([l,v])=>`<div class="small-row"><span>${esc(l)}</span><strong>${esc(n(v))}</strong></div>`).join('')}</div>`;
  const page2 = `<section class="page">${headerHtml('REPORT UFFICIALE DI TURNO - DETTAGLIO', subtitle)}${panel('Violazioni riscontrate','table',violTable)}<div class="detail-grid-2">${panel('Atti redatti','clip',attiBody)}<div>${panel('Esito turno','check',`<p>${esc(official.esiti || '-')}</p>`)}${panel('Comunicazioni E.Q.','doc',`<p>${esc(official.comunicazioneEq || '-')}</p>`)}</div></div><div class="detail-grid-2">${panel('Nota del comandante','clipboard',`<p>${esc(official.notaComandante || '-')}</p>`)}${panel('Responsabile di turno','user',`<p class="center"><strong>${esc(official.qualifica || '-')}</strong></p><p class="center"><strong>${esc(official.ufficiale || '-')}</strong></p><div class="signature">Firma</div>`)}</div>${footerHtml(2)}</section>`;
  return printShell('Report ufficiale di turno', page1 + page2);
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
  doc.text('Tel. 039 28161', 145, 17);
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
    doc.text('Tel. 039 28161', 86, 291);
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
  const atti = ['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].reduce((s, k) => s + n((report.counters || {})[k]), 0);
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

function buildVerbaliPdf(report) {
  const title = 'DISTINTA VIOLAZIONI DEL TURNO';
  const subtitle = `${report.data} | Turno ${turnoLabel(report)} | ${repartoLabel(report)}`;
  const doc = makePdf(title, subtitle);
  const c = report.counters || emptyCounters();
  let y = 54;

  y = section(doc, 'Dati generali', y, title, subtitle);
  y = kvGrid(doc, [
    { label: 'Data servizio', value: report.data },
    { label: 'Turno', value: turnoLabel(report) },
    { label: 'Tipologia orario', value: report.orarioTipo || '-' },
    { label: 'Reparto / servizio', value: repartoLabel(report) },
    ...(report.zonaServizio ? [{ label: 'Zona di servizio', value: report.zonaServizio }] : []),
    { label: 'Operatori', value: operatorNames(report).join(', ') || '-' },
  ], y, 2, title, subtitle) + 2;

  y = section(doc, 'Riepilogo smart attività sanzionatoria', y, title, subtitle);
  y = simpleTable(doc, ['Tipologia', 'N.'], [
    ['Preavvisi CdS', c.preavvisiCds],
    ['Verbali CdS', c.vdcCds],
    ['Verbali Regolamento Polizia', c.regPolizia],
    ['Verbali Regolamento Edilizio', c.regEdilizio],
    ['Verbali Regolamento Benessere Animali', c.regBenessereAnimali],
    ['Verbali Annonaria / commercio', c.annonaria],
    [`Altre violazioni ${c.altreNormeDescrizione || ''}`, c.altreNorme],
    ['TOTALE', getTotaleViolazioni(report)]
  ], y, [150, 36], title, subtitle);

  y = section(doc, 'Note', y, title, subtitle);
  y = paragraph(doc, 'La presente distinta è generata automaticamente dai dati inseriti nella sezione “Violazioni” del report di servizio.', y, title, subtitle);
  addFooter(doc);
  return doc;
}

function buildServicePdf_old(report) {
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
    ...(report.zonaServizio ? [{ label: 'Zona di servizio', value: report.zonaServizio }] : []),
  ], y, 2, title, subtitle) + 2;

  y = section(doc, 'Operatori', y, title, subtitle);
  const operatorRows = (report.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).map(o => [o.nome, o.matricola, o.qualifica]);
  y = simpleTable(doc, ['Nominativo', 'Matricola', 'Qualifica'], operatorRows.length ? operatorRows : [['-', '-', '-']], y, [90, 40, 56], title, subtitle);

  y = section(doc, 'Veicoli, chilometraggio e carburante', y, title, subtitle);
  const vehicleRows = (report.veicoli || []).map(v => [v.sigla || '-', v.kmInizio || '-', v.kmFine || '-', km(v), v.carburante || 'No', v.importoCarburante || '-', v.oraPrelievoCard || '-', v.oraRestituzioneCard || '-']);
  y = simpleTable(doc, ['Veicolo', 'Km inizio', 'Km fine', 'Km', 'Carburante', 'Importo', 'Card presa', 'Card resa'], vehicleRows.length ? vehicleRows : [['-', '-', '-', '-', '-', '-', '-', '-']], y, [34, 23, 23, 18, 24, 22, 21, 21], title, subtitle);
  const anomalieVeicoli = (report.veicoli || []).filter(v => v.anomaliaVeicolo).map(v => `${v.sigla || 'Veicolo'}: ${v.anomaliaVeicolo}`).join('\n');
  y = kvGrid(doc, [{ label: 'Totale km percorsi', value: getKmTotali(report) }, { label: 'Anomalie / danni veicolo', value: anomalieVeicoli || '-' }], y, 1, title, subtitle) + 2;

  y = section(doc, 'Interventi effettuati', y, title, subtitle);
  (report.interventi || []).forEach((i, idx) => {
    y = serviceInterventionCard(doc, i, idx, y, title, subtitle);
  });

  y = section(doc, 'Atti redatti', y, title, subtitle);
  const c = report.counters || emptyCounters();
  y = simpleTable(doc, ['Tipologia', 'N.'], [
    ['Relazioni di servizio', c.relazioni], ['Annotazioni di servizio', c.annotazioni],
    ['Sequestri amministrativi', c.sequestriAmministrativi], ['Fermi amministrativi', c.fermiAmministrativi], ['Sequestri penali', c.sequestriPenali], ['C.N.R.', c.cnr], [`Altri atti ${c.altriAttiDescrizione || ''}`, c.altriAttiNumero]
  ], y, [150, 36], title, subtitle);

  y = section(doc, 'Violazioni', y, title, subtitle);
  y = simpleTable(doc, ['Tipologia', 'N.'], [
    ['Preavvisi CdS', c.preavvisiCds], ['VdC CdS', c.vdcCds], ['Regolamento Polizia', c.regPolizia], ['Regolamento Edilizio', c.regEdilizio],
    ['Regolamento Benessere Animali', c.regBenessereAnimali], ['Annonaria / commercio', c.annonaria], [`Altre norme ${c.altreNormeDescrizione || ''}`, c.altreNorme], ['TOTALE', getTotaleViolazioni(report)]
  ], y, [150, 36], title, subtitle);

  y = section(doc, 'Documenti ritirati', y, title, subtitle);
  const docRows = (report.documentiRitirati || []).filter(d => d.tipo || d.quantita || d.note).map(d => [d.tipo || '-', d.quantita || '-', d.note || '-']);
  y = simpleTable(doc, ['Tipo documento', 'Quantità', 'Note'], docRows.length ? docRows : [['Nessun documento ritirato', '-', '-']], y, [70, 28, 88], title, subtitle);

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
    { label: 'Violazioni', value: aggregate.totaleViolazioni },
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
    reportRows.push({ Data: r.data || '', Turno: turnoLabel(r), 'Tipologia orario': r.orarioTipo || '', 'Reparto / servizio': repartoLabel(r), 'Zona di servizio': r.zonaServizio || '', Operatori: operatori, 'Numero operatori': (r.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).length, 'Interventi totali': (r.interventi || []).length, 'Violazioni': getTotaleViolazioni(r), 'Km percorsi': getKmTotali(r), 'Note UDT': r.noteUdt || '' });
    Object.entries(r.counters || {}).forEach(([key, value]) => { if (typeof value === 'number' && n(value) > 0) countersRows.push({ Data: r.data || '', Turno: turnoLabel(r), Reparto: repartoLabel(r), Voce: LABELS[key] || key, Totale: n(value) }); });
    (r.interventi || []).forEach((i, idx) => {
      const durata = durataMinuti(i.oraInizio, i.oraFine);
      const fascia = fasciaOraria(i.oraInizio);
      addCount(byFascia, fascia);
      const origine = i.origine === 'Altro' ? `Altro: ${i.origineAltro || '-'}` : (i.origine || 'Non indicata');
      const scuole = i.tipo === 'Servizio scuole' ? (i.scuole || []).filter(s => s.nome || s.momento || s.orario || s.criticita).map((s, pos) => `Scuola ${pos + 1}: ${s.nome || '-'} (${s.momento || '-'} ${s.orario || '-'}) Criticità: ${s.criticita || '-'}`).join(' | ') : '';
      const criticita = isInterventoCritico(i) ? 'Sì' : 'No';
      if (durata !== '' && durata > interventoPiuLungo.durata) interventoPiuLungo = { durata, label: `${r.data || ''} ${turnoLabel(r)} - ${i.tipo || '-'} (${durata} min)` };
      const row = { Data: r.data || '', Turno: turnoLabel(r), 'Tipologia orario': r.orarioTipo || '', 'Reparto / servizio': repartoLabel(r), 'Zona di servizio': r.zonaServizio || '', Operatori: operatori, 'N. report': reportIndex + 1, 'N. intervento': idx + 1, 'Tipo intervento': i.tipo || '', Origine: origine, 'Ora inizio': i.oraInizio || '', 'Ora fine': i.oraFine || '', 'Durata minuti': durata, 'Fascia oraria': fascia, Luogo: i.luogo || '', Descrizione: i.descrizione || '', Esito: i.esito || '', Note: i.note || '', Criticità: criticita, Feriti: i.conFeriti || '', 'Veicoli coinvolti': i.veicoliCoinvolti || '', 'Rilievi effettuati': i.rilievi || '', 'Veicoli controllati': i.veicoliControllati || '', 'Persone controllate': i.personeControllate || '', 'Verbali elevati': i.verbaliElevati || '', 'Fermi / sequestri intervento': i.fermiSequestri || '', 'Motivo viabilità': i.motivoViabilita || '', 'Strade interessate': i.strade || '', 'Scuole presidiate': scuole };
      interventiRows.push(row);
      if (criticita === 'Sì') criticitaRows.push({ Data: row.Data, Turno: row.Turno, Reparto: row['Reparto / servizio'], Zona: row['Zona di servizio'] || '', 'Tipo intervento': row['Tipo intervento'], Orario: `${row['Ora inizio']} - ${row['Ora fine']}`, Luogo: row.Luogo, Motivo: `${row.Descrizione} ${row.Note}`.trim() });
    });
  });
  const maxTipo = Math.max(1, ...Object.values(aggregate.byTipo || {}).map(n));
  const maxOrigine = Math.max(1, ...Object.values(aggregate.byOrigine || {}).map(n));
  const maxReparto = Math.max(1, ...Object.values(aggregate.byReparto || {}).map(n));
  const maxFascia = Math.max(1, ...Object.values(byFascia).map(n));
  const topTipo = orderedObjectRows(aggregate.byTipo)[0]?.Voce || '-';
  const topReparto = orderedObjectRows(aggregate.byReparto)[0]?.Voce || '-';
  const topFascia = orderedObjectRows(byFascia)[0]?.Voce || '-';
  const dashboardRows = [['REPORT AGGREGATO POLIZIA LOCALE - DASHBOARD EXCEL', ''], ['Periodo / data', aggregate.dateLabel], ['Report ricevuti', reports.length], ['Interventi totali', aggregate.totalInterventi], ['Violazioni', aggregate.totaleViolazioni], ['Km totali percorsi', aggregate.kmTotali], ['Tipologia intervento prevalente', topTipo], ['Reparto più impegnato', topReparto], ['Fascia oraria più intensa', topFascia], ['Intervento più lungo', interventoPiuLungo.label], ['Note ufficiale', commanderNotes || '-']];
  const graficiRows = [['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byTipo).map(r => ['Tipologia interventi', r.Voce, r.Totale, makeBar(r.Totale, maxTipo)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byOrigine).map(r => ['Origine interventi', r.Voce, r.Totale, makeBar(r.Totale, maxOrigine)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(aggregate.byReparto).map(r => ['Reparti / servizi', r.Voce, r.Totale, makeBar(r.Totale, maxReparto)]), [], ['Sezione', 'Voce', 'Totale', 'Grafico'], ...orderedObjectRows(byFascia).map(r => ['Fasce orarie', r.Voce, r.Totale, makeBar(r.Totale, maxFascia)])];
  const riepilogoRows = [{ Sezione: 'KPI', Voce: 'Report ricevuti', Totale: reports.length }, { Sezione: 'KPI', Voce: 'Interventi totali', Totale: aggregate.totalInterventi }, { Sezione: 'KPI', Voce: 'Violazioni', Totale: aggregate.totaleViolazioni }, { Sezione: 'KPI', Voce: 'Km totali percorsi', Totale: aggregate.kmTotali }, ...orderedObjectRows(aggregate.byTipo).map(r => ({ Sezione: 'Tipologia interventi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.byOrigine).map(r => ({ Sezione: 'Origine interventi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.byReparto).map(r => ({ Sezione: 'Reparti / servizi', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(byFascia).map(r => ({ Sezione: 'Fasce orarie', Voce: r.Voce, Totale: r.Totale })), ...orderedObjectRows(aggregate.counters, LABELS).map(r => ({ Sezione: 'Atti e violazioni', Voce: r.Voce, Totale: r.Totale }))];
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
  return ['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme'].reduce((s, k) => s + n(c[k]), 0);
}
function operatorNames(report) { return (report.operatori || []).filter(o => o.nome || o.matricola || o.qualifica).map(o => `${o.nome || 'Operatore'}${o.matricola ? ` mtr. ${o.matricola}` : ''}`); }
function extraDetails(i) {
  if (i.tipo === 'Codice della strada') {
    const dettaglio = i.cdsDettaglio || 'Non specificato';
    const righe = [`   Dettaglio CdS: ${dettaglio}`];
    if (dettaglio === 'Controllo soste') {
      righe.push(`   Rimozione veicolo: ${i.cdsRimozione || 'No'}`);
      righe.push(`   Motivazione: ${i.cdsMotivazione || '-'}`);
    }
    if (dettaglio === 'Buca su carreggiata') righe.push(`   Richiesto intervento per ripristino: ${i.cdsRipristino || 'No'}`);
    if (dettaglio === 'Sinistro stradale') {
      righe.push(`   Feriti: ${i.cdsFeriti || 'No'}`);
      righe.push(`   Veicoli coinvolti: ${i.cdsVeicoliCoinvolti || '-'}`);
    }
    if (dettaglio === 'Veicolo sospetto') {
      righe.push(`   Verifica effettuata: ${i.cdsVerificaEffettuata || 'No'}`);
      righe.push(`   Segnalazione: ${i.cdsSegnalazione || 'No'}`);
    }
    if (dettaglio === 'Segnaletica danneggiata') {
      righe.push(`   Pericolo: ${i.cdsPericolo || 'No'}`);
      righe.push(`   Intervento richiesto: ${i.cdsInterventoRichiesto || 'No'}`);
    }
    if (dettaglio === 'Guasto semaforo') {
      righe.push(`   Stato: ${i.cdsStatoSemaforo || '-'}`);
      righe.push(`   Intervento richiesto: ${i.cdsInterventoRichiesto || 'No'}`);
    }
    return righe.join('\n') + '\n';
  }
  if (i.tipo === 'Sinistro stradale') return `   Dettagli: ${i.conFeriti}; veicoli coinvolti ${i.veicoliCoinvolti || '-'}; rilievi ${i.rilievi}
`;
  if (i.tipo === 'Posto di controllo') return `   Controlli: veicoli ${i.veicoliControllati || '0'}; persone ${i.personeControllate || '0'}; verbali ${i.verbaliElevati || '0'}; fermi/sequestri ${i.fermiSequestri || '0'}
`;
  if (i.tipo === 'Viabilità') return `   Motivo: ${i.motivoViabilita || '-'}; strade interessate: ${i.strade || '-'}
`;
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
  return `REPORT DI SERVIZIO - POLIZIA LOCALE\n\nDATA: ${report.data}\nTURNO: ${turnoLabel(report)} (${report.orarioTipo})\nREPARTO: ${repartoLabel(report)}${report.zonaServizio ? '\nZONA DI SERVIZIO: ' + report.zonaServizio : ''}\n\nOPERATORI\n${ops}\n\nVEICOLI\n${mezzi || '- Non indicati'}\nTotale km percorsi: ${getKmTotali(report)}\n\nINTERVENTI EFFETTUATI\n${interventi}\n\nATTI REDATTI\nRelazioni: ${c.relazioni}\nAnnotazioni: ${c.annotazioni}\nSequestri amministrativi: ${c.sequestriAmministrativi}\nFermi amministrativi: ${c.fermiAmministrativi}\nSequestri penali: ${c.sequestriPenali}\nCNR: ${c.cnr}\nAltri atti: ${c.altriAttiNumero} ${c.altriAttiDescrizione || ''}\n\nVIOLAZIONI / PROVVEDIMENTI\nPreavvisi CdS: ${c.preavvisiCds}\nVdC CdS: ${c.vdcCds}\nRegolamento Polizia: ${c.regPolizia}\nRegolamento Edilizio: ${c.regEdilizio}\nRegolamento Benessere Animali: ${c.regBenessereAnimali}\nAnnonaria / commercio: ${c.annonaria}\nAltre norme: ${c.altreNorme} ${c.altreNormeDescrizione || ''}\nFermi: ${c.fermi}\nSequestri: ${c.sequestri}\nTOTALE: ${getTotaleViolazioni(report)}\n\nNOTE PER UDT / UFFICIALE DI COORDINAMENTO\n${report.noteUdt || '-'}\n\nDICHIARAZIONE\nGli operatori dichiarano che quanto riportato corrisponde fedelmente alle attività effettivamente svolte e riscontrate durante il turno di servizio.\nConferma dichiarazione: ${report.dichiarazione ? 'SI' : 'NO'}\n`;
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
      const tipoAggregato = i.tipo === 'Codice della strada' && i.cdsDettaglio ? `Codice della strada - ${i.cdsDettaglio}` : (i.tipo || 'Non indicato');
      aggregate.byTipo[tipoAggregato] = n(aggregate.byTipo[tipoAggregato]) + 1;
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
  const details = reports.map((r, idx) => `${idx + 1}. ${r.data} | ${turnoLabel(r)} | ${r.orarioTipo} | ${repartoLabel(r)} | Operatori: ${operatorNames(r).join(', ') || '-'} | Interventi: ${(r.interventi || []).length} | Violazioni: ${getTotaleViolazioni(r)} | Km: ${getKmTotali(r)}`).join('\n') || '- Nessun report caricato';
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
  return `Nel turno indicato sono stati acquisiti ${reports.length} report degli operatori, per complessivi ${aggregate.totalInterventi} interventi rendicontati. L'attività prevalente risulta: ${topTipo}. Sono state registrate ${aggregate.totaleViolazioni} violazioni e ${aggregate.kmTotali} km complessivi. Si segnalano ${conFeriti} sinistri con feriti, ${tso} TSO e ${aso} ASO.`;
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
  return `REPORT UFFICIALE DI TURNO\n\nDATA E TURNO\n${official.data || aggregate.dateLabel} ${official.turno || ''}\n\nBRIEFING OPERATIVO\n${official.briefing || '-'}\n\nA.P.L. ASSENTI\n${official.assenti || '-'}\n\nA.P.L. IN RITARDO\n${official.ritardi || '-'}\n\nNOTE\n${official.noteGenerali || '-'}\n\nSINTESI OPERATIVA\n${autoSintesi}\n${official.eventiManuali ? '\nIntegrazioni: ' + official.eventiManuali : ''}\n\nEVENTI DEGNI DI RILIEVO\n${autoEventi || '- Nessun evento rilevante automatico rilevato'}\n\nANOMALIE RISCONTRATE DURANTE IL TURNO\n${official.anomalie || '-'}\n\nATTIVITÀ ISPETTIVE\n${attivita}\n\nESITI\n${official.esiti || '-'}\n\nCOMUNICAZIONE ALL'E.Q. DI TURNO\n${official.comunicazioneEq || '-'}\n\nNOTA PER IL COMANDANTE\n${official.notaComandante || '-'}\n\nVIOLAZIONI RISCONTRATE\nTotale violazioni: ${aggregate.totaleViolazioni}\n\nFIRMA\n${official.qualifica || ''}\n${official.ufficiale || ''}`;
}
function buildOfficialShiftPdf_old(aggregate, reports, official, autoSintesi, autoEventi) {
  const title = 'REPORT UFFICIALE DI TURNO';
  const subtitle = `${official.data || aggregate.dateLabel} | ${official.turno || ''}`;
  const doc = makePdf(title, subtitle);
  let y = 54;
  y = section(doc, 'Data e turno', y, title, subtitle);
  y = kvGrid(doc, [{ label: 'Data', value: official.data || aggregate.dateLabel }, { label: 'Turno', value: official.turno || '-' }, { label: 'Ufficiale di turno', value: official.ufficiale || '-' }, { label: 'Qualifica', value: official.qualifica || '-' }], y, 2, title, subtitle) + 2;
  y = section(doc, 'Riepilogo rapido', y, title, subtitle);
  y = kvGrid(doc, [{ label: 'Report operatori', value: reports.length }, { label: 'Interventi totali', value: aggregate.totalInterventi }, { label: 'Violazioni', value: aggregate.totaleViolazioni }, { label: 'Km complessivi', value: aggregate.kmTotali }], y, 4, title, subtitle) + 2;
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
  y = section(doc, 'Violazioni riscontrate per pattuglia / reparto', y, title, subtitle);
  const violationRows = reports.map(r => { const c = r.counters || {}; return [operatorNames(r).join(' / ') || '-', repartoLabel(r), n(c.preavvisiCds), n(c.vdcCds), n(c.regPolizia), n(c.regEdilizio), n(c.regBenessereAnimali), n(c.annonaria), n(c.altreNorme), getTotaleViolazioni(r)]; });
  const tot = reports.reduce((acc, r) => { const c = r.counters || {}; ['preavvisiCds','vdcCds','regPolizia','regEdilizio','regBenessereAnimali','annonaria','altreNorme'].forEach(k => acc[k] = n(acc[k]) + n(c[k])); return acc; }, {});
  if (reports.length) violationRows.push(['TOTALE COMPLESSIVO', '-', n(tot.preavvisiCds), n(tot.vdcCds), n(tot.regPolizia), n(tot.regEdilizio), n(tot.regBenessereAnimali), n(tot.annonaria), n(tot.altreNorme), reports.reduce((sum, r) => sum + getTotaleViolazioni(r), 0)]);
  y = simpleTable(doc, ['Pattuglia / Operatori', 'Reparto', 'Prev.', 'VdC', 'Reg.PU', 'Edil.', 'Anim.', 'Ann.', 'Altre', 'Tot.'], violationRows.length ? violationRows : [['-', '-', '-', '-', '-', '-', '-', '-', '-', '-']], y, [40, 34, 14, 14, 16, 14, 16, 16, 16, 16], title, subtitle);
  y = section(doc, 'Atti redatti - riepilogo separato', y, title, subtitle);
  const attiTot = reports.reduce((acc, r) => { const c = r.counters || {}; ['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].forEach(k => acc[k] = n(acc[k]) + n(c[k])); return acc; }, {});
  y = simpleTable(doc, ['Tipologia atto', 'Totale'], [['Relazioni di servizio', n(attiTot.relazioni)], ['Annotazioni di servizio', n(attiTot.annotazioni)], ['Sequestri amministrativi', n(attiTot.sequestriAmministrativi)], ['Fermi amministrativi', n(attiTot.fermiAmministrativi)], ['Sequestri penali', n(attiTot.sequestriPenali)], ['C.N.R.', n(attiTot.cnr)], ['Altri atti', n(attiTot.altriAttiNumero)]], y, [150, 36], title, subtitle);
  y = ensureSpace(doc, y, 22, title, subtitle); doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.text('FIRMA:', 16, y + 4); doc.text(official.qualifica || '-', 16, y + 12); doc.text(official.ufficiale || '-', 16, y + 18); doc.line(80, y + 18, 190, y + 18);
  addFooter(doc); return doc;
}



// Modern coordinated PDF layout override
function themeColors() {
  return { blue: [12,47,97], green: [18,121,78], orange: [234,88,12], red: [220,38,38], gray: [245,247,250], line: [215,222,232], text: [20,28,40], muted: [90,99,115] };
}
function setC(doc, arr) { doc.setTextColor(arr[0], arr[1], arr[2]); }
function fillC(doc, arr) { doc.setFillColor(arr[0], arr[1], arr[2]); }
function drawMiniIcon(doc, type, x, y, color = [12,47,97]) {
  const c = color; doc.setDrawColor(c[0],c[1],c[2]); doc.setFillColor(c[0],c[1],c[2]); doc.setLineWidth(0.55);
  if (type === 'car') { doc.roundedRect(x, y+2.8, 8, 4, 1,1,'S'); doc.line(x+1.5,y+2.8,x+3,y+0.8); doc.line(x+3,y+0.8,x+6,y+0.8); doc.line(x+6,y+0.8,x+7.2,y+2.8); doc.circle(x+2,y+7.2,0.9,'F'); doc.circle(x+6.2,y+7.2,0.9,'F'); }
  else if (type === 'doc') { doc.rect(x+1,y,6.2,8.2,'S'); doc.line(x+5.2,y,x+7.2,y+2); doc.line(x+5.2,y,x+5.2,y+2); doc.line(x+5.2,y+2,x+7.2,y+2); doc.line(x+2.3,y+4,x+6,y+4); doc.line(x+2.3,y+6,x+6,y+6); }
  else if (type === 'warn') { doc.triangle(x+4,y,x+8,y+8,x,y+8,'S'); doc.line(x+4,y+2.4,x+4,y+5.2); doc.circle(x+4,y+6.7,0.35,'F'); }
  else if (type === 'clip') { doc.setLineWidth(0.8); doc.line(x+2,y+2,x+5.8,y+6); doc.circle(x+5.8,y+6,1.1,'S'); doc.circle(x+2,y+2,1.1,'S'); }
  else if (type === 'people') { doc.circle(x+3,y+2,1.3,'S'); doc.circle(x+6,y+2.6,1.0,'S'); doc.roundedRect(x+0.8,y+4.5,4.7,3.2,1,1,'S'); doc.roundedRect(x+4.6,y+5,4.1,2.7,1,1,'S'); }
  else if (type === 'list') { doc.rect(x+1,y,6.5,8,'S'); doc.line(x+2.4,y+2.4,x+6.3,y+2.4); doc.line(x+2.4,y+4.2,x+6.3,y+4.2); doc.line(x+2.4,y+6,x+6.3,y+6); }
  else if (type === 'mail') { doc.rect(x,y+1.2,8,5.8,'S'); doc.line(x,y+1.2,x+4,y+4.2); doc.line(x+8,y+1.2,x+4,y+4.2); }
  else if (type === 'check') { doc.circle(x+4,y+4,4,'S'); doc.line(x+2.1,y+4.2,x+3.5,y+5.6); doc.line(x+3.5,y+5.6,x+6.3,y+2.8); }
  else if (type === 'user') { doc.circle(x+4,y+2.2,1.6,'S'); doc.roundedRect(x+1,y+5,6,3,1,1,'S'); }
  else { doc.circle(x+4,y+4,3,'S'); }
}
function drawHeaderModern(doc, title, subtitle = '', accent = [12,47,97]) {
  const C = themeColors();
  fillC(doc, [255,255,255]); doc.rect(0,0,210,52,'F');
  try { const img = document.getElementById('pdfLogo'); if (img && img.complete) doc.addImage(img, 'PNG', 12, 7, 24, 24); } catch(e) {}
  doc.setDrawColor(C.blue[0],C.blue[1],C.blue[2]); doc.setLineWidth(0.4); doc.line(42,8,42,32);
  doc.setFont('helvetica','bold'); doc.setFontSize(16); setC(doc,C.blue); doc.text('COMUNE DI MONZA', 48, 17);
  doc.setFontSize(10.5); doc.text('Polizia Locale', 48, 24);
  doc.setFont('helvetica','normal'); doc.setFontSize(8.2); doc.text('Settore Polizia Locale e Protezione Civile', 48, 30);
  doc.setFontSize(7.8); setC(doc,C.text); doc.text('Via Marsala 13', 160, 12); doc.text('20900 Monza',160,16.5); doc.text('Tel. 039 28161',160,23); doc.text('polizialocale@comune.monza.it',160,29.5);
  doc.setDrawColor(C.blue[0],C.blue[1],C.blue[2]); doc.line(12,37,198,37);
  fillC(doc, accent); doc.roundedRect(12,42,186,9,1.4,1.4,'F');
  doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(255,255,255); doc.text(title,75,48.2,{align:'center'});
  if (subtitle) { doc.setFontSize(8.2); doc.text(subtitle,195,48.2,{align:'right'}); }
}
function footerModern(doc) {
  const C = themeColors(); const pages = doc.internal.getNumberOfPages();
  for (let p=1;p<=pages;p++) { doc.setPage(p); doc.setDrawColor(C.blue[0],C.blue[1],C.blue[2]); doc.setLineWidth(0.35); doc.line(12,284,198,284); doc.setFont('helvetica','normal'); doc.setFontSize(7.2); setC(doc,C.blue); doc.text('Settore Polizia Locale, Protezione Civile',12,289); doc.text('Via Marsala 13 | 20900 Monza',12,293); doc.text('Tel. 039 28161',91,291); doc.text('polizialocale@comune.monza.it',132,291); fillC(doc,C.blue); doc.roundedRect(178,287,20,7,1.2,1.2,'F'); doc.setTextColor(255,255,255); doc.setFont('helvetica','bold'); doc.text(`Pag. ${p} di ${pages}`,188,291.5,{align:'center'}); }
}
function drawPanel(doc, x,y,w,h,title, icon='list', opts={}) {
  const C = themeColors(); const bg = opts.bg || [255,255,255]; const border = opts.border || C.line; const accent = opts.accent || C.blue;
  fillC(doc,bg); doc.setDrawColor(border[0],border[1],border[2]); doc.setLineWidth(0.35); doc.roundedRect(x,y,w,h,1.6,1.6,'FD');
  if (opts.leftStripe) { fillC(doc, opts.leftStripe); doc.rect(x,y,2.2,h,'F'); }
  drawMiniIcon(doc, icon, x+4, y+4, accent);
  doc.setFont('helvetica','bold'); doc.setFontSize(9); setC(doc,C.blue); doc.text(title.toUpperCase(), x+15, y+10);
  setC(doc,C.text); doc.setFont('helvetica','normal');
}
function drawKpiBox(doc,x,y,w,h,icon,label,value,color=[12,47,97]) {
  const C = themeColors(); fillC(doc,[255,255,255]); doc.setDrawColor(C.line[0],C.line[1],C.line[2]); doc.roundedRect(x,y,w,h,1.6,1.6,'FD');
  drawMiniIcon(doc, icon, x+w/2-4, y+5, color);
  doc.setFont('helvetica','bold'); doc.setFontSize(8.2); setC(doc,color); doc.text(label.toUpperCase(), x+w/2, y+20,{align:'center'});
  doc.setFontSize(20); doc.text(String(value ?? '-'), x+w/2, y+32,{align:'center'});
  doc.setFont('helvetica','normal'); doc.setFontSize(7.2); setC(doc,themeColors().muted); doc.text('Totali', x+w/2, y+38,{align:'center'});
}
function writeTextInBox(doc, text, x, y, w, maxLines=8, fontSize=8) {
  doc.setFont('helvetica','normal'); doc.setFontSize(fontSize); setC(doc,themeColors().text);
  let lines = doc.splitTextToSize(String(text || '-'), w);
  if (lines.length > maxLines) lines = lines.slice(0,maxLines-1).concat(['…']);
  doc.text(lines, x, y);
  return lines.length;
}
function countTextItems(text) { const s = String(text || '').trim(); if (!s || /^nessun[oa]$/i.test(s) || s === '-') return 0; return s.split(/\n|;/).map(x=>x.trim()).filter(Boolean).length; }
function formatDateIT(value) {
  const s = String(value || '').trim();
  if (!s) return '-';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return String(d.getDate()).padStart(2,'0') + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
  return s;
}
function listFromText(text) {
  return String(text || '').split(/\n|;/).map(x => x.trim()).filter(x => x && x !== '-');
}
function totalKmFromReports(reports) { return reports.reduce((sum,r)=>sum+getKmTotali(r),0); }
function totalVehiclesFromReports(reports) { return reports.reduce((sum,r)=>sum+((r.veicoli||[]).filter(v=>v.sigla||v.kmInizio||v.kmFine).length),0); }

function totalAttiFromReports(reports) { return reports.reduce((sum,r)=>{ const c=r.counters||{}; return sum + ['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].reduce((s,k)=>s+n(c[k]),0); },0); }
function attiObjectFromReports(reports) { return reports.reduce((acc,r)=>{ const c=r.counters||{}; ['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].forEach(k=>acc[k]=n(acc[k])+n(c[k])); return acc; },{}); }
function buildViolationRows(reports) { const rows = reports.map(r=>{ const c=r.counters||{}; return [operatorNames(r).join(' / ') || '-', repartoLabel(r), n(c.preavvisiCds), n(c.vdcCds), n(c.regPolizia), n(c.annonaria), n(c.altreNorme), getTotaleViolazioni(r)]; }); const totals = reports.reduce((acc,r)=>{ const c=r.counters||{}; ['preavvisiCds','vdcCds','regPolizia','annonaria','altreNorme'].forEach(k=>acc[k]=n(acc[k])+n(c[k])); acc.tot += getTotaleViolazioni(r); return acc; },{preavvisiCds:0,vdcCds:0,regPolizia:0,annonaria:0,altreNorme:0,tot:0}); if (rows.length) rows.push(['TOTALE COMPLESSIVO','-',totals.preavvisiCds,totals.vdcCds,totals.regPolizia,totals.annonaria,totals.altreNorme,totals.tot]); return rows; }
function drawModernTable(doc, x,y,w,headers,rows,widths,opts={}) {
  const C=themeColors(); const rowH=8; const headerH=10; fillC(doc,C.blue); doc.roundedRect(x,y,w,headerH,1.2,1.2,'F');
  doc.setFont('helvetica','bold'); doc.setFontSize(7.2); doc.setTextColor(255,255,255); let xx=x;
  headers.forEach((h,i)=>{ doc.text(String(h), xx+widths[i]/2, y+6.2,{align:'center', maxWidth: widths[i]-2}); xx+=widths[i]; });
  y+=headerH; doc.setFont('helvetica','normal'); doc.setFontSize(7.2); rows.forEach((row,ri)=>{ if (ri===rows.length-1 && opts.totalLast) { fillC(doc,C.blue); doc.setTextColor(255,255,255); doc.setFont('helvetica','bold'); } else { fillC(doc, ri%2 ? [255,255,255] : [248,250,252]); setC(doc,C.text); doc.setFont('helvetica','normal'); } doc.setDrawColor(C.line[0],C.line[1],C.line[2]); doc.rect(x,y,w,rowH,'FD'); let cx=x; row.forEach((cell,i)=>{ const align = i>=2 ? 'center':'left'; doc.text(String(cell ?? '-'), align==='center'?cx+widths[i]/2:cx+2, y+5.4,{align, maxWidth: widths[i]-3}); if (i<row.length-1) { doc.setDrawColor(C.line[0],C.line[1],C.line[2]); doc.line(cx+widths[i],y,cx+widths[i],y+rowH); } cx+=widths[i]; }); y+=rowH; }); setC(doc,C.text); return y;
}
function newCleanDoc() { const doc = new jsPDF({unit:'mm', format:'a4'}); doc.setProperties({title:'Report Polizia Locale', author:'Polizia Locale'}); return doc; }
function buildOfficialShiftPdf(aggregate, reports, official, autoSintesi, autoEventi) {
  const C=themeColors(); const doc=newCleanDoc(); const subtitle=`${formatDateIT(official.data || aggregate.dateLabel)} | ${official.turno || ''}`;
  drawHeaderModern(doc,'REPORT UFFICIALE DI TURNO',subtitle,C.blue);
  drawKpiBox(doc,12,58,40,38,'car','Interventi',aggregate.totalInterventi,C.blue);
  drawKpiBox(doc,58,58,40,38,'doc','Verbali',aggregate.totaleViolazioni,C.blue);
  drawKpiBox(doc,104,58,40,38,'warn','Eventi',relevantInterventions(reports).length,C.orange);
  drawKpiBox(doc,150,58,48,38,'clip','Atti redatti',totalAttiFromReports(reports),C.blue);
  drawPanel(doc,12,104,67,55,'Personale','people');
  const presenti = new Set(reports.flatMap(r=>(r.operatori||[]).map(o=>o.matricola||o.nome).filter(Boolean))).size;
  const ritardiList=listFromText(official.ritardi); const assentiList=listFromText(official.assenti);
  const ritardi=ritardiList.length; const assenti=assentiList.length;
  doc.setFontSize(8); setC(doc,C.text); doc.setFont('helvetica','normal');
  doc.text('Presenti',18,121); setC(doc,C.green); doc.setFont('helvetica','bold'); doc.text(String(presenti || '-'),72,121,{align:'right'});
  doc.setDrawColor(C.line[0],C.line[1],C.line[2]); doc.line(18,124,73,124);
  setC(doc,C.text); doc.setFont('helvetica','normal'); doc.text('Ritardo',18,130); setC(doc,C.orange); doc.setFont('helvetica','bold'); doc.text(String(ritardi),72,130,{align:'right'});
  if (ritardiList.length) { setC(doc,C.muted); doc.setFont('helvetica','normal'); doc.setFontSize(6.6); doc.text(doc.splitTextToSize(ritardiList.join(', '),50).slice(0,2),18,135); }
  doc.setDrawColor(C.line[0],C.line[1],C.line[2]); doc.line(18,142,73,142);
  setC(doc,C.text); doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.text('Assenti',18,148); setC(doc,C.red); doc.setFont('helvetica','bold'); doc.text(String(assenti),72,148,{align:'right'});
  if (assentiList.length) { setC(doc,C.muted); doc.setFont('helvetica','normal'); doc.setFontSize(6.6); doc.text(doc.splitTextToSize(assentiList.join(', '),50).slice(0,2),18,153); }
  drawPanel(doc,84,104,114,55,'Briefing operativo','list'); writeTextInBox(doc, official.briefing || '-', 90, 122, 100, 7, 8);
  drawPanel(doc,12,166,186,45,'Eventi / anomalie degne di rilievo','warn',{leftStripe:C.orange,bg:[255,251,245],border:[245,208,166],accent:C.orange}); writeTextInBox(doc, `${autoEventi || 'Nessun evento rilevante automatico rilevato.'}${official.anomalie ? '\nAnomalie: '+official.anomalie : ''}`, 18, 184, 172, 5, 8.2);
  drawPanel(doc,12,218,90,38,'Note generali','doc'); writeTextInBox(doc, official.noteGenerali || '-', 18,236,78,4,8);
  drawPanel(doc,108,218,90,38,'Sintesi operativa','list'); writeTextInBox(doc, `${autoSintesi}${official.eventiManuali ? '\n'+official.eventiManuali : ''}\nKm totali veicoli: ${totalKmFromReports(reports)} km`, 114,236,78,4,8);
  footerModern(doc);
  doc.addPage(); drawHeaderModern(doc,'REPORT UFFICIALE DI TURNO - DETTAGLIO',subtitle,C.blue);
  drawPanel(doc,12,58,186,10,'Violazioni riscontrate','list');
  const rows=buildViolationRows(reports); drawModernTable(doc,12,72,186,['Pattuglia','Reparto','Prev.','C.d.S.','Urbana','Annon.','Altre','Tot.'], rows.length?rows:[['-','-','0','0','0','0','0','0']], [42,35,18,18,20,20,18,15], {totalLast:true});
  drawPanel(doc,12,168,82,45,'Atti redatti','clip'); const at=attiObjectFromReports(reports); const attiLines=[['Fermi amministrativi',at.fermiAmministrativi],['Sequestri amministrativi',at.sequestriAmministrativi],['Sequestri penali',at.sequestriPenali],['Notizie di reato',at.cnr]]; doc.setFontSize(8); attiLines.forEach((r,i)=>{ setC(doc,C.text); doc.setFont('helvetica','normal'); doc.text(r[0],18,186+i*7); doc.setFont('helvetica','bold'); doc.text(String(n(r[1])),88,186+i*7,{align:'right'}); });
  drawPanel(doc,101,168,97,30,'Esito turno','check',{bg:[247,253,250],border:[187,223,206],accent:C.green}); writeTextInBox(doc, official.esiti || '-', 107,186,84,3,8);
  drawPanel(doc,101,204,97,25,"Comunicazioni E.Q.",'mail'); writeTextInBox(doc, official.comunicazioneEq || '-', 107,222,84,2,8);
  drawPanel(doc,12,236,124,34,'Nota del Comandante','user'); writeTextInBox(doc, official.notaComandante || '-', 18,254,110,3,8);
  drawPanel(doc,144,236,54,34,'Responsabile di turno','user'); doc.setFontSize(7.5); setC(doc,C.text); doc.text(official.qualifica || '-',171,253,{align:'center'}); doc.setFont('helvetica','bold'); doc.text(official.ufficiale || '-',171,259,{align:'center'}); doc.line(154,265,188,265);
  footerModern(doc); return doc;
}
function buildServicePdf(report) {
  const C=themeColors(); const doc=newCleanDoc(); const subtitle=`${formatDateIT(report.data)} | ${turnoLabel(report)} | ${report.orarioTipo}`;
  drawHeaderModern(doc,'REPORT DI SERVIZIO',subtitle,C.green);
  const interventi=(report.interventi||[]).length; const violazioni=getTotaleViolazioni(report); const c=report.counters||emptyCounters(); const atti=['relazioni','annotazioni','sequestriAmministrativi','fermiAmministrativi','sequestriPenali','cnr','altriAttiNumero'].reduce((s,k)=>s+n(c[k]),0); const eventi=(report.interventi||[]).filter(isInterventoCritico).length;
  drawKpiBox(doc,12,58,40,38,'car','Interventi',interventi,C.green); drawKpiBox(doc,58,58,40,38,'doc','Violazioni',violazioni,C.green); drawKpiBox(doc,104,58,40,38,'clip','Atti redatti',atti,C.green); drawKpiBox(doc,150,58,48,38,'warn','Eventi',eventi,C.orange);
  drawPanel(doc,12,104,60,48,'Veicoli','car',{accent:C.green});
  const v=(report.veicoli||[])[0]||{}; const veicoliUsati=(report.veicoli||[]).filter(vv=>vv.sigla||vv.kmInizio||vv.kmFine).length;
  doc.setFontSize(8); setC(doc,C.text); doc.setFont('helvetica','normal');
  doc.text(`Veicoli impiegati: ${veicoliUsati || '-'}`,18,122);
  doc.setFont('helvetica','bold'); doc.setFontSize(11); setC(doc,C.green); doc.text(`${getKmTotali(report)} km`,18,134);
  setC(doc,C.text); doc.setFont('helvetica','normal'); doc.setFontSize(7.3); doc.text('Totale km percorsi nel turno',18,143);
  drawPanel(doc,78,104,60,48,'Carburante','list',{accent:C.green}); doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.text(`Effettuato: ${v.carburante || 'No'}`,84,122); doc.text(`Importo: ${v.importoCarburante || '-'}`,84,129); doc.text(`Card presa: ${v.oraPrelievoCard || '-'}`,84,136); doc.text(`Card resa: ${v.oraRestituzioneCard || '-'}`,84,143);
  drawPanel(doc,144,104,54,48,'Anomalie veicolo','warn',{accent:C.green}); writeTextInBox(doc, (report.veicoli||[]).filter(x=>x.anomaliaVeicolo).map(x=>`${x.sigla||'Veicolo'}: ${x.anomaliaVeicolo}`).join('\n') || 'Nessuna anomalia segnalata.', 150,122,42,4,8);
  drawPanel(doc,12,160,186,45,'Note di servizio','list',{accent:C.green}); writeTextInBox(doc, report.noteUdt || '-',18,178,174,5,8);
  drawPanel(doc,12,212,186,46,'Operatori','people',{accent:C.green}); const opLines=(report.operatori||[]).filter(o=>o.nome||o.matricola).map(o=>`${o.nome || '-'} ${o.matricola? '— mtr. '+o.matricola:''} ${o.qualifica? '— '+o.qualifica:''}`).join('\n') || '-'; writeTextInBox(doc, `Reparto: ${repartoLabel(report)}${report.zonaServizio ? '\nZona di servizio: ' + report.zonaServizio : ''}\n${opLines}`,18,230,172,5,8);
  footerModern(doc);
  doc.addPage(); drawHeaderModern(doc,'REPORT DI SERVIZIO - DETTAGLIO',subtitle,C.green);
  drawPanel(doc,12,58,88,120,'Interventi effettuati','car',{accent:C.green}); let yy=77; doc.setFontSize(7.7); (report.interventi||[]).slice(0,9).forEach((i,idx)=>{ fillC(doc,C.green); doc.circle(18,yy-1.5,0.9,'F'); setC(doc,C.text); doc.setFont('helvetica','bold'); doc.text(`${i.oraInizio || '--'} ${i.luogo || i.tipo || '-'}`,22,yy,{maxWidth:70}); doc.setFont('helvetica','normal'); const desc=`${i.tipo || ''}${i.esito ? ' — '+i.esito : ''}`; doc.text(doc.splitTextToSize(desc,70).slice(0,2),22,yy+4); yy+=12; }); if (!(report.interventi||[]).length) writeTextInBox(doc,'Nessun intervento inserito.',18,77,75,3,8);
  drawPanel(doc,108,58,90,58,'Violazioni contestate','doc',{accent:C.green}); drawModernTable(doc,114,75,78,['Tipo violazione','Nr.'],[['Codice della Strada',n(c.preavvisiCds)+n(c.vdcCds)],['Regolamenti comunali',n(c.regPolizia)+n(c.regEdilizio)+n(c.regBenessereAnimali)],['Annonaria / commercio',n(c.annonaria)],['Altro',n(c.altreNorme)],['TOTALE',violazioni]],[58,20],{totalLast:true});
  drawPanel(doc,108,125,90,48,'Atti redatti','clip',{accent:C.green}); const serviceAtti=[['Relazioni',c.relazioni],['Annotazioni',c.annotazioni],['Fermi amministrativi',c.fermiAmministrativi],['Sequestri amministrativi',c.sequestriAmministrativi],['Sequestri penali',c.sequestriPenali],['C.N.R.',c.cnr]].filter(r=>n(r[1])>0); doc.setFontSize(8); (serviceAtti.length?serviceAtti:[['Nessun atto redatto',0]]).slice(0,5).forEach((r,i)=>{ setC(doc,C.text); doc.setFont('helvetica','normal'); doc.text(r[0],114,143+i*6.5); doc.setFont('helvetica','bold'); doc.text(String(r[1]),191,143+i*6.5,{align:'right'}); });
  drawPanel(doc,12,188,112,40,'Osservazioni','list',{accent:C.green}); writeTextInBox(doc, report.noteUdt || 'Nessuna osservazione particolare da segnalare.',18,206,100,4,8);
  drawPanel(doc,132,188,66,40,'Firma agente','user',{accent:C.green}); const names=operatorNames(report).join(' / ') || '-'; doc.setFontSize(7.4); setC(doc,C.text); doc.text(doc.splitTextToSize(names,54).slice(0,2),138,206); doc.line(142,221,190,221);
  drawPanel(doc,12,238,186,28,'Documenti ritirati','doc',{accent:C.green}); const docs=(report.documentiRitirati||[]).filter(d=>d.tipo||d.quantita||d.note).map(d=>`${d.tipo||'-'} (${d.quantita||'-'}): ${d.note||'-'}`).join(' | ') || 'Nessun documento ritirato.'; writeTextInBox(doc, docs,18,256,174,2,8);
  footerModern(doc); return doc;
}


createRoot(document.getElementById('root')).render(<App />);
