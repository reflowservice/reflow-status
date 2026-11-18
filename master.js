/* REFLOW STATUS — lettura del REFLOW MASTER (sola lettura).
   Trasforma il foglio "Impianti " (un blocco di 8 righe per impianto, colonna A unita)
   nell'elenco impianti mostrato dall'app. Stessa logica dello script
   aggiorna_master_status.py (rebuild_dati_flat + build_clients_json), portata in JavaScript
   per leggere il file direttamente da SharePoint senza passare da Dati_Flat.
   Nessun dato aziendale in questo file: solo regole di lettura e riferimenti geografici pubblici. */
(function (root) {
  'use strict';

  const SHEET = 'Impianti ';
  // Colonne (1 = A). Intestazioni attese, controllate a ogni lettura: se il Master cambia
  // struttura l'app si ferma con un messaggio invece di mostrare dati sfalsati.
  const COL = { cliente: 1, regione: 2, coord: 3, tipologia: 4, potenza: 5, ettari: 6, pr: 7, pod: 8,
    referente: 9, contratto: 10, decorrenza: 11, corrispettivo: 12, tempi: 13, cctv: 33, suf: 34,
    layout: 35, pm: 36, pm2: 37 };
  const EXPECTED_HEADERS = { 1: 'CLIENTE', 2: 'REGIONE', 4: 'TIPOLOGIA IMPIANTO', 7: 'PR CONTRATTUALE',
    10: 'TIPOLOGIA DI CONTRATTO', 14: 'MANUTENZIONE ORDINARIA', 18: 'MANUTENZIONE VERDE', 21: 'LAVAGGIO MODULI',
    23: 'DERATTIZZAZIONE', 25: 'MANUTENZIONE ESTINTORI', 29: 'MANUTENZIONE TVCC', 31: 'TERMOGRAFIA CON DRONE',
    33: 'SERVIZIO DI CCTV E VIGILANZA', 36: 'PM', 37: 'PM SECONDA LINEA' };
  const GROUP_COLS = [
    // ordine del Master dal 18/11/2026: Derattizzazione spostata accanto a Lavaggio Moduli
    ['Manutenzione Ordinaria', 14], ['Manutenzione Verde', 18], ['Lavaggio Moduli', 21],
    ['Derattizzazione', 23], ['Manutenzione Estintori', 25], ['Manutenzione TVCC', 29],
    ['Termografia con Drone', 31]
  ];
  const SERVICE_META = [
    ['Manutenzione Ordinaria', 'ORD'], ['Manutenzione Verde', 'VERDE'], ['Lavaggio Moduli', 'LAV'],
    ['Derattizzazione', 'DERAT'], ['Manutenzione Estintori', 'ESTINT'], ['Manutenzione TVCC', 'TVCC'],
    ['Termografia con Drone', 'TERMO'], ['Servizio di CCTV e vigilanza', 'CCTV']
  ];
  const STATUS_PATTERNS = [
    [/\bGARANZIA ESPE ATTIVA\b/i, 'garanzia_espe', 'Garanzia costruttiva ESPE attiva'],
    [/\bZONA GRIGIA\b/i, 'zona_grigia', 'Zona grigia — contratto O&M non attivato'],
    [/\bPROPOSTA NON FIRMATA\b/i, 'proposta_non_firmata', 'Proposta commerciale non firmata'],
    [/proposta non ancora firmata/i, 'proposta_non_firmata', 'Proposta commerciale non firmata'],
    [/\bPROPOSTA IN FORSE\b/i, 'proposta_in_forse', 'Proposta in forse'],
    [/\bPOTENZIALE\b/i, 'potenziale', 'Potenziale — nessuna proposta ancora avviata']
  ];

  // ---- coordinate (colonna 3): decimale "lat, lon" o DMS, con controllo di plausibilità sull'Italia
  const DMS = /(\d{1,3})\s*°\s*(\d{1,2}(?:\.\d+)?)\s*['’]\s*(\d{1,2}(?:\.\d+)?)?\s*(?:"|''|″)?\s*([NSEW])/g;
  const DMS_LOOSE = /(\d{1,3})\s*°\s*(\d{1,2}(?:\.\d+)?)\s*['’]\s*(\d{1,2}(?:\.\d+)?)?\s*(?:"|''|″)?\s*([NSEW])?/g;
  const DEC = /^\s*(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)/;
  const inItaly = (lat, lon) => lat > 30 && lat < 50 && lon > 5 && lon < 20;
  function fromDms(list) {
    const v = {};
    for (const [deg, mn, sec, h] of list) {
      const d = +deg + (+mn) / 60 + (sec ? +sec : 0) / 3600;
      v[h] = (h === 'S' || h === 'W') ? -d : d;
    }
    const lat = v.N !== undefined ? v.N : v.S, lon = v.E !== undefined ? v.E : v.W;
    return (lat !== undefined && lon !== undefined && inItaly(lat, lon)) ? [lat, lon] : null;
  }
  function parseCoord(raw) {
    if (raw === null || raw === undefined || raw === '') return null;
    const text = String(raw).trim();
    const first = text.split(/\r?\n/)[0].trim();
    const m = DEC.exec(first);
    if (m) { const lat = +m[1], lon = +m[2]; return inItaly(lat, lon) ? [lat, lon] : null; }
    const all = [...text.matchAll(DMS)].map(x => [x[1], x[2], x[3], x[4]]);
    if (all.length >= 2) { const r = fromDms(all.slice(0, 2)); if (r) return r; }
    // una sola coppia DMS senza N/E: impianti tutti in Italia, emisfero deducibile senza indovinare il valore
    const loose = [...text.matchAll(DMS_LOOSE)].map(x => [x[1], x[2], x[3], x[4] || '']).filter(t => t[0]);
    if (loose.length === 2 && loose.filter(t => !t[3]).length === 1) {
      loose[0][3] = (loose[0][3] || 'N').toUpperCase(); loose[1][3] = (loose[1][3] || 'E').toUpperCase();
      return fromDms(loose);
    }
    return null;
  }

  // ---- posizione approssimata quando mancano le coordinate (comune > provincia > regione)
  const COMUNE_COORDS = [
    ['SAN BIAGIO DI CALLALTA', 45.6667, 12.4167], ['MONTICELLO CONTE OTTO', 45.5833, 11.5333],
    ['TEMPIO PAUSANIA', 40.9022, 9.1028], ['ALBA ADRIATICA', 42.8333, 13.9167], ['GUALDO TADINO', 43.2333, 12.7833],
    ['CIVITAVECCHIA', 42.0928, 11.7961], ['PIOMBINO DESE', 45.6167, 11.9833], ['VALDOBBIADENE', 45.9019, 11.9967],
    ['PORTO VIRO', 45.0333, 12.2333], ['COLLECCHIO', 44.7756, 10.2242], ['GAMBETTOLA', 44.1167, 12.3333],
    ['FERRANDINA', 40.4936, 16.4597], ['MONSELICE', 45.2333, 11.75], ['MOSCIANO', 42.7333, 13.8667],
    ['BIENTINA', 43.7167, 10.6167], ['CALDOGNO', 45.6167, 11.5333], ['VALDAGNO', 45.65, 11.3333],
    ['FABRICO', 44.8386, 10.7397], ['MEDICINA', 44.4692, 11.6417], ['BOLOGNA', 44.4949, 11.3426],
    ['FIDENZA', 44.8667, 10.0667], ['SINALUNGA', 43.3333, 11.7333], ['BRINDISI', 40.632, 17.937],
    ['ASCOLI', 42.8536, 13.575], ['RESANA', 45.6167, 11.9333], ['THIENE', 45.7061, 11.4794],
    ['SALEMI', 37.8167, 12.8], ['PARMA', 44.8015, 10.3279], ['TREIA', 43.3167, 13.3333], ['URAS', 39.9167, 8.75],
    ['ISOLA RIZZA', 45.2833, 11.1667]
  ];
  const PROVINCE_CAPITALS = {
    VI: [45.5455, 11.5354], TV: [45.6667, 12.25], PD: [45.4064, 11.8768], FC: [44.2226, 12.0407],
    MT: [40.6664, 16.6043], SS: [40.7259, 8.5557], MC: [43.3006, 13.4531], BO: [44.4949, 11.3426],
    PR: [44.8015, 10.3279], SI: [43.3188, 11.3308], VR: [45.4384, 10.9916], TP: [38.0176, 12.5365],
    RM: [41.9028, 12.4964], PI: [43.7228, 10.4017], RE: [44.6979, 10.6307], RO: [45.0705, 11.7905],
    OR: [39.9033, 8.591], BR: [40.632, 17.937], AP: [42.8536, 13.575], TE: [42.6614, 13.6985], PG: [43.1122, 12.3888]
  };
  const REGION_CENTROIDS = {
    'VENETO': [45.5, 11.9], 'EMILIA ROMAGNA': [44.5, 11.0], 'TOSCANA': [43.4, 11.1], 'MARCHE': [43.3, 13.2],
    'BASILICATA': [40.6, 16.2], 'SARDEGNA': [40.1, 9.1], 'SICILIA': [37.6, 14.0], 'LAZIO': [41.9, 12.6],
    'UMBRIA': [42.9, 12.6], 'ABRUZZO': [42.2, 13.8], 'PUGLIA': [41.1, 16.6]
  };
  const JITTER = { comune: 0.02, provincia: 0.06, regione: 0.18 };
  function geoFallback(cliente, regione) {
    const up = String(cliente || '').toUpperCase();
    for (const [k, lat, lon] of COMUNE_COORDS) if (up.includes(k)) return [lat, lon, 'comune'];
    for (const m of up.matchAll(/\(([A-Z]{2})\)/g)) if (PROVINCE_CAPITALS[m[1]]) return [...PROVINCE_CAPITALS[m[1]], 'provincia'];
    const r = String(regione || '').trim().toUpperCase();
    if (REGION_CENTROIDS[r]) return [...REGION_CENTROIDS[r], 'regione'];
    return null;
  }
  // spostamento deterministico (stessa riga, stesso spostamento) per non sovrapporre i punti approssimati
  function jitter(riga, scale) {
    let h1 = 2166136261, h2 = 5381; const s = String(riga);
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619) >>> 0; h2 = (Math.imul(h2, 33) + c) >>> 0; }
    h1 = Math.imul(h1 ^ (h1 >>> 13), 0x5bd1e995) >>> 0; h2 = Math.imul(h2 ^ (h2 >>> 15), 0x27d4eb2d) >>> 0;
    return [(h1 / 0xFFFFFFFF - 0.5) * scale, (h2 / 0xFFFFFFFF - 0.5) * scale];
  }

  function classifyStatus(text) {
    for (const [re, cat, label] of STATUS_PATTERNS) if (re.test(text || '')) return [cat, label];
    return ['', ''];
  }

  // ---- lettura del foglio
  function parse(XLSX, data) {
    const wb = XLSX.read(data, { type: 'array', cellDates: true });
    const ws = wb.Sheets[SHEET];
    if (!ws) throw new Error('Nel Master non c\'è il foglio "Impianti " (nome cambiato?).');
    const addr = (r, c) => XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
    const cell = (r, c) => ws[addr(r, c)];
    const val = (r, c) => { const x = cell(r, c); if (!x || x.v === undefined || x.v === null) return ''; if (x.v instanceof Date) return x.w || x.v.toISOString().slice(0, 10); return typeof x.v === 'string' ? x.v.replace(/\r\n?/g, '\n') : x.v; };
    const str = (r, c) => { const v = val(r, c); return typeof v === 'string' ? v : (v === '' ? '' : String(v)); };
    const comment = (r, c) => { const x = cell(r, c); return x && x.c && x.c.length ? x.c.map(k => k.t || '').join('\n').trim() : ''; };

    const wrong = Object.entries(EXPECTED_HEADERS)
      .filter(([c, h]) => str(1, +c).trim().toUpperCase() !== h).map(([c, h]) => `${XLSX.utils.encode_col(c - 1)} (atteso "${h}", trovato "${str(1, +c)}")`);
    if (wrong.length) throw new Error('Le colonne del Master non sono quelle attese: ' + wrong.join('; ') + '. L\'app va allineata alla nuova struttura.');

    const blocks = (ws['!merges'] || []).filter(m => m.s.c === 0 && m.e.c === 0 && m.s.r >= 1)
      .map(m => [m.s.r + 1, m.e.r + 1]).sort((a, b) => a[0] - b[0]);
    if (!blocks.length) throw new Error('Nessun blocco impianto trovato nel Master (colonna A senza celle unite).');

    const clients = blocks.map(([r0]) => {
      const cliente = val(r0, COL.cliente), regione = val(r0, COL.regione);
      const tipoContratto = str(r0, COL.contratto);
      const [cat, label] = classifyStatus(tipoContratto);
      const coordRaw = val(r0, COL.coord);
      const coord = parseCoord(coordRaw);
      const pmv = (c) => { const v = val(r0, c); return typeof v === 'string' ? v.trim() : v; };
      const c = {
        cliente: cliente || '', regione: regione || '', tipologia: val(r0, COL.tipologia) || '',
        potenza: val(r0, COL.potenza) || 0, pm: pmv(COL.pm) || '', pm2: pmv(COL.pm2) || '',
        pr: val(r0, COL.pr) || '', pr_comment: comment(r0, COL.pr),
        stato_cat: cat, stato_label: label, stato_dettaglio: cat ? tipoContratto.trim() : '',
        lat: coord ? coord[0] : null, lon: coord ? coord[1] : null,
        coord_raw: !coord && coordRaw ? String(coordRaw).trim() : '',
        geo_precision: coord ? 'esatta' : null, riga: r0, servizi: [],
        // campi in più rispetto al vecchio sito, per la scheda impianto
        ettari: val(r0, COL.ettari) || '', pod: str(r0, COL.pod).trim(), referente: str(r0, COL.referente).trim(),
        contratto: tipoContratto.trim(), decorrenza: val(r0, COL.decorrenza) || '', corrispettivo: val(r0, COL.corrispettivo) || '',
        tempi: str(r0, COL.tempi).trim()
      };
      if (c.lat === null) {
        const fb = geoFallback(cliente, regione);
        if (fb) { const [dl, dn] = jitter(r0, JITTER[fb[2]]); c.lat = fb[0] + dl; c.lon = fb[1] + dn; c.geo_precision = fb[2]; }
      }
      for (const [nome, col] of GROUP_COLS) {
        const freq = val(r0, col);
        const previsto = freq === 'Non prevista' ? 'No' : (['Annuale', 'Semestrale', 'Quadrimestrale'].includes(freq) ? 'Sì' : '');
        c.servizi.push({ nome, previsto, frequenza: freq || '', u1: val(r0 + 2, col) || '', u2: val(r0 + 3, col) || '',
          u3: val(r0 + 4, col) || '', subtec: val(r0 + 5, col) || '', contatto: val(r0 + 6, col) || '', note: val(r0 + 7, col) || '' });
      }
      const cctv = val(r0, COL.cctv);
      c.servizi.push({ nome: 'Servizio di CCTV e vigilanza', previsto: cctv === 'Previsto' ? 'Sì' : (cctv === 'Non previsto' ? 'No' : ''),
        frequenza: '', u1: '', u2: '', u3: '', subtec: '', contatto: '', note: comment(r0, COL.cctv) });
      return c;
    });
    return { clients, blocks: blocks.length };
  }

  const api = { parse, SERVICE_META, parseCoord, classifyStatus };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.STATUS_MASTER = api;
})(typeof window !== 'undefined' ? window : globalThis);
