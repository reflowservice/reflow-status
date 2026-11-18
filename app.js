/* REFLOW STATUS — avvio: accesso, lettura del Master, rendering.
   La grafica e la logica di calcolo (scadenze, copertura, mappe, filtri) stanno in RUN() dentro index.html. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const gate = $('gate'), msg = $('gateMsg'), btn = $('gateBtn');
  // Battito della linea a impulso (deve coincidere con --beat nel CSS di .gate-pulse)
  const BEAT_MS = 1600;
  // Dopo "Aggiorna" la pagina si ricarica: la schermata col polso resta per 3 battiti e poi si chiude da sola
  let refresh = false;
  try { refresh = sessionStorage.getItem('status_refresh') === '1'; sessionStorage.removeItem('status_refresh'); } catch (e) {}
  const say = (t, err) => { msg.textContent = t; msg.classList.toggle('err', !!err); };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const fmtData = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' alle ' + d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' }); };

  // Dati anagrafici e contrattuali in testa alla scheda impianto (campi del blocco nel Master)
  window.anagraficaHtml = function (c) {
    const num = (v, dec) => typeof v === 'number' ? v.toLocaleString('it-IT', { maximumFractionDigits: dec }) : String(v || '').trim();
    const euro = v => typeof v === 'number' ? v.toLocaleString('it-IT', { style: 'currency', currency: 'EUR' }) : String(v || '').trim();
    const items = [
      ['Potenza', c.potenza ? num(c.potenza, 2) + ' kWp' : ''],
      ['Ettari', c.ettari !== '' ? num(c.ettari, 2) : ''],
      ['POD', c.pod, 'mono'],
      ['PM', c.pm], ['PM seconda linea', c.pm2],
      ['Tipologia di contratto', c.nonContrattualizzato ? '' : c.contratto],
      ['Decorrenza', c.decorrenza], ['Corrispettivo annuo', euro(c.corrispettivo)],
      ["Tempi d'intervento", c.tempi],
      ['Referente impianto', c.referente, 'wide']
    ].filter(x => x[1] !== '' && x[1] != null);
    if (!items.length) return '';
    return '<div class="ana-grid">' + items.map(([k, v, cls]) =>
      `<div class="ana-item${cls === 'wide' ? ' wide' : ''}"><div class="ana-k">${esc(k)}</div><div class="ana-v${cls === 'mono' ? ' mono' : ''}">${esc(v)}</div></div>`).join('') + '</div>';
  };

  function done() {
    gate.classList.add('out');
    setTimeout(() => { gate.hidden = true; }, 500);
  }

  async function load() {
    btn.hidden = true;
    say(refresh ? 'Aggiornamento dal REFLOW MASTER…' : 'Lettura del REFLOW MASTER…');
    try {
      const m = await STATUS.loadMaster();
      say('Elaborazione dei dati…');
      const { clients } = STATUS_MASTER.parse(XLSX, m.data);
      document.querySelectorAll('.n-clienti').forEach(n => { n.textContent = clients.length; });
      $('genDate').textContent = 'Master aggiornato il ' + fmtData(m.modified) + (m.by ? ' da ' + m.by : '');
      if (m.url) $('bMaster').href = m.url; else $('bMaster').hidden = true;
      RUN(clients, STATUS_MASTER.SERVICE_META);
      rememberPm();
      const quando = 'Master aggiornato il ' + fmtData(m.modified);
      if (refresh) {
        // chiude alla fine del terzo battito dall'apertura della pagina
        say(quando);
        const an = (document.querySelector('.gate-pulse path').getAnimations || (() => []))
          .call(document.querySelector('.gate-pulse path'))[0];
        const t = an && an.currentTime != null ? an.currentTime : performance.now();
        const fine = Math.ceil(Math.max(t, 3 * BEAT_MS - 1) / BEAT_MS) * BEAT_MS;
        setTimeout(done, fine - t);
      } else {
        // la schermata iniziale resta finché non si clicca
        say(quando);
        btn.textContent = 'Entra'; btn.hidden = false; btn.onclick = done; btn.focus();
      }
    } catch (e) {
      console.error(e);
      say('Non riesco a leggere il Master.\n' + (e && e.message ? e.message : e), true);
      btn.textContent = 'Riprova'; btn.hidden = false; btn.onclick = () => location.reload();
    }
  }

  // PM scelto ricordato su questo dispositivo, come in ORBIT
  function rememberPm() {
    const sel = $('pmSelect'); if (!sel) return;
    let saved = ''; try { saved = localStorage.getItem('status_pm') || ''; } catch (e) {}
    if (saved && [...sel.options].some(o => o.value === saved)) { sel.value = saved; sel.dispatchEvent(new Event('change', { bubbles: true })); }
    sel.addEventListener('change', () => { try { localStorage.setItem('status_pm', sel.value); } catch (e) {} });
    $('resetBtn') && $('resetBtn').addEventListener('click', () => { try { localStorage.removeItem('status_pm'); } catch (e) {} });
  }

  $('bReload').addEventListener('click', () => { try { sessionStorage.setItem('status_refresh', '1'); } catch (e) {} location.reload(); });
  $('bOut').addEventListener('click', () => STATUS.logout());

  STATUS.ready.then(ok => {
    if (!ok) { say(''); btn.textContent = 'Accedi'; btn.hidden = false; btn.onclick = () => STATUS.login(); $('bOut').hidden = true; return; }
    load();
  }).catch(e => { console.error(e); say('Accesso non riuscito: ' + (e && e.message ? e.message : e), true); btn.textContent = 'Accedi'; btn.hidden = false; btn.onclick = () => STATUS.login(); });
})();
