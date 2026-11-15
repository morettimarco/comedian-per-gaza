/**
 * Dati per la pagina degli MC (mc.html): rubrica dei comici in line up.
 *
 * Restituisce nome, telefono, Instagram e fascia dei comici in scaletta nel
 * foglio "Elenco Comici" (posizioni 1-6). MC e headliner NON sono inclusi.
 * Sono dati personali: la funzione risponde SOLO con il codice MC giusto.
 *
 * INSTALLAZIONE (una volta sola)
 * 1. Incolla questo file nel progetto Apps Script del form, accanto a Citta.gs.
 * 2. Impostazioni progetto (ingranaggio) → Proprietà script →
 *    Aggiungi proprietà: CODICE_MC = <il codice da dare agli MC>
 *    Il codice NON sta nel sorgente né su GitHub (il repo è pubblico).
 * 3. Aggiungi `getElencoMC: getElencoMC` all'oggetto `consentite` di doPost().
 * 4. Ripubblica: Gestisci deployment → Modifica → Versione: Nuova versione.
 * Per cambiare codice basta modificare CODICE_MC, senza ripubblicare.
 *
 * COME LEGGE "Elenco Comici" (colonne cercate per intestazione)
 * Ogni fascia è un blocco di righe: la fascia (colonna "doll" o "Fascia")
 * sta solo sulla PRIMA riga del blocco; sotto ci sono "POS IN LINE UP"
 * (1-6, Headliner) e "Nome".
 * - la fascia si trascina in giù finché "POS IN LINE UP" è compilato;
 *   una riga senza posizione chiude il blocco (note in fondo al foglio escluse)
 * - entrano solo le posizioni numeriche: headliner e MC restano fuori
 * - righe con il nome tutto maiuscolo e senza contatti (es. "CHIUSURA SERATA")
 *   sono segnaposto e vengono saltate
 * - telefono o Instagram mancanti vengono cercati per nome in "Iscrizioni"
 */

var MC_MAX_ERRORI = 10;        // tentativi sbagliati tollerati...
var MC_FINESTRA_SEC = 15 * 60; // ...in 15 minuti, poi blocco per 15 minuti

function getElencoMC(codice) {
  verificaCodiceMC_(codice);

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var foglio = libro.getSheetByName('Elenco Comici');
  if (!foglio) throw new Error('Foglio "Elenco Comici" non trovato');

  // getDisplayValues: niente apostrofi di txt(), fasce come "14-15" e non date
  var righe = foglio.getDataRange().getDisplayValues();
  var intest = righe.shift().map(function (h) { return String(h).trim().toLowerCase(); });
  var col = function (nomi) {
    for (var i = 0; i < nomi.length; i++) { var j = intest.indexOf(nomi[i]); if (j >= 0) return j; }
    return -1;
  };
  var c = {
    fascia:    col(['fascia', 'doll', 'orario assegnato', 'orario']),
    posizione: col(['pos in line up', 'posizione']),
    nome:      col(['nome']),
    stato:     col(['stato']),
    telefono:  col(['telefono']),
    social:    col(['social', 'instagram'])
  };
  if (c.nome < 0) throw new Error('Colonna "Nome" non trovata in Elenco Comici');
  if (c.posizione < 0) throw new Error('Colonna "POS IN LINE UP" non trovata in Elenco Comici');
  if (c.fascia < 0) c.fascia = 0; // nel 2026 la fascia è la prima colonna
  var v = function (r, k) { return c[k] >= 0 ? String(r[c[k]] == null ? '' : r[c[k]]).trim() : ''; };

  var rubrica = rubricaIscrizioni_(libro);
  var persone = [];
  var visti = {};

  var fascia = '';
  righe.forEach(function (r) {
    var f = v(r, 'fascia');
    var pos = v(r, 'posizione');
    if (f) fascia = f;
    else if (!pos) { fascia = ''; return; } // fine blocco
    if (!fascia || !/^\d+$/.test(pos)) return; // solo line up: niente headliner

    var nome = v(r, 'nome');
    var tel = v(r, 'telefono');
    var ig = v(r, 'social');
    if (!nome) return;
    if (!tel && !ig && nome === nome.toUpperCase() && /[A-Z]{4}/.test(nome)) return; // segnaposto

    var k = chiaveNome_(nome);
    if (visti[k]) return;
    visti[k] = true;
    var form = rubrica[k] || {};
    persone.push({ nome: nome, fascia: fascia, posizione: pos, stato: v(r, 'stato'),
                   telefono: tel || form.telefono || '', instagram: ig || form.social || '' });
  });

  return persone;
}

/** { chiaveNome: {telefono, social} } dal foglio "Iscrizioni" (se c'è) */
function rubricaIscrizioni_(libro) {
  var out = {};
  var f = libro.getSheetByName('Iscrizioni');
  if (!f) return out;
  var righe = f.getDataRange().getDisplayValues();
  var intest = righe.shift().map(function (h) { return String(h).trim().toLowerCase(); });
  var n = intest.indexOf('nome'), t = intest.indexOf('telefono'), s = intest.indexOf('social');
  if (n < 0) return out;
  righe.forEach(function (r) {
    var k = chiaveNome_(r[n]);
    if (k) out[k] = { telefono: t >= 0 ? String(r[t]).trim() : '', social: s >= 0 ? String(r[s]).trim() : '' };
  });
  return out;
}

/** "Beppe Freda (Giuseppe Ciuffreda)" / "beppe  fréda" -> "beppe freda" */
function chiaveNome_(s) {
  return String(s || '').replace(/\(.*?\)/g, ' ').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

/** Lancia un errore se il codice è sbagliato; dopo troppi errori blocca tutti per un po'. */
function verificaCodiceMC_(codice) {
  var atteso = PropertiesService.getScriptProperties().getProperty('CODICE_MC');
  if (!atteso) throw new Error('Codice MC non ancora configurato: avvisa gli organizzatori.');

  var cache = CacheService.getScriptCache();
  var errori = Number(cache.get('mc_errori') || 0);
  if (errori >= MC_MAX_ERRORI) {
    throw new Error('Troppi tentativi sbagliati. Riprova fra 15 minuti.');
  }

  var dato = String(codice == null ? '' : codice).trim();
  if (dato.toLowerCase() !== String(atteso).trim().toLowerCase()) {
    cache.put('mc_errori', String(errori + 1), MC_FINESTRA_SEC);
    Utilities.sleep(800); // rallenta chi prova a indovinare
    throw new Error('Codice MC non valido');
  }
}
