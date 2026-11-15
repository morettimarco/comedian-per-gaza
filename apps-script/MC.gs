/**
 * Dati per la pagina degli MC (mc.html): rubrica di chi è in scaletta.
 *
 * Restituisce nome, telefono, Instagram e fascia di tutte le persone del
 * foglio "Elenco Comici" che hanno una fascia assegnata (come comico o come MC).
 * Sono dati personali: la funzione risponde SOLO con il codice MC giusto.
 *
 * INSTALLAZIONE (una volta sola)
 * 1. Incolla questo file nel progetto Apps Script del form, accanto a Citta.gs.
 * 2. Scegli il codice da dare agli MC e salvalo nelle proprietà dello script:
 *      Impostazioni progetto (ingranaggio) → Proprietà script →
 *      Aggiungi proprietà: CODICE_MC = <il codice>
 *    Il codice NON sta nel codice sorgente né su GitHub (il repo è pubblico).
 * 3. Aggiungi la funzione all'oggetto `consentite` di doPost():
 *
 *      var consentite = {
 *        getConfig: getConfig,
 *        submitIscrizione: submitIscrizione,
 *        getCitta: getCitta,
 *        getElencoMC: getElencoMC
 *      };
 *
 * 4. Ripubblica: Gestisci deployment → Modifica → Versione: Nuova versione.
 *
 * Per cambiare codice (es. se gira troppo) basta modificare CODICE_MC:
 * non serve ripubblicare, e i vecchi link smettono di funzionare subito.
 *
 * Colonne cercate per intestazione (l'ordine può cambiare):
 *   Nome | Orario assegnato | Stato | MC | Orario MC | POS IN LINE UP |
 *   Telefono | Social
 */

var MC_MAX_ERRORI = 10;        // tentativi sbagliati tollerati...
var MC_FINESTRA_SEC = 15 * 60; // ...in 15 minuti, poi blocco per 15 minuti

function getElencoMC(codice) {
  verificaCodiceMC_(codice);

  var foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elenco Comici');
  if (!foglio) throw new Error('Foglio "Elenco Comici" non trovato');

  // getDisplayValues: niente apostrofi di txt(), fasce come "14-15" e non date
  var righe = foglio.getDataRange().getDisplayValues();
  var intest = righe.shift().map(function (h) { return String(h).trim().toLowerCase(); });

  var campi = {
    nome:      'nome',
    fascia:    'orario assegnato',
    stato:     'stato',
    mc:        'mc',
    orarioMC:  'orario mc',
    posizione: 'pos in line up',
    telefono:  'telefono',
    instagram: 'social'
  };
  var idx = {};
  Object.keys(campi).forEach(function (k) { idx[k] = intest.indexOf(campi[k]); });
  if (idx.nome < 0) throw new Error('Colonna "Nome" non trovata in Elenco Comici');

  var leggi = function (r, k) { return idx[k] >= 0 ? String(r[idx[k]]).trim() : ''; };
  var vero = /^(true|vero|si|sì|x|1)$/i;

  return righe
    .map(function (r) {
      var mc = vero.test(leggi(r, 'mc'));
      var orarioMC = leggi(r, 'orarioMC');
      return {
        nome:      leggi(r, 'nome'),
        fascia:    leggi(r, 'fascia'),
        posizione: leggi(r, 'posizione'),
        // "Orario MC" conta se la casella MC è spuntata oppure se è compilato
        orarioMC:  (mc || orarioMC) ? orarioMC : '',
        mc:        mc,
        stato:     leggi(r, 'stato'),
        telefono:  leggi(r, 'telefono'),
        instagram: leggi(r, 'instagram')
      };
    })
    // solo chi è pianificato: ha un nome e almeno una fascia
    .filter(function (p) { return p.nome && (p.fascia || p.orarioMC); });
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
