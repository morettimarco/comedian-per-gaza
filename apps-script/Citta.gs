/**
 * Dati per la landing page delle città (citta.html).
 *
 * Da incollare nel progetto Apps Script che serve il form, accanto a
 * getConfig() e submitIscrizione(). Va poi aggiunta all'oggetto
 * `consentite` di doPost(), altrimenti l'API risponde
 * "Funzione non consentita: getCitta":
 *
 *   var consentite = {
 *     getConfig: getConfig,
 *     submitIscrizione: submitIscrizione,
 *     getCitta: getCitta
 *   };
 *
 * Infine ripubblica il deployment (Gestisci deployment → Modifica →
 * Versione: Nuova versione): salvare il codice non aggiorna l'URL /exec.
 *
 * Legge il foglio "Città" cercando le colonne per intestazione, quindi
 * l'ordine delle colonne può cambiare. Restituisce SOLO i campi pubblici:
 * "Persone da taggare" e "Note" restano nel foglio.
 *
 * Colonne facoltative lette se presenti:
 *   - "Link Form"             → URL del form per i comici di quella città
 *   - "Contatti Prenotazione" → telefono, mail, link o @instagram per prenotare
 */
function getCitta() {
  var foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Città');
  if (!foglio) throw new Error('Foglio "Città" non trovato');

  // getDisplayValues: le date arrivano come le vedi nel foglio (es. "22/11")
  var righe = foglio.getDataRange().getDisplayValues();
  var intest = righe.shift().map(function (h) { return String(h).trim().toLowerCase(); });

  var campi = {
    citta:      'città',
    data:       'data',
    referente:  'nome referente',
    igReferente:'instagram referente',
    locale:     'nome locale',
    indirizzo:  'indirizzo locale',
    igLocale:   'instagram locale',
    sito:       'sito locale',
    form:       'link form',
    contatti:   'contatti prenotazione'
  };
  var idx = {};
  Object.keys(campi).forEach(function (k) { idx[k] = intest.indexOf(campi[k]); });

  return righe
    .filter(function (r) { return idx.citta >= 0 && String(r[idx.citta]).trim(); })
    .map(function (r) {
      var o = {};
      Object.keys(idx).forEach(function (k) {
        o[k] = idx[k] >= 0 ? String(r[idx[k]]).trim() : '';
      });
      return o;
    });
}
