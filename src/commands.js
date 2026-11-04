/* Pulsante "Traduci" nella barra della mail: traduce con un clic, senza aprire pannelli. */
/* global Office, TradOutlook */

function traduciEmail(event) {
  TradOutlook.translateCurrentItem()
    .catch(function () { /* l'errore è già mostrato nella barra notifiche della mail */ })
    .then(function () { event.completed(); });
}

function ripristinaOriginale(event) {
  TradOutlook.restoreOriginal()
    .catch(function (e) {
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("trad", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: e.message.slice(0, 150)
        });
      } catch (x) { /* */ }
    })
    .then(function () { event.completed(); });
}

// Registrazione: globale (manifest XML) + Office.actions (runtime recenti)
window.traduciEmail = traduciEmail;
window.ripristinaOriginale = ripristinaOriginale;
Office.onReady(function () {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("traduciEmail", traduciEmail);
    Office.actions.associate("ripristinaOriginale", ripristinaOriginale);
  }
});
