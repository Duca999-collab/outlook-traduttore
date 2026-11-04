/*
 * Ponte tra Outlook (Office.js) e il motore di traduzione.
 * Usato sia dal pulsante "Traduci" (comando a un clic) sia dal pannello Opzioni.
 */
(function (root) {
  "use strict";
  var C = root.TradCore;
  var SETTINGS_KEY = "tradSettings";
  var UNDO_KEY = "tradUndo";

  function p(fn) { // promisifica le API async di Office
    return new Promise(function (resolve, reject) {
      fn(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(new Error(res.error && res.error.message || "Errore Outlook"));
      });
    });
  }

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(Office.context.roamingSettings.get(SETTINGS_KEY) || "{}"); } catch (e) { /* */ }
    return Object.assign({}, C.DEFAULT_SETTINGS, s);
  }

  function saveSettings(s) {
    Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify(s));
    return p(function (cb) { Office.context.roamingSettings.saveAsync(cb); });
  }

  function notify(type, message) {
    var item = Office.context.mailbox.item;
    var T = Office.MailboxEnums.ItemNotificationMessageType;
    var msg = { type: T[type], message: message.slice(0, 150) };
    if (type === "InformationalMessage") { msg.icon = "Icon.16"; msg.persistent = false; }
    return new Promise(function (resolve) {
      try { item.notificationMessages.replaceAsync("trad", msg, function () { resolve(); }); }
      catch (e) { resolve(); }
    });
  }

  function saveUndo(snapshot) {
    try { root.localStorage.setItem(UNDO_KEY, JSON.stringify(snapshot)); } catch (e) { /* */ }
  }
  function getUndo() {
    try { return JSON.parse(root.localStorage.getItem(UNDO_KEY) || "null"); } catch (e) { return null; }
  }

  /**
   * Traduce la mail aperta in composizione.
   * - se c'è testo selezionato nel corpo: traduce solo la selezione
   * - altrimenti: traduce tutto il testo scritto, escludendo firma e thread citato
   * overrides: impostazioni temporanee (es. lingua scelta nel pannello)
   */
  function translateCurrentItem(overrides) {
    var item = Office.context.mailbox.item;
    var s = Object.assign(loadSettings(), overrides || {});
    var started = Date.now();
    var langLabel = (C.LANGUAGE_LABELS[s.target] || s.target);

    return notify("ProgressIndicator", "Traduzione in " + langLabel + " in corso…")
      .then(function () { return p(function (cb) { item.body.getTypeAsync(cb); }); })
      .then(function (bodyType) {
        var isHtml = bodyType === Office.CoercionType.Html || bodyType === "html";
        var coercion = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

        return Promise.all([
          p(function (cb) { item.body.getAsync(coercion, cb); }),
          p(function (cb) { item.subject.getAsync(cb); }),
          p(function (cb) { item.getSelectedDataAsync(Office.CoercionType.Text, cb); }).catch(function () { return null; })
        ]).then(function (r) {
          var fullBody = r[0], subject = r[1], sel = r[2];
          var undo = { bodyType: isHtml ? "html" : "text", body: fullBody, subject: subject, at: Date.now() };
          var hasSelection = sel && sel.data && /\S/.test(sel.data);

          // --- Selezione nell'oggetto
          if (hasSelection && sel.sourceProperty === "subject") {
            return C.translateText(sel.data, s, { fragment: true }).then(function (res) {
              saveUndo(undo);
              return p(function (cb) { item.setSelectedDataAsync(res.text, { coercionType: Office.CoercionType.Text }, cb); })
                .then(function () { return "oggetto selezionato"; });
            });
          }

          // --- Selezione nel corpo
          if (hasSelection) {
            var getSel = isHtml
              ? p(function (cb) { item.getSelectedDataAsync(Office.CoercionType.Html, cb); })
              : Promise.resolve(sel);
            return getSel.then(function (selData) {
              var job = isHtml
                ? C.translateHtml(selData.data, s, { fragment: true }).then(function (x) { return x.html; })
                : C.translateText(selData.data, s, { fragment: true }).then(function (x) { return x.text; });
              return job.then(function (out) {
                if (out === null) return "niente da tradurre nella selezione";
                saveUndo(undo);
                return p(function (cb) { item.setSelectedDataAsync(out, { coercionType: coercion }, cb); })
                  .then(function () { return "testo selezionato"; });
              });
            });
          }

          // --- Mail intera. L'oggetto si traduce solo per mail nuove (non risposte/inoltri)
          var isNew = !item.conversationId;
          var subj = s.translateSubject && isNew && subject && /[A-Za-zÀ-ÿ]/.test(subject) ? subject : null;

          var job = isHtml
            ? C.translateHtml(fullBody, s, { subject: subj })
            : C.translateText(fullBody, s, { subject: subj });

          return job.then(function (res) {
            if (res.skipped) return "niente da tradurre (solo firma o thread)";
            saveUndo(undo);
            var newBody = isHtml ? res.html : res.text;
            var steps = [p(function (cb) { item.body.setAsync(newBody, { coercionType: coercion }, cb); })];
            if (subj && res.subject && res.subject !== subject) {
              steps.push(p(function (cb) { item.subject.setAsync(res.subject, cb); }));
            }
            return Promise.all(steps).then(function () {
              return "mail" + (subj ? " e oggetto" : "") +
                (res.boundary ? " (" + (res.boundary === "firma" ? "firma e thread esclusi" : "thread escluso") + ")" : "");
            });
          });
        });
      })
      .then(function (what) {
        var secs = ((Date.now() - started) / 1000).toFixed(1);
        var done = /^niente/.test(what) ? what : "Tradotto in " + langLabel + ": " + what + " · " + secs + " s";
        return notify("InformationalMessage", done).then(function () { return done; });
      })
      .catch(function (err) {
        var m = (err && err.message) || String(err);
        return notify("ErrorMessage", "Traduzione non riuscita: " + m).then(function () { throw err; });
      });
  }

  function restoreOriginal() {
    var item = Office.context.mailbox.item;
    var u = getUndo();
    if (!u) return Promise.reject(new Error("Nessuna versione originale salvata su questo dispositivo."));
    var coercion = u.bodyType === "html" ? Office.CoercionType.Html : Office.CoercionType.Text;
    return Promise.all([
      p(function (cb) { item.body.setAsync(u.body, { coercionType: coercion }, cb); }),
      p(function (cb) { item.subject.setAsync(u.subject || "", cb); })
    ]).then(function () { return notify("InformationalMessage", "Testo originale ripristinato."); });
  }

  root.TradOutlook = {
    loadSettings: loadSettings, saveSettings: saveSettings,
    translateCurrentItem: translateCurrentItem, restoreOriginal: restoreOriginal, getUndo: getUndo
  };
})(window);
