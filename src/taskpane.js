/* global Office, TradCore, TradOutlook */
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };

  function status(text, kind) {
    var el = $("status");
    el.textContent = text || "";
    el.className = kind || "";
  }

  function fill(select, map) {
    Object.keys(map).forEach(function (k) {
      var o = document.createElement("option");
      o.value = k;
      o.textContent = typeof map[k] === "string" ? map[k] : map[k].label;
      select.appendChild(o);
    });
  }

  function readForm(base) {
    return Object.assign({}, base, {
      apiKey: $("apiKey").value.trim(),
      model: $("model").value,
      target: $("target").value,
      style: $("style").value,
      glossary: $("glossary").value,
      instructions: $("instructions").value,
      excludeSignature: $("excludeSignature").checked,
      translateSubject: $("translateSubject").checked
    });
  }

  function writeForm(s) {
    $("apiKey").value = s.apiKey;
    $("model").value = s.model;
    $("target").value = s.target;
    $("style").value = s.style;
    $("glossary").value = s.glossary;
    $("instructions").value = s.instructions;
    $("excludeSignature").checked = !!s.excludeSignature;
    $("translateSubject").checked = !!s.translateSubject;
  }

  function persist(quiet) {
    var s = readForm(TradOutlook.loadSettings());
    return TradOutlook.saveSettings(s).then(function () {
      if (!quiet) status("Impostazioni salvate.", "ok");
      return s;
    }).catch(function (e) { status("Salvataggio non riuscito: " + e.message, "err"); throw e; });
  }

  function busy(on) {
    ["btnTranslate", "btnUndo", "btnTest", "btnSave"].forEach(function (id) { $(id).disabled = on; });
  }

  Office.onReady(function () {
    fill($("target"), TradCore.LANGUAGE_LABELS);
    fill($("model"), TradCore.MODELS);
    var s = TradOutlook.loadSettings();
    writeForm(s);
    if (!s.apiKey) {
      $("settings").open = true;
      status("Prima di iniziare incolla la chiave API nelle Impostazioni.", "busy");
    }

    // le scelte rapide si salvano da sole
    ["target", "excludeSignature", "translateSubject"].forEach(function (id) {
      $(id).addEventListener("change", function () { persist(true); });
    });

    $("btnTranslate").addEventListener("click", function () {
      busy(true);
      status("Traduzione in corso…", "busy");
      persist(true)
        .then(function (cur) { return TradOutlook.translateCurrentItem(cur); })
        .then(function (msg) { status(msg + ".", "ok"); })
        .catch(function (e) { status(e.message, "err"); })
        .then(function () { busy(false); });
    });

    $("btnUndo").addEventListener("click", function () {
      busy(true);
      TradOutlook.restoreOriginal()
        .then(function () { status("Testo originale ripristinato.", "ok"); })
        .catch(function (e) { status(e.message + " Prova con Cmd/Ctrl+Z.", "err"); })
        .then(function () { busy(false); });
    });

    $("btnSave").addEventListener("click", function () { persist(false); });

    $("btnShowKey").addEventListener("click", function () {
      var k = $("apiKey");
      k.type = k.type === "password" ? "text" : "password";
    });

    $("btnTest").addEventListener("click", function () {
      var cur = readForm(TradOutlook.loadSettings());
      busy(true);
      status("Verifica in corso…", "busy");
      var req = TradCore.buildRequest(cur, "text", "Ciao, come stai?", null);
      req.max_tokens = 64;
      TradCore.callClaude(cur, req)
        .then(function () { return persist(true); })
        .then(function () { status("Chiave valida, impostazioni salvate.", "ok"); })
        .catch(function (e) { status(e.message, "err"); })
        .then(function () { busy(false); });
    });
  });
})();
