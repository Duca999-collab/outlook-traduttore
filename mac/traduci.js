/*
 * Traduttore Mail per macOS
 * Seleziona il testo in qualsiasi app → scorciatoia da tastiera → il testo viene
 * sostituito dalla traduzione, mantenendo la formattazione.
 *
 * Avvio (da Comandi Rapidi → "Esegui script shell"):
 *   osascript -l JavaScript "$HOME/.local/share/traduttore-mail/traduci.js"
 * Modalità diagnostica (non tocca la tastiera, traduce gli appunti):
 *   osascript -l JavaScript .../traduci.js appunti
 *
 * Avviato da Comandi Rapidi, lo script NON invia direttamente i tasti ⌘C/⌘V
 * (macOS li attribuirebbe a osascript): costruisce una volta la piccola app
 * "Traduttore Mail.app" e la lancia in background. Il permesso di Accessibilità
 * si dà solo a quell'app. L'app carica questo stesso file a ogni esecuzione,
 * quindi gli aggiornamenti non richiedono di ricostruirla.
 */
ObjC.import("AppKit");
ObjC.import("Foundation");

var window = this; // per core.js / clip-engine.js
var DIR = ObjC.unwrap($("~/.local/share/traduttore-mail").stringByExpandingTildeInPath);
var IN_APPLET = (typeof __TRAD_APPLET !== "undefined");
var APPLET = DIR + "/Traduttore Mail.app";
var APPLET_VERSION = "1";
var APPLET_LOADER = [
  "var __TRAD_APPLET = true;",
  "ObjC.import('Foundation');",
  "var __p = ObjC.unwrap($('~/.local/share/traduttore-mail/traduci.js').stringByExpandingTildeInPath);",
  "eval(ObjC.unwrap($.NSString.stringWithContentsOfFileEncodingError(__p, $.NSUTF8StringEncoding, null)));",
  "function run(argv) { return __tradRun(argv); }"
].join("\n");
var KC_SERVICE = "traduttore-mail", KC_ACCOUNT = "anthropic";
var app = Application.currentApplication();
app.includeStandardAdditions = true;

function readFile(p) {
  var s = $.NSString.stringWithContentsOfFileEncodingError(p, $.NSUTF8StringEncoding, null);
  return s.isNil() ? null : ObjC.unwrap(s);
}
function writeFile(p, txt) {
  $(txt).writeToFileAtomicallyEncodingError(p, true, $.NSUTF8StringEncoding, null);
}
function q(s) { return "'" + String(s).replace(/'/g, "'\\''") + "'"; }
function sh(cmd) { return app.doShellScript(cmd, { alteringLineEndings: false }); }
function log(msg) {
  try {
    var line = new Date().toISOString() + "  " + msg + "\n";
    var p = DIR + "/log.txt";
    var old = readFile(p) || "";
    writeFile(p, (old + line).split("\n").slice(-200).join("\n"));
  } catch (e) { /* */ }
}
function notify(msg, sub) {
  try { app.displayNotification(msg, { withTitle: "Traduttore Mail", subtitle: sub || "" }); } catch (e) { /* */ }
}

eval(readFile(DIR + "/core.js"));
eval(readFile(DIR + "/clip-engine.js"));

// ---------------------------------------------------------------- impostazioni
function loadSettings() {
  var s = {};
  var raw = readFile(DIR + "/config.json");
  if (raw) { try { s = JSON.parse(raw); } catch (e) { log("config.json non valido: " + e.message); } }
  return Object.assign({}, TradCore.DEFAULT_SETTINGS, s);
}

function getKey() {
  try { return sh("security find-generic-password -s " + KC_SERVICE + " -a " + KC_ACCOUNT + " -w").trim(); }
  catch (e) { return null; }
}

function askKey() {
  app.activate();
  var r = app.displayDialog(
    "Incolla la chiave API Anthropic (inizia con sk-ant-).\nViene salvata nel Portachiavi di questo Mac.",
    { defaultAnswer: "", hiddenAnswer: true, buttons: ["Annulla", "Salva"], defaultButton: "Salva",
      cancelButton: "Annulla", withTitle: "Traduttore Mail" });
  var k = (r.textReturned || "").trim();
  if (!/^sk-ant-/.test(k)) throw new Error("La chiave deve iniziare con sk-ant-");
  sh("security add-generic-password -U -s " + KC_SERVICE + " -a " + KC_ACCOUNT + " -w " + q(k));
  return k;
}

// ---------------------------------------------------------------- API
function callApi(req, key) {
  var tmp = ObjC.unwrap($.NSTemporaryDirectory()) + "traduttore-" + Date.now();
  writeFile(tmp + ".h", "x-api-key: " + key + "\ncontent-type: application/json\nanthropic-version: 2023-06-01\n");
  writeFile(tmp + ".json", JSON.stringify(req));
  var out;
  try {
    out = sh("chmod 600 " + q(tmp + ".h") + "; curl -sS --max-time 180 -H @" + q(tmp + ".h") +
             " --data-binary @" + q(tmp + ".json") + " -w '\\n%{http_code}' https://api.anthropic.com/v1/messages");
  } finally {
    try { sh("rm -f " + q(tmp + ".h") + " " + q(tmp + ".json")); } catch (e) { /* */ }
  }
  var nl = out.lastIndexOf("\n");
  var code = parseInt(out.slice(nl + 1), 10);
  var json = null;
  try { json = JSON.parse(out.slice(0, nl)); } catch (e) { /* */ }
  if (code !== 200) {
    var msg = (json && json.error && json.error.message) || ("HTTP " + code);
    if (code === 401) msg = "chiave API non valida o revocata";
    else if (code === 429) msg = "troppe richieste, riprova tra poco";
    else if (code === 529 || code === 503) msg = "servizio sovraccarico, riprova";
    else if (/credit/i.test(msg)) msg = "credito API esaurito (console.anthropic.com → Billing)";
    var err = new Error(msg); err.code = code; throw err;
  }
  if (json.stop_reason === "max_tokens") throw new Error("testo troppo lungo: traducilo a pezzi");
  return (json.content || []).filter(function (c) { return c.type === "text"; })
    .map(function (c) { return c.text; }).join("");
}

// ---------------------------------------------------------------- appunti
var pb = $.NSPasteboard.generalPasteboard;

function pbString(type) {
  var s = pb.stringForType(type);
  return s.isNil() ? null : ObjC.unwrap(s);
}

function readClipboard() {
  var html = pbString("public.html");
  if (html && /\S/.test(html)) return { kind: "html", html: html };
  var rtf = pb.dataForType("public.rtf");
  if (!rtf.isNil()) {
    var p = ObjC.unwrap($.NSTemporaryDirectory()) + "traduttore-in.rtf";
    rtf.writeToFileAtomically(p, true);
    var conv = sh("textutil -convert html -stdout " + q(p) + "; rm -f " + q(p));
    return { kind: "html", html: conv, fromRtf: true };
  }
  var txt = pbString("public.utf8-plain-text");
  if (txt && /\S/.test(txt)) return { kind: "text", text: txt };
  return null;
}

function writeClipboard(clip) {
  var tmp = ObjC.unwrap($.NSTemporaryDirectory()) + "traduttore-out";
  pb.clearContents;
  if (clip.kind === "html") {
    writeFile(tmp + ".html", clip.html);
    var plain = "", rtfData = null;
    try { plain = sh("textutil -convert txt -stdout -inputencoding UTF-8 " + q(tmp + ".html")); } catch (e) { /* */ }
    try {
      sh("textutil -convert rtf -inputencoding UTF-8 -output " + q(tmp + ".rtf") + " " + q(tmp + ".html"));
      rtfData = $.NSData.dataWithContentsOfFile(tmp + ".rtf");
    } catch (e) { /* */ }
    pb.setStringForType($(clip.html), "public.html");
    if (rtfData && !rtfData.isNil()) pb.setDataForType(rtfData, "public.rtf");
    pb.setStringForType($(plain.replace(/\n$/, "")), "public.utf8-plain-text");
    try { sh("rm -f " + q(tmp + ".html") + " " + q(tmp + ".rtf")); } catch (e) { /* */ }
  } else {
    pb.setStringForType($(clip.text), "public.utf8-plain-text");
  }
}

// ---------------------------------------------------------------- tastiera
var se = Application("System Events");

function frontBundle() {
  try { return se.processes.whose({ frontmost: true })[0].bundleIdentifier(); } catch (e) { return null; }
}

function waitModifiersReleased() {
  // la scorciatoia (es. ⌃⌥T) non deve "sporcare" ⌘C/⌘V
  for (var i = 0; i < 60; i++) {
    if ((Number($.NSEvent.modifierFlags) & 0x1E0000) === 0) return;
    delay(0.05);
  }
}

function copySelection() {
  var before = Number(pb.changeCount);
  se.keystroke("c", { using: "command down" });
  for (var i = 0; i < 40; i++) {
    delay(0.05);
    if (Number(pb.changeCount) !== before) return true;
  }
  return false;
}

// ---------------------------------------------------------------- traduzione
function translate(clip, settings, key) {
  var send = function (req) { return callApi(req, key); };
  if (clip.kind === "html") {
    var r = TradClip.translateHtml(clip.html, settings, send);
    return r.skipped ? null : { kind: "html", html: r.html, boundary: r.boundary };
  }
  var t = TradClip.translatePlain(clip.text, settings, send);
  return t.skipped ? null : { kind: "text", text: t.text, boundary: t.boundary };
}

function openPrivacy(pane) {
  try { sh("open 'x-apple.systempreferences:com.apple.preference.security?" + pane + "'"); } catch (e) { /* */ }
}

function friendly(e) {
  var m = (e && e.message) || String(e);
  var n = e && e.errorNumber;
  if (n === 1002 || /1002|non è consentit|non è autorizzat|not allowed to send keystrokes|assistive/i.test(m)) {
    openPrivacy("Privacy_Accessibility");
    return "Attiva \"Traduttore Mail\" in Privacy e sicurezza → Accessibilità, poi riprova.";
  }
  if (n === -1743 || /-1743|Not authorized to send Apple events|non è autorizzato a inviare eventi/i.test(m)) {
    openPrivacy("Privacy_Automation");
    return "Consenti a \"Traduttore Mail\" di controllare System Events (Privacy e sicurezza → Automazione), poi riprova.";
  }
  return m;
}

// ---------------------------------------------------------------- app dedicata
function ensureApplet() {
  var marker = DIR + "/.applet-version";
  var exists = $.NSFileManager.defaultManager.fileExistsAtPath(APPLET + "/Contents/Info.plist");
  if (exists && (readFile(marker) || "").trim() === APPLET_VERSION) return;
  var src = DIR + "/applet-loader.js";
  writeFile(src, APPLET_LOADER);
  var plist = APPLET + "/Contents/Info.plist";
  sh("rm -rf " + q(APPLET) +
     " && /usr/bin/osacompile -l JavaScript -o " + q(APPLET) + " " + q(src) +
     " && /usr/bin/plutil -replace LSUIElement -bool true " + q(plist) +
     " && /usr/bin/plutil -replace CFBundleIdentifier -string it.riccardo.traduttore-mail " + q(plist) +
     " && /usr/bin/codesign -f -s - " + q(APPLET) + " 2>&1");
  writeFile(marker, APPLET_VERSION);
  log("app Traduttore Mail costruita");
}

function launchApplet() {
  ensureApplet();
  sh("/usr/bin/open -g " + q(APPLET));
}

function __tradRun(argv) {
  var mode = (argv && argv[0]) || "";
  if (!IN_APPLET && mode !== "appunti" && mode !== "diretto") {
    try { launchApplet(); }
    catch (e) { log("ERRORE avvio app: " + ((e && e.message) || e)); notify((e && e.message) || String(e), "Avvio non riuscito"); }
    return;
  }
  return main(argv);
}

function run(argv) { return __tradRun(argv); }

function main(argv) {
  var mode = (argv && argv[0]) || "";
  var t0 = Date.now();
  try {
    var settings = loadSettings();
    var key = getKey();
    if (!key) {
      key = askKey();
      notify("Chiave salvata. Seleziona il testo e ripeti la scorciatoia.");
      log("chiave salvata nel Portachiavi");
      return;
    }

    var bundle = null;
    if (mode !== "appunti") {
      bundle = frontBundle();
      waitModifiersReleased();
      if (!copySelection()) {
        notify("Seleziona prima il testo da tradurre (⌘A per tutto).");
        log("nessuna selezione in " + bundle);
        return;
      }
    }

    var clip = readClipboard();
    if (!clip) { notify("Niente da tradurre negli appunti."); return; }
    var lang = TradCore.LANGUAGE_LABELS[settings.target] || settings.target;
    notify("Traduzione in " + lang + " in corso…");

    var res = translate(clip, settings, key);
    if (!res) { notify("Niente da tradurre (solo firma o thread)."); return; }
    writeClipboard(res);

    var secs = ((Date.now() - t0) / 1000).toFixed(1);
    var note = res.boundary ? (res.boundary === "firma" ? " · firma e thread esclusi" : " · thread escluso") : "";
    if (mode === "appunti") {
      notify("Traduzione pronta negli appunti" + note, secs + " s");
    } else if (frontBundle() !== bundle) {
      notify("Hai cambiato finestra: la traduzione è negli appunti, incollala con ⌘V.", secs + " s");
    } else {
      waitModifiersReleased();
      se.keystroke("v", { using: "command down" });
      notify("Tradotto in " + lang + note, secs + " s");
    }
    log("ok " + (bundle || mode) + " " + clip.kind + (clip.fromRtf ? "(rtf)" : "") + " " + secs + "s" + note);
  } catch (e) {
    if (e && e.errorNumber === -128) return; // Annulla
    var m = friendly(e);
    log("ERRORE " + (e && e.errorNumber ? "[" + e.errorNumber + "] " : "") + ((e && e.message) || e));
    notify(m, "Traduzione non riuscita");
  }
}
