/*
 * Traduttore Mail per macOS — motore per HTML copiato negli appunti.
 * Non usa il DOM (osascript non ne ha): lavora su token.
 *  - testa (<html><head>…<body…>) e coda (</body></html>) restano fuori dalla richiesta
 *  - thread citato / firma: tutto ciò che segue il primo marcatore resta fuori
 *  - ogni tag d'apertura diventa <tag k="n"> e torna identico all'originale
 *  - commenti, <style>, <script> diventano segnaposto <x-k x="n"></x-k>
 * Riusa prompt e parametri API di core.js (TradCore).
 */
(function (root) {
  "use strict";
  var C = root.TradCore;

  var TOKEN_RE = /<!--[\s\S]*?-->|<(style|script|xml|title)\b[^>]*>[\s\S]*?<\/\1\s*>|<![^>]*>|<\/?[A-Za-z][^>]*>|[^<]+|</gi;

  // Marcatori di inizio thread citato (su tag)
  var QUOTE_TAG_RE = /\b(id\s*=\s*["']?(x_)?(appendonsend|divRplyFwdMsg|mail-editor-reference-message-container|OLK_SRC_BODY_SECTION)\b|class\s*=\s*["'][^"']*\b(gmail_quote|gmail_quote_container|moz-cite-prefix)\b|type\s*=\s*["']?cite\b|border-top:\s*solid\s+#(E1E1E1|B5C4DF))/i;
  // Marcatori di firma
  var SIG_TAG_RE = /\b(id\s*=\s*["']?(x_)?Signature\b|class\s*=\s*["'][^"']*\bgmail_signature\b|name\s*=\s*["']?_MailAutoSig\b)/i;
  // Marcatori di thread nel testo
  var QUOTE_TEXT_RE = /^\s*(-{3,}\s*(Original Message|Messaggio originale)|(Il giorno|On)\s.{5,200}(ha scritto|wrote):\s*$|(Da|From):\s.+(Inviato|Sent|Data|Date):)/i;

  function tokenize(html) {
    var out = [], m;
    TOKEN_RE.lastIndex = 0;
    while ((m = TOKEN_RE.exec(html))) out.push(m[0]);
    return out;
  }

  function tagName(tok) {
    var m = /^<\/?([A-Za-z][\w:.-]*)/.exec(tok);
    return m ? m[1].toLowerCase() : null;
  }

  /** Divide il documento copiato in testa / corpo / coda */
  function splitDocument(html) {
    var head = "", tail = "", body = html;
    var b = /<body\b[^>]*>/i.exec(html);
    if (b) {
      head = html.slice(0, b.index + b[0].length);
      body = html.slice(b.index + b[0].length);
      var e = /<\/body\s*>[\s\S]*$/i.exec(body);
      if (e) { tail = e[0]; body = body.slice(0, e.index); }
    } else {
      // frammento senza <body>: tieni fuori un eventuale <meta>/<html> iniziale
      var lead = /^(\s*(<!DOCTYPE[^>]*>|<html\b[^>]*>|<head\b[\s\S]*?<\/head>|<meta\b[^>]*>)\s*)+/i.exec(html);
      if (lead) { head = lead[0]; body = html.slice(lead[0].length); }
      var trail = /(\s*<\/html>\s*)$/i.exec(body);
      if (trail) { tail = trail[0]; body = body.slice(0, trail.index); }
    }
    return { head: head, body: body, tail: tail };
  }

  function prepare(html, opts) {
    opts = opts || {};
    var doc = splitDocument(html);
    var toks = tokenize(doc.body);
    var tags = [], frozen = [], parts = [];
    var cut = toks.length, boundary = null;

    for (var i = 0; i < toks.length; i++) {
      var t = toks[i];
      if (t.charAt(0) === "<" && t.charAt(1) !== "/" && t.charAt(1) !== "!") {
        if (QUOTE_TAG_RE.test(t)) { cut = i; boundary = "citazione"; break; }
        if (opts.excludeSignature && SIG_TAG_RE.test(t)) { cut = i; boundary = "firma"; break; }
      } else if (t.charAt(0) !== "<" && QUOTE_TEXT_RE.test(t)) {
        cut = i; boundary = "citazione"; break;
      }
    }

    for (var j = 0; j < cut; j++) {
      var tok = toks[j];
      var name = tagName(tok);
      if (/^<!--/.test(tok) || /^<!/.test(tok) || /^<(style|script|xml|title)\b/i.test(tok)) {
        parts.push('<x-k x="' + frozen.length + '"></x-k>');
        frozen.push(tok);
      } else if (name && tok.charAt(1) === "/") {
        parts.push("</" + name + ">");
      } else if (name) {
        var selfClose = /\/\s*>$/.test(tok);
        parts.push("<" + name + ' k="' + tags.length + '"' + (selfClose ? " /" : "") + ">");
        tags.push(tok);
      } else {
        parts.push(tok === "<" ? "&lt;" : tok);
      }
    }

    var payload = parts.join("");
    var rest = toks.slice(cut).join("");
    return {
      head: doc.head, tail: rest + doc.tail, payload: payload.trim(),
      leading: (payload.match(/^\s*/) || [""])[0], trailing: (payload.match(/\s*$/) || [""])[0],
      tags: tags, frozen: frozen, boundary: boundary,
      hasText: /[A-Za-zÀ-ÿ]/.test(payload.replace(/<[^>]*>/g, "").replace(/&[a-z]+;/gi, ""))
    };
  }

  function StructureError(msg) { this.name = "StructureError"; this.message = "Struttura alterata: " + msg; }
  StructureError.prototype = Object.create(Error.prototype);

  function restore(state, translated) {
    // 1) segnaposto: ognuno esattamente una volta
    var seen = {}, re = /<x-k\s+x="(\d+)"\s*>\s*<\/x-k>/gi, m;
    while ((m = re.exec(translated))) {
      if (seen[m[1]]) throw new StructureError("segnaposto duplicato " + m[1]);
      seen[m[1]] = true;
    }
    for (var f = 0; f < state.frozen.length; f++) {
      if (!seen[f]) throw new StructureError("segnaposto mancante " + f);
    }
    // 2) tag con k → tag originale (solo se il nome coincide)
    var out = translated.replace(/<([A-Za-z][\w:.-]*)\s+k="(\d+)"\s*\/?>/g, function (all, name, k) {
      var orig = state.tags[parseInt(k, 10)];
      return orig && tagName(orig) === name.toLowerCase() ? orig : "<" + name + ">";
    });
    // 3) segnaposto → originali
    out = out.replace(/<x-k\s+x="(\d+)"\s*>\s*<\/x-k>/gi, function (all, n) { return state.frozen[parseInt(n, 10)]; });
    return state.head + state.leading + out + state.trailing + state.tail;
  }

  var CLIP_NOTE = "\n\nThe HTML is a fragment copied from an email editor: tags may be unbalanced. Keep every tag exactly where it is relative to the text; never add closing tags or wrappers that are not in the input.";

  /** send: function(requestBody) -> Promise<string>|string  (sincrono in osascript) */
  function translateHtml(html, settings, send) {
    var st = prepare(html, { excludeSignature: settings.excludeSignature });
    if (!st.hasText) return { html: null, skipped: true, boundary: st.boundary };
    var req = C.buildRequest(settings, "html", st.payload, null);
    req.system += CLIP_NOTE;
    var out = send(req);
    var body = extract(out, "body_html");
    try {
      if (body === null) throw new StructureError("risposta senza <body_html>");
      return { html: restore(st, body), boundary: st.boundary };
    } catch (e) {
      if (!(e instanceof StructureError)) throw e;
      req.messages.push({ role: "assistant", content: out },
        { role: "user", content: e.message + ". Every x=\"n\" placeholder must appear exactly once. Reply again with the complete translation in the required format." });
      var out2 = send(req);
      var body2 = extract(out2, "body_html");
      if (body2 === null) throw new StructureError("risposta senza <body_html>");
      return { html: restore(st, body2), boundary: st.boundary };
    }
  }

  function translatePlain(text, settings, send) {
    if (!/[A-Za-zÀ-ÿ]/.test(text)) return { text: null, skipped: true };
    var parts = C.splitPlainText(text);
    var lead = (parts.head.match(/^\s*/) || [""])[0];
    var trail = (parts.head.match(/\s*$/) || [""])[0];
    var req = C.buildRequest(settings, "text", parts.head.trim(), null);
    var out = send(req);
    var t = extract(out, "body_text");
    if (t === null) throw new Error("Risposta non valida dal modello.");
    return { text: lead + t + trail + parts.tail, boundary: parts.tail ? "citazione" : null };
  }

  function extract(text, tag) {
    var a = text.indexOf("<" + tag + ">"), b = text.lastIndexOf("</" + tag + ">");
    if (a === -1 || b === -1 || b < a) return null;
    return text.slice(a + tag.length + 2, b).replace(/^\n/, "").replace(/\n$/, "");
  }

  root.TradClip = { prepare: prepare, restore: restore, translateHtml: translateHtml,
                    translatePlain: translatePlain, splitDocument: splitDocument, StructureError: StructureError };
})(typeof window !== "undefined" ? window : globalThis);
