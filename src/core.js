/*
 * Traduttore Mail — motore di traduzione
 * --------------------------------------
 * Strategia per mantenere la formattazione:
 *  1. Il corpo HTML della mail viene analizzato come DOM.
 *  2. Tutto ciò che NON va tradotto (thread citato, firma, stili, commenti)
 *     viene "congelato": sostituito da un segnaposto vuoto con attributo x="n".
 *  3. Su tutti gli altri elementi gli attributi (stili, classi, href...) vengono
 *     tolti e messi da parte; al loro posto resta solo un indice k="n".
 *     Claude vede quindi un HTML minimale (<p k="0">Ciao <b k="1">Mario</b></p>)
 *     e può tradurre frasi intere, spostando i tag inline se la sintassi lo richiede.
 *  4. Al ritorno gli attributi originali vengono rimessi al loro posto e i
 *     segnaposto sostituiti dai nodi originali, identici al byte.
 *  5. Se un segnaposto manca o è duplicato la traduzione viene rifiutata:
 *     la mail non viene mai toccata con un risultato incompleto.
 *
 * Funziona sia nel browser (add-in Outlook) sia in Node (test, con jsdom).
 */
(function (root) {
  "use strict";

  var API_URL = "https://api.anthropic.com/v1/messages";

  var MODELS = {
    "claude-sonnet-5-5": {
      label: "Sonnet 5.5 — qualità (consigliato)",
      extra: { thinking: { type: "between_tools" }, output_config: { effort: "low" } }
    },
    "claude-haiku-5-5": {
      label: "Haiku 5.5 — più veloce ed economico",
      extra: { thinking: { type: "disabled" }, output_config: { effort: "low" } }
    }
  };

  var LANGUAGES = {
    "en-GB": "British English",
    "en-US": "American English",
    "es": "Spanish (Spain)",
    "fr": "French",
    "de": "German",
    "it": "Italian"
  };

  var LANGUAGE_LABELS = {
    "en-GB": "Inglese (UK)", "en-US": "Inglese (USA)", "es": "Spagnolo",
    "fr": "Francese", "de": "Tedesco", "it": "Italiano"
  };

  var STYLES = {
    business: "Professional business register, clear and natural, accessible rather than stiff. Sound like a competent native-speaking manager wrote it.",
    formale: "Formal register, suitable for authorities, legal counterparts and senior management. Polite and precise.",
    diretto: "Friendly and direct register for colleagues: concise, natural, no unnecessary formality."
  };

  var DEFAULT_SETTINGS = {
    apiKey: "",
    model: "claude-sonnet-5-5",
    target: "en-GB",
    style: "business",
    glossary: "PCS = GCV (Gross Calorific Value)",
    instructions: "",
    excludeSignature: true,
    translateSubject: true
  };

  // Nodi da congelare sempre (non contengono testo da tradurre o non vanno toccati)
  var ALWAYS_FROZEN = { STYLE: 1, SCRIPT: 1, XML: 1, TITLE: 1, META: 1, LINK: 1, TEMPLATE: 1, NOSCRIPT: 1 };

  // ---------------------------------------------------------------------------
  // Individuazione di thread citato e firma
  // ---------------------------------------------------------------------------

  function findQuoteStart(body) {
    var candidates = [];
    var selectors = [
      "#appendonsend",                              // Outlook web / nuovo Outlook
      "#divRplyFwdMsg", "#x_divRplyFwdMsg",          // intestazione risposta/inoltro
      "#mail-editor-reference-message-container",    // nuovo Outlook per Mac
      "#OLK_SRC_BODY_SECTION",                      // Outlook per Mac (vecchio)
      "#divtagdefaultwrapper + hr",
      ".gmail_quote", "blockquote[type='cite']",
      "div.moz-cite-prefix"
    ];
    selectors.forEach(function (sel) {
      try {
        var el = body.querySelector(sel);
        if (el) candidates.push(el);
      } catch (e) { /* selettore non supportato */ }
    });

    // Outlook classico (Windows): <div style="border:none;border-top:solid #E1E1E1 1.0pt;...">
    var divs = body.getElementsByTagName("div");
    for (var i = 0; i < divs.length; i++) {
      var st = (divs[i].getAttribute("style") || "").replace(/\s+/g, " ");
      if (/border-top:\s*solid\s+#(E1E1E1|B5C4DF)/i.test(st)) { candidates.push(divs[i]); break; }
    }

    // OWA: <hr style="display:inline-block;width:98%"> che precede l'intestazione
    var hrs = body.getElementsByTagName("hr");
    for (var h = 0; h < hrs.length; h++) {
      var hs = (hrs[h].getAttribute("style") || "").replace(/\s+/g, "");
      if (/display:inline-block;width:98%/i.test(hs)) { candidates.push(hrs[h]); break; }
    }

    // Ultima risorsa: un blocco che inizia con "Da:/From:" e contiene "Inviato:/Sent:"
    if (!candidates.length) {
      var blocks = body.querySelectorAll("p, div");
      for (var b = 0; b < blocks.length; b++) {
        var t = (blocks[b].textContent || "").replace(/\s+/g, " ").trim();
        if (/^(From|Da|De|Von):\s/.test(t) && /(Sent|Inviato|Data|Enviado|Gesendet|Date):\s/.test(t.slice(0, 400))) {
          candidates.push(blocks[b]);
          break;
        }
      }
    }
    return earliest(candidates);
  }

  function findSignatureStart(body) {
    var candidates = [];
    ["#Signature", "#signature", "[id^='Signature']", "div.x_Signature", "#OutlookSignature"].forEach(function (sel) {
      try { var el = body.querySelector(sel); if (el) candidates.push(el); } catch (e) { /* */ }
    });
    // Outlook classico: segnalibro _MailAutoSig
    var sigAnchor = body.querySelector("a[name='_MailAutoSig']");
    if (sigAnchor) candidates.push(closestBlock(sigAnchor, body));
    return earliest(candidates);
  }

  function closestBlock(el, stop) {
    var n = el;
    while (n && n !== stop) {
      if (/^(P|DIV|TABLE|LI|H[1-6])$/.test(n.nodeName)) return n;
      n = n.parentNode;
    }
    return el;
  }

  function earliest(nodes) {
    var best = null;
    nodes.forEach(function (n) {
      if (!n) return;
      // DOCUMENT_POSITION_PRECEDING = 2
      if (!best || (best.compareDocumentPosition(n) & 2)) best = n;
    });
    return best;
  }

  // ---------------------------------------------------------------------------
  // Preparazione: congelamento + compattazione
  // ---------------------------------------------------------------------------

  function freezeNode(state, node) {
    var doc = node.ownerDocument;
    var idx = state.frozen.length;
    state.frozen.push(node);
    var ph;
    if (node.nodeType === 1) {
      // Segnaposto con lo stesso nome di tag: il parser HTML non lo sposta
      // (es. un <tr> resta dentro la tabella)
      ph = doc.createElement(node.nodeName.toLowerCase());
    } else {
      ph = doc.createElement("x-k");
    }
    ph.setAttribute("x", String(idx));
    node.parentNode.replaceChild(ph, node);
    return ph;
  }

  // Congela il nodo "boundary" e tutto ciò che segue, fino a <body>
  function freezeFrom(state, boundary, rootEl) {
    var n = boundary;
    var includeSelf = true;
    while (n && n !== rootEl) {
      var parent = n.parentNode;
      var s = includeSelf ? n : n.nextSibling;
      while (s) {
        var next = s.nextSibling;
        if (!(s.nodeType === 3 && !/\S/.test(s.nodeValue))) freezeNode(state, s);
        s = next;
      }
      n = parent;
      includeSelf = false;
    }
  }

  function freezeSpecial(state, rootEl) {
    // Commenti (anche i condizionali di Outlook classico) e tag tecnici
    var walker = rootEl.ownerDocument.createTreeWalker(rootEl, 1 | 128, null, false);
    var toFreeze = [];
    var node;
    while ((node = walker.nextNode())) {
      if (node.nodeType === 8) toFreeze.push(node);
      else if (ALWAYS_FROZEN[node.nodeName.toUpperCase()] && !node.hasAttribute("x")) toFreeze.push(node);
    }
    toFreeze.forEach(function (n) { if (n.parentNode) freezeNode(state, n); });
  }

  function compact(state, rootEl) {
    var all = rootEl.getElementsByTagName("*");
    var list = [];
    for (var i = 0; i < all.length; i++) list.push(all[i]);
    list.forEach(function (el) {
      if (el.hasAttribute("x")) return;          // segnaposto
      var saved = [];
      for (var a = 0; a < el.attributes.length; a++) {
        saved.push([el.attributes[a].name, el.attributes[a].value]);
      }
      var k = state.attrs.length;
      state.attrs.push(saved);
      while (el.attributes.length) el.removeAttribute(el.attributes[0].name);
      el.setAttribute("k", String(k));
    });
  }

  function hasLetters(rootEl) {
    var walker = rootEl.ownerDocument.createTreeWalker(rootEl, 4, null, false);
    var n;
    while ((n = walker.nextNode())) {
      if (/[A-Za-zÀ-ÿ]/.test(n.nodeValue)) return true;
    }
    return false;
  }

  /**
   * html: corpo HTML (documento completo o frammento)
   * opts: { fragment: bool, excludeSignature: bool }
   */
  function prepare(html, opts) {
    opts = opts || {};
    var Parser = root.DOMParser;
    var isFull = !opts.fragment;
    var src = isFull ? html : "<!DOCTYPE html><html><head></head><body>" + html + "</body></html>";
    var doc = new Parser().parseFromString(src, "text/html");
    var body = doc.body;
    var state = { doc: doc, body: body, frozen: [], attrs: [], isFull: isFull,
                  hadDoctype: /^\s*<!doctype/i.test(html), boundary: null };

    if (isFull) {
      var quote = findQuoteStart(body);
      var sig = opts.excludeSignature ? findSignatureStart(body) : null;
      // se la "firma" trovata sta dentro il thread citato, ignorala
      if (sig && quote && (quote.compareDocumentPosition(sig) & 4)) sig = null;
      var boundary = earliest([quote, sig]);
      if (boundary) {
        state.boundary = boundary === sig ? "firma" : "citazione";
        freezeFrom(state, boundary, body);
      }
    }
    freezeSpecial(state, body);
    state.hasText = hasLetters(body);
    compact(state, body);
    state.payload = body.innerHTML.trim();
    return state;
  }

  /**
   * Rimette insieme il risultato tradotto. Lancia un errore se la struttura
   * non è integra (segnaposto mancanti o duplicati).
   */
  function restore(state, translatedHtml) {
    var doc = state.doc;
    var tmp = new root.DOMParser().parseFromString(
      "<!DOCTYPE html><html><head></head><body>" + translatedHtml + "</body></html>", "text/html").body;

    // 1) verifica segnaposto
    var phs = tmp.querySelectorAll("[x]");
    var seen = {};
    for (var i = 0; i < phs.length; i++) {
      var x = phs[i].getAttribute("x");
      if (seen[x]) throw new StructureError("segnaposto duplicato (" + x + ")");
      seen[x] = true;
    }
    for (var f = 0; f < state.frozen.length; f++) {
      if (!seen[String(f)]) throw new StructureError("parte non tradotta mancante (" + f + ")");
    }

    // 2) importa nel documento originale
    var body = state.body;
    while (body.firstChild) body.removeChild(body.firstChild);
    var nodes = Array.prototype.slice.call(tmp.childNodes);
    nodes.forEach(function (n) { body.appendChild(doc.importNode(n, true)); });

    // 3) ripristina attributi
    var withK = Array.prototype.slice.call(body.querySelectorAll("[k]"));
    withK.forEach(function (el) {
      var k = parseInt(el.getAttribute("k"), 10);
      el.removeAttribute("k");
      var saved = state.attrs[k];
      if (saved) saved.forEach(function (p) {
        try { el.setAttribute(p[0], p[1]); } catch (e) { /* nome attributo non valido */ }
      });
    });

    // 4) rimetti i nodi congelati originali
    var placeholders = Array.prototype.slice.call(body.querySelectorAll("[x]"));
    placeholders.forEach(function (ph) {
      var original = state.frozen[parseInt(ph.getAttribute("x"), 10)];
      if (original) ph.parentNode.replaceChild(original, ph);
    });

    if (!state.isFull) return body.innerHTML;
    return (state.hadDoctype ? "<!DOCTYPE html>\n" : "") + doc.documentElement.outerHTML;
  }

  function StructureError(msg) {
    this.name = "StructureError";
    this.message = "Struttura della mail alterata: " + msg;
  }
  StructureError.prototype = Object.create(Error.prototype);

  // ---------------------------------------------------------------------------
  // Testo semplice (mail in formato "Testo normale")
  // ---------------------------------------------------------------------------

  function splitPlainText(text) {
    var re = /^(-{3,}\s*(Original Message|Messaggio originale)\s*-{3,}|_{10,}|(From|Da):\s.+|On .+ wrote:|Il giorno .+ ha scritto:)\s*$/im;
    var m = re.exec(text);
    if (!m) return { head: text, tail: "" };
    return { head: text.slice(0, m.index), tail: text.slice(m.index) };
  }

  // ---------------------------------------------------------------------------
  // Prompt e chiamata API
  // ---------------------------------------------------------------------------

  function parseGlossary(g) {
    return (g || "").split(/\r?\n/).map(function (l) { return l.trim(); })
      .filter(function (l) { return l && !/^#/.test(l); });
  }

  function buildSystemPrompt(settings, mode) {
    var lang = LANGUAGES[settings.target] || settings.target;
    var style = STYLES[settings.style] || STYLES.business;
    var lines = [
      "You are the translation engine of an Outlook add-in. The author writes business emails (mostly in Italian) for the biogas/biomethane and energy sector and needs them in " + lang + ".",
      "",
      "Translate the email draft into " + lang + ".",
      "Style: " + style,
      "Be faithful to the author's meaning, intent and level of politeness. Do not add, remove or summarise content. Do not add greetings, sign-offs or comments. Text already in " + lang + " stays unchanged.",
      "Keep unchanged: personal and company names, email addresses, URLs, numbers, amounts, dates, units (Sm³/h, kWh, mbar...), codes, file names, and the names of Italian institutions, laws and programmes (GSE, ARERA, Snam, PNRR, D.Lgs., DM...). Use the target-language conventions for decimal separators only if the number is not a code.",
    ];
    var glossary = parseGlossary(settings.glossary);
    if (glossary.length) {
      lines.push("", "Mandatory glossary (source = target):");
      glossary.forEach(function (g) { lines.push("- " + g); });
    }
    if (settings.instructions && settings.instructions.trim()) {
      lines.push("", "Additional instructions from the author: " + settings.instructions.trim());
    }
    if (mode === "html") {
      lines.push(
        "",
        "FORMAT RULES (critical):",
        "- The draft is minimal HTML. Every element has a k=\"n\" attribute: keep each tag with its exact k value. Do not add attributes, styles or new wrapper elements.",
        "- Translate only the human-readable text between tags. Keep the same block structure (paragraphs, line breaks, lists, tables) in the same order.",
        "- Inline formatting tags (b, strong, i, em, u, span, a, font) must wrap the same words in the translation; you may move them within the sentence when word order requires it.",
        "- Elements with an x=\"n\" attribute are frozen placeholders (signature, quoted thread, images, technical blocks). Copy each one exactly once, unchanged and empty, in the same position.",
        "- Keep HTML entities valid (&amp; &lt; &gt;). &nbsp; may become a normal space.",
        "",
        "Reply with exactly:",
        "<body_html>…translated HTML…</body_html>",
        "and, only if a <subject> was provided, also <subject>…translated subject…</subject>. Keep reply prefixes such as R:/RE:/I:/FW: in their usual " + lang + " form. Output nothing else."
      );
    } else {
      lines.push(
        "",
        "The draft is plain text. Keep line breaks, blank lines, bullets and indentation exactly as they are.",
        "Reply with exactly <body_text>…translated text…</body_text> and, only if a <subject> was provided, <subject>…</subject>. Output nothing else."
      );
    }
    return lines.join("\n");
  }

  function extractTag(text, tag) {
    var open = "<" + tag + ">", close = "</" + tag + ">";
    var a = text.indexOf(open);
    var b = text.lastIndexOf(close);
    if (a === -1 || b === -1 || b < a) return null;
    return text.slice(a + open.length, b).trim();
  }

  function buildRequest(settings, mode, content, subject) {
    var model = MODELS[settings.model] ? settings.model : DEFAULT_SETTINGS.model;
    var tag = mode === "html" ? "body_html" : "body_text";
    var user = (subject ? "<subject>" + subject + "</subject>\n" : "") +
               "<" + tag + ">\n" + content + "\n</" + tag + ">";
    var maxTokens = Math.min(32000, Math.max(2048, Math.ceil(content.length / 1.5) + 1024));
    var body = {
      model: model,
      max_tokens: maxTokens,
      system: buildSystemPrompt(settings, mode),
      messages: [{ role: "user", content: user }]
    };
    var extra = MODELS[model].extra;
    Object.keys(extra).forEach(function (k) { body[k] = extra[k]; });
    return body;
  }

  function callClaude(settings, requestBody) {
    if (!settings.apiKey) {
      return Promise.reject(new Error("Chiave API mancante: aprila in Opzioni traduttore e incollala."));
    }
    return root.fetch(API_URL, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": settings.apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify(requestBody)
    }).then(function (res) {
      return res.text().then(function (txt) {
        var json = null;
        try { json = JSON.parse(txt); } catch (e) { /* */ }
        if (!res.ok) {
          var msg = (json && json.error && json.error.message) || txt.slice(0, 200);
          if (res.status === 401) msg = "chiave API non valida o revocata";
          else if (res.status === 429) msg = "limite di richieste raggiunto, riprova tra poco";
          else if (res.status === 529 || res.status === 503) msg = "servizio momentaneamente sovraccarico, riprova";
          else if (res.status === 400 && /credit/i.test(msg)) msg = "credito API esaurito (console.anthropic.com → Billing)";
          throw new Error("Errore API (" + res.status + "): " + msg);
        }
        var text = (json.content || []).filter(function (c) { return c.type === "text"; })
          .map(function (c) { return c.text; }).join("");
        if (json.stop_reason === "max_tokens") throw new Error("Mail troppo lunga: risposta troncata. Seleziona e traduci una parte alla volta.");
        return text;
      });
    });
  }

  /**
   * Traduce HTML preservando la formattazione.
   * translator (opzionale, per i test): function(requestBody) -> Promise<string>
   */
  function translateHtml(html, settings, opts) {
    opts = opts || {};
    var send = opts.translator || function (req) { return callClaude(settings, req); };
    var state = prepare(html, { fragment: !!opts.fragment, excludeSignature: settings.excludeSignature });
    if (!state.hasText) {
      return Promise.resolve({ html: null, subject: null, skipped: true, boundary: state.boundary });
    }
    var subject = opts.subject || null;
    var attempt = function (n, extraNote) {
      var req = buildRequest(settings, "html", state.payload, subject);
      if (extraNote) req.messages.push({ role: "assistant", content: extraNote.prev },
                                       { role: "user", content: extraNote.msg });
      return send(req).then(function (out) {
        var bodyHtml = extractTag(out, "body_html");
        if (bodyHtml === null) throw new StructureError("risposta senza <body_html>");
        var result;
        try {
          // restore() valida PRIMA di modificare il documento: se fallisce, state è intatto
          result = restore(state, bodyHtml);
        } catch (e) {
          if (n === 0 && e instanceof StructureError) {
            // un secondo tentativo con il dettaglio dell'errore
            return attempt(1, {
              prev: out,
              msg: "Your output broke the structure: " + e.message + ". Every x=\"n\" placeholder must appear exactly once. Reply again with the complete translation in the required format."
            });
          }
          throw e;
        }
        return { html: result, subject: subject ? (extractTag(out, "subject") || subject) : null, boundary: state.boundary };
      });
    };
    return attempt(0);
  }

  function translateText(text, settings, opts) {
    opts = opts || {};
    var send = opts.translator || function (req) { return callClaude(settings, req); };
    var parts = opts.fragment ? { head: text, tail: "" } : splitPlainText(text);
    if (!/[A-Za-zÀ-ÿ]/.test(parts.head)) return Promise.resolve({ text: null, skipped: true });
    var req = buildRequest(settings, "text", parts.head.replace(/\s+$/, ""), opts.subject || null);
    var trailing = (parts.head.match(/\s+$/) || [""])[0];
    return send(req).then(function (out) {
      var t = extractTag(out, "body_text");
      if (t === null) throw new Error("Risposta non valida dal modello.");
      return { text: t + trailing + parts.tail, subject: opts.subject ? (extractTag(out, "subject") || opts.subject) : null };
    });
  }

  root.TradCore = {
    MODELS: MODELS, LANGUAGES: LANGUAGES, LANGUAGE_LABELS: LANGUAGE_LABELS, STYLES: STYLES, DEFAULT_SETTINGS: DEFAULT_SETTINGS,
    prepare: prepare, restore: restore, translateHtml: translateHtml, translateText: translateText,
    buildRequest: buildRequest, buildSystemPrompt: buildSystemPrompt, callClaude: callClaude,
    splitPlainText: splitPlainText, StructureError: StructureError
  };
})(typeof window !== "undefined" ? window : globalThis);
