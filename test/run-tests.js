// Test del motore: node test/run-tests.js  (richiede: npm i jsdom)
const path = require("path");
const assert = require("assert");
const { JSDOM } = require(process.env.JSDOM_PATH || "jsdom");
const win = new JSDOM("").window;
globalThis.DOMParser = win.DOMParser;
require(path.join(__dirname, "../src/core.js"));
const C = globalThis.TradCore;
const F = require("./fixtures");

const settings = Object.assign({}, C.DEFAULT_SETTINGS, { apiKey: "test" });

// Finto modello: antepone [EN] a ogni testo con lettere, mantiene i tag
function mockTranslator(opts = {}) {
  let calls = 0;
  const fn = async (req) => {
    calls++;
    const u = req.messages[req.messages.length - 1].content;
    const src = req.messages[0].content;
    const subj = (src.match(/<subject>([\s\S]*?)<\/subject>/) || [])[1];
    let body = (src.match(/<body_html>\n?([\s\S]*?)\n?<\/body_html>/) || [])[1];
    if (body === undefined) {
      const t = src.match(/<body_text>\n?([\s\S]*?)\n?<\/body_text>/)[1];
      return `<body_text>${t.replace(/^(.*[A-Za-z].*)$/gm, "[EN]$1")}</body_text>` + (subj ? `<subject>[EN]${subj}</subject>` : "");
    }
    const d = new JSDOM(`<body>${body}</body>`).window.document;
    const w = d.createTreeWalker(d.body, 4);
    let n; while ((n = w.nextNode())) if (/[A-Za-zÀ-ÿ]/.test(n.nodeValue)) n.nodeValue = "[EN]" + n.nodeValue.trim();
    let out = d.body.innerHTML;
    if (opts.dropPlaceholder && (opts.dropPlaceholder === "always" || calls === 1)) {
      out = out.replace(/<([a-z0-9:-]+) x="0"[^>]*>(<\/\1>)?/, "");
    }
    return `<body_html>${out}</body_html>` + (subj ? `<subject>[EN]${subj}</subject>` : "");
  };
  fn.calls = () => calls;
  return fn;
}

const tests = [];
const test = (name, f) => tests.push([name, f]);

test("Outlook web: risposta con firma e thread", async () => {
  const st = C.prepare(F.owaReply, { excludeSignature: true });
  assert(!/style=|class=|href=/.test(st.payload), "il payload non deve contenere attributi");
  assert(!/Any news|pablo@example/i.test(st.payload), "il thread citato non va inviato");
  assert(!/Operations Manager/.test(st.payload), "la firma non va inviata");
  assert.strictEqual(st.boundary, "firma");
  const r = await C.translateHtml(F.owaReply, settings, { translator: mockTranslator() });
  const h = r.html;
  assert(h.includes("[EN]Ciao Pablo,"), "testo tradotto");
  assert(/<b>\[EN\]test di prestazione<\/b>/.test(h), "grassetto mantenuto");
  assert(/<span style="color: rgb\(192, 0, 0\);">\[EN\]250 Sm³\/h<\/span>/.test(h), "colore mantenuto");
  assert(h.includes('href="https://www.gse.it/doc.pdf"'), "link mantenuto");
  assert(h.includes('font-family: Aptos'), "font mantenuto");
  assert(h.includes("<p>Riccardo<br>Operations Manager</p>"), "firma intatta");
  assert(h.includes("Hi Riccardo, any news on the test?") && !h.includes("[EN]Hi"), "thread intatto");
  assert(h.includes("P {margin-top:0;margin-bottom:0;}"), "stile in head intatto");
  assert(h.includes('<hr style="display:inline-block;width:98%" tabindex="-1">'));
  console.log("   payload inviato al modello:", st.payload.length, "caratteri vs", F.owaReply.length, "originali");
});

test("Outlook web: firma inclusa se l'opzione è disattivata", async () => {
  const st = C.prepare(F.owaReply, { excludeSignature: false });
  assert.strictEqual(st.boundary, "citazione");
  assert(/Operations Manager/.test(st.payload));
});

test("Outlook classico Windows: _MailAutoSig, o:p, commenti condizionali", async () => {
  const st = C.prepare(F.classicReply, { excludeSignature: true });
  assert(!/Green Farm|Prodeval|Dear Riccardo/.test(st.payload));
  const r = await C.translateHtml(F.classicReply, settings, { translator: mockTranslator() });
  const h = r.html;
  assert(/<b>\[EN\]verificare la discrepanza<\/b>/.test(h));
  assert(h.includes('class="MsoNormal"'));
  assert(h.includes("<o:p>"), "tag o:p mantenuti");
  assert(h.includes("[if gte mso 9]"), "commenti condizionali mantenuti");
  assert(h.includes("Dear Riccardo, please find attached.") && !h.includes("[EN]Dear"));
  assert(h.includes("border-top:solid #E1E1E1"));
  assert(!h.includes("[EN]Green Farm"), "firma non tradotta");
});

test("Nuova mail: tabella e immagine inline (cid)", async () => {
  const r = await C.translateHtml(F.newMessage, settings, { translator: mockTranslator(), subject: "Aggiornamento costi" });
  assert(r.html.includes('src="cid:image001.png@01DA"'));
  assert(/<td style="padding:4px"><b>\[EN\]Voce<\/b><\/td>/.test(r.html));
  assert(r.html.includes("€ 120.000"));
  assert.strictEqual(r.subject, "[EN]Aggiornamento costi");
});

test("Firma dentro tabella: segnaposto non spostati dal parser", async () => {
  const r = await C.translateHtml(F.signatureInTable, settings, { translator: mockTranslator() });
  assert(r.html.includes("[EN]Grazie mille."));
  assert(/<table><tbody><tr><td><div id="Signature">Riccardo<\/div><\/td><\/tr><tr><td>Tel\. 333<\/td><\/tr><\/tbody><\/table>/.test(r.html), r.html);
});

test("Solo testo selezionato (frammento)", async () => {
  const r = await C.translateHtml(F.selection, settings, { translator: mockTranslator(), fragment: true });
  assert.strictEqual(r.html, '<span style="font-size:12pt">[EN]Il<b>[EN]verbale</b>[EN]è pronto.</span>');
});

test("Struttura rotta al 1° tentativo → secondo tentativo riuscito", async () => {
  const m = mockTranslator({ dropPlaceholder: "first" });
  const r = await C.translateHtml(F.owaReply, settings, { translator: m });
  assert.strictEqual(m.calls(), 2);
  assert(r.html.includes("Hi Riccardo, any news"));
});

test("Struttura sempre rotta → errore, nessuna modifica", async () => {
  await assert.rejects(
    C.translateHtml(F.owaReply, settings, { translator: mockTranslator({ dropPlaceholder: "always" }) }),
    (e) => e instanceof C.StructureError);
});

test("Mail in testo semplice: il thread non viene tradotto", async () => {
  const r = await C.translateText(F.plainReply, settings, { translator: mockTranslator() });
  assert(r.text.startsWith("[EN]Ciao Jorge,"));
  assert(r.text.includes("\n\n[EN]Riccardo\n\n-----Original Message-----\nFrom: Jorge\nSent: Monday\nCan you send"));
});

test("Parametri API per modello", () => {
  const s = C.buildRequest(settings, "html", "<p k=\"0\">ciao</p>");
  assert.deepStrictEqual(s.thinking, { type: "between_tools" });
  assert.deepStrictEqual(s.output_config, { effort: "low" });
  assert(!("temperature" in s));
  const h = C.buildRequest(Object.assign({}, settings, { model: "claude-haiku-5-5" }), "html", "x");
  assert.deepStrictEqual(h.thinking, { type: "disabled" });
  assert(s.system.includes("PCS = GCV"), "glossario nel prompt");
  assert(s.system.includes("British English"));
});

test("Mail senza testo da tradurre", async () => {
  const r = await C.translateHtml('<html><body><div id="Signature">Riccardo</div></body></html>', settings, { translator: mockTranslator() });
  assert(r.skipped);
});

(async () => {
  let fail = 0;
  for (const [name, f] of tests) {
    try { await f(); console.log("✓", name); }
    catch (e) { fail++; console.log("✗", name, "\n  ", e.message); }
  }
  console.log(fail ? `\n${fail} test falliti` : `\nTutti i ${tests.length} test superati`);
  process.exit(fail ? 1 : 0);
})();
