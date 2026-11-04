// node test/test-clip.js  — test del motore appunti (senza DOM)
const path = require("path");
const assert = require("assert");
require(process.env.CORE_JS || path.join(__dirname, "../../src/core.js"));
require(path.join(__dirname, "../clip-engine.js"));
const T = globalThis.TradClip, C = globalThis.TradCore;
const settings = Object.assign({}, C.DEFAULT_SETTINGS);

// Finto modello: antepone [EN] a ogni testo con lettere fuori dai tag
function mock(opts = {}) {
  let n = 0;
  const f = (req) => {
    n++;
    const u = req.messages[0].content;
    const m = u.match(/<body_html>\n([\s\S]*)\n<\/body_html>/);
    if (!m) return "<body_text>[EN]" + u.match(/<body_text>\n([\s\S]*)\n<\/body_text>/)[1] + "</body_text>";
    let out = m[1].replace(/>([^<]*[A-Za-zÀ-ÿ][^<]*)</g, (a, t) => ">[EN]" + t + "<").replace(/^([^<]*[A-Za-z][^<]*)</, "[EN]$1<");
    if (opts.drop && (opts.drop === "always" || n === 1)) out = out.replace(/<x-k x="0"><\/x-k>/, "");
    return "<body_html>" + out + "</body_html>";
  };
  f.calls = () => n;
  return f;
}

const tests = [];
const test = (n, f) => tests.push([n, f]);

// Copia da Gmail in Chrome (formato reale degli appunti public.html)
const gmail = `<meta charset='utf-8'><div dir="ltr" style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small;">Ciao Jorge,<div>ti mando il <b>report</b> di <span style="color:rgb(204,0,0)">ottobre</span>.</div><div><br></div><div>Saluti</div><span class="gmail_signature_prefix">-- </span><br><div dir="ltr" class="gmail_signature">Riccardo Stocca</div><br><div class="gmail_quote gmail_quote_container"><div dir="ltr" class="gmail_attr">Il giorno lun 3 nov 2026 alle ore 10:00 Jorge &lt;j@x.com&gt; ha scritto:<br></div><blockquote class="gmail_quote">Can you send the report?</blockquote></div></div>`;

test("Gmail: testo tradotto, firma e citazione intatte, stili ripristinati", () => {
  const st = T.prepare(gmail, { excludeSignature: true });
  assert(!/style=|class=/.test(st.payload), st.payload);
  assert(!/Riccardo Stocca|Can you send/.test(st.payload));
  assert.strictEqual(st.boundary, "firma");
  const r = T.translateHtml(gmail, settings, mock());
  assert(r.html.startsWith("<meta charset='utf-8'><div dir=\"ltr\" style=\"color: rgb(34, 34, 34);"), r.html.slice(0, 80));
  assert(r.html.includes("[EN]Ciao Jorge,"));
  assert(r.html.includes('<span style="color:rgb(204,0,0)">[EN]ottobre</span>'));
  assert(r.html.includes('<div dir="ltr" class="gmail_signature">Riccardo Stocca</div>'));
  assert(r.html.includes('<blockquote class="gmail_quote">Can you send the report?</blockquote>'));
  assert(r.html.endsWith("</div></div>"));
});

test("Gmail: firma inclusa se l'opzione è disattivata", () => {
  const st = T.prepare(gmail, { excludeSignature: false });
  assert.strictEqual(st.boundary, "citazione");
  assert(/Riccardo Stocca/.test(st.payload));
});

// Copia dalla finestra di composizione di Outlook (documento completo con commenti StartFragment)
const outlook = `<html><head><meta http-equiv="content-type" content="text/html; charset=utf-8"><style>p{margin:0}</style></head><body><!--StartFragment--><div class="elementToProof" style="font-family:Aptos;font-size:12pt">Buongiorno,</div><div class="elementToProof">il <u>collaudo</u> è fissato per <b>giovedì</b>.</div><div id="appendonsend"></div><hr style="display:inline-block;width:98%"><div id="divRplyFwdMsg"><b>From:</b> Prodeval</div><div>Dear all</div><!--EndFragment--></body></html>`;

test("Outlook: testa, commenti e thread conservati", () => {
  const r = T.translateHtml(outlook, settings, mock());
  assert(r.html.startsWith('<html><head><meta http-equiv="content-type" content="text/html; charset=utf-8"><style>p{margin:0}</style></head><body><!--StartFragment--><div class="elementToProof" style="font-family:Aptos;font-size:12pt">[EN]Buongiorno,</div>'), r.html);
  assert(r.html.includes("<u>[EN]collaudo</u>"));
  assert(r.html.includes('<div id="appendonsend"></div><hr style="display:inline-block;width:98%"><div id="divRplyFwdMsg"><b>From:</b> Prodeval</div><div>Dear all</div><!--EndFragment--></body></html>'));
  assert.strictEqual(r.boundary, "citazione");
});

// HTML prodotto da textutil a partire da RTF (Mail, Notes, Outlook)
const fromRtf = `<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title></title>
<style type="text/css">
p.p1 {margin: 0.0px 0.0px 0.0px 0.0px; font: 12.0px Helvetica}
span.s1 {font: 12.0px 'Helvetica-Bold'}
</style>
</head>
<body>
<p class="p1">Vi chiedo la <span class="s1">conferma</span> entro venerdì.</p>
<ul class="ul1">
<li class="li1">punto uno</li>
</ul>
</body>
</html>
`;

test("HTML da RTF (textutil)", () => {
  const r = T.translateHtml(fromRtf, settings, mock());
  assert(r.html.includes('<p class="p1">[EN]Vi chiedo la <span class="s1">[EN]conferma</span>[EN] entro venerdì.</p>'), r.html);
  assert(r.html.includes('<li class="li1">[EN]punto uno</li>'));
  assert(r.html.includes("span.s1 {font: 12.0px 'Helvetica-Bold'}"));
});

test("Frammento con tag non chiusi e '<' nel testo", () => {
  const frag = `<span style="x">a < b, quindi <i>ok</i>`;
  const r = T.translateHtml(frag, settings, mock());
  assert(r.html.includes('<span style="x">') && r.html.includes("<i>"));
});

test("Segnaposto perso al 1° tentativo → riprova", () => {
  const m = mock({ drop: "first" });
  const r = T.translateHtml(outlook, settings, m);
  assert.strictEqual(m.calls(), 2);
  assert(r.html.includes("<!--StartFragment-->"));
});

test("Segnaposto sempre perso → errore", () => {
  assert.throws(() => T.translateHtml(outlook, settings, mock({ drop: "always" })), T.StructureError);
});

test("Testo semplice con thread", () => {
  const r = T.translatePlain("Ciao,\nfatto.\n\nOn Mon, Nov 3, 2026 at 10:00 Jorge wrote:\n> ok?", settings, mock());
  assert.strictEqual(r.text, "[EN]Ciao,\nfatto.\n\nOn Mon, Nov 3, 2026 at 10:00 Jorge wrote:\n> ok?");
});

test("Prompt: nota frammento + glossario", () => {
  let req;
  T.translateHtml("<p>ciao</p>", settings, (r) => { req = r; return "<body_html><p k=\"0\">hi</p></body_html>"; });
  assert(/unbalanced/.test(req.system) && /PCS = GCV/.test(req.system));
  assert.strictEqual(req.model, "claude-sonnet-5-5");
});

let fail = 0;
for (const [n, f] of tests) { try { f(); console.log("✓", n); } catch (e) { fail++; console.log("✗", n, "\n  ", e.message); } }
console.log(fail ? `\n${fail} falliti` : `\nTutti i ${tests.length} test superati`);
process.exit(fail ? 1 : 0);
