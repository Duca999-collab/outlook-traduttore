// Test del ponte Outlook con un finto Office.js e un finto endpoint API
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require(process.env.JSDOM_PATH || "jsdom");
const F = require("./fixtures");

function makeEnv({ bodyType = "html", body, subject = "", selection = "", selectionHtml = "", conversationId = null, apiStatus = 200 }) {
  const dom = new JSDOM("<!doctype html><body></body>", { runScripts: "outside-only", url: "https://example.github.io/" });
  const w = dom.window;
  const ok = (v) => ({ status: "succeeded", value: v });
  const state = { body, subject, notifications: [], requests: [], setSelected: null };
  const roaming = {};
  w.Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { ItemNotificationMessageType: { ProgressIndicator: "progressIndicator", InformationalMessage: "informationalMessage", ErrorMessage: "errorMessage" } },
    context: {
      roamingSettings: { get: (k) => roaming[k], set: (k, v) => { roaming[k] = v; }, saveAsync: (cb) => cb(ok()) },
      mailbox: { item: {
        conversationId,
        body: {
          getTypeAsync: (cb) => cb(ok(bodyType)),
          getAsync: (t, cb) => cb(ok(state.body)),
          setAsync: (v, o, cb) => { state.body = v; cb(ok()); }
        },
        subject: { getAsync: (cb) => cb(ok(state.subject)), setAsync: (v, cb) => { state.subject = v; cb(ok()); } },
        getSelectedDataAsync: (t, cb) => cb(ok({ data: t === "html" ? selectionHtml : selection, sourceProperty: "body" })),
        setSelectedDataAsync: (v, o, cb) => { state.setSelected = v; cb(ok()); },
        notificationMessages: { replaceAsync: (k, m, cb) => { state.notifications.push(m); cb && cb(ok()); } }
      } }
    }
  };
  // finta API Anthropic: traduce anteponendo [EN] ai nodi di testo
  w.fetch = async (url, init) => {
    const req = JSON.parse(init.body);
    state.requests.push({ url, headers: init.headers, req });
    if (apiStatus !== 200) return { ok: false, status: apiStatus, text: async () => JSON.stringify({ error: { message: "x" } }) };
    const u = req.messages[0].content;
    const subj = (u.match(/<subject>([\s\S]*?)<\/subject>/) || [])[1];
    const b = u.match(/<body_html>\n?([\s\S]*?)\n?<\/body_html>/);
    let out;
    if (b) {
      const d = new JSDOM(`<body>${b[1]}</body>`).window.document;
      const tw = d.createTreeWalker(d.body, 4); let n;
      while ((n = tw.nextNode())) if (/[A-Za-zÀ-ÿ]/.test(n.nodeValue)) n.nodeValue = "[EN]" + n.nodeValue;
      out = `<body_html>${d.body.innerHTML}</body_html>`;
    } else {
      out = `<body_text>[EN]${u.match(/<body_text>\n?([\s\S]*?)\n?<\/body_text>/)[1]}</body_text>`;
    }
    if (subj) out += `<subject>[EN]${subj}</subject>`;
    return { ok: true, status: 200, text: async () => JSON.stringify({ content: [{ type: "text", text: out }], stop_reason: "end_turn" }) };
  };
  for (const f of ["core.js", "office-bridge.js"]) w.eval(fs.readFileSync(path.join(__dirname, "../src", f), "utf8"));
  w.Office.context.roamingSettings.set("tradSettings", JSON.stringify({ apiKey: "sk-ant-test" }));
  return { w, state };
}

const tests = [];
const test = (n, f) => tests.push([n, f]);

test("Mail intera (risposta): corpo tradotto, oggetto NON toccato, header API corretti", async () => {
  const { w, state } = makeEnv({ body: F.owaReply, subject: "RE: Performance test", conversationId: "AAQk" });
  const msg = await w.TradOutlook.translateCurrentItem();
  assert(state.body.includes("[EN]Ciao Pablo,"));
  assert(state.body.includes("Hi Riccardo, any news"));
  assert.strictEqual(state.subject, "RE: Performance test");
  const r = state.requests[0];
  assert.strictEqual(r.url, "https://api.anthropic.com/v1/messages");
  assert.strictEqual(r.headers["anthropic-dangerous-direct-browser-access"], "true");
  assert.strictEqual(r.headers["x-api-key"], "sk-ant-test");
  assert.strictEqual(r.req.model, "claude-sonnet-5-5");
  assert(/firma e thread esclusi/.test(msg), msg);
  assert.strictEqual(state.notifications.at(-1).type, "informationalMessage");
  console.log("   notifica:", state.notifications.at(-1).message);
});

test("Mail nuova: traduce anche l'oggetto", async () => {
  const { w, state } = makeEnv({ body: F.newMessage, subject: "Aggiornamento costi" });
  await w.TradOutlook.translateCurrentItem();
  assert.strictEqual(state.subject, "[EN]Aggiornamento costi");
});

test("Testo selezionato: sostituisce solo la selezione", async () => {
  const { w, state } = makeEnv({ body: F.owaReply, selection: "Il verbale è pronto.", selectionHtml: F.selection });
  const msg = await w.TradOutlook.translateCurrentItem();
  assert.strictEqual(state.setSelected, '<span style="font-size:12pt">[EN]Il <b>[EN]verbale</b>[EN] è pronto.</span>');
  assert.strictEqual(state.body, F.owaReply, "il resto del corpo non deve cambiare");
  assert(/selezionato/.test(msg));
});

test("Ripristino del testo originale", async () => {
  const { w, state } = makeEnv({ body: F.classicReply, subject: "Oggetto" });
  await w.TradOutlook.translateCurrentItem();
  assert.notStrictEqual(state.body, F.classicReply);
  await w.TradOutlook.restoreOriginal();
  assert.strictEqual(state.body, F.classicReply);
  assert.strictEqual(state.subject, "Oggetto");
});

test("Mail in testo semplice", async () => {
  const { w, state } = makeEnv({ bodyType: "text", body: F.plainReply });
  await w.TradOutlook.translateCurrentItem();
  assert(state.body.startsWith("[EN]Ciao Jorge,"));
  assert(state.body.includes("-----Original Message-----\nFrom: Jorge"));
});

test("Chiave errata: errore chiaro nella barra della mail, corpo intatto", async () => {
  const { w, state } = makeEnv({ body: F.owaReply, apiStatus: 401 });
  await assert.rejects(w.TradOutlook.translateCurrentItem(), /chiave API non valida/);
  assert.strictEqual(state.body, F.owaReply);
  assert.strictEqual(state.notifications.at(-1).type, "errorMessage");
});

test("Lingua scelta nel pannello (spagnolo)", async () => {
  const { w, state } = makeEnv({ body: F.newMessage });
  await w.TradOutlook.translateCurrentItem({ target: "es" });
  assert(state.requests[0].req.system.includes("Spanish"));
});

(async () => {
  let fail = 0;
  for (const [n, f] of tests) {
    try { await f(); console.log("✓", n); } catch (e) { fail++; console.log("✗", n, "\n  ", e.message); }
  }
  console.log(fail ? `\n${fail} test falliti` : `\nTutti i ${tests.length} test superati`);
  process.exit(fail ? 1 : 0);
})();
