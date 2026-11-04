# Traduttore Mail per Outlook

Add-in per Outlook (Mac, Windows, web) che traduce la bozza con Claude **direttamente dentro la mail**, mantenendo la formattazione.

**Uso:** scrivi in italiano → clic su **Traduci** nella barra del messaggio → dopo 5–15 secondi il testo è sostituito dalla versione inglese.

- Grassetti, colori, elenchi, link, tabelle e immagini restano dove li hai messi.
- Thread citato (risposte/inoltri) e firma non vengono toccati né inviati a Claude.
- Se selezioni una parte del testo, traduce solo quella.
- L'oggetto viene tradotto solo nelle mail nuove (nelle risposte resta invariato per non rompere il thread).
- Se Claude restituisce una struttura incompleta, la mail **non viene modificata** (un tentativo automatico di correzione, poi messaggio di errore).

---

## Cosa serve

| | |
|---|---|
| Chiave API Anthropic | console.anthropic.com → API Keys. È un account a consumo, separato dall'abbonamento Claude: serve caricare un credito (Billing). |
| Costo indicativo | Sonnet 5.5 costa 2 $ per milione di token in ingresso e 10 $ in uscita: una mail di mezza pagina costa circa 1–2 centesimi. |
| Account GitHub (gratuito) | Per pubblicare i file dell'add-in su un indirizzo HTTPS (GitHub Pages). |
| Permesso del tenant | La casella Office 365 deve consentire gli add-in personalizzati (la maggior parte dei tenant lo consente). |

---

## Installazione (circa 10 minuti, una volta sola)

### 1. Pubblica i file su GitHub Pages

1. Su github.com crea un repository **pubblico** chiamato `outlook-traduttore`.
2. *Add file → Upload files*: trascina **tutto il contenuto** di questa cartella (cartelle `src`, `assets`, file `index.html` ecc.) → *Commit changes*.
3. *Settings → Pages* → Source: *Deploy from a branch* → Branch `main`, cartella `/ (root)` → *Save*.
4. Dopo 1–2 minuti apri `https://TUO-UTENTE.github.io/outlook-traduttore/`: deve comparire la pagina "Traduttore Mail per Outlook".

Il codice è pubblico ma non contiene nulla di riservato: la chiave API non sta nei file.

### 2. Crea il manifest

Nel Terminale del Mac, dentro questa cartella:

```bash
./crea-manifest.sh https://TUO-UTENTE.github.io/outlook-traduttore
```

Viene creato `manifest.xml`. In alternativa apri `manifest.template.xml` con un editor di testo, sostituisci ogni `__BASE_URL__` con l'indirizzo sopra e salvalo come `manifest.xml`.

### 3. Installa l'add-in in Outlook

Il modo più affidabile è da Outlook sul web, poi si propaga da solo a Outlook per Mac e Windows della stessa casella:

1. Accedi a outlook.office.com con l'account di lavoro.
2. Apri **https://aka.ms/olksideload** → *Componenti aggiuntivi personali* → *Aggiungi un componente aggiuntivo personalizzato* → *Aggiungi da file* → scegli `manifest.xml` → *Installa*.

Su Outlook per Mac può servire qualche minuto (o un riavvio) perché compaia.

### 4. Prima configurazione

1. Apri una nuova mail. Nella barra del messaggio trovi **Traduci** e **Opzioni traduttore** (su Outlook per Mac sono sotto il pulsante *App* / *…* finché non li fissi nella barra).
2. *Opzioni traduttore* → *Impostazioni* → incolla la chiave API → **Verifica chiave**.

Consiglio: fissa il pulsante **Traduci** nella barra del messaggio (personalizza barra / clic destro → aggiungi), così basta un clic.

---

## Pannello Opzioni

| Opzione | Default | Note |
|---|---|---|
| Lingua di destinazione | Inglese (UK) | Anche USA, spagnolo, francese, tedesco. Il pulsante Traduci usa l'ultima lingua scelta. |
| Non tradurre la firma | sì | |
| Traduci l'oggetto | sì | Solo mail nuove. |
| Modello | Sonnet 5.5 | Haiku 5.5 è più veloce ed economico, un po' meno curato nello stile. |
| Registro | Business | Formale (enti, legali, direzione) o Diretto (colleghi). |
| Glossario | `PCS = GCV (Gross Calorific Value)` | Una voce per riga, Claude è obbligato a rispettarle. |
| Istruzioni aggiuntive | — | Testo libero passato a Claude a ogni traduzione. |
| Ripristina testo originale | | Riporta corpo e oggetto a com'erano prima dell'ultima traduzione (sullo stesso dispositivo). In alternativa Cmd+Z. |

Impostazioni e glossario si salvano nella casella di posta e valgono su tutti i dispositivi.

---

## Se qualcosa non va

| Sintomo | Causa / soluzione |
|---|---|
| In aka.ms/olksideload manca "Aggiungi un componente aggiuntivo personalizzato" | Il tenant blocca gli add-in personali. Chiedi all'IT di distribuire lo stesso `manifest.xml` da *Microsoft 365 admin center → App integrate* (solo per te). |
| "Chiave API non valida" | Ricopia la chiave, controlla che non sia stata revocata. |
| "Credito API esaurito" | console.anthropic.com → Billing. |
| "Struttura della mail alterata" | Nessuna modifica è stata fatta. Riprova, oppure seleziona il testo e traduci a pezzi. |
| La firma viene tradotta lo stesso | Il tuo client usa un marcatore diverso: seleziona solo il testo da tradurre. |
| "Mail troppo lunga" | Seleziona e traduci una parte alla volta. |
| Ho aggiornato i file ma Outlook usa la vecchia versione | Cache di Outlook: chiudi e riapri Outlook. Se cambi il manifest, rimuovi e reinstalla l'add-in. |

---

## Privacy

Il testo che scrivi (non il thread citato, non la firma se esclusa) viene inviato all'API di Anthropic solo quando premi Traduci, direttamente dal tuo Outlook senza server intermedi. Verifica che sia compatibile con le policy aziendali sull'uso di servizi esterni. La chiave API è salvata nelle impostazioni dell'add-in della tua casella Exchange.

---

## Come funziona (per chi vuole modificarlo)

```
src/core.js           motore: analisi HTML, esclusione firma/thread, chiamata API, ricostruzione
src/office-bridge.js  collegamento con Outlook (Office.js): legge/scrive corpo, oggetto, selezione
src/commands.*        pulsante Traduci (un clic, senza pannello)
src/taskpane.*        pannello Opzioni
manifest.template.xml manifest dell'add-in (validato con office-addin-manifest)
test/                 test automatici: npm i jsdom && node test/run-tests.js && node test/run-office-tests.js
```

Per mantenere la formattazione, `core.js` toglie da ogni tag gli attributi (stili, classi, link) e li sostituisce con un numero (`<b k="4">`). Claude riceve un HTML minimale, traduce frasi intere e può spostare un grassetto se l'ordine delle parole cambia; al ritorno ogni tag riprende i suoi attributi originali. Firma, thread citato e blocchi tecnici diventano segnaposto vuoti (`x="n"`) che vengono rimessi identici: se ne manca anche uno, la traduzione viene scartata.
