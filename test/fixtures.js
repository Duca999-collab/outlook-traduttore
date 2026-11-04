// Bozze nel formato HTML reale generato dai diversi client Outlook
module.exports = {
  // Outlook web / nuovo Outlook: risposta con firma e thread citato
  owaReply: `<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style type="text/css" style="display:none;"> P {margin-top:0;margin-bottom:0;} </style></head><body dir="ltr">
<div class="elementToProof" style="font-family: Aptos, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Ciao Pablo,</div>
<div class="elementToProof" style="font-family: Aptos, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">ti confermo che il <b>test di prestazione</b> è stato superato: <span style="color: rgb(192, 0, 0);">250 Sm³/h</span> per sei ore.</div>
<ul><li>PCS conforme</li><li>Collaudo previsto entro <u>venerdì</u></li></ul>
<div class="elementToProof"><a href="https://www.gse.it/doc.pdf" title="doc">documento GSE</a></div>
<div id="Signature"><p>Riccardo<br>Operations Manager</p></div>
<div id="appendonsend"></div>
<hr style="display:inline-block;width:98%" tabindex="-1">
<div id="divRplyFwdMsg" dir="ltr"><font face="Calibri, sans-serif" style="font-size:11pt" color="#000000"><b>From:</b> Pablo &lt;pablo@example.com&gt;<br><b>Sent:</b> Tuesday<br><b>Subject:</b> Performance test</font><div>&nbsp;</div></div>
<div><p>Hi Riccardo, any news on the test?</p></div>
</body></html>`,

  // Outlook classico Windows (Word): risposta con firma _MailAutoSig
  classicReply: `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if gte mso 9]><xml><o:shapedefaults v:ext="edit" spidmax="1026" /></xml><![endif]--><style><!-- p.MsoNormal {margin:0cm; font-size:11.0pt;} --></style></head><body lang=IT link="#0563C1" vlink="#954F72" style='word-wrap:break-word'><div class=WordSection1>
<p class=MsoNormal>Buongiorno,<o:p></o:p></p>
<p class=MsoNormal>vi chiedo di <b>verificare la discrepanza</b> tra l’analizzatore e il gascromatografo.<o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal><a name="_MailAutoSig"><span style='color:#1F3864'>Riccardo</span></a><o:p></o:p></p>
<p class=MsoNormal><span style='color:#1F3864'>Green Farm</span><o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<div><div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm'><p class=MsoNormal><b>Da:</b> Prodeval<br><b>Inviato:</b> lunedì<br><b>Oggetto:</b> H2S</p></div></div>
<p class=MsoNormal>Dear Riccardo, please find attached.<o:p></o:p></p>
</div></body></html>`,

  // Nuova mail con tabella e immagine inline
  newMessage: `<html><head></head><body><div style="font-family:Aptos">Gentili tutti,</div>
<table border="1" style="border-collapse:collapse"><tbody><tr><td style="padding:4px"><b>Voce</b></td><td>Importo</td></tr><tr><td>Opere civili</td><td>€ 120.000</td></tr></tbody></table>
<p>In allegato lo schema <img src="cid:image001.png@01DA" width="200"> aggiornato.</p></body></html>`,

  // Firma dentro una tabella (caso limite per i segnaposto)
  signatureInTable: `<html><head></head><body><p>Grazie mille.</p><table><tbody><tr><td><div id="Signature">Riccardo</div></td></tr><tr><td>Tel. 333</td></tr></tbody></table></body></html>`,

  selection: `<span style="font-size:12pt">Il <b>verbale</b> è pronto.</span>`,

  plainReply: `Ciao Jorge,
ti mando i dati aggiornati.

- Produzione: 250 Sm³/h
- Disponibilità: 98%

Riccardo

-----Original Message-----
From: Jorge
Sent: Monday
Can you send the update?`
};
