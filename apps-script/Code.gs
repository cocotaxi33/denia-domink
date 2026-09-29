/**
 * RSVP-Backend für denia-und-dominik.de
 *
 * Einrichtung (einmalig):
 * 1. Google Sheet öffnen → Erweiterungen → Apps Script → diesen Code einfügen.
 * 2. CONFIG unten ausfüllen (E-Mail-Adressen von Denia & Dominik).
 * 3. Oben die Funktion „setup“ auswählen → Ausführen → Berechtigungen erlauben.
 *    Legt die Tabs Antworten, Personen, Gästeliste und Übersicht an.
 * 4. Bereitstellen → Neue Bereitstellung → Typ „Web-App“
 *    Ausführen als: Ich · Zugriff: Jeder → Bereitstellen → Web-App-URL kopieren.
 * 5. Die URL in rsvp.html bei API_URL eintragen.
 * Nach Code-Änderungen: Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Neue Version.
 */

const CONFIG = {
  NOTIFY: ['DENIA@BEISPIEL.DE', 'DOMINIK@BEISPIEL.DE'], // Benachrichtigung bei jeder Antwort
  REPLY_TO: 'DENIA@BEISPIEL.DE',                       // Antworten der Gäste auf die Bestätigungsmail gehen hierhin
  SENDER_NAME: 'Denia & Dominik',
  SITE: 'https://denia-und-dominik.de'
};

const SH = { ANTWORTEN: 'Antworten', PERSONEN: 'Personen', GAESTE: 'Gästeliste', UEBERSICHT: 'Übersicht' };
const H_ANTWORTEN = ['Antwort-ID', 'Erstellt', 'Geändert', 'Status', 'Personen', 'Erwachsene', 'Kinder', 'Namen', 'Kuchen', 'Welcher Kuchen', 'E-Mail', 'Telefon', 'Nachricht', 'Daten (JSON)'];
const H_PERSONEN = ['Antwort-ID', 'Zeitpunkt', 'Quelle', 'Status', 'Vorname', 'Nachname', 'Typ', 'Alter', 'Kuchen', 'Kontakt', 'Hinweis'];
const H_GAESTE = ['Einladung (Haushalt)', 'Eingeladene Personen', 'Einlegeblatt (Abend)', 'Nachname für Abgleich', 'Status (automatisch)', 'Zugesagte Personen', 'Status manuell', 'Notiz'];
const DUP = 'Mögliche Dublette';

/* ================= Einrichtung ================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const red = '#C8231A';

  const a = sheet_(ss, SH.ANTWORTEN, H_ANTWORTEN);
  a.getRange('A:N').setNumberFormat('@');
  a.hideColumns(14);

  const p = sheet_(ss, SH.PERSONEN, H_PERSONEN);
  p.getRange('A:K').setNumberFormat('@');
  const list = (vals) => SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build();
  p.getRange('C2:C').setDataValidation(list(['Formular', 'Telefon']));
  p.getRange('D2:D').setDataValidation(list(['Zusage', 'Absage']));
  p.getRange('G2:G').setDataValidation(list(['Erwachsen', 'Kind']));
  p.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$K2="' + DUP + '"')
      .setBackground('#FEF7E0').setRanges([p.getRange('A2:K')]).build()
  ]);

  const g = sheet_(ss, SH.GAESTE, H_GAESTE);
  g.getRange('C2:C').setDataValidation(list(['ja', 'nein']));
  g.getRange('G2:G').setDataValidation(list(['Zugesagt', 'Abgesagt', 'Offen']));
  g.getRange('E2').setFormula('=ARRAYFORMULA(IF(D2:D="","",IF(COUNTIFS(Personen!F2:F,D2:D,Personen!D2:D,"Zusage")>0,"Zugesagt",IF(COUNTIFS(Personen!F2:F,D2:D,Personen!D2:D,"Absage")>0,"Abgesagt","Offen"))))');
  g.getRange('F2').setFormula('=ARRAYFORMULA(IF(D2:D="","",COUNTIFS(Personen!F2:F,D2:D,Personen!D2:D,"Zusage")))');
  g.getRange('E1:F1').setBackground('#F1F3F4');
  g.getRange('A1').setNote('Pro eingeladenem Haushalt eine Zeile. „Nachname für Abgleich“ wird mit den Antworten verglichen. Wenn der automatische Status nicht passt, „Status manuell“ setzen.');
  g.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2<>"",$G2="",$E2="Offen")')
      .setBackground('#FCE8E6').setRanges([g.getRange('A2:H')]).build()
  ]);

  const u = sheet_(ss, SH.UEBERSICHT, ['Kennzahl', 'Wert']);
  const rows = [
    ['Zugesagte Personen', '=COUNTIF(Personen!D2:D,"Zusage")'],
    ['davon Erwachsene', '=COUNTIFS(Personen!D2:D,"Zusage",Personen!G2:G,"Erwachsen")'],
    ['davon Kinder', '=COUNTIFS(Personen!D2:D,"Zusage",Personen!G2:G,"Kind")'],
    ['Zusagen (Antworten)', '=COUNTIF(Antworten!D2:D,"Zusage")'],
    ['Absagen (Antworten)', '=COUNTIF(Antworten!D2:D,"Absage")'],
    ['Kuchen zugesagt', '=COUNTA(Personen!I2:I)'],
    ['Kuchenliste', '=TEXTJOIN(", ",TRUE,Personen!I2:I)'],
    ['Einladungen ohne Antwort', '=COUNTIFS(\'Gästeliste\'!D2:D,"<>",\'Gästeliste\'!G2:G,"",\'Gästeliste\'!E2:E,"Offen")+COUNTIF(\'Gästeliste\'!G2:G,"Offen")'],
    ['Mögliche Dubletten (Zeilen)', '=COUNTIF(Personen!K2:K,"' + DUP + '")']
  ];
  u.getRange(2, 1, rows.length, 2).setValues(rows);
  u.setColumnWidth(1, 240); u.setColumnWidth(2, 420);

  const ours = Object.keys(SH).map(k => SH[k]);
  ss.getSheets().forEach(s => { if (ours.indexOf(s.getName()) < 0 && s.getLastRow() <= 1) ss.deleteSheet(s); });
  ss.setActiveSheet(u);
  [a, p, g].forEach(s => s.getRange(1, 1, 1, s.getLastColumn()).setBackground('#F1F3F4'));
  u.getRange('A1:B1').setBackground('#F1F3F4');
  g.getRange('E1:F1').setFontColor(red);
}

function sheet_(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  s.setFrozenRows(1);
  return s;
}

/* ================= Web-App ================= */

function doGet(e) {
  const id = e && e.parameter && e.parameter.id;
  if (!id) return json_({ ok: true });
  const row = findRow_(SpreadsheetApp.getActive().getSheetByName(SH.ANTWORTEN), id);
  if (!row) return json_({ ok: false, error: 'not-found' });
  const data = JSON.parse(row.values[13]);
  return json_({ ok: true, data: data });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad-request' }); }
  if (d.website) return json_({ ok: true, id: 'x' }); // Spam-Falle

  const err = check_(d);
  if (err) return json_({ ok: false, error: err });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id, isUpdate = false;
  try {
    const ss = SpreadsheetApp.getActive();
    const a = ss.getSheetByName(SH.ANTWORTEN);
    const p = ss.getSheetByName(SH.PERSONEN);
    const now = Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
    const clean = sanitize_(d);

    let row = null;
    if (d.id) {
      row = findRow_(a, d.id);
      if (!row) return json_({ ok: false, error: 'not-found' });
      id = d.id; isUpdate = true;
    } else {
      id = newId_();
    }

    const yes = clean.attending === 'yes';
    const adults = clean.persons.filter(x => x.kind === 'adult').length;
    const kids = clean.persons.length - adults;
    const names = yes ? clean.persons.map(x => x.first + ' ' + x.last).join(', ') : clean.absName;
    const created = row ? row.values[1] : now;
    const line = [id, created, isUpdate ? now : '', yes ? 'Zusage' : 'Absage', yes ? clean.persons.length : 0, adults, kids, names,
      clean.cake ? 'ja' : 'nein', clean.cakeKind, clean.email, clean.phone, clean.message, JSON.stringify(clean)].map(safe_);
    if (row) a.getRange(row.index, 1, 1, line.length).setValues([line]);
    else a.appendRow(line);

    // Personen-Zeilen dieser Antwort ersetzen (Telefon-Einträge bleiben unberührt)
    const pv = p.getDataRange().getValues();
    for (let i = pv.length - 1; i >= 1; i--) if (pv[i][0] === id) p.deleteRow(i + 1);
    const contact = [clean.email, clean.phone].filter(String).join(' / ');
    const cakeCell = clean.cake ? (clean.cakeKind || 'ja') : '';
    let add;
    if (yes) {
      add = clean.persons.map((x, i) => [id, now, 'Formular', 'Zusage', x.first, x.last, x.kind === 'child' ? 'Kind' : 'Erwachsen',
        x.kind === 'child' ? String(x.age) : '', i === 0 ? cakeCell : '', i === 0 ? contact : '', '']);
    } else {
      const parts = clean.absName.split(/\s+/);
      add = [[id, now, 'Formular', 'Absage', clean.absName, parts[parts.length - 1], '', '', '', contact, '']];
    }
    add = add.map(r => r.map(safe_));
    p.getRange(p.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
    markDuplicates_(p);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try { notifyCouple_(sanitize_(d), id, isUpdate); } catch (err) { console.error(err); }
  try { if (d.email) confirmGuest_(sanitize_(d), id, isUpdate); } catch (err) { console.error(err); }
  return json_({ ok: true, id: id });
}

/* Telefon-Einträge von Hand: Dubletten beim Bearbeiten neu markieren */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === SH.PERSONEN && e.range.getColumn() <= 6) markDuplicates_(e.range.getSheet());
}

/* ================= Hilfsfunktionen ================= */

function check_(d) {
  if (d.attending !== 'yes' && d.attending !== 'no') return 'attending';
  if (d.attending === 'yes') {
    if (!Array.isArray(d.persons) || !d.persons.length || d.persons.length > 20) return 'persons';
    for (const x of d.persons) if (!str_(x.first) || !str_(x.last)) return 'name';
    if (!str_(d.email) && !str_(d.phone)) return 'contact';
  } else if (!str_(d.absName)) return 'name';
  return null;
}

function sanitize_(d) {
  const yes = d.attending === 'yes';
  return {
    attending: yes ? 'yes' : 'no',
    persons: yes ? d.persons.slice(0, 20).map(x => ({
      first: str_(x.first, 80), last: str_(x.last, 80),
      kind: x.kind === 'child' ? 'child' : 'adult',
      age: x.kind === 'child' ? Math.max(0, Math.min(17, parseInt(x.age, 10) || 0)) : ''
    })) : [],
    cake: yes && !!d.cake,
    cakeKind: yes && d.cake ? str_(d.cakeKind, 120) : '',
    absName: yes ? '' : str_(d.absName, 200),
    email: str_(d.email, 200),
    phone: str_(d.phone, 50),
    message: str_(d.message, 2000)
  };
}

function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 500); }
function safe_(v) { const s = String(v); return /^[=+\-@]/.test(s) ? "'" + s : v; } // keine Formeln aus Eingaben
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function findRow_(sheet, id) {
  const v = sheet.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === id) return { index: i + 1, values: v[i] };
  return null;
}

function markDuplicates_(p) {
  const last = p.getLastRow();
  if (last < 2) return;
  const v = p.getRange(2, 1, last - 1, 11).getValues();
  const seen = {};
  v.forEach(r => {
    const key = (String(r[4]) + '|' + String(r[5])).toLowerCase().replace(/\s+/g, ' ').trim();
    if (key === '|') return;
    (seen[key] = seen[key] || new Set()).add(r[0] || ('row' + Math.random()));
  });
  const out = v.map(r => {
    const key = (String(r[4]) + '|' + String(r[5])).toLowerCase().replace(/\s+/g, ' ').trim();
    return [seen[key] && seen[key].size > 1 ? DUP : ''];
  });
  p.getRange(2, 11, out.length, 1).setValues(out);
}

function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function ageText_(a) { return a === 0 ? 'unter 1 Jahr' : a + (a === 1 ? ' Jahr' : ' Jahre'); }
function personLine_(x) { return esc_(x.first + ' ' + x.last) + (x.kind === 'child' ? ' · Kind, ' + ageText_(x.age) : ''); }

function totals_() {
  const v = SpreadsheetApp.getActive().getSheetByName(SH.PERSONEN).getDataRange().getValues().slice(1);
  const z = v.filter(r => r[3] === 'Zusage');
  const kids = z.filter(r => r[6] === 'Kind').length;
  const absagen = SpreadsheetApp.getActive().getSheetByName(SH.ANTWORTEN).getDataRange().getValues().slice(1).filter(r => r[3] === 'Absage').length;
  const cakes = v.filter(r => String(r[8]).trim() !== '').length;
  return z.length + ' Personen zugesagt (' + (z.length - kids) + ' Erwachsene, ' + kids + ' Kinder) · ' + absagen + ' Absagen · ' + cakes + ' Kuchen';
}

function wrap_(inner, bg) {
  return '<div style="background:#EFE9E6;padding:24px 12px;font-family:Helvetica,Arial,sans-serif;color:#2B2220">' +
    '<div style="max-width:560px;margin:0 auto;background:' + (bg || '#FFF4EF') + ';border-radius:10px;padding:32px 28px;font-size:16px;line-height:1.5">' +
    inner + '</div></div>';
}
const H1 = 'font-family:Georgia,serif;font-weight:bold;font-size:28px;line-height:1.2;color:#D8261D;margin:0 0 16px';
const BTN = 'display:block;text-align:center;background:#C8231A;color:#ffffff;text-decoration:none;font-weight:bold;font-size:17px;padding:15px;border-radius:12px;margin:8px 0 20px';

function confirmGuest_(d, id, isUpdate) {
  const yes = d.attending === 'yes';
  const link = CONFIG.SITE + '/rsvp?id=' + id;
  const first = yes ? d.persons[0].first : d.absName.split(/\s+/)[0];
  let body = '<h1 style="' + H1 + '">' + (yes ? 'Danke, ' + esc_(first) + '!' : 'Danke fürs Bescheidgeben') + '</h1>';
  if (yes) {
    body += '<p style="margin:0 0 16px">' + (isUpdate ? 'Eure Änderungen sind bei uns angekommen.' : 'Wir haben eure Antwort bekommen und freuen uns riesig, dass ihr dabei seid.') + '</p>' +
      '<div style="background:#ffffff;border:1px solid #EBD8CF;border-radius:12px;padding:14px 18px;margin:0 0 16px;font-size:15px">' +
      '<div style="font-weight:bold;color:#8A1710;margin-bottom:6px">Eure Angaben</div>' +
      d.persons.map(x => '<div>' + personLine_(x) + '</div>').join('') +
      '<div>Kuchen: ' + (d.cake ? esc_(d.cakeKind || 'ja') : 'nein') + '</div></div>' +
      '<p style="margin:0 0 16px;font-size:15px"><b>Samstag, 29. Mai 2027</b><br>Einlass 10:30 Uhr, Beginn 11:00 Uhr<br>K3 Kirche Schwelm, Sedanstraße 12, 58332 Schwelm</p>' +
      '<p style="margin:0">Etwas hat sich geändert? Über diesen Link könnt ihr eure Angaben jederzeit anpassen:</p>' +
      '<a href="' + link + '" style="' + BTN + '">Angaben ändern</a>';
  } else {
    body += '<p style="margin:0 0 16px">Schade, dass ihr nicht dabei sein könnt – wir werden an euch denken.</p>' +
      '<p style="margin:0">Falls sich doch noch etwas ändert, könnt ihr über diesen Link jederzeit nachträglich zusagen:</p>' +
      '<a href="' + link + '" style="' + BTN + '">Doch zusagen</a>';
  }
  body += '<p style="margin:0 0 12px;font-size:14px;color:#5A4A45">Lieber telefonisch? Denia +49 172 6743037 · Dominik +49 1520 6423680</p>' +
    '<p style="margin:0">Liebe Grüße<br>Denia &amp; Dominik</p>';
  MailApp.sendEmail({
    to: d.email,
    subject: yes ? (isUpdate ? 'Eure geänderte Zusage zur Hochzeit von Denia & Dominik' : 'Eure Zusage zur Hochzeit von Denia & Dominik') : 'Eure Rückmeldung zur Hochzeit von Denia & Dominik',
    htmlBody: wrap_(body),
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

function notifyCouple_(d, id, isUpdate) {
  const yes = d.attending === 'yes';
  const names = yes ? d.persons.map(x => x.first + ' ' + x.last) : [d.absName];
  const kind = (isUpdate ? 'Geändert – ' : '') + (yes ? 'Zusage' : 'Absage');
  const subject = kind + ': ' + (yes ? shortNames_(d.persons) : d.absName);
  const tag = yes ? '<span style="background:#DDF0E2;color:#1E5B30;font-weight:bold;font-size:12px;padding:3px 9px;border-radius:99px">' + kind.toUpperCase() + '</span>'
    : '<span style="background:#FCE3DE;color:#8A1710;font-weight:bold;font-size:12px;padding:3px 9px;border-radius:99px">' + kind.toUpperCase() + '</span>';
  const nameHtml = yes
    ? d.persons.map(x => '<div style="font-family:Georgia,serif;font-weight:bold;font-size:26px;line-height:1.25;color:#D8261D">' + esc_(x.first + ' ' + x.last) +
      (x.kind === 'child' ? ' <span style="font-family:Helvetica,Arial,sans-serif;font-weight:normal;font-size:16px;color:#5A4A45">(Kind, ' + ageText_(x.age) + ')</span>' : '') + '</div>').join('')
    : '<div style="font-family:Georgia,serif;font-weight:bold;font-size:26px;line-height:1.25;color:#D8261D">' + esc_(d.absName) + '</div>';
  const cell = (l, v) => '<td style="background:#FFF4EF;border-radius:8px;padding:10px 12px;width:50%;vertical-align:top"><div style="font-size:12px;color:#5A4A45">' + l + '</div><div style="font-weight:bold;font-size:15px">' + v + '</div></td>';
  let body = '<div style="margin-bottom:10px">' + tag + '</div>' + nameHtml +
    '<table role="presentation" cellspacing="8" style="width:100%;margin:14px -8px 6px;border-collapse:separate">';
  if (yes) {
    const kids = d.persons.filter(x => x.kind === 'child').length;
    body += '<tr>' + cell('Kuchen', d.cake ? esc_(d.cakeKind || 'ja') : '–') + cell('Personen', (d.persons.length - kids) + ' Erwachsene, ' + kids + ' Kinder') + '</tr>';
  }
  body += '<tr>' + cell('E-Mail', esc_(d.email || '–')) + cell('Telefon', esc_(d.phone || '–')) + '</tr></table>';
  if (d.message) body += '<div style="font-size:12px;color:#5A4A45;margin-top:6px">Nachricht</div><div style="font-style:italic;margin-bottom:12px">„' + esc_(d.message).replace(/\n/g, '<br>') + '“</div>';
  body += '<div style="border-top:1px solid #EBD8CF;padding-top:12px;margin-top:8px;font-size:14px;color:#5A4A45">Stand gesamt: ' + totals_() + '</div>' +
    '<a href="' + SpreadsheetApp.getActive().getUrl() + '" style="display:block;text-align:center;background:#2B2220;color:#ffffff;text-decoration:none;font-weight:bold;padding:13px;border-radius:10px;margin-top:16px">Google Sheet öffnen</a>';
  MailApp.sendEmail({
    to: CONFIG.NOTIFY.join(','),
    subject: subject,
    htmlBody: wrap_(body, '#ffffff'),
    name: 'RSVP-Formular',
    replyTo: d.email || CONFIG.REPLY_TO
  });
}

function shortNames_(persons) {
  const byLast = {};
  persons.forEach(x => (byLast[x.last] = byLast[x.last] || []).push(x.first));
  return Object.keys(byLast).map(l => {
    const f = byLast[l];
    return (f.length > 1 ? f.slice(0, -1).join(', ') + ' & ' + f[f.length - 1] : f[0]) + ' ' + l;
  }).join(', ');
}
