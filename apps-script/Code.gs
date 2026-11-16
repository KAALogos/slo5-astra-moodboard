/**
 * SLO nr 5 · Moodboard Astra — zapis głosów komisji.
 * Skrypt powiązany z arkuszem „SLO5 Moodboard Astra — głosy komisji”.
 * Wdrożenie: Aplikacja internetowa, wykonuj jako: ja, dostęp: każdy.
 *
 * Arkusz „Glosy” = dziennik zdarzeń (każda zmiana to nowy wiersz).
 * GET zwraca aktualny stan: ostatni wpis dla pary (głosujący, szablon).
 */
const SHEET = 'Glosy';
const HEADER = ['czas', 'glosujacy', 'szablon_id', 'ocena_1_5', 'finalista_0_1'];
const IDS = ['01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function state_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const voters = {};
  for (const r of rows) {
    const name = String(r[1] || '').trim();
    const id = String(r[2] || '').padStart(2, '0');
    if (!name || IDS.indexOf(id) < 0) continue;
    const v = voters[name] = voters[name] || {};
    const e = v[id] = v[id] || {};
    const score = Number(r[3]);
    if (r[3] !== '' && Number.isInteger(score)) {
      if (score >= 1 && score <= 5) e.score = score; else delete e.score;
    }
    if (r[4] !== '') e.fav = Number(r[4]) === 1;
    if (!('score' in e) && !e.fav) delete v[id];
  }
  return { voters: voters, updatedAt: new Date().toISOString() };
}

function doGet() {
  return json_(state_());
}

/** Body (text/plain JSON): {voter, entries:[{id, score?, fav?}]} ; score 0 = usuń ocenę */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const voter = String(body.voter || '').trim().slice(0, 80);
    if (!voter) return json_({ ok: false, error: 'brak_glosujacego' });
    const entries = Array.isArray(body.entries) ? body.entries.slice(0, 200) : [];
    const now = new Date();
    const out = [];
    for (const x of entries) {
      const id = String(x && x.id || '');
      if (IDS.indexOf(id) < 0) continue;
      const score = Number.isInteger(x.score) && x.score >= 0 && x.score <= 5 ? x.score : '';
      const fav = typeof x.fav === 'boolean' ? (x.fav ? 1 : 0) : '';
      if (score === '' && fav === '') continue;
      out.push([now, voter, "'" + id, score, fav]);
    }
    if (out.length) {
      const sh = sheet_();
      sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADER.length).setValues(out);
    }
    return json_({ ok: true, saved: out.length, state: state_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
