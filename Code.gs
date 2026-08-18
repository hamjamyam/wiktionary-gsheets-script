/**
 * Word Tracker — auto-fill vocabulary details in Google Sheets.
 * Single data source: Wiktionary (one API call per word).
 *
 * Type a word in column A (row 2 or below). The row auto-fills:
 *   B: Part of speech   C: Definition   D: IPA
 *   E: Example          F: Etymology    G: Date added
 *
 * All fields are parsed from one fetch of the word's rendered Wiktionary
 * page (English section): first sense per part of speech, first IPA, first
 * usage example, and the first Etymology section.
 *
 * SETUP (one time):
 *   1. In your Sheet: Extensions → Apps Script, paste this whole file, save.
 *   2. Run the `setup` function once from the editor (▶ button) and approve
 *      the permission prompts. This installs the edit trigger.
 *      (A plain onEdit trigger can't call external APIs — installable ones can.)
 *   3. Back in the Sheet, type a word in A2.
 */

var COL = { WORD: 1, POS: 2, DEF: 3, IPA: 4, EXAMPLE: 5, ETYMOLOGY: 6, ADDED: 7 };

var POS_NAMES = 'Noun|Verb|Adjective|Adverb|Interjection|Preposition|Conjunction|' +
  'Pronoun|Determiner|Numeral|Particle|Proper_noun|Phrase';

/** Run this once from the Apps Script editor. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  sheet.getRange(1, 1, 1, 7).setValues([[
    'Word', 'Part of speech', 'Definition', 'IPA', 'Example', 'Etymology', 'Added'
  ]]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(COL.DEF, 320);
  sheet.setColumnWidth(COL.ETYMOLOGY, 420);
  sheet.getRange('C:F').setWrap(true);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onWordEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onWordEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/** Installable edit trigger: fires when column A gets a new word. */
function onWordEdit(e) {
  var range = e.range;
  if (range.getColumn() !== COL.WORD) return;

  var sheet = range.getSheet();
  for (var i = 0; i < range.getNumRows(); i++) {
    var row = range.getRow() + i;
    if (row < 2) continue;
    var word = String(sheet.getRange(row, COL.WORD).getValue()).trim();
    if (!word) {
      // Word deleted → clear the rest of the row
      sheet.getRange(row, COL.POS, 1, 6).clearContent();
      continue;
    }
    // Dupe protection: if the word already lives in another row, flag
    // this one instead of fetching
    var dupRow = findDuplicateRow(sheet, row, word);
    if (dupRow) {
      sheet.getRange(row, COL.POS, 1, 6).clearContent();
      sheet.getRange(row, COL.DEF).setValue('Duplicate — already in row ' + dupRow);
      continue;
    }
    // Changing the word on a filled row refetches with the new word
    fillRow(sheet, row, word);
  }
}

/**
 * Returns the row number of an existing entry for this word (case-insensitive),
 * or 0 if none. The earlier/filled row counts as the original; when the same
 * word is pasted into several rows at once, the first row wins and the rest
 * get flagged.
 */
function findDuplicateRow(sheet, row, word) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var values = sheet.getRange(2, COL.WORD, last - 1, 1).getValues();
  var target = word.toLowerCase();
  for (var i = 0; i < values.length; i++) {
    var r = i + 2;
    if (r === row) continue;
    if (String(values[i][0]).trim().toLowerCase() !== target) continue;
    if (r < row || sheet.getRange(r, COL.DEF).getValue()) return r;
  }
  return 0;
}

function fillRow(sheet, row, word) {
  sheet.getRange(row, COL.DEF).setValue('Looking up…');
  SpreadsheetApp.flush();

  var w = fetchWiktionary(word);

  sheet.getRange(row, COL.POS, 1, 6).setValues([[
    w.pos,
    w.def || 'Not found on Wiktionary',
    w.ipa,
    w.example,
    w.etymology,
    new Date()
  ]]);
}

/**
 * One fetch of the rendered Wiktionary page; everything is parsed from the
 * English section of the HTML.
 *
 * Why full-page: parsing a single section in isolation breaks Wiktionary's
 * templates ("Lua error … does not match the L2 header") on multi-language
 * pages like "run". disableeditsection removes "[edit]" links at the source.
 */
function fetchWiktionary(word) {
  var empty = { pos: '', def: '', ipa: '', example: '', etymology: '' };
  var url = 'https://en.wiktionary.org/w/api.php?format=json&action=parse' +
    '&redirects=1&prop=text&disableeditsection=true' +
    '&page=' + encodeURIComponent(word.toLowerCase());
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return empty;
  var data = JSON.parse(resp.getContentText());
  var html = data.parse && data.parse.text ? data.parse.text['*'] : '';
  if (!html) return empty;

  // Slice out the English language section
  var start = html.indexOf('id="English"');
  if (start === -1) return empty;
  var end = html.indexOf('<h2', start);
  var en = html.substring(start, end === -1 ? html.length : end);

  var defs = extractDefinitions(en);
  return {
    pos: defs.posList.join(', '),
    def: defs.lines.join('\n'),
    ipa: extractIpa(en),
    example: defs.example,
    etymology: extractEtymology(en)
  };
}

/** First IPA transcription in the English section. */
function extractIpa(en) {
  var m = en.match(/<span class="IPA[^"]*"[^>]*>([^<]+)<\/span>/);
  return m ? m[1] : '';
}

/**
 * First sense of each part of speech (deduped), plus the first usage example.
 * POS headings are h3 (or h4/h5 under "Etymology N" groupings); the sense
 * text is the first <li> of the following <ol>, cut at the first nested
 * block so quotations and sub-senses don't leak in.
 */
function extractDefinitions(en) {
  var out = { posList: [], lines: [], example: '' };
  var seen = {};
  var headingRe = new RegExp('<h[3-5] id="(' + POS_NAMES + ')(?:_\\d+)?"', 'g');
  var m;
  while ((m = headingRe.exec(en)) !== null) {
    var pos = m[1].replace(/_/g, ' ').toLowerCase();
    if (seen[pos]) continue;
    var rest = en.substring(m.index + m[0].length);
    var ol = rest.match(/<ol[^>]*>([\s\S]*?)<\/ol>/);
    if (!ol) continue;
    // First NON-empty sense (pages like "bean" open with an empty <li>)
    var liRe = /<li[^>]*>([\s\S]*?)<\/li>/g;
    var li, body = '', text = '';
    while ((li = liRe.exec(ol[1])) !== null) {
      body = li[1];
      var cut = body.search(/<(ol|ul|dl|div)[\s>]/);
      text = htmlToText(cut === -1 ? body : body.substring(0, cut));
      if (text) break;
    }
    if (!text) continue;
    seen[pos] = true;
    out.posList.push(pos);
    out.lines.push(pos + ': ' + text);
    if (!out.example) {
      var ex = body.match(/class="[^"]*e-example[^"]*"[^>]*>([\s\S]*?)<\/i>/);
      if (ex) out.example = htmlToText(ex[1]);
    }
  }
  return out;
}

/** Plain text of the first Etymology section in the English slice. */
function extractEtymology(en) {
  var ety = en.indexOf('id="Etymology');
  if (ety === -1) return '';
  var start = en.indexOf('</div>', ety) + 6;
  var end = en.indexOf('<div class="mw-heading', start);
  var sec = en.substring(start, end === -1 ? en.length : end);
  // Skip the collapsible "Etymology tree" box (and any styling) by starting
  // at the first prose paragraph
  var p = sec.indexOf('<p');
  if (p > 0) sec = sec.substring(p);
  return htmlToText(sec);
}

function htmlToText(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<li[^>]*>/gi, '\n• ')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&#(\d+);/g, function (m, code) { return String.fromCharCode(code); })
    .replace(/&quot;/g, '"').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}
