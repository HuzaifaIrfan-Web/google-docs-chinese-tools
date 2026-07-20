/**
 * 🇨🇳 Google Docs Chinese Tools — by Huzaifa Irfan
 * Editor Add-on for Docs, Sheets, and Slides: tone-colors Chinese text
 * (Pleco-style) and adds pinyin/zhuyin annotations.
 *
 * ARCHITECTURE: this is a STANDALONE add-on, not bound to one file.
 * Settings live in UserProperties (getUserProperties), so they are the
 * same for you across every Doc, Sheet, and Slide deck you use this in.
 *
 * PER-HOST LIMITATIONS (platform limits, not fixable in script):
 * - Sheets: no rich-text superscript for cells -> annotations are always
 *   inline "字(zì)", never floating superscript.
 * - Slides: same -> inline only, no superscript, no "line below" mode
 *   (a shape has no fixed "paragraph flow" to insert a line under).
 * - Docs: full feature set (superscript char mode + whole-line mode).
 * - Tables/footnotes skipped everywhere in v1.
 * - Polyphonic characters (多音字) use the first/most common reading only.
 * - Slides "Copy Selection" copies the whole text box's content (Slides'
 *   selection API doesn't expose partial in-shape character ranges the
 *   way Docs/Sheets do).
 */

/* ============================== CONFIG ============================== */

const TONE_COLORS_DEFAULT = {
  1: '#E53935', 2: '#43A047', 3: '#1E88E5', 4: '#8E24AA', 5: '#757575'
};

const PRESETS = {
  // Similar to Pleco
  pleco: {
    1: '#E53935', // Red
    2: '#43A047', // Green
    3: '#1E88E5', // Blue
    4: '#8E24AA', // Purple
    5: '#757575'  // Gray
  },

  // Strong colors, excellent readability
  vivid: {
    1: '#C62828', // Dark Red
    2: '#2E7D32', // Dark Green
    3: '#1565C0', // Dark Blue
    4: '#6A1B9A', // Dark Purple
    5: '#616161'  // Gray
  },

  // Soft but still readable
  pastel: {
    1: '#D84315', // Soft Orange-Red
    2: '#558B2F', // Olive Green
    3: '#1976D2', // Medium Blue
    4: '#7B1FA2', // Medium Purple
    5: '#9E9E9E'
  },

  // Black & White friendly
  grayscale: {
    1: '#000000',
    2: '#303030',
    3: '#505050',
    4: '#707070',
    5: '#A0A0A0'
  },

  // Colorblind-friendly (Okabe-Ito palette)
  colorblind: {
    1: '#D55E00', // Vermillion
    2: '#009E73', // Bluish Green
    3: '#0072B2', // Blue
    4: '#CC79A7', // Reddish Purple
    5: '#7F7F7F'
  },

  // Material Design palette
  material: {
    1: '#D32F2F',
    2: '#388E3C',
    3: '#1976D2',
    4: '#7B1FA2',
    5: '#757575'
  }
};

const PINYIN_DATA_URL = 'https://raw.githubusercontent.com/mozillazg/pinyin-data/master/pinyin.txt';
const CACHE_PREFIX = 'PYDICT_';
const CACHE_CHUNKS_KEY = 'PYDICT_CHUNKS';
const CACHE_TTL_SEC = 21600;
const MARKER = '\u200B';

/* ========================== ADD-ON ENTRY POINTS ========================== */
/*
 * This add-on's manifest registers homepageTrigger entries under
 * addOns.common and addOns.docs/sheets/slides, which makes this a unified
 * Google Workspace Add-on. That means its entry point is the icon in the
 * RIGHT-HAND SIDE PANEL of Docs/Sheets/Slides (the same strip Calendar,
 * Keep, and Tasks live in) -- NOT the Extensions menu. There is
 * intentionally no onOpen()/createAddonMenu() here anymore; that's the
 * older, separate "Editor Add-on" framework and mixing the two is what
 * caused the menu to never appear.
 */

function onHomepage(e) {
  const host = getActiveHost_();
  const hostLabel = {
    docs: 'Google Docs',
    sheets: 'Google Sheets',
    slides: 'Google Slides'
  }[host] || 'this app';

  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle('🇨🇳 Google Docs Chinese Tools')
        .setSubtitle('by Huzaifa Irfan')
    )
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText(
            'Tone-color Chinese text and add Pinyin/Zhuyin annotations in ' +
            hostLabel +
            '.'
          )
        )
        .addWidget(
          CardService.newTextButton()
            .setText('🧰 Open Chinese Tools')
            .setOnClickAction(
              CardService.newAction().setFunctionName('showSidebar_')
            )
        )
    )
    .addSection(
      CardService.newCardSection()
        .setHeader('Resources')
        .addWidget(
          CardService.newTextButton()
            .setText('📦 GitHub Repository')
            .setOpenLink(
              CardService.newOpenLink().setUrl(
                'https://github.com/HuzaifaIrfan-Web/google-docs-chinese-tools'
              )
            )
        )
        .addWidget(
          CardService.newTextButton()
            .setText('🌐 huzaifairfan.com')
            .setOpenLink(
              CardService.newOpenLink().setUrl(
                'https://huzaifairfan.com/'
              )
            )
        )
        .addWidget(
          CardService.newTextButton()
            .setText('✉️ huzaifairfan2001@gmail.com')
            .setOpenLink(
              CardService.newOpenLink().setUrl(
                'mailto:huzaifairfan2001@gmail.com'
              )
            )
        )
    )
    .build();
}

function showSidebar_() {
  showSidebar();
  return CardService.newActionResponseBuilder().build();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('🇨🇳 Google Docs Chinese Tools')
    .setWidth(340);
  const ui = getUi_();
  if (ui) ui.showSidebar(html);
}

/* ============================ HOST DETECTION ============================ */

function getActiveHost_() {
  try { if (DocumentApp.getActiveDocument()) return 'docs'; } catch (e) { }
  try { if (SpreadsheetApp.getActiveSpreadsheet()) return 'sheets'; } catch (e) { }
  try { if (SlidesApp.getActivePresentation()) return 'slides'; } catch (e) { }
  return null;
}

function getUi_() {
  const host = getActiveHost_();
  if (host === 'docs') return DocumentApp.getUi();
  if (host === 'sheets') return SpreadsheetApp.getUi();
  if (host === 'slides') return SlidesApp.getUi();
  return null;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================ SETTINGS (USER-SCOPED) ============================ */

function getDefaultSettings_() {
  return {
    preset: 'pleco',
    colors: Object.assign({}, TONE_COLORS_DEFAULT),
    colorEnabled: true,
    annotationEnabled: false,
    annotationType: 'pinyin',
    annotationMode: 'char',    // 'line' only actually applies on Docs
    colorAnnotation: false,
    autoMode: false,           // Docs-only (time trigger)
    copyMode: 'both'
  };
}

function getSettings() {
  const raw = PropertiesService.getUserProperties().getProperty('CT_SETTINGS');
  const defaults = getDefaultSettings_();
  let merged = defaults;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      merged = Object.assign({}, defaults, parsed, {
        colors: Object.assign({}, defaults.colors, parsed.colors || {})
      });
    } catch (e) { merged = defaults; }
  }
  return { settings: merged, host: getActiveHost_() };
}

function saveSettings(settings) {
  PropertiesService.getUserProperties().setProperty('CT_SETTINGS', JSON.stringify(settings));
  const host = getActiveHost_();
  if (host === 'docs') {
    if (settings.autoMode) installAutoTrigger_(); else removeAutoTrigger_();
  }
  return { ok: true };
}

function getPresetColors(presetName) {
  return PRESETS[presetName] || TONE_COLORS_DEFAULT;
}

/* ======================= PINYIN DATA (cached, shared across hosts) ======================= */

function fetchAndCachePinyinDict_() {
  const resp = UrlFetchApp.fetch(PINYIN_DATA_URL, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to fetch pinyin dataset: HTTP ' + resp.getResponseCode());
  }
  const lines = resp.getContentText().split('\n');
  const dict = {};
  const lineRe = /^U\+[0-9A-Fa-f]+:\s*([^#]+)#\s*(\S+)/;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line || line.charAt(0) === '#') continue;
    const m = lineRe.exec(line);
    if (!m) continue;
    const readings = m[1].trim().split(',');
    const char = m[2].trim();
    if (char && readings.length && readings[0].trim()) dict[char] = readings[0].trim();
  }
  const json = JSON.stringify(dict);
  const cache = CacheService.getScriptCache();
  const chunkSize = 90000;
  const chunks = [];
  for (let i = 0; i < json.length; i += chunkSize) chunks.push(json.substring(i, i + chunkSize));
  const payload = {};
  chunks.forEach(function (c, idx) { payload[CACHE_PREFIX + idx] = c; });
  payload[CACHE_CHUNKS_KEY] = String(chunks.length);
  cache.putAll(payload, CACHE_TTL_SEC);
  return dict;
}

function getPinyinDict_() {
  const cache = CacheService.getScriptCache();
  const countStr = cache.get(CACHE_CHUNKS_KEY);
  if (countStr) {
    const count = parseInt(countStr, 10);
    const keys = [];
    for (let i = 0; i < count; i++) keys.push(CACHE_PREFIX + i);
    const parts = cache.getAll(keys);
    let json = '', ok = true;
    for (let i = 0; i < count; i++) {
      const part = parts[CACHE_PREFIX + i];
      if (!part) { ok = false; break; }
      json += part;
    }
    if (ok) { try { return JSON.parse(json); } catch (e) { } }
  }
  return fetchAndCachePinyinDict_();
}

function refreshPinyinData() {
  fetchAndCachePinyinDict_();
  return { ok: true, message: 'Pinyin data refreshed.' };
}

/* ======================= TONE PARSING (shared) ======================= */

const TONE_MARKS_ = {
  'ā': ['a', 1], 'á': ['a', 2], 'ǎ': ['a', 3], 'à': ['a', 4],
  'ē': ['e', 1], 'é': ['e', 2], 'ě': ['e', 3], 'è': ['e', 4],
  'ī': ['i', 1], 'í': ['i', 2], 'ǐ': ['i', 3], 'ì': ['i', 4],
  'ō': ['o', 1], 'ó': ['o', 2], 'ǒ': ['o', 3], 'ò': ['o', 4],
  'ū': ['u', 1], 'ú': ['u', 2], 'ǔ': ['u', 3], 'ù': ['u', 4],
  'ǖ': ['ü', 1], 'ǘ': ['ü', 2], 'ǚ': ['ü', 3], 'ǜ': ['ü', 4],
  'ń': ['n', 2], 'ň': ['n', 3], 'ǹ': ['n', 4], 'ḿ': ['m', 2]
};

function parsePinyinSyllable_(raw) {
  if (!raw) return { plain: '', tone: 5 };
  const numMatch = raw.match(/^([a-zA-Züv]+)([1-5])$/i);
  if (numMatch) return { plain: numMatch[1].toLowerCase(), tone: parseInt(numMatch[2], 10) };
  let tone = 5, plain = '';
  for (const ch of raw) {
    if (TONE_MARKS_[ch]) { plain += TONE_MARKS_[ch][0]; tone = TONE_MARKS_[ch][1]; }
    else plain += ch;
  }
  return { plain: plain.toLowerCase(), tone: tone };
}

const ZHUYIN_INITIALS_ = [['zh', 'ㄓ'], ['ch', 'ㄔ'], ['sh', 'ㄕ'], ['b', 'ㄅ'], ['p', 'ㄆ'], ['m', 'ㄇ'], ['f', 'ㄈ'],
['d', 'ㄉ'], ['t', 'ㄊ'], ['n', 'ㄋ'], ['l', 'ㄌ'], ['g', 'ㄍ'], ['k', 'ㄎ'], ['h', 'ㄏ'], ['j', 'ㄐ'], ['q', 'ㄑ'],
['x', 'ㄒ'], ['r', 'ㄖ'], ['z', 'ㄗ'], ['c', 'ㄘ'], ['s', 'ㄙ']];

const ZHUYIN_SPECIAL_SYLLABLES_ = {
  'zhi': 'ㄓ', 'chi': 'ㄔ', 'shi': 'ㄕ', 'ri': 'ㄖ', 'zi': 'ㄗ', 'ci': 'ㄘ', 'si': 'ㄙ',
  'yi': 'ㄧ', 'ya': 'ㄧㄚ', 'ye': 'ㄧㄝ', 'yao': 'ㄧㄠ', 'you': 'ㄧㄡ', 'yan': 'ㄧㄢ', 'yin': 'ㄧㄣ',
  'yang': 'ㄧㄤ', 'ying': 'ㄧㄥ', 'yong': 'ㄩㄥ', 'yu': 'ㄩ', 'yue': 'ㄩㄝ', 'yuan': 'ㄩㄢ', 'yun': 'ㄩㄣ',
  'wu': 'ㄨ', 'wa': 'ㄨㄚ', 'wo': 'ㄨㄛ', 'wai': 'ㄨㄞ', 'wei': 'ㄨㄟ', 'wan': 'ㄨㄢ', 'wen': 'ㄨㄣ',
  'wang': 'ㄨㄤ', 'weng': 'ㄨㄥ', 'er': 'ㄦ', 'e': 'ㄜ', 'a': 'ㄚ', 'o': 'ㄛ', 'ai': 'ㄞ', 'ei': 'ㄟ',
  'ao': 'ㄠ', 'ou': 'ㄡ', 'an': 'ㄢ', 'en': 'ㄣ', 'ang': 'ㄤ', 'eng': 'ㄥ'
};

const ZHUYIN_FINALS_ = [['iang', 'ㄧㄤ'], ['iong', 'ㄩㄥ'], ['uang', 'ㄨㄤ'], ['ueng', 'ㄨㄥ'],
['ian', 'ㄧㄢ'], ['iao', 'ㄧㄠ'], ['uai', 'ㄨㄞ'], ['uan', 'ㄨㄢ'], ['van', 'ㄩㄢ'], ['ing', 'ㄧㄥ'],
['ang', 'ㄤ'], ['eng', 'ㄥ'], ['ong', 'ㄨㄥ'], ['ia', 'ㄧㄚ'], ['ie', 'ㄧㄝ'], ['iu', 'ㄧㄡ'], ['in', 'ㄧㄣ'],
['ua', 'ㄨㄚ'], ['uo', 'ㄨㄛ'], ['ui', 'ㄨㄟ'], ['un', 'ㄨㄣ'], ['ve', 'ㄩㄝ'], ['ue', 'ㄩㄝ'], ['vn', 'ㄩㄣ'],
['ai', 'ㄞ'], ['ei', 'ㄟ'], ['ao', 'ㄠ'], ['ou', 'ㄡ'], ['an', 'ㄢ'], ['en', 'ㄣ'], ['er', 'ㄦ'],
['a', 'ㄚ'], ['o', 'ㄛ'], ['e', 'ㄜ'], ['i', 'ㄧ'], ['u', 'ㄨ'], ['ü', 'ㄩ'], ['v', 'ㄩ']];

const TONE_MARK_SUFFIX_ = { 1: '', 2: 'ˊ', 3: 'ˇ', 4: 'ˋ', 5: '' };

function pinyinToZhuyin_(plain, tone) {
  let base = ZHUYIN_SPECIAL_SYLLABLES_[plain];
  if (!base) {
    let rest = plain, initialZhuyin = '';
    for (const pair of ZHUYIN_INITIALS_) {
      if (rest.indexOf(pair[0]) === 0) { initialZhuyin = pair[1]; rest = rest.substring(pair[0].length); break; }
    }
    let finalZhuyin = '';
    for (const pair of ZHUYIN_FINALS_) { if (rest === pair[0]) { finalZhuyin = pair[1]; break; } }
    if (!finalZhuyin && rest) finalZhuyin = rest;
    base = initialZhuyin + finalZhuyin;
  }
  if (!base) base = plain;
  if (tone === 5) return '˙' + base;
  return base + (TONE_MARK_SUFFIX_[tone] || '');
}

function buildAnnotationString_(rawReading, parsed, type) {
  const zhuyin = pinyinToZhuyin_(parsed.plain, parsed.tone);
  if (type === 'zhuyin') return zhuyin;
  if (type === 'both') return rawReading + '·' + zhuyin;
  return rawReading;
}

function isChinese_(ch) {
  if (!ch) return false;
  const code = ch.codePointAt(0);
  return (code >= 0x4E00 && code <= 0x9FFF) || (code >= 0x3400 && code <= 0x4DBF);
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ============================ PUBLIC DISPATCHERS ============================ */
/* Client calls these regardless of host; they route to the right engine. */

function convertSelection() {
  const host = getActiveHost_();
  if (host === 'docs') return DocsEngine.convertSelection();
  if (host === 'sheets') return SheetsEngine.convertSelection();
  if (host === 'slides') return SlidesEngine.convertSelection();
  return { ok: false, message: 'Could not detect the active app.' };
}

function convertDocument() {
  const host = getActiveHost_();
  if (host === 'docs') return DocsEngine.convertAll();
  if (host === 'sheets') return SheetsEngine.convertAll();
  if (host === 'slides') return SlidesEngine.convertAll();
  return { ok: false, message: 'Could not detect the active app.' };
}

function clearAll() {
  const host = getActiveHost_();
  let result;
  if (host === 'docs') result = DocsEngine.clearAll();
  else if (host === 'sheets') result = SheetsEngine.clearAll();
  else if (host === 'slides') result = SlidesEngine.clearAll();
  else return { ok: false, message: 'Could not detect the active app.' };

  const settings = getSettings().settings;
  settings.autoMode = false;
  saveSettings(settings);
  return result;
}

function getSelectionMarkup(copyMode) {
  const host = getActiveHost_();
  if (host === 'docs') return DocsEngine.getSelectionMarkup(copyMode);
  if (host === 'sheets') return SheetsEngine.getSelectionMarkup(copyMode);
  if (host === 'slides') return SlidesEngine.getSelectionMarkup(copyMode);
  return { ok: false, message: 'Could not detect the active app.' };
}

/* ============================ DOCS ENGINE ============================ */

const DocsEngine = {
  collectTextElements_: function (element, out) {
    const type = element.getType();
    if (type === DocumentApp.ElementType.TEXT) { out.push(element.asText()); return; }
    if (typeof element.getNumChildren === 'function') {
      const count = element.getNumChildren();
      for (let i = 0; i < count; i++) this.collectTextElements_(element.getChild(i), out);
    }
  },

  processTextElement_: function (textEl, settings, dict, colorsMap, boundStart, boundEnd) {
    const str = textEl.getText();
    const n = str.length;
    if (n === 0) return;
    const s = (boundStart === undefined) ? 0 : boundStart;
    const e = (boundEnd === undefined) ? n - 1 : boundEnd;

    if (settings.colorEnabled) {
      for (let i = s; i <= e; i++) {
        const ch = str[i];
        if (isChinese_(ch)) {
          const reading = dict[ch];
          const tone = reading ? parsePinyinSyllable_(reading).tone : 5;
          textEl.setForegroundColor(i, i, colorsMap[tone] || colorsMap[5]);
        }
      }
    }

    if (settings.annotationEnabled && settings.annotationMode === 'char') {
      for (let i = e; i >= s; i--) {
        const ch = str[i];
        if (!isChinese_(ch)) continue;
        if (str[i + 1] === MARKER) continue;
        const reading = dict[ch];
        if (!reading) continue;
        const parsed = parsePinyinSyllable_(reading);
        const annotation = buildAnnotationString_(reading, parsed, settings.annotationType);
        const insertStr = '(' + annotation + ')' + MARKER;
        textEl.insertText(i + 1, insertStr);
        const startA = i + 1, endA = startA + insertStr.length - 1;
        textEl.setTextAlignment(startA, endA, DocumentApp.TextAlignment.SUPERSCRIPT);
        if (settings.colorAnnotation) textEl.setForegroundColor(startA, endA, colorsMap[parsed.tone] || colorsMap[5]);
      }
    }
  },

  applyLineAnnotations_: function (body, settings, dict, colorsMap) {
    const n = body.getNumChildren();
    for (let idx = n - 1; idx >= 0; idx--) {
      const child = body.getChild(idx);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      const text = child.asParagraph().getText();
      if (text.indexOf(MARKER) === 0) continue;

      const items = [];
      for (const ch of text) { const reading = dict[ch]; if (isChinese_(ch) && reading) items.push({ ch, reading }); }
      if (items.length === 0) continue;

      if (idx + 1 < body.getNumChildren()) {
        const next = body.getChild(idx + 1);
        if (next.getType() === DocumentApp.ElementType.PARAGRAPH && next.asParagraph().getText().indexOf(MARKER) === 0) {
          body.removeChild(next);
        }
      }

      let lineStr = MARKER;
      const segments = [];
      items.forEach(function (item, i) {
        const parsed = parsePinyinSyllable_(item.reading);
        const ann = buildAnnotationString_(item.reading, parsed, settings.annotationType);
        const segStart = lineStr.length;
        lineStr += ann;
        segments.push({ start: segStart, end: lineStr.length - 1, tone: parsed.tone });
        if (i < items.length - 1) lineStr += '  ';
      });

      const newPara = body.insertParagraph(idx + 1, lineStr);
      newPara.setFontSize(9);
      newPara.setItalic(true);
      const t = newPara.editAsText();
      t.setForegroundColor(0, lineStr.length - 1, '#9E9E9E');
      if (settings.colorAnnotation) {
        segments.forEach(function (seg) { t.setForegroundColor(seg.start, seg.end, colorsMap[seg.tone] || colorsMap[5]); });
      }
    }
  },

  convertSelection: function () {
    const settings = getSettings().settings;
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) return { ok: false, message: 'Select some Chinese text first, then click Convert Selection.' };
    const dict = getPinyinDict_();
    const colorsMap = settings.colors;
    let count = 0;
    selection.getRangeElements().forEach(re => {
      const el = re.getElement();
      if (el.getType() !== DocumentApp.ElementType.TEXT) return;
      const textEl = el.asText();
      const start = re.isPartial() ? re.getStartOffset() : 0;
      const end = re.isPartial() ? re.getEndOffsetInclusive() : textEl.getText().length - 1;
      this.processTextElement_(textEl, settings, dict, colorsMap, start, end);
      count++;
    });
    if (settings.annotationEnabled && settings.annotationMode === 'line') {
      return { ok: true, message: 'Colored selection. Whole-line annotation applies doc-wide — use "Convert Whole Document" for that.' };
    }
    return { ok: true, message: 'Selection updated (' + count + ' text run(s)).' };
  },

  convertAll: function () {
    const settings = getSettings().settings;
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const dict = getPinyinDict_();
    const colorsMap = settings.colors;
    const textElements = [];
    this.collectTextElements_(body, textElements);
    textElements.forEach(t => this.processTextElement_(t, settings, dict, colorsMap));
    if (settings.annotationEnabled && settings.annotationMode === 'line') this.applyLineAnnotations_(body, settings, dict, colorsMap);
    return { ok: true, message: 'Document updated (' + textElements.length + ' text runs processed).' };
  },

  clearAll: function () {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    for (let idx = body.getNumChildren() - 1; idx >= 0; idx--) {
      const child = body.getChild(idx);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH && child.asParagraph().getText().indexOf(MARKER) === 0) {
        body.removeChild(child);
      }
    }
    const textElements = [];
    this.collectTextElements_(body, textElements);
    const re = /([\u4E00-\u9FFF\u3400-\u4DBF])\(([^)]*)\)\u200B/;
    textElements.forEach(textEl => {
      let changed = true;
      while (changed) {
        changed = false;
        const m = re.exec(textEl.getText());
        if (m) {
          const start = m.index + m[1].length;
          const end = start + ('(' + m[2] + ')' + MARKER).length - 1;
          textEl.deleteText(start, end);
          changed = true;
        }
      }
      const n = textEl.getText().length;
      if (n > 0) textEl.setForegroundColor(0, n - 1, '#000000');
    });
    removeAutoTrigger_();
    return { ok: true, message: 'Cleared all tone colors and annotations from the document.' };
  },

  getSelectionMarkup: function (copyMode) {
    const settings = getSettings().settings;
    copyMode = copyMode || 'both';
    const doc = DocumentApp.getActiveDocument();
    const selection = doc.getSelection();
    if (!selection) return { ok: false, message: 'Select some text first, then click Copy Selection.' };
    const dict = getPinyinDict_();
    const colorsMap = settings.colors;
    let hanziHtml = '', hanziText = '', hanziInlineHtml = '', hanziInlineText = '';
    const annItems = [];

    selection.getRangeElements().forEach(re => {
      const el = re.getElement();
      if (el.getType() !== DocumentApp.ElementType.TEXT) return;
      const textEl = el.asText();
      const str = textEl.getText();
      const start = re.isPartial() ? re.getStartOffset() : 0;
      const end = re.isPartial() ? re.getEndOffsetInclusive() : str.length - 1;
      for (let i = start; i <= end; i++) {
        const ch = str[i];
        if (isChinese_(ch)) {
          const reading = dict[ch];
          const parsed = reading ? parsePinyinSyllable_(reading) : { plain: '', tone: 5 };
          const toneColor = colorsMap[parsed.tone] || colorsMap[5];
          const charSpan = settings.colorEnabled ? '<span style="color:' + toneColor + '">' + escapeHtml_(ch) + '</span>' : escapeHtml_(ch);
          hanziHtml += charSpan; hanziText += ch;
          if (reading) {
            const ann = buildAnnotationString_(reading, parsed, settings.annotationType);
            const annStyled = settings.colorAnnotation ? '<span style="color:' + toneColor + '">' + escapeHtml_(ann) + '</span>' : escapeHtml_(ann);
            hanziInlineHtml += charSpan + '<sup style="font-size:0.65em">(' + annStyled + ')</sup>';
            hanziInlineText += ch + '(' + ann + ')';
            annItems.push({ html: annStyled, text: ann });
          } else { hanziInlineHtml += charSpan; hanziInlineText += ch; }
        } else {
          const esc = escapeHtml_(ch);
          hanziHtml += esc; hanziText += ch; hanziInlineHtml += esc; hanziInlineText += ch;
        }
      }
    });

    if (copyMode === 'annotation') {
      if (!annItems.length) return { ok: false, message: 'No recognized Chinese characters in the selection.' };
      return { ok: true, html: '<div>' + annItems.map(a => a.html).join(' ') + '</div>', text: annItems.map(a => a.text).join(' ') };
    }
    if (copyMode === 'chinese') return { ok: true, html: '<div>' + hanziHtml + '</div>', text: hanziText };
    if (settings.annotationMode === 'line' && annItems.length) {
      return {
        ok: true,
        html: '<div>' + hanziHtml + '<br><span style="font-size:0.8em;color:#9E9E9E">' + annItems.map(a => a.html).join('&nbsp;&nbsp;') + '</span></div>',
        text: hanziText + '\n' + annItems.map(a => a.text).join('  ')
      };
    }
    return { ok: true, html: '<div>' + hanziInlineHtml + '</div>', text: hanziInlineText };
  }
};

/* ============================ SHEETS ENGINE ============================ */
/* Sheets rich text has no superscript, so annotation is always inline "字(zì)" */

const SheetsEngine = {
  colorAndAnnotateCell_: function (cell, settings, dict, colorsMap) {
    const str = cell.getValue();
    if (typeof str !== 'string' || !str) return false;
    const built = this.buildAnnotatedString_(str, settings, dict);
    const finalBuilder = SpreadsheetApp.newRichTextValue().setText(built.text);
    built.runs.forEach(run => {
      const style = SpreadsheetApp.newTextStyle();
      if (run.color) style.setForegroundColor(run.color);
      finalBuilder.setTextStyle(run.start, run.end + 1, style.build());
    });
    cell.setRichTextValue(finalBuilder.build());
    return true;
  },

  // Builds the final string (original text + inline annotations if enabled)
  // plus a list of {start,end,color} runs to style.
  buildAnnotatedString_: function (str, settings, dict) {
    const colorsMap = settings.colors;
    let out = '';
    const runs = [];
    for (const ch of str) {
      const charStart = out.length;
      out += ch;
      if (isChinese_(ch)) {
        const reading = dict[ch];
        const tone = reading ? parsePinyinSyllable_(reading).tone : 5;
        if (settings.colorEnabled) runs.push({ start: charStart, end: out.length - 1, color: colorsMap[tone] || colorsMap[5] });
        if (settings.annotationEnabled && reading) {
          const parsed = parsePinyinSyllable_(reading);
          const ann = buildAnnotationString_(reading, parsed, settings.annotationType);
          const annStart = out.length;
          out += '(' + ann + ')';
          if (settings.colorAnnotation) runs.push({ start: annStart, end: out.length - 1, color: colorsMap[parsed.tone] || colorsMap[5] });
        }
      }
    }
    return { text: out, runs: runs };
  },

  convertSelection: function () {
    const settings = getSettings().settings;
    const range = SpreadsheetApp.getActiveRange();
    if (!range) return { ok: false, message: 'Select some cells first, then click Convert Selection.' };
    const dict = getPinyinDict_();
    let count = 0;
    const numRows = range.getNumRows(), numCols = range.getNumColumns();
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const cell = range.getCell(r + 1, c + 1);
        if (this.colorAndAnnotateCell_(cell, settings, dict, settings.colors)) count++;
      }
    }
    return { ok: true, message: 'Updated ' + count + ' cell(s).' };
  },

  convertAll: function () {
    const settings = getSettings().settings;
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const dict = getPinyinDict_();
    let count = 0;
    const numRows = range.getNumRows(), numCols = range.getNumColumns();
    for (let r = 1; r <= numRows; r++) {
      for (let c = 1; c <= numCols; c++) {
        const cell = sheet.getRange(r, c);
        if (this.colorAndAnnotateCell_(cell, settings, dict, settings.colors)) count++;
      }
    }
    return { ok: true, message: 'Active sheet updated (' + count + ' cell(s) processed).' };
  },

  clearAll: function () {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const numRows = range.getNumRows(), numCols = range.getNumColumns();
    const re = /([\u4E00-\u9FFF\u3400-\u4DBF])\(([^)]*)\)/g;
    let count = 0;
    for (let r = 1; r <= numRows; r++) {
      for (let c = 1; c <= numCols; c++) {
        const cell = sheet.getRange(r, c);
        const val = cell.getValue();
        if (typeof val !== 'string' || !val) continue;
        const stripped = val.replace(re, '$1');
        if (stripped !== val) { cell.setValue(stripped); count++; }
        else {
          const rich = SpreadsheetApp.newRichTextValue().setText(val)
            .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('#000000').build())
            .build();
          cell.setRichTextValue(rich);
        }
      }
    }
    return { ok: true, message: 'Cleared annotations and reset colors on the active sheet (' + count + ' cell(s) had annotations removed).' };
  },

  getSelectionMarkup: function (copyMode) {
    const settings = getSettings().settings;
    copyMode = copyMode || 'both';
    const range = SpreadsheetApp.getActiveRange();
    if (!range) return { ok: false, message: 'Select some cells first, then click Copy Selection.' };
    const dict = getPinyinDict_();
    const colorsMap = settings.colors;
    const values = range.getValues();
    let hanziHtml = '', hanziText = '', inlineHtml = '', inlineText = '';
    const annItems = [];
    values.forEach(row => {
      row.forEach(val => {
        if (typeof val !== 'string') return;
        for (const ch of val) {
          if (isChinese_(ch)) {
            const reading = dict[ch];
            const parsed = reading ? parsePinyinSyllable_(reading) : { plain: '', tone: 5 };
            const toneColor = colorsMap[parsed.tone] || colorsMap[5];
            const span = settings.colorEnabled ? '<span style="color:' + toneColor + '">' + escapeHtml_(ch) + '</span>' : escapeHtml_(ch);
            hanziHtml += span; hanziText += ch;
            if (reading) {
              const ann = buildAnnotationString_(reading, parsed, settings.annotationType);
              const annStyled = settings.colorAnnotation ? '<span style="color:' + toneColor + '">' + escapeHtml_(ann) + '</span>' : escapeHtml_(ann);
              inlineHtml += span + '(' + annStyled + ')';
              inlineText += ch + '(' + ann + ')';
              annItems.push({ html: annStyled, text: ann });
            } else { inlineHtml += span; inlineText += ch; }
          } else {
            const esc = escapeHtml_(ch);
            hanziHtml += esc; hanziText += ch; inlineHtml += esc; inlineText += ch;
          }
        }
        hanziHtml += ' '; hanziText += ' '; inlineHtml += ' '; inlineText += ' ';
      });
    });
    if (copyMode === 'annotation') {
      if (!annItems.length) return { ok: false, message: 'No recognized Chinese characters in the selection.' };
      return { ok: true, html: '<div>' + annItems.map(a => a.html).join(' ') + '</div>', text: annItems.map(a => a.text).join(' ') };
    }
    if (copyMode === 'chinese') return { ok: true, html: '<div>' + hanziHtml + '</div>', text: hanziText };
    return { ok: true, html: '<div>' + inlineHtml + '</div>', text: inlineText };
  }
};

/* ============================ SLIDES ENGINE ============================ */
/* Slides TextRange has no superscript either -> inline-only annotation. */

const SlidesEngine = {
  processTextRange_: function (textRange, settings, dict, colorsMap) {
    const str = textRange.asString();
    if (!str) return;
    // Color first (doesn't change length), forward.
    if (settings.colorEnabled) {
      for (let i = 0; i < str.length; i++) {
        const ch = str[i];
        if (isChinese_(ch)) {
          const reading = dict[ch];
          const tone = reading ? parsePinyinSyllable_(reading).tone : 5;
          textRange.getRange(i, i + 1).getTextStyle().setForegroundColor(colorsMap[tone] || colorsMap[5]);
        }
      }
    }
    // Then annotate backward so earlier indices stay valid.
    if (settings.annotationEnabled) {
      for (let i = str.length - 1; i >= 0; i--) {
        const ch = str[i];
        if (!isChinese_(ch)) continue;
        if (str[i + 1] === MARKER) continue;
        const reading = dict[ch];
        if (!reading) continue;
        const parsed = parsePinyinSyllable_(reading);
        const ann = buildAnnotationString_(reading, parsed, settings.annotationType);
        const insertStr = '(' + ann + ')' + MARKER;
        textRange.getRange(i + 1, i + 1).insertText(insertStr); // insert after char i
        const startA = i + 1, endA = startA + insertStr.length;
        if (settings.colorAnnotation) {
          textRange.getRange(startA, endA).getTextStyle().setForegroundColor(colorsMap[parsed.tone] || colorsMap[5]);
        }
      }
    }
  },

  collectShapes_: function (slide, out) {
    slide.getShapes().forEach(shape => { if (shape.getText) out.push(shape); });
  },

  convertSelection: function () {
    const settings = getSettings().settings;
    const selection = SlidesApp.getActivePresentation().getSelection();
    const pageElements = selection.getPageElementRange();
    if (!pageElements) return { ok: false, message: 'Select a text box (or text inside one) first, then click Convert Selection.' };
    const dict = getPinyinDict_();
    let count = 0;
    pageElements.getPageElements().forEach(pe => {
      if (pe.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = pe.asShape();
        if (shape.getText) { this.processTextRange_(shape.getText(), settings, dict, settings.colors); count++; }
      }
    });
    return { ok: true, message: 'Updated ' + count + ' text box(es).' };
  },

  convertAll: function () {
    const settings = getSettings().settings;
    const pres = SlidesApp.getActivePresentation();
    const dict = getPinyinDict_();
    let count = 0;
    pres.getSlides().forEach(slide => {
      const shapes = [];
      this.collectShapes_(slide, shapes);
      shapes.forEach(shape => { this.processTextRange_(shape.getText(), settings, dict, settings.colors); count++; });
    });
    return { ok: true, message: 'Presentation updated (' + count + ' text box(es) processed).' };
  },

  clearAll: function () {
    const pres = SlidesApp.getActivePresentation();
    const re = /([\u4E00-\u9FFF\u3400-\u4DBF])\(([^)]*)\)\u200B/;
    let count = 0;
    pres.getSlides().forEach(slide => {
      const shapes = [];
      this.collectShapes_(slide, shapes);
      shapes.forEach(shape => {
        const tr = shape.getText();
        let changed = true;
        while (changed) {
          changed = false;
          const str = tr.asString();
          const m = re.exec(str);
          if (m) {
            const start = m.index + m[1].length;
            const end = start + ('(' + m[2] + ')' + MARKER).length;
            tr.getRange(start, end).clear();
            changed = true;
          }
        }
        const n = tr.asString().length;
        if (n > 0) tr.getRange(0, n).getTextStyle().setForegroundColor('#000000');
        count++;
      });
    });
    return { ok: true, message: 'Cleared all tone colors and annotations from the presentation (' + count + ' text box(es)).' };
  },

  getSelectionMarkup: function (copyMode) {
    const settings = getSettings().settings;
    copyMode = copyMode || 'both';
    const selection = SlidesApp.getActivePresentation().getSelection();
    const pageElements = selection.getPageElementRange();
    if (!pageElements) return { ok: false, message: 'Select a text box first, then click Copy Selection.' };
    const dict = getPinyinDict_();
    const colorsMap = settings.colors;
    let hanziHtml = '', hanziText = '', inlineHtml = '', inlineText = '';
    const annItems = [];
    pageElements.getPageElements().forEach(pe => {
      if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const shape = pe.asShape();
      if (!shape.getText) return;
      const str = shape.getText().asString();
      for (const ch of str) {
        if (isChinese_(ch)) {
          const reading = dict[ch];
          const parsed = reading ? parsePinyinSyllable_(reading) : { plain: '', tone: 5 };
          const toneColor = colorsMap[parsed.tone] || colorsMap[5];
          const span = settings.colorEnabled ? '<span style="color:' + toneColor + '">' + escapeHtml_(ch) + '</span>' : escapeHtml_(ch);
          hanziHtml += span; hanziText += ch;
          if (reading) {
            const ann = buildAnnotationString_(reading, parsed, settings.annotationType);
            const annStyled = settings.colorAnnotation ? '<span style="color:' + toneColor + '">' + escapeHtml_(ann) + '</span>' : escapeHtml_(ann);
            inlineHtml += span + '(' + annStyled + ')';
            inlineText += ch + '(' + ann + ')';
            annItems.push({ html: annStyled, text: ann });
          } else { inlineHtml += span; inlineText += ch; }
        } else {
          const esc = escapeHtml_(ch);
          hanziHtml += esc; hanziText += ch; inlineHtml += esc; inlineText += ch;
        }
      }
    });
    if (copyMode === 'annotation') {
      if (!annItems.length) return { ok: false, message: 'No recognized Chinese characters in the selection.' };
      return { ok: true, html: '<div>' + annItems.map(a => a.html).join(' ') + '</div>', text: annItems.map(a => a.text).join(' ') };
    }
    if (copyMode === 'chinese') return { ok: true, html: '<div>' + hanziHtml + '</div>', text: hanziText };
    return { ok: true, html: '<div>' + inlineHtml + '</div>', text: inlineText };
  }
};

/* ============================ AUTO MODE (Docs only) ============================ */

function installAutoTrigger_() {
  removeAutoTrigger_();
  ScriptApp.newTrigger('autoConvert').timeBased().everyMinutes(1).create();
}
function removeAutoTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'autoConvert') ScriptApp.deleteTrigger(t); });
}
function autoConvert() {
  const settings = getSettings().settings;
  if (!settings.autoMode) { removeAutoTrigger_(); return; }
  if (getActiveHost_() !== 'docs') return; // auto mode is Docs-only
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const dict = getPinyinDict_();
  const textElements = [];
  DocsEngine.collectTextElements_(body, textElements);
  const colorOnlySettings = Object.assign({}, settings, { annotationEnabled: false });
  textElements.forEach(t => DocsEngine.processTextElement_(t, colorOnlySettings, dict, settings.colors));
}