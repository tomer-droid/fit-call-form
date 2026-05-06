/**
 * Apps Script — קולט תגובות מהטפסים של Kamir Group:
 *   - הטופס הארוך (8 שאלות) → טאב Google Forms הקיים
 *   - הטופס הקצר (3 שאלות, סינון מהיר) → טאב נפרד "סינון מהיר (תגובות)"
 *
 * התקנה ראשונית:
 *   1. גיליון "שאלות לקראת השיחת התאמה (תגובות)"
 *   2. Extensions → Apps Script → מדביקים את הקובץ הזה (מחליפים את הקיים) → Save
 *   3. Deploy → Manage deployments → ליד הפריסה הקיימת לחיצה על אייקון העריכה (✏️) →
 *      Version: New version → Deploy. (ה-URL נשאר זהה!)
 *   4. אופציונלי: Run → testSubmit_ או testShort_ לבדיקה.
 */

// ====== הגדרות גיליונות =================================================
var LONG_SHEET_CANDIDATES = ['תגובות לטופס 1', 'תגובות הטופס 1', 'Form Responses 1', 'Form_Responses'];
var LONG_SHEET_DEFAULT    = LONG_SHEET_CANDIDATES[0];
var SHORT_SHEET_NAME      = 'סינון מהיר (תגובות)';

// ====== מיפוי שדות → כותרות ============================================
// טופס ארוך (קיים)
var LONG_FIELDS = {
  name:       'שם מלא',
  email:      'כתובת אימייל',
  experience: 'ניסיון בהשקעות',
  priority:   'מה חשוב לך במיוחד בליווי',
  why_now:    'מה גורם לך לרצות להתחיל עכשיו ולא עוד שנה?',
  focus:      'על מה תרצה שנשים דגש בשיחה שלנו?',
  budget:     'טווח השקעה מוערך להתחלה',
  success:    'אם נצא לדרך - מה ייחשב מבחינתך הצלחה בתהליך'
};

// טופס קצר (סינון מהיר)
var SHORT_FIELDS = {
  timeline: 'מתי מוכן לצאת לדרך'
};

// ====== Entry points ===================================================
function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p && p.source === 'fit-call-form-short') {
      writeRow_(getOrCreateShortSheet_(), SHORT_FIELDS, p);
    } else {
      writeRow_(getOrCreateLongSheet_(), LONG_FIELDS, p);
    }
    return jsonOut_({ ok:true });
  } catch (err) {
    Logger.log('doPost error: ' + err + '\n' + (err.stack || ''));
    return jsonOut_({ ok:false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return jsonOut_({ ok:true, service:'fit-call-form', long: LONG_SHEET_DEFAULT, short: SHORT_SHEET_NAME });
}

// ====== Sheet helpers ==================================================
function getOrCreateLongSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < LONG_SHEET_CANDIDATES.length; i++) {
    var sh = ss.getSheetByName(LONG_SHEET_CANDIDATES[i]);
    if (sh) return sh;
  }
  var nsh = ss.insertSheet(LONG_SHEET_DEFAULT);
  var hdrs = ['חותמת זמן'];
  Object.keys(LONG_FIELDS).forEach(function(k){ hdrs.push(LONG_FIELDS[k]); });
  nsh.appendRow(hdrs);
  return nsh;
}

function getOrCreateShortSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHORT_SHEET_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(SHORT_SHEET_NAME);
  var hdrs = ['חותמת זמן'];
  Object.keys(SHORT_FIELDS).forEach(function(k){ hdrs.push(SHORT_FIELDS[k]); });
  sh.appendRow(hdrs);
  return sh;
}

// ====== Generic row writer =============================================
function normalize_(s){
  return String(s || '')
    .replace(/[ ‎‏]/g, ' ')
    .replace(/[״"]/g, '"')
    .replace(/[׳']/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function writeRow_(sh, fieldMap, p) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var rawHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                     .map(function(h){ return String(h || ''); });
  var normHeaders = rawHeaders.map(normalize_);

  var fieldNorm = {};
  Object.keys(fieldMap).forEach(function(f){ fieldNorm[f] = normalize_(fieldMap[f]); });

  // נוסיף עמודות חסרות בסוף
  var missing = [];
  Object.keys(fieldMap).forEach(function(field){
    if (normHeaders.indexOf(fieldNorm[field]) === -1) missing.push(fieldMap[field]);
  });
  if (missing.length) {
    sh.getRange(1, rawHeaders.length + 1, 1, missing.length).setValues([missing]);
    rawHeaders  = rawHeaders.concat(missing);
    normHeaders = normHeaders.concat(missing.map(normalize_));
  }

  var TS_NORMS = [normalize_('חותמת זמן'), 'timestamp'];

  var row = rawHeaders.map(function(h, i){
    var hn = normHeaders[i];
    if (!hn) return '';
    if (TS_NORMS.indexOf(hn) !== -1) return formatTimestamp_(new Date());
    var match = null;
    Object.keys(fieldNorm).forEach(function(f){
      if (fieldNorm[f] === hn) match = f;
    });
    if (match && p[match] !== undefined && p[match] !== null && p[match] !== '') return p[match];
    if (p[h] !== undefined && p[h] !== null && p[h] !== '') return p[h];
    return '';
  });

  sh.appendRow(row);
}

function formatTimestamp_(d) {
  function pad(n){ return ('0' + n).slice(-2); }
  return pad(d.getDate()) + '/' + pad(d.getMonth()+1) + '/' + d.getFullYear() +
         ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== בדיקות ידניות ==================================================
function testSubmit_() {
  writeRow_(getOrCreateLongSheet_(), LONG_FIELDS, {
    name: 'בדיקת התקנה (טופס ארוך)',
    email: 'test@example.com',
    experience: 'משקיע בשוק ההון',
    priority: 'שקיפות מלאה',
    why_now: 'בדיקה ידנית מ-Apps Script',
    focus: 'לוודא שהשורה מתווספת לעמודות הנכונות',
    budget: 'עוד לא החלטתי / צריך להבין את זה בשיחה',
    success: 'שורה אחת מופיעה בגיליון בעמודות הנכונות'
  });
  Logger.log('Long test row appended.');
}

function testShort_() {
  writeRow_(getOrCreateShortSheet_(), SHORT_FIELDS, {
    timeline: 'ב-3 החודשים הקרובים'
  });
  Logger.log('Short test row appended.');
}
