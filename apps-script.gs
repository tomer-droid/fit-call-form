/**
 * Apps Script — קולט תגובות מ-fit-call-form ומוסיף שורה לאותו גיליון
 * שכבר מחובר ל-Google Form הקיים, באותן עמודות בדיוק.
 *
 * התקנה:
 *   1. פותחים את הגיליון "שאלות לקראת השיחת התאמה (תגובות)" (id 1cwKBNFSfAvVgtZR0Wh0_-q9VyDnM_l6fQxssoAWjaeM)
 *   2. Extensions → Apps Script
 *   3. מדביקים את כל הקובץ הזה כ-Code.gs (מוחקים את ברירת המחדל)
 *   4. Deploy → New deployment → Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *   5. מעתיקים את ה-URL שמתקבל ושותלים אותו ב-index.html במשתנה SCRIPT_URL
 *   6. אופציונלי: מריצים מהעורך את testSubmit_() פעם אחת כדי לאשר הרשאות
 *      ולוודא שמתווספת שורת בדיקה לגיליון.
 */

// שמות אפשריים של הטאב שאליו Google Forms כותב. נחפש לפי הסדר.
var SHEET_NAME_CANDIDATES = ['תגובות לטופס 1', 'תגובות הטופס 1', 'Form Responses 1', 'Form_Responses'];
var SHEET_NAME = SHEET_NAME_CANDIDATES[0]; // ברירת מחדל ליצירה אם אף אחד לא נמצא

// מיפוי field-id (מה שה-frontend שולח) → שם הכותרת בעברית בגיליון.
// נשען על הכותרות הקיימות שראינו: חותמת זמן, שם מלא, כתובת אימייל, ניסיון בהשקעות,
// מה חשוב לך במיוחד בליווי, מה גורם לך לרצות להתחיל עכשיו ולא עוד שנה?,
// על מה תרצה שנשים דגש בשיחה שלנו?, טווח השקעה מוערך להתחלה, אם נצא לדרך - מה ייחשב מבחינתך הצלחה בתהליך.
var FIELD_TO_HEADER = {
  name:       'שם מלא',
  email:      'כתובת אימייל',
  experience: 'ניסיון בהשקעות',
  priority:   'מה חשוב לך במיוחד בליווי',
  why_now:    'מה גורם לך לרצות להתחיל עכשיו ולא עוד שנה?',
  focus:      'על מה תרצה שנשים דגש בשיחה שלנו?',
  budget:     'טווח השקעה מוערך להתחלה',
  success:    'אם נצא לדרך - מה ייחשב מבחינתך הצלחה בתהליך'
};

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    appendRow_(p);
    return jsonOut_({ ok:true });
  } catch (err) {
    Logger.log('doPost error: ' + err + '\n' + (err.stack || ''));
    return jsonOut_({ ok:false, error: String(err && err.message || err) });
  }
}

// GET עוזר לבדוק במהירות שה-deploy חי (פתח את ה-URL בדפדפן).
function doGet() {
  return jsonOut_({ ok:true, service:'fit-call-form', sheet: SHEET_NAME });
}

function normalize_(s){
  // מנרמל רווחים/טאבים/שורות חדשות וגרשיים שונים כדי שכותרות יותאמו גם אם יש שינויים זעירים
  return String(s || '')
    .replace(/[ ‎‏]/g, ' ')   // nbsp + RTL/LTR marks
    .replace(/[״"]/g, '"')
    .replace(/[׳']/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function appendRow_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = null;
  for (var i = 0; i < SHEET_NAME_CANDIDATES.length; i++) {
    sh = ss.getSheetByName(SHEET_NAME_CANDIDATES[i]);
    if (sh) break;
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var hdrs = ['חותמת זמן'];
    Object.keys(FIELD_TO_HEADER).forEach(function(k){ hdrs.push(FIELD_TO_HEADER[k]); });
    sh.appendRow(hdrs);
  }

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var rawHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0]
                     .map(function(h){ return String(h || ''); });
  var normHeaders = rawHeaders.map(normalize_);

  // מיפוי: הכותרת המנורמלת של כל field
  var fieldNorm = {};
  Object.keys(FIELD_TO_HEADER).forEach(function(f){ fieldNorm[f] = normalize_(FIELD_TO_HEADER[f]); });

  // נוסיף עמודות חסרות בסוף (לפי כותרת מנורמלת)
  var missing = [];
  Object.keys(FIELD_TO_HEADER).forEach(function(field){
    if (normHeaders.indexOf(fieldNorm[field]) === -1) missing.push(FIELD_TO_HEADER[field]);
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
    // התאמה לפי field
    var match = null;
    Object.keys(fieldNorm).forEach(function(f){
      if (fieldNorm[f] === hn) match = f;
    });
    if (match && p[match] !== undefined && p[match] !== null && p[match] !== '') return p[match];
    // גמיש: גם payload שמכיל את שם הכותרת ישירות
    if (p[h] !== undefined && p[h] !== null && p[h] !== '') return p[h];
    return '';
  });

  sh.appendRow(row);
}

function formatTimestamp_(d) {
  function pad(n){ return ('0' + n).slice(-2); }
  // פורמט תואם לקיים בגיליון: dd/MM/yyyy HH:mm:ss
  return pad(d.getDate()) + '/' + pad(d.getMonth()+1) + '/' + d.getFullYear() +
         ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** הרצה ידנית מהעורך כדי לאמת הרשאות + שורת בדיקה. */
function testSubmit_() {
  appendRow_({
    name: 'בדיקת התקנה',
    email: 'test@example.com',
    experience: 'משקיע בשוק ההון',
    priority: 'שקיפות מלאה',
    why_now: 'בדיקה ידנית מ-Apps Script',
    focus: 'לוודא שהשורה מתווספת לעמודות הנכונות',
    budget: 'עוד לא החלטתי / צריך להבין את זה בשיחה',
    success: 'שורה אחת מופיעה בגיליון בעמודות הנכונות'
  });
  Logger.log('Test row appended.');
}
