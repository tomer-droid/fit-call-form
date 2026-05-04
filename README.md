# שאלות לקראת שיחת התאמה

טופס סטטי בעברית RTL, ממותג Kamir Group, שמחליף את ה-Google Form הקיים.
התשובות נשמרות ב**אותו** Google Sheet שכבר מחובר לטופס המקורי.

## מבנה
- `index.html` — האתר עצמו (wizard 8 שאלות, localStorage, Tailwind CDN לא בשימוש — CSS מקומי)
- `apps-script.gs` — קוד שמדביקים ב-Sheet → Extensions → Apps Script
- `assets/kamir-logo.png` — לוגו
- `netlify.toml` — הגדרות deploy

## חיבור ל-Sheet (פעם אחת)
1. פותחים את הגיליון `שאלות לקראת השיחת התאמה (תגובות)` (id `1cwKBNFSfAvVgtZR0Wh0_-q9VyDnM_l6fQxssoAWjaeM`)
2. Extensions → Apps Script → מדביקים את `apps-script.gs` (במקום ברירת המחדל)
3. Save → Run `testSubmit_` פעם אחת ומאשרים הרשאות → בודקים שהופיעה שורת "בדיקת התקנה" בגיליון
4. Deploy → New deployment → Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. מעתיקים את ה-URL ושותלים אותו ב-`index.html`:
   ```js
   const SCRIPT_URL = 'https://script.google.com/macros/s/.../exec';
   ```

## Deploy ל-Netlify
- אופציה הכי מהירה: גוררים את התיקייה כולה ל-https://app.netlify.com/drop
- אופציה עם git: `git init && git add . && git commit -m "init" && netlify deploy --prod`

## בדיקה מקומית
פותחים את `index.html` ישירות בדפדפן. לפני ש-`SCRIPT_URL` הוגדר — שליחה תרשום את ה-payload ל-console במקום לשלוח. localStorage שומר התקדמות בין refresh-ים.
