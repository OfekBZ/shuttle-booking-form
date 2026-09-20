/**
 * Code.gs
 * -----------------------------------------------------------
 * סקריפט Google Apps Script לקבלת נתוני "הזמנת הסעה"
 * מהאתר (index.html) ורישומם כשורה חדשה בגיליון Google Sheets.
 *
 * התקנה:
 * 1. פתחו את גיליון ה-Google Sheets שאליו תרצו לרשום את ההזמנות.
 * 2. תפריט: תוספים (Extensions) > Apps Script.
 * 3. מחקו כל קוד קיים בקובץ Code.gs, והדביקו את כל הקוד הזה במקומו.
 * 4. שמרו את הפרויקט (בשם כלשהו, למשל "טופס הסעות").
 * 5. עברו לשלב הפריסה (ראו הוראות מלאות בסוף הקובץ).
 * -----------------------------------------------------------
 */

// שם הגיליון (Sheet/Tab) שבו יירשמו ההזמנות. שנו אם הטאב נקרא אחרת.
const SHEET_NAME = 'הזמנות';

// כותרות העמודות בגיליון, בהתאמה לשדות הטופס
const HEADERS = [
  'חותמת זמן',
  'שם הנוסע',
  'טלפון',
  'תאריך נסיעה',
  'שעת איסוף',
  'מוצא',
  'יעד',
  'מספר נוסעים',
  'שעת חזור',
  'הערות'
];

/**
 * מטפל בבקשות POST שמגיעות מהטופס באתר.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('לא התקבל תוכן בבקשה (postData חסר)');
    }

    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date(),                    // חותמת זמן
      data.passengerName || '',
      data.phone || '',
      data.date || '',
      data.pickupTime || '',
      data.origin || '',
      data.destination || '',
      data.passengers || '',
      data.returnTime || '',
      data.notes || ''
    ]);

    return jsonResponse({ result: 'success' });

  } catch (err) {
    return jsonResponse({ result: 'error', error: err.message || String(err) });
  }
}

/**
 * מאתר את הגיליון היעד, ויוצר אותו עם כותרות אם הוא לא קיים עדיין.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * עוזר קטן להחזרת תשובת JSON תקינה.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * פונקציית בדיקה אופציונלית - הריצו אותה מתוך עורך ה-Apps Script
 * (לחצן ▶ Run) כדי לבדוק שהרישום לגיליון עובד, בלי צורך בטופס.
 */
function testAppendRow() {
  const sheet = getOrCreateSheet();
  sheet.appendRow([
    new Date(),
    'בדיקה - ישראל ישראלי',
    '050-1234567',
    '2026-10-01',
    '08:30',
    'תל אביב',
    'ירושלים',
    3,
    '18:00',
    'הזמנת בדיקה'
  ]);
}

/*
=====================================================================
 הוראות פריסה (Deploy) כ-Web App
=====================================================================

1. בעורך ה-Apps Script, לחצו על הכפתור הכחול "Deploy" (פרוס)
   שבפינה הימנית העליונה, ואז על "New deployment" (פריסה חדשה).

2. לצד "Select type" (בחירת סוג), לחצו על סמל גלגל השיניים ⚙️
   ובחרו באפשרות "Web app".

3. הגדירו את הפרטים הבאים:
     - Description:      כל תיאור שתרצו, למשל "טופס הזמנת הסעה"
     - Execute as:        Me (החשבון שלכם)
     - Who has access:    Anyone   <-- חשוב מאוד לבחור באפשרות הזו

4. לחצו על "Deploy" (פרוס).

5. ייתכן שתתבקשו לאשר הרשאות (Authorize access):
     - בחרו את חשבון ה-Google שלכם.
     - אם תופיע אזהרה "Google hasn't verified this app",
       לחצו על "Advanced" ואז על "Go to [שם הפרויקט] (unsafe)".
       זהו אזהרה סטנדרטית לסקריפטים אישיים - הסקריפט הוא שלכם ובטוח.
     - אשרו את ההרשאות המבוקשות.

6. בסיום, תופיע כתובת "Web app URL" - זו הכתובת שאתם צריכים.
   העתיקו אותה.

7. פתחו את קובץ index.html, ומצאו את השורה:
       const GOOGLE_SCRIPT_URL = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
   הדביקו את הכתובת שהעתקתם בין הגרשיים, כך שתיראה בערך כך:
       const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycb.../exec";

8. שמרו את הקובץ. הטופס מוכן לשליחת הזמנות אל הגיליון.

הערה חשובה - עדכון עתידי של הקוד:
אם בעתיד תשנו את הקוד ב-Code.gs, יצירת גרסה חדשה של הפריסה חייבת
להתבצע דרך "Deploy" > "Manage deployments" > עריכה (סמל העיפרון) >
"Version: New version" > "Deploy". פריסה חדשה לגמרי (New deployment)
תיצור כתובת URL אחרת, ותצטרכו לעדכן אותה גם בקובץ index.html.
=====================================================================
*/
