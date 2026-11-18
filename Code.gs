// סקר לקוחות שקד – קליטת תשובות לגיליון "תשובות סקר שקד"
const SHEET_ID = '10L15Yy41rdujyGkvT-K0VY7vY8NuxXN4iDVv8zYCQ8M';

function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  const r = d.rate || {};
  // מונע הרצת נוסחאות שהוקלדו בטופס
  const clean = v => {
    const s = Array.isArray(v) ? v.join(', ') : String(v == null ? '' : v);
    return s.slice(0, 1500).replace(/^[=+\-@]/, "'$&");
  };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sheet.appendRow([
    new Date(),
    clean(d.name),
    clean(d.sat),
    clean(d.good),
    clean(d.good_other),
    r['מחיר'] || '',
    r['שירות'] || '',
    r['איכות'] || '',
    r['זמני אספקה'] || '',
    r['מגוון'] || '',
    clean(d.value),
    clean(d.improve),
    clean(d.cats),
    clean(d.top),
    clean(d.notes)
  ]);
  return ContentService.createTextOutput('ok');
}
