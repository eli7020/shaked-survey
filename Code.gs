// סקר לקוחות שקד – קליטת תשובות לגיליון "תשובות סקר שקד"
const SHEET_ID = '10L15Yy41rdujyGkvT-K0VY7vY8NuxXN4iDVv8zYCQ8M';

function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  const r = d.rate || {};
  const clean = v => {
    const s = Array.isArray(v) ? v.join(', ') : String(v == null ? '' : v);
    return s.slice(0, 1500).replace(/^[=+\-@]/, "'$&");
  };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getRange(1, 16).getValue() === '') sheet.getRange(1, 16).setValue('סוג העסק');
  const biz = d.biz === 'אחר' ? ('אחר: ' + (d.biz_other || '')) : d.biz;
  const name = String(d.name || '').replace(/\s*\([^)]*\)\s*$/, '');
  sheet.appendRow([
    new Date(), clean(name), clean(d.sat), clean(d.good), clean(d.good_other),
    r['מחיר'] || '', r['שירות'] || '', r['איכות'] || '', r['זמני אספקה'] || '', r['מגוון'] || '',
    clean(d.value), clean(d.improve), clean(d.cats), clean(d.top), clean(d.notes),
    clean(biz)
  ]);
  return ContentService.createTextOutput('ok');
}
