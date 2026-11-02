/**
 * اختياري — Google Apps Script يستقبل الطلبات من الصفحة ويكتبها في الشيت حسب أسماء الأعمدة (الصف الأول).
 * الحقول الجديدة (paymentMethod, paidNow, dueOnDelivery, receipt) تُضاف كأعمدة تلقائياً إن لم تكن موجودة.
 * إن كان عندك سكربت يعمل حالياً بترتيب أعمدة ثابت، أضف له هذه الحقول الأربعة أو استبدله بهذا ثم: نشر ← نشر جديد.
 * الوصولات (الصور) لا تُحفظ هنا: تصل إلى واتساب عبر api/receipt.js
 */
function doPost(e) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var p = e.parameter;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
  Object.keys(p).forEach(function (k) { if (head.indexOf(k) < 0) head.push(k); });
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  var row = head.map(function (k) { return p[k] == null ? '' : p[k]; });
  sh.appendRow(row);
  return ContentService.createTextOutput('ok');
}
