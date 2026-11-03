// ===== طلبات حنين → Google Sheet (orderId) =====
var SHEET_ID = '';

// الأعمدة القديمة كما هي (1-13). الأعمدة الجديدة أُضيفت في النهاية حتى لا يتغير ترتيب جدولك الحالي.
var HEADERS = ['التاريخ','الاسم','الهاتف','هاتف ثاني','الولاية','البلدية',
  'نوع التوصيل','شركة التوصيل','الدورة','السعر','سعر التوصيل','المجموع','الحالة',
  'طريقة الدفع','المدفوع الآن','المتبقي عند الاستلام','الوصل','رقم الطلب'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    if (!SHEET_ID) throw new Error('افتح الكود من داخل الجدول: الإضافات > Apps Script');
    ss = SpreadsheetApp.openById(SHEET_ID);
  }
  return ss.getSheets()[0];
}

// يحذف عمود "orderId" القديم تلقائياً إن كان موجوداً
function removeOrderIdColumn_(sh) {
  var last = sh.getLastColumn();
  if (sh.getLastRow() === 0 || last === 0) return;
  var hdr = sh.getRange(1, 1, 1, last).getValues()[0];
  var i = hdr.indexOf('order_id');
  if (i > -1) sh.deleteColumn(i + 1);
}

// يكتب عناوين الأعمدة الجديدة (14-18) إن لم تكن موجودة في جدول قديم
function ensureHeaders_(sh) {
  var n = HEADERS.length;
  var cur = sh.getRange(1, 1, 1, n).getValues()[0];
  for (var i = 13; i < n; i++) {
    if (!cur[i]) sh.getRange(1, i + 1).setValue(HEADERS[i]);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    var sh = getSheet_();

    removeOrderIdColumn_(sh);

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    } else {
      ensureHeaders_(sh);
    }

    var row = sh.getLastRow() + 1;
    var date = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd HH:mm');

    sh.getRange(row, 3, 1, 2).setNumberFormat('@');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      date,
      p.fullName || '', p.phone || '', p.phone2 || '',
      p.wilaya || '', p.commune || '',
      p.delivery || '', p.deliveryCompany || '',
      p.product || '',
      Number(p.price) || 0, Number(p.shipping) || 0, Number(p.total) || 0,
      'جديد',
      p.paymentMethod || 'الدفع عند الاستلام',
      Number(p.paidNow) || 0, Number(p.dueOnDelivery) || 0,
      p.receipt || '',
      p.orderId || ''
    ]]);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('الخدمة تعمل');
}

// شغّل هذه الدالة مرة واحدة: تنظّف الجدول القديم من عمود "orderId" فوراً
function fixSheet() {
  removeOrderIdColumn_(getSheet_());
  ensureHeaders_(getSheet_());
}

function test() {
  doPost({ parameter: {
    fullName: 'طلب تجريبي', phone: '0550123456', phone2: '',
    wilaya: 'قسنطينة', commune: 'الخروب', delivery: 'للمنزل',
    deliveryCompany: 'DHD', product: 'حقيبة المهندس برو',
    price: '0', shipping: '700', total: '700',
    paymentMethod: 'بريدي موب / CCP', paidNow: '7500', dueOnDelivery: '700',
    receipt: 'أُرسل عبر واتساب', orderId: 'HN-TEST'
  }});
}
