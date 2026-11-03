// ===== طلبات حنين → Google Sheet (orderId) =====
var SHEET_ID = '';

// ترتيب جدولك الحالي (A–M) كما هو، وعمودان جديدان بعد الحالة: N طريقة الدفع، O المدفوع الآن، P رقم الطلب
var HEADERS = ['التاريخ','الاسم','الهاتف','الولاية','البلدية',
  'نوع التوصيل','شركة التوصيل','الدورة','الدورات','السعر','سعر التوصيل','المجموع','الحالة',
  'طريقة الدفع','المدفوع الآن','رقم الطلب'];

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

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    // يتجاهل أي طلب فارغ (بدون اسم ولا هاتف) حتى لا تظهر صفوف فارغة
    if (!p.fullName && !p.phone) return ContentService.createTextOutput('empty');

    var sh = getSheet_();

    removeOrderIdColumn_(sh);

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    }

    var row = sh.getLastRow() + 1;
    var date = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd HH:mm');

    // ينسخ تنسيق وقائمة "الحالة" من الصف السابق ليبقى الجدول متناسقاً
    if (row > 2) {
      var prev = sh.getRange(row - 1, 1, 1, HEADERS.length);
      prev.copyTo(sh.getRange(row, 1, 1, HEADERS.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      prev.copyTo(sh.getRange(row, 1, 1, HEADERS.length), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    sh.getRange(row, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      date,
      p.fullName || '',
      (p.phone || '') + (p.phone2 ? ' / ' + p.phone2 : ''),
      p.wilaya || '', p.commune || '',
      p.delivery || '', p.deliveryCompany || '',
      p.product || '', p.items || '',
      Number(p.price) || 0, Number(p.shipping) || 0, Number(p.total) || 0,
      'طلب جديد',
      p.paymentMethod || 'الدفع عند الاستلام', Number(p.paidNow) || 0,
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

// شغّل هذه الدالة مرة واحدة: تنظّف عمود "orderId" القديم وتكتب عناوين الأعمدة الجديدة (N و O و P)
function fixSheet() {
  var sh = getSheet_();
  removeOrderIdColumn_(sh);
  if (sh.getLastRow() > 0) sh.getRange(1, 14, 1, 3).setValues([[HEADERS[13], HEADERS[14], HEADERS[15]]]);
}

// اختياري، مرة واحدة: يحذف الصفوف الفارغة (بدون اسم ولا هاتف) التي ظهرت سابقاً
function deleteEmptyRows() {
  var sh = getSheet_();
  for (var r = sh.getLastRow(); r >= 2; r--) {
    var v = sh.getRange(r, 2, 1, 2).getValues()[0];
    if (!v[0] && !v[1]) sh.deleteRow(r);
  }
}

function test() {
  doPost({ parameter: {
    fullName: 'طلب تجريبي', phone: '0550123456', phone2: '',
    wilaya: 'قسنطينة', commune: 'الخروب', delivery: 'للمنزل',
    deliveryCompany: 'DHD', product: 'حقيبة المهندس برو', items: 'Robot + AutoCAD',
    price: '7500', shipping: '700', total: '8200',
    paymentMethod: 'الدفع عند الاستلام', paidNow: '0', orderId: 'HN-TEST'
  }});
}
