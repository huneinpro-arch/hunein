// ===== طلبات حنين → Google Sheet (order_id) =====
var SHEET_ID = '';

// ترتيب الأعمدة كما في جدولك: date, order_id, name, phone, wilaya, baladia, address, product, qty, base_price, offer_price
// ثم الأعمدة التي بعدها (غيّر أسماءها وترتيبها هنا بما يوافق جدولك)
var HEADERS = ['date','order_id','name','phone','wilaya','baladia','address','product','qty','base_price','offer_price',
  'shipping','total','delivery','company','status','payment_method','paid_now'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    if (!SHEET_ID) throw new Error('افتح الكود من داخل الجدول: الإضافات > Apps Script');
    ss = SpreadsheetApp.openById(SHEET_ID);
  }
  return ss.getSheets()[0];
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    // يتجاهل أي طلب فارغ (بدون اسم ولا هاتف)
    if (!p.fullName && !p.phone) return ContentService.createTextOutput('empty');

    var sh = getSheet_();

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    }

    var row = sh.getLastRow() + 1;
    var date = Utilities.formatDate(new Date(), 'Africa/Algiers', 'dd/MM/yyyy');

    // ينسخ تنسيق الصف السابق ليبقى الجدول متناسقاً
    if (row > 2) {
      sh.getRange(row - 1, 1, 1, HEADERS.length)
        .copyTo(sh.getRange(row, 1, 1, HEADERS.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    sh.getRange(row, 4).setNumberFormat('@');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      date,
      p.orderId || '',
      p.fullName || '',
      (p.phone || '') + (p.phone2 ? ' / ' + p.phone2 : ''),
      p.wilaya || '', p.commune || '', '',
      p.product || '',
      Number(p.qty) || 1,
      Number(p.basePrice) || 0, Number(p.price) || 0,
      Number(p.shipping) || 0, Number(p.total) || 0,
      p.delivery || '', p.deliveryCompany || '',
      'طلب جديد',
      p.paymentMethod || 'الدفع عند الاستلام', Number(p.paidNow) || 0
    ]]);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('الخدمة تعمل');
}

function test() {
  doPost({ parameter: {
    orderId: 'HN-TEST', fullName: 'طلب تجريبي', phone: '0550123456', phone2: '',
    wilaya: 'قسنطينة', commune: 'الخروب', delivery: 'للمنزل',
    deliveryCompany: 'DHD', product: 'حقيبة المهندس برو', qty: '3',
    basePrice: '15000', price: '7500', shipping: '700', total: '8200',
    paymentMethod: 'الدفع عند الاستلام', paidNow: '0'
  }});
}
