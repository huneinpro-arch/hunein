// ===== طلبات حنين → Google Sheet (order id) =====
// السكربت يقرأ عناوين الصف الأول في جدولك ويضع كل قيمة تحت عنوانها، فيبقى كل شيء في مكانه
// حتى لو غيّرت ترتيب الأعمدة أو أسماءها (الأسماء المعروفة مذكورة في fields_ بالأسفل).
var SHEET_ID = '';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    if (!SHEET_ID) throw new Error('افتح الكود من داخل الجدول: الإضافات > Apps Script');
    ss = SpreadsheetApp.openById(SHEET_ID);
  }
  return ss.getSheets()[0];
}

// يوحّد اسم العمود: بدون مسافات أو _ أو - وبحروف صغيرة
function norm_(h) {
  return String(h).toLowerCase().replace(/[\s_\-]/g, '');
}

// كل اسم عمود (موحّد) ← قيمته. أضف اسماً بديلاً هنا إن غيّرت اسم عمود في الجدول.
function fields_(p, date, hasOffer) {
  var price = Number(p.price) || 0, base = Number(p.basePrice) || 0;
  var f = {
    'date': date, 'التاريخ': date,
    'orderid': p.orderId || '', 'رقمالطلب': p.orderId || '',
    'name': p.fullName || '', 'الاسم': p.fullName || '',
    'phone': (p.phone || '') + (p.phone2 ? ' / ' + p.phone2 : ''), 'الهاتف': (p.phone || '') + (p.phone2 ? ' / ' + p.phone2 : ''),
    'wilaya': p.wilaya || '', 'الولاية': p.wilaya || '',
    'baladia': p.commune || '', 'commune': p.commune || '', 'البلدية': p.commune || '',
    'delivery': p.delivery || '', 'نوعالتوصيل': p.delivery || '',
    'socitedelivraison': p.deliveryCompany || '', 'societedelivraison': p.deliveryCompany || '', 'sociétédelivraison': p.deliveryCompany || '',
    'company': p.deliveryCompany || '', 'شركةالتوصيل': p.deliveryCompany || '',
    'product': p.product || '', 'الدورة': p.product || '', 'المنتج': p.product || '',
    'items': p.items || '', 'الدورات': p.items || '',
    'qty': Number(p.qty) || 1,
    'baseprice': hasOffer ? base : price, 'السعر': hasOffer ? base : price,
    'offerprice': price,
    'shippingcost': Number(p.shipping) || 0, 'shipping': Number(p.shipping) || 0, 'سعرالتوصيل': Number(p.shipping) || 0,
    'totalprix': Number(p.total) || 0, 'total': Number(p.total) || 0, 'المجموع': Number(p.total) || 0,
    'status': 'طلب جديد', 'الحالة': 'طلب جديد',
    'طريقةالدفع': p.paymentMethod || 'الدفع عند الاستلام', 'paymentmethod': p.paymentMethod || 'الدفع عند الاستلام',
    'المدفوعالآن': Number(p.paidNow) || 0, 'paidnow': Number(p.paidNow) || 0,
    'الوصل': p.receipt || '', 'receipt': p.receipt || ''
  };
  return f;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    // يتجاهل أي طلب فارغ (بدون اسم ولا هاتف)
    if (!p.fullName && !p.phone) return ContentService.createTextOutput('empty');

    var sh = getSheet_();
    var cols = sh.getLastColumn();
    if (sh.getLastRow() === 0 || cols === 0) throw new Error('اكتب عناوين الأعمدة في الصف الأول من الجدول');

    var heads = sh.getRange(1, 1, 1, cols).getValues()[0].map(norm_);
    // عمود بلا عنوان: بعد product يأخذ أسماء الدورات، وبعد total Prix يأخذ الحالة
    for (var k = 1; k < heads.length; k++) {
      if (heads[k] === '' && (heads[k - 1] === 'product' || heads[k - 1] === 'الدورة')) heads[k] = 'items';
      if (heads[k] === '' && (heads[k - 1] === 'totalprix' || heads[k - 1] === 'total' || heads[k - 1] === 'المجموع')) heads[k] = 'status';
    }
    var hasOffer = heads.indexOf('offerprice') > -1;
    var date = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd HH:mm');
    var f = fields_(p, date, hasOffer);

    var row = sh.getLastRow() + 1;

    // ينسخ تنسيق الصف السابق ليبقى الجدول متناسقاً
    if (row > 2) {
      sh.getRange(row - 1, 1, 1, cols).copyTo(sh.getRange(row, 1, 1, cols), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    var phoneCol = heads.indexOf('phone') > -1 ? heads.indexOf('phone') : heads.indexOf('الهاتف');
    if (phoneCol > -1) sh.getRange(row, phoneCol + 1).setNumberFormat('@');

    var vals = heads.map(function (h) { return f.hasOwnProperty(h) ? f[h] : ''; });
    sh.getRange(row, 1, 1, cols).setValues([vals]);

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
    deliveryCompany: 'DHD', product: 'حقيبة المهندس برو', items: 'Robot + AutoCAD', qty: '3',
    basePrice: '15000', price: '7500', shipping: '700', total: '8200',
    paymentMethod: 'الدفع عند الاستلام', paidNow: '0', receipt: ''
  }});
}
