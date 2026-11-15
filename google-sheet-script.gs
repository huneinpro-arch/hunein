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

// يضيف عمودي name و phone بعد order id إن لم يكونا موجودين في الجدول (مرة واحدة فقط)
function ensureNamePhone_(sh) {
  var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  var at = heads.indexOf('orderid');
  if (at < 0) at = heads.indexOf('رقمالطلب');
  if (at < 0) at = 0;
  if (heads.indexOf('name') < 0 && heads.indexOf('الاسم') < 0) {
    sh.insertColumnAfter(at + 1);
    sh.getRange(1, at + 2).setValue('name');
    sh.getRange(1, at + 2).setBackground(sh.getRange(1, at + 1).getBackground());
    at++;
  } else {
    at = Math.max(heads.indexOf('name'), heads.indexOf('الاسم'));
  }
  if (heads.indexOf('phone') < 0 && heads.indexOf('الهاتف') < 0) {
    sh.insertColumnAfter(at + 1);
    sh.getRange(1, at + 2).setValue('phone');
    sh.getRange(1, at + 2).setBackground(sh.getRange(1, at + 1).getBackground());
  }
}

// ===== التحقق في الخادم (لا نثق بما يرسله المتصفح) =====
var PHONE_RE_ = /^0[567]\d{8}$/;
var ORDER_RE_ = /^HN-\d{6}-[A-Z2-9]{4}$/;
var DELIVERY_OK_ = ['للمنزل', 'لمكتب التوصيل'];
var PAYMENT_OK_ = ['الدفع عند الاستلام', 'بريدي موب / CCP'];
var RISKY_RE_ = /^[=+\-@\t\r]/; // نص يبدأ بهذه الرموز قد يُفسَّر كصيغة في الجدول

// يزيل محارف التحكم ويقصّ الطول
function txt_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

// رقم بين 0 وحدّ أعلى معقول، وإلا 0
function num_(v, max) {
  var n = Number(v);
  return isFinite(n) && n >= 0 && n <= max ? n : 0;
}

// يعيد نسخة منظّفة من الطلب، أو null إن كان غير صالح
function validate_(raw) {
  var p = {
    orderId: txt_(raw.orderId, 20),
    fullName: txt_(raw.fullName, 80),
    phone: txt_(raw.phone, 10),
    phone2: txt_(raw.phone2, 10),
    wilaya: txt_(raw.wilaya, 40),
    commune: txt_(raw.commune, 60),
    delivery: txt_(raw.delivery, 30),
    deliveryCompany: txt_(raw.deliveryCompany, 30),
    product: txt_(raw.product, 120),
    items: txt_(raw.items, 250),
    qty: Math.min(Math.max(Math.floor(num_(raw.qty, 20)), 1), 20),
    basePrice: num_(raw.basePrice, 1000000),
    price: num_(raw.price, 1000000),
    shipping: num_(raw.shipping, 20000),
    total: num_(raw.total, 1000000),
    paidNow: num_(raw.paidNow, 1000000),
    paymentMethod: txt_(raw.paymentMethod, 30) || 'الدفع عند الاستلام',
    receipt: txt_(raw.receipt, 40)
  };
  if (!ORDER_RE_.test(p.orderId)) return null;
  if (p.fullName.length < 3) return null;
  if (!PHONE_RE_.test(p.phone)) return null;
  if (p.phone2 && !PHONE_RE_.test(p.phone2)) return null;
  if (!p.wilaya || !p.commune) return null;
  if (DELIVERY_OK_.indexOf(p.delivery) < 0) return null;
  if (PAYMENT_OK_.indexOf(p.paymentMethod) < 0) return null;
  return p;
}

var RC_MIMES_ = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
function attachReceipt_(raw) {
  var oid = txt_(raw.orderId, 20);
  if (!ORDER_RE_.test(oid) || !raw.rcData || RC_MIMES_.indexOf(raw.rcMime) < 0) return 'invalid';
  var it = DriveApp.getFoldersByName('Hunein Receipts');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Hunein Receipts');
  var ext = raw.rcMime === 'application/pdf' ? 'pdf' : raw.rcMime.split('/')[1];
  var url = folder.createFile(Utilities.newBlob(Utilities.base64Decode(raw.rcData), raw.rcMime, oid + '.' + ext)).getUrl();
  var sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return 'norow';
  var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  var oc = Math.max(heads.indexOf('orderid'), heads.indexOf('رقمالطلب'));
  var rc = Math.max(heads.indexOf('الوصل'), heads.indexOf('receipt'));
  if (oc < 0 || rc < 0) return 'nocol';
  var ids = sh.getRange(2, oc + 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === oid) { sh.getRange(i + 2, rc + 1).setValue(url); return 'ok'; }
  }
  return 'norow';
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var raw = (e && e.parameter) || {};
    // رفع وصل الدفع لطلب مسجَّل: يُحفظ في Drive ويوضع رابطه في صف الطلب
    if (raw.rcOnly) return ContentService.createTextOutput(attachReceipt_(raw));
    // يتجاهل أي طلب فارغ (بدون اسم ولا هاتف)
    if (!raw.fullName && !raw.phone) return ContentService.createTextOutput('empty');
    // يرفض أي طلب غير صالح قبل أن يلمس الجدول
    var p = validate_(raw);
    if (!p) return ContentService.createTextOutput('invalid');
    // يمنع التكرار: الموقع قد يعيد إرسال الطلب نفسه عند ضعف الشبكة (نفس رقم الطلب يُسجَّل مرة واحدة)
    var cache = CacheService.getScriptCache();
    if (cache.get('o_' + p.orderId)) return ContentService.createTextOutput('dup');

    var sh = getSheet_();
    if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) throw new Error('اكتب عناوين الأعمدة في الصف الأول من الجدول');
    ensureNamePhone_(sh);
    var cols = sh.getLastColumn();

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
    // أي خلية نصية تبدأ برمز صيغة (= + - @) تُحفظ كنص عادي حتى لا تُنفَّذ كصيغة
    for (var c = 0; c < vals.length; c++) {
      if (typeof vals[c] === 'string' && RISKY_RE_.test(vals[c])) {
        sh.getRange(row, c + 1).setNumberFormat('@');
      }
    }
    sh.getRange(row, 1, 1, cols).setValues([vals]);
    cache.put('o_' + p.orderId, '1', 21600);

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
    orderId: 'HN-261114-AB23', fullName: 'طلب تجريبي', phone: '0550123456', phone2: '',
    wilaya: 'قسنطينة', commune: 'الخروب', delivery: 'للمنزل',
    deliveryCompany: 'DHD', product: 'حقيبة المهندس برو', items: 'Robot + AutoCAD', qty: '3',
    basePrice: '15000', price: '7500', shipping: '700', total: '8200',
    paymentMethod: 'الدفع عند الاستلام', paidNow: '0', receipt: ''
  }});
}
