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
  var smsJob = null;
  try {
    var raw = (e && e.parameter) || {};
    // ردّ الزبون على الرسالة (webhook من تطبيق البوابة): 1 = تأكيد، 0 = طلب إلغاء
    var body = null;
    if (e && e.postData && /json/i.test(e.postData.type || '')) {
      try { body = JSON.parse(e.postData.contents); } catch (x) { body = null; }
    }
    if (body && body.event === 'sms:received') return ContentService.createTextOutput(handleSmsReply_(body, raw.k));
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
    smsJob = { phone: p.phone, orderId: p.orderId, ccp: p.paymentMethod === PAYMENT_OK_[1] };
  } finally {
    lock.releaseLock();
  }
  // يُرسَل بعد تحرير القفل كي لا يبطّئ باقي الطلبات، وفشله لا يؤثر على تسجيل الطلب
  if (smsJob) { try { sendOrderSms_(smsJob); } catch (err) { Logger.log('sms error: ' + err); } }
  return ContentService.createTextOutput('ok');
}

// ===== رسائل SMS التلقائية عبر تطبيق "SMS Gateway for Android" (sms-gate.app) =====
// الإعدادات تُوضع في: إعدادات المشروع (⚙) ← خصائص النص (Script properties). لا تكتبها داخل الكود:
//   SMS_ENABLED = 1          (أي قيمة أخرى توقف الإرسال)
//   SMS_USER / SMS_PASS      بيانات الدخول الظاهرة في تطبيق البوابة (وضع Cloud server)
//   WEBHOOK_KEY              كلمة سرّ عشوائية طويلة لحماية استقبال الردود
//   SMS_TEST_PHONE           رقمك (بصيغة 0550123456) لدالة testSms
var SMS_API_ = 'https://api.sms-gate.app/3rdparty/v1';
var EXEC_URL_ = 'https://script.google.com/macros/s/AKfycbzNR9VjoDsoEdvt6erg0DQ1p9np2KMuqsWcPHM3xocvfG3xG0HvH8CCvbpKC6wxUOU6/exec';
var SMS_DAILY_MAX_ = 100;      // سقف الرسائل يومياً
var SMS_PHONE_TTL_ = 21600;    // رسالة واحدة لكل رقم كل 6 ساعات (أقصى مدة يسمح بها CacheService)

function smsAuth_() {
  var pr = PropertiesService.getScriptProperties();
  var u = pr.getProperty('SMS_USER'), pw = pr.getProperty('SMS_PASS');
  if (!u || !pw) return null;
  return 'Basic ' + Utilities.base64Encode(u + ':' + pw);
}

// 0550123456 -> +213550123456
function intl_(local) { return '+213' + String(local).replace(/\D/g, '').slice(-9); }
// أي صيغة (+213..., 213..., 0...) -> 0550123456
function local_(any) { return '0' + String(any || '').replace(/\D/g, '').slice(-9); }

function smsText_(job) {
  return job.ccp
    ? 'حنين: استلمنا طلبك ' + job.orderId + ' ووصل الدفع. سنتحقق منه ونتصل بك قريبا. شكرا لثقتك!'
    : 'حنين: تم تسجيل طلبك ' + job.orderId + '. للتأكيد ردّ بـ 1، وللإلغاء ردّ بـ 0. شكرا لثقتك!';
}

function sendOrderSms_(job) {
  var pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('SMS_ENABLED') !== '1') return 'off';
  var auth = smsAuth_();
  if (!auth || !PHONE_RE_.test(job.phone)) return 'skip';

  // حماية من إساءة الاستعمال: رقم واحد لا يتلقى إلا رسالة واحدة كل 6 ساعات، وسقف يومي عام
  var cache = CacheService.getScriptCache();
  if (cache.get('sp_' + job.phone)) return 'phone-limit';
  var day = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd');
  var cur = String(pr.getProperty('SMS_DAY') || '').split(':');
  var n = cur[0] === day ? Number(cur[1]) || 0 : 0;
  if (n >= SMS_DAILY_MAX_) return 'daily-limit';
  pr.setProperty('SMS_DAY', day + ':' + (n + 1));
  cache.put('sp_' + job.phone, '1', SMS_PHONE_TTL_);

  var res = UrlFetchApp.fetch(SMS_API_ + '/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: auth },
    payload: JSON.stringify({ textMessage: { text: smsText_(job) }, phoneNumbers: [intl_(job.phone)] })
  });
  Logger.log('sms ' + job.orderId + ' -> ' + res.getResponseCode());
  return res.getResponseCode();
}

// أعمدة الجدول بنفس منطق doPost (العمود بلا عنوان بعد المجموع = الحالة)
function headsOf_(sh) {
  var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  for (var k = 1; k < heads.length; k++) {
    if (heads[k] === '' && (heads[k - 1] === 'product' || heads[k - 1] === 'الدورة')) heads[k] = 'items';
    if (heads[k] === '' && (heads[k - 1] === 'totalprix' || heads[k - 1] === 'total' || heads[k - 1] === 'المجموع')) heads[k] = 'status';
  }
  return heads;
}

// ردّ الزبون: 1/نعم -> "ردّ بـ 1 (SMS)"، 0/لا -> "ردّ بـ 0 (SMS)". عمداً لا نكتب "مؤكد" حتى لا يُطلق أي إرسال تلقائي لشركة التوصيل؛ أنت تراجع وتؤكد بنفسك.
// لا يغيّر إلا الطلبات التي ما زالت "طلب جديد".
function handleSmsReply_(body, key) {
  var want = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
  if (!want || key !== want) return 'forbidden';
  var pl = body.payload || {};
  var text = String(pl.message || '').trim()
    .replace(/[٠-٩]/g, function (d) { return String(d.charCodeAt(0) - 1632); })
    .replace(/[۰-۹]/g, function (d) { return String(d.charCodeAt(0) - 1776); })
    .replace(/[.!،\s]+$/g, '').toLowerCase();
  var label = null;
  if (/^(1|نعم|اي|أي|ايه|oui|ok|yes)$/.test(text)) label = 'ردّ بـ 1 (SMS)';
  else if (/^(0|لا|no|non|الغاء|إلغاء)$/.test(text)) label = 'ردّ بـ 0 (SMS)';
  if (!label) return 'ignored';

  var phone = local_(pl.sender || pl.phoneNumber);
  if (!PHONE_RE_.test(phone)) return 'bad-phone';
  var sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return 'norow';
  var heads = headsOf_(sh);
  var pc = Math.max(heads.indexOf('phone'), heads.indexOf('الهاتف'));
  var sc = Math.max(heads.indexOf('status'), heads.indexOf('الحالة'));
  if (pc < 0 || sc < 0) return 'nocol';
  var ph = sh.getRange(2, pc + 1, last - 1, 1).getValues();
  var st = sh.getRange(2, sc + 1, last - 1, 1).getValues();
  for (var i = ph.length - 1; i >= 0; i--) {          // أحدث طلب لهذا الرقم
    if (String(ph[i][0]).indexOf(phone) < 0) continue;
    var cur = String(st[i][0]).trim();
    if (cur === '' || cur === 'طلب جديد') { sh.getRange(i + 2, sc + 1).setValue(label); return 'ok'; }
    return 'kept';                                    // حالة لاحقة (تأكيد/توصيل...) لا نغيّرها
  }
  return 'nomatch';
}

// ===== دوال تُشغَّل يدوياً من المحرر مرة واحدة =====
// تسجّل في تطبيق البوابة عنوان استقبال الردود (بعد ضبط SMS_USER وSMS_PASS وWEBHOOK_KEY)
function registerSmsWebhook() {
  var key = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
  var auth = smsAuth_();
  if (!key || !auth) throw new Error('اضبط SMS_USER وSMS_PASS وWEBHOOK_KEY أولاً في Script properties');
  var res = UrlFetchApp.fetch(SMS_API_ + '/webhooks', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: auth },
    payload: JSON.stringify({ url: EXEC_URL_ + '?k=' + encodeURIComponent(key), event: 'sms:received' })
  });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText());
}

// يرسل رسالة تجريبية إلى SMS_TEST_PHONE (يتجاوز حدّ الرقم الواحد)
function testSms() {
  var pr = PropertiesService.getScriptProperties();
  var phone = pr.getProperty('SMS_TEST_PHONE'), auth = smsAuth_();
  if (!auth || !PHONE_RE_.test(phone || '')) throw new Error('اضبط SMS_USER وSMS_PASS وSMS_TEST_PHONE (مثل 0550123456)');
  var res = UrlFetchApp.fetch(SMS_API_ + '/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: auth },
    payload: JSON.stringify({ textMessage: { text: smsText_({ orderId: 'HN-000000-TEST', ccp: false }) }, phoneNumbers: [intl_(phone)] })
  });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText());
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
