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
    // اسم البلدية يُوحَّد مع قائمة DHD (فروق الحروف والشرطات فقط)؛ ولا يعطّل الطلب إن تعذّر
    try { p.commune = dhdCommune_(p.wilaya, p.commune); } catch (errC) { Logger.log('commune normalize: ' + errC); }
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


// ===== رسائل تتبع الشحنة: تُقرأ من ورقة Suivi وتُرسل حسب قواعد ورقة "SMS Rules" =====
// تشغيلها يحتاج: SMS_TRACK_ENABLED = 1 في Script properties، وتشغيل baselineSmsTracking مرة واحدة،
// ثم installSmsTrackingTrigger مرة واحدة (مؤقّت كل 10 دقائق).
var TRACK_SHEET_ = 'Suivi';
var RULES_SHEET_ = 'SMS Rules';
var LOG_SHEET_ = 'SMS Log';
var TRACK_MAX_PER_RUN_ = 20;   // أقصى عدد رسائل في كل دورة

// القواعد الافتراضية: STATUT (الكلمة كما تكتبها DHD) | التوصيل | مفعّل | نص الرسالة
// {id} رقم الطلب  {track} رقم التتبع  {amount} المبلغ عند الاستلام  {office} اسم المكتب  {commune} البلدية
function trackRulesDefaults_() {
  return [
    ['STATUT', 'التوصيل', 'مفعّل', 'نص الرسالة'],
    ['Drawn', 'الكل', 'نعم', 'حنين: تم شحن طلبك {id} مع DHD (تتبع: {track}). سيتصل بك الموزع قريبا، يرجى الرد على اتصاله وتجهيز المبلغ: {amount} دج.'],
    ['En livraison', 'A domicile', 'لا', 'حنين: طلبك {id} مع الموزع اليوم. يرجى إبقاء هاتفك مفتوحا وتجهيز المبلغ: {amount} دج. شكرا لثقتك.'],
    ['En Station', 'Stopdesk', 'لا', 'حنين: وصل طلبك {id} إلى مكتب DHD ({office}). يرجى استلامه في أقرب وقت ومعك بطاقة الهوية. المبلغ: {amount} دج.'],
    ['Suspendus', 'الكل', 'لا', 'حنين: تعذّر تسليم طلبك {id}. يرجى الرد على اتصال الموزع أو زيارة المكتب قريبا لتفادي إرجاع الطلب.'],
    ['Livre non encaissé', 'الكل', 'لا', 'حنين: تم تسليم طلبك {id}. شكرا لثقتك! لأي مساعدة تواصل معنا عبر واتساب.']
  ];
}

function ssForTrack_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
}

function sheetOrCreate_(ss, name, rows) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  if (rows && rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return sh;
}

function colOf_(heads, names) {
  for (var i = 0; i < names.length; i++) {
    var at = heads.indexOf(names[i]);
    if (at > -1) return at;
  }
  return -1;
}

function fmtMoney_(n) { return String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

function fillTpl_(tpl, v) {
  return String(tpl).replace(/\{(\w+)\}/g, function (m, k) { return v.hasOwnProperty(k) ? v[k] : ''; });
}

// يقرأ ورقة Suivi ويعيد صفوف الطلبات الصالحة (رقم طلب + هاتف + حالة)
function trackRows_(ss) {
  var sh = ss.getSheetByName(TRACK_SHEET_);
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow(), cols = sh.getLastColumn();
  var heads = sh.getRange(1, 1, 1, cols).getValues()[0].map(norm_);
  var cId = colOf_(heads, ['orderid', 'code']), cPh = colOf_(heads, ['phone', 'الهاتف']), cSt = colOf_(heads, ['statut']);
  var cTr = colOf_(heads, ['numerosuivi']), cTot = colOf_(heads, ['prixtotal', 'prixtota']);
  var cDel = colOf_(heads, ['stopdesk']), cOff = colOf_(heads, ['nomstopdesk']), cCom = colOf_(heads, ['commune']);
  if (cId < 0 || cPh < 0 || cSt < 0) return [];
  var data = sh.getRange(2, 1, last - 1, cols).getValues(), out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i], id = String(r[cId]).trim(), st = String(r[cSt]).trim();
    if (!ORDER_RE_.test(id) || !st) continue;
    var phone = local_(String(r[cPh]).split('/')[0]);
    if (!PHONE_RE_.test(phone)) continue;
    out.push({
      id: id, statut: st, phone: phone,
      track: cTr > -1 ? String(r[cTr]).trim() : '',
      amount: cTot > -1 ? fmtMoney_(r[cTot]) : '',
      del: cDel > -1 ? String(r[cDel]).trim().toLowerCase() : '',
      office: cOff > -1 ? String(r[cOff]).trim() : '',
      commune: cCom > -1 ? String(r[cCom]).trim() : ''
    });
  }
  return out;
}

function trackRules_(ss) {
  var sh = sheetOrCreate_(ss, RULES_SHEET_, trackRulesDefaults_());
  if (sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(), rules = [];
  for (var i = 0; i < v.length; i++) {
    var st = String(v[i][0]).trim().toLowerCase();
    if (!st || String(v[i][2]).trim() !== 'نعم' || !String(v[i][3]).trim()) continue;
    rules.push({ statut: st, del: String(v[i][1]).trim().toLowerCase(), text: String(v[i][3]) });
  }
  return rules;
}

function ruleFor_(rules, row) {
  var st = row.statut.toLowerCase();
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (r.statut === st && (r.del === 'الكل' || r.del === '' || r.del === row.del)) return r;
  }
  return null;
}

function trackDailyTake_(pr) {
  var day = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd');
  var cur = String(pr.getProperty('SMS_DAY') || '').split(':');
  var n = cur[0] === day ? Number(cur[1]) || 0 : 0;
  if (n >= SMS_DAILY_MAX_) return false;
  pr.setProperty('SMS_DAY', day + ':' + (n + 1));
  return true;
}

function sendSmsRaw_(phone, text) {
  var res = UrlFetchApp.fetch(SMS_API_ + '/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: smsAuth_() },
    payload: JSON.stringify({ textMessage: { text: text }, phoneNumbers: [intl_(phone)] })
  });
  return res.getResponseCode();
}

// baseline=true: يسجّل الوضع الحالي كأنه "تمّ" دون إرسال أي شيء (حتى لا تصل رسائل لطلبات قديمة)
function trackPass_(baseline) {
  var ss = ssForTrack_(), pr = PropertiesService.getScriptProperties();
  var rows = trackRows_(ss), rules = trackRules_(ss);   // تُنشئ ورقة القواعد عند أول تشغيل
  var log = sheetOrCreate_(ss, LOG_SHEET_, [['مفتاح', 'وقت', 'رقم الطلب', 'STATUT', 'الهاتف', 'النتيجة']]);
  var seen = {};
  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().forEach(function (x) { seen[String(x[0])] = 1; });
  }
  var now = Utilities.formatDate(new Date(), 'Africa/Algiers', 'yyyy-MM-dd HH:mm'), added = [], sent = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], key = r.id + '|' + r.statut.toLowerCase();
    if (seen[key]) continue;
    if (baseline) { seen[key] = 1; added.push([key, now, r.id, r.statut, r.phone, 'baseline']); continue; }
    var rule = ruleFor_(rules, r);
    if (!rule) continue;
    if (sent >= TRACK_MAX_PER_RUN_ || !trackDailyTake_(pr)) break;
    var text = fillTpl_(rule.text, { id: r.id, track: r.track, amount: r.amount, office: r.office, commune: r.commune });
    var code = sendSmsRaw_(r.phone, text);
    if (code >= 200 && code < 300) { seen[key] = 1; sent++; added.push([key, now, r.id, r.statut, r.phone, 'sent ' + code]); }
    else { Logger.log('track sms failed ' + r.id + ' -> ' + code); break; }   // يُعاد المحاولة في الدورة القادمة
  }
  if (added.length) log.getRange(log.getLastRow() + 1, 1, added.length, 6).setValues(added);
  return baseline ? 'baseline:' + added.length : 'sent:' + sent;
}

// تُشغَّل من المؤقّت كل 10 دقائق
function runSmsTracking() {
  var pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('SMS_TRACK_ENABLED') !== '1') return 'off';
  if (pr.getProperty('SMS_TRACK_BASELINE') !== '1') return 'no-baseline';
  if (!smsAuth_()) return 'no-auth';
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 'busy';
  try { return trackPass_(false); } finally { lock.releaseLock(); }
}

// تُشغَّل مرة واحدة يدوياً قبل التفعيل: تُنشئ ورقتي "SMS Rules" و"SMS Log" وتسجّل الوضع الحالي بلا إرسال
function baselineSmsTracking() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var out = trackPass_(true);
    PropertiesService.getScriptProperties().setProperty('SMS_TRACK_BASELINE', '1');
    Logger.log(out);
    return out;
  } finally { lock.releaseLock(); }
}

// تُشغَّل مرة واحدة: تنشئ مؤقّتاً كل 10 دقائق
function installSmsTrackingTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runSmsTracking') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runSmsTracking').timeBased().everyMinutes(10).create();
}

// ===== فحص هيكل إعدادات شركات التوصيل دون كشف أي سرّ =====
// تُطبع أسماء الحقول وأنواعها فقط. القيم النصية تُخفى (يظهر طولها)، ما عدا الروابط (بلا parameters)
// وبضع كلمات معروفة وغير سرّية (مثل Authorization و Bearer).
function redactCfg_(v) {
  var SAFE = /^(authorization|bearer|x-[\w-]+|content-type|get|post|json|true|false)$/i;
  if (v === null || typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'string') {
    if (/^https?:\/\//i.test(v)) return v.split('?')[0].slice(0, 120);
    if (SAFE.test(v)) return v;
    return '<' + v.length + ' chars>';
  }
  if (Array.isArray(v)) return v.map(redactCfg_);
  var o = {};
  Object.keys(v).forEach(function (k) { o[k] = redactCfg_(v[k]); });
  return o;
}

function inspectShippingConfig() {
  var raw = PropertiesService.getScriptProperties().getProperty('SHIPPING_COMPANIES_V1');
  if (!raw) { Logger.log('SHIPPING_COMPANIES_V1 غير موجودة'); return; }
  var data;
  try { data = JSON.parse(raw); } catch (e) { Logger.log('ليست JSON، طولها ' + raw.length); return; }
  Logger.log(JSON.stringify(redactCfg_(data)));
}

// ===== تصدير بلديات DHD (مرة واحدة) إلى ورقة "DHD Communes" =====
// يقرأ مفتاح DHD من إعداداتك (SHIPPING_COMPANIES_V1) ولا يطبعه أبداً. شغّل exportDhdCommunes مرة واحدة.
var COMMUNES_SHEET_ = 'DHD Communes';

function dhdPickName_(it) {
  if (typeof it === 'string') return it.trim();
  if (!it || typeof it !== 'object') return '';
  var keys = ['nom', 'name', 'commune_name', 'nom_commune', 'commune', 'libelle', 'label', 'designation'];
  for (var i = 0; i < keys.length; i++) {
    if (typeof it[keys[i]] === 'string' && it[keys[i]].trim()) return it[keys[i]].trim();
  }
  return '';
}

function dhdPickWilaya_(it) {
  if (!it || typeof it !== 'object') return 0;
  var keys = ['wilaya_id', 'id_wilaya', 'code_wilaya', 'wilaya'];
  for (var i = 0; i < keys.length; i++) {
    var n = Number(it[keys[i]]);
    if (n > 0 && n < 100) return n;
  }
  return 0;
}

// يحوّل أي شكل شائع للردّ (مصفوفة، {data:[...]}، كائن مفاتيحه معرّفات) إلى قائمة عناصر
function dhdList_(j) {
  if (Array.isArray(j)) return j;
  if (j && Array.isArray(j.data)) return j.data;
  if (j && typeof j === 'object') {
    var vals = Object.keys(j).map(function (k) { return j[k]; });
    if (vals.length && vals.every(function (v) { return v && typeof v === 'object'; })) {
      var flat = [];
      vals.forEach(function (v) { flat = Array.isArray(v) ? flat.concat(v) : flat.concat([v]); });
      return flat;
    }
    return Object.keys(j).map(function (k) { return typeof j[k] === 'string' ? j[k] : null; }).filter(Boolean);
  }
  return [];
}

function exportDhdCommunes() {
  var raw = PropertiesService.getScriptProperties().getProperty('SHIPPING_COMPANIES_V1');
  if (!raw) throw new Error('SHIPPING_COMPANIES_V1 غير موجودة');
  var cfg = JSON.parse(raw).dhd;
  if (!cfg || !cfg.api_url || !cfg.api_key) throw new Error('إعدادات DHD ناقصة');
  var base = String(cfg.api_url).replace(/\/+$/, '');
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(COMMUNES_SHEET_) || ss.insertSheet(COMMUNES_SHEET_);

  // يكمل من حيث توقف: الولايات الموجودة في الورقة تُتخطّى، والجديد يُضاف بعدها ويُحفظ ولاية بولاية
  var seen = {}, done = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      seen[r[0] + '|' + r[1]] = 1; done[Number(r[0])] = 1;
    });
  } else {
    sh.getRange(1, 1, 1, 3).setValues([['wilaya_id', 'name', 'raw']]);
  }

  var fails = 0, firstLogged = false;
  for (var w = 1; w <= 58; w++) {
    if (done[w]) continue;
    var code = 0, body = '';
    for (var t = 0; t < 3; t++) {                     // DHD تحدّ الطلبات: عند 429 ننتظر ثم نعيد
      var res = UrlFetchApp.fetch(base + '/api/v1/get/communes?wilaya_id=' + w, {
        method: 'get', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + cfg.api_key, Accept: 'application/json' }
      });
      code = res.getResponseCode(); body = res.getContentText();
      if (code !== 429) break;
      Logger.log('wilaya ' + w + ' -> 429، انتظار ثم إعادة المحاولة');
      Utilities.sleep(25000 * (t + 1));
    }
    if (!firstLogged) { firstLogged = true; Logger.log('wilaya ' + w + ' -> ' + code + ' ' + body.length + ' chars: ' + body.slice(0, 200)); }
    if (code === 429) { Logger.log('ما زالت DHD تحدّ الطلبات؛ أعد التشغيل بعد دقائق لإكمال الباقي'); break; }
    if (code !== 200) {
      Logger.log('wilaya ' + w + ' failed: ' + code + ' ' + body.slice(0, 120));
      if (++fails >= 3) break;
      continue;
    }
    var j;
    try { j = JSON.parse(body); } catch (e) { Logger.log('wilaya ' + w + ' ليس JSON'); if (++fails >= 3) break; continue; }
    var chunk = [];
    dhdList_(j).forEach(function (it) {
      var name = dhdPickName_(it), ww = dhdPickWilaya_(it) || w;
      if (!name) return;
      var key = ww + '|' + name;
      if (seen[key]) return;
      seen[key] = 1;
      chunk.push([ww, name, typeof it === 'object' ? JSON.stringify(it).slice(0, 300) : '']);
    });
    if (chunk.length) sh.getRange(sh.getLastRow() + 1, 1, chunk.length, 3).setValues(chunk);
    Logger.log('wilaya ' + w + ': ' + chunk.length);
    Utilities.sleep(1500);                            // إيقاع أبطأ تجنباً للحدّ
  }
  var missing = [];
  for (var k = 1; k <= 58; k++) {
    var has = false;
    if (done[k]) has = true;
    else { for (var key2 in seen) { if (key2.indexOf(k + '|') === 0) { has = true; break; } } }
    if (!has) missing.push(k);
  }
  var total = Math.max(sh.getLastRow() - 1, 0);
  Logger.log('TOTAL ' + total + (missing.length ? ' | ناقص: ' + missing.join(',') : ' | كل الولايات موجودة'));
  return total;
}

// يقدّم الأسماء فقط (بلا الحقول الخام) لمن يطلب ?dhdcommunes=1
function dhdCommunesJson_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(COMMUNES_SHEET_);
  var out = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { out.push([Number(r[0]), String(r[1])]); });
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ===== توحيد اسم البلدية مع قائمة DHD (ورقة "DHD Communes") =====
// أسماء الولايات كما يرسلها الموقع، والترتيب = رقم الولاية عند DHD (1..58)
var WILAYAS_ = ["Adrar", "Chlef", "Laghouat", "Oum El Bouaghi", "Batna", "Béjaïa", "Biskra", "Béchar", "Blida", "Bouira", "Tamanrasset", "Tébessa", "Tlemcen", "Tiaret", "Tizi Ouzou", "Alger", "Djelfa", "Jijel", "Sétif", "Saïda", "Skikda", "Sidi Bel Abbès", "Annaba", "Guelma", "Constantine", "Médéa", "Mostaganem", "M'Sila", "Mascara", "Ouargla", "Oran", "El Bayadh", "Illizi", "Bordj Bou Arréridj", "Boumerdès", "El Tarf", "Tindouf", "Tissemsilt", "El Oued", "Khenchela", "Souk Ahras", "Tipaza", "Mila", "Aïn Defla", "Naâma", "Aïn Témouchent", "Ghardaïa", "Relizane", "Timimoun", "Bordj Badji Mokhtar", "Ouled Djellal", "Béni Abbès", "In Salah", "In Guezzam", "Touggourt", "Djanet", "El M'Ghair", "El Menia"];

function dhdKey_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

// "15 - Tizi Ouzou" أو "Tizi Ouzou" أو "15" -> رقم الولاية، وإلا 0
function wilayaIdOf_(v) {
  var s = String(v == null ? '' : v).trim();
  var m = /^(\d{1,2})\b/.exec(s);
  if (m && Number(m[1]) >= 1 && Number(m[1]) <= 58) return Number(m[1]);
  var k = dhdKey_(s.replace(/^\d+\s*-\s*/, ''));
  for (var i = 0; i < WILAYAS_.length; i++) if (dhdKey_(WILAYAS_[i]) === k) return i + 1;
  return 0;
}

// بلديات ولاية واحدة من ورقة DHD Communes (مخزَّنة مؤقتاً 6 ساعات)
function dhdCommunesOf_(id) {
  var cache = CacheService.getScriptCache(), ck = 'dc_' + id, hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(COMMUNES_SHEET_), out = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (Number(r[0]) === id) out.push(String(r[1]));
    });
  }
  if (out.length) { try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {} }
  return out;
}

// يعيد اسم DHD الدقيق إن طابق الاسم بعد توحيد الحروف والشرطات؛ وإلا يعيده كما هو (لا نخمّن)
function dhdCommune_(wilaya, commune) {
  var id = wilayaIdOf_(wilaya), c = String(commune == null ? '' : commune).trim();
  if (!id || !c) return c;
  var names = dhdCommunesOf_(id);
  if (!names.length) return c;
  if (names.indexOf(c) > -1) return c;
  var k = dhdKey_(c), hit = names.filter(function (n) { return dhdKey_(n) === k; });
  return hit.length === 1 ? hit[0] : c;
}

// معاينة: تكتب في ورقة "Communes Check" كل طلب اسم بلديته يخالف DHD، دون تغيير أي طلب
function checkCommunes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var rep = [['ورقة', 'صف', 'الولاية', 'البلدية الحالية', 'اسم DHD المقترح', 'النوع', 'بلدية Management لنفس الطلب', 'رقم الطلب']];
  // ما كُتب في Management لكل رقم طلب (للتشخيص، حتى لو لم يكن اسماً صالحاً)
  var mgAny = {}, msh = ss.getSheetByName('Management');
  if (msh && msh.getLastRow() > 1) {
    var mh = msh.getRange(1, 1, 1, msh.getLastColumn()).getValues()[0].map(norm_);
    var mi = mh.indexOf('orderid'), mb = mh.indexOf('baladia');
    if (mi > -1 && mb > -1) {
      msh.getRange(2, 1, msh.getLastRow() - 1, Math.max(mi, mb) + 1).getValues().forEach(function (r) {
        var id = String(r[mi]).trim(); if (id) mgAny[id] = String(r[mb]).trim();
      });
    }
  }
  ['Processing Orders', 'Management'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
    var cw = heads.indexOf('wilaya'), cb = heads.indexOf('baladia'), ci = heads.indexOf('orderid');
    if (cw < 0 || cb < 0) return;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(cw, cb, ci) + 1).getValues();
    data.forEach(function (r, i) {
      var cur = String(r[cb]).trim(), id = ci > -1 ? String(r[ci]).trim() : '', wid = wilayaIdOf_(r[cw]);
      if (!cur || !wid) return;
      var names = dhdCommunesOf_(wid);
      if (!names.length || names.indexOf(cur) > -1) return;
      var fixed = dhdCommune_(r[cw], cur);
      var inMg = Object.prototype.hasOwnProperty.call(mgAny, id) ? (mgAny[id] || '(فارغة)') : '(الطلب غير موجود في Management)';
      rep.push([name, i + 2, String(r[cw]), cur, fixed !== cur ? fixed : '', fixed !== cur ? 'يُصحَّح تلقائياً' : 'يحتاج مراجعتك', inMg, id]);
    });
  });
  var sh2 = ss.getSheetByName('Communes Check') || ss.insertSheet('Communes Check');
  sh2.clear();
  sh2.getRange(1, 1, rep.length, rep[0].length).setValues(rep);
  var auto = rep.filter(function (r) { return r[5] === 'يُصحَّح تلقائياً'; }).length;
  Logger.log('طلبات تخالف DHD: ' + (rep.length - 1) + ' | يُصحَّح تلقائياً: ' + auto + ' | يحتاج مراجعة: ' + (rep.length - 1 - auto));
  return rep.length - 1;
}

// تطبيق: لكل صف اسم بلديته ليس من قائمة DHD:
//  1) في Processing Orders تُؤخذ البلدية المؤكَّدة من ورقة Management لنفس رقم الطلب (راجعها المؤكِّدون)
//  2) وإلا يُصحَّح الفرق في الحروف والشرطات فقط. وما سوى ذلك لا يُمَسّ. كل تغيير يُسجَّل في ورقة "Communes Fixed".
function applyCommuneFixes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var audit = [['ورقة', 'صف', 'رقم الطلب', 'قبل', 'بعد', 'المصدر']], fromMg = 0, auto = 0;

  // البلديات المؤكَّدة في Management: رقم الطلب -> اسم DHD صالح
  var mg = {}, msh = ss.getSheetByName('Management');
  if (msh && msh.getLastRow() > 1) {
    var mh = msh.getRange(1, 1, 1, msh.getLastColumn()).getValues()[0].map(norm_);
    var mi = mh.indexOf('orderid'), mb = mh.indexOf('baladia'), mw = mh.indexOf('wilaya');
    if (mi > -1 && mb > -1 && mw > -1) {
      msh.getRange(2, 1, msh.getLastRow() - 1, Math.max(mi, mb, mw) + 1).getValues().forEach(function (r) {
        var id = String(r[mi]).trim(), val = String(r[mb]).trim(), wid = wilayaIdOf_(r[mw]);
        if (id && val && wid && dhdCommunesOf_(wid).indexOf(val) > -1) mg[id] = val;
      });
    }
  }

  ['Processing Orders', 'Management'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
    var cw = heads.indexOf('wilaya'), cb = heads.indexOf('baladia'), ci = heads.indexOf('orderid');
    if (cw < 0 || cb < 0) return;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(cw, cb, ci) + 1).getValues();
    data.forEach(function (r, i) {
      var cur = String(r[cb]).trim(), id = ci > -1 ? String(r[ci]).trim() : '', wid = wilayaIdOf_(r[cw]);
      if (!cur || !wid) return;
      var names = dhdCommunesOf_(wid);
      if (!names.length || names.indexOf(cur) > -1) return;           // سليم أو لا قائمة للمقارنة
      var fixed = '', src = '';
      if (name === 'Processing Orders' && mg[id]) { fixed = mg[id]; src = 'Management'; }
      else { var f = dhdCommune_(r[cw], cur); if (f !== cur) { fixed = f; src = 'تلقائي'; } }
      if (!fixed || fixed === cur) return;
      sh.getRange(i + 2, cb + 1).setValue(fixed);
      if (src === 'Management') fromMg++; else auto++;
      audit.push([name, i + 2, id, cur, fixed, src]);
    });
  });

  if (audit.length > 1) {
    var as = ss.getSheetByName('Communes Fixed') || ss.insertSheet('Communes Fixed');
    as.clear();
    as.getRange(1, 1, audit.length, 6).setValues(audit);
  }
  Logger.log('تم تصحيح ' + (fromMg + auto) + ' خلية | من Management: ' + fromMg + ' | تلقائي: ' + auto);
  return fromMg + auto;
}

// ===== إحصائيات الشراء (للعرض على الموقع: عدد الطلبة لكل دورة + الأكثر مبيعاً) =====
// تُرجع أعداداً مجمّعة فقط (بدون أسماء ولا هواتف). تُحسب من جدول الطلبات الحقيقي وتُخزَّن 10 دقائق.
var STATS_EXCLUDE_ = /ملغ|إلغاء|الغاء|مرتجع|راجع|رفض|ردّ بـ 0|annul|cancel|retour|return|refus/i;
function hasTok_(t, code) { return new RegExp('(^|[^A-Za-z])' + code + '([^A-Za-z]|$)').test(t); }
// يحوّل نص الدورات (مثل: "Robot Structural Analysis + AutoCAD ...") إلى رموز الدورات الموجودة فيه
function courseIds_(s) {
  var ids = {};
  String(s || '').split('+').forEach(function (t) {
    t = String(t).trim();
    if (!t) return;
    if (/robot/i.test(t) || hasTok_(t, 'RB')) ids.rsa = 1;
    else if (/detailing/i.test(t) || hasTok_(t, 'ASD')) ids.acd = 1;
    else if (/charpente|hangar|hanger/i.test(t) || hasTok_(t, 'CM')) ids.cmh = 1;
    else if (/autocad|dessin|رسم/i.test(t) || hasTok_(t, 'ATCD')) ids.cad = 1;
  });
  return Object.keys(ids);
}
function statsJson_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('stats_v2');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var out = { ok: 1, total: 0, c: { rsa: 0, acd: 0, cmh: 0, cad: 0 }, k: { 1: 0, 2: 0, 3: 0, 4: 0 }, t: Date.now() };
  var sh = getSheet_(), last = sh.getLastRow();
  if (last >= 2) {
    var heads = headsOf_(sh);
    var col = function (names) { for (var i = 0; i < names.length; i++) { var k = heads.indexOf(names[i]); if (k > -1) return k; } return -1; };
    var ci = col(['items', 'الدورات']), cp = col(['product', 'الدورة', 'المنتج']),
        cs = col(['status', 'الحالة']), cn = col(['name', 'الاسم']);
    var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    data.forEach(function (r) {
      var st = cs > -1 ? String(r[cs]) : '', nm = cn > -1 ? String(r[cn]) : '';
      if (STATS_EXCLUDE_.test(st) || /تجريب/.test(nm)) return;           // ملغى/مرتجع/طلب تجريبي لا يُحتسب
      var ids = courseIds_((ci > -1 && r[ci]) ? r[ci] : (cp > -1 ? r[cp] : ''));
      if (!ids.length) return;
      out.total++;
      out.k[Math.min(ids.length, 4)]++;
      ids.forEach(function (id) { out.c[id]++; });
    });
  }
  var js = JSON.stringify(out);
  cache.put('stats_v2', js, 600);
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.dhdcommunes === '1') return dhdCommunesJson_();
  if (e && e.parameter && e.parameter.stats === '1') return statsJson_();
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
