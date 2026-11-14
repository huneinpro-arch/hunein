// Vercel Serverless Function — يستقبل وصل الدفع (بريدي موب / CCP) ويرسله إلى واتساب صاحب المشروع
// عبر WhatsApp Cloud API. متغيّرات البيئة المطلوبة في Vercel (لا تضع المفاتيح داخل الملف أبداً):
//   WA_TOKEN      = رمز الوصول الدائم (System User token) لتطبيق واتساب
//   WA_PHONE_ID   = Phone Number ID لرقم واتساب للأعمال المرسِل
//   OWNER_PHONE   = رقمك الذي يستقبل الوصولات بصيغة دولية بدون +، مثال: 213550123456
// ملاحظة: واتساب يسمح بإرسال صورة/ملف حر فقط خلال 24 ساعة من آخر رسالة أرسلتها أنت إلى رقم الأعمال.
const { hostAllowed, clientIp, allow } = require('../lib/guard');

const WA_TOKEN = process.env.WA_TOKEN;
const PHONE_ID = process.env.WA_PHONE_ID;
const OWNER = String(process.env.OWNER_PHONE || '').replace(/\D/g, '');
const V = 'v21.0';
const MAX_BYTES = 3 * 1024 * 1024;
const MIMES = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];

const clean = (v, n = 120) => (typeof v === 'string' ? v.replace(/[\r\u0000-\u001f\u007f]+/g, ' ').trim().slice(0, n) : '');

// رقم بين 0 و1,000,000 (دج) أو null إن كان غير صالح
function money(v) {
  const n = Number(v);
  return isFinite(n) && n >= 0 && n <= 1000000 ? Math.round(n) : null;
}

// يتأكد أن بداية الملف تطابق النوع المعلن (لا نثق بالنوع الذي يرسله الزائر)
function magicOk(buf, mime) {
  if (mime === 'image/jpeg') return buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff;
  if (mime === 'image/png') return buf.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  if (mime === 'image/webp') return buf.slice(0, 4).toString('latin1') === 'RIFF' && buf.slice(8, 12).toString('latin1') === 'WEBP';
  if (mime === 'application/pdf') return buf.slice(0, 5).toString('latin1') === '%PDF-';
  return false;
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  if (!hostAllowed(req)) return res.status(403).json({ ok: false });
  // حدّ صارم: وصل واحد حقيقي لكل طلب؛ نسمح بـ 15 كل 10 دقائق لكل عنوان لأن عنوان الجوال قد يشاركه كثيرون
  if (!allow('receipt:' + clientIp(req), 15, 10 * 60 * 1000)) return res.status(429).json({ ok: false });
  if (!WA_TOKEN || !PHONE_ID || !OWNER) return res.status(500).json({ ok: false });

  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  b = b && typeof b === 'object' ? b : {};

  const mime = clean(b.mime, 40);
  if (!MIMES.includes(mime)) return res.status(400).json({ ok: false, error: 'file type not allowed' });
  const b64 = String(b.file || '');
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(b64)) return res.status(400).json({ ok: false, error: 'bad file' });
  const buf = Buffer.from(b64, 'base64');
  if (!buf.length || buf.length > MAX_BYTES) return res.status(400).json({ ok: false, error: 'bad file size' });
  if (!magicOk(buf, mime)) return res.status(400).json({ ok: false, error: 'file content does not match type' });

  const orderId = clean(b.orderId, 30);
  if (!/^HN-\d{6}-[A-Z2-9]{4}$/.test(orderId)) return res.status(400).json({ ok: false, error: 'bad order id' });
  const amount = money(b.amount), due = money(b.due);
  if (amount === null || due === null) return res.status(400).json({ ok: false, error: 'bad amount' });

  const ext = mime === 'application/pdf' ? 'pdf' : mime.split('/')[1];
  const filename = orderId + '.' + ext;
  const caption = [
    '🧾 وصل دفع جديد (بريدي موب / CCP)',
    'رقم الطلب: ' + orderId,
    'الاسم: ' + clean(b.name, 80),
    'الهاتف: ' + clean(b.phone, 20),
    'الدورة: ' + clean(b.course, 160),
    'المدفوع: ' + amount + ' دج',
    'المتبقي عند الاستلام: ' + due + ' دج',
    'الولاية: ' + clean(b.wilaya, 40),
  ].join('\n');

  try {
    const fd = new FormData();
    fd.append('messaging_product', 'whatsapp');
    fd.append('type', mime);
    fd.append('file', new Blob([buf], { type: mime }), filename);
    const up = await fetch(`https://graph.facebook.com/${V}/${PHONE_ID}/media`, { method: 'POST', headers: { Authorization: 'Bearer ' + WA_TOKEN }, body: fd });
    const upj = await up.json().catch(() => ({}));
    if (!up.ok || !upj.id) {
      console.error('receipt: media upload failed', up.status, upj.error && upj.error.message); // التفاصيل في سجلات Vercel فقط
      return res.status(502).json({ ok: false, error: 'media upload failed' });
    }

    const isImg = mime.startsWith('image/');
    const msg = { messaging_product: 'whatsapp', to: OWNER, type: isImg ? 'image' : 'document' };
    msg[msg.type] = isImg ? { id: upj.id, caption } : { id: upj.id, caption, filename };
    const sd = await fetch(`https://graph.facebook.com/${V}/${PHONE_ID}/messages`, { method: 'POST', headers: { Authorization: 'Bearer ' + WA_TOKEN, 'Content-Type': 'application/json' }, body: JSON.stringify(msg) });
    const sdj = await sd.json().catch(() => ({}));
    if (!sd.ok) {
      console.error('receipt: send failed', sd.status, sdj.error && sdj.error.message);
      return res.status(502).json({ ok: false, error: 'send failed' });
    }
    return res.status(200).json({ ok: true });
  } catch (e) {
    console.error('receipt: upstream error');
    return res.status(502).json({ ok: false, error: 'upstream error' });
  }
};
