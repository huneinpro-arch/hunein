// Vercel Serverless Function — يستقبل وصل الدفع (بريدي موب / CCP) ويرسله إلى واتساب صاحب المشروع
// عبر WhatsApp Cloud API. متغيّرات البيئة المطلوبة في Vercel (لا تضع المفاتيح داخل الملف أبداً):
//   WA_TOKEN      = رمز الوصول الدائم (System User token) لتطبيق واتساب
//   WA_PHONE_ID   = Phone Number ID لرقم واتساب للأعمال المرسِل
//   OWNER_PHONE   = رقمك الذي يستقبل الوصولات بصيغة دولية بدون +، مثال: 213550123456
// ملاحظة: واتساب يسمح بإرسال صورة/ملف حر فقط خلال 24 ساعة من آخر رسالة أرسلتها أنت إلى رقم الأعمال.
const WA_TOKEN = process.env.WA_TOKEN;
const PHONE_ID = process.env.WA_PHONE_ID;
const OWNER = String(process.env.OWNER_PHONE || '').replace(/\D/g, '');
const V = 'v21.0';
const MAX_BYTES = 3 * 1024 * 1024;
const MIMES = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];

const clean = (v, n = 120) => (typeof v === 'string' ? v.replace(/[\r\u0000-\u001f]+/g, ' ').trim().slice(0, n) : '');

function hostAllowed(req) {
  const src = req.headers.origin || req.headers.referer || '';
  try {
    const h = new URL(src).hostname;
    return h === 'hunein.com' || h.endsWith('.hunein.com') || h.endsWith('.vercel.app') || h === 'localhost';
  } catch (e) {
    return false;
  }
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  if (!WA_TOKEN || !PHONE_ID || !OWNER) return res.status(500).json({ ok: false, error: 'WhatsApp env vars are not set' });
  if (!hostAllowed(req)) return res.status(403).json({ ok: false });

  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  b = b || {};

  const mime = clean(b.mime, 40);
  if (!MIMES.includes(mime)) return res.status(400).json({ ok: false, error: 'file type not allowed' });
  const buf = Buffer.from(String(b.file || ''), 'base64');
  if (!buf.length || buf.length > MAX_BYTES) return res.status(400).json({ ok: false, error: 'bad file size' });

  const ext = mime === 'application/pdf' ? 'pdf' : mime.split('/')[1];
  const filename = (clean(b.orderId, 30) || 'receipt') + '.' + ext;
  const caption = [
    '🧾 وصل دفع جديد (بريدي موب / CCP)',
    'رقم الطلب: ' + clean(b.orderId, 30),
    'الاسم: ' + clean(b.name),
    'الهاتف: ' + clean(b.phone, 20),
    'الدورة: ' + clean(b.course, 160),
    'المدفوع: ' + clean(String(b.amount), 12) + ' دج',
    'المتبقي عند الاستلام: ' + clean(String(b.due), 12) + ' دج',
    'الولاية: ' + clean(b.wilaya, 40),
  ].join('\n');

  try {
    const fd = new FormData();
    fd.append('messaging_product', 'whatsapp');
    fd.append('type', mime);
    fd.append('file', new Blob([buf], { type: mime }), filename);
    const up = await fetch(`https://graph.facebook.com/${V}/${PHONE_ID}/media`, { method: 'POST', headers: { Authorization: 'Bearer ' + WA_TOKEN }, body: fd });
    const upj = await up.json().catch(() => ({}));
    if (!up.ok || !upj.id) return res.status(502).json({ ok: false, error: 'media upload failed', detail: upj.error && upj.error.message });

    const isImg = mime.startsWith('image/');
    const msg = { messaging_product: 'whatsapp', to: OWNER, type: isImg ? 'image' : 'document' };
    msg[msg.type] = isImg ? { id: upj.id, caption } : { id: upj.id, caption, filename };
    const sd = await fetch(`https://graph.facebook.com/${V}/${PHONE_ID}/messages`, { method: 'POST', headers: { Authorization: 'Bearer ' + WA_TOKEN, 'Content-Type': 'application/json' }, body: JSON.stringify(msg) });
    const sdj = await sd.json().catch(() => ({}));
    if (!sd.ok) return res.status(502).json({ ok: false, error: 'send failed', detail: sdj.error && sdj.error.message });
    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'upstream error' });
  }
};
