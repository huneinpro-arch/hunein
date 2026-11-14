// Vercel Serverless Function — Meta Conversions API (server-side events)
// يحتاج متغيّر البيئة META_ACCESS_TOKEN في Vercel (لا تضع المفتاح داخل الملف أبداً).
const crypto = require('crypto');
const { hostAllowed, clientIp, allow } = require('../lib/guard');

const PIXEL_ID = process.env.META_PIXEL_ID || '1473274858152641';
const TOKEN = process.env.META_ACCESS_TOKEN;
const TEST_CODE = process.env.META_TEST_CODE || ''; // اختياري: لتجربة الأحداث في Test Events
const API_VERSION = 'v21.0';

const ALLOWED_EVENTS = ['PageView', 'ViewContent', 'AddToCart', 'InitiateCheckout', 'Purchase', 'Lead', 'Contact'];

const sha = (v) => crypto.createHash('sha256').update(String(v).trim().toLowerCase()).digest('hex');
const clean = (v, n = 120) => (typeof v === 'string' ? v.replace(/[\u0000-\u001f\u007f]+/g, ' ').trim().slice(0, n) : '');

function normPhone(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.startsWith('00')) d = d.slice(2);
  if (d.startsWith('0')) d = '213' + d.slice(1);
  return d.length >= 11 && d.length <= 15 ? d : '';
}

// يبني custom_data من قيم مفحوصة النوع والمدى بدل نسخ ما يرسله الزائر كما هو
function buildCustom(c) {
  const out = {};
  if (!c || typeof c !== 'object') return out;
  const v = Number(c.value);
  if (c.value !== undefined && isFinite(v) && v >= 0 && v <= 1000000) out.value = v;
  if (c.currency === 'DZD') out.currency = 'DZD';
  if (Array.isArray(c.content_ids)) out.content_ids = c.content_ids.slice(0, 10).map((x) => clean(String(x), 40)).filter(Boolean);
  if (c.content_name !== undefined) out.content_name = clean(String(c.content_name), 120);
  if (c.content_type !== undefined) out.content_type = clean(String(c.content_type), 30);
  if (c.content_category !== undefined) out.content_category = clean(String(c.content_category), 60);
  const n = Number(c.num_items);
  if (c.num_items !== undefined && Number.isInteger(n) && n >= 0 && n <= 20) out.num_items = n;
  return out;
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  if (!hostAllowed(req)) return res.status(403).json({ ok: false });
  // حدّ لكل زائر: زيارة عادية ترسل بضعة أحداث؛ الحدّ سخي لأن شبكات الجوال بالجزائر تتشارك عنوانًا واحدًا بين كثيرين
  const ip = clientIp(req);
  if (!allow('capi:' + ip, 120, 60 * 1000)) return res.status(429).json({ ok: false });
  if (!TOKEN) return res.status(500).json({ ok: false });

  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  b = b && typeof b === 'object' ? b : {};

  const name = clean(b.event, 40);
  if (!ALLOWED_EVENTS.includes(name)) return res.status(400).json({ ok: false });

  const ud = {
    client_ip_address: ip === 'unknown' ? undefined : ip,
    client_user_agent: clean(req.headers['user-agent'], 400),
    country: [sha('dz')],
  };
  const fbp = clean(b.fbp, 200), fbc = clean(b.fbc, 300);
  if (/^fb\.\d\.\d+\.\d+$/.test(fbp)) ud.fbp = fbp;
  if (/^fb\.\d\.\d+\.[\w-]+$/.test(fbc)) ud.fbc = fbc;

  const u = b.user;
  if (u && typeof u === 'object') {
    const ph = normPhone(u.phone);
    if (ph) { ud.ph = [sha(ph)]; ud.external_id = [sha(ph)]; }
    const parts = clean(u.name, 120).split(/\s+/).filter(Boolean);
    if (parts[0]) ud.fn = [sha(parts[0])];
    if (parts.length > 1) ud.ln = [sha(parts[parts.length - 1])];
    if (clean(u.commune, 60)) ud.ct = [sha(clean(u.commune, 60))];
    if (clean(u.wilaya, 40)) ud.st = [sha(clean(u.wilaya, 40))];
  }

  // رابط الصفحة يجب أن يكون من نطاقك فقط
  let srcUrl;
  try {
    const su = new URL(clean(b.url, 500));
    if (su.protocol === 'https:' && hostAllowed({ headers: { origin: su.origin } })) srcUrl = su.href;
  } catch (e) { /* يُترك فارغًا */ }

  const eid = clean(b.event_id, 120);
  const payload = {
    data: [{
      event_name: name,
      event_time: Math.floor(Date.now() / 1000),
      event_id: /^[\w-]{1,120}$/.test(eid) ? eid : undefined,
      event_source_url: srcUrl,
      action_source: 'website',
      user_data: ud,
      custom_data: buildCustom(b.custom),
    }],
  };
  if (TEST_CODE) payload.test_event_code = TEST_CODE;

  try {
    const r = await fetch(
      `https://graph.facebook.com/${API_VERSION}/${PIXEL_ID}/events?access_token=${encodeURIComponent(TOKEN)}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }
    );
    if (!r.ok) console.error('capi: Meta responded', r.status); // التفاصيل في سجلات Vercel فقط
    return res.status(r.ok ? 200 : 502).json({ ok: r.ok });
  } catch (e) {
    console.error('capi: upstream error');
    return res.status(502).json({ ok: false });
  }
};
