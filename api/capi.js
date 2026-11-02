// Vercel Serverless Function — Meta Conversions API (server-side events)
// يحتاج متغيّر البيئة META_ACCESS_TOKEN في Vercel (لا تضع المفتاح داخل الملف أبداً).
const crypto = require('crypto');

const PIXEL_ID = process.env.META_PIXEL_ID || '1473274858152641';
const TOKEN = process.env.META_ACCESS_TOKEN;
const TEST_CODE = process.env.META_TEST_CODE || ''; // اختياري: لتجربة الأحداث في Test Events
const API_VERSION = 'v21.0';

const ALLOWED_EVENTS = ['PageView', 'ViewContent', 'AddToCart', 'InitiateCheckout', 'Purchase', 'Lead', 'Contact'];
const CUSTOM_KEYS = ['value', 'currency', 'content_ids', 'content_name', 'content_type', 'content_category', 'num_items'];

const sha = (v) => crypto.createHash('sha256').update(String(v).trim().toLowerCase()).digest('hex');
const clean = (v, n = 120) => (typeof v === 'string' ? v.trim().slice(0, n) : '');

function normPhone(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.startsWith('00')) d = d.slice(2);
  if (d.startsWith('0')) d = '213' + d.slice(1);
  return d.length >= 11 ? d : '';
}

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
  if (!TOKEN) return res.status(500).json({ ok: false, error: 'META_ACCESS_TOKEN is not set' });
  if (!hostAllowed(req)) return res.status(403).json({ ok: false });

  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  b = b || {};

  const name = clean(b.event, 40);
  if (!ALLOWED_EVENTS.includes(name)) return res.status(400).json({ ok: false, error: 'event not allowed' });

  const ip = (req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.socket?.remoteAddress || '';
  const ud = {
    client_ip_address: ip,
    client_user_agent: clean(req.headers['user-agent'], 400),
    country: [sha('dz')],
  };
  const fbp = clean(b.fbp, 200), fbc = clean(b.fbc, 300);
  if (fbp) ud.fbp = fbp;
  if (fbc) ud.fbc = fbc;

  const u = b.user;
  if (u && typeof u === 'object') {
    const ph = normPhone(u.phone);
    if (ph) { ud.ph = [sha(ph)]; ud.external_id = [sha(ph)]; }
    const parts = clean(u.name, 120).split(/\s+/).filter(Boolean);
    if (parts[0]) ud.fn = [sha(parts[0])];
    if (parts.length > 1) ud.ln = [sha(parts[parts.length - 1])];
    if (clean(u.commune)) ud.ct = [sha(u.commune)];
    if (clean(u.wilaya)) ud.st = [sha(u.wilaya)];
  }

  const custom = {};
  if (b.custom && typeof b.custom === 'object') {
    for (const k of CUSTOM_KEYS) if (b.custom[k] !== undefined) custom[k] = b.custom[k];
  }

  const payload = {
    data: [{
      event_name: name,
      event_time: Math.floor(Date.now() / 1000),
      event_id: clean(b.event_id, 120) || undefined,
      event_source_url: clean(b.url, 500) || undefined,
      action_source: 'website',
      user_data: ud,
      custom_data: custom,
    }],
  };
  if (TEST_CODE) payload.test_event_code = TEST_CODE;

  try {
    const r = await fetch(
      `https://graph.facebook.com/${API_VERSION}/${PIXEL_ID}/events?access_token=${encodeURIComponent(TOKEN)}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }
    );
    const j = await r.json().catch(() => ({}));
    return res.status(r.ok ? 200 : 502).json({ ok: r.ok, fb: r.ok ? undefined : j });
  } catch (e) {
    return res.status(502).json({ ok: false });
  }
};
