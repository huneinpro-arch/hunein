// أدوات حماية مشتركة لدوال /api (ليست مسارًا عامًا لأنها خارج مجلد api).
// ملاحظة: حدّ الطلبات هنا في ذاكرة كل نسخة من الدالة (best-effort): يوقف الإغراق العادي،
// لكنه يُصفَّر عند إعادة تشغيل النسخة وقد تتعدد النسخ. للحماية الصارمة استعمل Upstash/Vercel KV.

const ALLOWED_HOSTS = new Set(
  ['hunein.com', 'www.hunein.com']
    .concat(String(process.env.ALLOWED_HOSTS || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean))
);

// localhost مسموح خارج الإنتاج فقط (للتجربة المحلية)
function hostOk(h) {
  if (ALLOWED_HOSTS.has(h)) return true;
  return process.env.VERCEL_ENV !== 'production' && (h === 'localhost' || h === '127.0.0.1');
}

// يقبل الطلب فقط إن جاء Origin (أو Referer) من نطاقاتك. لا يوقف curl المزوِّر وحده، لذلك نضيف حدّ الطلبات.
function hostAllowed(req) {
  const src = req.headers.origin || req.headers.referer || '';
  try {
    const u = new URL(src);
    return (u.protocol === 'https:' || u.hostname === 'localhost' || u.hostname === '127.0.0.1') && hostOk(u.hostname.toLowerCase());
  } catch (e) {
    return false;
  }
}

// عنوان الزائر كما تضعه Vercel (لا يمكن للزائر تزويره في x-real-ip / x-vercel-forwarded-for)
function clientIp(req) {
  const h = req.headers;
  const v = h['x-real-ip'] || h['x-vercel-forwarded-for'] || String(h['x-forwarded-for'] || '').split(',')[0] || (req.socket && req.socket.remoteAddress) || '';
  return String(v).trim().slice(0, 64) || 'unknown';
}

const buckets = new Map(); // key -> [timestamps]

// true = مسموح، false = تجاوز الحدّ
function allow(key, limit, windowMs) {
  const now = Date.now();
  const arr = (buckets.get(key) || []).filter((t) => now - t < windowMs);
  if (arr.length >= limit) { buckets.set(key, arr); return false; }
  arr.push(now);
  buckets.set(key, arr);
  if (buckets.size > 5000) { // تنظيف حتى لا تكبر الذاكرة
    for (const [k, v] of buckets) if (!v.length || now - v[v.length - 1] > windowMs) buckets.delete(k);
  }
  return true;
}

module.exports = { hostAllowed, clientIp, allow };
