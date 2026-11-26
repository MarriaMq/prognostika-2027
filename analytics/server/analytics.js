/* =====================================================================
   Прогностика 2027 · серверная часть аналитики (Node.js 18+, без зависимостей)

   1) handleTrack(body)          → POST /api/track из мини-аппа
   2) trackServer(user, ev, ...) → события бота и оплаты (вызывать из своего кода)
   Всё складывается в очередь и раз в 5 секунд пачкой уходит в Google-таблицу
   (Apps Script из analytics/apps-script/Code.gs).

   Переменные окружения:
     BOT_TOKEN           токен бота (для проверки initData)
     SHEETS_WEBHOOK_URL  адрес веб-приложения Apps Script (…/exec)
     SHEETS_SECRET       тот же секрет, что в свойствах скрипта (SECRET)
   ===================================================================== */
import crypto from 'node:crypto';

const URL_ = process.env.SHEETS_WEBHOOK_URL;
const SECRET = process.env.SHEETS_SECRET;
const BOT_TOKEN = process.env.BOT_TOKEN;

/* Какие события может прислать мини-апп и какие поля у них разрешены.
   Всё остальное отбрасывается: сервер не доверяет странице. */
const APP_EVENTS = {
  app_open: ['start', 'platform', 'returning', 'lang'],
  app_hide: ['sec', 'screen'],
  lang_select: ['lang'], lang_change: ['lang'],
  onb_done: ['skipped', 'slide'], onb_reopen: [],
  reg_name: ['for'], reg_gender: ['for'], reg_dob: ['for'], reg_confirm: ['for'],
  reg_done: ['for', 'gender', 'age'],
  tariffs_view: ['for', 'access'],
  tariff_click: ['tariff', 'amount', 'for'],
  pay_click: ['tariff', 'amount', 'for'],
  year_lock_click: ['from'],
  promo_open: [], promo_apply: ['ok'],
  pay_success_view: ['tariff'], pay_fail_view: [],
  forecast_open: ['tab', 'month', 'for'],
  month_open: ['month', 'locked', 'for'],
  tab_switch: ['tab'],
  year_topic: ['topic', 'period', 'for'],
  year_period: ['period'],
  section_open: ['topic', 'month', 'section', 'from'],
  profile_open: [], add_person: [], switch_person: ['locked'], support_click: ['from']
};
/* События, которые шлёт только сервер (мини-апп их прислать не может) */
const SERVER_EVENTS = new Set(['bot_start', 'bot_blocked', 'link_activated', 'invoice_created', 'payment_success', 'payment_failed']);

/* ---------- проверка initData (документация Telegram: Validating data received via the Mini App) ---------- */
export function checkInitData(initData, botToken = BOT_TOKEN, maxAgeSec = 24 * 3600) {
  if (typeof initData !== 'string' || initData.length > 4096 || !botToken) return null;
  const params = new URLSearchParams(initData);
  const hash = params.get('hash'); if (!hash) return null;
  params.delete('hash');
  const check = [...params.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => `${k}=${v}`).join('\n');
  const secret = crypto.createHmac('sha256', 'WebAppData').update(botToken).digest();
  const calc = crypto.createHmac('sha256', secret).update(check).digest('hex');
  if (calc.length !== hash.length || !crypto.timingSafeEqual(Buffer.from(calc), Buffer.from(hash))) return null;
  const age = Date.now() / 1000 - Number(params.get('auth_date') || 0);
  if (!(age >= 0 && age <= maxAgeSec)) return null;
  try { const u = JSON.parse(params.get('user') || 'null'); return u && u.id ? u : null; } catch { return null; }
}

const clean = v => {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return Number.isFinite(v) ? Math.round(v * 100) / 100 : undefined;
  if (typeof v === 'string') return v.slice(0, 80);
  return undefined;
};

/* ---------- 1) POST /api/track ----------
   body: {initData, events:[{ev, ts, ...поля}]}
   Ответ всегда 204, чтобы страница не ждала. Пример для Express:
     app.post('/api/track', express.json({limit: '64kb'}), (req, res) => { handleTrack(req.body); res.sendStatus(204); });
*/
export function handleTrack(body) {
  const user = checkInitData(body && body.initData);
  if (!user || !Array.isArray(body.events)) return 0;
  const now = Date.now(); let n = 0;
  for (const e of body.events.slice(0, 50)) {
    const allow = e && APP_EVENTS[e.ev]; if (!allow) continue;
    const props = {};
    for (const k of allow) { const v = clean(e[k]); if (v !== undefined) props[k] = v; }
    /* время с телефона берём, только если оно правдоподобное (не старше суток и не из будущего) */
    const ts = Number(e.ts); const t = ts > now - 864e5 && ts <= now + 6e4 ? ts : now;
    push(user, e.ev, props, 'app', t); n++;
  }
  return n;
}

/* ---------- 2) события бота и оплаты ----------
   user: объект пользователя Telegram ({id, username, first_name}) или просто ID.
   Где вызывать:
     /start [код]                 → trackServer(from, 'bot_start', {start: код || ''})
     ссылка-доступ сработала      → trackServer(from, 'link_activated', {tariff, link: название})
     создан счёт monoPay          → trackServer(user, 'invoice_created', {tariff, amount, order_id})
     вебхук: оплата успешна        → trackServer(user, 'payment_success', {tariff, amount, order_id})
     вебхук: отказ/ошибка          → trackServer(user, 'payment_failed', {tariff, amount, order_id, reason})
     sendMessage вернул 403 (блок) → trackServer(id, 'bot_blocked')
   tariff: 'months' | 'year' | 'all' | 'promo'; amount в гривнах. */
export function trackServer(user, ev, props = {}) {
  if (!SERVER_EVENTS.has(ev)) throw new Error('unknown server event ' + ev);
  const p = {}; for (const [k, v] of Object.entries(props)) { const c = clean(v); if (c !== undefined) p[k] = c; }
  push(typeof user === 'object' ? user : {id: user}, ev, p, 'bot', Date.now());
}

/* ---------- очередь и отправка в таблицу ---------- */
const Q = [];
function push(u, ev, props, source, ts) {
  if (Q.length > 20000) Q.shift(); /* таблица недоступна долго: старые события выкидываем, а не роняем сервер */
  Q.push({ts, tg_id: String(u.id), username: u.username || '', first_name: u.first_name || '', ev, props, source});
}
let busy = false;
export async function flush() {
  if (busy || !Q.length || !URL_ || !SECRET) return;
  busy = true;
  const batch = Q.splice(0, 500);
  try {
    const r = await fetch(URL_, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({secret: SECRET, events: batch}), redirect: 'follow', signal: AbortSignal.timeout(30000)});
    const j = await r.json().catch(() => ({}));
    if (!j.ok) throw new Error(j.error || 'HTTP ' + r.status);
  } catch (err) {
    Q.unshift(...batch); /* не потеряли: попробуем в следующий раз */
    console.warn('[analytics] таблица не приняла события:', err.message);
  } finally { busy = false; }
}
setInterval(flush, 5000).unref();
process.on('beforeExit', flush);
