/* =====================================================================
   Прогностика 2027 · приёмник аналитики для Google-таблицы
   Ставится в саму таблицу: Расширения → Apps Script → вставить этот файл.
   Сервер бота шлёт сюда POST с событиями, скрипт пишет их на лист «События»
   и обновляет строку человека на листе «Пользователи».
   Мини-апп сюда НЕ ходит напрямую: только сервер, у которого есть SECRET.
   ===================================================================== */

const SH_EV = 'События';
const SH_US = 'Пользователи';

/* Колонки листа «События» (порядок важен, формулы дашбордов на него опираются) */
const EV_COLS = ['Время', 'TG ID', 'Username', 'Событие', 'Что произошло', 'Тариф', 'Сумма, грн',
  'Месяц', 'Раздел', 'Период', 'Закрыто', 'Откуда', 'Детали', 'Источник', 'Демо'];

/* Колонки листа «Пользователи»: одна строка на человека, скрипт держит её в актуальном виде */
const US_COLS = ['TG ID', 'Username', 'Имя в Telegram', 'Статус', 'Откуда пришёл', 'Первый визит',
  'Последняя активность', 'Зашёл в бот', 'Открыл мини-апп', 'Открытий мини-аппа', 'Регистрация',
  'Пол', 'Возраст', 'Язык', 'Смотрел тарифы', 'Нажал «Оплатить»', 'Неудачных оплат', 'Тариф',
  'Оплачено, грн', 'Дата оплаты', 'Просмотров после оплаты', 'Последний просмотр прогноза',
  'Заблокировал бота', 'Демо'];
const U = {}; US_COLS.forEach((c, i) => U[c] = i);

/* Человеческие названия событий: колонка «Что произошло» */
const LABEL = {
  bot_start: 'Зашёл в бот (/start)',
  bot_blocked: 'Заблокировал бота',
  link_activated: 'Получил доступ по ссылке',
  invoice_created: 'Создан счёт monoPay',
  payment_success: 'Оплатил',
  payment_failed: 'Оплата не прошла',
  app_open: 'Открыл мини-апп',
  app_hide: 'Свернул мини-апп',
  lang_select: 'Выбрал язык',
  onb_done: 'Прошёл онбординг',
  reg_name: 'Регистрация: имя',
  reg_gender: 'Регистрация: пол',
  reg_dob: 'Регистрация: дата рождения',
  reg_confirm: 'Регистрация: проверка данных',
  reg_done: 'Зарегистрировался',
  tariffs_view: 'Смотрел тарифы',
  tariff_click: 'Выбрал тариф',
  pay_click: 'Нажал «Оплатить»',
  year_lock_click: 'Нажал на закрытый год',
  promo_open: 'Открыл ввод промокода',
  promo_apply: 'Ввёл промокод',
  pay_success_view: 'Увидел «Оплата прошла»',
  pay_fail_view: 'Увидел «Оплата не прошла»',
  forecast_open: 'Открыл прогноз',
  month_open: 'Открыл месяц',
  tab_switch: 'Переключил вкладку',
  year_topic: 'Открыл годовой прогноз',
  year_period: 'Переключил период года',
  section_open: 'Раскрыл раздел',
  profile_open: 'Открыл профиль',
  add_person: 'Купить для близкого',
  switch_person: 'Переключил человека',
  lang_change: 'Сменил язык',
  support_click: 'Написать в поддержку',
  onb_reopen: 'Открыл «Что такое прогностика»'
};
const TARIFF = {months: 'Месяцы', year: 'Год', all: 'Всё вместе', promo: 'Промокод'};
const TOPIC = {general: 'Общий', personal: 'Личный'};

/* Статусы (сегменты). Порядок = приоритет: человек попадает в первый подходящий */
const STATUS = {
  blocked: 'Заблокировал бота',
  paidNoRead: 'Оплатил, но не открывал прогноз',
  paidReads: 'Оплатил и читает',
  linkNoRead: 'Доступ по ссылке, не открывал',
  linkReads: 'Доступ по ссылке, читает',
  tried: 'Пытался оплатить, не оплатил',
  sawTariffs: 'Смотрел тарифы, не оплатил',
  registered: 'Зарегистрировался, до тарифов не дошёл',
  openedApp: 'Открыл мини-апп, не зарегистрировался',
  botOnly: 'Зашёл в бот, мини-апп не открыл'
};

/* ---------------- вход ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
    if (!secret || body.secret !== secret) return json({ok: false, error: 'forbidden'});
    const evs = (Array.isArray(body.events) ? body.events : []).slice(0, 1000).filter(validEvent);
    if (!evs.length) return json({ok: true, n: 0});
    const ss = SpreadsheetApp.getActive();
    if (!body.demo) clearDemoOnce(ss);
    appendEvents(ss, evs);
    upsertUsers(ss, evs);
    return json({ok: true, n: evs.length});
  } catch (err) {
    return json({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}
function doGet() { return json({ok: true, service: 'prognostika-analytics'}); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* Событие от сервера: {ts, tg_id, username, first_name, ev, props, source} */
function validEvent(x) {
  return x && typeof x.ev === 'string' && /^[a-z_]{2,32}$/.test(x.ev) && /^\d{1,15}$/.test(String(x.tg_id || ''));
}

/* ---------------- лист «События» ---------------- */
function eventRow(x) {
  const p = x.props || {};
  const rest = Object.assign({}, p);
  ['tariff', 'amount', 'month', 'topic', 'section', 'period', 'locked', 'from'].forEach(k => delete rest[k]);
  if (p.section && p.topic) rest.topic = p.topic; /* у раздела года колонка «Раздел» = название рубрики */
  return [
    new Date(Number(x.ts) || Date.now()),
    String(x.tg_id),
    x.username ? '@' + String(x.username).replace(/^@/, '') : '',
    x.ev,
    LABEL[x.ev] || x.ev,
    TARIFF[p.tariff] || p.tariff || '',
    p.amount != null && p.amount !== '' ? Number(p.amount) : '',
    p.month ? Number(p.month) : '',
    p.section ? String(p.section) : (TOPIC[p.topic] || p.topic || ''),
    p.period || '',
    p.locked === true || p.locked === 1 ? 1 : '',
    p.from || '',
    Object.keys(rest).length ? JSON.stringify(rest) : '',
    x.source || '',
    x.demo ? 1 : ''
  ];
}
function appendEvents(ss, evs) {
  const sh = ss.getSheetByName(SH_EV);
  const rows = evs.map(eventRow);
  ensureRows(sh, rows.length);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, EV_COLS.length).setValues(rows);
}

/* ---------------- лист «Пользователи» ---------------- */
function upsertUsers(ss, evs) {
  const sh = ss.getSheetByName(SH_US);
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const byId = {};
  evs.slice().sort((a, b) => a.ts - b.ts).forEach(x => (byId[x.tg_id] = byId[x.tg_id] || []).push(x));
  const fresh = [];
  Object.keys(byId).forEach(id => {
    const i = ids.indexOf(String(id));
    let row = i >= 0 ? sh.getRange(i + 2, 1, 1, US_COLS.length).getValues()[0] : emptyUser(id);
    byId[id].forEach(x => { row = applyEvent(row, x); });
    row[U['Статус']] = statusOf(row);
    if (i >= 0) sh.getRange(i + 2, 1, 1, US_COLS.length).setValues([row]);
    else fresh.push(row);
  });
  if (fresh.length) ensureRows(sh, fresh.length);
  if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, US_COLS.length).setValues(fresh);
}
/* лист сам растёт, когда строки заканчиваются */
function ensureRows(sh, n) { const free = sh.getMaxRows() - sh.getLastRow(); if (free < n) sh.insertRowsAfter(sh.getMaxRows(), Math.max(n - free, 5000)); }
function emptyUser(id) { const r = US_COLS.map(() => ''); r[U['TG ID']] = String(id); return r; }
const num = v => Number(v) || 0;
const asDate = v => v instanceof Date ? v : (v ? new Date(v) : null);
const READ = {forecast_open: 1, month_open: 1, year_topic: 1, section_open: 1, year_period: 1, tab_switch: 1};

/* Вся логика «кто есть кто» в одном месте. Чистая функция: строка + событие → новая строка */
function applyEvent(row, x) {
  const r = row.slice(), p = x.props || {}, t = new Date(Number(x.ts) || Date.now());
  const setFirst = col => { if (!r[U[col]]) r[U[col]] = t; };
  if (x.username) r[U['Username']] = '@' + String(x.username).replace(/^@/, '');
  if (x.first_name) r[U['Имя в Telegram']] = String(x.first_name).slice(0, 64);
  if (x.demo) r[U['Демо']] = 1;
  setFirst('Первый визит');
  if (!asDate(r[U['Первый визит']]) || t < asDate(r[U['Первый визит']])) r[U['Первый визит']] = t;
  if (!asDate(r[U['Последняя активность']]) || t > asDate(r[U['Последняя активность']])) r[U['Последняя активность']] = t;
  if (x.ev !== 'bot_blocked' && x.source === 'app') r[U['Заблокировал бота']] = '';

  switch (x.ev) {
    case 'bot_start':
      setFirst('Зашёл в бот');
      if (!r[U['Откуда пришёл']]) r[U['Откуда пришёл']] = p.start ? String(p.start).slice(0, 64) : 'Напрямую';
      r[U['Заблокировал бота']] = '';
      break;
    case 'app_open':
      setFirst('Открыл мини-апп');
      r[U['Открытий мини-аппа']] = num(r[U['Открытий мини-аппа']]) + 1;
      if (p.lang) r[U['Язык']] = p.lang === 'uk' ? 'UA' : 'RU';
      if (!r[U['Откуда пришёл']]) r[U['Откуда пришёл']] = p.start ? String(p.start).slice(0, 64) : 'Напрямую';
      break;
    case 'lang_select': case 'lang_change':
      if (p.lang) r[U['Язык']] = p.lang === 'uk' ? 'UA' : 'RU';
      break;
    case 'reg_done':
      if (p.for !== 'other') {
        setFirst('Регистрация');
        if (p.gender) r[U['Пол']] = p.gender === 'm' ? 'М' : 'Ж';
        if (p.age) r[U['Возраст']] = num(p.age);
      }
      break;
    case 'tariffs_view': r[U['Смотрел тарифы']] = num(r[U['Смотрел тарифы']]) + 1; break;
    case 'pay_click': r[U['Нажал «Оплатить»']] = num(r[U['Нажал «Оплатить»']]) + 1; break;
    case 'payment_failed': r[U['Неудачных оплат']] = num(r[U['Неудачных оплат']]) + 1; break;
    case 'payment_success':
      r[U['Оплачено, грн']] = num(r[U['Оплачено, грн']]) + num(p.amount);
      r[U['Тариф']] = mergeTariff(r[U['Тариф']], p.tariff);
      setFirst('Дата оплаты');
      break;
    case 'link_activated':
      r[U['Тариф']] = mergeTariff(r[U['Тариф']], p.tariff);
      setFirst('Дата оплаты');
      r[U['Откуда пришёл']] = 'Ссылка: ' + String(p.link || '').slice(0, 48);
      break;
    case 'bot_blocked': r[U['Заблокировал бота']] = t; break;
  }
  const paidAt = asDate(r[U['Дата оплаты']]);
  if (READ[x.ev] && paidAt && t >= paidAt) {
    r[U['Просмотров после оплаты']] = num(r[U['Просмотров после оплаты']]) + 1;
    r[U['Последний просмотр прогноза']] = t;
  }
  return r;
}
/* Месяцы + Год = Всё вместе. Промокод даёт всё вместе */
function mergeTariff(cur, add) {
  const a = TARIFF[add] || add || '', c = cur || '';
  if (!c) return a === 'Промокод' ? 'Всё вместе (промокод)' : a;
  if (c.indexOf('Всё вместе') === 0) return c;
  if ((c === 'Месяцы' && a === 'Год') || (c === 'Год' && a === 'Месяцы')) return 'Всё вместе';
  if (a === 'Всё вместе' || a === 'Промокод') return a === 'Промокод' ? 'Всё вместе (промокод)' : a;
  return c;
}
function statusOf(r) {
  const paid = num(r[U['Оплачено, грн']]) > 0, link = !paid && !!r[U['Дата оплаты']];
  const reads = num(r[U['Просмотров после оплаты']]) > 0;
  if (r[U['Заблокировал бота']]) return STATUS.blocked;
  if (paid) return reads ? STATUS.paidReads : STATUS.paidNoRead;
  if (link) return reads ? STATUS.linkReads : STATUS.linkNoRead;
  if (num(r[U['Нажал «Оплатить»']]) > 0 || num(r[U['Неудачных оплат']]) > 0) return STATUS.tried;
  if (num(r[U['Смотрел тарифы']]) > 0) return STATUS.sawTariffs;
  if (r[U['Регистрация']]) return STATUS.registered;
  if (r[U['Открыл мини-апп']]) return STATUS.openedApp;
  return STATUS.botOnly;
}

/* ---------------- обслуживание ---------------- */
/* Демо-данные (колонка «Демо» = 1) удаляются сами при первом настоящем событии */
function clearDemoOnce(ss) {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('DEMO_CLEARED')) return;
  clearDemo(ss);
  pr.setProperty('DEMO_CLEARED', '1');
}
function clearDemo(ss) {
  ss = ss || SpreadsheetApp.getActive();
  [[SH_EV, EV_COLS.length], [SH_US, US_COLS.length]].forEach(([name, n]) => {
    const sh = ss.getSheetByName(name), last = sh.getLastRow();
    if (last < 2) return;
    const keep = sh.getRange(2, 1, last - 1, n).getValues().filter(r => r[n - 1] !== 1 && r[n - 1] !== '1');
    sh.getRange(2, 1, last - 1, n).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, n).setValues(keep);
  });
}
/* Пересобрать «Пользователи» с нуля по всем событиям (если что-то правили руками) */
function rebuildUsers() {
  const ss = SpreadsheetApp.getActive(), ev = ss.getSheetByName(SH_EV), us = ss.getSheetByName(SH_US);
  const last = ev.getLastRow();
  if (us.getLastRow() > 1) us.getRange(2, 1, us.getLastRow() - 1, US_COLS.length).clearContent();
  if (last < 2) return;
  const evs = ev.getRange(2, 1, last - 1, EV_COLS.length).getValues().map(rowToEvent);
  const byId = {}, order = [];
  evs.sort((a, b) => a.ts - b.ts).forEach(x => { if (!byId[x.tg_id]) { byId[x.tg_id] = emptyUser(x.tg_id); order.push(x.tg_id); } byId[x.tg_id] = applyEvent(byId[x.tg_id], x); });
  const rows = order.map(id => { const r = byId[id]; r[U['Статус']] = statusOf(r); return r; });
  us.getRange(2, 1, rows.length, US_COLS.length).setValues(rows);
}
/* Обратное преобразование строки «Событий» в событие (для пересборки) */
function rowToEvent(r) {
  const inv = o => { const m = {}; Object.keys(o).forEach(k => m[o[k]] = k); return m; };
  const props = r[12] ? JSON.parse(r[12]) : {};
  if (r[5]) props.tariff = inv(TARIFF)[r[5]] || r[5];
  if (r[6] !== '') props.amount = r[6];
  if (r[7] !== '') props.month = r[7];
  if (r[8]) { if (r[3] === 'section_open') props.section = r[8]; else props.topic = inv(TOPIC)[r[8]] || r[8]; }
  if (r[9]) props.period = r[9];
  if (r[10]) props.locked = true;
  if (r[11]) props.from = r[11];
  return {ts: asDate(r[0]).getTime(), tg_id: String(r[1]), username: String(r[2] || '').replace(/^@/, ''), ev: r[3], props, source: r[13], demo: r[14] === 1};
}
/* Проверка без сервера: Выполнить → testPost. Добавит одно демо-событие */
function testPost() {
  const s = PropertiesService.getScriptProperties().getProperty('SECRET');
  const r = doPost({postData: {contents: JSON.stringify({secret: s, demo: true, events: [{ts: Date.now(), tg_id: '1', username: 'test', ev: 'app_open', props: {lang: 'ru'}, source: 'app', demo: true}]})}});
  Logger.log(r.getContent());
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Прогностика')
    .addItem('Пересобрать «Пользователи»', 'rebuildUsers')
    .addItem('Удалить демо-данные', 'clearDemo')
    .addToUi();
}

/* для проверки в Node (в Apps Script не используется) */
if (typeof module !== 'undefined') module.exports = {applyEvent, statusOf, emptyUser, eventRow, mergeTariff, US_COLS, EV_COLS, U, LABEL, STATUS, TARIFF};
