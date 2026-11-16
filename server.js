// Punjab Mart AI: website + inventory assistant server
// Run: npm install && npm start   (needs ANTHROPIC_API_KEY in .env or the environment)
import 'dotenv/config';
import express from 'express';
import ExcelJS from 'exceljs';
import Anthropic from '@anthropic-ai/sdk';
import fs from 'node:fs';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;
const MODEL = process.env.ANTHROPIC_MODEL || 'claude-haiku-5-5';
const INVENTORY_FILE = process.env.INVENTORY_FILE || path.join(__dirname, 'data', 'inventory.xlsx');
const STORE = {
  name: 'Punjab Mart Super Store',
  address: '4906 30 Ave, Beaumont, AB T4X 1T9',
  hours: 'Open 7 days a week, closes 9 PM',
  phone: '(780) 737-0424',
};

const anthropic = process.env.ANTHROPIC_API_KEY ? new Anthropic() : null;

// WhatsApp + Instagram (all optional: each channel switches on when its settings are present)
const GRAPH_VERSION = process.env.GRAPH_VERSION || 'v25.0';
const FB_HOST = process.env.FB_GRAPH_HOST || 'https://graph.facebook.com';
const IG_HOST = process.env.IG_GRAPH_HOST || 'https://graph.instagram.com';
const WA = {
  token: process.env.WA_TOKEN, phoneId: process.env.WA_PHONE_NUMBER_ID,
  publicNumber: (process.env.WA_PUBLIC_NUMBER || '').replace(/\D/g, ''),
  alertTo: (process.env.STORE_ALERT_NUMBER || '').replace(/\D/g, ''),
  alertTemplate: process.env.WA_ALERT_TEMPLATE, alertLang: process.env.WA_ALERT_TEMPLATE_LANG || 'en',
};
WA.on = !!(WA.token && WA.phoneId);
const IG = { token: process.env.IG_TOKEN, accountId: process.env.IG_ACCOUNT_ID, username: process.env.IG_USERNAME || 'punjabmart_beaumont1' };
IG.on = !!(IG.token && IG.accountId);
const VERIFY_TOKEN = process.env.META_VERIFY_TOKEN || '';
const APP_SECRETS = [process.env.META_APP_SECRET, process.env.IG_APP_SECRET].filter(Boolean);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const ORDERS_FILE = path.join(__dirname, 'data', 'orders.jsonl');

/* ---------------- inventory (reads the Excel file, reloads when it changes) ---------------- */
let INV = [];
let BY = {};
let loadedAt = 0;

const COLS = {
  id: 'ID', name: 'Product (English)', pa: 'Punjabi Name', hi: 'Hindi Name', tags: 'Search Keywords',
  aisle: 'Aisle', emoji: 'Emoji', price: 'Price (CAD)', unit: 'Unit', stock: 'Stock Qty',
  fresh: 'Freshness / Best Before', active: 'Active',
};

function cellText(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if ('result' in v) return cellText(v.result);
    if ('richText' in v) return v.richText.map(r => r.text).join('');
    if ('text' in v) return String(v.text);
  }
  return String(v).trim();
}

async function loadInventory() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(INVENTORY_FILE);
  const ws = wb.getWorksheet('Inventory') || wb.worksheets[0];
  const header = {};
  ws.getRow(1).eachCell((c, i) => { header[cellText(c.value)] = i; });
  const items = [];
  ws.eachRow((row, r) => {
    if (r === 1) return;
    const get = k => cellText(row.getCell(header[COLS[k]] || 0).value);
    const id = get('id');
    if (!id || get('active').toLowerCase() === 'no') return;
    items.push({
      id, name: get('name'), pa: get('pa'), hi: get('hi'), tags: get('tags'), aisle: get('aisle'),
      emoji: get('emoji') || '🛒', price: Number(get('price')) || 0, unit: get('unit'),
      stock: Math.max(0, Math.round(Number(get('stock')) || 0)), fresh: get('fresh'),
    });
  });
  INV = items;
  BY = Object.fromEntries(items.map(i => [i.id, i]));
  loadedAt = Date.now();
  console.log(`[inventory] loaded ${items.length} items from ${path.basename(INVENTORY_FILE)}`);
}

let lastMtime = 0;
async function refreshIfChanged() {
  try {
    const m = fs.statSync(INVENTORY_FILE).mtimeMs;
    if (m !== lastMtime) { lastMtime = m; await loadInventory(); }
  } catch (e) { console.error('[inventory] could not read file:', e.message); }
}
setInterval(refreshIfChanged, 30_000);

/* ---------------- search ---------------- */
const norm = s => (s || '').toLowerCase().normalize('NFC');
const splitWords = s => s.split(/[^\p{L}\p{M}\p{N}]+/u).filter(Boolean);

function search(query) {
  const qn = norm(query);
  const words = splitWords(qn).filter(w => w.length > 1);
  return INV.map(i => {
    const hay = norm([i.name, i.pa, i.hi, i.tags].join(' '));
    const hayWords = splitWords(hay);
    let s = 0;
    if (qn.length > 2 && hay.includes(qn)) s += 5;
    [i.pa, i.hi].forEach(n => n.split(/[\s/()]+/).filter(p => p.length > 1).forEach(p => { if (qn.includes(norm(p))) s += 3; }));
    words.forEach(w => { if (w.length > 2 && hayWords.some(h => h === w || (w.length > 3 && h.startsWith(w)))) s += 2; });
    return { i, s };
  }).filter(x => x.s > 0).sort((a, b) => b.s - a.s).slice(0, 8).map(x => x.i);
}

const brief = i => ({
  id: i.id, name: i.name, punjabi: i.pa, hindi: i.hi, aisle: i.aisle,
  price: `$${i.price.toFixed(2)} per ${i.unit}`, in_stock: i.stock, freshness: i.fresh,
});

/* ---------------- the agent ---------------- */
const SYSTEM = `You are "Sardar Ji", the warm, friendly shop helper for ${STORE.name}, an Indian & Punjabi grocery at ${STORE.address} (${STORE.hours}, phone ${STORE.phone}).
Rules:
- Reply in the SAME language and script the customer used (Punjabi in Gurmukhi, Hindi in Devanagari, English, Hinglish, Urdu, Gujarati, French or anything else). Keep replies short and warm: 1 to 3 sentences, like a friendly Punjabi uncle at the counter (an occasional "ji" or "haanji" is nice). Never use em dashes. You may use **bold** for item names.
- ALWAYS call search_inventory before saying whether something is available. Search with simple English/romanized product words (e.g. "methi", "masoor dal", "paneer"). For recipes or occasions, search each main ingredient.
- Never invent products, prices, stock or freshness. Only use tool results. If an item is out of stock, mention the freshness note (it often says the next delivery) and suggest an in-stock alternative.
- After finding relevant items, call show_products with their ids so the customer sees cards (stock, aisle, price, freshness). Don't repeat prices in text when cards are shown.
- If the customer asks to add items, or wants a list for a recipe or party, call add_to_list. Scale quantities sensibly with the number of people.
- If something isn't stocked at all, say so politely and offer to note a request for the team.
- Only help with shopping at this store. Politely decline unrelated requests.`;

const CHANNEL_NOTE = `
You are chatting on {CHANNEL}. The customer sees plain text only, no product cards, so after searching mention the key items as short lines (name, price, in stock or not, aisle). Keep it under 6 lines.
To take a pickup order: add items with add_to_list, read the list back with quantities and ask "Shall I confirm this order for pickup?". Only after the customer clearly says yes, call place_order. Then share the pickup code and say the team will have it ready. Payment happens in store at pickup.`;

const PLACE_ORDER_TOOL = {
  name: 'place_order',
  description: 'Confirm the customer\'s current pickup list as an order. Only call after the customer explicitly confirms. Returns a pickup code.',
  input_schema: { type: 'object', properties: { note: { type: 'string', description: 'Optional note from the customer, e.g. pickup time' } } },
};

const TOOLS = [
  {
    name: 'search_inventory',
    description: 'Search the store\'s live inventory. Returns matching items with id, names, aisle, price, units in stock and freshness. Use romanized/English product words.',
    input_schema: { type: 'object', properties: { query: { type: 'string' } }, required: ['query'] },
  },
  {
    name: 'show_products',
    description: 'Show product cards to the customer for the given item ids (max 8).',
    input_schema: { type: 'object', properties: { ids: { type: 'array', items: { type: 'string' } } }, required: ['ids'] },
  },
  {
    name: 'add_to_list',
    description: 'Add items to the customer\'s pickup list.',
    input_schema: {
      type: 'object',
      properties: { items: { type: 'array', items: { type: 'object', properties: { id: { type: 'string' }, qty: { type: 'number' } }, required: ['id'] } } },
      required: ['items'],
    },
  },
];

async function runAgent(history, ctx = {}) {
  const show = new Set();
  const add = [];
  const messages = history.map(m => ({ role: m.role, content: String(m.content).slice(0, 2000) }));

  for (let step = 0; step < 6; step++) {
    const system = ctx.channel ? SYSTEM + CHANNEL_NOTE.replace('{CHANNEL}', ctx.channel === 'whatsapp' ? 'WhatsApp' : 'Instagram DMs') : SYSTEM;
    const tools = ctx.channel ? [...TOOLS, PLACE_ORDER_TOOL] : TOOLS;
    const res = await anthropic.messages.create({ model: MODEL, max_tokens: 700, system, tools, messages });
    if (res.stop_reason !== 'tool_use') {
      const reply = res.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
      return { reply, products: [...show].map(id => BY[id]).filter(Boolean), add };
    }
    messages.push({ role: 'assistant', content: res.content });
    const results = res.content.filter(b => b.type === 'tool_use').map(b => {
      let out;
      if (b.name === 'search_inventory') {
        const r = search(b.input.query || '');
        out = r.length ? r.map(brief) : { results: [], note: 'No matching items in inventory' };
      } else if (b.name === 'show_products') {
        (b.input.ids || []).slice(0, 8).forEach(id => BY[id] && show.add(id));
        out = 'shown';
      } else if (b.name === 'add_to_list') {
        const added = [];
        (b.input.items || []).forEach(x => {
          const it = BY[x.id];
          if (it && it.stock > 0) {
            const qty = Math.max(1, Math.round(x.qty || 1));
            add.push({ id: x.id, qty }); added.push(x.id); show.add(x.id);
            if (ctx.session) ctx.session.cart[x.id] = (ctx.session.cart[x.id] || 0) + qty;
          }
        });
        out = { added, current_list: ctx.session ? listSummary(ctx.session.cart) : undefined };
      } else if (b.name === 'place_order' && ctx.session) {
        const items = Object.entries(ctx.session.cart).map(([id, qty]) => ({ id, qty }));
        if (!items.length) out = { error: 'The pickup list is empty. Add items first.' };
        else {
          const order = createOrder({ channel: ctx.channel, name: ctx.session.name, phone: ctx.session.phone, userId: ctx.session.userId, items, note: b.input.note });
          ctx.session.cart = {};
          out = { ok: true, pickup_code: order.ref, total_estimate: '$' + order.total.toFixed(2) };
        }
      } else out = { error: 'unknown tool' };
      return { type: 'tool_result', tool_use_id: b.id, content: JSON.stringify(out) };
    });
    messages.push({ role: 'user', content: results });
  }
  return { reply: 'Sorry, that took too long. Could you ask again in a simpler way?', products: [...show].map(id => BY[id]).filter(Boolean), add };
}

/* ---------------- orders ---------------- */
let ORDERS = [];
try { ORDERS = fs.readFileSync(ORDERS_FILE, 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l)); } catch {}
const saveOrders = () => fs.writeFileSync(ORDERS_FILE, ORDERS.map(o => JSON.stringify(o)).join('\n') + (ORDERS.length ? '\n' : ''));
function listSummary(cart) { return Object.entries(cart).map(([id, qty]) => BY[id] ? `${qty} x ${BY[id].name}` : null).filter(Boolean); }

function createOrder({ channel, name, phone, userId, items, note }) {
  const clean = items.filter(x => BY[x.id]).map(x => ({ id: x.id, name: BY[x.id].name, qty: Math.max(1, Math.round(x.qty || 1)), price: BY[x.id].price }));
  const order = {
    ref: 'PM-' + Date.now().toString(36).toUpperCase().slice(-6), at: new Date().toISOString(), status: 'new',
    channel, name: String(name || '').slice(0, 80), phone: String(phone || '').slice(0, 30), userId: userId || null,
    note: String(note || '').slice(0, 200), items: clean, total: clean.reduce((t, x) => t + x.price * x.qty, 0),
  };
  ORDERS.push(order); saveOrders();
  console.log('[order]', order.ref, channel, clean.length, 'items');
  alertStore(order).catch(e => console.error('[alert]', e.message));
  return order;
}

/* ---------------- WhatsApp + Instagram senders ---------------- */
async function graph(url, body, token) {
  const r = await fetch(url, { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { const e = new Error(j?.error?.message || `HTTP ${r.status}`); e.code = j?.error?.code; throw e; }
  return j;
}
const waSendText = (to, text) => graph(`${FB_HOST}/${GRAPH_VERSION}/${WA.phoneId}/messages`,
  { messaging_product: 'whatsapp', recipient_type: 'individual', to, type: 'text', text: { preview_url: false, body: text.slice(0, 4000) } }, WA.token);
const waSendTemplate = (to, name, lang, params) => graph(`${FB_HOST}/${GRAPH_VERSION}/${WA.phoneId}/messages`,
  { messaging_product: 'whatsapp', recipient_type: 'individual', to, type: 'template',
    template: { name, language: { code: lang }, components: [{ type: 'body', parameters: params.map(t => ({ type: 'text', text: t })) }] } }, WA.token);
function igSendText(igsid, text) {
  let t = text; while (Buffer.byteLength(t) > 990) t = t.slice(0, -20); // Instagram limit: 1000 bytes
  return graph(`${IG_HOST}/${GRAPH_VERSION}/${IG.accountId}/messages`, { recipient: { id: igsid }, message: { text: t } }, IG.token);
}

// tell the store team about every new order on WhatsApp
async function alertStore(order) {
  if (!WA.on || !WA.alertTo) return;
  const via = { website: 'website', whatsapp: 'WhatsApp', instagram: 'Instagram' }[order.channel] || order.channel;
  const who = [order.name, order.phone].filter(Boolean).join(', ') || 'Customer';
  const lines = order.items.map(x => `• ${x.qty} x ${x.name}`).join('\n');
  const text = `🛒 New pickup order ${order.ref}\nVia ${via}: ${who}\n${lines}\nEstimated total: $${order.total.toFixed(2)}${order.note ? '\nNote: ' + order.note : ''}`;
  try { await waSendText(WA.alertTo, text); }
  catch (e) {
    // outside the 24-hour window WhatsApp only allows approved templates
    if (!WA.alertTemplate) throw e;
    const itemsOneLine = order.items.map(x => `${x.qty} x ${x.name}`).join(', ').slice(0, 900);
    await waSendTemplate(WA.alertTo, WA.alertTemplate, WA.alertLang, [order.ref, `${who} (${via})`, itemsOneLine]);
  }
}

/* ---------------- chat sessions for WhatsApp / Instagram ---------------- */
const sessions = new Map(); // key -> { history, cart, name, phone, userId, greeted, last }
const seen = new Set();
const channelHits = new Map();
function getSession(key, init) {
  let s = sessions.get(key);
  if (!s || Date.now() - s.last > 24 * 3600_000) { s = { history: [], cart: {}, greeted: false, ...init }; sessions.set(key, s); }
  s.last = Date.now(); return s;
}
setInterval(() => { const cut = Date.now() - 24 * 3600_000; for (const [k, v] of sessions) if (v.last < cut) sessions.delete(k); }, 3600_000);

const GREETING = 'ਸਤਿ ਸ੍ਰੀ ਅਕਾਲ ਜੀ! 🙏 I\'m Sardar Ji, Punjab Mart\'s automated shop helper.';
function productLines(products) {
  return products.slice(0, 5).map(p => `• ${p.name}: $${p.price.toFixed(2)}/${p.unit}, ${p.stock > 0 ? (p.stock < 6 ? `only ${p.stock} left` : 'in stock') : 'out of stock'} (${p.aisle})`).join('\n');
}

async function handleChannelMessage(channel, userId, name, text, phone) {
  if (!seenRate(channel + userId)) return;
  const send = channel === 'whatsapp' ? (t => waSendText(userId, t)) : (t => igSendText(userId, t));
  const s = getSession(channel + ':' + userId, { name, phone, userId });
  if (!anthropic) { await send(`${GREETING} Our assistant is resting right now. Please call ${STORE.phone} and we'll help you.`); return; }
  s.history.push({ role: 'user', content: text.slice(0, 1500) });
  s.history = s.history.slice(-14);
  try {
    await refreshIfChanged();
    const r = await runAgent(s.history, { channel, session: s });
    let reply = r.reply || 'Haanji, one moment please.';
    s.history.push({ role: 'assistant', content: reply });
    if (r.products.length && !/PM-[A-Z0-9]{4,}/.test(reply)) reply += '\n\n' + productLines(r.products);
    if (!s.greeted) { reply = GREETING + '\n\n' + reply; s.greeted = true; }
    await send(reply);
  } catch (e) {
    console.error(`[${channel}]`, e.message);
    await send(`Sorry ji, something went wrong on our side. Please try again or call ${STORE.phone}.`).catch(() => {});
  }
}
function seenRate(key) {
  const now = Date.now(); const arr = (channelHits.get(key) || []).filter(t => now - t < 5 * 60_000);
  arr.push(now); channelHits.set(key, arr); return arr.length <= 30;
}

/* ---------------- web server ---------------- */
const app = express();
app.use(express.json({ limit: '256kb', verify: (req, _res, buf) => { req.rawBody = buf; } }));

// light per-IP rate limit: 20 chat requests per 5 minutes
const hits = new Map();
function rateLimited(ip) {
  const now = Date.now();
  const arr = (hits.get(ip) || []).filter(t => now - t < 5 * 60_000);
  arr.push(now); hits.set(ip, arr);
  return arr.length > 20;
}

app.get('/api/inventory', async (_req, res) => {
  await refreshIfChanged();
  res.set('Cache-Control', 'no-store').json({ items: INV, loadedAt, ai: !!anthropic });
});

app.post('/api/chat', async (req, res) => {
  if (!anthropic) return res.status(503).json({ error: 'ai_not_configured' });
  if (rateLimited(req.ip)) return res.status(429).json({ error: 'rate_limited' });
  const history = Array.isArray(req.body?.messages) ? req.body.messages.slice(-12) : [];
  if (!history.length || history[history.length - 1].role !== 'user') return res.status(400).json({ error: 'bad_request' });
  try {
    await refreshIfChanged();
    res.json(await runAgent(history));
  } catch (e) {
    console.error('[chat]', e.status || '', e.message);
    res.status(502).json({ error: 'ai_error' });
  }
});

// pickup reservations from the website
app.post('/api/reserve', (req, res) => {
  const items = (req.body?.items || []).filter(x => BY[x.id]);
  if (!items.length) return res.status(400).json({ error: 'empty' });
  const order = createOrder({ channel: 'website', name: req.body?.name, phone: req.body?.phone, items });
  res.json({ ok: true, ref: order.ref });
});

// public settings for the website (which chat apps are switched on)
app.get('/api/config', (_req, res) => res.json({
  whatsapp: WA.on && WA.publicNumber ? WA.publicNumber : null,
  instagram: IG.username || null,
}));

/* ---------------- Meta webhooks (WhatsApp + Instagram) ---------------- */
app.get('/webhook', (req, res) => {
  if (req.query['hub.mode'] === 'subscribe' && VERIFY_TOKEN && req.query['hub.verify_token'] === VERIFY_TOKEN) return res.status(200).send(req.query['hub.challenge']);
  res.sendStatus(403);
});

function validSignature(req) {
  if (!APP_SECRETS.length) return true; // not configured: accept (set META_APP_SECRET in production)
  const sig = req.get('x-hub-signature-256') || '';
  return APP_SECRETS.some(secret => {
    const expected = 'sha256=' + crypto.createHmac('sha256', secret).update(req.rawBody || '').digest('hex');
    return sig.length === expected.length && crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected));
  });
}

app.post('/webhook', (req, res) => {
  if (!validSignature(req)) return res.sendStatus(401);
  res.sendStatus(200); // answer Meta right away, then work
  const body = req.body || {};
  try {
    if (body.object === 'whatsapp_business_account' && WA.on) {
      for (const entry of body.entry || []) for (const ch of entry.changes || []) {
        const v = ch.value || {};
        const names = Object.fromEntries((v.contacts || []).map(c => [c.wa_id, c.profile?.name]));
        for (const m of v.messages || []) {
          if (seen.has(m.id)) continue; seen.add(m.id); if (seen.size > 5000) seen.clear();
          const text = m.type === 'text' ? m.text?.body : m.type === 'interactive' ? (m.interactive?.button_reply?.title || m.interactive?.list_reply?.title) : m.type === 'button' ? m.button?.text : null;
          if (!text) { waSendText(m.from, 'Haanji! I can read text messages only for now. Please type what you are looking for 🙏').catch(() => {}); continue; }
          handleChannelMessage('whatsapp', m.from, names[m.from] || '', text, '+' + m.from);
        }
      }
    } else if (body.object === 'instagram' && IG.on) {
      for (const entry of body.entry || []) for (const ev of entry.messaging || []) {
        const m = ev.message; if (!m || m.is_echo || m.is_deleted) continue;
        if (seen.has(m.mid)) continue; seen.add(m.mid);
        if (!m.text) { igSendText(ev.sender.id, 'Haanji! Please type what you are looking for and I\'ll check the shelf 🙏').catch(() => {}); continue; }
        handleChannelMessage('instagram', ev.sender.id, '', m.text, '');
      }
    }
  } catch (e) { console.error('[webhook]', e.message); }
});

/* ---------------- orders screen for the store team ---------------- */
function adminAuth(req, res, next) {
  if (!ADMIN_PASSWORD) return res.status(503).send('Set ADMIN_PASSWORD in .env to use the orders screen.');
  const [, b64] = (req.get('authorization') || '').split(' ');
  const pass = b64 ? Buffer.from(b64, 'base64').toString().split(':').slice(1).join(':') : '';
  if (pass === ADMIN_PASSWORD) return next();
  res.set('WWW-Authenticate', 'Basic realm="Punjab Mart orders"').status(401).send('Login required');
}
app.get('/api/orders', adminAuth, (_req, res) => res.set('Cache-Control', 'no-store').json({ orders: [...ORDERS].reverse().slice(0, 200), wa: WA.on, ig: IG.on }));
app.post('/api/orders/:ref/status', adminAuth, async (req, res) => {
  const o = ORDERS.find(x => x.ref === req.params.ref);
  const status = req.body?.status;
  if (!o || !['new', 'ready', 'picked', 'cancelled'].includes(status)) return res.status(400).json({ error: 'bad_request' });
  o.status = status; saveOrders();
  let notified = false, notifyError = null;
  if (status === 'ready') {
    const msg = `Haanji ${o.name || ''}! 🙏 Your Punjab Mart order ${o.ref} is packed and ready for pickup at 4906 30 Ave, Beaumont. See you soon!`.replace('  ', ' ');
    try {
      if (o.channel === 'whatsapp' && WA.on && o.userId) { await waSendText(o.userId, msg); notified = true; }
      else if (o.channel === 'instagram' && IG.on && o.userId) { await igSendText(o.userId, msg); notified = true; }
    } catch (e) { notifyError = e.message; }
  }
  res.json({ ok: true, notified, notifyError });
});
app.get('/orders', adminAuth, (_req, res) => res.sendFile(path.join(__dirname, 'admin', 'orders.html')));

app.use(express.static(path.join(__dirname, 'public'), { extensions: ['html'] }));

await refreshIfChanged();
app.listen(PORT, () => {
  console.log(`Punjab Mart site running on http://localhost:${PORT}`);
  console.log(anthropic ? `[ai] live, model ${MODEL}` : '[ai] ANTHROPIC_API_KEY not set: assistant runs in offline mode');
  console.log(`[whatsapp] ${WA.on ? 'on' : 'off'}${WA.on && WA.alertTo ? ', order alerts to +' + WA.alertTo : ''} · [instagram] ${IG.on ? 'on' : 'off'} · [orders screen] ${ADMIN_PASSWORD ? '/orders' : 'off (set ADMIN_PASSWORD)'}`);
});
