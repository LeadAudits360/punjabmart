# Punjab Mart: website + AI inventory assistant

An animated website for Punjab Mart Super Store (Beaumont, AB) with a multilingual AI shop assistant. Customers ask in Punjabi, Hindi, English or any mix. The assistant checks the inventory sheet and answers with stock, aisle, price and freshness. It also builds pickup lists and takes pickup reservations.

## What's inside

```
server.js             Node server: website, AI chat, WhatsApp + Instagram webhooks, orders
admin/orders.html     Orders screen for the store team (yoursite.com/orders)
public/index.html     The website (all animations, the chat, the pickup list)
public/logo.webp      Punjab Mart logo
data/inventory.xlsx   The inventory the AI reads (60 sample products)
data/orders.jsonl     Created automatically: one line per pickup reservation
.env.example          Settings template
```

## Run it locally (5 minutes)

1. Install Node.js 18 or newer.
2. In this folder:
   ```bash
   npm install
   cp .env.example .env      # then paste your Anthropic API key into .env
   npm start
   ```
3. Open http://localhost:3000

Get an API key at https://console.anthropic.com (API keys). The key stays on the server and is never sent to the browser.

## Put it on your subdomain

Any host that runs Node.js works. Pick whichever you already use:

- **cPanel hosting with "Setup Node.js App"**: create an app, set the application root to this folder, startup file `server.js`, run `npm install` from the panel, add `ANTHROPIC_API_KEY` under environment variables, then point the subdomain at the app.
- **VPS (DigitalOcean, Lightsail, etc.)**: upload the folder, `npm install`, then run it with PM2 (`npm i -g pm2 && pm2 start server.js --name punjabmart`). Put Nginx in front for the subdomain and HTTPS.
- **Render / Railway**: create a new Web Service from this folder (or a Git repo), build command `npm install`, start command `npm start`, and add the `ANTHROPIC_API_KEY` environment variable. Then add your subdomain as a custom domain.

Only have static hosting? Upload just the `public/` folder. The site and chat still work in offline mode using a built-in copy of the sample inventory, but without live AI, Excel updates or reservations.

## Updating the inventory

Open `data/inventory.xlsx`, edit it, save it, and replace the file on the server. The server notices the change within 30 seconds, with no restart needed.

- **Stock Qty** 0 = out of stock; under 6 shows "Only X left".
- **Freshness / Best Before** is shown on each card, e.g. "Arrived today" or "Next delivery Thursday".
- **Active** = No hides an item without deleting it.
- Keep the column headers exactly as they are; the server finds columns by header name.

The "How to use" tab inside the workbook explains every column.

## Settings (.env)

| Variable | Default | What it does |
|---|---|---|
| `ANTHROPIC_API_KEY` | (none) | Required for live AI. Without it, the chat runs in offline mode. |
| `ANTHROPIC_MODEL` | `claude-haiku-5-5` | Fast, low-cost model. Use `claude-sonnet-5-5` for smarter answers. |
| `PORT` | `3000` | Port the server listens on. |
| `INVENTORY_FILE` | `data/inventory.xlsx` | Path to the inventory workbook. |

## Orders

Every pickup order (from the website, WhatsApp or Instagram) is saved in `data/orders.jsonl` and shows up on the **orders screen** at `https://yoursite.com/orders`. Log in with any username and your `ADMIN_PASSWORD`.

- New orders appear automatically every 15 seconds. Tap "Turn on sound alerts" for a beep.
- **Packed · notify customer** marks the order ready and, for WhatsApp and Instagram orders, messages the customer that it's ready for pickup.
- **Picked up** and **Cancel** close the order.

If WhatsApp is set up with `STORE_ALERT_NUMBER`, every new order is also sent to that WhatsApp number.

## WhatsApp + Instagram (Sardar Ji on both apps)

Customers message the store on WhatsApp or Instagram; Sardar Ji answers from the same inventory sheet, builds a pickup list, and places the order after the customer confirms. The team gets a WhatsApp alert and sees it on the orders screen.

Your site must be live on **HTTPS** first (Meta only sends webhooks to HTTPS). Your webhook address is:

```
https://YOUR-SUBDOMAIN/webhook
```

### Step 1: Meta developer app
1. Go to developers.facebook.com, log in, click **My Apps → Create app**, pick **Business** type, and connect the store's Business portfolio (create one at business.facebook.com if needed).
2. App settings → Basic: copy the **App secret** into `META_APP_SECRET`.
3. Make up a secret phrase and put it in `META_VERIFY_TOKEN`.

### Step 2: WhatsApp
1. In the app, add the **WhatsApp** product. Add and verify the store's business phone number (a number not already on the regular WhatsApp app, or migrate it).
2. Create a **System User** (Business settings → Users → System users), give it the app, and generate a token with `whatsapp_business_messaging` and `whatsapp_business_management`. Put it in `WA_TOKEN`.
3. WhatsApp → API setup: copy the **Phone number ID** into `WA_PHONE_NUMBER_ID`. Put the public number in `WA_PUBLIC_NUMBER`.
4. WhatsApp → Configuration → Webhook: Callback URL `https://YOUR-SUBDOMAIN/webhook`, Verify token = your `META_VERIFY_TOKEN`. Click **Verify and save**, then **Manage** and subscribe to **messages**.
5. Put the owner's WhatsApp number in `STORE_ALERT_NUMBER` (digits with country code).
6. **Order alert template** (needed because WhatsApp only lets a business message someone freely within 24 hours of their last message). In WhatsApp Manager → Message templates, create a **Utility** template named `new_order_alert`, language English, body:
   ```
   New pickup order {{1}} from {{2}}. Items: {{3}}. Please check the orders screen.
   ```
   Once approved, alerts always get through. Tip: if the owner sends any message to the store's WhatsApp each morning, alerts go as normal messages that day.

### Step 3: Instagram
1. The store's Instagram must be a **Professional (Business or Creator)** account.
2. In the same app, add **Instagram → API setup with Instagram login**. Add the store's Instagram account and generate a token (permissions `instagram_business_basic`, `instagram_business_manage_messages`). Put it in `IG_TOKEN`, the account ID in `IG_ACCOUNT_ID`, and the Instagram app secret in `IG_APP_SECRET`.
3. Configure webhooks there with the same Callback URL and Verify token, and subscribe to **messages**.
4. In the Instagram app: Settings → Messages and story replies → Message controls → turn on **Allow access to messages** (connected tools).

### Step 4: Restart and test
Restart the server. The log should show `[whatsapp] on` and `[instagram] on`. Message the store from a personal phone: "Do you have masoor dal?", then "add 2", then "yes". You'll get a pickup code, the alert arrives on the owner's WhatsApp, and the order shows on `/orders`.

While the Meta app is in **Development** mode, only people with a role on the app can message it. Submit for **App Review / go Live** before customers use it.

### Good to know
- Replies must go out within 24 hours of the customer's last message on both apps. Sardar Ji replies instantly, so this only matters for "order ready" messages on very old orders.
- Sardar Ji introduces himself as an automated helper in the first message, as Meta recommends.
- WhatsApp charges per message for some categories; check current rates at business.whatsapp.com before launch.

## Next steps toward production

- Replace the sample products with Punjab Mart's real list, or sync from their POS (Square, Clover, Lightspeed and Shopify POS all have inventory APIs) instead of the Excel file.
- Send reservations to the store's WhatsApp or email.
- Connect the same assistant to WhatsApp Business and Instagram DMs.
- The server has a simple rate limit (20 chats per visitor per 5 minutes) to keep AI costs predictable.
