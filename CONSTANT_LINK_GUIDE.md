# Keeping One Constant Student Link (While Running Locally on Your Laptop)

Previously, `AppUpdate.py` printed `http://<laptop-LAN-IP>:5000`. That link changes every time you switch Wi-Fi networks, restart your hotspot, or move buildings — and students on mobile data (4G/5G) cannot open a local IP at all.

Because `AppUpdate.py` still runs **100% on your laptop** (saving `registrations.jsonl` and `registrations.xlsx` directly onto your drive), you can lock **one permanent link** in front of your laptop using any of the 4 options below.

---

## Option 1 (Best Overall): Tailscale Funnel — Free, Permanent HTTPS, Zero Interstitials
- **Why it wins:** Free permanent `https://<your-laptop>.<tailnet>.ts.net` URL that **never changes**, works on **both campus Wi-Fi and 4G/5G**, shows **no warning page**, and keeps all Excel/JSONL files locally on your laptop.
- **Setup (once):**
  1. Install Tailscale from [tailscale.com](https://tailscale.com/download) and sign in.
  2. Enable **Funnel** in your Tailscale Admin Console (`DNS` -> enable HTTPS certificates + Funnel ACL if prompted).
  3. Run once in Terminal:
     ```bash
     tailscale funnel 5000
     ```
  4. Tailscale prints your permanent HTTPS URL (for example `https://abhigyan-air.tail1234.ts.net`).
  5. Paste it into `.env` next to `AppUpdate.py`:
     ```env
     PUBLIC_URL=https://abhigyan-air.tail1234.ts.net
     ```
  6. Start `AppUpdate.py` (or double-click `start_mac.command`). The Admin Console's **Student Link & QR** standee will automatically encode your permanent link!

---

## Option 2: ngrok Free Static Domain — Quick 1-Command Setup
- **Why it works:** Every free ngrok account includes **1 static domain** (`*.ngrok-free.app`) that never changes across restarts.
- **Setup (once):**
  1. Create a free account at [ngrok.com](https://ngrok.com), claim your free static domain in the dashboard, and run `ngrok config add-authtoken <TOKEN>`.
  2. Whenever you run the photoshoot desk:
     ```bash
     ngrok http --domain=your-static-name.ngrok-free.app 5000
     ```
  3. Set `PUBLIC_URL=https://your-static-name.ngrok-free.app` in `.env` (or `AppUpdate.py` will also auto-detect active ngrok tunnels on `127.0.0.1:4040`).
- **Note:** Free ngrok shows a one-tap *"Visit Site"* interstitial the first time a student opens it on a phone browser.

---

## Option 3: Cloudflare Named Tunnel (If You Own a Domain)
- **Why it works:** Point a subdomain like `photoshoot.yearsatflame.com` directly to `localhost:5000` with zero warning pages and free HTTPS.
- **Setup:**
  1. `brew install cloudflared` && `cloudflared tunnel login`
  2. Create a named tunnel (`cloudflared tunnel create flame-shoot`), route DNS (`cloudflared tunnel route dns flame-shoot photoshoot.yourdomain.com`), and run `cloudflared tunnel run --url http://localhost:5000 flame-shoot`.
  3. Set `PUBLIC_URL=https://photoshoot.yourdomain.com` in `.env`.

---

## Option 4 (Offline / Dedicated Hotspot): mDNS `.local` Hostname
- **Why it works:** If the photoshoot desk has no internet, use a **dedicated hotspot** or same Wi-Fi network. On macOS, your laptop already broadcasts a stable hostname:
  - Example: `http://Admins-MacBook-Air.local:5000`
- `AppUpdate.py` automatically detects this `.local` address and uses it instead of a shifting IP when no public tunnel is active.
- **Caveat:** Some restricted campus enterprise Wi-Fi networks block mDNS/client-to-client traffic — if that happens, use a personal phone hotspot or Options 1–3.

---

## Day-of-Shoot Checklist
1. Open `.env` (copied from `.env.example`) and set `ADMIN_PASSWORD=your-desk-password` so you don't have to re-type a random password if you restart the server.
2. Start your tunnel (Option 1/2/3) if using public HTTPS.
3. Double-click `start_mac.command` (macOS) or `start_windows.bat` (Windows).
4. Open `http://localhost:5000/admin` -> click **Student Link & QR** -> **Print QR Standee** for the registration table.
5. **Security tip:** Don't share `/admin` or your password; only share the root student link (`/`).
