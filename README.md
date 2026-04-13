# OutreachAI — Local Email Outreach Agent

### A terminal-based AI agent for personalised cold email outreach

---

## How it works

1. **Reads** your client spreadsheet (CSV or Excel)
2. **Calls Claude AI** to write a personalised email for each company
3. **Shows you the draft** in the terminal — subject, body, and your signature
4. You press **`[o]`** → Gmail compose opens in your browser, pre-filled with everything
5. You review (and optionally edit) inside Gmail, then click **Send**
6. Back in terminal, press **`[y]`** to confirm it was sent
7. **Spreadsheet updates** automatically with Sent / Skipped + timestamp

> No passwords stored. No SMTP config. Gmail handles the actual sending — you stay in control.

---

## Setup (~3 minutes)

### Step 1 — Install Python dependencies

Open Terminal and run:

```bash
pip3 install anthropic openpyxl
```

- `anthropic` — Claude AI SDK (for email personalisation)
- `openpyxl` — Excel file support (skip if you use CSV)

---

### Step 2 — Get your Anthropic API key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up / log in → **API Keys** → **Create Key**
3. Copy the key (starts with `sk-ant-...`)

Cost: ~$0.002 per email. Extremely cheap for 100 clients.

---

### Step 3 — Configure the script

Open `outreach_agent.py` in any text editor and fill in the `CONFIG` section at the top:

```python
CONFIG = {
    "spreadsheet_path": "clients.csv",   # ← your file name

    "sender_name":     "Your Name",
    "sender_title":    "Digital Marketing Specialist",
    "sender_company":  "Your Company",
    "sender_website":  "yourwebsite.com",
    "booking_link":    "https://calendly.com/yourname",
    "booking_label":   "Book a free 30-min consultation",

    "gmail_address":   "you@gmail.com",   # ← just your address, no password needed

    "anthropic_api_key": "sk-ant-...",    # ← paste your key here
    ...
}
```

Also customise the **email template** in CONFIG to match your service and tone.

---

### Step 4 — Prepare your spreadsheet

Your CSV or Excel file needs at least these columns (names are flexible):

| company        | email                 | website          |
| -------------- | --------------------- | ---------------- |
| Acme Coffee    | hello@acmecoffee.com  | acmecoffee.com   |
| Blue Sky Salon | info@blueskysalon.com |                  |
| TechStart Pty  | contact@techstart.com | techstart.com.au |

Save it as `clients.csv` (or update `spreadsheet_path` in CONFIG to match your filename).

---

## Running the agent

Put `outreach_agent.py` and your spreadsheet in the same folder, then:

```bash
cd /path/to/your/folder
python3 outreach_agent.py
```

---

## During the run — your options per client

```
  [o]  Open in Gmail       ← opens browser with email pre-filled
  [y]  Confirm sent        ← after you clicked Send inside Gmail
  [n]  Skip this client
  [e]  Edit subject or body before opening
  [q]  Quit and save progress
```

**Typical flow per email:**

1. Read the preview in terminal
2. Press `o` → Gmail opens in browser
3. Check it looks good, click **Send** inside Gmail
4. Switch back to terminal, press `y`
5. Move to next client

---

## Email template placeholders

| Placeholder            | Replaced with                                 |
| ---------------------- | --------------------------------------------- |
| `{{company_name}}`   | Company name from spreadsheet                 |
| `{{website}}`        | Website from spreadsheet                      |
| `{{ai_insight}}`     | AI-generated insight specific to that company |
| `{{sender_name}}`    | Your name from CONFIG                         |
| `{{sender_title}}`   | Your title from CONFIG                        |
| `{{sender_company}}` | Your company from CONFIG                      |
| `{{booking_link}}`   | Your Calendly / Google Meet URL from CONFIG   |

---

## Tips

- **Start with 5 clients** to verify everything looks right before running the full list
- Press `e` to edit an email before opening it in Gmail if you want to tweak the AI draft
- Press `q` to quit safely at any time — progress is saved after every client
- Already-sent rows are automatically skipped on the next run
- The 2-second delay between emails is configurable in CONFIG
