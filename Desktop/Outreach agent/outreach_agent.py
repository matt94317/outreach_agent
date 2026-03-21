#!/usr/bin/env python3
"""
OutreachAI — Local Email Outreach Agent
========================================
Reads your client spreadsheet, personalises emails with Claude AI,
opens a pre-filled Gmail compose window in your browser for review,
waits for your confirmation in the terminal, then marks the
spreadsheet as Sent or Skipped.

Flow per client:
  1. AI generates personalised email
  2. Terminal shows a preview
  3. Press [o] → Gmail compose opens in your browser, pre-filled
  4. Review / edit inside Gmail, then send from there
  5. Back in terminal: [y] confirm sent · [n] didn't send / skip · [e] regenerate

Usage:
    python outreach_agent.py
"""

import csv
import json
import time
import webbrowser
import urllib.parse
import textwrap
import sys
from datetime import datetime
from pathlib import Path

# ── Optional Excel support ────────────────────────────────────────────────────
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# ── Optional Anthropic SDK (falls back to raw HTTP) ───────────────────────────
try:
    import anthropic
    USE_SDK = True
except ImportError:
    import urllib.request
    USE_SDK = False

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION  —  Edit this section before running
# ══════════════════════════════════════════════════════════════════════════════

CONFIG = {
    # ── Spreadsheet ───────────────────────────────────────────────────────────
    "spreadsheet_path": "/Users/matt/Desktop/Outreach agent/Test_B2B Tracking Sheet - China copy.xlsx",   # CSV or .xlsx file
    # Column names in your spreadsheet (case-insensitive, partial match ok)
    "col_company":          "Company Name",
    "col_contact_name":     "Contact Name",
    "col_role":             "Role",
    "col_email":            "Email",            # matches "Email(s)"
    "col_linkedin":         "LinkedIn",         # matches "Linke..."
    "col_date_contacted":   "Date Last Contacted",
    "col_website":      "Website",
    "col_status":           "Status",           # Will be added/updated automatically
    "col_action_required":  "Action Required",

    # ── Your details ─────────────────────────────────────────────────────────
    "sender_name":     "Renee Yu",
    "sender_title":    "Hookd 共同創辦人",
    "sender_company":  "Hookd",
    "sender_website":  "https://hookdugc.framer.website/",
    "booking_link":    "https://calendar.app.google/2Gh1kTgRiH88sn5J6",   # Google Meet / Calendly
    "booking_label":   "Book a free 30-min consultation",

    # ── Gmail ─────────────────────────────────────────────────────────────────
    # Just your Gmail address — used to pre-fill the From field in compose URLs.
    # No password needed. Gmail opens in your browser and you send from there.
    "gmail_address":   "renee@hookdugc.com",

    # ── Anthropic API key ─────────────────────────────────────────────────────
    # Get one at console.anthropic.com
    "anthropic_api_key": "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",

    # ── Email template ────────────────────────────────────────────────────────
    # Use {{company_name}}, {{website}}, {{ai_insight}} as placeholders.
    "email_subject_template": "Helping {{company_name}} grow with Digital Marketing",

    "email_body_template": """Hi {{company_name}} team,

I came across your business and wanted to reach out personally.

{{ai_insight}}

At {{sender_company}}, we specialise in digital marketing strategies tailored for businesses like yours — from SEO and social media to paid ads and content marketing. We'd love to show you what's possible.

Would you be open to a quick 30-minute chat? You can book a time that suits you here:
{{booking_link}}

Looking forward to connecting!

Best regards,
{{sender_name}}
{{sender_title}} | {{sender_company}}
{{sender_website}}
""",

    # ── Behaviour ─────────────────────────────────────────────────────────────
    "delay_between_sends": 2,     # seconds to wait between sent emails
    "skip_already_sent":   True,  # skip rows already marked "Sent"
}

# ══════════════════════════════════════════════════════════════════════════════
#  COLOURS & TERMINAL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    DIM    = "\033[2m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    CYAN   = "\033[96m"
    WHITE  = "\033[97m"
    BLUE   = "\033[94m"
    MAGENTA= "\033[95m"

def banner():
    print(f"""
{C.CYAN}{C.BOLD}
  ╔═══════════════════════════════════════╗
  ║        OutreachAI  —  v2.0            ║
  ║   Local Email Outreach Agent 🤖       ║
  ║   Gmail Compose Mode                  ║
  ╚═══════════════════════════════════════╝
{C.RESET}""")

def section(title):
    print(f"\n{C.BOLD}{C.WHITE}{'─'*50}{C.RESET}")
    print(f"{C.BOLD}{C.CYAN}  {title}{C.RESET}")
    print(f"{C.BOLD}{C.WHITE}{'─'*50}{C.RESET}")

def info(msg):    print(f"  {C.BLUE}ℹ{C.RESET}  {msg}")
def success(msg): print(f"  {C.GREEN}✓{C.RESET}  {C.GREEN}{msg}{C.RESET}")
def warn(msg):    print(f"  {C.YELLOW}⚠{C.RESET}  {C.YELLOW}{msg}{C.RESET}")
def error(msg):   print(f"  {C.RED}✗{C.RESET}  {C.RED}{msg}{C.RESET}")
def dim(msg):     print(f"  {C.DIM}{msg}{C.RESET}")

def divider(): print(f"  {C.DIM}{'·'*46}{C.RESET}")

# ══════════════════════════════════════════════════════════════════════════════
#  SPREADSHEET  READ / WRITE
# ══════════════════════════════════════════════════════════════════════════════

def find_col(headers, keyword):
    """Find a column header that contains the keyword (case-insensitive)."""
    kw = keyword.lower()
    for h in headers:
        if kw in h.lower():
            return h
    return None

def load_clients(path):
    path = Path(path)
    if not path.exists():
        error(f"Spreadsheet not found: {path}")
        error("Please set 'spreadsheet_path' in CONFIG and make sure the file exists.")
        sys.exit(1)

    ext = path.suffix.lower()
    rows = []
    headers = []

    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames or []
            rows = list(reader)

    elif ext in (".xlsx", ".xls"):
        if not EXCEL_SUPPORT:
            error("openpyxl not installed. Run: pip install openpyxl")
            sys.exit(1)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        raw = list(ws.values)
        if not raw:
            error("Spreadsheet appears to be empty.")
            sys.exit(1)
        headers = [str(h) if h is not None else "" for h in raw[0]]
        for row in raw[1:]:
            rows.append({headers[i]: (str(v) if v is not None else "") for i, v in enumerate(row)})
    else:
        error(f"Unsupported file type: {ext}. Use .csv or .xlsx")
        sys.exit(1)

    # Normalise column mapping
    cfg = CONFIG
    col_company        = find_col(headers, cfg["col_company"]) or (headers[0] if headers else "company")
    col_contact_name   = find_col(headers, cfg["col_contact_name"])
    col_role           = find_col(headers, cfg["col_role"])
    col_email          = find_col(headers, cfg["col_email"])
    col_linkedin       = find_col(headers, cfg["col_linkedin"])
    col_date_contacted = find_col(headers, cfg["col_date_contacted"])
    col_description    = find_col(headers, cfg["col_website"])
    col_status         = find_col(headers, cfg["col_status"])
    col_action         = find_col(headers, cfg["col_action_required"])

    clients = []
    for i, row in enumerate(rows):
        company      = row.get(col_company, "").strip()
        email        = row.get(col_email, "").strip()          if col_email          else ""
        contact_name = row.get(col_contact_name, "").strip()   if col_contact_name   else ""
        role         = row.get(col_role, "").strip()           if col_role           else ""
        linkedin     = row.get(col_linkedin, "").strip()       if col_linkedin       else ""
        date_contacted = row.get(col_date_contacted, "").strip() if col_date_contacted else ""
        description  = row.get(col_description, "").strip()   if col_description    else ""
        status       = row.get(col_status, "").strip()         if col_status         else ""
        action       = row.get(col_action, "").strip()         if col_action         else ""
        if not company and not email:
            continue
        clients.append({
            "_row_index":     i,
            "_raw":           row,
            "company":        company,
            "contact_name":   contact_name,
            "role":           role,
            "email":          email,
            "linkedin":       linkedin,
            "date_contacted": date_contacted,
            "website":        description,
            "status":         status,
            "action":         action,
        })

    return clients, headers, rows, col_company, col_email, col_linkedin, col_status, path, ext

def save_clients(clients, headers, rows, col_status, path, ext):
    """Write status updates back to the spreadsheet."""
    # Ensure status column exists in headers
    if col_status is None:
        col_status_name = "status"
        headers = list(headers) + [col_status_name]
    else:
        col_status_name = col_status

    # Update rows
    for client in clients:
        idx = client["_row_index"]
        if idx < len(rows):
            rows[idx][col_status_name] = client["status"]

    if ext == ".csv":
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                if col_status_name not in row:
                    row[col_status_name] = ""
                writer.writerow(row)
    else:
        if not EXCEL_SUPPORT:
            warn("openpyxl not installed, cannot save .xlsx. Install with: pip install openpyxl")
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for row in rows:
            ws.append([row.get(h, "") for h in headers])
        wb.save(path)

# ══════════════════════════════════════════════════════════════════════════════
#  AI EMAIL GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def generate_email(client):
    """Call Claude to personalise the email for a specific client."""
    cfg = CONFIG
    api_key = cfg["anthropic_api_key"]
    template = cfg["email_body_template"]
    subject_tpl = cfg["email_subject_template"]

    website_ctx = f"Their website is: {client['website']}." if client["website"] else "No website provided."

    prompt = f"""You are an expert digital marketing outreach specialist writing a personalised cold email.

Client details:
- Company name: {client['company'] or 'Unknown'}
- Email: {client['email'] or 'Unknown'}
- {website_ctx}

Email body template to personalise:
---
{template}
---

Subject template:
{subject_tpl}

Instructions:
- Personalise the subject and body specifically for this company
- Replace {{{{company_name}}}} with the actual company name
- Replace {{{{website}}}} with their website if available
- Replace {{{{ai_insight}}}} with 2-3 specific, compelling sentences about a digital marketing opportunity relevant to THIS type of business (based on their name/website)
- Keep {{{{sender_name}}}}, {{{{sender_title}}}}, {{{{sender_company}}}}, {{{{sender_website}}}}, {{{{booking_link}}}} as-is — these will be filled in separately
- Keep the tone professional yet warm and conversational
- Do NOT be generic — make it feel like you genuinely researched them
- Keep the body under 200 words

Return ONLY valid JSON, nothing else:
{{"subject": "...", "body": "..."}}"""

    if USE_SDK:
        client_sdk = anthropic.Anthropic(api_key=api_key)
        message = client_sdk.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )
        text = message.content[0].text
    else:
        import urllib.request, json as jsonlib
        payload = jsonlib.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1024,
            "messages": [{"role": "user", "content": prompt}]
        }).encode()
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            }
        )
        with urllib.request.urlopen(req) as resp:
            data = jsonlib.loads(resp.read())
        text = data["content"][0]["text"]

    # Parse JSON response
    clean = text.strip()
    if clean.startswith("```"):
        clean = clean.split("```")[1]
        if clean.startswith("json"):
            clean = clean[4:]
    clean = clean.strip()
    result = json.loads(clean)
    return result["subject"], result["body"]

def fill_sender_placeholders(text):
    """Replace sender-side placeholders with config values."""
    cfg = CONFIG
    replacements = {
        "{{sender_name}}":    cfg["sender_name"],
        "{{sender_title}}":   cfg["sender_title"],
        "{{sender_company}}": cfg["sender_company"],
        "{{sender_website}}": cfg["sender_website"],
        "{{booking_link}}":   cfg["booking_link"],
        "{{booking_label}}":  cfg["booking_label"],
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text

def build_signature():
    cfg = CONFIG
    sig = f"""
--
📅 {cfg['booking_label']}
   30 minutes · Google Meet
   Book here: {cfg['booking_link']}

{cfg['sender_name']}
{cfg['sender_title']} | {cfg['sender_company']}
{cfg['sender_website']}
"""
    return sig

# ══════════════════════════════════════════════════════════════════════════════
#  TERMINAL REVIEW UI
# ══════════════════════════════════════════════════════════════════════════════

def display_email_preview(client, subject, body, index, total):
    section(f"Client {index}/{total}  —  {client['company'] or client['email']}")

    print(f"\n  {C.DIM}Company :{C.RESET} {C.WHITE}{client['company']}{C.RESET}")
    print(f"  {C.DIM}Email   :{C.RESET} {C.CYAN}{client['email']}{C.RESET}")
    if client["website"]:
        print(f"  {C.DIM}Website :{C.RESET} {C.BLUE}{client['website']}{C.RESET}")

    divider()
    print(f"\n  {C.BOLD}Subject:{C.RESET} {C.YELLOW}{subject}{C.RESET}\n")

    # Word-wrap body for terminal display
    for line in body.splitlines():
        if line.strip():
            wrapped = textwrap.fill(line, width=60, initial_indent="  ", subsequent_indent="  ")
            print(f"{C.WHITE}{wrapped}{C.RESET}")
        else:
            print()

    # Show signature preview
    sig = build_signature()
    print(f"\n{C.DIM}", end="")
    for line in sig.strip().splitlines():
        print(f"  {line}")
    print(C.RESET)

def prompt_approval(client):
    """
    New two-step Gmail flow:
      [o]  Open Gmail compose in browser (pre-filled)
      [y]  Confirm you sent it from Gmail
      [n]  Skip this client
      [e]  Edit subject/body then re-preview
      [q]  Quit and save progress
    """
    while True:
        print(f"\n  {C.BOLD}What would you like to do?{C.RESET}")
        print(f"  {C.CYAN}[o]{C.RESET} Open in Gmail  {C.DIM}(pre-fills To, Subject & Body){C.RESET}")
        print(f"  {C.GREEN}[y]{C.RESET} Confirm sent   {C.DIM}(after you clicked Send in Gmail){C.RESET}")
        print(f"  {C.RED}[n]{C.RESET} Skip")
        print(f"  {C.YELLOW}[e]{C.RESET} Edit subject or body")
        print(f"  {C.MAGENTA}[q]{C.RESET} Quit and save progress")

        choice = input(f"\n  {C.BOLD}Your choice (o/y/n/e/q): {C.RESET}").strip().lower()

        if choice == "o":
            return "open", None, None
        elif choice in ("y", "yes"):
            return "confirm_sent", None, None
        elif choice in ("n", "no", "s", "skip"):
            return "skip", None, None
        elif choice == "e":
            return "edit", None, None
        elif choice == "q":
            return "quit", None, None
        else:
            warn("Please type o, y, n, e, or q")

def edit_email(subject, body):
    """Allow inline editing of subject and body."""
    print(f"\n  {C.YELLOW}Edit subject{C.RESET} (press Enter to keep current):")
    print(f"  Current: {subject}")
    new_subject = input(f"  New subject: ").strip()
    if not new_subject:
        new_subject = subject

    print(f"\n  {C.YELLOW}Edit body{C.RESET}")
    print(f"  (Tip: open your editor, paste new content, then paste it here)")
    print(f"  Type your new body below, then type {C.BOLD}END{C.RESET} on a new line (or press Enter to keep current):\n")

    lines = []
    while True:
        line = input()
        if line.strip().upper() == "END":
            break
        lines.append(line)

    new_body = "\n".join(lines).strip() if lines else body
    return new_subject, new_body

# ══════════════════════════════════════════════════════════════════════════════
#  GMAIL COMPOSE  —  Opens pre-filled compose window in browser
# ══════════════════════════════════════════════════════════════════════════════

def open_gmail_compose(to_email, subject, body):
    """
    Opens Gmail compose in the default browser with To, Subject, and Body
    pre-filled. The user reviews and clicks Send inside Gmail.
    Works whether Gmail is open or not — browser handles the login state.
    """
    signature = build_signature()
    full_body = body.rstrip() + "\n" + signature

    params = urllib.parse.urlencode({
        "to":   to_email,
        "su":   subject,
        "body": full_body,
    })
    url = f"https://mail.google.com/mail/?view=cm&fs=1&{params}"
    webbrowser.open(url)

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN AGENT LOOP
# ══════════════════════════════════════════════════════════════════════════════

def main():
    banner()

    # ── Load spreadsheet ──────────────────────────────────────────────────────
    section("Loading client list")
    (clients, headers, rows,
     col_company, col_email, _,
     col_status, path, ext) = load_clients(CONFIG["spreadsheet_path"])

    total = len(clients)
    info(f"Loaded {C.BOLD}{total}{C.RESET} clients from {C.CYAN}{path.name}{C.RESET}")

    # Filter out already sent if configured
    to_process = []
    skipped_count = 0
    for c in clients:
        if CONFIG["skip_already_sent"] and c["status"].lower() in ("sent", "✓ sent"):
            skipped_count += 1
        else:
            to_process.append(c)

    if skipped_count:
        info(f"Skipping {skipped_count} already-sent clients")

    if not to_process:
        success("All clients have already been contacted. Nothing to do!")
        return

    info(f"{C.BOLD}{len(to_process)}{C.RESET} clients to process\n")

    # ── Validate config ───────────────────────────────────────────────────────
    if CONFIG["anthropic_api_key"].startswith("sk-ant-..."):
        error("Please set your Anthropic API key in CONFIG['anthropic_api_key']")
        sys.exit(1)
    if CONFIG["gmail_address"] == "you@gmail.com":
        error("Please set your Gmail address in CONFIG['gmail_address']")
        sys.exit(1)
    # ── Stats tracking ────────────────────────────────────────────────────────
    stats = {"sent": 0, "skipped": 0, "errors": 0}

    # ── Process each client ───────────────────────────────────────────────────
    for i, client in enumerate(to_process, 1):
        try:
            # Generate email
            info(f"Generating email for {C.BOLD}{client['company'] or client['email']}{C.RESET}...")
            subject, body = generate_email(client)
            subject = fill_sender_placeholders(subject)
            body    = fill_sender_placeholders(body)

        except Exception as e:
            warn(f"AI generation failed for {client['company']}: {e}")
            warn("Using template fallback")
            subject = fill_sender_placeholders(
                CONFIG["email_subject_template"].replace("{{company_name}}", client["company"])
            )
            body = fill_sender_placeholders(
                CONFIG["email_body_template"]
                    .replace("{{company_name}}", client["company"])
                    .replace("{{website}}", client["website"])
                    .replace("{{ai_insight}}", "We see great potential to help grow your digital presence.")
            )

        # Review loop
        gmail_opened = False
        while True:
            display_email_preview(client, subject, body, i, len(to_process))
            action, _, _ = prompt_approval(client)

            if action == "edit":
                subject, body = edit_email(subject, body)
                gmail_opened = False   # reset so they can re-open with edits
                continue

            elif action == "open":
                if not client["email"]:
                    warn(f"No email address for {client['company']} — cannot open Gmail")
                    continue
                info(f"Opening Gmail compose for {C.CYAN}{client['email']}{C.RESET}...")
                open_gmail_compose(client["email"], subject, body)
                gmail_opened = True
                info("Gmail opened in your browser.")
                info(f"Review the email, then click {C.BOLD}Send{C.RESET} inside Gmail.")
                info(f"Come back here and press {C.GREEN}[y]{C.RESET} to confirm, or {C.RED}[n]{C.RESET} to skip.")
                continue

            elif action == "confirm_sent":
                if not gmail_opened:
                    warn("You haven't opened Gmail yet — press [o] first to open and send the email.")
                    continue
                client["status"] = f"Sent {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                stats["sent"] += 1
                success(f"Marked as sent · {client['email']}")
                time.sleep(CONFIG["delay_between_sends"])
                break

            elif action == "skip":
                client["status"] = "Skipped"
                stats["skipped"] += 1
                info("Skipped.")
                break

            elif action == "quit":
                info("Saving progress and quitting...")
                save_clients(clients, headers, rows, col_status, path, ext)
                success(f"Progress saved to {path.name}")
                print_summary(stats, len(to_process))
                return

        # Auto-save after every client
        save_clients(clients, headers, rows, col_status, path, ext)

    # ── Done ──────────────────────────────────────────────────────────────────
    section("All done!")
    save_clients(clients, headers, rows, col_status, path, ext)
    success(f"Spreadsheet updated: {path.name}")
    print_summary(stats, len(to_process))

def print_summary(stats, total):
    print(f"""
  {C.BOLD}Summary{C.RESET}
  {C.GREEN}✓ Sent   : {stats['sent']}{C.RESET}
  {C.DIM}— Skipped: {stats['skipped']}{C.RESET}
  {C.RED}✗ Errors : {stats['errors']}{C.RESET}
  {C.DIM}─ Total  : {total}{C.RESET}
""")

if __name__ == "__main__":
    main()
