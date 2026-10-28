"""Bulk outreach sender: reads the leads sheet, personalises, sends via your own SMTP mailbox.

Safe by default: without --send it only prints what it WOULD send.

Setup:
  pip install openpyxl anthropic        (anthropic only needed for --ai)
  export SMTP_HOST=smtp.gmail.com SMTP_PORT=587
  export SMTP_USER=you@yourdomain.com SMTP_PASS='app-password'   # Gmail: create an App Password
  export FROM_NAME='Your Name' REPLY_TO=you@yourdomain.com
  export ANTHROPIC_API_KEY=...          # only for --ai

Usage:
  python send_emails.py --leads dental_lab_emails.xlsx            # preview
  python send_emails.py --leads dental_lab_emails.xlsx --send --daily-limit 30
"""
import argparse, csv, os, random, re, smtplib, ssl, sys, time
from datetime import date
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

import openpyxl

if os.path.exists(".env"):  # simple KEY=value loader, so no extra packages are needed
    for _l in open(".env"):
        if "=" in _l and not _l.lstrip().startswith("#"):
            _k, _v = _l.strip().split("=", 1)
            os.environ.setdefault(_k, _v.strip().strip("'\""))

LOG = "sent_log.csv"
OPTOUT = "optout.txt"  # one address per line; add anyone who replies "unsubscribe"
LEADS_SHEET = "Dental Labs - SEND"


def load_leads(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if LEADS_SHEET not in wb.sheetnames:
        raise SystemExit(
            f'Sheet "{LEADS_SHEET}" not found in {path}. Sheets: {", ".join(wb.sheetnames)}'
        )
    ws = wb[LEADS_SHEET]
    head = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(head, row))
        raw = r.get("Email")
        if raw is None:
            continue
        email = str(raw).strip()
        if not email or "@" not in email:
            continue
        yield {
            "name": r.get("Name") or "",
            "category": r.get("Category") or "",
            "city": r.get("City") or "",
            "email": email,
        }


def already_done():
    done = set()
    if os.path.exists(LOG):
        done |= {r[1] for r in csv.reader(open(LOG)) if len(r) > 1}
    if os.path.exists(OPTOUT):
        done |= {l.strip().lower() for l in open(OPTOUT) if l.strip()}
    return done


def render(tpl, lead):
    subject, body = tpl.split("\n---\n", 1)
    subs = {"{name}": lead["name"] or "", "{city}": lead["city"] or "", "{category}": lead["category"] or ""}
    for k, v in subs.items():
        subject, body = subject.replace(k, v), body.replace(k, v)
    return subject.replace("SUBJECT:", "").strip(), body.strip()


def ai_personalise(lead, subject, body):
    """Optional: let Claude add one natural opening sentence. Never changes the offer or the facts."""
    import anthropic
    msg = anthropic.Anthropic().messages.create(
        model="claude-sonnet-5-5", max_tokens=120,
        messages=[{"role": "user", "content":
                   f"Write ONE short, genuine opening sentence (German if the business is German) for a cold email to "
                   f"'{lead['name']}', a {lead['category']} in {lead['city']}. No flattery, no invented facts, "
                   f"no claims about them beyond name/type/city. Output only the sentence."}])
    return subject, msg.content[0].text.strip() + "\n\n" + body


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--leads", required=True)
    ap.add_argument("--template", default="template.txt")
    ap.add_argument("--send", action="store_true", help="actually send (default: preview only)")
    ap.add_argument("--daily-limit", type=int, default=30)
    ap.add_argument("--min-delay", type=int, default=60)
    ap.add_argument("--max-delay", type=int, default=180)
    ap.add_argument("--ai", action="store_true")
    ap.add_argument("--test-to", help="send ONE sample email (first lead's data) to this address")
    ap.add_argument("--limit", type=int, default=5, help="preview count when not sending")
    a = ap.parse_args()

    tpl = open(a.template, encoding="utf-8").read()

    if a.test_to:  # one sample email to a single address; not logged, leads untouched
        env = {k: os.environ[k] for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASS", "FROM_NAME")}
        lead = next(load_leads(a.leads))
        subject, body = render(tpl, lead)
        m = EmailMessage()
        m["From"] = f"{env['FROM_NAME']} <{env['SMTP_USER']}>"
        m["To"], m["Subject"] = a.test_to, subject
        m["Reply-To"] = os.environ.get("REPLY_TO", env["SMTP_USER"])
        m["Date"], m["Message-ID"] = formatdate(localtime=True), make_msgid()
        m.set_content(body)
        port = int(os.environ.get("SMTP_PORT", 587))
        if port == 465:
            s = smtplib.SMTP_SSL(env["SMTP_HOST"], port, context=ssl.create_default_context())
        else:
            s = smtplib.SMTP(env["SMTP_HOST"], port)
            s.starttls(context=ssl.create_default_context())
        s.login(env["SMTP_USER"], env["SMTP_PASS"])
        s.send_message(m)
        s.quit()
        print(f"Test email sent to {a.test_to}")
        return
    done = already_done()
    today = date.today().isoformat()
    sent_today = sum(1 for r in csv.reader(open(LOG)) if r and r[0] == today) if os.path.exists(LOG) else 0

    leads = [l for l in load_leads(a.leads) if l["email"].lower() not in done]
    print(f"{len(leads)} unsent leads. Already sent today: {sent_today}")

    server = None
    if a.send:
        env = {k: os.environ[k] for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASS", "FROM_NAME")}
        port = int(os.environ.get("SMTP_PORT", 587))
        if port == 465:
            server = smtplib.SMTP_SSL(env["SMTP_HOST"], port, context=ssl.create_default_context())
        else:
            server = smtplib.SMTP(env["SMTP_HOST"], port)
            server.starttls(context=ssl.create_default_context())
        server.login(env["SMTP_USER"], env["SMTP_PASS"])

    n = 0
    for lead in leads:
        if a.send and sent_today + n >= a.daily_limit:
            print("Daily limit reached; run again tomorrow.")
            break
        if not a.send and n >= a.limit:
            break
        subject, body = render(tpl, lead)
        if a.ai:
            subject, body = ai_personalise(lead, subject, body)
        if not a.send:
            print(f"\n--- To: {lead['email']} ({lead['name']})\nSubject: {subject}\n\n{body}")
            n += 1
            continue
        m = EmailMessage()
        m["From"] = f"{env['FROM_NAME']} <{env['SMTP_USER']}>"
        m["To"] = lead["email"]
        m["Subject"] = subject
        m["Reply-To"] = os.environ.get("REPLY_TO", env["SMTP_USER"])
        m["Date"], m["Message-ID"] = formatdate(localtime=True), make_msgid()
        m["List-Unsubscribe"] = f"<mailto:{os.environ.get('REPLY_TO', env['SMTP_USER'])}?subject=unsubscribe>"
        m.set_content(body)
        try:
            server.send_message(m)
            csv.writer(open(LOG, "a", newline="")).writerow([today, lead["email"], lead["name"], "sent"])
            print(f"sent -> {lead['email']}")
        except Exception as e:
            csv.writer(open(LOG, "a", newline="")).writerow([today, lead["email"], lead["name"], f"error: {e}"])
            print(f"FAILED {lead['email']}: {e}")
        n += 1
        time.sleep(random.randint(a.min_delay, a.max_delay))
    if server:
        server.quit()


if __name__ == "__main__":
    main()
