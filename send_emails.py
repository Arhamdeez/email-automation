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
import argparse, csv, imaplib, os, random, re, smtplib, socket, ssl, sys, time
from datetime import date
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import formatdate, make_msgid
from html import escape

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


def log_rows():
    if not os.path.exists(LOG):
        return
    with open(LOG, encoding="utf-8-sig", newline="") as f:
        for r in csv.reader(f):
            if len(r) >= 5 and r[1] != "email":
                yield r


def was_sent(status):
    return status.strip().lower().startswith("sent")


def already_done():
    """Addresses that actually received mail, plus opt-outs. Failures stay eligible for retry."""
    done = {r[1].strip().lower() for r in log_rows() if was_sent(r[4])}
    if os.path.exists(OPTOUT):
        done |= {l.strip().lower() for l in open(OPTOUT) if l.strip()}
    return done


def log_sent(row):
    new = not os.path.exists(LOG) or os.path.getsize(LOG) == 0
    with open(LOG, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date", "email", "name", "city", "status"])
        w.writerow(row)
        f.flush()


def imap_host():
    if os.environ.get("IMAP_HOST"):
        return os.environ["IMAP_HOST"]
    smtp = os.environ.get("SMTP_HOST", "")
    if smtp.startswith("smtp.") and "." in smtp:
        return "imap." + smtp.split(".", 1)[1]
    return "imap.hostinger.com"


def sent_mailbox(imap):
    status, boxes = imap.list()
    if status == "OK":
        for raw in boxes or []:
            line = raw.decode(errors="replace") if isinstance(raw, bytes) else str(raw)
            flags = line.split(")", 1)[0]
            if "\\Sent" not in flags:
                continue
            name = line.rsplit(" ", 1)[-1].strip().strip('"')
            if name:
                return name
    return "INBOX.Sent"


def save_to_sent(message):
    """Copy an already-accepted message into the mailbox Sent folder. SMTP does not do this."""
    port = int(os.environ.get("IMAP_PORT", "993"))
    imap = imaplib.IMAP4_SSL(imap_host(), port, timeout=30)
    try:
        imap.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
        folder = sent_mailbox(imap)
        status, data = imap.append(folder, "\\Seen", None, message.as_bytes(policy=SMTP))
        if status != "OK":
            raise RuntimeError(data)
        return folder
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def send_and_close(env, port, message):
    """One short SMTP session per message. Holding a connection across the delay lets Hostinger idle-timeout it."""
    timeout = 30
    previous_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    server = None
    try:
        if port == 465:
            server = smtplib.SMTP_SSL(env["SMTP_HOST"], port, timeout=timeout, context=ssl.create_default_context())
        else:
            server = smtplib.SMTP(env["SMTP_HOST"], port, timeout=timeout)
            server.starttls(context=ssl.create_default_context())
        if server.sock is not None:
            server.sock.settimeout(timeout)
        server.login(env["SMTP_USER"], env["SMTP_PASS"])
        server.send_message(message)
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                try:
                    server.close()
                except Exception:
                    pass
        socket.setdefaulttimeout(previous_timeout)


APP_STORE = "https://apps.apple.com/de/app/elite-dent/id6803564344"


def pick(options, lead, salt):
    """Same lab always gets the same wording. Neighbouring labs do not."""
    key = f"{lead.get('email') or ''}|{salt}".lower()
    return options[sum(ord(c) for c in key) % len(options)]


def render(tpl, lead):
    subject_tpl, body = tpl.split("\n---\n", 1)
    name = str(lead["name"] or "").strip()
    city = str(lead["city"] or "").strip()
    who = f"{name} in {city}" if city and city.lower() not in name.lower() else name
    subjects = [
        f"{who} und eine eigene App",
        f"Für {name}: Praxis und Labor gemeinsam",
        f"Vorschlag einer App für {who}",
    ]
    openings = [
        f"ich schreibe {who} an, weil eine eigene App mit Ihrem Namen den Austausch mit Ihren Praxen bündeln kann.",
        f"kurz zu {who}: eine App unter Ihrem Namen, genutzt von Ihren Praxen.",
        f"zu {who} hätte ich einen Vorschlag, eine eigene App für den Ablauf zwischen Labor und Praxis.",
    ]
    products = [
        'Für einen deutschen Kunden haben wir das in "Elite Dent" umgesetzt. Scans und Fotos werden hochgeladen, die Scan-Qualität automatisch geprüft, dazu VITA-Farbbestimmung, Fallstatus, Termine und Chat zwischen Praxis und Labor.',
        '"Elite Dent" ist bei einem deutschen Kunden live. Praxis und Dentallabor arbeiten dort an denselben Fällen: Scans, Fotos, automatische Qualitätsprüfung, VITA-Farbe, Status, Termine und Chat.',
        'Ein deutscher Kunde nutzt "Elite Dent" genau dafür. Upload von Scans und Fotos, automatische Scan-Prüfung, VITA-Farbbestimmung, Fallstatus, Termine und der Chat mit der Praxis.',
    ]
    offers = [
        f"Dasselbe können wir für {name} unter Ihrem Namen bauen.\n"
        "– Logo und Name Ihres Labors, Praxen als Nutzer\n"
        "– der Funktionsumfang folgt Ihrem Ablauf\n"
        "– App-Store-Veröffentlichung, Betrieb und Wartung inklusive\n"
        "– auch als monatliches Abo, ohne hohe Einmalkosten",
        f"Für {name} wäre das eine eigene Version.\n"
        "– Ihre Praxen arbeiten in Ihrer App\n"
        "– Funktionen passen wir an Ihren Ablauf an\n"
        "– Veröffentlichung und laufender Betrieb liegen bei uns\n"
        "– auf Wunsch monatlich statt einer großen Einmalzahlung",
        f"{name} bekommt damit eine eigene App, kein fertiges Standardprodukt.\n"
        "– Ihr Name und Ihr Logo\n"
        "– Praxen als Nutzer, Ablauf nach Ihrer Arbeitsweise\n"
        "– wir veröffentlichen, betreiben und warten sie\n"
        "– Monatsabo möglich, wenn Sie keine hohe Startzahlung wollen",
    ]
    teams = [
        "Wir sind VAWCOM, ein kleines Gründerteam. Web, Apps und KI kommen aus einer Hand: https://vawcom.com",
        "Hinter der Umsetzung steht VAWCOM, ein kleines Team für Web, Apps und KI: https://vawcom.com",
        "VAWCOM ist klein und gründergeführt. Websites, Apps und KI machen wir selbst: https://vawcom.com",
    ]
    ctas = [
        "Haben Sie 15 Minuten? Eine Antwort mit einem Wunschtermin reicht.",
        "Für ein kurzes Gespräch von 15 Minuten schreiben Sie mir gern einen Termin.",
        "Wenn ein Termin in den nächsten Tagen passt, antworten Sie einfach mit Uhrzeit und Tag.",
    ]
    subs = {
        "{subject}": pick(subjects, lead, "subject"),
        "{opening}": pick(openings, lead, "opening"),
        "{product}": pick(products, lead, "product"),
        "{offer}": pick(offers, lead, "offer"),
        "{team}": pick(teams, lead, "team"),
        "{cta}": pick(ctas, lead, "cta"),
        "{name}": name,
        "{city}": city,
    }
    for k, v in subs.items():
        subject_tpl = subject_tpl.replace(k, v)
        body = body.replace(k, v)
    return subject_tpl.replace("SUBJECT:", "").strip(), body.strip()


def html_from_plain(plain):
    """HTML version so App Store, website, and email addresses are real links."""
    blocks = []
    for para in plain.split("\n\n"):
        lines = []
        for line in para.split("\n"):
            text = escape(line)
            text = re.sub(
                r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}",
                lambda m: f'<a href="mailto:{m.group(0)}">{m.group(0)}</a>',
                text,
            )
            text = re.sub(
                r"https?://[^\s<]+",
                lambda m: f'<a href="{m.group(0)}">{m.group(0)}</a>',
                text,
            )
            lines.append(text)
        blocks.append("<p>" + "<br>\n".join(lines) + "</p>")
    return (
        '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5;color:#222;">'
        + "".join(blocks)
        + "</body></html>"
    )


def attach_body(message, plain):
    message.set_content(plain)
    message.add_alternative(html_from_plain(plain), subtype="html")


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
        attach_body(m, body)
        port = int(os.environ.get("SMTP_PORT", 587))
        send_and_close(env, port, m)
        try:
            folder = save_to_sent(m)
        except Exception as e:
            raise SystemExit(f"Test email sent to {a.test_to}, but copying it to Sent failed: {e}")
        print(f"Test email sent to {a.test_to} and saved in {folder}")
        return
    done = already_done()
    today = date.today().isoformat()
    sent_today = sum(1 for r in log_rows() if r[0] == today and was_sent(r[4]))

    leads = [l for l in load_leads(a.leads) if l["email"].lower() not in done]
    print(f"{len(leads)} unsent leads. Already sent today: {sent_today}")

    env = port = None
    if a.send:
        env = {k: os.environ[k] for k in ("SMTP_HOST", "SMTP_USER", "SMTP_PASS", "FROM_NAME")}
        port = int(os.environ.get("SMTP_PORT", 587))

    n = 0
    consecutive_failures = 0
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
        attach_body(m, body)
        err = None
        for attempt in (1, 2):
            try:
                send_and_close(env, port, m)
                err = None
                break
            except Exception as e:
                err = e
                print(f"FAILED {lead['email']} (attempt {attempt}): {e}", flush=True)
                if attempt == 1:
                    time.sleep(10)
        if err is not None:
            log_sent([today, lead["email"], lead["name"], lead["city"], f"error: {err}"])
            consecutive_failures += 1
            if consecutive_failures >= 3:
                print("Stopping after 3 failed sends in a row. Failed addresses stay unsent and can be retried.")
                break
            time.sleep(random.randint(a.min_delay, a.max_delay))
            continue
        consecutive_failures = 0
        try:
            save_to_sent(m)
        except Exception as e:
            log_sent([today, lead["email"], lead["name"], lead["city"], f"sent (not in Sent folder: {e})"])
            print(f"sent -> {lead['email']} (Sent folder copy failed: {e})", flush=True)
        else:
            log_sent([today, lead["email"], lead["name"], lead["city"], "sent"])
            print(f"sent -> {lead['email']}", flush=True)
        n += 1
        time.sleep(random.randint(a.min_delay, a.max_delay))


if __name__ == "__main__":
    main()
