import smtplib
import time
import random
import email.utils
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Set TESTING_MODE = True for testing (short delays) and False for production (1 and 2 days)
TESTING_MODE = False
followup_delay = 10 if TESTING_MODE else 86400  # 86400 seconds = 1 day

SENDER_ACCOUNT = {
    "email": "neal@filldesignprojects.website",
    "password": "Fdg@9874#",
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587
}

def get_random_sender():
    return SENDER_ACCOUNT

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    greetings = [f"Hi {person_name},", f"Hello {person_name},", f"Dear {person_name},"]
    sentence1 = random.choice([
        "Hope you’re doing well! I’m Neal from Fill Design Group — we help spa owners get more bookings and boost revenue through smart digital marketing.",
        "I hope you’re having a great day! I’m Neal at Fill Design Group. We specialize in helping spa owners increase their bookings and drive higher revenue through effective digital marketing",
        "Hope all is well! I’m Neal from Fill Design Group. We help spas like yours maximize bookings and revenue using tailored digital marketing strategies."
    ])

    # Bullet-pointed plain-text options for sentence2:
    sentence2_text = random.choice([
        (
            "We’ve recently partnered with RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL. "
            "By combining solid SEO, targeted Facebook & Instagram ad campaigns, and automated follow‑up sequences, we were able to:\n"
            "• Attract a consistent stream of high‑quality leads\n"
            "• Turn those leads into confirmed appointments\n"
            "• Boost their revenue by 38% in just 90 days"
        ),
        (
            "Just recently, we teamed up with RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL. "
            "Through a combination of ongoing SEO work, precision‑targeted Facebook & Instagram ads, and automated follow‑ups, we:\n"
            "• Generated a steady influx of qualified inquiries\n"
            "• Converted those inquiries into booked appointments\n"
            "• Increased their revenue by 38% within 90 days"
        ),
        (
            "We recently delivered great results for RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL. "
            "Our mix of continuous SEO, pinpoint Facebook & Instagram advertising, and automated follow‑up campaigns helped them:\n"
            "• Maintain a reliable pipeline of quality leads\n"
            "• Turn those leads into confirmed appointments\n"
            "• Achieve a 38% revenue uplift in just 90 days"
        ),
    ])

    # Bullet-pointed HTML options for sentence2:
    sentence2_html = random.choice([
        """
        <p>We’ve recently partnered with RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL.
        By combining solid SEO, targeted Facebook & Instagram ad campaigns, and automated follow‑up sequences, we were able to:</p>
        <ul>
          <li>Attract a consistent stream of high‑quality leads</li>
          <li>Turn those leads into confirmed appointments</li>
          <li>Boost their revenue by 38% in just 90 days</li>
        </ul>
        """,
        """
        <p>Just recently, we teamed up with RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL.
        Through a combination of ongoing SEO work, precision‑targeted Facebook & Instagram ads, and automated follow‑ups, we:</p>
        <ul>
          <li>Generated a steady influx of qualified inquiries</li>
          <li>Converted those inquiries into booked appointments</li>
          <li>Increased their revenue by 38% within 90 days</li>
        </ul>
        """,
        """
        <p>We recently delivered great results for RiverMed Spa in Cocoa Village, FL, and Medloft Spa in Melbourne, FL.
        Our mix of continuous SEO, pinpoint Facebook & Instagram advertising, and automated follow‑up campaigns helped them:</p>
        <ul>
          <li>Maintain a reliable pipeline of quality leads</li>
          <li>Turn those leads into confirmed appointments</li>
          <li>Achieve a 38% revenue uplift in just 90 days</li>
        </ul>
        """
    ])

    sentence3 = random.choice([
        "If you’d like to see how we can fill your appointment book and grow your spa’s bottom line, let’s chat!",
        "If you're looking to fill your schedule and grow your spa, I’d love to show you what we can do.",
        "If you’re interested in filling up your calendar and boosting your spa’s growth, I’d love to share our approach."
    ])

    extra = f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email." if is_followup and followup_number else ""

    # Assemble plain-text body
    text = (
        f"{random.choice(greetings)}\n\n"
        f"{sentence1}\n\n"
        f"{sentence2_text}\n\n"
        f"{sentence3}\n\n"
        f"{extra}\n\n"
        "Looking forward to hearing from you.\n\n"
        "Best regards,\n"
        "Neal\n"
        "https://filldesigngroup.com/\n"
    )

    # Assemble HTML body
    html = f"""
<html><body>
  <p>{random.choice(greetings)}</p>
  <p>{sentence1}</p>
  {sentence2_html}
  <p>{sentence3}</p>
  {f"<p>{extra}</p>" if extra else ""}
  <p>Looking forward to hearing from you.<br>Best regards,<br>Neal<br>
     <a href="https://filldesigngroup.com/">Fill Design Group</a></p>
</body></html>
"""
    return text, html

def choose_subject(company):
    return random.choice([
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]).format(Company=company)

def check_reply(email_address):
    return False  # placeholder

def send_initial_email(row):
    company = row['company']
    name    = row['name']
    to_addr = row['email']
    subject = choose_subject(company)
    text, html = spin_email_template(name, company)

    sender = get_random_sender()
    msg = MIMEMultipart('alternative')
    msg['From']    = sender['email']
    msg['To']      = to_addr
    msg['Subject'] = subject
    msg_id = email.utils.make_msgid()
    msg['Message-ID'] = msg_id
    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        # Use SMTP + STARTTLS on port 587 for Office 365
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Initial email sent to {to_addr} from {sender['email']}")
    except Exception as e:
        print(f"Error sending to {to_addr}: {e}")
        return None, None, None

    return msg_id, subject, sender

def send_followup(to_addr, msg_id, name, company, num, sender, orig_subj):
    text, html = spin_email_template(name, company, True, num)
    msg = MIMEMultipart('alternative')
    msg['From']        = sender['email']
    msg['To']          = to_addr
    msg['Subject']     = "Re: " + orig_subj
    msg['In-Reply-To'] = msg_id
    msg['References']  = msg_id
    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Follow-up #{num} sent to {to_addr}")
    except Exception as e:
        print(f"Error sending follow-up #{num} to {to_addr}: {e}")

def followup_scheduler(to_addr, msg_id, name, company, sender, subj):
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 1, sender, subj)
    else:
        return
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 2, sender, subj)

def send_emails(xlsx_path):
    df = pd.read_excel(xlsx_path, engine='openpyxl')
    for _, row in df.iterrows():
        company = row['company']
        name    = row['name']
        email   = row['email']
        print(f"Processing: {company} | {name} | {email}")

        msg_id, subj, sender = send_initial_email(row)
        if not msg_id:
            continue

<<<<<<< HEAD
        # Always schedule follow-ups in the same thread (sequential blocking)
        followup_scheduler(email, msg_id, name, company, sender, subj)

        # Pause for 60 seconds between sends
=======
        if TESTING_MODE:
            # In testing mode, run follow-ups sequentially (short delays)
            followup_scheduler(email, msg_id, name, company, sender, subj)
        else:
            # In production mode, do not spawn threads. Handle follow-ups via external scheduler.
            print(
                "NOTE: follow-up emails are not scheduled automatically. "
                "Use a separate scheduler or cron job to send follow-ups."
            )

        # Randomize pause between sends
>>>>>>> ce992704d5182ba987a4ea8297743bbe89071867
        time.sleep(60)

if __name__ == "__main__":
    send_emails("test Email.xlsx")
