import os
import smtplib
import pandas as pd
import yaml
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load email recipients
with open("config/email_recipients.yaml", "r") as f:
    config = yaml.safe_load(f)
recipients = config.get("recipients", [])

# Exit if no recipients
if not recipients:
    print("[!] No email recipients found.")
    exit(1)

# Email credentials from secrets
sender = os.environ.get("EMAIL_USER")
password = os.environ.get("EMAIL_PASSWORD")

if not sender or not password:
    print("[!] EMAIL_USER or EMAIL_PASSWORD is missing.")
    exit(1)

# Load Excel data
df = pd.read_excel("data/data.xlsx", engine="openpyxl")
df.columns = [col.strip().replace(" ", "_") for col in df.columns]
df["Next_Review_Date"] = pd.to_datetime(df["Next_Review_Date"], dayfirst=True, errors="coerce")
df["Date_Changed"] = pd.to_datetime(df["Date_Changed"], dayfirst=True, errors="coerce")

# Reference dates
today = pd.Timestamp.now().normalize()
in_1_day = today + pd.Timedelta(days=1) # 1 day
in_3_days = today + pd.Timedelta(days=3) # 3 days
in_7_days = today + pd.Timedelta(days=7) # 1 week
long_ago = today - pd.Timedelta(days=90) # 3 months

# Categorize
review_1_day = df[df["Next_Review_Date"] <= in_1_day]
review_3_days = df[(df["Next_Review_Date"] > in_1_day) & (df["Next_Review_Date"] <= in_3_days)]
review_1_week = df[(df["Next_Review_Date"] > in_3_days) & (df["Next_Review_Date"] <= in_7_days)]
old_passwords = df[df["Date_Changed"] < long_ago]

# Format table for HTML
def format_df(sub_df):
    cols = ["LastPass_ID", "Changed_By", "Next_Review_Date", "Service_Website", "Date_Changed"]
    sub_df = sub_df[cols].copy()
    sub_df["Next_Review_Date"] = pd.to_datetime(sub_df["Next_Review_Date"]).dt.strftime("%d %B %Y")
    sub_df["Date_Changed"] = pd.to_datetime(sub_df["Date_Changed"]).dt.strftime("%d %B %Y")
    return sub_df.to_html(index=False, border=0, justify="center", classes="styled-table")

# Build email sections
sections = {
    "within 1 day": (review_1_day, "You have {count} password(s) that need to be reviewed within 1 day:"),
    "within 3 days": (review_3_days, "You have {count} password(s) that need to be reviewed within 3 days:"),
    "within 1 week": (review_1_week, "You have {count} password(s) that need to be reviewed within 1 week:"),
    "old": (old_passwords, "You have {count} password(s) that haven't been updated in a long time, please review them for security:")
}

html_sections = ""
total_count = 0
for _, (subset, message) in sections.items():
    if not subset.empty:
        total_count += len(subset)
        html_sections += f"<p>{message.format(count=len(subset))}</p>{format_df(subset)}"

# Exit if nothing to send
if not html_sections:
    print("[✓] No relevant passwords found. Email not sent.")
    exit(0)

# Final footer note
html_note = """
<p class="note">
    If you're not already using it, I highly recommend using <strong>LastPass</strong> for managing and securely storing your passwords.<br>
    You can access it here: <a href="https://lastpass.com/vault/" target="_blank">https://lastpass.com/vault/</a>
</p>
"""

# Build full HTML email
html_body = f"""
<html>
<head>
  <style>
    body {{
      font-family: Arial, sans-serif;
    }}
    .styled-table {{
      border-collapse: collapse;
      margin: 25px 0;
      font-size: 14px;
      width: 100%;
    }}
    .styled-table th, .styled-table td {{
      border: 1px solid #dddddd;
      text-align: left;
      padding: 8px;
    }}
    .styled-table th {{
      background-color: #f2f2f2;
    }}
    .note {{
      margin-top: 30px;
      font-size: 13px;
      color: #555;
    }}
  </style>
</head>
<body>
  <p>Hey,</p>
  {html_sections}
  {html_note}
</body>
</html>
"""

# Create email message
msg = MIMEMultipart("alternative")
msg["Subject"] = "Password Review Reminder"
msg["From"] = sender
msg["To"] = ", ".join(recipients)
msg.attach(MIMEText(html_body, "html"))

# Send email
try:
    print("[*] Connecting to SMTP...")
    with smtplib.SMTP_SSL("%SMTP_HOST%", 465) as smtp:
        smtp.login(sender, password)
        smtp.sendmail(sender, recipients, msg.as_string())
        print(f"[✓] Email sent to: {', '.join(recipients)}")
except Exception as e:
    print(f"[X] Failed to send email: {e}")
