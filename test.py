from shipping import send_email

send_email(
    subject="[Adscraper] SMTP test",
    body="If you got this, SMTP is configured correctly.",
    attachments=[]
)
print("Sent.")
