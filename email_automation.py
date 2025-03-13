import win32com.client
from datetime import datetime, timedelta

# Calculate last working date (previous weekday)
today = datetime.today()
last_working_day = today - timedelta(days=1 if today.weekday() != 0 else 3)
last_working_date = last_working_day.strftime("%Y-%m-%d")

# Email details
subject = f"Daily Billing Update - {last_working_date}"

to_recipients = ["recipient1@example.com", "recipient2@example.com"]
cc_recipients = ["cc1@example.com", "cc2@example.com"]

# Email body
body = f"""
<html>
<head>
    <style>
        body {{
            font-family: Calibri, sans-serif;
            font-size: 11pt;
        }}
    </style>
</head>
<body>
    Hi Team,<br><br>
    Please find the billing update for <b>{last_working_date}</b>.<br><br>
</body>
</html>
"""

# Initialize Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = subject
mail.To = ";".join(to_recipients)
mail.CC = ";".join(cc_recipients)
mail.HTMLBody = body

# Open the email for review
mail.Display()

print("Email draft is ready in Outlook.")
