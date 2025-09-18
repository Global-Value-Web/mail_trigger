import smtplib
import os
from flask import Flask, request, jsonify
from email.mime.text import MIMEText

app = Flask(__name__)

FROM_ADDRESS = os.getenv("FROM_ADDRESS")
ACCOUNT_NAME = os.getenv("ACCOUNT_NAME")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

#Done
@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    value_id = data.get("value_id")
    follow_up_message = data.get("follow_up_message")
    new_recipient = data.get("new_recipient")

    try:
        msg = MIMEText(follow_up_message, "html")
        msg["Subject"] = value_id
        msg["From"] = FROM_ADDRESS
        msg["To"] = new_recipient

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(ACCOUNT_NAME, EMAIL_PASSWORD)
        server.sendmail(FROM_ADDRESS, [new_recipient], msg.as_string())
        server.quit()

        return jsonify({"message": "Email sent successfully."})

    except Exception as e:
        return jsonify({"error": str(e)})
