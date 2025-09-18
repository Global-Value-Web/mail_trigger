from flask import Flask, request, jsonify
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

app = Flask(__name__)

def send_first_mail(value_id, follow_up_message, from_address, account_name, new_recipient, product_name, follow_up_link):
    try:
        # Create the email message
        msg = MIMEMultipart("alternative")
        msg["From"] = from_address
        msg["To"] = new_recipient
        msg["Subject"] = f"{value_id}"

        # Email body (HTML)
        html_body = MIMEText(follow_up_message, "html")
        msg.attach(html_body)

        # Connect to Office 365 SMTP server
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        smtp_user = from_address  # your email (same as login)
        smtp_password = os.environ.get("EMAIL_PASSWORD")  # stored securely in env/secret

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(from_address, new_recipient, msg.as_string())

        return {"message": "Email sent successfully."}

    except Exception as e:
        return {"error": str(e)}

@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json

    value_id = data.get("value_id")
    follow_up_message = data.get("follow_up_message")
    new_recipient = data.get("new_recipient")
    product_name = data.get("product_name")
    follow_up_link = data.get("follow_up_link")

    from_address = os.environ.get("FROM_ADDRESS", "gvw-one@outlook.com")
    account_name = os.environ.get("ACCOUNT_NAME", "gvw-one@outlook.com")

    result = send_first_mail(value_id, follow_up_message, from_address, account_name, new_recipient, product_name, follow_up_link)

    return jsonify(result)

if __name__ == "__main__":
    app.run(port=6020, host="0.0.0.0")
