from flask import Flask, request, jsonify
import json
import os

# Try importing Windows-only libraries
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

import smtplib
from email.message import EmailMessage

# Create an instance of the Flask class
app = Flask(__name__)

def send_first_mail(value_id, follow_up_message, from_address, account_name, new_recipient, product_name, follow_up_link):
    try:
        if win32com:  # Windows Outlook COM
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            account = None
            for acc in namespace.Folders:
                if acc.Name == account_name:
                    account = acc
                    break

            if not account:
                return {"error": f"No account found with the name {account_name}"}

            new_mail = outlook.CreateItem(0)  # Mail item
            new_mail.HTMLBody = follow_up_message
            new_mail.SentOnBehalfOfName = from_address
            new_mail.To = new_recipient
            new_mail.Subject = f"{value_id}"
            new_mail.Send()
            return {"message": "First email sent successfully (via Outlook COM)."}

        else:  # Linux/Docker fallback → SMTP
            msg = EmailMessage()
            msg["Subject"] = f"{value_id}"
            msg["From"] = from_address
            msg["To"] = new_recipient
            msg.set_content(follow_up_message)

            smtp_host = os.environ.get("SMTP_HOST")
            smtp_port = int(os.environ.get("SMTP_PORT", "587"))
            smtp_user = os.environ.get("SMTP_USER")
            smtp_pass = os.environ.get("SMTP_PASS")

            with smtplib.SMTP(smtp_host, smtp_port) as s:
                s.starttls()
                if smtp_user and smtp_pass:
                    s.login(smtp_user, smtp_pass)
                s.send_message(msg)

            return {"message": "First email sent successfully (via SMTP)."}

    except Exception as e:
        return {"error": str(e)}

    finally:
        if win32com:
            pythoncom.CoUninitialize()

# API route
@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    value_id = data.get('value_id')
    follow_up_message = data.get('follow_up_message')
    new_recipient = data.get('new_recipient')
    product_name = data.get('product_name')
    follow_up_link = data.get('follow_up_link')

    from_address = os.environ.get("SMTP_FROM", "gvw-one@outlook.com")
    account_name = "gvw-one@outlook.com"

    result = send_first_mail(
        value_id, follow_up_message, from_address,
        account_name, new_recipient, product_name, follow_up_link
    )

    return jsonify(result)

if __name__ == '__main__':
    app.run(port=6020, host="0.0.0.0")
