from flask import Flask, request, jsonify
import win32com.client
import json
import pythoncom  # Import pythoncom to handle COM initialization

# Create an instance of the Flask class
app = Flask(__name__)

# Define the function to send the first email
def send_first_mail(value_id, follow_up_message, from_address, account_name, new_recipient, product_name, follow_up_link):
    try:
        # Initialize COM library
        pythoncom.CoInitialize()

        # Create an Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Get the namespace (MAPI)
        namespace = outlook.GetNamespace("MAPI")
        
        # Identify the correct account
        account = None
        for acc in namespace.Folders:
            if acc.Name == account_name:
                account = acc
                break
        
        if not account:
            return {"error": f"No account found with the name {account_name}"}
        
        # Create a new mail item
        new_mail = outlook.CreateItem(0)  # 0 represents a mail item
        
        # Set up the dynamic body content with HTML formatting
        # body_content = f"""
        # <p>Thank you for sharing your experience with {product_name}.</p>
        # <p>The follow-up form <a href="{follow_up_link}" target="_blank">here</a> will help us gather additional valuable information that will help us thoroughly assess the relationship between {product_name} and the reported adverse event(s) or side effect(s).</p>
        # <p>The form has been prefilled with the information you provided in the initial report. We kindly ask you to complete only the sections requiring additional details, which will remain strictly confidential.</p>
        # """
        body_content = follow_up_message
        
        # Set the HTML body of the message
        new_mail.HTMLBody = body_content
        
        # Set the from address
        new_mail.SentOnBehalfOfName = from_address
        
        # Set the recipient address
        new_mail.To = new_recipient
        
        # Set the subject of the email
        new_mail.Subject = f"{value_id}"
        
        # Send the email
        new_mail.Send()
        return {"message": "First email sent successfully."}
        
    except Exception as e:
        return {"error": str(e)}
    
    finally:
        # Uninitialize COM library after use
        pythoncom.CoUninitialize()

# Define the route for sending email
@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    print(data)
    # Extract the relevant fields from the incoming JSON data
    value_id = data.get('value_id')
    follow_up_message = data.get('follow_up_message')
    new_recipient = data.get('new_recipient')
    product_name = data.get('product_name')
    follow_up_link = data.get('follow_up_link')
    
    # Example usage
    from_address = "gvw-one@outlook.com"
    account_name = "gvw-one@outlook.com"

    # Call the send_first_mail function
    result = send_first_mail(value_id, follow_up_message, from_address, account_name, new_recipient, product_name, follow_up_link)
    
    # Return the result as JSON
    return jsonify(result)

# Run the Flask app
if __name__ == '__main__':
    app.run(port=6020)
