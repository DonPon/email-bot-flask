from flask import Flask, render_template, request
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_email', methods=['POST'])
def send_email():
    sender_email = request.form['sender_email']
    email_token = request.form['email_token']
    attachment = request.files['attachment']
    email_list_file = request.files['email_list']
    email_subject = request.form['email_subject']
    email_message = request.form['email_message']

    # Setup email server and credentials
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, email_token)

    # Read email list from Excel file
    email_list = read_email_list_from_excel(email_list_file)

    for recipient_email in email_list:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = email_subject

        # Attach file
        attachment_part = MIMEApplication(attachment.read())
        attachment_part.add_header('Content-Disposition', 'attachment', filename=attachment.filename)
        msg.attach(attachment_part)

        # Email body
        msg.attach(MIMEText(email_message, 'plain'))

        # Send email
        server.sendmail(sender_email, recipient_email, msg.as_string())

    # Close the server
    server.quit()

    return 'Emails sent successfully!'

def read_email_list_from_excel(file):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.active

    # Assuming email addresses are in the first column (Column A)
    email_list = [str(cell.value) for cell in sheet['A'] if cell.value is not None]

    return email_list

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=True)
