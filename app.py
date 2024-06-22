from flask import Flask, render_template, request, flash, redirect, url_for
import os
import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image as PILImage
import io
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

app = Flask(__name__, template_folder='templates',static_folder='static')
app.secret_key = '123'

@app.route('/')
def index():
    return render_template('index.html')

def send_certificate_email(email, certificate_path):
    smtp_server = "smtp.gmail.com"  # Use the SMTP server of your email provider
    smtp_port = 587  # Port for TLS encryption
    smtp_username = ""  # Your email address
    smtp_password = ""  # Your email password

    subject = "Certificate Attached"
    sender_email = "" # Your email address
    receiver_email = email  # The recipient's email address

    # Create a multipart message and set the headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add a plain-text message (optional)
    message.attach(MIMEText("Please find your certificate attached."))

    # Attach the certificate as a PDF
    with open(certificate_path, "rb") as cert_file:
        part = MIMEApplication(cert_file.read(), Name=os.path.basename(certificate_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(certificate_path)}"'
        message.attach(part)

    # Establish a secure connection with the SMTP server
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        text = message.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending email: {str(e)}")
        return False

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        participants_excel = request.files['participants_excel']
        logo_image = request.files['logo_image']
        mentor_signature = request.files['mentor_signature']
        head_signature = request.files['head_signature']
        output_dir = 'certificates'  # Create a "certificates" folder in your project directory

        if participants_excel and logo_image and mentor_signature and head_signature:
            participants_df = pd.read_excel(participants_excel)

            for _, row in participants_df.iterrows():
                student = row['Student Name']
                school = row['School or University Name']
                date = row['Competition_Date']
                clas = row['Class']
                comp = row['Competition Name']
                position = row['Position Secured']
                email = row['email']  # Assuming your DataFrame has an 'Email' column

                packet = io.BytesIO()
                width, height = letter
                c = canvas.Canvas(packet, pagesize=(width * 2, height * 2))

                pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
                pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
                pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

                c.setFillColorRGB(0 / 255, 0 / 255, 0 / 255)
                c.setFont('VeraBd', 16)
                c.drawCentredString(390, 350, student)


                c.setFillColorRGB(0 / 255, 0 / 255, 0 / 255)
                c.setFont('Vera', 15)
                c.drawCentredString(440, 313, school)

                c.setFillColorRGB(1 / 255, 1 / 255, 1 / 255)
                c.setFont('Vera', 14)
                c.drawCentredString(350, 273, date)

                c.setFillColorRGB(1 / 255, 1 / 255, 1 / 255)
                c.setFont('Vera', 14)
                c.drawCentredString(140, 312, clas)

                c.setFillColorRGB(1 / 255, 1 / 255, 1 / 255)
                c.setFont('Vera', 14)
                c.drawCentredString(630, 273, comp)

                c.setFillColorRGB(1 / 255, 1 / 255, 1 / 255)
                c.setFont('Vera', 14)
                position_str = str(position)
                c.drawCentredString(182, 235, position_str)

                

                # Draw the image from the uploaded logo_image
                logo = PILImage.open(logo_image)
                logo = logo.convert("RGB")
                logo.save(packet, "JPEG")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_image:
                  logo.save(temp_image, format="JPEG")

                # Draw the image from the temporary file on the canvas
                c.drawImage(temp_image.name, 550,447, width=100, height=100, mask=None)

                sign = PILImage.open(mentor_signature)
                sign = sign.convert("RGB")
                sign.save(packet,"JPEG")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_image:
                    sign.save(temp_image, format="JPEG")

                c.drawImage(temp_image.name, 100,190, width=100, height=30, mask=None)

                signature = PILImage.open(head_signature)
                signature = sign.convert("RGB")
                signature.save(packet,"JPEG")

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_image:
                    sign.save(temp_image, format="JPEG")

                c.drawImage(temp_image.name, 550,190, width=100, height=30, mask=None)


                




                # Add other certificate details...

                # After saving the certificate, send it by email
                c.save()
                existing_pdf = PdfReader(open('test.pdf', 'rb'))
                page = existing_pdf.pages[0]
                packet.seek(0)
                new_pdf = PdfReader(packet)
                page.merge_page(new_pdf.pages[0])
                file_name = student.replace(" ", "_")
                os.makedirs(output_dir, exist_ok=True)
                certificado = os.path.join(output_dir, file_name + "_certificate.pdf")
                outputStream = open(certificado, "wb")
                output = PdfWriter()
                output.add_page(page)
                output.write(outputStream)
                outputStream.close()

                # Send the certificate by email
                if send_certificate_email(email, certificado):
                    flash(f"Certificate sent to {email} successfully!", "success")
                else:
                    flash(f"Failed to send certificate to {email}", "error")
            flash("Certificates generated successfully!", "success")
        else:
            flash("Please upload both participants Excel file and logo image.", "error")

    except Exception as e:
        flash(f"An error occurred: {str(e)}", "error")

    return redirect(url_for('index'))

