import pandas as pd
import smtplib
import os
import glob
from email.message import EmailMessage

# ================= CONFIGURATION =================
SENDER_EMAIL = "noresponder@tierradelfuego.edu.ar"
SENDER_PASSWORD = "M3cccyt_N0_R3sp0nd3r"

EXCEL_FILE = "datos.xlsx"
PDF_FOLDER = "certificados_emitidos"
TEMPLATE_FILE = "email_template.html"  # Name of your HTML file

SMTP_SERVER = "mail.tierradelfuego.edu.ar"
SMTP_PORT = 587
# =================================================

def send_emails():
    # 1. Read the Excel file
    try:
        df = pd.read_excel(EXCEL_FILE, dtype={'DNI': str})
    except FileNotFoundError:
        print(f"Error: Could not find Excel file '{EXCEL_FILE}'")
        return

    # 2. Read the HTML Template
    try:
        with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
            html_template = f.read()
    except FileNotFoundError:
        print(f"Error: Could not find HTML template '{TEMPLATE_FILE}'")
        return

    # 3. Connect to Server (Open connection once for efficiency)
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
    except Exception as e:
        print(f"Error connecting to email server: {e}")
        return

    # 4. Iterate through the list
    for index, row in df.iterrows():
        nombre = str(row['Nombre'])
        apellido = str(row['Apellido'])
        dni = str(row['DNI']).strip()
        recipient_email = row['Email']

        print(f"Processing: {nombre} {apellido} (DNI: {dni})...")

        # Search for PDF
        search_pattern = os.path.join(PDF_FOLDER, f"*_{dni}.pdf")
        found_files = glob.glob(search_pattern)

        if not found_files:
            print(f"  [!] PDF not found for DNI {dni}. Skipping.")
            continue
        
        pdf_path = found_files[0]
        pdf_filename = os.path.basename(pdf_path)

        # 5. Customize the HTML Content
        # We replace the {{placeholders}} with actual data
        body_html = html_template.replace("{{nombre}}", nombre)
        body_html = body_html.replace("{{apellido}}", apellido)
        body_html = body_html.replace("{{dni}}", dni)

        # Create the Email Object
        msg = EmailMessage()
        msg['Subject'] = "Certificado Congreso Docente 2025"
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email

        # Set a plain text fallback (for email clients that don't read HTML)
        #msg.set_content(f"Hello {nombre}, please find attached your certificate for DNI {dni}.")
        
        # Add the HTML version
        msg.add_alternative(body_html, subtype='html')

        # Attach PDF
        with open(pdf_path, 'rb') as f:
            file_data = f.read()
            msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=pdf_filename)

        # Send
        try:
            server.send_message(msg)
            print(f"  [+] Email sent to {recipient_email}")
        except Exception as e:
            print(f"  [!] Failed to send to {recipient_email}: {e}")

    # Close connection
    server.quit()
    print("Done.")

if __name__ == "__main__":
    send_emails()