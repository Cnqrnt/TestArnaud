from flask import Flask, render_template, request, send_file
from fpdf import FPDF
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import unicodedata

app = Flask(__name__)

def send_email_with_pdf(pdf_path, recipient_emails):
    sender_email = "X"  # Remplacez par votre adresse Gmail
    sender_password = "X"  # Remplacez par votre mot de passe Gmail

    # Création du message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = "Attestation de remise/restitution du matériel informatique"

    body = "Veuillez trouver ci-joint l'attestation en PDF."
    msg.attach(MIMEText(body, 'plain'))

    # Pièce jointe
    with open(pdf_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(pdf_path)}")
        msg.attach(part)

    # Connexion au serveur SMTP et envoi
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Serveur SMTP de Gmail
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_emails, text)
        server.quit()
        print("Email envoyé avec succès.")
    except Exception as e:
        print(f"Erreur lors de l'envoi de l'email: {e}")

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        'name': request.form.get('name', ''),
        'direction': request.form.get('direction', ''),
        'site': request.form.get('site', ''),
        'matricule': request.form.get('matricule', ''),
        'equipment': request.form.getlist('equipment'),
        'equipment_status': request.form.get('equipment_status', ''),
        'serial_number_ordinateur': request.form.get('serial_number_ordinateur', ''),
        'dsi_code_ordinateur': request.form.get('dsi_code_ordinateur', ''),
        'serial_number_telephone': request.form.get('serial_number_telephone', ''),
        'dsi_code_telephone': request.form.get('dsi_code_telephone', ''),
        'serial_number_imprimante': request.form.get('serial_number_imprimante', ''),
        'dsi_code_imprimante': request.form.get('dsi_code_imprimante', ''),
        'serial_number_casque': request.form.get('serial_number_casque', ''),
        'dsi_code_casque': request.form.get('dsi_code_casque', ''),
        'serial_number_autre': request.form.get('serial_number_autre', ''),
        'dsi_code_autre': request.form.get('dsi_code_autre', ''),
        'support_steps': request.form.getlist('support_steps'),
        'return_steps': request.form.getlist('return_steps'),
        'date_lieu': request.form.get('date', ''),
        'tech_signature': request.form.get('tech_signature', ''),
        'agent_signature': request.form.get('agent_signature', '')
    }


    def normalize_name(name):
        # Supprimer les accents et convertir en ASCII
        name = unicodedata.normalize('NFKD', name).encode('ASCII', 'ignore').decode('utf-8')
        # Séparer le prénom et le nom par un point et utiliser des tirets pour les parties composées
        name_parts = name.split()
        if len(name_parts) > 1:
            name = name_parts[0] + "." + "-".join(name_parts[1:])
        name = name.lower()
        return name

    prenom_nom = normalize_name(data['name'])
    pdf_filename = f"{prenom_nom} - {data['date_lieu']}.pdf"

    # Création du PDF
    pdf = FPDF()
    pdf.add_page()
    page_width = pdf.w
    logo_width = 20
    x_position = (page_width - logo_width) / 2
    pdf.image('logo.png', x=x_position, y=10, w=logo_width)

    # Move cursor down to start text after the logo
    pdf.ln(20)

    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)

    # Title section with a border
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, txt="Attestation de remise/restitution du matériel informatique", ln=True, align='C', fill=True)
    pdf.ln(5)

    # Agent information section
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "Informations de l'agent:", ln=True, border=1, fill=True)
    pdf.set_font("Arial", size=12)
    pdf.cell(100, 8, f"Nom et prénom: {data['name']}", border=1)
    pdf.cell(90, 8, f"Direction: {data['direction']}", border=1, ln=True)
    pdf.cell(100, 8, f"Site de l'agent: {data['site']}", border=1)
    pdf.cell(90, 8, f"Matricule agent: {data['matricule']}", border=1, ln=True)
    pdf.ln(5)

    # Equipment Allocation with additional fields
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "1. Attribution ou réstitution du matériel informatique:", ln=True, border=1, fill=True)
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 8, f"Statut du matériel: {data['equipment_status']}", ln=True, border=1)

    # Add equipment details in a table format
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(60, 8, "Matériel", 1, 0, 'C')
    pdf.cell(65, 8, "Numéro de série", 1, 0, 'C')
    pdf.cell(65, 8, "Code DSI", 1, 1, 'C')

    pdf.set_font("Arial", size=12)
    equipment_details = {
        'Ordinateur': {'serial_number': data['serial_number_ordinateur'], 'dsi_code': data['dsi_code_ordinateur']},
        'Téléphone': {'serial_number': data['serial_number_telephone'], 'dsi_code': data['dsi_code_telephone']},
        'Imprimante': {'serial_number': data['serial_number_imprimante'], 'dsi_code': data['dsi_code_imprimante']},
        'Casque audio': {'serial_number': data['serial_number_casque'], 'dsi_code': data['dsi_code_casque']},
        'Autre': {'serial_number': data['serial_number_autre'], 'dsi_code': data['dsi_code_autre']}
    }

    for item, details in equipment_details.items():
        checkbox = "Oui - " if item in data['equipment'] else "Non - "
        pdf.cell(60, 8, f"{checkbox} {item}", 1)
        pdf.cell(65, 8, details['serial_number'], 1)
        pdf.cell(65, 8, details['dsi_code'], 1, 1)

    pdf.ln(5)

    # Support steps section
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "2. Accompagnement des agents - Lors de la dotation:", ln=True, border=1, fill=True)
    pdf.set_font("Arial", size=12)
    support_step_list = [
        "Documentation nouvel arrivant",
        "Charte informatique",
        "Comment contacter le support informatique",
        "Sensibilisation aux données",
        "Présentation de la suite Office 365",
        "Présentation de l'intranet Vikings",
        "Vérification du compte informatique",
        "Configuration badge impression",
        "Information sur les demandes de logiciels",
        "Configuration de la double authentification",
        "Renseignement du matricule et données agent"
    ]
    for step in support_step_list:
        checkbox = "Oui - " if item in data['support_steps'] else "Non - "
        pdf.cell(0, 8, f"{checkbox} {step}", ln=True, border=1)
    pdf.ln(5)

    # Steps during return
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "3. Lors de la restitution du matériel:", ln=True, border=1, fill=True)
    pdf.set_font("Arial", size=12)
    return_step_list = [
        "Sauvegarde des données",
        "Avertissement sur accès aux données pro par hiérarchie",
        "Message d'absence dans la messagerie",
        "Renvoi de ligne téléphonique",
        "Résiliation abonnement téléphonie mobile"
    ]
    for step in return_step_list:
        checkbox = "Oui - " if item in data['return_steps'] else "Non - "
        pdf.cell(0, 8, f"{checkbox} {step}", ln=True, border=1)

    # Signature section
    pdf.ln(15)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "Signatures :", ln=True, border=1, fill=True)
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 8, f"Date: {data['date_lieu']}", ln=True, border=1)
    pdf.cell(0, 8, f"Signature du technicien: {data['tech_signature']}", ln=True, border=1)
    pdf.cell(0, 8, f"Signature de l'agent: {data['agent_signature']}", ln=True, border=1)

    # Définir le chemin du fichier PDF et le sauvegarder
    pdf_file_path = os.path.join(os.getcwd(), pdf_filename)
    pdf.output(pdf_file_path)

    # Définir la liste des destinataires
    recipient_emails = [f"{prenom_nom}@normandie.fr", "X@normandie.fr"]  # Ajoutez d'autres emails si nécessaire

    # Envoyer le PDF par email
    send_email_with_pdf(pdf_file_path, recipient_emails)

    # Retourner le fichier à télécharger
    return send_file(pdf_file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)