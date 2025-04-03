from flask import Flask, render_template, request, send_file
import vobject
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'vcf_file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    
    vcf_file = request.files['vcf_file']
    if vcf_file.filename == '':
        return {'error': 'No file selected'}, 400
    
    if not vcf_file.filename.lower().endswith('.vcf'):
        return {'error': 'Invalid file type. Please upload a .vcf file'}, 400
    
    try:
        # Read and parse VCF file
        vcf_data = vcf_file.read().decode('utf-8')
        vcards = vobject.readComponents(vcf_data)
        
        # Extract all available contact information
        contacts = []
        for vcard in vcards:
            contact = {}
            # Standard fields
            contact['Name'] = getattr(vcard, 'fn', None) and vcard.fn.value or ''
            contact['Organization'] = getattr(vcard, 'org', None) and vcard.org.value or ''
            contact['Title'] = getattr(vcard, 'title', None) and vcard.title.value or ''
            
            # Handle multiple phone numbers - each in separate column
            for tel in getattr(vcard, 'tel_list', []):
                phone_type = tel.type_param.lower() if hasattr(tel, 'type_param') else 'other'
                contact[f'Phone_{phone_type}'] = tel.value
            
            # Handle multiple emails
            emails = []
            for email in getattr(vcard, 'email_list', []):
                emails.append(email.value)
            contact['Emails'] = ' | '.join(emails) if emails else ''
            
            # Handle addresses
            addresses = []
            for adr in getattr(vcard, 'adr_list', []):
                addresses.append(format_address(adr.value))
            contact['Addresses'] = ' | '.join(addresses) if addresses else ''
            
            # Additional fields
            contact['Note'] = getattr(vcard, 'note', None) and vcard.note.value or ''
            contact['URL'] = getattr(vcard, 'url', None) and vcard.url.value or ''
            contact['Birthday'] = getattr(vcard, 'bday', None) and vcard.bday.value or ''
            
            contacts.append(contact)
        
        # Create DataFrame and Excel file
        df = pd.DataFrame(contacts)
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='contacts.xlsx'
        )
        
    except Exception as e:
        return {'error': f'Error processing file: {str(e)}'}, 500

def format_address(address):
    if not address:
        return ''
    return ', '.join(filter(None, [
        address.street,
        address.city,
        address.region,
        address.country
    ]))

if __name__ == '__main__':
    app.run(debug=True, port=8000)
