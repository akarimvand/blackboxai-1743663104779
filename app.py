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
        
        # Extract contact information
        contacts = []
        for vcard in vcards:
            contact = {
                'Name': getattr(vcard, 'fn', None) and vcard.fn.value or '',
                'Phone': getattr(vcard, 'tel', None) and vcard.tel.value or '',
                'Email': getattr(vcard, 'email', None) and vcard.email.value or '',
                'Address': getattr(vcard, 'adr', None) and format_address(vcard.adr.value) or ''
            }
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
    app.run(debug=True)