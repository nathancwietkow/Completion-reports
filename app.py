#!/usr/bin/env python3
from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

REMEDIAL_WORKS = {
    "pipework_reroute": "Pipework rerouting",
    "vent_reroute": "Vent back rerouting",
    "pipework_disinfect": "Pipework disinfections",
    "outlet_flush": "Outlet flushing",
    "insulation_wrap": "Insulation wrapping",
    "lid_replace": "Lid replacement",
    "lid_vents": "Lid vents/overflow screens",
    "support_replace": "Support replacements",
    "fibreglass_repair": "Fibreglass repair",
    "access_hatches": "Installing Access Hatches"
}

SUBCONTRACTORS = {
    "Echo Square Services Limited": "0800 7720 628"
}

def replace_text_in_document(doc, old_text, new_text):
    """Replace all instances of old_text with new_text in entire document."""
    for paragraph in doc.paragraphs:
        if old_text in paragraph.text:
            # Replace in runs
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

def replace_placeholder_with_photo(doc, placeholder, photo_files):
    """Replace a placeholder with photo(s)."""
    if not photo_files or len(photo_files) == 0:
        return

    for para in doc.paragraphs:
        if placeholder in para.text:
            # Get the index of this paragraph
            para_index = doc.paragraphs.index(para)

            # Clear the paragraph
            p = para._element
            p.getparent().remove(p)

            # Insert photos at this location
            parent = doc.paragraphs[0]._element.getparent()
            insert_index = list(parent).index(doc.paragraphs[para_index - 1]._element) + 1

            for photo_file in photo_files:
                if photo_file and photo_file.filename:
                    # Save temporarily
                    temp_path = f"/tmp/{photo_file.filename}"
                    photo_file.save(temp_path)

                    # Create new paragraph with image
                    new_para = doc.add_paragraph()
                    run = new_para.add_run()
                    try:
                        run.add_picture(temp_path, width=Inches(5.5))
                    except:
                        new_para.text = "[Photo could not be loaded]"

                    # Move new paragraph to correct position
                    parent.insert(insert_index, new_para._element)
                    insert_index += 1

                    # Cleanup
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
            break

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    try:
        data = request.form

        # Collect form values
        site_name = data.get('siteName', '')
        address_line_1 = data.get('addressLine1', '')
        address_line_2 = data.get('addressLine2', '')
        address_line_3 = data.get('addressLine3', '')
        postcode = data.get('postcode', '')

        asset_type = data.get('assetType', '')
        asset_subcategory = data.get('assetSubcategory', '')
        asset_location = data.get('assetLocation', '')
        asset_quantity = data.get('assetQuantity', '1')
        major_project = data.get('majorProject', '')
        project_number = data.get('projectNumber', '')
        completion_date = data.get('completionDate', '')

        work_by = data.get('workCompletedBy', 'key_environmental')
        subcon = data.get('subcontractorName', '')
        company = subcon if work_by == 'subcontracted' and subcon else "Key Environmental Services Ltd"

        # Set company contact details based on selection
        if work_by == 'key_environmental':
            company_contact_details = "01789 330830\nservice@key-environmental.co.uk\nwww.watertankrelining.co.uk"
        else:
            # Check if subcontractor is in our list
            if subcon in SUBCONTRACTORS:
                company_contact_details = SUBCONTRACTORS[subcon]
            else:
                company_contact_details = "[Subcontractor contact details to be added]"

        # Get issues list
        issues = data.getlist('assetIssues[]')

        if not os.path.exists('completion_report_template.docx'):
            return jsonify({'error': 'Template not found'}), 400

        doc = Document('completion_report_template.docx')

        # STEP 1: Replace all simple text placeholders
        replacements = {
            '[SITE_NAME]': site_name,
            '[ADDRESS_LINE_1]': address_line_1,
            '[ADDRESS_LINE_2]': address_line_2,
            '[ADDRESS_LINE_3]': address_line_3,
            '[POSTCODE]': postcode,
            '[MAIN_PROJECT]': major_project,
            '[ASSET_QUANTITY]': asset_quantity,
            '[ASSET_SUBCATEGORY]': asset_subcategory,
            '[ASSET_TYPE]': asset_type,
            '[ASSET_LOCATION]': asset_location,
            '[PROJECT_NUMBER]': project_number,
            '[COMPLETION_DATE]': completion_date,
            '[COMPANY_NAME]': company,
            '[COMPANY_CONTACT_DETAILS]': company_contact_details,
        }

        for placeholder, value in replacements.items():
            replace_text_in_document(doc, placeholder, value)

        # STEP 2: Handle [ISSUES_RAISED] - replace EACH instance with a different issue
        issue_count = 0
        for para in doc.paragraphs:
            while '[ISSUES_RAISED]' in para.text:
                if issue_count < len(issues):
                    for run in para.runs:
                        if '[ISSUES_RAISED]' in run.text:
                            run.text = run.text.replace('[ISSUES_RAISED]', issues[issue_count], 1)
                            issue_count += 1
                            break
                else:
                    # Remove remaining [ISSUES_RAISED] if not enough issues
                    for run in para.runs:
                        if '[ISSUES_RAISED]' in run.text:
                            run.text = run.text.replace('[ISSUES_RAISED]', '')
                    break

        # STEP 3: Handle remedial works
        remedial = [REMEDIAL_WORKS.get(w, w) for w in data.getlist('remedialWorks[]')]
        if remedial:
            remedial_text = "Remedial works completed: " + ", ".join(remedial)
        else:
            remedial_text = ""
        replace_text_in_document(doc, '[REMEDIAL_WORKS_DESCRIPTION]', remedial_text)

        # STEP 4: Insert photos
        photo_mapping = {
            'photo_front_building': '[FRONT_BUILDING_PHOTO]',
            'photo_before': '[BEFORE_PHOTO_PLACEHOLDER]',
            'photo_surface_prep': '[PREPARATION_PHOTO_PLACEHOLDER]',
            'photo_basecoat': '[BASE_COAT_PHOTO_PLACEHOLDER]',
            'photo_topcoat': '[TOP_COAT_PHOTO_PLACEHOLDER]',
            'photo_remedials': '[REMEDIAL_WORKS_PHOTO_PLACEHOLDER]',
        }

        for form_field, template_placeholder in photo_mapping.items():
            photo_files = request.files.getlist(form_field)
            if photo_files and any(f.filename for f in photo_files):
                replace_placeholder_with_photo(doc, template_placeholder, photo_files)

        # Save report
        filename = f"{project_number} - {site_name} - Completion Report.docx"
        path = os.path.join('reports', filename)
        os.makedirs('reports', exist_ok=True)
        doc.save(path)

        return jsonify({'success': True, 'filename': filename, 'filepath': path})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<path:filepath>')
def download(filepath):
    try:
        return send_file(filepath, as_attachment=True)
    except:
        return jsonify({'error': 'Download failed'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
