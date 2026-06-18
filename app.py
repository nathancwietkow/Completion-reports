#!/usr/bin/env python3
from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Inches
import os
from io import BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

COATINGS = {
    "resichem_507": {"name": "RESICHEM 507 DWPU", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "Extremely Fast Curing Times", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Fast refill - 6hrs minimum", "Long life performance", "Easy clean finish", "Brush/roller or spray", "Excellent track record"]},
    "maxline_100": {"name": "MAXLINE 100 DWPU", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Long life performance", "Easy clean finish", "Brush/roller or spray", "Excellent track record"]},
    "dwpu": {"name": "DRINKING WATER POLYURETHANE (DWPU)", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Long life performance", "Easy clean finish", "Excellent track record"]}
}

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

def replace_text_in_paragraph(paragraph, key, value):
    """Replace text in a paragraph."""
    if key in paragraph.text:
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)

def replace_placeholder_with_photos(doc, placeholder, photo_files):
    """Replace a placeholder with multiple photos."""
    if not photo_files or len(photo_files) == 0:
        return

    for para in doc.paragraphs:
        if placeholder in para.text:
            # Clear the placeholder
            para.clear()

            # Insert all photos
            for photo_file in photo_files:
                if photo_file and photo_file.filename:
                    # Save uploaded file temporarily
                    temp_path = f"/tmp/{photo_file.filename}"
                    photo_file.save(temp_path)

                    # Add image to document
                    run = para.add_run()
                    try:
                        run.add_picture(temp_path, width=Inches(5.5))
                    except:
                        run.text = "[Photo could not be loaded]"

                    # Add line break between photos
                    para.add_run('\n')

                    # Clean up
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
        files = request.files

        site_name = data.get('siteName', '')
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

        # Get asset issues
        issues = data.getlist('assetIssues[]')

        # Get process steps selected
        process_steps = data.getlist('processSteps[]')

        if not os.path.exists('completion_report_template.docx'):
            return jsonify({'error': 'Template not found'}), 400

        doc = Document('completion_report_template.docx')

        # Replace simple text placeholders
        replace_map = {
            '[MAIN_PROJECT]': major_project,
            '[ASSET_QUANTITY]': asset_quantity,
            '[ASSET_SUBCATEGORY]': asset_subcategory,
            '[ASSET_TYPE]': asset_type,
            '[ASSET_LOCATION]': asset_location,
            '[PROJECT_NUMBER]': project_number,
            '[COMPLETION_DATE]': completion_date,
            '[SITE_NAME]': site_name,
            '[COMPANY_NAME]': company,
        }

        # Replace all text placeholders
        for para in doc.paragraphs:
            for key, value in replace_map.items():
                replace_text_in_paragraph(para, key, value)

        # Handle [ISSUES_RAISED] - replace each instance with an issue
        for para in doc.paragraphs:
            if '[ISSUES_RAISED]' in para.text:
                issue_index = 0
                for run in para.runs:
                    if '[ISSUES_RAISED]' in run.text:
                        if issue_index < len(issues):
                            run.text = run.text.replace('[ISSUES_RAISED]', issues[issue_index], 1)
                            issue_index += 1
                        else:
                            run.text = run.text.replace('[ISSUES_RAISED]', '')

        # Handle descriptions for each process step
        descriptions = {
            'before': '[BEFORE_DESCRIPTION]',
            'surface_prep': '[PREPARATION_DESCRIPTION]',
            'basecoat': '[BASE_COAT_DESCRIPTION]',
            'topcoat': '[TOP_COAT_DESCRIPTION]',
        }

        # Placeholder text for descriptions (user can customize in template)
        default_descriptions = {
            'before': 'Initial condition of asset prior to works.',
            'surface_prep': 'Surface preparation work completed.',
            'basecoat': 'Basecoat application.',
            'topcoat': 'Topcoat application.',
        }

        for step, placeholder in descriptions.items():
            if step in process_steps:
                desc = default_descriptions.get(step, '')
                for para in doc.paragraphs:
                    replace_text_in_paragraph(para, placeholder, desc)

        # Handle remedial works description
        remedial = [REMEDIAL_WORKS.get(w, w) for w in data.getlist('remedialWorks[]')]
        if remedial:
            remedial_text = "Remedial works completed: " + ", ".join(remedial)
        else:
            remedial_text = ""

        for para in doc.paragraphs:
            replace_text_in_paragraph(para, '[REMEDIAL_WORKS_DESCRIPTION]', remedial_text)

        # Insert photos at placeholders
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
            if photo_files:
                replace_placeholder_with_photos(doc, template_placeholder, photo_files)

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
