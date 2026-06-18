#!/usr/bin/env python3
from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

COATINGS = {
    "resichem_507": {"name": "RESICHEM 507 DWPU", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "Extremely Fast Curing Times", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Fast refill - 6hrs minimum", "Long life performance", "Easy clean finish", "Brush/roller or spray", "Excellent track record"]},
    "maxline_100": {"name": "MAXLINE 100 DWPU", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Long life performance", "Easy clean finish", "Brush/roller or spray", "Excellent track record"]},
    "dwpu": {"name": "DRINKING WATER POLYURETHANE (DWPU)", "benefits": ["WRAS Approved", "Complete Solvent Free Technology", "High flexibility", "Superb adhesion", "High impact/chemical resistance", "Long life performance", "Easy clean finish", "Excellent track record"]}
}

ADDITIONAL_PRODUCTS = {"jotamastic_90": {"name": "JOTAMASTIC 90", "reason": "to use with JOTUN HARDTOP XP"}, "jotun_hardtop": {"name": "JOTUN HARDTOP XP", "reason": "which is a long-lasting topcoat that provides UV and weather resistant protection"}}

REMEDIAL_WORKS = {"pipework_reroute": "Pipework rerouting", "vent_reroute": "Vent back rerouting", "pipework_disinfect": "Pipework disinfections", "outlet_flush": "Outlet flushing", "insulation_wrap": "Insulation wrapping", "lid_replace": "Lid replacement", "lid_vents": "Lid vents/overflow screens", "support_replace": "Support replacements", "fibreglass_repair": "Fibreglass repair", "access_hatches": "Installing Access Hatches"}

def generate_project_brief(company, project, qty, asset_type, location, coating_id, products, issues):
    coating = COATINGS.get(coating_id, {})
    asset_word = "asset" if qty == "1" else "assets"

    # Build brief based on whether coating was selected
    if coating_id:
        benefits = "\n".join([f"• {b}" for b in coating.get('benefits', [])])
        brief = f"{company} were asked to provide a solution to {project} {qty} {asset_type} {asset_word} located in {location}. We proposed that each asset be given long term protection by the application of a solvent free polyurethane coating.\n\n"
        brief += f"The product employed was {coating.get('name', 'coating')}, which offers numerous economic, technical and environmental features and benefits:\n\n{benefits}\n\n"
    else:
        brief = f"{company} were asked to provide a solution to {project} {qty} {asset_type} {asset_word} located in {location}.\n\n"

    # Add additional products if selected
    for pid in products:
        p = ADDITIONAL_PRODUCTS.get(pid, {})
        if p:
            brief += f"Additionally, {p.get('name', '')} was applied {p.get('reason', '')}.\n\n"

    brief += "The following pages show the methodology employed:"
    return brief

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

        job = {
            'project_number': data.get('projectNumber', ''),
            'site_name': site_name,
            'asset_type': asset_type,
            'asset_subcategory': asset_subcategory,
            'asset_location': asset_location,
            'asset_quantity': asset_quantity,
            'date': data.get('completionDate', ''),
        }

        work_by = data.get('workCompletedBy', 'key_environmental')
        subcon = data.get('subcontractorName', '')
        company = subcon if work_by == 'subcontracted' and subcon else "Key Environmental Services Ltd"
        footer = f"For and on behalf of {company}"

        # Get asset issues
        issues = data.getlist('assetIssues[]')
        issues_text = ", ".join(issues) if issues else "General maintenance"

        # Get coating (now optional)
        coating_id = data.get('coating', '')

        brief = generate_project_brief(
            company,
            data.get('majorProject', ''),
            asset_quantity,
            asset_type,
            asset_location,
            coating_id,
            data.getlist('additionalProducts[]'),
            issues
        )

        remedial = [REMEDIAL_WORKS.get(w, w) for w in data.getlist('remedialWorks[]')]

        if not os.path.exists('completion_report_template.docx'):
            return jsonify({'error': 'Template not found'}), 400

        doc = Document('completion_report_template.docx')

        replace = {
            '[CUSTOMER_NAME]': site_name,
            '[PROJECT_NUMBER]': job['project_number'],
            '[COMPLETION_DATE]': job['date'],
            '[COMPANY_NAME]': company,
            '[PROJECT_BRIEF_DESCRIPTION]': brief,
            '[SIGNATURE_NAME]': 'CON O\'SHEA',
            '[SIGNATURE_TITLE]': 'DIRECTOR',
            '[FOOTER_LINE]': footer,
            '[ASSET_ISSUES]': issues_text,
        }

        for para in doc.paragraphs:
            for run in para.runs:
                for k, v in replace.items():
                    if k in run.text:
                        run.text = run.text.replace(k, v)

        if remedial:
            text = "Remedial works completed: " + ", ".join(remedial)
            for para in doc.paragraphs:
                if '[REMEDIAL_WORKS_DESCRIPTION]' in para.text:
                    for run in para.runs:
                        run.text = run.text.replace('[REMEDIAL_WORKS_DESCRIPTION]', text)

        filename = f"{job['project_number']} - {site_name} - Completion Report.docx"
        path = os.path.join('reports', filename)
        os.makedirs('reports', exist_ok=True)
        doc.save(path)

        return jsonify({'success': True, 'filename': filename, 'filepath': path})
    except Exception as e:
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
