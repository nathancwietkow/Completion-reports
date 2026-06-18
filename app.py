#!/usr/bin/env python3
"""
Flask app for Completion Report Generator with Full Integration
"""

from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Inches, Pt
import os
from pathlib import Path
import json
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# ===== COATINGS DATA =====
COATINGS = {
    "resichem_507": {
        "name": "RESICHEM 507 DWPU",
        "benefits": [
            "WRAS Approved",
            "Complete Solvent Free Technology",
            "Extremely Fast Curing Times, even at low temperatures",
            "High degree of flexibility @ > 35%, capable of accommodating structural movement compared with resin / epoxy / bitumastic coatings of <1.0%",
            "Superb adhesion to steel, concrete, GRP and many other substrates",
            "High levels of impact and chemical resistance",
            "Reduced downtime of storage tanks and process vessels - refill after minimum of 6hrs after completion",
            "Long life performance with minimal maintenance",
            "Easy clean, ceramic tile-like finish",
            "Application by brush / roller or plural component spray equipment",
            "Excellent track record for applications such as drinking water tanks, cooling towers, reservoirs etc."
        ]
    },
    "maxline_100": {
        "name": "MAXLINE 100 DWPU",
        "benefits": [
            "WRAS Approved",
            "Complete Solvent Free Technology",
            "High degree of flexibility @ > 35%, capable of accommodating structural movement compared with resin / epoxy / bitumastic coatings of <1.0%",
            "Superb adhesion to steel, concrete, GRP and many other substrates",
            "High levels of impact and chemical resistance",
            "Long life performance with minimal maintenance",
            "Easy clean, ceramic tile-like finish",
            "Application by brush / roller or plural component spray equipment",
            "Excellent track record for applications such as drinking water tanks, cooling towers, reservoirs etc."
        ]
    },
    "dwpu": {
        "name": "DRINKING WATER POLYURETHANE (DWPU)",
        "benefits": [
            "WRAS Approved",
            "Complete Solvent Free Technology",
            "High degree of flexibility @ > 35%, capable of accommodating structural movement compared with resin / epoxy / bitumastic coatings of <1.0%",
            "Superb adhesion to steel, concrete, GRP and many other substrates",
            "High levels of impact and chemical resistance",
            "Long life performance with minimal maintenance",
            "Easy clean, ceramic tile-like finish",
            "Excellent track record for applications such as drinking water tanks, cooling towers, reservoirs etc."
        ]
    }
}

# ===== ADDITIONAL PRODUCTS DATA =====
ADDITIONAL_PRODUCTS = {
    "jotamastic_90": {
        "name": "JOTAMASTIC 90",
        "reason": "to use in conjunction with JOTUN HARDTOP XP"
    },
    "jotun_hardtop": {
        "name": "JOTUN HARDTOP XP",
        "reason": "which is a long-lasting topcoat that provides UV and weather resistant protection"
    }
}

# ===== PROCESS STEPS DATA =====
PROCESS_STEPS = {
    "surface_prep": "The surfaces were manually prepared to provide a surface profile of 0.75mm to accept the new coating.",
    "repairs": "Localised repairs were completed to restore structural strength and provide increased protection.",
    "sealing": "Sealing was applied to provide structural strength and improve substrate adhesion of the coating.",
    "basecoat": "A basecoat of [coating] was then applied to all internal surfaces at a nominal thickness of 300-500 micron.",
    "topcoat": "A topcoat of [coating] was then applied to all internal surfaces at a nominal thickness of 300-500 micron.",
    "refill": "The system was disinfected, and refilled, ready to go back online.",
    "other_tanks": "The same methodology was applied to [asset_name].",
    "remedials": "The following remedial works were also completed during this project."
}

# ===== TANK TYPES & SUBTYPES =====
TANK_TYPES = {
    "gsm": "Galvanised Steel Tank (GSM)",
    "grp": "Glass Reinforced Plastic Tank (GRP)",
    "concrete": "Concrete Tank",
    "plastic": "Plastic Tank",
    "asbestos": "Asbestos Tank",
    "other": "Other"
}

TANK_SUBTYPES = {
    "one_piece": "One Piece",
    "sectional_tif": "Sectional – Totally Internally Flanged",
    "sectional_pif": "Sectional – Partially Internally Flanged",
    "sectional_ef": "Sectional – Externally Flanged",
    "sectional": "Sectional",
    "other": "Other"
}

# ===== MAJOR PROJECTS =====
MAJOR_PROJECTS = {
    "full_reline": "Full Tank Reline",
    "bolt_reline": "Bolt restoration and lining",
    "steel_reline": "Steel restorations and lining",
    "clean_disinfect": "Tank clean and disinfections",
    "tank_install": "Tank installation",
    "tank_remove": "Tank removal",
    "crack_repair": "Repairing cracks/patch repairs on fibreglass tanks",
    "dividing_wall": "Dividing wall repairs on duplex fibreglass tanks",
    "tanking_floor": "Tanking plant room floors",
}

# ===== REMEDIAL WORKS =====
REMEDIAL_WORKS = {
    "pipework_reroute": "Pipework rerouting",
    "vent_reroute": "Vent back rerouting",
    "pipework_disinfect": "Pipework disinfections",
    "outlet_flush": "Outlet flushing",
    "insulation_wrap": "Insulation wrapping pipework/tank",
    "lid_replace": "Lid replacement",
    "lid_vents": "Installation of lid vents and/or overflow screens",
    "support_replace": "Hollow support replacements",
    "fibreglass_repair": "Fibreglass repair and strengthening",
    "access_hatches": "Installing Access Hatches",
}

def generate_project_brief(company_name, major_project, tank_quantity, tank_type, tank_location, coating_id, additional_products_list):
    """Generate dynamic project brief based on selections"""

    coating = COATINGS.get(coating_id, {})
    coating_name = coating.get('name', 'coating')
    coating_benefits = coating.get('benefits', [])

    benefits_text = "\n".join([f"• {benefit}" for benefit in coating_benefits])

    project_brief = f"{company_name} were asked to provide a solution to {major_project} {tank_quantity} {tank_type} located in {tank_location}. We proposed that each tank be given long term protection by the application of a solvent free polyurethane coating.\n\n"

    project_brief += f"The product employed was {coating_name}, which offers numerous economic, technical and environmental features and benefits, some of which are highlighted below:\n\n{benefits_text}\n\n"

    if additional_products_list:
        for product_id in additional_products_list:
            product = ADDITIONAL_PRODUCTS.get(product_id, {})
            if product:
                product_name = product.get('name', '')
                product_reason = product.get('reason', '')
                project_brief += f"Additionally, {product_name} was applied {product_reason}.\n\n"

    project_brief += "The following pages show the methodology employed:"

    return project_brief

@app.route('/')
def index():
    return render_template('index.html',
                         tank_types=TANK_TYPES,
                         tank_subtypes=TANK_SUBTYPES,
                         coatings=COATINGS,
                         additional_products=ADDITIONAL_PRODUCTS,
                         major_projects=MAJOR_PROJECTS,
                         remedial_works=REMEDIAL_WORKS)

@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """Generate completion report from form data"""
    try:
        data = request.form
        files = request.files

        job_details = {
            'project_number': data.get('projectNumber', ''),
            'customer_name': data.get('customerName', ''),
            'address_line1': data.get('addressLine1', ''),
            'address_line2': data.get('addressLine2', ''),
            'address_line3': data.get('addressLine3', ''),
            'postcode': data.get('postcode', ''),
            'completion_date': data.get('completionDate', ''),
            'tank_quantity': data.get('tankQuantity', '1'),
            'tank_location': data.get('tankLocation', ''),
        }

        major_project = data.get('majorProject', '')
        tank_type = data.get('tankType', '')
        work_completed_by = data.get('workCompletedBy', 'key_environmental')
        subcontractor_name = data.get('subcontractorName', '')
        coating_id = data.get('coating', '')
        additional_products = data.getlist('additionalProducts[]')
        remedial_works = data.getlist('remedialWorks[]')
        include_guarantee = data.get('includeGuarantee') == 'true'
        include_disinfection = data.get('includeDisinfection') == 'true'

        if work_completed_by == 'subcontracted' and subcontractor_name:
            company_name = subcontractor_name
            footer_line = f"For and on behalf of {subcontractor_name}"
        else:
            company_name = "Key Environmental Services Ltd"
            footer_line = "For and on behalf of Key Environmental Services Ltd"

        project_brief_text = generate_project_brief(
            company_name,
            MAJOR_PROJECTS.get(major_project, major_project),
            job_details['tank_quantity'],
            TANK_TYPES.get(tank_type, tank_type),
            job_details['tank_location'],
            coating_id,
            additional_products
        )

        remedial_names = [REMEDIAL_WORKS.get(w, w) for w in remedial_works]

        template_path = 'completion_report_template.docx'
        if not os.path.exists(template_path):
            return jsonify({'error': 'Template file not found'}), 400

        doc = Document(template_path)

        replacements = {
            '[CUSTOMER_NAME]': job_details['customer_name'],
            '[ADDRESS_LINE_1]': job_details['address_line1'],
            '[ADDRESS_LINE_2]': job_details['address_line2'],
            '[ADDRESS_LINE_3]': job_details['address_line3'],
            '[POSTCODE]': job_details['postcode'],
            '[PROJECT_NUMBER]': job_details['project_number'],
            '[COMPLETION_DATE]': job_details['completion_date'],
            '[COMPANY_NAME]': company_name,
            '[PROJECT_BRIEF_DESCRIPTION]': project_brief_text,
            '[SIGNATURE_NAME]': 'CON O\'SHEA',
            '[SIGNATURE_TITLE]': 'DIRECTOR',
            '[FOOTER_LINE]': footer_line,
        }

        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                for key, value in replacements.items():
                    if key in run.text:
                        run.text = run.text.replace(key, value)

        if remedial_names:
            remedial_text = "Remedial works completed: " + ", ".join(remedial_names)
            for paragraph in doc.paragraphs:
                if '[REMEDIAL_WORKS_DESCRIPTION]' in paragraph.text:
                    for run in paragraph.runs:
                        run.text = run.text.replace('[REMEDIAL_WORKS_DESCRIPTION]', remedial_text)

        output_filename = f"{job_details['project_number']} - {job_details['customer_name']} - Completion Report.docx"
        output_path = os.path.join('reports', output_filename)

        os.makedirs('reports', exist_ok=True)
        doc.save(output_path)

        return jsonify({
            'success': True,
            'filename': output_filename,
            'filepath': output_path
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<path:filepath>')
def download_report(filepath):
    """Download generated report"""
    try:
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
