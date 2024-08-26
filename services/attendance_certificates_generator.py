from docx import Document
from datetime import datetime
import os

doc_path = 'static/INFIRMIER_ATTESTATION_DE_PARTICIPATION_A_UN_PROGRAMME_DE_DPC.docx'

def generate_attendance_certificate(participant, formation):
    doc = Document(doc_path)
    formation_titre = formation['titre'].replace('\n', '')
    formation_orientation = formation['orientation']
    formation_date_debut = formation.get('date_debut').strftime('%d/%m/%Y')
    if formation.get('date_fin'):
        formation_date_fin = formation.get('date_fin').strftime('%d/%m/%Y')
    else:
        formation_date_fin = formation_date_debut
    
    info_to_add = {
        "Nom :": participant['nom'].upper(),
        "Prénom :": participant['prenoms'].upper(),
        "Adresse électronique :": participant['email'],
        "N° RPPS :": participant['rpps'],
        "Date de début :": formation_date_debut,
        "Date de fin:": formation_date_fin,
        "Année(s) civile(s) de participation :": datetime.now().year,
        "Intitulé du programme :": formation_titre,
        "Orientation nationale dans laquelle le programme s’inscrit :": formation_orientation,
        "//2024": datetime.now().strftime("%d/%m/%Y")
    }

    # for paragraph in doc.paragraphs:
    #     for placeholder, actual_value in info_to_add.items():
    #         if placeholder in paragraph.text:
    #             paragraph.text = paragraph.text.replace(placeholder, actual_value)
    #             for run in paragraph.runs:
    #                 if actual_value in run.text:
    #                     run.bold = True


    for paragraph in doc.paragraphs:
        for placeholder, actual_value in info_to_add.items():
            
            if placeholder in paragraph.text:
                if placeholder == "//2024":
                    paragraph.text = paragraph.text.replace(placeholder, actual_value)
                else:
                    parts = paragraph.text.split(placeholder)
                    paragraph.text = parts[0] + placeholder
                    paragraph.add_run(f" {actual_value}").bold = True
                    if len(parts) > 1:
                        paragraph.add_run(parts[1]).bold = True

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_directory = os.path.join(script_dir, "../downloads")
    output_file = os.path.join(output_directory, f"../downloads/{formation.get('code')}_ATTESTATION_DE_PARTICIPATION_A_UN_PROGRAMME_DE_DPC_{participant['nom_complet']}.docx")                        

    doc.save(output_file)