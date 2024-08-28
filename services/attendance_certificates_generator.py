from docx import Document
from datetime import datetime, timedelta
import os

doc_path = 'static/IDEL_ATTESTATION_DE_PARTICIPATION_A_UN_PROGRAMME_DE_DPC.docx'

def generate_attendance_certificate(participant, formation):
    doc = Document(doc_path)
    formation_titre = formation['titre'].replace('\n', '')
    formation_orientation = formation['orientation']
    formation_date_debut = formation.get('date_debut').strftime('%d/%m/%Y')
    if formation.get('date_fin'):
        formation_date_fin = formation.get('date_fin').strftime('%d/%m/%Y')
        date_signature = formation.get('date_fin') + timedelta(days=1)
    else:
        formation_date_fin = formation_date_debut
        date_signature = formation.get('date_debut') + timedelta(days=1)

    

    info_to_add = {
        "Nom :": participant['nom'].upper(),
        "Prénom :": participant['prenoms'].upper(),
        "Adresse électronique :": participant['email'],
        "N° RPPS :": participant['rpps'],
        "DD/DD/DDDD": formation_date_debut,
        "FF/FF/FFFF": formation_date_fin,
        "Année(s) civile(s) de participation :": datetime.now().year,
        "Intitulé du programme :": formation_titre,
        "Orientation nationale dans laquelle le programme s’inscrit :": formation_orientation,
        "Nom/sigle :": formation['nom'],
        "Adresse :": formation['adresse'],
        "N° enregistrement OGDPC / Agence nationale du DPC :": formation['numero dpc'],
        "//2024": date_signature.strftime("%d/%m/%Y")
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
                elif placeholder == "N° RPPS :":
                    # Séparer le texte en deux parties, avant et après le placeholder
                    parts = paragraph.text.split(placeholder)
                
                    # Nettoyer le paragraphe
                    paragraph.clear()

                    # Ajouter la première partie (avant le placeholder)
                    if parts[0]:
                        paragraph.add_run(parts[0])


                    # Ajouter le placeholder sans gras
                    paragraph.add_run(placeholder)

                    # Ajouter le numéro RPPS en gras
                    actual_value = str(actual_value)
                    run = paragraph.add_run(actual_value)
                    run.bold = True

                    # Ajouter le reste du texte sans gras (après le numéro RPPS)
                    remaining_text = parts[1].replace(actual_value, '')
                    paragraph.add_run(remaining_text)
                elif placeholder == "DD/DD/DDDD" or placeholder == "FF/FF/FFFF":
                    # Si le paragraphe contient des dates à remplacer

                    # Remplacer d'abord la date de début, puis la date de fin en nettoyant le texte mais en réécrivant proprement
                    text_parts = paragraph.text.split("DD/DD/DDDD")

                    # Nettoyer le paragraphe
                    paragraph.clear()

                    # Ajouter le texte avant la date de début, puis la date de début en gras
                    paragraph.add_run(text_parts[0])
                    run = paragraph.add_run(formation_date_debut)
                    run.bold = True

                    # Si la phrase continue avec "Date de fin :", traiter cela
                    if len(text_parts) > 1:
                        remaining_text = text_parts[1].split("FF/FF/FFFF")
                        
                        # Ajouter le texte entre les deux dates (souvent " Date de fin :")
                        paragraph.add_run(remaining_text[0])

                        # Ajouter la date de fin en gras
                        run = paragraph.add_run(formation_date_fin)
                        run.bold = True

                        # Ajouter tout texte restant après la date de fin
                        if len(remaining_text) > 1:
                            paragraph.add_run(remaining_text[1])
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