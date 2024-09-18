import os
import subprocess
import services.source_parser as source_parser
import services.timesheet_generator as timesheet_generator
from services.attendance_certificates_generator import generate_attendance_certificate, merge_pdfs

pdf_files = []

def generate_timesheet_zoom():
    filepath = 'uploads/'
    full_meetings_and_participants = {}

    for filename in os.listdir(filepath):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            formation = source_parser.create_formation(filepath + filename)
            participants = source_parser.create_participants(filepath + filename)

            if len(formation) == 0 or len(participants) == 0:
                continue

            wb, ws, full_meetings_and_participants = timesheet_generator.create_zoom_timesheet(filename, formation, participants)
            
            # Appel de la fonction pour générer les tableaux pour chaque demi-journée
#            virtualclass_synthese_generator.generate_tables_for_each_meeting(filename, full_meetings_and_participants)


            output_excel = "downloads/{}_zoom_timesheet_{}".format(formation['code'], filename)
            wb.save(output_excel)
            convert_excel_to_pdf(output_excel)
            print("zoom_timesheet_{} generated".format(filename))
        else:
            print(filename, 'cannot be used!')
    
    print('Timesheet generated done!')

    return full_meetings_and_participants

def generate_attendance_certificates():
    filepath = 'uploads/'

    for filename in os.listdir(filepath):
        if filename.endswith('.xlsx'):
            formation = source_parser.create_formation(filepath + filename)
            participants = source_parser.create_participants(filepath + filename)
            if len(formation) == 0 or len(participants) == 0:
                continue

            formation_titre = formation['titre'].replace('\n', '')
            print(f"Formation - {formation_titre}")
            
            for participant in participants:
                print(f"Attestation de présence généré pour {participant['nom_complet']}")
                pdf_file = generate_attendance_certificate(participant, formation)
                pdf_files.append(pdf_file)


            # Fusionner tous les PDF générés pour la formation
            merge_pdfs(formation['code'], 'downloads/')


def convert_excel_to_pdf(input_file):
    """
    Convertit un fichier Excel en PDF en utilisant LibreOffice en mode headless.
    """
    try:
        # Créer le chemin de sortie en remplaçant l'extension par ".pdf"
        output_file = input_file.replace(".xlsx", ".pdf").replace(".xls", ".pdf")
        
        # Commande LibreOffice en mode headless pour convertir en PDF
        command = [
            "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(output_file), input_file
        ]

        # Exécuter la commande LibreOffice
        subprocess.run(command, check=True)
        
        print(f"Conversion réussie : {output_file}")
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de la conversion de {input_file} en PDF : {e}")