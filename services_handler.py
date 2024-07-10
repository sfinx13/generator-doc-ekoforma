import os
from openpyxl.styles import Font, PatternFill
import services.source_parser as source_parser
import services.timesheet_generator as timesheet_generator
import services.attendance_certificates as attendance_certificates


def generate_timesheet_zoom():
    filepath = 'assets/source/'

    for filename in os.listdir(filepath):
        formation = source_parser.create_formation(filepath + filename)
        participants = source_parser.create_participants(filepath + filename)

        wb, ws = timesheet_generator.create_zoom_timesheet(filename, formation, participants)

        gray_fill = PatternFill(start_color='DCDCDC', end_color='DCDCDC', fill_type='solid')
        bold_font = Font(name='Calibri', size=11, bold=True)
        for row in range(3, ws.max_row + 1):
            cell = ws.cell(row=row, column=1)
            if cell.value and cell.value != 'N° de réunion' and not cell.value.endswith('zoom'):
                cell.fill = gray_fill
                cell.font = bold_font
                if cell.value == 'empty':
                    cell.value = ''

        wb.save("public/zoom_timesheet_{}".format(filename))
        print("zoom_timesheet_{} generated".format(filename))

    print('Timesheet generated done!')


def generate_attendance_certificates():
    filepath = 'assets/source/'

    for filename in os.listdir(filepath):
        formation = source_parser.create_formation(filepath + filename)
        participants = source_parser.create_participants(filepath + filename)
        formation_titre = formation['titre'].replace('\n', '')
        print(f"Formation - {formation_titre}")
        
        for participant in participants:
            attendance_certificates.generate_attendance_certificate(participant, formation)
            print(f"    Attestation de présence généré pour {participant['nom_complet']}")
