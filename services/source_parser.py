from openpyxl import load_workbook
from datetime import datetime


def parse_sheet(filepath, sheet_name):
    wb = load_workbook(filename=filepath)
    
    sheet = wb[sheet_name]

    def extract_data(sheet):
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)
        return data

    data_sheet = extract_data(sheet)

    return data_sheet

def create_formation(filepath):
    formation = {}
    try:
        data = parse_sheet(filepath, 'formation')
        for row in data:
            formation[row[0].lower()] = row[1]
    except Exception as e:
        print(f"Erreur: {e}")
    
    days, month, year = formation['date'].split('/')
    day_start = days
    day_end = None

    if ("-" in days):
        day_start, day_end = days.split('-')
    
    formation['date_debut'] = datetime.strptime(f"{day_start}/{month}/{year}", "%d/%m/%y")
    
    if (day_end is not None):
        formation['date_fin'] = datetime.strptime(f"{day_end}/{month}/{year}", "%d/%m/%y")
    
    return formation

def create_participants(filepath):
    participants = []
    try:
        data = parse_sheet(filepath, 'participants')        
        for row in data:
            participant = {}
            participant['civilite'] = row[0]
            participant['nom_complet'] = row[1]
            participant['email'] = row[2]
            participant['rpps'] = row[3]
            participant['phone'] = row[4]
            participant['financement'] = row[5]
            participants.append(participant)
    except Exception as e:
        print(f"Erreur: {e}")

    return participants