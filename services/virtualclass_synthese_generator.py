from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def create_table(ws, start_row, start_col, title, data):
    # Calculer les cellules de début et de fin pour la fusion
    start_cell = ws.cell(row=start_row, column=start_col).coordinate
    end_cell = ws.cell(row=start_row, column=start_col + 6).coordinate
    ws.merge_cells(f'{start_cell}:{end_cell}')
    
    # Ajouter le titre
    ws[start_cell] = title

    # Appliquer les styles : police Calibri, taille 16, en gras, centré
    font = Font(name='Calibri', size=16, bold=True)
    alignment = Alignment(horizontal='center', vertical='center')
    
    ws[start_cell].font = font
    ws[start_cell].alignment = alignment
    
    # Définir une bordure de 2pt
    border_style = Side(style='medium', color='000000', border_style='medium')
    border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
    
    # Appliquer la bordure à toutes les cellules fusionnées
    for col in range(start_col, start_col + 7):
        cell = ws.cell(row=start_row, column=col)
        cell.border = border

    # Ajouter les en-têtes de colonne
    headers = data[0]
    for col_index, header in enumerate(headers, start=start_col):
        cell = ws.cell(row=start_row + 1, column=col_index)
        cell.value = header

    # Ajouter les données
    for row_index, row_data in enumerate(data[1:], start=start_row + 2):
        for col_index, cell_value in enumerate(row_data, start=start_col):
            ws.cell(row=row_index, column=col_index).value = cell_value
            # cell.value = cell_value

                        # Appliquer le style en gras pour les deuxième et troisième lignes
            if row_index in (start_row + 2 - 1, start_row + 3 - 1):
                cell.font = Font(bold=True)
            
            # Appliquer la couleur de fond et les styles à la quatrième ligne
            if row_index == start_row + 4 - 1:
                cell.font = Font(name='Calibri', size=12, bold=True)
                cell.fill = PatternFill(start_color='DEEBF7', end_color='DEEBF7', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')            

    # Définir les largeurs des colonnes
    ws.column_dimensions[ws.cell(row=start_row + 1, column=start_col).column_letter].width = 35
    ws.column_dimensions[ws.cell(row=start_row + 1, column=start_col + 1).column_letter].width = 40

        # Fusionner les colonnes spécifiées pour la quatrième ligne
    ws.merge_cells(start_row=start_row + 4, start_column=start_col + 2, end_row=start_row + 4, end_column=start_col + 3)

# Créer un nouveau Workbook
wb = Workbook()
ws = wb.active

# Les données pour le tableau
data = [
    ["Nom de l'organisme :", "EKOFORMA"],
    ["Identifiant", "99LH"],
    [],
    ["N° Action / Programme : 99LH2325001", "", "N° de session : 	24.015", "N° de l’unité: 1", "Volume horaire déclaré :", "25/04/2024 : 3 H MATIN"]
]

# Ajouter plusieurs tableaux
create_table(ws, start_row=1, start_col=1, title="Synthèse de suivi de classe virtuelle", data=data)


# Sauvegarder le fichier
file_path = "exemple_dynamique.xlsx"
wb.save(file_path)

file_path
