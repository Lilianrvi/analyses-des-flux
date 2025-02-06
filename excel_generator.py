# excel_generator.py
import os
from openpyxl import load_workbook
import config

def load_template_workbook():
    template_path = "template.xlsx"
    if not os.path.exists(template_path):
        raise FileNotFoundError("Le fichier 'template.xlsx' est introuvable.")
    # Chargez le workbook sans options spéciales (cela préserve les formules et formats)
    wb = load_workbook(filename=template_path)
    if config.EXCEL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"La feuille '{config.EXCEL_SHEET_NAME}' est manquante.")
    return wb

def set_cell_value(ws, cell_coord, value):
    """
    Met à jour uniquement la valeur de la cellule existante, sans modifier aucun style ni format.
    Si la cellule fait partie d'une plage fusionnée, la cellule en haut à gauche de la plage est utilisée.
    """
    # Vérifier si la cellule est dans une plage fusionnée
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            cell.value = value
            return
    # Sinon, mettre à jour directement la cellule
    ws[cell_coord].value = value

def fill_excel_workbook(wb, data_par_produit, client_info):
    ws = wb[config.EXCEL_SHEET_NAME]
    # Mettre à jour les champs globaux sans modifier le style
    set_cell_value(ws, config.GLOBAL_FIELDS["Nom du client"], client_info.get("Nom du client", ""))
    comptes = client_info.get("Comptes clients", [])
    set_cell_value(ws, config.GLOBAL_FIELDS["Comptes clients"], ", ".join(comptes) if comptes else "")
    set_cell_value(ws, config.GLOBAL_FIELDS["Périodicité"], client_info.get("Périodicité", ""))
    
    # Pour I6, extraire le mois (nombre) de la dernière date de la période (format mm/aaaa)
    period_str = client_info.get("Périodicité", "")
    parts = period_str.split()
    if len(parts) < 9:
        raise ValueError("Format de période invalide.")
    quoted_date = parts[8]  # par exemple "12/2024"
    try:
        last_month = int(quoted_date.split("/")[0])
    except Exception:
        last_month = 0
    set_cell_value(ws, config.GLOBAL_FIELDS["Dernier mois"], last_month)
    
    # Mise à jour des en-têtes (Année N et Année N-1)
    quoted_date_N_1 = parts[1]  # ex: "01/2023" pour N-1
    quoted_date_N   = parts[8]  # ex: "12/2024" pour N
    year_N_1 = quoted_date_N_1.split("/")[1]
    year_N = quoted_date_N.split("/")[1]
    if year_N_1 == year_N:
        raise ValueError("Les années de comparaison sont identiques dans la période.")
    header_val_N_1 = int(year_N_1)
    header_val_N = int(year_N)
    cells_N = ["D9", "F9", "L9", "N9", "D36", "F36", "L36", "N36", "R9", "S37", "U37", "W37"]
    cells_N_1 = ["E9", "G9", "M9", "O9", "E36", "G36", "M36", "O36", "T37", "V37", "X37"]
    for cell in cells_N:
        set_cell_value(ws, cell, header_val_N)
    for cell in cells_N_1:
        set_cell_value(ws, cell, header_val_N_1)
    
    # Remplissage des données variables dans les cellules des tableaux définies dans EXCEL_STRUCTURE
    for tableau, annees in config.EXCEL_STRUCTURE.items():
        for annee, produits in annees.items():
            for produit, cell in produits.items():
                if produit in data_par_produit and annee in data_par_produit[produit]:
                    if tableau == "RC":
                        valeur = data_par_produit[produit][annee].get("RC", 0)
                    elif tableau == "Tonnage":
                        valeur = data_par_produit[produit][annee].get("Tonnage", 0)
                    elif tableau == "CA":
                        valeur = data_par_produit[produit][annee].get("CA", 0)
                    else:
                        valeur = 0
                else:
                    valeur = 0
                set_cell_value(ws, cell, valeur)
    return wb
