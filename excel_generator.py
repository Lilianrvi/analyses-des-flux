# excel_generator.py
import os
import sys
from openpyxl import load_workbook
import config

def load_template_workbook():
    template_path = config.TEMPLATE_PATH if hasattr(config, "TEMPLATE_PATH") else "template.xlsx"
    if not os.path.exists(template_path):
        raise FileNotFoundError("Le fichier 'template.xlsx' est introuvable.")
    wb = load_workbook(filename=template_path)
    if config.EXCEL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"La feuille '{config.EXCEL_SHEET_NAME}' est manquante.")
    return wb

def set_cell_value(ws, cell_coord, value):
    """
    Met à jour uniquement la valeur de la cellule existante (ou celle en haut à gauche
    d'une plage fusionnée) sans modifier les styles ni les règles conditionnelles.
    """
    try:
        new_value = float(value)
    except (ValueError, TypeError):
        new_value = value

    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            cell.value = new_value
            return
    ws[cell_coord].value = new_value

def fill_excel_workbook(wb, data_par_produit, client_info):
    """
    Utilisé pour l'extraction depuis PDF (avec openpyxl).
    Met à jour le template en modifiant uniquement les valeurs.
    """
    ws = wb[config.EXCEL_SHEET_NAME]
    # Champs globaux
    set_cell_value(ws, config.GLOBAL_FIELDS["Nom du client"], client_info.get("Nom du client", ""))
    comptes = client_info.get("Comptes clients", [])
    set_cell_value(ws, config.GLOBAL_FIELDS["Comptes clients"], ", ".join(comptes) if comptes else "")
    set_cell_value(ws, config.GLOBAL_FIELDS["Périodicité"], client_info.get("Périodicité", ""))
    
    # Pour I6 : extraire le mois de la dernière date (format mm/aaaa)
    period_str = client_info.get("Périodicité", "")
    parts = period_str.split()
    if len(parts) < 9:
        raise ValueError("Format de période invalide.")
    quoted_date = parts[8]  # ex: "12/2024"
    try:
        last_month = int(quoted_date.split("/")[0])
    except Exception:
        last_month = 0
    set_cell_value(ws, config.GLOBAL_FIELDS["Dernier mois"], last_month)
    
    # En-têtes : Année N et Année N-1
    quoted_date_N_1 = parts[1]  # ex: "01/2023"
    quoted_date_N = parts[8]    # ex: "12/2024"
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
    
    # Remplissage des données des tableaux
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
    
    # Forcer Excel à recalculer à l'ouverture (si possible)
    try:
        wb.calculation_properties.fullCalcOnLoad = True
    except AttributeError:
        pass
    
    return wb

# Définir update_excel_with_xlwings uniquement si l'environnement est Windows
if sys.platform.startswith('win'):
    import xlwings as xw
    def update_excel_with_xlwings(combined_data, client_info, output_path):
        """
        Utilise xlwings pour mettre à jour le template Excel afin de préserver exactement les règles conditionnelles.
        Nécessite Windows avec Excel installé.
        """
        wb = xw.Book(config.TEMPLATE_PATH if hasattr(config, "TEMPLATE_PATH") else "template.xlsx")
        ws = wb.sheets[config.EXCEL_SHEET_NAME]
        # Mise à jour des champs globaux
        ws.range(config.GLOBAL_FIELDS["Nom du client"]).value = client_info.get("Nom du client", "")
        comptes = client_info.get("Comptes clients", [])
        ws.range(config.GLOBAL_FIELDS["Comptes clients"]).value = ", ".join(comptes) if comptes else ""
        ws.range(config.GLOBAL_FIELDS["Périodicité"]).value = client_info.get("Périodicité", "")
        
        # Pour I6, extraire le mois
        period_str = client_info.get("Périodicité", "")
        parts = period_str.split()
        if len(parts) < 9:
            wb.close()
            raise ValueError("Format de période invalide.")
        quoted_date = parts[8]
        try:
            last_month = int(quoted_date.split("/")[0])
        except Exception:
            last_month = 0
        ws.range(config.GLOBAL_FIELDS["Dernier mois"]).value = last_month
        
        # En-têtes
        quoted_date_N_1 = parts[1]
        quoted_date_N = parts[8]
        year_N_1 = quoted_date_N_1.split("/")[1]
        year_N = quoted_date_N.split("/")[1]
        if year_N_1 == year_N:
            wb.close()
            raise ValueError("Les années de comparaison sont identiques dans la période.")
        header_val_N_1 = int(year_N_1)
        header_val_N = int(year_N)
        cells_N = ["D9", "F9", "L9", "N9", "D36", "F36", "L36", "N36", "R9", "S37", "U37", "W37"]
        cells_N_1 = ["E9", "G9", "M9", "O9", "E36", "G36", "M36", "O36", "T37", "V37", "X37"]
        for cell in cells_N:
            ws.range(cell).value = header_val_N
        for cell in cells_N_1:
            ws.range(cell).value = header_val_N_1
        
        # Remplissage des données
        for tableau, annees in config.EXCEL_STRUCTURE.items():
            for annee, produits in annees.items():
                for produit, cell in produits.items():
                    if produit in combined_data and annee in combined_data[produit]:
                        if tableau == "RC":
                            valeur = combined_data[produit][annee].get("RC", 0)
                        elif tableau == "Tonnage":
                            valeur = combined_data[produit][annee].get("Tonnage", 0)
                        elif tableau == "CA":
                            valeur = combined_data[produit][annee].get("CA", 0)
                        else:
                            valeur = 0
                    else:
                        valeur = 0
                    ws.range(cell).value = valeur
        
        wb.save(output_path)
        wb.close()
else:
    # Si l'environnement n'est pas Windows, on ne définit pas update_excel_with_xlwings.
    def update_excel_with_xlwings(*args, **kwargs):
        raise EnvironmentError("update_excel_with_xlwings est uniquement disponible sur Windows avec Excel.")

