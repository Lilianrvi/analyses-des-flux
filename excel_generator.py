# excel_generator.py
import xlwings as xw
import config

def update_template_with_xlwings(data_par_produit, client_info, output_path):
    """
    Ouvre le template Excel avec xlwings, met à jour uniquement les valeurs 
    (en conservant les formats et règles conditionnelles), et sauvegarde le résultat.
    Cette méthode nécessite Microsoft Excel et un environnement Windows.
    """
    # Ouvrir le template Excel
    wb = xw.Book(config.TEMPLATE_PATH)
    ws = wb.sheets[config.EXCEL_SHEET_NAME]
    
    # Mise à jour des champs globaux
    ws.range(config.GLOBAL_FIELDS["Nom du client"]).value = client_info.get("Nom du client", "")
    comptes = client_info.get("Comptes clients", [])
    ws.range(config.GLOBAL_FIELDS["Comptes clients"]).value = ", ".join(comptes) if comptes else ""
    ws.range(config.GLOBAL_FIELDS["Périodicité"]).value = client_info.get("Périodicité", "")
    
    # Pour I6, extraire le mois (nombre) de la dernière date (format mm/aaaa)
    period_str = client_info.get("Périodicité", "")
    parts = period_str.split()
    if len(parts) < 9:
        raise ValueError("Format de période invalide.")
    quoted_date = parts[8]  # Par exemple "12/2024"
    try:
        last_month = int(quoted_date.split("/")[0])
    except Exception:
        last_month = 0
    ws.range(config.GLOBAL_FIELDS["Dernier mois"]).value = last_month

    # Mise à jour des en-têtes pour Année N et Année N-1
    quoted_date_N_1 = parts[1]  # ex: "01/2023" pour N-1
    quoted_date_N = parts[8]    # ex: "12/2024" pour N
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

    # Remplissage des données des tableaux selon la structure EXCEL_STRUCTURE
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
                ws.range(cell).value = valeur

    # Sauvegarder le classeur et le fermer
    wb.save(output_path)
    wb.close()

def update_excel_with_xlwings(data_par_produit, client_info, output_path):
    update_template_with_xlwings(data_par_produit, client_info, output_path)
