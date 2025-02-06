# extraction.py
import pdfplumber
import re
import string
from config import PRODUCT_MAPPING, PRODUCT_TONNAGE_FIELD

def extract_data_from_pdf(pdf):
    data = {}
    with pdfplumber.open(pdf) as pdf_obj:
        first_page = pdf_obj.pages[0]
        text = first_page.extract_text()
        full_text = text  # Pour débogage

        print("----- Début du texte du PDF -----")
        print(text)
        print("----- Fin du texte du PDF -----")

        # Extraction du Nom du client
        client_match = re.search(r'Analyse des ventes par client\s+([^\s].*?)\s+\d{2}/\d{2}/\d{4}', text, re.IGNORECASE)
        if client_match:
            data['Nom du client'] = client_match.group(1).strip()
        else:
            client_match = re.search(r'Analyse des ventes par client\s+(.+?)\s+\d{2}/\d{2}/\d{4}', text, re.IGNORECASE)
            data['Nom du client'] = client_match.group(1).strip() if client_match else ''

        # Extraction des Comptes clients
        comptes_match = re.search(r'Compte(?:\(s\))?\s*:\s*\[([^\]]+)\]', text, re.IGNORECASE)
        if comptes_match:
            comptes_str = comptes_match.group(1)
            comptes = re.findall(r'\d+', comptes_str)
            data['Comptes clients'] = comptes
            print(f"Comptes clients trouvés : {comptes}")
        else:
            data['Comptes clients'] = []
            print("Comptes clients non trouvés")

        # Extraction du Produit concerné
        if comptes_match:
            start_pos = comptes_match.end()
            text_after = text[start_pos:]
            produit_match = re.search(r'\[([^\]]+)\]', text_after, re.IGNORECASE)
            if produit_match:
                produit_full = produit_match.group(1).strip()
                produit = ' '.join(produit_full.split()[:2]).strip().lower().strip(string.punctuation)
                print(f"Produit extrait : {produit}")
                data['Produit concerné'] = PRODUCT_MAPPING.get(produit, None)
            else:
                data['Produit concerné'] = None
                print("Produit non trouvé après comptes")
        else:
            produit_match = re.search(r'\[([^\]]+)\]', text, re.IGNORECASE)
            if produit_match:
                produit_full = produit_match.group(1).strip()
                produit = ' '.join(produit_full.split()[:2]).strip().lower().strip(string.punctuation)
                print(f"Produit extrait (alternative) : {produit}")
                data['Produit concerné'] = PRODUCT_MAPPING.get(produit, None)
            else:
                data['Produit concerné'] = None
                print("Produit non trouvé dans le texte")
        
        # Extraction des données du tableau (année, RC, Tonnage, CA)
        tables = pdf_obj.pages[0].extract_tables()
        analyse_table = None
        for table in tables:
            headers = table[0]
            if any('mois' in header.lower() for header in headers):
                analyse_table = table
                break
        
        if analyse_table:
            print("Table 'Analyse par mois de transport' trouvée")
            tonnage_field = PRODUCT_TONNAGE_FIELD.get(data.get('Produit concerné'), 'Tonnage')
            header_map = {}
            for idx, header in enumerate(analyse_table[0]):
                header_clean = header.strip().lower()
                if 'nb rc' in header_clean or 'nb dossier' in header_clean:
                    header_map['RC'] = idx
                if header_clean == tonnage_field.lower():
                    header_map['Tonnage'] = idx
                if 'ca ht facturé' in header_clean or 'ca ht facture' in header_clean:
                    header_map['CA'] = idx
            print(f"Mapping des en-têtes : {header_map}")
            
            if not all(k in header_map for k in ['RC', 'Tonnage', 'CA']):
                data['RC'] = data['Tonnage'] = data['CA'] = 0
            else:
                years = set()
                for row in analyse_table[1:]:
                    mois_val = row[0]
                    if mois_val and mois_val.strip() and not mois_val.strip().lower().startswith("total"):
                        match = re.search(r'(\d{4})', mois_val)
                        if match:
                            year = match.group(1)
                            if year in ["2023", "2024"]:
                                years.add(year)
                data['Année'] = years.pop() if len(years)==1 else None
                total_row = None
                for row in analyse_table:
                    if any(k.lower() in row[0].lower() for k in ['totaux', 'total']):
                        total_row = row
                        break
                if total_row:
                    try:
                        data['RC'] = int(total_row[header_map['RC']].replace(',', '').strip())
                        data['Tonnage'] = float(total_row[header_map['Tonnage']].replace(',', '.').strip())
                        data['CA'] = float(total_row[header_map['CA']].replace(',', '.').strip())
                    except Exception as e:
                        print(f"Erreur extraction valeurs : {e}")
                        data['RC'] = data['Tonnage'] = data['CA'] = 0
                else:
                    data['RC'] = data['Tonnage'] = data['CA'] = 0
        else:
            data['Année'] = None
            data['RC'] = 0
            data['Tonnage'] = 0.0
            data['CA'] = 0

        print(f"Fichier PDF: {pdf.name}")
        print(f"Nom du client: {data.get('Nom du client')}")
        return data, full_text

def validate_client_info(extracted_data):
    client_names = {data.get('Nom du client', '') for data in extracted_data}
    client_accounts = {tuple(sorted(data.get('Comptes clients', []))) for data in extracted_data}
    if len(client_names) > 1:
        return False, "Les fichiers PDF contiennent des noms de clients différents."
    if len(client_accounts) > 1:
        return False, "Les fichiers PDF contiennent des comptes clients différents."
    return True, ""
