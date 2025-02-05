# app.py
import streamlit as st
import io
from extraction import extract_data_from_pdf, validate_client_info
from excel_generator import load_template_workbook, fill_excel_workbook, fill_excel_workbook_addition
from openpyxl import load_workbook
import config
from openpyxl.styles import Alignment

# Fonction pour formater automatiquement la saisie d'une date au format mm/aaaa
def format_date_field(key):
    val = st.session_state.get(key, "")
    digits = "".join(ch for ch in val if ch.isdigit())
    # Format attendu : mm/aaaa (2 chiffres pour le mois, 4 pour l'année)
    formatted = digits[:2]
    if len(digits) > 2:
        formatted += "/" + digits[2:6]
    st.session_state[key] = formatted

# Configurer la page
st.set_page_config(page_title="Extraction et Addition Excel", layout="wide")

# Mode de sélection
mode = st.radio("Sélectionnez le mode", options=["Extraction depuis PDF", "Addition de fichiers Excel"])

if mode == "Extraction depuis PDF":
    st.title("Extraction Automatisée de Données PDF vers Excel")
    st.write("Cette application vous permet d'extraire des données spécifiques de fichiers PDF et de générer un fichier Excel basé sur un template.")
    
    st.subheader("Définissez la période d'analyse (format mm/aaaa)")
    col1, col2, col3, col4 = st.columns(4)
    date1 = col1.text_input("Du (mm/aaaa)", "", key="date1", on_change=lambda: format_date_field("date1"))
    date2 = col2.text_input("Au (mm/aaaa)", "", key="date2", on_change=lambda: format_date_field("date2"))
    date3 = col3.text_input("Et du (mm/aaaa)", "", key="date3", on_change=lambda: format_date_field("date3"))
    date4 = col4.text_input("Au (mm/aaaa)", "", key="date4", on_change=lambda: format_date_field("date4"))
    
    if not (date1 and date2 and date3 and date4):
        st.info("Veuillez remplir toutes les cases pour définir la période.")
        st.stop()
        
    period_string = f"Du {date1} au {date2} et du {date3} au {date4}"
    
    # File uploader pour PDF
    uploader_container = st.empty()
    uploaded_files = uploader_container.file_uploader(
        "Drag and drop files here (Limit 200MB per file • PDF)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"uploaded_files_{st.session_state.get('file_uploader_key', 0)}",
        help="Les nouveaux fichiers remplaceront les précédents."
    )
    
    # Bouton Clear Files pour PDF
    clear = st.button("Clear Files")
    if clear:
        st.session_state.file_uploader_key = st.session_state.get("file_uploader_key", 0) + 1
        uploader_container.empty()
        uploader_container.file_uploader(
            "Drag and drop files here (Limit 200MB per file • PDF)",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"uploaded_files_{st.session_state.file_uploader_key}",
            help="Les nouveaux fichiers remplaceront les précédents."
        )
        st.info("Fichiers effacés. Veuillez importer de nouveaux fichiers.")
        st.stop()
    
    if not uploaded_files:
        st.warning("Veuillez télécharger au moins un fichier PDF.")
        st.stop()
    if len(uploaded_files) > 12:
        st.error("Vous pouvez télécharger au maximum 12 fichiers PDF.")
        st.stop()
    
    with st.spinner("Traitement des fichiers PDF..."):
        try:
            extracted_data = []
            for pdf_file in uploaded_files:
                data, _ = extract_data_from_pdf(pdf_file)
                extracted_data.append(data)
            
            client_info = {
                "Nom du client": extracted_data[0].get("Nom du client", ""),
                "Comptes clients": extracted_data[0].get("Comptes clients", []),
                "Périodicité": period_string
            }
            
            st.subheader("Informations du Client Confirmées")
            st.write(f"**Nom du client :** {client_info['Nom du client']}")
            st.write(f"**Comptes clients :** {', '.join(client_info['Comptes clients']) if client_info['Comptes clients'] else 'Non trouvé'}")
            st.write(f"**Périodicité :** {client_info['Périodicité']}")
            
            data_par_produit = {}
            for data in extracted_data:
                produit = data.get("Produit concerné")
                annee = data.get("Année")
                if not produit or not annee:
                    st.warning(f"Produit ou année non reconnu dans le fichier {data.get('Nom du client', 'Unknown')}.")
                    continue
                if produit not in data_par_produit:
                    data_par_produit[produit] = {}
                if annee not in data_par_produit[produit]:
                    data_par_produit[produit][annee] = {"RC": 0, "Tonnage": 0, "CA": 0}
                data_par_produit[produit][annee]["RC"] += data.get("RC", 0)
                data_par_produit[produit][annee]["Tonnage"] += data.get("Tonnage", 0)
                data_par_produit[produit][annee]["CA"] += data.get("CA", 0.0)
            
            valid, error_msg = validate_client_info(extracted_data)
            if not valid:
                st.error(error_msg)
                st.stop()
            
            wb = load_template_workbook()
            wb = fill_excel_workbook(wb, data_par_produit, client_info)
            if wb is None:
                st.error("Erreur lors de la création du classeur Excel.")
                st.stop()
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.success("Le fichier Excel a été généré avec succès !")
            st.download_button(
                label="Télécharger le fichier Excel",
                data=excel_buffer,
                file_name=f"ANALYSES DES FLUX {client_info['Nom du client']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("### Détails des fichiers")
            for idx, data in enumerate(extracted_data):
                with st.expander(f"Fichier {idx + 1} : {uploaded_files[idx].name}", expanded=False):
                    st.write(f"**Nom du client :** {data.get('Nom du client', '')}")
                    st.write(f"**Comptes clients :** {', '.join(data.get('Comptes clients', [])) if data.get('Comptes clients') else 'Non trouvé'}")
                    st.write(f"**Produit concerné :** {data.get('Produit concerné', 'Non reconnu')}")
                    st.write(f"**Année :** {data.get('Année', 'Non reconnue')}")
                    st.write(f"**NB Dossier :** {data.get('RC', 0)}")
                    st.write(f"**TONNAGE :** {data.get('Tonnage', 0)}")
                    st.write(f"**CA HT Facturé :** {data.get('CA', 0.0)}")
        except Exception as e:
            st.error(f"Une erreur s'est produite lors du traitement : {e}")

else:  # Mode "Addition de fichiers Excel"
    st.title("Addition de Fichiers Excel")
    st.write("Cette option vous permet de combiner les données de plusieurs fichiers Excel (templates préremplis) en additionnant uniquement les cellules à l'intérieur des tableaux.")
    
    st.subheader("Importer vos fichiers Excel")
    excel_container = st.empty()
    excel_files = excel_container.file_uploader(
        "Téléchargez vos fichiers Excel (format .xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="excel_files"
    )
    
    # Bouton Clear Files pour Excel
    clear_excel = st.button("Clear Excel Files")
    if clear_excel:
        st.session_state.excel_uploader_key = st.session_state.get("excel_uploader_key", 0) + 1
        excel_container.empty()
        excel_container.file_uploader(
            "Téléchargez vos fichiers Excel (format .xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"excel_files_{st.session_state.excel_uploader_key}"
        )
        st.info("Fichiers Excel effacés. Veuillez importer de nouveaux fichiers.")
        st.stop()
    
    if not excel_files:
        st.warning("Veuillez télécharger au moins un fichier Excel.")
        st.stop()
    
    with st.spinner("Combinaison des fichiers Excel..."):
        try:
            # Initialiser la structure pour additionner uniquement les cellules des tableaux
            combined_data = {}
            for table_key, years in config.EXCEL_STRUCTURE.items():
                combined_data[table_key] = {}
                for year, products in years.items():
                    combined_data[table_key][year] = {}
                    for product, cell in products.items():
                        combined_data[table_key][year][product] = 0.0
            
            combined_global_names = []
            combined_global_accounts = []
            combined_period = None  # On utilisera la période du premier fichier
            
            # Parcourir chaque fichier Excel et additionner les valeurs dans les cellules indiquées
            for idx, file in enumerate(excel_files):
                wb_file = load_workbook(filename=io.BytesIO(file.read()), data_only=True)
                ws_file = wb_file[config.EXCEL_SHEET_NAME]
                client_name = ws_file["G3"].value
                client_accounts = ws_file["G4"].value
                period = ws_file["G5"].value
                if idx == 0:
                    combined_period = period  # On prend la période du premier fichier
                if client_name:
                    combined_global_names.append(str(client_name))
                if client_accounts:
                    combined_global_accounts.append(str(client_accounts))
                # Additionner uniquement les cellules correspondant à la structure (tableaux)
                for table_key, years in config.EXCEL_STRUCTURE.items():
                    for year, products in years.items():
                        for product, cell in products.items():
                            try:
                                val = ws_file[cell].value
                                if val is None:
                                    val = 0.0
                                else:
                                    val = float(val)
                            except (ValueError, TypeError):
                                val = 0.0
                            combined_data[table_key][year][product] += val
            
            new_wb = load_template_workbook()
            # Pour les champs globaux, on reprend ceux du premier fichier
            new_client_name = combined_global_names[0] if combined_global_names else ""
            new_client_accounts = combined_global_accounts[0] if combined_global_accounts else ""
            new_period = combined_period if combined_period else "Période inconnue"
            
            new_ws = new_wb[config.EXCEL_SHEET_NAME]
            new_ws["G3"].value = new_client_name
            new_ws["G4"].value = new_client_accounts
            new_ws["G5"].value = new_period
            
            # Ajuster la cellule G3 (nom) pour renvoyer le texte à la ligne et centré
            new_ws["G3"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            new_ws.row_dimensions[3].height = 150
            
            # Mise à jour des en-têtes à reprendre depuis la période du premier fichier
            period_parts = new_period.split()
            if len(period_parts) < 9:
                raise ValueError("Format de période invalide dans le fichier Excel.")
            quoted_date_N_1 = period_parts[1]  # ex: "01/2023" pour N-1
            quoted_date_N = period_parts[8]      # ex: "12/2024" pour N
            year_N_1 = quoted_date_N_1.split("/")[1]
            year_N = quoted_date_N.split("/")[1]
            if year_N_1 == year_N:
                raise ValueError("Les années de comparaison sont identiques dans le fichier Excel.")
            header_val_N_1 = int(year_N_1)
            header_val_N = int(year_N)
            cells_N = ["D9", "F9", "L9", "N9", "D36", "F36", "L36", "N36", "R9", "S37", "U37", "W37"]
            cells_N_1 = ["E9", "G9", "M9", "O9", "E36", "G36", "M36", "O36", "T37", "V37", "X37"]
            for cell in cells_N:
                new_ws[cell].value = header_val_N
            for cell in cells_N_1:
                new_ws[cell].value = header_val_N_1
            # Pour I6, extraire le mois de la dernière date (quoted_date_N)
            try:
                header_I6 = int(quoted_date_N.split("/")[0])
            except Exception:
                header_I6 = 0
            new_ws[config.GLOBAL_FIELDS["Dernier mois"]].value = header_I6
            
            # Remplir les cellules des tableaux avec les valeurs additionnées
            for table_key, years in config.EXCEL_STRUCTURE.items():
                for year, products in years.items():
                    for product, cell in products.items():
                        new_ws[cell].value = combined_data[table_key][year][product]
            
            output_filename = f"ANALYSES DES FLUX {new_client_name}.xlsx"
            output_buffer = io.BytesIO()
            new_wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("Les fichiers Excel ont été combinés avec succès !")
            st.download_button(
                label="Télécharger le fichier Excel combiné",
                data=output_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Une erreur s'est produite lors de la combinaison des fichiers Excel : {e}")
