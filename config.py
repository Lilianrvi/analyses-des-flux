# config.py
# --------------------------------
# Mapping des produits (en minuscules lors de l'extraction)
PRODUCT_MAPPING = {
    "premium 13": "Premium France",
    "direct inter": "Direct Inter",
    "direct france": "Direct France",
    "system export": "Systeme et Prem Inter",
    "system france": "Systeme France",  # Sans accent
    "pallet france": "Pallet France",
    "system import": "Systeme et Prem Inter",
    "system home": "Système et Prem Inter",
    # Ajoutez d'autres variantes si nécessaire
}

# Pour le tonnage, spécifier le nom exact de la colonne à utiliser
PRODUCT_TONNAGE_FIELD = {
    "Premium France": "TONNAGE",
    "Systeme et Prem Inter": "TONNAGE",
    "Systeme France": "TONNAGE",
    "Pallet France": "TONNAGE",
    "Direct France": "Tonnage",
    "Direct Inter": "Tonnage",
    # Ajoutez d'autres produits si nécessaire
}

# Structure Excel pour les tableaux par produit et par année
EXCEL_STRUCTURE = {
    "RC": {
        "2025": {
            "Systeme France": "D11",
            "Pallet France": "D12",
            "Premium France": "D13",
            "Systeme et Prem Inter": "D14",
            "Direct France": "D15",
            "Direct Inter": "D16"
        },
        "2024": {
            "Systeme France": "E11",
            "Pallet France": "E12",
            "Premium France": "E13",
            "Systeme et Prem Inter": "E14",
            "Direct France": "E15",
            "Direct Inter": "E16"
        }
    },
    "Tonnage": {
        "2025": {
            "Systeme France": "L11",
            "Pallet France": "L12",
            "Premium France": "L13",
            "Systeme et Prem Inter": "L14",
            "Direct France": "L15",
            "Direct Inter": "L16"
        },
        "2024": {
            "Systeme France": "M11",
            "Pallet France": "M12",
            "Premium France": "M13",
            "Systeme et Prem Inter": "M14",
            "Direct France": "M15",
            "Direct Inter": "M16"
        }
    },
    "CA": {
        "2025": {
            "Systeme France": "D38",
            "Pallet France": "D39",
            "Premium France": "D40",
            "Systeme et Prem Inter": "D41",
            "Direct France": "D42",
            "Direct Inter": "D43"
        },
        "2024": {
            "Systeme France": "E38",
            "Pallet France": "E39",
            "Premium France": "E40",
            "Systeme et Prem Inter": "E41",
            "Direct France": "E42",
            "Direct Inter": "E43"
        }
    }
}

# Mapping des champs globaux à écrire dans Excel
GLOBAL_FIELDS = {
    "Nom du client": "G3",
    "Comptes clients": "G4",
    "Périodicité": "G5",
    "Dernier mois": "I6"  # Le dernier mois sera écrit ici sous format numérique
}

# Nom de la feuille Excel contenant toutes les tables
EXCEL_SHEET_NAME = "KPI activité client"
