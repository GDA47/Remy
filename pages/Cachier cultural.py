import pandas as pd
import streamlit as st
import io
import xlsxwriter
from datetime import datetime


def charger_fichier(uploaded_file):
    """Load and validate the input file"""
    try:
        df = pd.read_csv(uploaded_file,sep='\t',encoding='cp1252',na_values=['', 'NA', 'N/A', 'NaN', 'None', ' '],keep_default_na=False)
        if df.empty:
            st.error("Le fichier est vide ou ne contient pas de données valides")
            return None

        return df

    except Exception as e:
        st.error(f"❌ Erreur lors du chargement du fichier : {str(e)}")
        return None


def nettoyer_noms_colonnes(df):
    """Clean column names by fixing common errors"""
    df.columns = df.columns.str.replace("Prvisionnelle", "Prévisionnelle") \
        .str.replace("dbut", "début") \
        .str.replace("Unit", "Unité") \
        .str.replace("Unitéé", "Unité") \
        .str.replace("l'intrant", "à l'intrant") \
        .str.replace("à à", "à")
    return df


# def traiter_donnees(df):
#     """Process and filter the data"""
#     # Column definitions
#     col_date = "Interventions des parcelles culturales.Date début"
#     col_prev = "Interventions des parcelles culturales.Prévisionnelle"
#     col_dose = "Intrants des parcelles culturales.Dose"
#     col_unite = "Intrants des parcelles culturales.Unité"
#
#     # Filter for "Non" values in Prévisionnelle column
#     if col_prev in df.columns:
#         df = df[df[col_prev].astype(str).str.strip().str.lower() == "non"]
#         df = df[~df[col_prev].isna()]
#
#         if df.empty:
#             st.warning("Aucune donnée avec 'Prévisionnelle = Non' trouvée")
#             return df
#
#     # Date processing
#     if col_date in df.columns:
#         df[col_date] = pd.to_datetime(
#             df[col_date],
#             dayfirst=True,
#             errors='coerce',
#             format='mixed'
#         )
#
#         # Remove rows with invalid dates
#         initial_count = len(df)
#         df = df.dropna(subset=[col_date])
#         if len(df) < initial_count:
#             st.warning(f"{initial_count - len(df)} lignes supprimées (dates invalides)")
#
#         # Standardize year
#         if not df.empty:
#             df['Year'] = df[col_date].dt.year
#             max_year = df['Year'].max()
#             df['Year'] = max_year
#             df[col_date] = df[col_date].apply(
#                 lambda x: x.replace(year=max_year) if pd.notnull(x) else pd.NaT
#             )
#             df = df.sort_values(by=col_date, ascending=True)
#             df[col_date] = df[col_date].dt.strftime('%d/%m/%Y')
#
#     # Merge dose and unit columns
#     if col_dose in df.columns and col_unite in df.columns:
#         dose_str = df[col_dose].astype(str).str.strip().replace('nan', '')
#         unit_str = df[col_unite].astype(str).str.strip().replace('nan', '')
#         df[col_dose] = (dose_str + ' ' + unit_str).str.strip()
#         df.drop(columns=[col_unite], inplace=True, errors='ignore')
#
#     return df
import streamlit as st
import pandas as pd

def traiter_donnees(df):
    """Process and filter the data"""

    # Column definitions
    col_date = "Interventions des parcelles culturales.Date début"
    col_prev = "Interventions des parcelles culturales.Prévisionnelle"
    col_dose = "Intrants des parcelles culturales.Dose"
    col_unite = "Intrants des parcelles culturales.Unité"

    # Show original dataframe
    st.subheader("Tableau original")
    st.dataframe(df)

    # Filter for "Non" values in Prévisionnelle column
    if col_prev in df.columns:
        df = df[df[col_prev].astype(str).str.strip().str.lower() == "non"]
        df = df[~df[col_prev].isna()]

        if df.empty:
            st.warning("Aucune donnée avec 'Prévisionnelle = Non' trouvée")
            return df

    # Date processing
    if col_date in df.columns:
        df[col_date] = pd.to_datetime(
            df[col_date],
            dayfirst=True,
            errors='coerce',
            format='mixed'
        )

        # Remove rows with invalid dates
        initial_count = len(df)
        df = df.dropna(subset=[col_date])
        if len(df) < initial_count:
            st.warning(f"{initial_count - len(df)} lignes supprimées (dates invalides)")

        # Standardize year
        if not df.empty:
            df['Year'] = df[col_date].dt.year
            max_year = df['Year'].max()
            df['Year'] = max_year
            df[col_date] = df[col_date].apply(
                lambda x: x.replace(year=max_year) if pd.notnull(x) else pd.NaT
            )
            df = df.sort_values(by=col_date, ascending=True)
            df[col_date] = df[col_date].dt.strftime('%d/%m/%Y')

    # Merge dose and unit columns
    if col_dose in df.columns and col_unite in df.columns:
        dose_str = df[col_dose].astype(str).str.strip().replace('nan', '')
        unit_str = df[col_unite].astype(str).str.strip().replace('nan', '')
        df[col_dose] = (dose_str + ' ' + unit_str).str.strip()
        df.drop(columns=[col_unite], inplace=True, errors='ignore')

    # # Show filtered and processed dataframe
    # st.subheader("Tableau filtré et traité")
    # st.dataframe(df)

    return df


def get_table_exploitations_parcelles(df):
    """Generate farm information table"""
    rename_dict = {
        "Exploitations.Raison sociale": "Raison sociale",
        "Exploitations.Adresse_exploitant": "Adresse",
        "Exploitations.Téléphone": "Téléphone",
        "Exploitations.Code SIRET": "Numéro SIRET",
        "Parcelles culturales.Culture": "Espèce"
    }

    cols = [col for col in rename_dict.keys() if col in df.columns]

    if not cols:
        st.error("Aucune colonne valide trouvée pour le tableau des exploitations")
        return None

    result = []
    for col in cols:
        valeurs = df[col].dropna().astype(str).str.strip()
        valeurs = valeurs[valeurs != ''].unique()
        nom_affiche = rename_dict[col]
        for val in valeurs:
            result.append([nom_affiche, val])

    if not result:
        return None

    table = pd.DataFrame(result, columns=["Élément", "Valeur"])

    # Insert empty rows for organization
    if "Téléphone" in table["Élément"].values:
        idx_tel = table[table["Élément"] == "Téléphone"].index.max()
        lignes_insertion = pd.DataFrame([["Organisation de producteur", ""], ["Service technique", ""]],
                                        columns=["Élément", "Valeur"])
        part1 = table.iloc[:idx_tel + 1]
        part2 = table.iloc[idx_tel + 1:]
        table = pd.concat([part1, lignes_insertion, part2], ignore_index=True)

    # Add year
    max_year = df['Year'].max() if 'Year' in df.columns else "N/A"
    table = pd.concat([table, pd.DataFrame([["Année", max_year]], columns=["Élément", "Valeur"])], ignore_index=True)

    return table


def get_table_codification_parcelles(df):
    """Generate parcel coding table"""
    parcelle_cols = [col for col in df.columns if "Parcelles" in col and "Nom" in col]

    if not parcelle_cols:
        st.warning("Colonne 'Nom de parcelle' introuvable")
        return None

    parcelle_col = parcelle_cols[0]
    parcelle_names = df[parcelle_col].dropna().astype(str).str.strip()
    parcelle_names = parcelle_names[parcelle_names != ''].unique()

    if len(parcelle_names) == 0:
        st.warning("Aucun nom de parcelle valide trouvé")
        return None

    df_codif = pd.DataFrame([parcelle_names, range(1, len(parcelle_names) + 1)])
    df_codif.index = ["Nom de la parcelle", "Code parcelle"]

    return df_codif


# def get_table_operations_agricoles_codifie(df):
#     """Generate agricultural operations table"""
#     # Find required columns
#     date_cols = [col for col in df.columns if "Date" in col and "début" in col]
#     type_col = "Types d'interventions.Nom"
#     parcelle_col = "Parcelles culturales.Nom"
#
#     if not date_cols or type_col not in df.columns or parcelle_col not in df.columns:
#         st.error("Colonnes requises manquantes")
#         return None
#
#     col_date = date_cols[0]
#
#     try:
#         df_op = df[[col_date, type_col, parcelle_col]].copy()
#         df_op[col_date] = pd.to_datetime(df_op[col_date], dayfirst=True, errors='coerce')
#         df_op = df_op.dropna(subset=[col_date])
#
#         if df_op.empty:
#             return None
#
#         # Create parcel coding
#         parcelle_names = df_op[parcelle_col].dropna().astype(str).str.strip()
#         parcelle_names = parcelle_names[parcelle_names != ''].unique()
#         codif_dict = {name: idx + 1 for idx, name in enumerate(parcelle_names)}
#
#         # Apply coding
#         df_op["Code parcelle"] = df_op[parcelle_col].map(codif_dict)
#         for code in codif_dict.values():
#             df_op[str(code)] = df_op["Code parcelle"].apply(lambda x: "x" if x == code else "")
#
#         # Group operations
#         grouped = df_op.groupby([col_date, type_col], dropna=False)
#         lignes_fusionnees = []
#
#         for (date, type_interv), group in grouped:
#             ligne = {"Date": date.strftime("%d/%m/%Y"), "Type d'intervention": type_interv}
#             for code in codif_dict.values():
#                 ligne[str(code)] = 'x' if (group[str(code)] == 'x').any() else ''
#             lignes_fusionnees.append(ligne)
#
#         df_result = pd.DataFrame(lignes_fusionnees)
#         columns_order = ["Date", "Type d'intervention"] + sorted(str(c) for c in codif_dict.values())
#         return df_result[columns_order]
#
#     except Exception as e:
#         st.error(f"Erreur: {str(e)}")
#         return None

def get_table_operations_agricoles_codifie(df):
    """Generate agricultural operations table"""

    # Define allowed operations
    operations = [
        "Arrachage culture pérenne",
        "Broyage des bois de taille",
        "Brûlage des bois de taille",
        "Ebourgeonnage",
        "Ébourgeonnage fructifère",
        "Écimage",
        "Éclaircissage manuel/physiologique",
        "Élagage",
        "Élagage double têtes",
        "Entreplantation-complantation-rebrochage",
        "Élagage",
        "Élagage double têtes",
        "Liage",
        "Marcotage",
        "Palissage",
        "Pré-taille",
        "Surgreffage",
        "Taille",
        "Taille au sabre",
        "Taille en vert",
        "Tirage des bois"
    ]

    # Find required columns
    date_cols = [col for col in df.columns if "Date" in col and "début" in col]
    type_col = "Types d'interventions.Nom"
    parcelle_col = "Parcelles culturales.Nom"

    if not date_cols or type_col not in df.columns or parcelle_col not in df.columns:
        st.error("Colonnes requises manquantes")
        return None

    col_date = date_cols[0]

    try:
        df_op = df[[col_date, type_col, parcelle_col]].copy()
        df_op[col_date] = pd.to_datetime(df_op[col_date], dayfirst=True, errors='coerce')
        df_op = df_op.dropna(subset=[col_date])

        # 💡 Apply the filter on intervention type
        df_op = df_op[df_op[type_col].isin(operations)]

        if df_op.empty:
            return None

        # Create parcel coding
        parcelle_names = df_op[parcelle_col].dropna().astype(str).str.strip()
        parcelle_names = parcelle_names[parcelle_names != ''].unique()
        codif_dict = {name: idx + 1 for idx, name in enumerate(parcelle_names)}

        # Apply coding
        df_op["Code parcelle"] = df_op[parcelle_col].map(codif_dict)
        for code in codif_dict.values():
            df_op[str(code)] = df_op["Code parcelle"].apply(lambda x: "x" if x == code else "")

        # Group operations
        grouped = df_op.groupby([col_date, type_col], dropna=False)
        lignes_fusionnees = []

        for (date, type_interv), group in grouped:
            ligne = {"Date": date.strftime("%d/%m/%Y"), "Type d'intervention": type_interv}
            for code in codif_dict.values():
                ligne[str(code)] = 'x' if (group[str(code)] == 'x').any() else ''
            lignes_fusionnees.append(ligne)

        df_result = pd.DataFrame(lignes_fusionnees)
        columns_order = ["Date", "Type d'intervention"] + sorted(str(c) for c in codif_dict.values())
        return df_result[columns_order]

    except Exception as e:
        st.error(f"Erreur: {str(e)}")
        return None


def get_table_irrigation(df):
    """Generate irrigation table"""
    type_col = "Types d'interventions.Nom"
    date_col = "Interventions des parcelles culturales.Date début"
    dose_col = "Intrants des parcelles culturales.Dose"
    parcelle_col = "Parcelles culturales.Nom"

    required_cols = [type_col, date_col, dose_col, parcelle_col]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"Colonnes manquantes: {', '.join(missing_cols)}")
        return None

    try:
        df_irrig = df[df[type_col].str.lower().str.strip() == "irrigation"].copy()

        if df_irrig.empty:
            return None

        df_irrig[date_col] = pd.to_datetime(df_irrig[date_col], dayfirst=True, errors='coerce')
        df_irrig = df_irrig.dropna(subset=[date_col])

        df_result = df_irrig[[date_col, dose_col, parcelle_col]].copy()
        df_result.rename(columns={
            date_col: "Date",
            dose_col: "Dose",
            parcelle_col: "Parcelle"
        }, inplace=True)

        df_result["Pluie (mm)"] = ""
        df_result["X"] = "x"

        df_pivot = df_result.pivot_table(
            index=["Date", "Dose", "Pluie (mm)"],
            columns="Parcelle",
            values="X",
            aggfunc="first",
            fill_value=""
        ).reset_index()

        df_pivot["Date"] = df_pivot["Date"].dt.strftime('%d/%m/%Y')
        return df_pivot

    except Exception as e:
        st.error(f"Erreur irrigation: {str(e)}")
        return None


def get_table_fertilisation(df):
    """Generate fertilization table"""
    fertilisation_types = [
        "Amendements calco-magnésiens", "Biostimulant", "Boues de station d'épuration/compost urbain",
        "Effluents d'élevage", "Fertilisation minérale", "Fertilisation minérale Bulk",
        "Fertirrigation", "Obligo-éléments", "Organo-minéral", "Fertilisation organique",
        "Sous-produits/déchets alimentaires", "Sous-produits/déchets non alimentaires", "Supports de culture"
    ]

    required_cols = {
        'type': "Types d'interventions.Nom",
        'date': "Interventions des parcelles culturales.Date début",
        'dose': "Intrants des parcelles culturales.Dose",
        'parcelle': "Parcelles culturales.Nom"
    }

    missing_cols = [col for col in required_cols.values() if col not in df.columns]
    if missing_cols:
        st.error(f"Colonnes manquantes: {', '.join(missing_cols)}")
        return None

    try:
        df_fert = df[df[required_cols['type']].isin(fertilisation_types)].copy()

        if df_fert.empty:
            return None

        df_fert[required_cols['date']] = pd.to_datetime(
            df_fert[required_cols['date']],
            dayfirst=True,
            errors='coerce'
        )
        df_fert = df_fert.dropna(subset=[required_cols['date']])

        # Prepare result
        column_mapping = {
            required_cols['date']: "📅 Date",
            "Traitements.Nom": "🧪 Produit",
            required_cols['dose']: "💧 Dose",
            "Engrais.N": "🧬 N",
            "Engrais.P2O5": "🧬 P₂O₅",
            "Engrais.K2O": "🧬 K₂O",
            "Engrais.CaO": "🧬 CaO",
            "Engrais.MgO": "🧬 MgO",
            required_cols['parcelle']: "🌿 Parcelle"
        }

        available_cols = [col for col in column_mapping.keys() if col in df_fert.columns]
        df_result = df_fert[available_cols].rename(columns=column_mapping)

        # Add parcel markers
        parcelles = df_fert[required_cols['parcelle']].dropna().unique()
        for parcelle in parcelles:
            df_result[parcelle] = df_fert[required_cols['parcelle']].apply(
                lambda x: 'x' if x == parcelle else '')

        # Group similar operations
        group_cols = [col for col in ["📅 Date", "💧 Dose", "🧪 Produit",
                                      "🧬 N", "🧬 P₂O₅", "🧬 K₂O", "🧬 CaO", "🧬 MgO"]
                      if col in df_result.columns]

        grouped = df_result.groupby(group_cols, dropna=False)
        lignes_fusionnees = []

        for _, group in grouped:
            ligne = group.iloc[0].copy()
            for parcelle in parcelles:
                ligne[parcelle] = 'x' if (group[parcelle] == 'x').any() else ''
            lignes_fusionnees.append(ligne)

        df_final = pd.DataFrame(lignes_fusionnees)
        if "📅 Date" in df_final.columns:
            df_final["📅 Date"] = df_final["📅 Date"].dt.strftime("%d/%m/%Y")
        df_final.drop(columns=["🌿 Parcelle"], inplace=True, errors='ignore')

        return df_final

    except Exception as e:
        st.error(f"Erreur fertilisation: {str(e)}")
        return None


def get_table_traitement(df):
    """Generate treatment table"""
    excluded_types = [
        "Amendements calco-magnésiens", "Biostimulant", "Boues de station d'épuration/compost urbain",
        "Effluents d'élevage", "Fertilisation minérale", "Fertilisation minérale Bulk",
        "Fertirrigation", "Obligo-éléments", "Organo-minéral", "Taille", "Fertilisation organique",
        "Irrigation", "Sous-produits/déchets alimentaires", "Sous-produits/déchets non alimentaires",
        "Supports de culture",
        "Arrachage culture pérenne", "Broyage des bois de taille", "Brûlage des bois de taille",
        "Ebourgeonnage", "Ébourgeonnage fructifère", "Écimage", "Eclaircissage manuel/physiologique",
        "Elagage", "Elagage double tétes", "Entreplantation-complantation-rebrochage",
        "Elagage double têtes", "Liage", "Marcotage", "Palissage", "Pré-taille", "Surgreffage",
        "Taille au sabre", "Taille en vert", "Tirage des bois"
    ]

    required_cols = {
        'type': "Types d'interventions.Nom",
        'date': "Interventions des parcelles culturales.Date début",
        'dose': "Intrants des parcelles culturales.Dose",
        'produit': "Traitements.Nom",
        'cible': "Cibles à l'intrant.Nom de la cible",
        'parcelle': "Parcelles culturales.Nom"
    }

    missing_cols = [col for col in required_cols.values() if col not in df.columns]
    if missing_cols:
        st.error(f"Colonnes manquantes: {', '.join(missing_cols)}")
        return None

    try:
        df_trait = df[~df[required_cols['type']].isin(excluded_types)].copy()

        if df_trait.empty:
            return None

        df_trait[required_cols['date']] = pd.to_datetime(
            df_trait[required_cols['date']],
            dayfirst=True,
            errors='coerce'
        )
        df_trait = df_trait.dropna(subset=[required_cols['date']])

        # Add parcel markers
        parcelles = df_trait[required_cols['parcelle']].dropna().unique()
        for parcelle in parcelles:
            df_trait[parcelle] = df_trait[required_cols['parcelle']].apply(
                lambda x: 'x' if x == parcelle else '')

        # Group treatments
        group_cols = [
            required_cols['date'],
            required_cols['produit'],
            required_cols['type'],
            required_cols['dose']
        ]

        grouped = df_trait.groupby(group_cols, dropna=False)
        lignes_fusionnees = []

        for _, group in grouped:
            ligne = group.iloc[0].copy()
            if required_cols['cible'] in group.columns:
                ligne["Cible"] = group[required_cols['cible']].dropna().iloc[0] \
                    if not group[required_cols['cible']].dropna().empty else ''
            else:
                ligne["Cible"] = ''

            for parcelle in parcelles:
                ligne[parcelle] = 'x' if (group[parcelle] == 'x').any() else ''
            lignes_fusionnees.append(ligne)

        df_result = pd.DataFrame(lignes_fusionnees)
        df_result["Date"] = df_result[required_cols['date']].dt.strftime("%d/%m/%Y")

        # Add empty columns
        df_result.insert(3, "DAR", "")
        df_result.insert(6, "Commentaire", "")

        # Organize columns
        final_order = ["Date", required_cols['produit'], required_cols['type'],
                       "DAR", required_cols['dose'], "Cible", "Commentaire"] + list(parcelles)
        df_result = df_result[final_order]

        # Rename columns
        df_result.rename(columns={
            required_cols['produit']: "Produit commercial",
            required_cols['type']: "Matiere active",
            required_cols['dose']: "Dose appliquée par ha"
        }, inplace=True)

        return df_result

    except Exception as e:
        st.error(f"Erreur traitement: {str(e)}")
        return None


def get_table_inventaire_parcelles(df):
    """Generate parcel inventory table"""
    column_mapping = {
        "Parcelles culturales.Nom": "Nom de la parcelle",
        "Variétés de parcelle.Nom": "Variété",
        "Parcelles culturales.Lieu-dit": "Lieu-dit",
        "Parcelles culturales.Surface": "Surface (ha)",
        "Parcelles culturales.PFI Verger éco responsable": "PFI Verger éco responsable",
        "Parcelles culturales.ZRP Zéro Résidu Pesticide": "ZRP Zéro Résidu Pesticide",
        "Parcelles culturales.Global Gap": "Global GAP",
        "Parcelles culturales.HVE 3": "HVE 3"
    }

    # Get available columns
    available_cols = [col for col in column_mapping.keys() if col in df.columns]

    if not available_cols:
        st.error("Aucune colonne valide trouvée pour l'inventaire des parcelles")
        return None

    # Create result dataframe
    result = df[available_cols].copy()
    result = result.rename(columns=column_mapping)
    result = result.drop_duplicates()

    # Add empty columns
    empty_cols = ["Autres", "Suivi 1", "Suivi 2", "Suivi 3", "Conformité C", "Conformité NC", "Motivation"]
    for col in empty_cols:
        result[col] = ""

    return result


def export_all_tables_to_excel(table_dict, raison_sociale):
    """Export all tables to an Excel file"""
    # Clean filename
    safe_name = "".join(c for c in raison_sociale if c.isalnum() or c in (' ', '_')).strip()
    nom_fichier = f"Cahier_Cultural_{safe_name}_{datetime.now().strftime('%Y')}.xlsx"

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in table_dict.items():
            if df is not None and not df.empty:
                sheet_name = sheet_name[:31]  # Excel sheet name limit
                df.to_excel(writer, index=False, sheet_name=sheet_name)

    st.download_button(
        label="📥 Télécharger toutes les tables (Excel)",
        data=output.getvalue(),
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def main():
    st.title("Cahier culturel")

    uploaded_file = st.file_uploader("Téléchargez un fichier .txt", type=["txt"])
    if uploaded_file is not None:
        df = charger_fichier(uploaded_file)

        if df is not None:
            if df.empty:
                st.error("Le fichier chargé est vide ou ne contient pas de données valides")
                return

            df = nettoyer_noms_colonnes(df)
            df_original = df.copy()
            df = traiter_donnees(df)

            if df.empty:
                st.error("Aucune donnée ne correspond au critère 'Prévisionnelle = Non'")
                return

            st.subheader("Tableau des Données Filtrées")
            st.dataframe(df)

            # Generate all tables
            tables = {
                "Exploitation": get_table_exploitations_parcelles(df),
                "Codification Parcelles": get_table_codification_parcelles(df),
                "Inventaire Parcelles": get_table_inventaire_parcelles(df),
                "Operation agricole": get_table_operations_agricoles_codifie(df),
                "Traitement": get_table_traitement(df),
                "Fertilisation": get_table_fertilisation(df),
                "Irrigation": get_table_irrigation(df),
            }

            # Filter out None or empty tables
            tables = {k: v for k, v in tables.items() if v is not None and not v.empty}

            if not tables:
                st.error("Aucun tableau n'a pu être généré à partir des données")
                return

            # Get company name
            raison_sociale = "EARL_de_Fleury"  # Default value
            if "Exploitation" in tables:
                try:
                    rs_row = tables["Exploitation"][tables["Exploitation"]["Élément"] == "Raison sociale"]
                    if not rs_row.empty:
                        raison_sociale = rs_row.iloc[0]["Valeur"]  # Corrected from "VALUE"
                        # Clean the name for filename
                        raison_sociale = raison_sociale.replace(" ", "_").replace("/", "_").strip()
                except Exception as e:
                    st.warning(f"Impossible de récupérer la raison sociale : {str(e)}")

            # Display tables
            for name, table in tables.items():
                st.subheader(name)
                st.dataframe(table)

            # Export button
            export_all_tables_to_excel(tables, raison_sociale)


if __name__ == "__main__":
    main()
