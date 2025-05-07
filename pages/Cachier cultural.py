import pandas as pd
import streamlit as st
import os
import xlsxwriter
import io
from datetime import datetime


# Fonction pour charger et afficher le fichier téléchargé
def charger_fichier(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file, sep='\t', encoding='cp1252')
        return df
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement du fichier : {e}")
        return None


# Fonction pour nettoyer les noms de colonnes
def nettoyer_noms_colonnes(df):
    # Nettoyage des noms de colonnes pour corriger les erreurs courantes
    df.columns = df.columns.str.replace("Prvisionnelle", "Prévisionnelle") \
        .str.replace("dbut", "début") \
        .str.replace("Unit", "Unité") \
        .str.replace("Unitéé", "Unité") \
        .str.replace("l'intrant", "à l'intrant") \
        .str.replace("à à", "à")
    return df


def traiter_donnees(df):
    # Colonnes importantes pour le traitement
    col_date = "Interventions des parcelles culturales.Date début"
    col_prev = "Interventions des parcelles culturales.Prévisionnelle"
    col_dose = "Intrants des parcelles culturales.Dose"
    col_unite = "Intrants des parcelles culturales.Unité"

    # Filtrer les données non prévisionnelles
    if col_prev in df.columns:
        df = df[df[col_prev].str.strip().str.lower() != "oui"]

    # Traitement des dates
    if col_date in df.columns:
        df[col_date] = pd.to_datetime(df[col_date], dayfirst=True, errors='coerce')

        # Standardiser l'année sur la plus récente trouvée
        df['Year'] = df[col_date].dt.year
        max_year = df['Year'].max()
        df['Year'] = max_year
        df[col_date] = df[col_date].apply(lambda x: x.replace(year=max_year) if pd.notnull(x) else x)

        # Trier par date et formater
        df = df.sort_values(by=col_date, ascending=True)
        df[col_date] = df[col_date].dt.strftime('%d/%m/%Y')

    # Fusionner les colonnes Dose et Unité
    if col_dose in df.columns and col_unite in df.columns:
        df[col_dose] = df[col_dose].fillna('').astype(str).str.strip() + ' ' + df[col_unite].fillna('').astype(
            str).str.strip()
        df.drop(columns=[col_unite], inplace=True)

    return df


def get_table_exploitations_parcelles(df):
    # Dictionnaire pour renommer les colonnes
    rename_dict = {
        "Exploitations.Raison sociale": "Raison sociale",
        "Exploitations.Adresse_exploitant": "Adresse",
        "Exploitations.Téléphone": "Téléphone",
        "Exploitations.Code SIRET": "Numéro SIRET",
        "Parcelles culturales.Culture": "Espèce"
    }

    cols = list(rename_dict.keys())

    # Vérification des colonnes requises
    missing_cols = [col for col in cols if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Certaines colonnes manquent : {', '.join(missing_cols)}")
        return None

    # Construction du tableau
    result = []
    for col in cols:
        valeurs = df[col].dropna().unique()
        nom_affiche = rename_dict[col]
        for val in valeurs:
            result.append([nom_affiche, val])

    table = pd.DataFrame(result, columns=["Élément", "Valeur"])

    # Insertion des lignes vides pour Organisation de producteur et Service technique
    idx_tel = table[table["Élément"] == "Téléphone"].index.max()
    lignes_insertion = pd.DataFrame([["Organisation de producteur", ""], ["Service technique", ""]],
                                    columns=["Élément", "Valeur"])

    part1 = table.iloc[:idx_tel + 1]
    part2 = table.iloc[idx_tel + 1:]
    table = pd.concat([part1, lignes_insertion, part2], ignore_index=True)

    # Ajout de l'année
    max_year = df['Year'].max() if 'Year' in df.columns else "N/A"
    table = pd.concat([table, pd.DataFrame([["Année", max_year]], columns=["Élément", "Valeur"])], ignore_index=True)

    return table


def get_table_codification_parcelles(df):
    # Trouver la colonne des noms de parcelles
    parcelle_col = next((col for col in df.columns if col.strip() == "Parcelles culturales.Nom"), None)

    if not parcelle_col:
        st.warning("🟡 Colonne 'Parcelles culturales.Nom' introuvable dans le fichier.")
        return None

    # Créer la table de codification
    parcelle_names = df[parcelle_col].dropna().astype(str).str.strip().unique()
    df_codif = pd.DataFrame([list(parcelle_names), list(range(1, len(parcelle_names) + 1))])
    df_codif.index = ["Nom de la parcelle", "Code parcelle"]

    return df_codif


def get_table_operations_agricoles_codifie(df):
    # Colonnes nécessaires
    col_date = "Interventions des parcelles culturales.Date début"
    col_type = "Types d'interventions.Nom"
    col_parcelle = "Parcelles culturales.Nom"

    # Vérification des colonnes
    for col in [col_date, col_type, col_parcelle]:
        if col not in df.columns:
            st.warning(f"Colonne manquante : {col}")
            return None

    # Préparation des données
    df_op = df[[col_date, col_type, col_parcelle]].copy()
    df_op[col_date] = pd.to_datetime(df_op[col_date], errors='coerce', dayfirst=True)
    df_op = df_op.dropna(subset=[col_date])

    # Création du dictionnaire de codification
    parcelle_names = df_op[col_parcelle].dropna().astype(str).str.strip().unique()
    codif_dict = dict(zip(parcelle_names, range(1, len(parcelle_names) + 1)))

    # Application de la codification
    df_op["Code parcelle"] = df_op[col_parcelle].map(codif_dict)
    for code in codif_dict.values():
        df_op[str(code)] = df_op["Code parcelle"].apply(lambda x: "x" if x == code else "")

    # Regroupement par date et type d'intervention
    grouped = df_op.groupby([col_date, col_type], dropna=False)
    lignes_fusionnees = []

    for (date, type_interv), group in grouped:
        ligne = {col_date: date, col_type: type_interv}
        for code in codif_dict.values():
            ligne[str(code)] = 'x' if (group[str(code)] == 'x').any() else ''
        lignes_fusionnees.append(ligne)

    # Création du DataFrame final
    df_result = pd.DataFrame(lignes_fusionnees)
    df_result["Date"] = df_result[col_date].dt.strftime("%d/%m/%Y")
    df_result = df_result[["Date", col_type] + [str(code) for code in codif_dict.values()]]
    df_result.rename(columns={col_type: "Type d'intervention"}, inplace=True)

    return df_result


def get_table_irrigation(df):
    # Colonnes nécessaires
    col_type = "Types d'interventions.Nom"
    col_date = "Interventions des parcelles culturales.Date début"

    if col_type not in df.columns:
        st.error(f"❌ Colonne '{col_type}' introuvable dans le fichier.")
        return None

    # Filtrage des irrigations
    df_irrigation = df[df[col_type].str.lower().str.strip() == "irrigation"]

    # Traitement des dates
    if col_date in df_irrigation.columns:
        df_irrigation[col_date] = pd.to_datetime(df_irrigation[col_date], dayfirst=True, errors='coerce')
        df_irrigation = df_irrigation.sort_values(by=col_date)

    # Colonnes à conserver
    columns_to_keep = [
        "Interventions des parcelles culturales.Date début",
        "Intrants des parcelles culturales.Dose",
        "Interventions des parcelles culturales.Motivation",
        "Parcelles culturales.Nom"
    ]

    # Vérification des colonnes
    missing_cols = [col for col in columns_to_keep if col not in df_irrigation.columns]
    if missing_cols:
        st.error(f"❌ Colonnes manquantes dans les données : {', '.join(missing_cols)}")
        return None

    # Construction du tableau final
    df_result = df_irrigation[columns_to_keep].copy()
    df_result.insert(1, "Pluie", "")

    # Renommage des colonnes
    rename_dict = {
        "Interventions des parcelles culturales.Date début": "Date",
        "Pluie": "Pluie (mm)",
        "Intrants des parcelles culturales.Dose": "Dose",
        "Interventions des parcelles culturales.Motivation": "Motif",
        "Parcelles culturales.Nom": "Parcelle"
    }
    df_result.rename(columns=rename_dict, inplace=True)

    # Formatage de la date
    if "Date" in df_result.columns:
        df_result["Date"] = df_result["Date"].dt.strftime('%d/%m/%Y')

    return df_result


def get_table_fertilisation(df):
    # Types d'interventions considérés comme fertilisation
    mots_fertilisation = [
        "Amendements calco-magnésiens", "Biostimulant", "Boues de station d'épuration/compost urbain",
        "Effluents d'élevage", "Fertilisation minérale", "Fertilisation minérale Bulk",
        "Fertirrigation", "Obligo-éléments", "Organo-minéral", "Fertilisation organique",
        "Sous-produits/déchets alimentaires", "Sous-produits/déchets non alimentaires", "Supports de culture"
    ]

    # Colonnes nécessaires
    col_type = "Types d'interventions.Nom"
    col_date = "Interventions des parcelles culturales.Date début"
    col_dose = "Intrants des parcelles culturales.Dose"
    col_parcelle = "Parcelles culturales.Nom"

    if not all(col in df.columns for col in [col_type, col_date, col_dose, col_parcelle]):
        st.error("❌ Colonnes essentielles manquantes dans les données.")
        return None

    # Filtrage des fertilisations
    df_fertilisation = df[df[col_type].isin(mots_fertilisation)].copy()
    df_fertilisation[col_date] = pd.to_datetime(df_fertilisation[col_date], errors="coerce", dayfirst=True)

    # Mapping des colonnes
    colonne_mapping = {
        "Interventions des parcelles culturales.Date début": "📅 Date de l'intervention",
        "Traitements.Nom": "🧪 Produit",
        "Intrants des parcelles culturales.Dose": "💧 Dose",
        "Engrais.N": "🧬 N",
        "Engrais.P2O5": "🧬 P₂O₅",
        "Engrais.K2O": "🧬 K₂O",
        "Engrais.CaO": "🧬 CaO",
        "Engrais.MgO": "🧬 MgO",
        "Interventions des parcelles culturales.Observations": "📝 Observations",
        "Parcelles culturales.Nom": "🌿 Parcelle"
    }

    # Sélection des colonnes disponibles
    available_cols = [col for col in colonne_mapping.keys() if col in df_fertilisation.columns]
    df_result = df_fertilisation[available_cols].rename(columns=colonne_mapping)

    # Ajout des marqueurs de parcelles
    parcelles_uniques = df_fertilisation[col_parcelle].dropna().unique()
    for parcelle in parcelles_uniques:
        df_result[parcelle] = df_fertilisation[col_parcelle].apply(lambda x: 'x' if x == parcelle else '')

    # Regroupement et fusion des lignes
    group_columns = [col for col in ["📅 Date de l'intervention", "💧 Dose", "🧪 Produit",
                                     "🧬 N", "🧬 P₂O₅", "🧬 K₂O", "🧬 CaO", "🧬 MgO"]
                     if col in df_result.columns]

    try:
        grouped = df_result.groupby(group_columns, dropna=False)
    except KeyError as e:
        st.error(f"❌ KeyError: {str(e)}")
        return None

    lignes_fusionnees = []
    for group, group_df in grouped:
        ligne = group_df.iloc[0].copy()
        for parcelle in parcelles_uniques:
            ligne[parcelle] = 'x' if (group_df[parcelle] == 'x').any() else ''
        if "📝 Observations" in ligne:
            ligne["📝 Observations"] = " / ".join(group_df["📝 Observations"].dropna().unique())
        lignes_fusionnees.append(ligne)

    # Création du DataFrame final
    df_fertilisation = pd.DataFrame(lignes_fusionnees)
    if "📅 Date de l'intervention" in df_fertilisation.columns:
        df_fertilisation["📅 Date de l'intervention"] = df_fertilisation["📅 Date de l'intervention"].dt.strftime(
            "%d/%m/%Y")
    df_fertilisation.drop(columns=["🌿 Parcelle"], inplace=True, errors="ignore")

    return df_fertilisation


def get_table_traitement(df):
    # Colonnes nécessaires
    col_type = "Types d'interventions.Nom"
    col_date = "Interventions des parcelles culturales.Date début"
    col_dose = "Intrants des parcelles culturales.Dose"
    col_produit = "Traitements.Nom"
    col_cible = "Cibles à l'intrant.Nom de la cible"
    col_parcelle = "Parcelles culturales.Nom"

    # Types d'interventions exclus (fertilisations)
    mots_fertilisation = [
        "Amendements calco-magnésiens", "Biostimulant", "Boues de station d'épuration/compost urbain",
        "Effluents d'élevage", "Fertilisation minérale", "Fertilisation minérale Bulk",
        "Fertirrigation", "Obligo-éléments", "Organo-minéral", "Fertilisation organique",
        "Sous-produits/déchets alimentaires", "Sous-produits/déchets non alimentaires", "Supports de culture"
    ]

    if not all(col in df.columns for col in [col_type, col_date, col_dose, col_produit, col_parcelle]):
        st.error("❌ Colonnes essentielles manquantes dans les données.")
        return None

    # Filtrage des traitements (excluant irrigation et fertilisations)
    df_traitement = df[~df[col_type].isin(["irrigation"] + mots_fertilisation)].copy()
    df_traitement[col_date] = pd.to_datetime(df_traitement[col_date], errors='coerce', dayfirst=True)
    df_traitement = df_traitement.dropna(subset=[col_date])

    # Codification des parcelles
    parcelles_uniques = df_traitement[col_parcelle].dropna().unique()
    for parcelle in parcelles_uniques:
        df_traitement[parcelle] = df_traitement[col_parcelle].apply(lambda x: 'x' if x == parcelle else '')

    # Regroupement par date, produit, type et dose
    grouped = df_traitement.groupby([col_date, col_produit, col_type, col_dose], dropna=False)
    lignes_fusionnees = []

    for _, group in grouped:
        ligne = group.iloc[0].copy()
        if col_cible in group.columns:
            ligne["Cible"] = group[col_cible].dropna().astype(str).iloc[0] if not group[
                col_cible].dropna().empty else ''
        else:
            ligne["Cible"] = ''
        for parcelle in parcelles_uniques:
            ligne[parcelle] = 'x' if (group[parcelle] == 'x').any() else ''
        lignes_fusionnees.append(ligne)

    # Création du DataFrame final
    df_result = pd.DataFrame(lignes_fusionnees)
    df_result[col_date] = pd.to_datetime(df_result[col_date], errors='coerce')
    df_result["Date"] = df_result[col_date].dt.strftime("%d/%m/%Y")

    # Ajout des colonnes vides
    df_result.insert(3, "DAR", "")
    df_result.insert(6, "Commentaire", "")

    # Tri et organisation des colonnes
    df_result = df_result.sort_values(by=col_date)
    final_order = ["Date", col_produit, col_type, "DAR", col_dose, "Cible", "Commentaire"] + list(parcelles_uniques)
    df_result = df_result[final_order]

    # Renommage des colonnes
    rename_dict = {
        col_produit: "Produit commercial",
        col_type: "Matiere active",
        col_dose: "Dose appliquée par ha"
    }
    df_result.rename(columns=rename_dict, inplace=True)

    return df_result


def get_table_inventaire_parcelles(df):
    # Dictionnaire pour mapper les noms de colonnes
    colonne_mapping = {
        "Parcelles culturales.Nom": "Nom de la parcelle",
        "Variétés de parcelle.Nom": "Variété",
        "Parcelles culturales.Lieu-dit": "Lieu-dit",
        "Parcelles culturales.Surface": "Surface (ha)",
        "Parcelles culturales.PFI Verger éco responsable": "PFI Verger éco responsable",
        "Parcelles culturales.ZRP Zéro Résidu Pesticide": "ZRP Zéro Résidu Pesticide",
        "Parcelles culturales.Global Gap": "Global GAP",
        "Parcelles culturales.HVE 3": "HVE 3",
        "Autres": "Autres",
        "Suivi 1": "Suivi 1",
        "Suivi 2": "Suivi 2",
        "Suivi 3": "Suivi 3",
        "Conformité C": "C",
        "Conformité NC": "NC",
        "Motivation": "Motivation"
    }

    # Colonnes à extraire du DataFrame original
    extracted_cols = [
        "Parcelles culturales.Nom",
        "Variétés de parcelle.Nom",
        "Parcelles culturales.Lieu-dit",
        "Parcelles culturales.Surface",
        "Parcelles culturales.PFI Verger éco responsable",
        "Parcelles culturales.ZRP Zéro Résidu Pesticide",
        "Parcelles culturales.Global Gap",
        "Parcelles culturales.HVE 3"
    ]

    # Création du DataFrame temporaire
    result = pd.DataFrame()
    for col in extracted_cols:
        result[col] = df[col] if col in df.columns else ""

    # Suppression des doublons
    result = result.drop_duplicates()

    # Ajout des colonnes vides supplémentaires
    result["Autres"] = ""
    result["Suivi 1"] = ""
    result["Suivi 2"] = ""
    result["Suivi 3"] = ""
    result["Conformité C"] = ""
    result["Conformité NC"] = ""
    result["Motivation"] = ""

    # Renommage final des colonnes
    result = result.rename(columns=colonne_mapping)

    return result


def export_all_tables_to_excel(table_dict, raison_sociale):
    # Création du nom de fichier avec raison sociale et date
    date_aujourdhui = datetime.now().strftime("%Y")
    nom_fichier = f"Cahier_Cultural_{raison_sociale}_{date_aujourdhui}.xlsx"

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in table_dict.items():
            if df is not None:
                df.to_excel(writer, index=False,
                            sheet_name=sheet_name[:31])  # Limite de 31 caractères pour les noms de feuilles

    st.download_button(
        label="📥 Télécharger toutes les tables (Excel)",
        data=output.getvalue(),
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def main():
    st.title("Cahier culturel")

    # Téléchargement du fichier
    uploaded_file = st.file_uploader("Téléchargez un fichier .txt", type=["txt"])
    if uploaded_file is not None:
        # Chargement des données
        df = charger_fichier(uploaded_file)

        if df is not None:
            # Nettoyage des noms de colonnes
            df = nettoyer_noms_colonnes(df)

            # Sauvegarde d'une copie originale
            df_original = df.copy()

            # Traitement des données
            df = traiter_donnees(df)

            # Affichage des données filtrées
            st.subheader("Tableau des Données Filtrées")
            st.write(df)

            # Génération de toutes les tables
            tables = {
                "Exploitation": get_table_exploitations_parcelles(df),
                "Codification Parcelles": get_table_codification_parcelles(df),
                "Inventaire Parcelles": get_table_inventaire_parcelles(df),
                "Operation agricole": get_table_operations_agricoles_codifie(df),
                "Traitement": get_table_traitement(df),
                "Fertilisation": get_table_fertilisation(df),
                "Irrigation": get_table_irrigation(df),
            }

            # Récupération de la raison sociale pour le nom du fichier
            raison_sociale = "EARL_de_Fleury"  # Valeur par défaut
            if tables["Exploitation"] is not None:
                try:
                    rs_row = tables["Exploitation"][tables["Exploitation"]["Élément"] == "Raison sociale"]
                    if not rs_row.empty:
                        raison_sociale = rs_row.iloc[0]["Valeur"]
                        # Nettoyage pour un nom de fichier valide
                        raison_sociale = raison_sociale.replace(" ", "_").replace("/", "_").strip()
                except Exception as e:
                    st.warning(f"Impossible de récupérer la raison sociale : {e}")

            # Affichage des tables
            for name, table in tables.items():
                if table is not None:
                    st.subheader(name)
                    st.dataframe(table)

            # Bouton d'export
            if all(table is not None for table in tables.values()):
                export_all_tables_to_excel(tables, raison_sociale)
            else:
                st.warning("Certaines tables n'ont pas pu être générées. Vérifiez les données d'entrée.")


if __name__ == "__main__":
    main()
