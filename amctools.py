import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font

# Fonction de traitement pour le fichier Excel
def process_excel(file):
    try:
        # Lire le fichier Excel sans header
        xls = pd.read_excel(file, header=None)
                
        # Trouver l'index de la ligne d'en-tête
        header_index = next(
            (idx for idx, row in xls.iterrows() if all(col in row.values for col in ['Code', 'Nom', 'Prénom'])),
            None
        )
        
        if header_index is None:
            st.error("Les colonnes 'Code', 'Nom', 'Prénom' sont introuvables dans le fichier.")
            return None, None
        
        # Redéfinir les en-têtes et supprimer les lignes précédentes
        xls.columns = xls.iloc[header_index]
        xls = xls.iloc[header_index + 1:].reset_index(drop=True)
        
        # Vérification si le fichier est vide après nettoyage
        if xls.empty:
            st.error("Aucune donnée valide après le traitement des lignes.")
            return None, None
        
        # Vérification des colonnes nécessaires (double vérification)
        required_columns = ['Nom', 'Prénom', 'Code']
        missing = [col for col in required_columns if col not in xls.columns]
        if missing:
            st.error(f"Colonnes manquantes après traitement : {', '.join(missing)}")
            return None, None
        
        # Nettoyage des données
        liste = xls.dropna(subset=['Nom', 'Prénom', 'Code'])
        liste['Name'] = liste['Code'].astype(str) + ' ' + liste['Nom'] + ' ' + liste['Prénom']
        liste = liste[['Code', 'Name']].drop_duplicates()
        return xls, liste
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None

# Fonction de traitement pour le fichier CSV
def process_csv(excel_file, csv_file):
    try:
        xls, liste = process_excel(excel_file) 
        
        csv = pd.read_csv(csv_file, delimiter=';', encoding='utf-8')

        # Initialiser un DataFrame pour les anomalies
        anomalies = pd.DataFrame()

        # 1. Filtrer les lignes avec 'NONE' ou valeurs vides
        csv = csv[['A:Code', 'Code', 'Nom', 'Note']]
        none_mask = (csv['A:Code'] == 'NONE') 
        anomalies = pd.concat([anomalies, csv[none_mask]])
        csv_clean = csv[~none_mask].copy()
        
        # Vérifier si le fichier nettoyé est vide
        if csv_clean.empty:
            st.error("Aucune donnée valide après le nettoyage !")
            return None, None, None
    
        # Fusionner les données sur la colonne 'Code'
        df_merged = pd.merge(xls, csv_clean[['Code', 'Note']], on='Code', how='left')
        if 'Note_x' in df_merged.columns and 'Note_y' in df_merged.columns:
            df_merged.rename(columns={'Note_y': 'Note'}, inplace=True)
            df_merged.drop(columns=['Note_x'], inplace=True)
        return csv_clean, df_merged, anomalies
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier Excel : {str(e)}")
        st.info("Assurez-vous que le fichier est bien formaté et contient les colonnes requises.")
        return None, None, None

# Fonction pour générer un fichier Excel valide avec un en-tête personnalisé "Medkour_DLGE15_2025"
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Écrire le DataFrame à partir de la ligne 3 (index 2) pour laisser la place à l'en-tête personnalisé
        df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=2)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Déterminer le nombre de colonnes du DataFrame
        num_cols = df.shape[1]
        last_col = get_column_letter(num_cols)
    
    output.seek(0)  # Repositionner le curseur au début du buffer
    return output.getvalue()

# ----------------- Interface utilisateur -----------------

st.title("Traitements de fichiers Excel et CSV pour AMC")

# Onglets pour séparer les sections
tab1, tab2 = st.tabs(["Liste des étudiants", "Traitement des notes"])

with tab1:
    st.header("Préparation de la liste des étudiants")
    st.info(
        """
        - Télécharger le fichier Excel de l'administration. 
        - Les en-têtes 'Nom', 'Prénom' et 'Code' seront détectés.
        - Les lignes avant les en-têtes seront automatiquement supprimées.
        - La liste des étudiants à fournir à AMC sera préparée au format Excel avec un en-tête personnalisé.
        """
    )
    
    uploaded_excel_file = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader"
    )
    
    if uploaded_excel_file is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file)
            
            if xls is not None:
                st.success(f"Lecture du fichier Excel réussie ! {len(xls)} étudiants trouvés.")  
                st.write("Aperçu de la base de données des étudiants avant traitement automatique :")
                st.write(xls.head(10))              
                st.write("Aperçu de la liste des étudiants à fournir à AMC:")
                st.write(liste.head(10))
                st.success(f"La liste contient {len(xls)} étudiants.")

                # Générer le fichier Excel avec en-tête personnalisé
                excel_data = to_excel(liste)
                st.download_button(
                    label="📥 Télécharger la liste des étudiants au format Excel",
                    data=excel_data,
                    file_name="liste_etudiants.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Générer le fichier CSV
                csv_data = liste.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Télécharger la liste des étudiants au format CSV",
                    data=csv_data,
                    file_name="liste_etudiants.csv",
                    mime="text/csv"
                )

with tab2:
    st.header("Traitement des notes")
    st.info(
        """
        - Télécharger le fichier Excel de l'administration. 
        - Télécharger le fichier CSV des notes calculées par AMC.
        - Les notes seront automatiquement associées aux étudiants.
        - Le nombre d'étudiants mal identifiés sera indiqué.
        """
    )

    uploaded_excel_file2 = st.file_uploader(
        "Télécharger le fichier Excel de l'administration", 
        type="xlsx", 
        key="excel_uploader2"
    )

    uploaded_csv_file = st.file_uploader(
        "Télécharger le fichier CSV des notes calculées par AMC", 
        type="csv", 
        key="csv_uploader"
    )

    if uploaded_excel_file2 is not None:
        with st.spinner("Traitement automatique du fichier Excel en cours..."):
            xls, liste = process_excel(uploaded_excel_file2)

    if uploaded_csv_file is not None and uploaded_excel_file2 is not None:
        with st.spinner("Intégration des notes aux étudiants..."):
            csv_clean, df_merged, anomalies = process_csv(uploaded_excel_file2, uploaded_csv_file)

            st.write("Aperçu de la base de données des étudiants :")
            st.write(xls.head(10))   
            
            st.write("Aperçu du fichier des notes :")
            st.write(csv_clean.head(10))

            st.write("Aperçu de la base de données des étudiants alimentée par les notes :")
            st.write(df_merged.head(10))
                
            # Affichage des statistiques
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Effectif total", len(df_merged))
            with col2:
                st.metric("Présents", len(csv_clean))
            with col3:
                st.metric("Absents", len(df_merged) - len(csv_clean) - len(anomalies))
            with col4:
                st.metric("Mal identifiés", len(anomalies))
                
            # Générer le fichier Excel final avec en-tête personnalisé
            excel_data = to_excel(df_merged)
            st.download_button(
                label="📥 Télécharger le fichier final des notes au format Excel",
                data=excel_data,
                file_name="etudiants_avec_notes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
                
            if len(anomalies) > 0:   
                st.error(f"Attention! {len(anomalies)} étudiants ont été mal identifiés. Vérifiez leurs copies.")