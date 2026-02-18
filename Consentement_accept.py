import pandas as pd
import os

# =======================================================
# 1. CONFIGURATION (À MODIFIER PAR VOS CHEMINS DE FICHIERS)
# =======================================================

# Ch
# Chemin vers le fichier toutes visites
FICHIER_TOUTES_LES_VISITES = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\Export_data\CP_pathologie_good_age_good_treatment_unique_visite.xlsx"

output_folder = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1"
file_c = os.path.join(output_folder, "CP_INCLUSION_visit_list.xlsx")

# Colonnes d'identification
COLONNE_ID = "ID_Patient"
COLONNE_DATE = "DateVisite"

# Date de référence pour le filtre (17 mars 2017)
DATE_LIMITE = pd.to_datetime("17.03.2017", format="%d.%m.%Y")

# =======================================================
# 2. CHARGEMENT ET PRÉPARATION DES DONNÉES
# =======================================================

try:
    df_inclusion_list = pd.read_excel(FICHIER_TOUTES_LES_VISITES)

    # Convertir la colonne de date en objet datetime (crucial)
    # Le format "%d.%m.%Y" gère votre format "jj.mm.aaaa"
    df_inclusion_list[COLONNE_DATE] = pd.to_datetime(
        df_inclusion_list[COLONNE_DATE],
        format="%d.%m.%Y",
        errors='coerce'  # Met 'NaT' si une date n'est pas valide
    )

except FileNotFoundError as e:
    print(f"ERREUR : Fichier non trouvé. Veuillez vérifier les chemins : {e}")
    exit()
except Exception as e:
    print(f"Une erreur est survenue lors du chargement ou de la conversion : {e}")
    exit()

# =======================================================
# 3. FILTRAGE ET IDENTIFICATION
# =======================================================

df_inclusion_list_only_CP = df_inclusion_list[df_inclusion_list['CP'] != 'No'].copy()

df_inclusion_list_only_CP_only_patient = df_inclusion_list_only_CP[df_inclusion_list_only_CP['Research'] != 'Yes'].copy()


# 2. On crée un masque (True/False) : est-ce que la visite est avant la limite ?
df_inclusion_list_only_CP_only_patient['est_avant_limite'] = df_inclusion_list_only_CP_only_patient['DateVisite'] < DATE_LIMITE

# 3. On vérifie si TOUTES les lignes de chaque enfant sont True
# Remplacez 'enfant' par le nom exact de votre colonne d'identifiant (ex: 'ID', 'Nom')
condition_toutes_visites = df_inclusion_list_only_CP_only_patient.groupby('ID_Patient')['est_avant_limite'].transform('all')

# 4. On remplit la colonne selon le résultat
condition_finale = condition_toutes_visites & df_inclusion_list_only_CP_only_patient['Consentement_Generalise'].isna()

df_inclusion_list_only_CP_only_patient.loc[condition_finale, 'Consentement_Generalise'] = 'Oui_by_default'

# Optionnel : supprimer la colonne temporaire de calcul
# df = df.drop(columns=['est_avant_limite'])

df_inclusion_list_only_CP_only_patient.to_excel(file_c, index=False)