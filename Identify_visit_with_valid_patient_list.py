import pandas as pd
import os

# ==========================================
# 1. CONFIGURATION
# ==========================================
# Chemins des fichiers (modifiez avec vos chemins réels)
file_a = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_INCLUSION_LIST.xlsx"
file_b = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_children_diagnostic_visit_no_wrong_traitement_unique_visit.xlsx"

# Dossier où enregistrer le fichier C
output_folder = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1"
file_c = os.path.join(output_folder, "CP_INCLUSION_visit_list.xlsx")

# Nom de la colonne commune (Identifiant)
col_id = "ID_Patient"

# ==========================================
# 2. CHARGEMENT ET FILTRAGE
# ==========================================
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

try:
    # Charger les données
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # On ne garde que la colonne des identifiants du fichier A et on enlève les doublons
    # pour éviter de multiplier les lignes inutilement lors de la jointure
    ids_a = df_a[[col_id]].drop_duplicates()

    # Jointure interne (Inner Merge) :
    # On garde toutes les lignes de B qui correspondent à un ID présent dans A
    df_c = pd.merge(ids_a, df_b, on=col_id, how='inner')

    # Sauvegarde du résultat
    df_c.to_excel(file_c, index=False)

    print(f"✅ Opération réussie !")
    print(f"Le fichier C contient {len(df_c)} lignes et a été enregistré ici : {file_c}")

except Exception as e:
    print(f"❌ Une erreur est survenue : {e}")