import pandas as pd
import numpy as np

# =======================================================
# 1. CONFIGURATION (À MODIFIER)
# =======================================================

# 1.1 Chemins des fichiers
# Chemin vers le fichier A (celui à modifier)
FICHIER_A_CHEMIN = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_children_diagnostic_visit_no_wrong_traitement_unique_patient.xlsx"
# Chemin vers le fichier B (la source de vérité)
FICHIER_B_CHEMIN = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\Draft\Patient_include.xlsx"

# 1.2 Noms des feuilles
FEUILLE_A = "Sheet 1"  # Nom de la feuille dans le fichier A
FEUILLE_B = "CP_Consentement"  # Nom de la feuille dans le fichier B

# 1.3 Noms des colonnes
COLONNE_PRENOM = "ID_Patient"  # Nom de la colonne contenant le prénom dans les deux fichiers
COLONNE_A_REMPLIR = "To_include"  # Nom de la colonne à remplir (ou à créer) dans le fichier A

# =======================================================
# 2. CHARGEMENT ET PRÉPARATION DES DONNÉES
# =======================================================

try:
    # 2.1 Charger les données
    df_A = pd.read_excel(FICHIER_A_CHEMIN, sheet_name=FEUILLE_A)
    df_B = pd.read_excel(FICHIER_B_CHEMIN, sheet_name=FEUILLE_B)

    # 2.2 S'assurer que les colonnes clés existent
    if COLONNE_PRENOM not in df_A.columns or COLONNE_PRENOM not in df_B.columns:
        raise ValueError("La colonne 'Prenom' n'est pas trouvée dans les deux fichiers.")

    # 2.3 Nettoyage des prénoms (uniformisation pour la jointure)
    # Convertir en minuscules et supprimer les espaces autour
    df_A[COLONNE_PRENOM] = df_A[COLONNE_PRENOM].astype(str).str.lower().str.strip()
    df_B[COLONNE_PRENOM] = df_B[COLONNE_PRENOM].astype(str).str.lower().str.strip()

except FileNotFoundError:
    print("Erreur : Un ou plusieurs fichiers Excel sont introuvables. Vérifiez les chemins.")
    exit()
except ValueError as e:
    print(f"Erreur de colonne : {e}")
    exit()
except Exception as e:
    print(f"Une erreur inattendue est survenue : {e}")
    exit()

# =======================================================
# 3. LOGIQUE DE VÉRIFICATION ET REMPLISSAGE
# =======================================================

# 3.1 Créer la liste unique des prénoms présents dans le Fichier B
prenoms_consentement = df_B[COLONNE_PRENOM].unique()

# 3.2 Créer la nouvelle colonne 'Include' dans df_A
# Utiliser .isin() pour vérifier si le prénom de df_A est dans la liste.
# Le résultat est une série de True/False.
est_present = df_A[COLONNE_PRENOM].isin(prenoms_consentement)

# 3.3 Remplir la colonne 'Include'
# Si c'est True, mettre 'Yes', sinon (False), mettre 'No' (ou utiliser NaN pour vide)
df_A[COLONNE_A_REMPLIR] = np.where(est_present, "Yes", "No")

# =======================================================
# 4. SAUVEGARDE DU RÉSULTAT
# =======================================================

# Définir le chemin du nouveau fichier de sortie (on ajoute '_Updated' pour ne pas écraser l'original)
FICHIER_A_SORTIE = FICHIER_A_CHEMIN.replace(".xlsx", "_Updated.xlsx")

try:
    # Sauvegarder le DataFrame A modifié.
    # NOTE: Ceci écrira SEULEMENT la feuille utilisée (FEUILLE_A) dans le nouveau fichier.
    df_A.to_excel(FICHIER_A_SORTIE, sheet_name=FEUILLE_A, index=False)

    print("=" * 70)
    print("✅ Opération terminée avec succès.")
    print(f"Le fichier A mis à jour est enregistré sous : {FICHIER_A_SORTIE}")
    print(f"La colonne '{COLONNE_A_REMPLIR}' a été remplie avec 'Yes' ou 'No'.")
    print("=" * 70)

except Exception as e:
    print(f"Erreur lors de la sauvegarde du fichier : {e}")