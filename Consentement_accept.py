import pandas as pd

# =======================================================
# 1. CONFIGURATION (À MODIFIER PAR VOS CHEMINS DE FICHIERS)
# =======================================================

# Chemin vers le fichier 1 : Patients à vérifier
FICHIER_PATIENTS_A_VERIFIER = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_children_diagnostic_visit_no_wrong_traitement_unique_patient_Updated.xlsx"

# Chemin vers le fichier 2 : Toutes les visites
FICHIER_TOUTES_LES_VISITES = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_children_diagnostic_visit_traitement_timing.xlsx"

# Colonnes d'identification
COLONNE_ID = "ID_Patient"
COLONNE_DATE = "DateVisite"

# Date de référence pour le filtre (17 mars 2017)
DATE_LIMITE = pd.to_datetime("17.03.2017", format="%d.%m.%Y")

# =======================================================
# 2. CHARGEMENT ET PRÉPARATION DES DONNÉES
# =======================================================

try:
    # Charger les deux fichiers Excel
    df_patients_a_verifier = pd.read_excel(FICHIER_PATIENTS_A_VERIFIER, sheet_name="Check_consentement")
    df_toutes_les_visites = pd.read_excel(FICHIER_TOUTES_LES_VISITES)

    # Convertir la colonne de date en objet datetime (crucial)
    # Le format "%d.%m.%Y" gère votre format "jj.mm.aaaa"
    df_toutes_les_visites[COLONNE_DATE] = pd.to_datetime(
        df_toutes_les_visites[COLONNE_DATE],
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

# 3.1. Filtrer les visites qui ont eu lieu APRÈS la date limite
# On utilise le signe '>' pour "strictement après"
df_visites_apres_limite = df_toutes_les_visites[
    df_toutes_les_visites[COLONNE_DATE] > DATE_LIMITE
    ].copy()

# 3.2. Récupérer la liste unique des ID des patients ayant une visite après la limite
patients_avec_visite_recente = df_visites_apres_limite[COLONNE_ID].unique()

# 3.3. Filtrer le fichier initial (Patients à vérifier)
# On garde uniquement les patients dont l'ID est dans la liste 'patients_avec_visite_recente'
df_patients_resultat = df_patients_a_verifier[
    df_patients_a_verifier[COLONNE_ID].isin(patients_avec_visite_recente)
].copy()

# =======================================================
# 4. AFFICHAGE ET SAUVEGARDE DES RÉSULTATS
# =======================================================

print("=" * 60)
print("ANALYSE DES VISITES APRÈS LE 17 MARS 2017")
print("=" * 60)

if not df_patients_resultat.empty:
    print(
        f"Nombre total de patients ayant eu une visite après le {DATE_LIMITE.strftime('%d/%m/%Y')} : {len(df_patients_resultat)}")
    print("\nListe des patients du premier fichier qui remplissent la condition (ID_Patient) :")
    print("-" * 50)
    # Afficher uniquement la colonne ID_Patient pour une liste claire
    print(df_patients_resultat[COLONNE_ID].to_string(index=False))

    # Suggestion de sauvegarde
    NOM_FICHIER_RESULTAT = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\Patients_Visites_Recentes.xlsx"
    df_patients_resultat.to_excel(NOM_FICHIER_RESULTAT, index=False)
    print(f"\n--- Le résultat a été sauvegardé dans : {NOM_FICHIER_RESULTAT} ---")
else:
    print(f"Aucun patient du fichier initial n'a eu de visite après le {DATE_LIMITE.strftime('%d/%m/%Y')}.")