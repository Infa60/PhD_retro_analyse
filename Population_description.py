import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.ticker as ticker
# ==========================================
# 1. CONFIGURATION
# ==========================================
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\CP_INCLUSION_VISIT_LIST.xlsx"
sheet_name = "Include_file"

# --- CHOISISSEZ VOTRE CHEMIN ICI ---
# Vous pouvez mettre un chemin complet comme r"C:\Users\Nom\Documents\Analyses"
output_folder = r"C:\Users\bourgema\OneDrive - Université de Genève\PHD\Part1\Population_visit_plot"

# ==========================================
# 2. DATA LOADING & PRE-PROCESSING
# ==========================================
# Vérification et création du dossier de sauvegarde
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print(f"Dossier créé : {output_folder}")

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df['DateVisite'] = pd.to_datetime(df['DateVisite'], errors='coerce')
    df['AnneeVisite'] = df['DateVisite'].dt.year
    df['CoteDiagnostic'] = df['CoteDiagnostic'].fillna('Missing')
except Exception as e:
    print(f"Error loading file: {e}")
    exit()


# ==========================================
# 3. ANALYSIS FUNCTION
# ==========================================
def generate_plots(data, group_title):
    if data.empty:
        print(f"No data found for group: {group_title}")
        return

    # Taille compacte (16x5)
    fig, axes = plt.subplots(1, 3, figsize=(16, 5))
    fig.suptitle(f"Analysis: {group_title}", fontsize=16, fontweight='bold', y=1.05)

    # --- A. Age Distribution ---
    age_min = int(data['Age'].min())
    age_max = int(data['Age'].max())
    sns_hist = sns.histplot(data=data, x='Age', bins=range(age_min, age_max + 2),
                            ax=axes[0], color="#4c72b0", edgecolor='white')

    for p in sns_hist.patches:
        val = p.get_height()
        if val > 0:
            if val < 3:
                axes[0].annotate(f'{int(val)}', (p.get_x() + p.get_width() / 2., val),
                                 ha='center', va='bottom', color='black',
                                 fontsize=9, fontweight='bold', xytext=(0, 2), textcoords='offset points')
            else:
                axes[0].annotate(f'{int(val)}', (p.get_x() + p.get_width() / 2., val / 2.),
                                 ha='center', va='center', color='white',
                                 fontsize=9, fontweight='bold')

    axes[0].set_title("Age Distribution", fontsize=12)
    axes[0].set_xlabel("Age (years)", fontsize=10)
    axes[0].set_ylabel("Count", fontsize=10)
    axes[0].yaxis.set_major_locator(ticker.MaxNLocator(integer=True))

    # --- B. Laterality Distribution ---
    sns_count = sns.countplot(data=data, x='CoteDiagnostic', ax=axes[1], palette="viridis")
    for p in axes[1].patches:
        val = p.get_height()
        if val > 0:
            if val < 3:
                axes[1].annotate(f'{int(val)}', (p.get_x() + p.get_width() / 2., val),
                                 ha='center', va='bottom', color='black',
                                 fontsize=9, fontweight='bold', xytext=(0, 2), textcoords='offset points')
            else:
                axes[1].annotate(f'{int(val)}', (p.get_x() + p.get_width() / 2., val / 2.),
                                 ha='center', va='center', color='white',
                                 fontsize=9, fontweight='bold')

    axes[1].set_title("Laterality", fontsize=12)
    axes[1].set_xlabel("Affected Side", fontsize=10)
    axes[1].set_ylabel("Count", fontsize=10)

    # --- C. Visit History ---
    visits_per_year = data.groupby('AnneeVisite').size().reset_index(name='VisitCount')
    visits_per_year = visits_per_year.dropna()

    sns.lineplot(data=visits_per_year, x='AnneeVisite', y='VisitCount',
                 ax=axes[2], marker='o', color='firebrick', linewidth=2, markersize=6)

    for x, y in zip(visits_per_year['AnneeVisite'], visits_per_year['VisitCount']):
        axes[2].annotate(f'{int(y)}', (x, y), textcoords="offset points",
                         xytext=(0, 10), ha='center', fontsize=8, fontweight='bold')

    axes[2].yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    axes[2].set_title("Visit Evolution", fontsize=12)
    axes[2].set_xlabel("Year", fontsize=10)
    axes[2].set_ylabel("Number of Visits", fontsize=10)
    axes[2].grid(True, linestyle='--', alpha=0.3)

    # Dates tous les 5 ans
    min_year = int(visits_per_year['AnneeVisite'].min())
    max_year = int(visits_per_year['AnneeVisite'].max())
    five_year_labels = range(min_year, max_year + 1, 5)
    axes[2].set_xticks(five_year_labels)
    plt.setp(axes[2].get_xticklabels(), rotation=90, fontsize=9)

    plt.tight_layout()

    # --- ENREGISTREMENT DANS LE DOSSIER CHOISI ---
    filename = f"{group_title.replace(' ', '_').lower()}_analysis.png"
    full_save_path = os.path.join(output_folder, filename)

    plt.savefig(full_save_path, bbox_inches='tight')
    print(f"Graphique enregistré dans : {full_save_path}")

    plt.show()


# ==========================================
# 4. EXECUTION
# ==========================================
df_hemi = df[df['Diagnostic'].str.startswith('hemi', na=False)].copy()
generate_plots(df_hemi, "Hemiplegic Group")

df_diplegie = df[df['Diagnostic'].str.startswith('diplegie', na=False)].copy()
generate_plots(df_diplegie, "Diplegic Group")

generate_plots(df, "All Children")