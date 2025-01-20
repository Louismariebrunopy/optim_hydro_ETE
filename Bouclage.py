import pandas as pd
import os
import subprocess
import time

# Définir le répertoire de travail
#os.chdir('/Users/j.coeuillet/Library/Mobile Documents/com~apple~CloudDocs/ISAE-SUPAERO/ETE/Julia/Optim Julia')

# Chemin du fichier des résultats
filePath = "Output/results_k_step.csv"
filePath_border = "Output/results_k_effetbord.csv"

# Convertir les résultats Julia en un DataFrame 
data = pd.read_csv(filePath, delimiter=",")
bord = pd.read_csv(filePath_border, delimiter=",")
bord = bord.round(0) # Pour mettre les valeurs en binaire, car sinon on a des e-15

# Définir le nombre de semaines simulées
nb = 52

# Créer un DataFrame vide pour accumuler les résultats
final_data = pd.DataFrame()

# Initialisation du temps de référence
previous_time = time.time()

# Boucle pour exécuter le script Julia et stocker les résultats
for semaine in range(1, nb + 1):  # On boucle sur nb semaines
    
    # Stocker les effets de bord dans un Excel pour Julia (pas encore utilisé)
    bord.to_excel("Données/bord.xlsx", index=False)

    # Stocker le numéro de semaine dans un Excel pour Julia
    colonne_semaine = {"Numéro de semaine courant": [semaine-1]}
    df = pd.DataFrame(colonne_semaine)
    df.to_excel("Données/week.xlsx", index=False)

    # Lancer la simulation Julia pour la k-ème semaine
    try:
        print(f"Exécution {semaine}/{nb}...")
        subprocess.run(["julia", "optim_hydro.jl"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de l'exécution du script Julia à l'itération {semaine}: {e}")
        continue

    # Charger les résultats
    try: 

        data = pd.read_csv(filePath, delimiter=",", decimal=".")
        data = data.round(7)
        bord = pd.read_csv(filePath_border, delimiter=",")
        bord = bord.round(0) # Pour mettre les valeurs en binaire, car sinon on a des e-15

        # Ajouter les colonnes Jour et Heure avec des valeurs vides
        if 'Semaine' not in data.columns:
            data.insert(0, 'Semaine', '')  # Insérer la colonne 'Semaine'
        if 'Jour' not in data.columns:    
            data.insert(1, 'Jour', '')  # Insérer la colonne 'Jour'
        if 'Heure' not in data.columns:
            data.insert(2, 'Heure', '')  # Insérer la colonne 'Heure'
        if 'Jour cumulé' not in data.columns:
            data.insert(3, 'Jour cumulé', '') # Insérer la colonne 'Jour cumulé'
        if 'Heure cumulée' not in data.columns:
            data.insert(4, 'Heure cumulée', '') # Insérer la colonne 'Heure cumulée'

        # Numéroter les semaines (de 1 à nb)
        data['Semaine'] = [semaine for i in range(168)]  # Diviser les heures totales en semaines

        # Numéroter les jours (de 1 à 7*nb)
        data['Jour'] = [(i // 24) % 7 + 1 for i in range(168)]  # Numéroter les jours dans la semaine (de 1 à 7)

        # Numéroter les heures (de 1 à 24)
        data['Heure'] = [i % 24 + 1 for i in range(168)]  # Numéroter les heures de la journée (de 1 à 24)

        # Numéroter les jours (de 1 à 365)
        data['Jour cumulé'] = data['Jour'] + [(semaine-1)*7]

        # Numéroter les heures (de 1 à 8760)
        data['Heure cumulée'] = data['Heure'] + [(semaine-1)*168]
        
        # Réorganiser les colonnes pour que Semaine, Jour et Heure soient à gauche
        #data = data[['Semaine', 'Jour', 'Heure'] + [col for col in data.columns if col not in ['Semaine', 'Jour', 'Heure']]]

        # Empiler les résultats dans le DataFrame final
        final_data = pd.concat([final_data, data], ignore_index=True)

    except FileNotFoundError:
        print(f"Fichier {filePath} non trouvé après l'itération {semaine}")
    except pd.errors.EmptyDataError:
        print(f"Le fichier {filePath} est vide après l'itération {semaine}")

    current_time = time.time()
    elapsed_time = current_time - previous_time
    print(f"…réalisée en {elapsed_time:.2f} secondes")
    previous_time = current_time

# Chemin vers le fichier Excel existant
fichier_excel = "Output/final_results.xlsx"

# Ajouter le DataFrame à la feuille de sortie
#final_data = final_data.round(7) # Pour ne pas finir avec des e-14…

# Convertir les colonnes numériques avec une virgule comme séparateur de décimale
#final_data = final_data.map(lambda x: f"{x:.7f}".replace(',', '.') if isinstance(x, float) else x)

with pd.ExcelWriter(fichier_excel, mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
    final_data.to_excel(writer, sheet_name="Sortie_Python", index=False)
    
subprocess.run(["open", "Output/final_results.xlsx"])