import requests
import openpyxl
from requests.exceptions import SSLError

# Charger le fichier Excel
workbook = openpyxl.load_workbook('urls.xlsx')
worksheet = workbook['Feuil1']  # Remplacez 'Feuil1' par le nom de votre feuille

# Boucle sur chaque ligne du fichier Excel
for row in range(2, worksheet.max_row + 1):  # Commence à la ligne 2 pour éviter l'en-tête
    url = worksheet.cell(row, 1).value

    try:
        # Obtenir le code HTTP
        response = requests.get(url, verify=False)  # Désactiver la vérification SSL
        http_code = response.status_code

        # Créer la chaîne de redirections
        redirection_chain = ' > '.join([str(resp.status_code) for resp in response.history] + [str(http_code)])

        # Obtenir l'URL finale après redirections
        final_url = response.url

        # Écrire les données dans les colonnes B, C et D
        worksheet.cell(row, 2).value = redirection_chain
        worksheet.cell(row, 3).value = len(response.history)
        worksheet.cell(row, 4).value = final_url

    except SSLError as ssl_err:
        # Gérer les erreurs de certificat SSL
        print(f"Erreur SSL pour l'URL {url}: {ssl_err}")
        worksheet.cell(row, 2).value = 'Erreur SSL'
        worksheet.cell(row, 3).value = 'N/A'
        worksheet.cell(row, 4).value = 'N/A'

    except requests.exceptions.RequestException as e:
        # Gérer les autres erreurs de requête
        print(f"Erreur pour l'URL {url}: {e}")
        worksheet.cell(row, 2).value = 'Erreur'
        worksheet.cell(row, 3).value = 'N/A'
        worksheet.cell(row, 4).value = 'N/A'

# Enregistrer les modifications dans le fichier Excel
workbook.save('urls.xlsx')