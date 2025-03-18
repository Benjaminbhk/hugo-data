# Hugo Data: Application de traitement de données Bloomberg

Cette application Streamlit permet de traiter des fichiers Excel (xlsx) provenant de Bloomberg, relatifs aux échanges financiers. Elle offre une interface utilisateur conviviale pour :

- **Charger** un ou plusieurs fichiers Excel (xlsx) contenant des données de Bloomberg.
- **Traiter** ces données selon des règles définies.
- **Générer** un fichier Excel (xlsx) avec les résultats du traitement, téléchargeable directement depuis l'interface.

## Fonctionnalités

- **Interface utilisateur intuitive** : Sélectionnez et chargez un ou plusieurs fichiers Excel.
- **Traitement des données** : Analyse et transformation des données Bloomberg selon des règles prédéfinies.
- **Export** : Génération d'un fichier Excel prêt à être utilisé dans votre workflow.

## Prérequis

- Python 3.7 ou supérieur.
- Les dépendances Python listées dans le fichier `requirements.txt`.

Pour les installer, utilisez la commande :

```bash
pip install -r requirements.txt
```
## Lancer l'application

Pour démarrer l'application Streamlit, exécutez la commande suivante à partir de la racine du dépôt :

```bash
streamlit run app/streamlite_app.py
```
