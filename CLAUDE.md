# hugo-data

Outil de traitement de fichiers Excel Bloomberg pour un ami (Hugo).
Application Streamlit déployée — **outil fini, pas en développement actif**.

## Ce que fait l'app

1. Charge un ou plusieurs fichiers `.xlsx` Bloomberg (trades financiers)
2. Détecte les paires de **Roll** (deux legs sur le même sous-jacent, prix et taille proches)
3. Catégorise chaque ligne : Roll (L1/L2), Roll Screen, Roll Client, Outright, Autre
4. Génère un fichier Excel avec formules Bloomberg (`BDH`) intégrées

## Structure

```
streamlit_app.py       # UI Streamlit (entry point déploiement)
app/
  processing.py        # Toute la logique métier — source unique
  recap.py             # Dédup multi-fichiers + génération du recap texte
  status.py            # Persistance des jours traités (~/.hugo-data)
  script-hugo.py       # UI tkinter (usage local)
```

## Règles importantes

- **Ne jamais modifier `main`** sans accord — Hugo utilise la branche main en production
- Travailler sur la branche `optimisation` ou une nouvelle branche
- Tag de sauvegarde : `backup/before-optimisation` (état avant refactor juin 2026)

## Logique métier clé

- **Détection Roll** : O(n²) stateful — pas vectorisable, intentionnel
- **Colonnes Excel fixes** : L = UndTkr, R = Date (formules BDH dépendent de cet ordre)
- `_reorder_columns()` doit toujours être appelé avant `build_excel()`
- `detect_roll_clients` : Roll Client si legs ont même Price **OU** même Notional

## Stack

- Python 3.x, pandas, numpy, openpyxl, streamlit
- Déploiement : Azure App Service (Web App `HugoData`, région Japan West)
  - URL : https://hugodata-bfaqf6buehfce8ad.japanwest-01.azurewebsites.net/
  - Déploiement auto via GitHub Actions (`.github/workflows/main_hugodata.yml`) à chaque push sur `main`
  - Pas de filesystem persistant → pas de CSV local
