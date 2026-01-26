# Guide d’utilisation du rapport Power BI

Power BI est un logiciel de visualisation de données (équivalent à Excel pour l’analyse interactive) permettant :
- d’afficher des graphiques dynamiques,
- d’appliquer des filtres,
- de naviguer entre plusieurs pages d’analyse.
**Aucune programmation n’est nécessaire pour consulter le rapport.**

# Ouvrir le projet
1. Installation (obligatoire)
Power BI fonctionne uniquement sous **Windows**.
Télécharger Power BI Desktop (gratuit)
https://powerbi.microsoft.com/desktop/

2. Télécharger le fichier du projet `Analyse_TBM.pbix`
3. Double-cliquer sur le fichier  
   → Il s’ouvre automatiquement dans **Power BI Desktop**

! Le chargement peut prendre quelques secondes.


## Important : ne PAS actualiser les données

**Ne pas cliquer sur le bouton « Actualiser »**

Raison :
- les fichiers de données sources (CSV / Excel) **ne sont pas inclus** dans le dépôt Git,
- les données sont déjà intégrées dans le rapport,
- une actualisation demanderait les fichiers originaux absents.

Le rapport est **prêt à être consulté tel quel**.


## Navigation dans le rapport

### Changer de page
- Les pages sont visibles **en bas de l’écran**
- Cliquer sur :
  - **Page 1 – Analyse globale des incidents**
  - **Page 2 – Géolocalisation des incidents**
  - **Page 3 – Conditions météorologiques**

## Interaction avec les graphiques

1. Filtres automatiques
- Cliquer sur un élément d’un graphique (ex. une ligne de tram)
- Tous les autres graphiques de la page se mettent à jour automatiquement

2. Annuler un filtre
- Cliquer à nouveau sur l’élément sélectionné
- Ou cliquer dans une zone vide du rapport