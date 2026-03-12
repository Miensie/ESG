# ESG Analyzer Pro — Office Add-in Excel

Complément Excel pour l'analyse ESG automatisée : bilan carbone (GHG Protocol), indicateurs KPI, détection d'anomalies, tableaux de bord et rapports.

## Architecture du projet

```
esg-addin/
├── manifest.xml               # Configuration Office Add-in
├── package.json               # Dépendances Node.js
├── webpack.config.js          # Build & dev server HTTPS
├── README.md
└── src/
    ├── taskpane.html          # Interface utilisateur (5 onglets)
    ├── taskpane.css           # Styles (thème dark industriel)
    ├── taskpane.js            # Logique principale + Office.js
    ├── esg-calculator.js      # Moteur calcul bilan carbone
    └── commands.html          # Stub commandes ruban
```

## Fonctionnalités

### 📥 Onglet Collecte
- Configuration de 5 sources de données (Énergie, Transport, Émissions, Déchets, Filiales)
- Lecture des plages Excel configurées
- Création automatique d'une feuille de structuration `ESG_Data_YYYY`

### 🌍 Onglet Carbone
- Facteurs d'émission ADEME Base Carbone 2024 (modifiables)
- Saisie manuelle OU calcul depuis données collectées
- Calcul Scope 1, 2, 3 + total + intensité carbone
- Export vers feuille Excel avec graphique camembert

### 📊 Onglet Dashboard
- Génération d'un tableau de bord ESG complet
- KPI cards colorés, graphiques (répartition scopes, évolution)
- Données historiques configurables

### 🔍 Onglet Analyse
- Détection d'anomalies statistiques (seuil configurable)
- Recommandations de réduction (priorisées par économie CO₂e)
- Export analyse vers Excel

### 📄 Onglet Rapport
- Rapport ESG structuré (4 sections : bilan, KPI, anomalies, plan d'action)
- Formaté pour impression (fond sombre, typographie claire)
- Contenu configurable (cases à cocher)

## Calcul bilan carbone

```
Scope 1 = (gaz × FE_gaz) + (fioul × FE_fioul) + (procédés × FE_process)
Scope 2 = (élec × FE_élec) + (chaleur × FE_chaleur)
Scope 3 = (transport × FE_transport) + (déchets × FE_déchets) + (achats × FE_achats)
Total = Scope 1 + Scope 2 + Scope 3
Intensité = Total / Valeur_activité
```

Tous les facteurs sont en kgCO₂e/unité, convertis en tCO₂e (÷1000).

## Installation

### Prérequis
- Node.js ≥ 18
- Excel Desktop (Windows/Mac) ou Excel Online

### 1. Installer les dépendances
```bash
npm install
```

### 2. Générer les certificats HTTPS
```bash
npx office-addin-dev-certs install
```

### 3. Lancer en mode développement
```bash
npm run start
# ou séparément :
npm run dev          # webpack dev server sur https://localhost:3000
```

### 4. Charger le complément dans Excel

**Excel Desktop :**
1. Excel → Fichier → Options → Centre de gestion de la confidentialité
2. Catalogues de compléments approuvés → Ajouter `\\votre-pc\AddIns` (partage réseau)
3. OU : Fichier → Options → Compléments → Compléments COM → Parcourir `manifest.xml`

**Excel Online (Office 365) :**
1. Insertion → Compléments → Mes compléments → Charger un complément
2. Sélectionner `manifest.xml`

**Mode développement (le plus simple) :**
```bash
npm run start
# Ouvre automatiquement Excel avec le complément chargé
```

## Déploiement

### GitHub Pages
```bash
# Dans manifest.xml, remplacer YOUR-DOMAIN par votre domaine
npm run build
npm run deploy
```

### Serveur intranet (IIS ou nginx)
```bash
npm run build
# Copier le contenu de /dist/ sur votre serveur web HTTPS
# Modifier manifest.xml avec l'URL de votre serveur
```

## Facteurs d'émission ADEME 2024

| Source | Facteur | Unité |
|--------|---------|-------|
| Gaz naturel | 0.205 kgCO₂e | MWh |
| Fioul lourd | 2.96 kgCO₂e | L |
| Électricité (RFE) | 0.0571 kgCO₂e | MWh |
| Transport routier | 0.0621 kgCO₂e | t·km |
| Déchets enfouis | 449 kgCO₂e | t |

Source : [Base Carbone ADEME](https://base-empreinte.ademe.fr/)

## Standards de référence
- **GHG Protocol** — Corporate Accounting and Reporting Standard
- **ISO 14064-1** — Quantification des GES au niveau des organisations
- **CSRD** — Corporate Sustainability Reporting Directive (UE)
- **TCFD** — Task Force on Climate-related Financial Disclosures

## Licence
MIT — Libre d'utilisation et de modification
