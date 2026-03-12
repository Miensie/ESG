# 🌿 ESG Analyzer Pro — Complément Excel

> Automatisation du bilan carbone et de l'analyse ESG directement dans Microsoft Excel.
> Scopes 1, 2 & 3 · GHG Protocol · Base Carbone ADEME 2024

---

## 📁 Structure du projet

```
esg-addin/
├── manifest.xml              ← Manifeste Office Add-in (GUID requis)
├── taskpane.html             ← Interface principale (panneau latéral)
├── commands.html             ← Commandes de ruban (requis par manifest)
├── package.json
├── src/
│   ├── app.js                ← Orchestrateur principal
│   ├── styles/
│   │   └── main.css          ← Design system ESG
│   └── modules/
│       ├── carbonCalc.js     ← Calcul bilan carbone (ADEME/GHG Protocol)
│       ├── excelBridge.js    ← Interaction Excel via Office.js
│       └── reportGenerator.js← Génération rapport HTML/PDF
└── backend/
    └── server.js             ← Serveur Node.js HTTPS de développement
```

---

## 🚀 Installation & Démarrage rapide

### 1. Cloner / déployer les fichiers

```bash
# Sur GitHub Pages (exemple)
git init
git add .
git commit -m "ESG Analyzer Pro v1.0"
git remote add origin https://github.com/miensie/ESG.git
git push -u origin main

# Activer GitHub Pages → Settings → Pages → Branch: main
```

### 2. Mettre à jour le manifest.xml

Remplacez toutes les URLs `https://miensie.github.io/ESG` par votre URL réelle.

### 3. Développement local (Node.js)

```bash
# Installer les dépendances
npm install

# Installer les certificats HTTPS locaux (OBLIGATOIRE pour Excel)
npm run certs
# ou :
npx office-addin-dev-certs install

# Démarrer le serveur de développement
npm start
# → https://localhost:3000
```

---

## 📥 Chargement du complément dans Excel

### Excel Web (Office 365 en ligne)

1. Ouvrir Excel sur **office.com**
2. Ouvrir un classeur
3. **Insérer** → **Compléments** → **Mes compléments**
4. Cliquer **Charger le complément** → **URL du manifeste**
5. Entrer : `https://VOTRE_DOMAINE/manifest.xml`
6. Valider → Le bouton ESG apparaît dans l'onglet **Accueil**

### Excel Desktop (Windows)

```bash
# Option A : Via script npm (recommandé)
npm run sideload-desktop

# Option B : Manuellement (Windows)
# 1. Créer un dossier partagé réseau : ex. \\MonPC\ESGAddin
# 2. Copier manifest.xml dans ce dossier
# 3. Excel → Fichier → Options → Centre de gestion de la confidentialité
#    → Catalogues de compléments approuvés
# 4. Ajouter l'URL : \\MonPC\ESGAddin
# 5. Redémarrer Excel → Insérer → Mes compléments
```

### Excel Mac

```bash
# Copier manifest.xml dans :
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# Redémarrer Excel → Insérer → Mes compléments
```

---

## 🔧 Workflow d'utilisation

```
1. Ouvrir ESG Analyzer Pro dans Excel
2. Onglet "Accueil" → Configurer entreprise + année
3. Cliquer "Initialiser le classeur ESG"
   → Crée automatiquement les feuilles : ESG_Energie, ESG_Transport, ESG_Dechets, ESG_Config
4. Remplir les feuilles de collecte avec vos données
   → Utiliser les clés de type fournies dans l'onglet "Collecte"
5. Onglet "Collecte" → "Lire les données depuis Excel"
6. Onglet "Bilan" → "Calculer le bilan carbone"
   → Affichage KPI Scope 1/2/3, graphiques, détail
7. "Écrire dans Excel" → Feuille ESG_Resultats créée
8. Onglet "Dashboard" → "Créer le dashboard Excel natif"
   → Graphiques Excel natifs dans ESG_Dashboard
9. Onglet "Analyse" → "Lancer l'analyse"
   → Anomalies + recommandations de réduction
10. Onglet "Rapport" → "Ouvrir le rapport"
    → Rapport HTML complet → Ctrl+P pour PDF
```

---

## 📊 Clés de types d'émission disponibles

### Scope 1 — Émissions directes
| Clé | Description | Unité |
|-----|-------------|-------|
| `naturalGas` | Gaz naturel | MWh |
| `diesel` | Gazole | MWh |
| `fuelOil` | Fioul lourd | MWh |
| `lpg` | GPL | MWh |
| `coal` | Charbon | MWh |
| `r410a` | Fuite R-410A | kg |
| `r32` | Fuite R-32 | kg |
| `cementProduction` | Production ciment | t |
| `steelProduction` | Production acier | t |

### Scope 2 — Énergie achetée
| Clé | Description | Unité |
|-----|-------------|-------|
| `electricityFrance` | Électricité France | MWh |
| `electricityEurope` | Électricité Europe | MWh |
| `electricityUSA` | Électricité USA | MWh |
| `districtHeat` | Chaleur réseau | MWh |
| `steamPurchased` | Vapeur achetée | MWh |

### Scope 3 — Émissions indirectes (sélection)
| Clé | Description | Unité |
|-----|-------------|-------|
| `roadFreightHeavy` | Fret routier lourd | t.km |
| `seaFreight` | Fret maritime | t.km |
| `airFreight` | Fret aérien | t.km |
| `businessCar` | Voiture essence (pro) | km |
| `businessAirShort` | Avion court-courrier | km |
| `wasteIncineration` | Incinération déchets | t |
| `wasteRecycling` | Recyclage (évité) | t |
| `steelPurchased` | Acier acheté | t |

---

## 🌐 Déploiement GitHub Pages

```yaml
# .github/workflows/deploy.yml
name: Deploy ESG Add-in
on:
  push:
    branches: [main]
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: peaceiris/actions-gh-pages@v3
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./
```

---

## 📋 Publication sur Microsoft AppSource

1. Créer un compte **Microsoft Partner Center**
2. Soumettre le `manifest.xml` + captures d'écran
3. Passer la validation Microsoft (2-3 semaines)
4. Fixer un prix ou distribuer gratuitement

**Exigences AppSource :**
- HTTPS obligatoire (certificat SSL valide)
- Icônes : 32x32, 64x64, 128x128 px (PNG transparent)
- Support en anglais minimum
- Politique de confidentialité publique

---

## 🏗️ Architecture technique

```
┌─────────────────────────────────────────────────────┐
│                    EXCEL / OFFICE.JS                │
│  ┌─────────────────┐    ┌────────────────────────┐  │
│  │  Feuilles Excel │◄──►│   ExcelBridge.js       │  │
│  │  ESG_Energie    │    │   (Lecture/Écriture)   │  │
│  │  ESG_Transport  │    └───────────┬────────────┘  │
│  │  ESG_Dechets    │                │               │
│  │  ESG_Resultats  │    ┌───────────▼────────────┐  │
│  │  ESG_Dashboard  │    │   app.js (Orchestrateur│  │
│  └─────────────────┘    │   Navigation / Events) │  │
│                          └───────┬──────────┬─────┘  │
│                    ┌─────────────▼─┐   ┌────▼──────┐  │
│                    │ carbonCalc.js │   │reportGen. │  │
│                    │ ADEME/GHG     │   │HTML→PDF   │  │
│                    │ Scope 1/2/3   │   └───────────┘  │
│                    └───────────────┘                  │
└─────────────────────────────────────────────────────┘
                           │ HTTPS
              ┌────────────▼──────────────┐
              │      Node.js Backend      │
              │   server.js (optionnel)   │
              │   API Python (optionnel)  │
              └───────────────────────────┘
```

---

## 🧪 Test rapide sans Excel

Ouvrez `taskpane.html` directement dans un navigateur :
- Cliquez **"Charger données démo"** → données en mémoire
- Allez dans **Bilan** → **Calculer** → KPI et graphiques s'affichent
- **Rapport** → **Ouvrir** → Rapport complet dans un nouvel onglet

> Note : Les fonctions Excel (lecture/écriture) ne fonctionnent qu'à l'intérieur d'Excel.
> En mode navigateur seul, les boutons Excel afficheront une erreur normale.

---

## 📞 Support & Évolutions

- Ajouter Python pour calculs avancés (ADEME API)
- Connecteur Power BI pour dashboards avancés
- Import automatique depuis ERP (SAP, Oracle)
- Module de suivi des objectifs Net Zero

---

*ESG Analyzer Pro v1.0 · Méthodologie GHG Protocol Corporate Standard · Base Carbone ADEME 2024*
