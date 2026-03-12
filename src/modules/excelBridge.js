/**
 * ============================================================
 * MODULE : excelBridge.js
 * Interaction avec Excel via Office.js
 * Lecture, écriture, création de feuilles et graphiques
 * ============================================================
 */

"use strict";

// ─── Noms des feuilles de collecte ───────────────────────────────────────────
const SHEET_NAMES = {
  CONFIG:     "ESG_Config",
  ENERGIE:    "ESG_Energie",
  TRANSPORT:  "ESG_Transport",
  EMISSIONS:  "ESG_EmissionsDirect",
  DECHETS:    "ESG_Dechets",
  FILIALES:   "ESG_Filiales",
  RESULTATS:  "ESG_Resultats",
  DASHBOARD:  "ESG_Dashboard",
};

// ─── Utilitaires bas niveau ───────────────────────────────────────────────────

/**
 * Exécute un bloc Excel.run avec gestion d'erreur
 */
async function xrun(fn) {
  try {
    return await Excel.run(fn);
  } catch (err) {
    console.error("[ExcelBridge] Erreur Excel.run :", err);
    throw err;
  }
}

/**
 * Récupère ou crée une feuille par nom
 */
async function getOrCreateSheet(context, name, visible = true) {
  let sheet;
  try {
    sheet = context.workbook.worksheets.getItem(name);
    await context.sync();
  } catch {
    sheet = context.workbook.worksheets.add(name);
    if (!visible) sheet.visibility = Excel.SheetVisibility.hidden;
    await context.sync();
  }
  return sheet;
}

/**
 * Vérifie si une feuille existe
 */
async function sheetExists(name) {
  return xrun(async (ctx) => {
    try {
      ctx.workbook.worksheets.getItem(name);
      await ctx.sync();
      return true;
    } catch { return false; }
  });
}

// ─── Initialisation du classeur ESG ──────────────────────────────────────────

/**
 * Crée toutes les feuilles ESG avec leur structure de données
 */
async function initWorkbook() {
  return xrun(async (ctx) => {
    // Feuille Configuration
    const sheetConfig = await getOrCreateSheet(ctx, SHEET_NAMES.CONFIG);
    sheetConfig.getRange("A1:B1").values = [["Paramètre", "Valeur"]];
    sheetConfig.getRange("A2:B8").values = [
      ["Entreprise",       "Mon Entreprise SAS"],
      ["Secteur",          "Industrie manufacturière"],
      ["Année de référence", new Date().getFullYear()],
      ["Chiffre d'affaires (€)", 45000000],
      ["Production (tonnes)", 12000],
      ["Pays / Région",    "France"],
      ["Responsable ESG",  ""],
    ];
    styleHeaderRow(ctx, sheetConfig, "A1:B1");

    // Feuille Énergie
    const sheetEnergie = await getOrCreateSheet(ctx, SHEET_NAMES.ENERGIE);
    sheetEnergie.getRange("A1:F1").values = [[
      "Source énergie", "Scope", "Type carburant/énergie",
      "Quantité consommée", "Unité", "Commentaire"
    ]];
    sheetEnergie.getRange("A2:F7").values = [
      ["Chaudière principale", "Scope 1", "naturalGas",      8500, "MWh", "Site production"],
      ["Groupe électrogène",   "Scope 1", "diesel",          1200, "MWh", "Secours"],
      ["Climatisation",        "Scope 1", "r410a",             25, "kg",  "Fuite annuelle estimée"],
      ["Production - Site A",  "Scope 2", "electricityFrance",6800, "MWh", "Compteur EDF"],
      ["Bureaux - Site B",     "Scope 2", "electricityFrance",1200, "MWh", "Télérelève"],
      ["Réseau chaleur",       "Scope 2", "districtHeat",     850, "MWh", "Contrat Dalkia"],
    ];
    styleHeaderRow(ctx, sheetEnergie, "A1:F1");
    sheetEnergie.getRange("A1:F1").format.columnWidth = 160;

    // Feuille Transport
    const sheetTransport = await getOrCreateSheet(ctx, SHEET_NAMES.TRANSPORT);
    sheetTransport.getRange("A1:G1").values = [[
      "Description", "Scope", "Mode transport", "Type clé",
      "Quantité", "Unité", "Commentaire"
    ]];
    sheetTransport.getRange("A2:G7").values = [
      ["Livraisons clients",   "Scope 3", "Routier",   "roadFreightHeavy",  420000, "t.km", "Via prestataires"],
      ["Import matières",      "Scope 3", "Maritime",  "seaFreight",        180000, "t.km", "Asie - Europe"],
      ["Déplacements pro avion","Scope 3", "Aérien",   "businessAirShort",   85000, "km",   "Court-courrier"],
      ["Voitures salariés",    "Scope 3", "Routier",   "businessCar",       124000, "km",   "Notes de frais"],
      ["Train déplacements",   "Scope 3", "Ferroviaire","trainTravel",        18000, "km",   "SNCF Voyageurs"],
      ["Sous-traitants livr.", "Scope 3", "Routier",   "roadFreightLight",   35000, "t.km", "Messagerie"],
    ];
    styleHeaderRow(ctx, sheetTransport, "A1:G1");

    // Feuille Déchets
    const sheetDechets = await getOrCreateSheet(ctx, SHEET_NAMES.DECHETS);
    sheetDechets.getRange("A1:F1").values = [[
      "Type déchet", "Scope", "Mode traitement", "Type clé",
      "Quantité (t)", "Commentaire"
    ]];
    sheetDechets.getRange("A2:F5").values = [
      ["Déchets industriels banals", "Scope 3", "Incinération", "wasteIncineration", 185, "DASRI exclus"],
      ["Métaux ferreux",             "Scope 3", "Recyclage",    "wasteRecycling",     95, "Revalorisation"],
      ["Boues STEP",                 "Scope 3", "Eaux usées",   "wasteWater",         42, "Traitement externe"],
      ["Gravats / inertes",          "Scope 3", "Enfouissement","wasteLandfill",       28, "ISDI"],
    ];
    styleHeaderRow(ctx, sheetDechets, "A1:F1");

    await ctx.sync();
    console.log("[ExcelBridge] Classeur ESG initialisé ✓");
    return true;
  });
}

/**
 * Style rapide pour une ligne d'en-tête
 */
function styleHeaderRow(ctx, sheet, rangeAddr) {
  const hdr = sheet.getRange(rangeAddr);
  hdr.format.fill.color        = "#1C2B3A";
  hdr.format.font.color        = "#2ECC8E";
  hdr.format.font.bold         = true;
  hdr.format.font.size         = 11;
  hdr.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = Excel.BorderLineStyle.medium;
  hdr.format.borders.getItem(Excel.BorderIndex.edgeBottom).color = "#2ECC8E";
}

// ─── Lecture des données ESG depuis Excel ─────────────────────────────────────

/**
 * Lit toutes les données ESG du classeur et les retourne structurées
 */
async function readESGData() {
  return xrun(async (ctx) => {
    const data = { scope1: [], scope2: [], scope3: [], annee: null, entreprise: null };

    // Lire config
    try {
      const cfgSheet = ctx.workbook.worksheets.getItem(SHEET_NAMES.CONFIG);
      const cfgRange = cfgSheet.getRange("B2:B8");
      cfgRange.load("values");
      await ctx.sync();
      const v = cfgRange.values;
      data.entreprise      = v[0][0] || "Mon Entreprise";
      data.secteur         = v[1][0] || "Industrie";
      data.annee           = parseInt(v[2][0]) || new Date().getFullYear();
      data.chiffreAffaires = parseFloat(v[3][0]) || 0;
      data.productionTonnes= parseFloat(v[4][0]) || 0;
    } catch (e) {
      console.warn("[ExcelBridge] Feuille Config manquante, valeurs par défaut.");
    }

    // Lire Énergie (Scope 1 & 2)
    await readSheetData(ctx, SHEET_NAMES.ENERGIE, data);

    // Lire Transport (Scope 3)
    await readSheetData(ctx, SHEET_NAMES.TRANSPORT, data);

    // Lire Déchets (Scope 3)
    await readSheetData(ctx, SHEET_NAMES.DECHETS, data);

    return data;
  });
}

/**
 * Lecture générique d'une feuille de collecte
 * Colonnes attendues : col[1]=Scope, col[2]=typeClé, col[3]=quantité, col[0]=source
 */
async function readSheetData(ctx, sheetName, data) {
  try {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRange();
    used.load("values");
    await ctx.sync();

    const rows = used.values.slice(1); // Skip header
    rows.forEach(row => {
      const source   = String(row[0] || "").trim();
      const scopeStr = String(row[1] || "").trim().toLowerCase();
      const typeKey  = String(row[2] || row[3] || "").trim(); // flexible selon feuille
      const quantity = parseFloat(row[3] || row[4] || 0);

      if (!typeKey || isNaN(quantity) || quantity <= 0) return;

      const entry = { type: typeKey, quantity, source };

      if (scopeStr.includes("scope 1") || scopeStr === "scope1") {
        data.scope1.push(entry);
      } else if (scopeStr.includes("scope 2") || scopeStr === "scope2") {
        data.scope2.push(entry);
      } else if (scopeStr.includes("scope 3") || scopeStr === "scope3") {
        data.scope3.push(entry);
      }
    });
  } catch (e) {
    console.warn(`[ExcelBridge] Feuille "${sheetName}" introuvable ou vide.`);
  }
}

// ─── Écriture des résultats dans Excel ────────────────────────────────────────

/**
 * Écrit le bilan carbone dans la feuille ESG_Resultats
 */
async function writeResultats(bilan) {
  return xrun(async (ctx) => {
    const sheet = await getOrCreateSheet(ctx, SHEET_NAMES.RESULTATS);
    sheet.getUsedRange().clear();

    // Titre
    const title = sheet.getRange("A1");
    title.values = [[`BILAN CARBONE ${bilan.annee} — ${bilan.entreprise}`]];
    title.format.font.bold = true;
    title.format.font.size = 14;
    title.format.font.color = "#2ECC8E";

    // KPI principaux
    sheet.getRange("A3:B3").values = [["Indicateur", "Valeur"]];
    styleHeaderRow(ctx, sheet, "A3:B3");

    const kpis = [
      ["Total Scope 1 (tCO2eq)",    bilan.scope1.total],
      ["Total Scope 2 (tCO2eq)",    bilan.scope2.total],
      ["Total Scope 3 (tCO2eq)",    bilan.scope3.total],
      ["TOTAL BILAN CARBONE",        bilan.grandTotal],
      ["Part Scope 1 (%)",           parseFloat(bilan.scope1.pct)],
      ["Part Scope 2 (%)",           parseFloat(bilan.scope2.pct)],
      ["Part Scope 3 (%)",           parseFloat(bilan.scope3.pct)],
      ["Intensité carbone",          bilan.intensite],
      ["Unité intensité",            bilan.intensiteUnit || "N/A"],
    ];

    sheet.getRange(`A4:B${3 + kpis.length}`).values = kpis;
    sheet.getRange(`A7:B7`).format.font.bold = true;

    // Détail Scope 1
    let row = 3 + kpis.length + 2;
    sheet.getRange(`A${row}`).values = [["DÉTAIL SCOPE 1"]];
    sheet.getRange(`A${row}`).format.font.bold = true;
    sheet.getRange(`A${row}`).format.font.color = "#E05252";
    row++;
    sheet.getRange(`A${row}:D${row}`).values = [["Source", "Type", "Quantité", "tCO2eq"]];
    styleHeaderRow(ctx, sheet, `A${row}:D${row}`);
    row++;
    bilan.scope1.lines.forEach(l => {
      sheet.getRange(`A${row}:D${row}`).values = [[l.source, l.label, l.quantity, l.tCO2eq]];
      row++;
    });

    // Détail Scope 2
    row++;
    sheet.getRange(`A${row}`).values = [["DÉTAIL SCOPE 2"]];
    sheet.getRange(`A${row}`).format.font.bold = true;
    sheet.getRange(`A${row}`).format.font.color = "#F5A623";
    row++;
    sheet.getRange(`A${row}:D${row}`).values = [["Source", "Type", "Quantité", "tCO2eq"]];
    styleHeaderRow(ctx, sheet, `A${row}:D${row}`);
    row++;
    bilan.scope2.lines.forEach(l => {
      sheet.getRange(`A${row}:D${row}`).values = [[l.source, l.label, l.quantity, l.tCO2eq]];
      row++;
    });

    // Détail Scope 3
    row++;
    sheet.getRange(`A${row}`).values = [["DÉTAIL SCOPE 3"]];
    sheet.getRange(`A${row}`).format.font.bold = true;
    sheet.getRange(`A${row}`).format.font.color = "#2ECC8E";
    row++;
    sheet.getRange(`A${row}:D${row}`).values = [["Source", "Type", "Quantité", "tCO2eq"]];
    styleHeaderRow(ctx, sheet, `A${row}:D${row}`);
    row++;
    bilan.scope3.lines.forEach(l => {
      sheet.getRange(`A${row}:D${row}`).values = [[l.source, l.label, l.quantity, l.tCO2eq]];
      row++;
    });

    // Autofit colonnes
    sheet.getRange("A:D").format.autofitColumns();

    await ctx.sync();
    console.log("[ExcelBridge] Résultats écrits ✓");
    return true;
  });
}

/**
 * Crée un dashboard interactif avec graphiques Excel natifs
 */
async function createDashboard(bilan) {
  return xrun(async (ctx) => {
    const sheet = await getOrCreateSheet(ctx, SHEET_NAMES.DASHBOARD);
    sheet.getUsedRange().clear();

    // Données pour graphiques
    sheet.getRange("A1:B4").values = [
      ["Scope", "tCO2eq"],
      ["Scope 1", bilan.scope1.total],
      ["Scope 2", bilan.scope2.total],
      ["Scope 3", bilan.scope3.total],
    ];

    // Graphique camembert Répartition Scopes
    const pieChart = sheet.charts.add(
      Excel.ChartType.pie,
      sheet.getRange("A1:B4"),
      Excel.ChartSeriesBy.auto
    );
    pieChart.title.text = "Répartition par Scope";
    pieChart.setPosition(sheet.getRange("D1"), sheet.getRange("L16"));
    pieChart.dataLabels.showPercentage = true;
    pieChart.dataLabels.showValue      = true;
    pieChart.legend.position = Excel.ChartLegendPosition.bottom;

    // Données détail pour histogramme
    const detailData = [["Source", "tCO2eq"]];
    [...bilan.scope1.lines, ...bilan.scope2.lines, ...bilan.scope3.lines].forEach(l => {
      detailData.push([l.source.substring(0, 20), l.tCO2eq]);
    });
    const detailStartRow = 6;
    sheet.getRange(`A${detailStartRow}:B${detailStartRow + detailData.length - 1}`).values = detailData;

    // Graphique barres Détail
    const barChart = sheet.charts.add(
      Excel.ChartType.barClustered,
      sheet.getRange(`A${detailStartRow}:B${detailStartRow + detailData.length - 1}`),
      Excel.ChartSeriesBy.auto
    );
    barChart.title.text = "Émissions par source (tCO2eq)";
    barChart.setPosition(
      sheet.getRange("D17"),
      sheet.getRange(`L${17 + detailData.length + 4}`)
    );

    // KPI zone texte
    sheet.getRange("A1").values = [[`Bilan Carbone ${bilan.annee}`]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.font.color = "#1A6B4A";

    await ctx.sync();
    console.log("[ExcelBridge] Dashboard créé ✓");
    return true;
  });
}

/**
 * Sélectionne et active une feuille Excel
 */
async function activateSheet(name) {
  return xrun(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(name);
    sheet.activate();
    await ctx.sync();
  });
}

/**
 * Récupère le nom du classeur actif
 */
async function getWorkbookName() {
  return xrun(async (ctx) => {
    const wb = ctx.workbook;
    wb.load("name");
    await ctx.sync();
    return wb.name;
  });
}

// Export global
window.ExcelBridge = {
  SHEET_NAMES,
  initWorkbook,
  readESGData,
  writeResultats,
  createDashboard,
  activateSheet,
  getWorkbookName,
  sheetExists,
};
