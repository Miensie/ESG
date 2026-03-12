/**
 * ============================================================
 * ESG ANALYZER PRO — Taskpane Principal
 * Office.js interactions, UI logic, Excel export
 * ============================================================
 */

"use strict";

// ============================================================
// ÉTAT GLOBAL
// ============================================================
const AppState = {
  sources: {
    energie:   { sheet: "", range: "", headerRow: 1, configured: false },
    transport: { sheet: "", range: "", headerRow: 1, configured: false },
    emissions: { sheet: "", range: "", headerRow: 1, configured: false },
    dechets:   { sheet: "", range: "", headerRow: 1, configured: false },
    filiales:  { sheet: "", range: "", headerRow: 1, configured: false },
  },
  currentSource: null,
  collectedData: {},
  carbonResults: null,
  anomalies: [],
  recommendations: [],
  reportFormat: "excel",
  calculator: null
};

// ============================================================
// INIT OFFICE.JS
// ============================================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    AppState.calculator = new ESGCalculator();
    initNavTabs();
    loadSheetList();
    console.log("ESG Analyzer Pro — Prêt");
  }
});

// ============================================================
// NAVIGATION
// ============================================================
function initNavTabs() {
  document.querySelectorAll(".nav-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      const panelId = "panel-" + tab.dataset.panel;
      document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(panelId).classList.add("active");
    });
  });
}

// ============================================================
// CHARGEMENT DES FEUILLES EXCEL
// ============================================================
async function loadSheetList() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const select = document.getElementById("config-sheet");
      select.innerHTML = '<option value="">Sélectionner une feuille...</option>';
      sheets.items.forEach(sh => {
        const opt = document.createElement("option");
        opt.value = opt.textContent = sh.name;
        select.appendChild(opt);
      });
    });
  } catch (e) {
    console.warn("Impossible de charger les feuilles :", e);
  }
}

// ============================================================
// CONFIGURATION DES SOURCES
// ============================================================
function configureSource(sourceKey) {
  AppState.currentSource = sourceKey;
  const labels = {
    energie:   "⚡ Consommation Énergie",
    transport: "🚛 Transport & Logistique",
    emissions: "🏭 Émissions Directes",
    dechets:   "♻️ Déchets & Recyclage",
    filiales:  "🏢 Activités Filiales"
  };

  document.getElementById("config-title").textContent = labels[sourceKey];

  const existing = AppState.sources[sourceKey];
  if (existing.configured) {
    document.getElementById("config-sheet").value = existing.sheet;
    document.getElementById("config-range").value = existing.range;
    document.getElementById("config-header").value = existing.headerRow;
  }

  document.getElementById("source-config").style.display = "block";
  loadSheetList();
}

function cancelSourceConfig() {
  document.getElementById("source-config").style.display = "none";
  AppState.currentSource = null;
}

function saveSourceConfig() {
  const key   = AppState.currentSource;
  const sheet = document.getElementById("config-sheet").value;
  const range = document.getElementById("config-range").value.trim().toUpperCase();
  const hdr   = parseInt(document.getElementById("config-header").value) || 1;

  if (!sheet || !range) {
    showToast("Veuillez renseigner la feuille et la plage", "warn");
    return;
  }

  AppState.sources[key] = { sheet, range, headerRow: hdr, configured: true };

  document.getElementById(`range-${key}`).textContent = `${sheet} › ${range}`;
  const statusEl = document.getElementById(`status-${key}`);
  statusEl.classList.add("ok");

  document.querySelector(`[data-source="${key}"]`).classList.add("configured");
  document.getElementById("source-config").style.display = "none";

  showToast(`Source "${sheet}" configurée`, "success");
  AppState.currentSource = null;
}

async function previewSource() {
  const sheet = document.getElementById("config-sheet").value;
  const range = document.getElementById("config-range").value.trim().toUpperCase();
  if (!sheet || !range) { showToast("Sélectionnez feuille et plage", "warn"); return; }

  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const rng = ws.getRange(range);
      rng.load("values");
      await ctx.sync();

      const preview = rng.values.slice(0, 3).map(r => r.join(" | ")).join("\n");
      showToast(`Aperçu (3 premières lignes) :\n${preview}`, "info");
    });
  } catch (e) {
    showToast("Erreur lecture plage : " + e.message, "error");
  }
}

// ============================================================
// COLLECTE DES DONNÉES
// ============================================================
async function collectAllData() {
  const configuredSources = Object.entries(AppState.sources).filter(([, v]) => v.configured);

  if (configuredSources.length === 0) {
    showToast("Configurez au moins une source de données", "warn");
    return;
  }

  setStatus("collect-status", "running", "Collecte en cours...");

  try {
    await Excel.run(async (ctx) => {
      for (const [key, src] of configuredSources) {
        const ws  = ctx.workbook.worksheets.getItem(src.sheet);
        const rng = ws.getRange(src.range);
        rng.load("values");
        await ctx.sync();

        // Supprimer la ligne d'en-tête
        const data = rng.values.slice(src.headerRow - 1 + 1);
        AppState.collectedData[key] = data;
      }
    });

    const total = Object.values(AppState.collectedData).reduce((s, d) => s + d.length, 0);
    setStatus("collect-status", "ok", `${total} lignes collectées depuis ${configuredSources.length} source(s)`);
    showToast(`Données collectées : ${total} lignes`, "success");

    // Créer feuille de structuration
    await createStructuredSheet();

  } catch (e) {
    setStatus("collect-status", "error", "Erreur : " + e.message);
    showToast("Erreur collecte : " + e.message, "error");
  }
}

async function createStructuredSheet() {
  const sheetName = `ESG_Data_${new Date().getFullYear()}`;

  await Excel.run(async (ctx) => {
    // Supprimer si existe
    try {
      const old = ctx.workbook.worksheets.getItem(sheetName);
      old.delete();
      await ctx.sync();
    } catch (_) {}

    const sheet = ctx.workbook.worksheets.add(sheetName);
    sheet.activate();

    // En-têtes
    const headers = ["Source", "Catégorie", "Quantité", "Unité", "Scope", "Période", "Commentaire"];
    const hdrRange = sheet.getRange("A1:G1");
    hdrRange.values = [headers];
    hdrRange.format.fill.color = "#0f1a14";
    hdrRange.format.font.color = "#00d97e";
    hdrRange.format.font.bold = true;
    hdrRange.format.font.size = 10;

    // Données structurées
    let row = 2;
    const categoryLabels = {
      energie:   "Énergie",
      transport: "Transport",
      emissions: "Émissions",
      dechets:   "Déchets",
      filiales:  "Filiales"
    };
    const scopeMap = {
      energie:   "Scope 2",
      transport: "Scope 3",
      emissions: "Scope 1",
      dechets:   "Scope 3",
      filiales:  "Scope 1/3"
    };

    const allRows = [];
    for (const [key, data] of Object.entries(AppState.collectedData)) {
      for (const r of data) {
        if (!r || r.every(c => c === null || c === "")) continue;
        allRows.push([
          String(r[0] || ""),
          categoryLabels[key] || key,
          parseFloat(r[1]) || 0,
          String(r[2] || ""),
          scopeMap[key] || "",
          String(r[3] || new Date().getFullYear()),
          String(r[4] || "")
        ]);
      }
    }

    if (allRows.length > 0) {
      const dataRange = sheet.getRange(`A2:G${allRows.length + 1}`);
      dataRange.values = allRows;
      dataRange.format.font.size = 10;

      // Alternance couleurs
      for (let i = 0; i < allRows.length; i++) {
        const rowRange = sheet.getRange(`A${i+2}:G${i+2}`);
        rowRange.format.fill.color = i % 2 === 0 ? "#0f1a14" : "#141f18";
        rowRange.format.font.color = "#e8f5ee";
      }
    }

    // Ajuster largeurs
    const cols = ["A","B","C","D","E","F","G"];
    const widths = [180, 120, 90, 80, 90, 80, 150];
    cols.forEach((c, i) => {
      sheet.getRange(`${c}:${c}`).format.columnWidth = widths[i];
    });

    await ctx.sync();
  });

  showToast(`Feuille "${sheetName}" créée`, "success");
}

// ============================================================
// CALCUL BILAN CARBONE
// ============================================================
function calculateCarbon() {
  const customFactors = {
    fe_gaz:       parseFloat(document.getElementById("fe-gaz").value)       || undefined,
    fe_fioul:     parseFloat(document.getElementById("fe-fioul").value)     || undefined,
    fe_process:   parseFloat(document.getElementById("fe-process").value)   || undefined,
    fe_elec:      parseFloat(document.getElementById("fe-elec").value)      || undefined,
    fe_chaleur:   parseFloat(document.getElementById("fe-chaleur").value)   || undefined,
    fe_transport: parseFloat(document.getElementById("fe-transport").value) || undefined,
    fe_dechets:   parseFloat(document.getElementById("fe-dechets").value)   || undefined,
    fe_achats:    parseFloat(document.getElementById("fe-achats").value)    || undefined,
  };

  const inputs = {
    qty_gaz:       parseFloat(document.getElementById("qty-gaz").value)       || 0,
    qty_fioul:     parseFloat(document.getElementById("qty-fioul").value)     || 0,
    qty_process:   parseFloat(document.getElementById("qty-process").value)   || 0,
    qty_elec:      parseFloat(document.getElementById("qty-elec").value)      || 0,
    qty_chaleur:   parseFloat(document.getElementById("qty-chaleur").value)   || 0,
    qty_transport: parseFloat(document.getElementById("qty-transport").value) || 0,
    qty_dechets:   parseFloat(document.getElementById("qty-dechets").value)   || 0,
    qty_achats:    parseFloat(document.getElementById("qty-achats").value)    || 0,
    year:          parseInt(document.getElementById("ref-year").value)        || 2024,
    activity_unit: document.getElementById("activity-unit").value             || "unité",
    activity_value: parseFloat(document.getElementById("activity-value").value) || 0,
  };

  const total = Object.entries(inputs).filter(([k]) => k.startsWith("qty")).reduce((s, [,v]) => s + v, 0);
  if (total === 0) {
    showToast("Saisissez au moins une quantité, ou collectez vos données", "warn");
    return;
  }

  const results = AppState.calculator.calculateFromManual(inputs, customFactors);
  AppState.carbonResults = results;

  displayCarbonResults(results);
  showToast("Bilan carbone calculé !", "success");
}

function displayCarbonResults(r) {
  document.getElementById("carbon-results").style.display = "block";

  document.getElementById("res-s1").textContent = fmtNum(r.scope1);
  document.getElementById("res-s2").textContent = fmtNum(r.scope2);
  document.getElementById("res-s3").textContent = fmtNum(r.scope3);
  document.getElementById("res-total").textContent = fmtNum(r.total) + " t";

  const pct1 = r.total > 0 ? Math.round(r.scope1 / r.total * 100) : 0;
  const pct3 = r.total > 0 ? Math.round(r.scope3 / r.total * 100) : 0;
  document.getElementById("res-pct1").textContent = pct1 + "%";
  document.getElementById("res-pct3").textContent = pct3 + "%";

  if (r.intensity) {
    document.getElementById("res-intensity").textContent = fmtNum(r.intensity);
    document.getElementById("res-intensity-unit").textContent = `tCO₂e/${r.metadata.activity_unit}`;
  }
}

async function exportCarbonToExcel() {
  if (!AppState.carbonResults) { showToast("Calculez d'abord le bilan carbone", "warn"); return; }
  const r = AppState.carbonResults;
  const date = formatDateSuffix();
  const sheetName = `BilanCarbone_${date}`;

  await Excel.run(async (ctx) => {
    try { ctx.workbook.worksheets.getItem(sheetName).delete(); await ctx.sync(); } catch(_) {}

    const sheet = ctx.workbook.worksheets.add(sheetName);
    sheet.activate();

    // Titre
    sheet.getRange("A1:F1").merge();
    sheet.getRange("A1").values = [["BILAN CARBONE — " + (r.metadata.year || 2024)]];
    styleTitle(sheet.getRange("A1:F1"), "#00d97e");

    // Sous-titre
    sheet.getRange("A2").values = [[`Calculé le ${new Date().toLocaleDateString("fr-FR")} — Total : ${fmtNum(r.total)} tCO₂e`]];
    sheet.getRange("A2").format.font.color = "#8aab97";
    sheet.getRange("A2").format.font.size = 9;

    // Résumé Scopes
    const summaryHeaders = [["Scope", "Émissions (tCO₂e)", "Part (%)","","",""]];
    sheet.getRange("A4:F4").values = summaryHeaders;
    styleHeaders(sheet.getRange("A4:F4"), "#141f18", "#8aab97");

    const summaryData = AppState.calculator.getScopesSummary();
    const summaryRows = summaryData.map(row => [row[0], row[1], Math.round(row[2] * 10)/10, "", "", ""]);
    sheet.getRange(`A5:F${4 + summaryRows.length}`).values = summaryRows;

    // Couleurs scopes
    const scopeColors = ["#00d97e", "#00b8d4", "#7c5cbf", "#e8f5ee"];
    summaryRows.forEach((_, i) => {
      sheet.getRange(`A${5+i}:F${5+i}`).format.font.color = scopeColors[i] || "#e8f5ee";
      sheet.getRange(`A${5+i}:F${5+i}`).format.fill.color = "#0f1a14";
    });

    // Détail des émissions
    const detailStart = 5 + summaryRows.length + 2;
    sheet.getRange(`A${detailStart}:F${detailStart}`).values = [["Source d'émission", "Scope", "Quantité", "Unité", "Facteur (kgCO₂e)", "Émissions (tCO₂e)"]];
    styleHeaders(sheet.getRange(`A${detailStart}:F${detailStart}`), "#141f18", "#8aab97");

    const detailRows = AppState.calculator.toExcelRows();
    if (detailRows.length > 0) {
      sheet.getRange(`A${detailStart+1}:F${detailStart + detailRows.length}`).values = detailRows;
      for (let i = 0; i < detailRows.length; i++) {
        const rng = sheet.getRange(`A${detailStart+1+i}:F${detailStart+1+i}`);
        rng.format.fill.color = i % 2 === 0 ? "#0f1a14" : "#141f18";
        rng.format.font.color = "#e8f5ee";
        rng.format.font.size = 10;
      }
    }

    // Intensité carbone
    if (r.intensity) {
      const intensityRow = detailStart + detailRows.length + 2;
      sheet.getRange(`A${intensityRow}`).values = [[`Intensité carbone : ${fmtNum(r.intensity)} tCO₂e/${r.metadata.activity_unit}`]];
      sheet.getRange(`A${intensityRow}`).format.font.color = "#ffb800";
      sheet.getRange(`A${intensityRow}`).format.font.bold = true;
    }

    // Graphique camembert Scopes
    const chartData = sheet.getRange("A5:B7");
    const chart = sheet.charts.add(Excel.ChartType.pie, chartData);
    chart.title.text = "Répartition par Scope";
    chart.setPosition("D4", "J14");
    chart.format.fill.setSolidColor("#0a0f0d");
    chart.title.format.font.color = "#e8f5ee";
    chart.title.format.font.size = 12;

    // Largeurs colonnes
    [["A",180],["B",110],["C",90],["D",80],["E",120],["F",100]].forEach(([c,w]) => {
      sheet.getRange(`${c}:${c}`).format.columnWidth = w;
    });

    // Fond global
    sheet.getRange("A1:J50").format.fill.color = "#0a0f0d";

    await ctx.sync();
  });

  showToast(`Bilan exporté → feuille "${sheetName}"`, "success");
}

// ============================================================
// DASHBOARD
// ============================================================
async function generateDashboardFromPanel() {
  if (!AppState.carbonResults) {
    showToast("Calculez d'abord le bilan carbone (onglet Carbone)", "warn");
    return;
  }

  setStatus("dash-status", "running", "Génération du dashboard...");

  const sheetName = document.getElementById("dash-sheet-name").value || "ESG_Dashboard";
  const r = AppState.carbonResults;
  const hist = getHistoricalData();

  try {
    await Excel.run(async (ctx) => {
      try { ctx.workbook.worksheets.getItem(sheetName).delete(); await ctx.sync(); } catch(_) {}

      const sheet = ctx.workbook.worksheets.add(sheetName);
      sheet.tabColor = "#00d97e";
      sheet.activate();

      // Fond global
      sheet.getRange("A1:Z100").format.fill.color = "#0a0f0d";

      // TITRE PRINCIPAL
      sheet.getRange("A1:M1").merge();
      sheet.getRange("A1").values = [["📊  TABLEAU DE BORD ESG — " + (r.metadata.year || 2024)]];
      styleTitle(sheet.getRange("A1:M1"), "#00d97e", 16);

      sheet.getRange("A2:M2").merge();
      sheet.getRange("A2").values = [[`Généré le ${new Date().toLocaleDateString("fr-FR")} | ${r.metadata.activity_unit} : ${fmtNum(r.metadata.activity_value)}`]];
      sheet.getRange("A2").format.font.color = "#4a7060";
      sheet.getRange("A2").format.font.size = 9;

      // KPI CARDS (ligne 4)
      const kpiData = [
        { label: "TOTAL CO₂e", value: fmtNum(r.total) + " t",   color: "#00d97e" },
        { label: "SCOPE 1",     value: fmtNum(r.scope1) + " t", color: "#00d97e" },
        { label: "SCOPE 2",     value: fmtNum(r.scope2) + " t", color: "#00b8d4" },
        { label: "SCOPE 3",     value: fmtNum(r.scope3) + " t", color: "#7c5cbf" },
      ];

      if (r.intensity) kpiData.push({ label: "INTENSITÉ", value: fmtNum(r.intensity) + " t/u", color: "#ffb800" });

      kpiData.forEach((kpi, i) => {
        const col = String.fromCharCode(65 + i * 2); // A, C, E, G, I...
        const col2 = String.fromCharCode(65 + i * 2 + 1);
        const kpiRange = sheet.getRange(`${col}4:${col2}5`);
        kpiRange.merge();
        kpiRange.values = [[`${kpi.label}\n${kpi.value}`]];
        kpiRange.format.fill.color = "#0f1a14";
        kpiRange.format.font.color = kpi.color;
        kpiRange.format.font.bold = true;
        kpiRange.format.font.size = 11;
        kpiRange.format.horizontalAlignment = "Center";
        kpiRange.format.verticalAlignment = "Center";
        kpiRange.format.wrapText = true;
      });

      // DONNÉES GRAPHIQUE 1 : Répartition Scopes
      if (document.getElementById("chart-scopes").checked) {
        sheet.getRange("A8").values = [["Scope"]];
        sheet.getRange("B8").values = [["tCO₂e"]];
        const scopeRows = [
          ["Scope 1 Directs", r.scope1],
          ["Scope 2 Énergie", r.scope2],
          ["Scope 3 Indirect", r.scope3]
        ];
        sheet.getRange("A9:B11").values = scopeRows;
        styleHeaders(sheet.getRange("A8:B8"), "#0f1a14", "#8aab97");
        sheet.getRange("A9:B11").format.fill.color = "#0f1a14";
        sheet.getRange("A9:B11").format.font.color = "#e8f5ee";
        sheet.getRange("A9:B11").format.font.size = 10;

        const chart1 = sheet.charts.add(Excel.ChartType.doughnut, sheet.getRange("A8:B11"));
        chart1.title.text = "Répartition Scopes";
        chart1.setPosition("D7", "L17");
        chart1.format.fill.setSolidColor("#0f1a14");
        chart1.title.format.font.color = "#e8f5ee";
      }

      // DONNÉES GRAPHIQUE 2 : Évolution annuelle
      if (document.getElementById("chart-evolution").checked && hist.length > 0) {
        const histStart = 14;
        sheet.getRange(`A${histStart}`).values = [["Année"]];
        sheet.getRange(`B${histStart}`).values = [["tCO₂e"]];
        styleHeaders(sheet.getRange(`A${histStart}:B${histStart}`), "#0f1a14", "#8aab97");

        const allYears = [...hist, [r.metadata.year || 2024, r.total]];
        sheet.getRange(`A${histStart+1}:B${histStart+allYears.length}`).values = allYears;
        sheet.getRange(`A${histStart+1}:B${histStart+allYears.length}`).format.fill.color = "#0f1a14";
        sheet.getRange(`A${histStart+1}:B${histStart+allYears.length}`).format.font.color = "#e8f5ee";
        sheet.getRange(`A${histStart+1}:B${histStart+allYears.length}`).format.font.size = 10;

        const chart2 = sheet.charts.add(
          Excel.ChartType.lineMarkers,
          sheet.getRange(`A${histStart}:B${histStart+allYears.length}`)
        );
        chart2.title.text = "Évolution des émissions CO₂e";
        chart2.setPosition("D18", "L28");
        chart2.format.fill.setSolidColor("#0f1a14");
        chart2.title.format.font.color = "#e8f5ee";
      }

      // TABLEAU RÉSUMÉ
      const tblStart = 20;
      sheet.getRange(`A${tblStart}:C${tblStart}`).values = [["Indicateur", "Valeur", "Unité"]];
      styleHeaders(sheet.getRange(`A${tblStart}:C${tblStart}`), "#141f18", "#8aab97");

      const kpiRows = [
        ["Émissions totales CO₂e",   fmtNum(r.total),      "tCO₂e"],
        ["Scope 1 — Émissions directes", fmtNum(r.scope1), "tCO₂e"],
        ["Scope 2 — Énergie achetée", fmtNum(r.scope2),    "tCO₂e"],
        ["Scope 3 — Émissions indirectes", fmtNum(r.scope3), "tCO₂e"],
        ["Part Scope 1", Math.round(r.scope1/r.total*100) + "%", ""],
        ["Part Scope 3", Math.round(r.scope3/r.total*100) + "%", ""],
      ];
      if (r.intensity) kpiRows.push(["Intensité carbone", fmtNum(r.intensity), `tCO₂e/${r.metadata.activity_unit}`]);

      sheet.getRange(`A${tblStart+1}:C${tblStart+kpiRows.length}`).values = kpiRows;
      for (let i = 0; i < kpiRows.length; i++) {
        sheet.getRange(`A${tblStart+1+i}:C${tblStart+1+i}`).format.fill.color = i%2===0 ? "#0f1a14" : "#141f18";
        sheet.getRange(`A${tblStart+1+i}:C${tblStart+1+i}`).format.font.color = "#e8f5ee";
        sheet.getRange(`A${tblStart+1+i}:C${tblStart+1+i}`).format.font.size = 10;
      }

      // Largeurs
      [["A",200],["B",120],["C",100],["D",100],["E",100],["F",100],["G",100],["H",100],["I",100],["J",100],["K",100],["L",100]].forEach(([c,w]) => {
        sheet.getRange(`${c}:${c}`).format.columnWidth = w;
      });

      // Hauteurs lignes KPI
      sheet.getRange("4:5").format.rowHeight = 35;

      await ctx.sync();
    });

    setStatus("dash-status", "ok", `Dashboard "${sheetName}" généré avec succès`);
    showToast("Dashboard créé !", "success");

  } catch (e) {
    setStatus("dash-status", "error", "Erreur : " + e.message);
    showToast("Erreur dashboard : " + e.message, "error");
  }
}

// Bouton commande ruban
async function generateDashboard(event) {
  await generateDashboardFromPanel();
  event?.completed();
}

function getHistoricalData() {
  const hist = [];
  const y1 = parseFloat(document.getElementById("hist-y1")?.value);
  const v1 = parseFloat(document.getElementById("hist-v1")?.value);
  const y2 = parseFloat(document.getElementById("hist-y2")?.value);
  const v2 = parseFloat(document.getElementById("hist-v2")?.value);
  if (y1 && v1) hist.push([y1, v1]);
  if (y2 && v2) hist.push([y2, v2]);
  return hist;
}

async function refreshDashboard() {
  showToast("Rafraîchissement du dashboard...", "info");
  await generateDashboardFromPanel();
}

// ============================================================
// ANALYSE — ANOMALIES & RECOMMANDATIONS
// ============================================================
async function detectAnomalies() {
  const threshold = parseFloat(document.getElementById("anomaly-threshold").value) || 20;

  // Utiliser données collectées ou simuler
  let allData = [];
  for (const data of Object.values(AppState.collectedData)) {
    allData = allData.concat(data);
  }

  if (allData.length === 0 && AppState.carbonResults) {
    // Mode démo : simuler anomalies sur les résultats carbon
    const r = AppState.carbonResults;
    allData = [
      ["Gaz naturel", r.scope1 * 0.6], ["Gaz naturel", r.scope1 * 0.7],
      ["Gaz naturel", r.scope1 * 1.8], // anomalie !
      ["Électricité",  r.scope2 * 0.9], ["Électricité",  r.scope2 * 1.1],
      ["Transport",    r.scope3 * 0.5], ["Transport",    r.scope3 * 0.55],
      ["Transport",    r.scope3 * 0.9], // anomalie !
    ];
  }

  if (allData.length === 0) {
    showToast("Collectez d'abord des données ou calculez le bilan carbone", "warn");
    return;
  }

  const anomalies = AppState.calculator.detectAnomalies(allData, threshold);
  AppState.anomalies = anomalies;

  displayAnomalies(anomalies);
}

function displayAnomalies(anomalies) {
  const container = document.getElementById("anomalies-container");
  const list      = document.getElementById("anomaly-list");

  if (anomalies.length === 0) {
    list.innerHTML = '<div class="anomaly-item info"><div class="anomaly-title">✅ Aucune anomalie détectée</div><div class="anomaly-desc">Les données sont dans les seuils normaux.</div></div>';
    container.style.display = "block";
    return;
  }

  list.innerHTML = anomalies.map(a => `
    <div class="anomaly-item ${a.severity}">
      <div class="anomaly-title">
        ${a.severity === "critical" ? "🔴" : a.severity === "warning" ? "🟠" : "🔵"}
        ${a.source} — Écart ${a.deviation}%
      </div>
      <div class="anomaly-desc">${a.message}</div>
      <div class="anomaly-action" onclick="applyAnomalyAction('${a.source}')">
        ➤ ${a.action}
      </div>
    </div>
  `).join("");

  container.style.display = "block";
  showToast(`${anomalies.length} anomalie(s) détectée(s)`, anomalies.some(a => a.severity === "critical") ? "warn" : "info");
}

function applyAnomalyAction(source) {
  showToast(`Action : vérifier la source "${source}" dans les données collectées`, "info");
}

function generateRecommendations() {
  if (!AppState.carbonResults) {
    showToast("Calculez d'abord le bilan carbone", "warn");
    return;
  }

  const target  = parseFloat(document.getElementById("reduction-target").value)  || 30;
  const horizon = parseInt(document.getElementById("reduction-horizon").value) || 3;

  const recos = AppState.calculator.generateRecommendations(AppState.carbonResults, target, horizon);
  AppState.recommendations = recos;

  displayRecommendations(recos);
}

function displayRecommendations(recos) {
  const container = document.getElementById("recos-container");
  const list      = document.getElementById("reco-list");

  list.innerHTML = recos.map(r => `
    <div class="reco-item">
      <div class="reco-icon">${r.icon}</div>
      <div class="reco-body">
        <div class="reco-title">${r.title}</div>
        <div class="reco-desc">${r.description}<br>
          <span style="color:var(--text-muted);font-size:9px;">Horizon : ${r.horizon} | Investissement : ${r.investment} | ${r.scope}</span>
        </div>
      </div>
      <div class="reco-saving">-${fmtNum(r.saving_tco2e)} t<br><span style="font-size:8px;">(-${r.saving_pct}%)</span></div>
    </div>
  `).join("");

  container.style.display = "block";
  showToast(`${recos.length} recommandation(s) générée(s)`, "success");
}

async function exportAnalysisToExcel() {
  if (AppState.anomalies.length === 0 && AppState.recommendations.length === 0) {
    showToast("Lancez d'abord l'analyse et les recommandations", "warn");
    return;
  }

  const sheetName = `ESG_Analyse_${formatDateSuffix()}`;

  await Excel.run(async (ctx) => {
    try { ctx.workbook.worksheets.getItem(sheetName).delete(); await ctx.sync(); } catch(_) {}
    const sheet = ctx.workbook.worksheets.add(sheetName);
    sheet.activate();
    sheet.getRange("A1:G1").format.fill.color = "#0a0f0d";

    // Titre
    sheet.getRange("A1:G1").merge();
    sheet.getRange("A1").values = [["ANALYSE ESG — ANOMALIES & RECOMMANDATIONS"]];
    styleTitle(sheet.getRange("A1:G1"), "#ff6b35");

    // Anomalies
    let row = 3;
    sheet.getRange(`A${row}:G${row}`).values = [["Source", "Valeur", "Moyenne", "Écart %", "Sévérité", "Action", ""]];
    styleHeaders(sheet.getRange(`A${row}:G${row}`), "#141f18", "#ff6b35");
    row++;

    if (AppState.anomalies.length > 0) {
      const anomRows = AppState.anomalies.map(a => [a.source, a.value, a.mean, a.deviation+"%", a.severity, a.action, ""]);
      sheet.getRange(`A${row}:G${row + anomRows.length - 1}`).values = anomRows;
      for (let i = 0; i < anomRows.length; i++) {
        const color = anomRows[i][4] === "critical" ? "#2a1010" : anomRows[i][4] === "warning" ? "#2a1a08" : "#0f1a20";
        sheet.getRange(`A${row+i}:G${row+i}`).format.fill.color = color;
        sheet.getRange(`A${row+i}:G${row+i}`).format.font.color = "#e8f5ee";
        sheet.getRange(`A${row+i}:G${row+i}`).format.font.size = 10;
      }
      row += anomRows.length + 2;
    } else {
      sheet.getRange(`A${row}`).values = [["Aucune anomalie détectée"]];
      sheet.getRange(`A${row}`).format.font.color = "#00d97e";
      row += 2;
    }

    // Recommandations
    sheet.getRange(`A${row}:G${row}`).values = [["Priorité", "Action", "Scope", "Économie (tCO₂e)", "Économie %", "Horizon", "Investissement"]];
    styleHeaders(sheet.getRange(`A${row}:G${row}`), "#141f18", "#00b8d4");
    row++;

    if (AppState.recommendations.length > 0) {
      const recoRows = AppState.recommendations.map((r, i) => [i+1, r.title, r.scope, r.saving_tco2e, r.saving_pct+"%", r.horizon, r.investment]);
      sheet.getRange(`A${row}:G${row + recoRows.length - 1}`).values = recoRows;
      for (let i = 0; i < recoRows.length; i++) {
        sheet.getRange(`A${row+i}:G${row+i}`).format.fill.color = i%2===0 ? "#0f1a14" : "#141f18";
        sheet.getRange(`A${row+i}:G${row+i}`).format.font.color = "#e8f5ee";
        sheet.getRange(`A${row+i}:G${row+i}`).format.font.size = 10;
      }
    }

    [["A",50],["B",220],["C",100],["D",130],["E",90],["F",90],["G",110]].forEach(([c,w]) => {
      sheet.getRange(`${c}:${c}`).format.columnWidth = w;
    });

    await ctx.sync();
  });

  showToast(`Analyse exportée → "${sheetName}"`, "success");
}

// ============================================================
// RAPPORT ESG
// ============================================================
let selectedFormat = "excel";
function selectFormat(fmt) {
  selectedFormat = fmt;
  ["excel","word","pdf"].forEach(f => {
    document.getElementById(`fmt-${f}`).classList.toggle("active", f === fmt);
  });
  const notes = {
    excel: "ℹ️ Excel : rapport intégré dans un nouvel onglet",
    word:  "ℹ️ Word : génère une feuille de données à copier dans Word (API Word non disponible en task pane basique)",
    pdf:   "ℹ️ PDF : génère une feuille formatée optimisée pour l'impression PDF"
  };
  document.getElementById("format-note").textContent = notes[fmt];
}

async function generateReport() {
  if (!AppState.carbonResults) {
    showToast("Calculez d'abord le bilan carbone (onglet Carbone)", "warn");
    return;
  }

  const company = document.getElementById("report-company").value || "Entreprise";
  const year    = document.getElementById("report-year").value    || 2024;
  const author  = document.getElementById("report-author").value  || "";
  const sector  = document.getElementById("report-sector").value  || "";

  setStatus("report-status", "running", "Génération du rapport en cours...");

  const sheetName = `Rapport_ESG_${year}`;

  try {
    await Excel.run(async (ctx) => {
      try { ctx.workbook.worksheets.getItem(sheetName).delete(); await ctx.sync(); } catch(_) {}
      const sheet = ctx.workbook.worksheets.add(sheetName);
      sheet.tabColor = "#00b8d4";
      sheet.activate();

      // Fond global
      sheet.getRange("A1:L80").format.fill.color = "#0a0f0d";

      let row = 1;

      // PAGE DE GARDE
      sheet.getRange(`A${row}:L${row+1}`).merge();
      sheet.getRange(`A${row}`).values = [["RAPPORT ESG"]];
      sheet.getRange(`A${row}`).format.font.color = "#00d97e";
      sheet.getRange(`A${row}`).format.font.size = 24;
      sheet.getRange(`A${row}`).format.font.bold = true;
      sheet.getRange(`A${row}`).format.font.name = "Segoe UI";
      sheet.getRange(`A${row}`).format.horizontalAlignment = "Center";
      sheet.getRange(`A${row}:L${row+1}`).format.rowHeight = 40;
      row += 2;

      sheet.getRange(`A${row}:L${row}`).merge();
      sheet.getRange(`A${row}`).values = [[company.toUpperCase()]];
      sheet.getRange(`A${row}`).format.font.color = "#e8f5ee";
      sheet.getRange(`A${row}`).format.font.size = 16;
      sheet.getRange(`A${row}`).format.font.bold = true;
      sheet.getRange(`A${row}`).format.horizontalAlignment = "Center";
      row++;

      sheet.getRange(`A${row}:L${row}`).merge();
      sheet.getRange(`A${row}`).values = [[`Exercice ${year} | ${sector} | Responsable : ${author}`]];
      sheet.getRange(`A${row}`).format.font.color = "#4a7060";
      sheet.getRange(`A${row}`).format.font.size = 10;
      sheet.getRange(`A${row}`).format.horizontalAlignment = "Center";
      row += 2;

      const r = AppState.carbonResults;

      // SECTION 1 : BILAN CARBONE
      if (document.getElementById("rpt-bilan").checked) {
        sheet.getRange(`A${row}:L${row}`).merge();
        sheet.getRange(`A${row}`).values = [["1. BILAN CARBONE — SYNTHÈSE PAR SCOPE"]];
        styleTitle(sheet.getRange(`A${row}:L${row}`), "#00d97e", 12);
        row += 2;

        sheet.getRange(`A${row}:D${row}`).values = [["Scope", "Émissions (tCO₂e)", "Part du total (%)", "Variation N-1"]];
        styleHeaders(sheet.getRange(`A${row}:D${row}`), "#141f18", "#8aab97");
        row++;

        const scopeData = [
          ["Scope 1 — Émissions directes (combustion, procédés)", fmtNum(r.scope1), (r.scope1/r.total*100).toFixed(1)+"%", "N/A"],
          ["Scope 2 — Énergie achetée (électricité, chaleur)",    fmtNum(r.scope2), (r.scope2/r.total*100).toFixed(1)+"%", "N/A"],
          ["Scope 3 — Émissions indirectes (transport, achats)",  fmtNum(r.scope3), (r.scope3/r.total*100).toFixed(1)+"%", "N/A"],
          ["TOTAL",                                               fmtNum(r.total),  "100%",                                "N/A"]
        ];
        const scopeColors = ["#00d97e","#00b8d4","#7c5cbf","#e8f5ee"];
        for (let i = 0; i < scopeData.length; i++) {
          sheet.getRange(`A${row}:D${row}`).values = [scopeData[i]];
          sheet.getRange(`A${row}:D${row}`).format.fill.color = i%2===0 ? "#0f1a14" : "#141f18";
          sheet.getRange(`A${row}:D${row}`).format.font.color = scopeColors[i];
          sheet.getRange(`A${row}:D${row}`).format.font.size = 10;
          if (i === 3) sheet.getRange(`A${row}:D${row}`).format.font.bold = true;
          row++;
        }

        if (r.intensity) {
          row++;
          sheet.getRange(`A${row}:D${row}`).merge();
          sheet.getRange(`A${row}`).values = [[`Intensité carbone : ${fmtNum(r.intensity)} tCO₂e par ${r.metadata.activity_unit}`]];
          sheet.getRange(`A${row}`).format.font.color = "#ffb800";
          sheet.getRange(`A${row}`).format.font.bold = true;
          sheet.getRange(`A${row}`).format.font.size = 11;
        }
        row += 3;
      }

      // SECTION 2 : KPI
      if (document.getElementById("rpt-kpi").checked) {
        sheet.getRange(`A${row}:L${row}`).merge();
        sheet.getRange(`A${row}`).values = [["2. INDICATEURS DE PERFORMANCE ESG"]];
        styleTitle(sheet.getRange(`A${row}:L${row}`), "#00b8d4", 12);
        row += 2;

        sheet.getRange(`A${row}:C${row}`).values = [["Indicateur", "Valeur", "Commentaire"]];
        styleHeaders(sheet.getRange(`A${row}:C${row}`), "#141f18", "#8aab97");
        row++;

        const kpiData = [
          ["Émissions totales CO₂e",          fmtNum(r.total) + " tCO₂e",      "GHG Protocol — Périmètre opérationnel"],
          ["Scope 1 — Directes",               fmtNum(r.scope1) + " tCO₂e",     "Combustion interne + procédés industriels"],
          ["Scope 2 — Énergie achetée",        fmtNum(r.scope2) + " tCO₂e",     "Location-based (RFE ADEME 2024)"],
          ["Scope 3 — Indirectes",             fmtNum(r.scope3) + " tCO₂e",     "Transport, achats, déchets"],
          ["Part du Scope 3 dans le total",   (r.scope3/r.total*100).toFixed(1)+"%", "Levier majeur de réduction"],
        ];
        if (r.intensity) kpiData.push([`Intensité carbone (/${r.metadata.activity_unit})`, fmtNum(r.intensity) + " tCO₂e", "Indicateur d'efficacité carbone"]);

        for (let i = 0; i < kpiData.length; i++) {
          sheet.getRange(`A${row}:C${row}`).values = [kpiData[i]];
          sheet.getRange(`A${row}:C${row}`).format.fill.color = i%2===0 ? "#0f1a14" : "#141f18";
          sheet.getRange(`A${row}:C${row}`).format.font.color = "#e8f5ee";
          sheet.getRange(`A${row}:C${row}`).format.font.size = 10;
          row++;
        }
        row += 2;
      }

      // SECTION 3 : ANOMALIES
      if (document.getElementById("rpt-anomalies").checked && AppState.anomalies.length > 0) {
        sheet.getRange(`A${row}:L${row}`).merge();
        sheet.getRange(`A${row}`).values = [["3. ANOMALIES & ALERTES"]];
        styleTitle(sheet.getRange(`A${row}:L${row}`), "#ff6b35", 12);
        row += 2;

        for (const a of AppState.anomalies) {
          sheet.getRange(`A${row}:E${row}`).values = [[
            a.severity === "critical" ? "⚠ CRITIQUE" : "⚡ ALERTE",
            a.source, `Écart : ${a.deviation}%`, a.message, a.action
          ]];
          const bgColor = a.severity === "critical" ? "#2a1010" : "#2a1a08";
          sheet.getRange(`A${row}:E${row}`).format.fill.color = bgColor;
          sheet.getRange(`A${row}:E${row}`).format.font.color = a.severity === "critical" ? "#e63946" : "#ff6b35";
          sheet.getRange(`A${row}:E${row}`).format.font.size = 10;
          row++;
        }
        row += 2;
      }

      // SECTION 4 : RECOMMANDATIONS
      if (document.getElementById("rpt-recos").checked && AppState.recommendations.length > 0) {
        sheet.getRange(`A${row}:L${row}`).merge();
        sheet.getRange(`A${row}`).values = [["4. PLAN D'ACTIONS — RECOMMANDATIONS DE RÉDUCTION"]];
        styleTitle(sheet.getRange(`A${row}:L${row}`), "#7c5cbf", 12);
        row += 2;

        sheet.getRange(`A${row}:E${row}`).values = [["Action", "Scope", "Économie tCO₂e", "Horizon", "Investissement"]];
        styleHeaders(sheet.getRange(`A${row}:E${row}`), "#141f18", "#8aab97");
        row++;

        for (let i = 0; i < AppState.recommendations.length; i++) {
          const rec = AppState.recommendations[i];
          sheet.getRange(`A${row}:E${row}`).values = [[rec.title, rec.scope, fmtNum(rec.saving_tco2e), rec.horizon, rec.investment]];
          sheet.getRange(`A${row}:E${row}`).format.fill.color = i%2===0 ? "#0f1a14" : "#141f18";
          sheet.getRange(`A${row}:E${row}`).format.font.color = "#e8f5ee";
          sheet.getRange(`A${row}:E${row}`).format.font.size = 10;
          row++;
        }
        row += 2;
      }

      // FOOTER
      sheet.getRange(`A${row}:L${row}`).merge();
      sheet.getRange(`A${row}`).values = [[`Rapport généré par ESG Analyzer Pro | ${new Date().toLocaleDateString("fr-FR")} | Facteurs ADEME Base Carbone 2024 | GHG Protocol`]];
      sheet.getRange(`A${row}`).format.font.color = "#2a3d32";
      sheet.getRange(`A${row}`).format.font.size = 8;
      sheet.getRange(`A${row}`).format.horizontalAlignment = "Center";

      // Largeurs colonnes
      [["A",220],["B",120],["C",120],["D",120],["E",150],["F",100],["G",100],["H",100],["I",100],["J",100],["K",100],["L",100]].forEach(([c,w]) => {
        sheet.getRange(`${c}:${c}`).format.columnWidth = w;
      });

      await ctx.sync();
    });

    setStatus("report-status", "ok", `Rapport "${sheetName}" généré avec succès`);
    showToast("Rapport ESG généré !", "success");

  } catch (e) {
    setStatus("report-status", "error", "Erreur : " + e.message);
    showToast("Erreur génération rapport : " + e.message, "error");
  }
}

// Commande ruban quickCarbonCalc
async function quickCarbonCalc(event) {
  showToast("Utilisez le panneau ESG pour le calcul complet", "info");
  event?.completed();
}

// ============================================================
// UTILITAIRES UI & EXCEL
// ============================================================
function styleTitle(range, color = "#00d97e", size = 13) {
  range.format.fill.color = "#0a0f0d";
  range.format.font.color = color;
  range.format.font.bold  = true;
  range.format.font.size  = size;
  range.format.font.name  = "Segoe UI";
  range.format.horizontalAlignment = "Left";
  range.format.rowHeight = 28;
}

function styleHeaders(range, bgColor = "#141f18", fontColor = "#8aab97") {
  range.format.fill.color = bgColor;
  range.format.font.color = fontColor;
  range.format.font.bold  = true;
  range.format.font.size  = 9;
}

function setStatus(id, type, msg) {
  const bar = document.getElementById(id);
  if (!bar) return;
  const dot  = bar.querySelector(".status-dot");
  const span = bar.querySelector("span");
  dot.className = "status-dot";
  if (type) dot.classList.add(type);
  span.textContent = msg;
}

function showToast(msg, type = "info", duration = 3500) {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), duration);
}

function fmtNum(n) {
  if (n === null || n === undefined) return "—";
  return new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 2 }).format(n);
}

function formatDateSuffix() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2,"0")}${String(d.getMonth()+1).padStart(2,"0")}${d.getFullYear()}`;
}

function addHistoricalYear() {
  showToast("Remplissez les lignes existantes pour l'historique", "info");
}
