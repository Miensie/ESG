/**
 * ESG Analyzer Pro — Backend Node.js
 * Port 3001 — API REST pour génération de rapports Word/PDF
 */

const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");

// Optionnel : génération Word avec docx
// const { Document, Packer, Paragraph, TextRun } = require("docx");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: ["https://miensie.github.io", "http://localhost:3000"] }));
app.use(express.json({ limit: "10mb" }));
app.use("/downloads", express.static(path.join(__dirname, "downloads")));

// Créer dossier downloads
if (!fs.existsSync(path.join(__dirname, "downloads"))) {
  fs.mkdirSync(path.join(__dirname, "downloads"));
}

/* ─── ROUTE SANTÉ ─────────────────────────────────────── */
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", version: "1.0.0", timestamp: new Date().toISOString() });
});

/* ─── ROUTE : FACTEURS D'ÉMISSION ────────────────────── */
app.get("/api/emission-factors", (req, res) => {
  const { ref = "ademe", year = "2024" } = req.query;

  // En production : charger depuis base de données ou fichier JSON ADEME
  const factors = require("./data/emission-factors.json");
  res.json({ ref, year, factors: factors[ref] || factors.ademe });
});

/* ─── ROUTE : CALCUL BILAN CARBONE ──────────────────── */
app.post("/api/carbon/calculate", (req, res) => {
  const { scope1Data, scope2Data, scope3Data, ref = "ademe" } = req.body;

  try {
    const result = calculateCarbonFootprint(scope1Data, scope2Data, scope3Data, ref);
    res.json({ success: true, result });
  } catch (e) {
    res.status(400).json({ success: false, error: e.message });
  }
});

function calculateCarbonFootprint(s1, s2, s3, ref) {
  // Calcul côté serveur si nécessaire (pour validation / audit trail)
  const scope1Total = (s1 || []).reduce((acc, row) => acc + (parseFloat(row.co2e) || 0), 0);
  const scope2Total = (s2 || []).reduce((acc, row) => acc + (parseFloat(row.co2e) || 0), 0);
  const scope3Total = (s3 || []).reduce((acc, row) => acc + (parseFloat(row.co2e) || 0), 0);
  const total = scope1Total + scope2Total + scope3Total;

  return {
    scope1: { total: scope1Total },
    scope2: { total: scope2Total },
    scope3: { total: scope3Total },
    total,
    methodology: ref,
    calculatedAt: new Date().toISOString(),
  };
}

/* ─── ROUTE : GÉNÉRATION RAPPORT ─────────────────────── */
app.post("/api/generate-report", async (req, res) => {
  const { company, period, format, data } = req.body;

  try {
    const filename = `ESG_Rapport_${company.replace(/\s+/g, "_")}_${period}.${format === "pdf" ? "pdf" : format === "word" ? "docx" : "xlsx"}`;
    const filePath = path.join(__dirname, "downloads", filename);

    if (format === "word") {
      await generateWordReport(filePath, company, period, data);
    } else if (format === "pdf") {
      await generatePDFReport(filePath, company, period, data);
    } else {
      // Excel géré côté client via Office.js
      return res.json({ success: true, message: "Rapport Excel généré côté client" });
    }

    const downloadUrl = `http://localhost:${PORT}/downloads/${filename}`;
    res.json({ success: true, downloadUrl, filename });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

async function generateWordReport(filePath, company, period, data) {
  // Implémentation avec la lib 'docx'
  // const { Document, Packer, Paragraph, HeadingLevel, Table, TableRow, TableCell } = require("docx");
  // ... génération document Word structuré
  // Pour l'exemple, créer un fichier texte simulé
  const content = generateReportText(company, period, data);
  fs.writeFileSync(filePath.replace(".docx", ".txt"), content);
  // En production : utiliser Packer.toBuffer(doc) et fs.writeFileSync(filePath, buffer)
}

async function generatePDFReport(filePath, company, period, data) {
  // Implémentation avec 'puppeteer' ou 'pdfkit'
  // const PDFDocument = require("pdfkit");
  // const doc = new PDFDocument();
  // doc.pipe(fs.createWriteStream(filePath));
  // ... contenu PDF
  // doc.end();
  const content = generateReportText(company, period, data);
  fs.writeFileSync(filePath.replace(".pdf", ".txt"), content);
}

function generateReportText(company, period, data) {
  const carbon = data?.carbon || { scope1: { total: 0 }, scope2: { total: 0 }, scope3: { total: 0 }, total: 0 };
  return `
RAPPORT ESG — ${company.toUpperCase()}
Période : ${period}
Généré le : ${new Date().toLocaleDateString("fr-FR")}
${"═".repeat(60)}

1. SYNTHÈSE EXÉCUTIVE
${company} publie son rapport ESG pour la période ${period}.
Ce rapport a été généré automatiquement par ESG Analyzer Pro.

2. BILAN CARBONE (GHG Protocol)
   Scope 1 (Émissions directes) : ${carbon.scope1.total.toFixed(1)} t CO₂e
   Scope 2 (Énergie achetée)    : ${carbon.scope2.total.toFixed(1)} t CO₂e
   Scope 3 (Émissions indirectes): ${carbon.scope3.total.toFixed(1)} t CO₂e
   TOTAL                         : ${carbon.total.toFixed(1)} t CO₂e

3. INDICATEURS KPI ESG
   Intensité carbone : ${(carbon.total / 125).toFixed(2)} t CO₂e / M€ CA
   Énergie totale    : 48 750 MWh
   % EnR             : 32%
   Recyclage déchets : 72%
`;
}

/* ─── ROUTE : ANALYSE ANOMALIES ───────────────────────── */
app.post("/api/analysis/anomalies", (req, res) => {
  const { data } = req.body;
  const anomalies = detectServerSideAnomalies(data);
  res.json({ success: true, anomalies });
});

function detectServerSideAnomalies(data) {
  // Logique de détection avancée côté serveur
  // (complément au module client)
  return [];
}

/* ─── ROUTE : FACTEURS SECTORIELS BENCHMARK ──────────── */
app.get("/api/benchmark/:sector", (req, res) => {
  const { sector } = req.params;
  const benchmarks = {
    industrie: { intensite_co2: 85.2, energie_mwh: 62000, recyclage: 68 },
    chimie: { intensite_co2: 142.0, energie_mwh: 95000, recyclage: 55 },
    agroalimentaire: { intensite_co2: 48.5, energie_mwh: 38000, recyclage: 78 },
    construction: { intensite_co2: 67.3, energie_mwh: 28000, recyclage: 82 },
  };
  res.json({ sector, benchmark: benchmarks[sector] || benchmarks.industrie });
});

/* ─── DÉMARRAGE ───────────────────────────────────────── */
app.listen(PORT, () => {
  console.log(`
  ╔══════════════════════════════════════╗
  ║  ESG Analyzer Pro — Backend         ║
  ║  Port : ${PORT}                         ║
  ║  Santé : http://localhost:${PORT}/api/health  ║
  ╚══════════════════════════════════════╝
  `);
});

module.exports = app;
