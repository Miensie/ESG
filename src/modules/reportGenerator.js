/**
 * ============================================================
 * MODULE : reportGenerator.js
 * Génération du rapport ESG synthétique
 * Sortie : HTML imprimable (convertible en PDF via Ctrl+P)
 * ============================================================
 */

"use strict";

/**
 * Génère le HTML complet du rapport ESG
 * @param {Object} bilan       - Résultat de computeFullBilan()
 * @param {Array}  anomalies   - Résultat de detectAnomalies()
 * @param {Array}  suggestions - Résultat de generateSuggestions()
 * @returns {string} HTML complet prêt à l'impression
 */
function generateReportHTML(bilan, anomalies, suggestions) {
  const now = new Date().toLocaleDateString("fr-FR", {
    day: "2-digit", month: "long", year: "numeric"
  });

  const scopeColor = { 1: "#E05252", 2: "#F5A623", 3: "#2ECC8E" };

  function barPct(val, max) {
    return Math.min(100, Math.round((val / (max || 1)) * 100));
  }

  const maxScope = Math.max(bilan.scope1.total, bilan.scope2.total, bilan.scope3.total);

  // Lignes de détail
  function detailRows(lines, scopeColor) {
    return lines.map(l => `
      <tr>
        <td>${l.source}</td>
        <td>${l.label}</td>
        <td style="text-align:right">${l.quantity.toLocaleString("fr-FR")}</td>
        <td style="text-align:right"><strong>${l.tCO2eq.toLocaleString("fr-FR", {minimumFractionDigits:2})}</strong></td>
        <td>
          <div style="background:#eee;border-radius:3px;height:8px;width:100px">
            <div style="background:${scopeColor};height:8px;border-radius:3px;width:${barPct(l.tCO2eq, lines.reduce((a,x)=>a+x.tCO2eq,0))}%"></div>
          </div>
        </td>
      </tr>
    `).join("");
  }

  function anomalyRows(list) {
    if (!list.length) return `<p style="color:#666;font-style:italic">Aucune anomalie détectée.</p>`;
    return list.map(a => `
      <div style="padding:10px 14px;margin-bottom:8px;border-left:3px solid ${
        a.severity==="error"?"#E05252":a.severity==="warn"?"#F5A623":"#2ECC8E"
      };background:#f9f9f9;border-radius:0 4px 4px 0">
        <span style="font-weight:700;font-size:11px;text-transform:uppercase;color:#666">${a.scope}</span><br>
        <span style="font-size:13px">${a.message}</span>
      </div>
    `).join("");
  }

  function suggestionRows(list) {
    const priorityColor = { high:"#E05252", medium:"#F5A623", low:"#2ECC8E" };
    const priorityLabel = { high:"Priorité haute", medium:"Priorité moyenne", low:"Priorité basse" };
    return list.map(s => `
      <div style="padding:12px;margin-bottom:10px;border:1px solid #ddd;border-radius:6px;background:#fff">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
          <span style="background:${priorityColor[s.priority]};color:#fff;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700">${priorityLabel[s.priority]}</span>
          <span style="background:#e8f5e9;color:#1A6B4A;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700">${s.scope}</span>
        </div>
        <div style="font-weight:700;font-size:14px;margin-bottom:4px">${s.action}</div>
        <div style="color:#555;font-size:12px;margin-bottom:4px">${s.detail}</div>
        <div style="color:#1A6B4A;font-size:12px;font-weight:600">💚 Potentiel : ${s.potentiel}</div>
      </div>
    `).join("");
  }

  return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Rapport ESG ${bilan.annee} — ${bilan.entreprise}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap');
  * { box-sizing: border-box; }
  body { font-family: 'Sora', sans-serif; font-size: 13px; color: #1a1a2e; margin: 0; padding: 0; background: #fff; }

  /* PAGE PRINT */
  @media print {
    .no-print { display: none !important; }
    body { padding: 0; }
    .page-break { page-break-before: always; }
  }

  .report-wrapper { max-width: 900px; margin: 0 auto; padding: 40px; }

  /* Cover */
  .cover { background: linear-gradient(135deg, #0D3B2E 0%, #1C2B3A 100%); color: #F0F4F1; padding: 60px; border-radius: 12px; margin-bottom: 40px; position: relative; overflow: hidden; }
  .cover::before { content:''; position:absolute; top:-40px; right:-40px; width:200px; height:200px; background:rgba(46,204,142,0.1); border-radius:50%; }
  .cover-tag { font-size: 10px; font-weight: 700; letter-spacing: 0.2em; text-transform: uppercase; color: #2ECC8E; margin-bottom: 12px; }
  .cover h1 { font-size: 32px; font-weight: 700; line-height: 1.2; margin-bottom: 8px; }
  .cover-sub { font-size: 16px; color: rgba(240,244,241,0.7); margin-bottom: 32px; }
  .cover-meta { display: flex; gap: 32px; }
  .cover-meta-item label { display: block; font-size: 10px; text-transform: uppercase; letter-spacing: 0.1em; color: rgba(240,244,241,0.5); margin-bottom: 2px; }
  .cover-meta-item span { font-family: 'DM Mono', monospace; font-size: 14px; color: #2ECC8E; }

  /* Sections */
  h2 { font-size: 18px; font-weight: 700; color: #0D3B2E; margin: 32px 0 16px; padding-bottom: 8px; border-bottom: 2px solid #2ECC8E; display: flex; align-items: center; gap: 8px; }
  h3 { font-size: 14px; font-weight: 700; margin: 20px 0 10px; }

  /* KPI Cards */
  .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
  .kpi-box { background: #f8fffe; border: 1px solid #e0f0ea; border-radius: 10px; padding: 16px; text-align: center; }
  .kpi-box.total { background: linear-gradient(135deg, #0D3B2E, #1A6B4A); color: #fff; border-color: transparent; }
  .kpi-box label { display: block; font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #666; margin-bottom: 4px; }
  .kpi-box.total label { color: rgba(255,255,255,0.7); }
  .kpi-box .val { font-family: 'DM Mono', monospace; font-size: 24px; font-weight: 500; color: #1A6B4A; }
  .kpi-box.total .val { color: #2ECC8E; font-size: 28px; }
  .kpi-box .unit { font-size: 10px; color: #999; }
  .kpi-box.total .unit { color: rgba(255,255,255,0.6); }

  /* Scope bars */
  .scope-section { margin-bottom: 16px; }
  .scope-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
  .scope-name { font-weight: 700; font-size: 13px; }
  .scope-value { font-family: 'DM Mono', monospace; font-weight: 500; }
  .scope-bar-bg { background: #f0f0f0; border-radius: 4px; height: 10px; }
  .scope-bar-fill { height: 10px; border-radius: 4px; transition: width 0.5s; }

  /* Tables */
  table { width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 12px; }
  th { background: #1C2B3A; color: #2ECC8E; font-size: 10px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; padding: 8px 10px; text-align: left; }
  td { padding: 7px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
  tr:hover td { background: #f8fffe; }
  .scope-badge-r { display:inline-block; padding:2px 8px; border-radius:12px; font-size:9px; font-weight:700; }

  /* Print button */
  .print-btn { no-print; background: linear-gradient(135deg,#1A6B4A,#2ECC8E); color:#fff; border:none; padding:12px 24px; border-radius:8px; font-family:'Sora',sans-serif; font-size:14px; font-weight:700; cursor:pointer; display:block; margin:0 auto 32px; }

  /* Footer */
  .report-footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; display: flex; justify-content: space-between; color: #999; font-size: 11px; }
</style>
</head>
<body>

<div class="report-wrapper">

  <!-- Bouton impression -->
  <button class="print-btn no-print" onclick="window.print()">🖨️ Imprimer / Exporter PDF</button>

  <!-- Couverture -->
  <div class="cover">
    <div class="cover-tag">Rapport ESG Annuel · GHG Protocol / ADEME</div>
    <h1>${bilan.entreprise}</h1>
    <div class="cover-sub">Bilan Carbone Scope 1, 2 & 3 — Exercice ${bilan.annee}</div>
    <div class="cover-meta">
      <div class="cover-meta-item">
        <label>Secteur</label>
        <span>${bilan.secteur}</span>
      </div>
      <div class="cover-meta-item">
        <label>Total émissions</label>
        <span>${bilan.grandTotal.toLocaleString("fr-FR")} tCO2eq</span>
      </div>
      <div class="cover-meta-item">
        <label>Intensité carbone</label>
        <span>${bilan.intensite} ${bilan.intensiteUnit || "—"}</span>
      </div>
      <div class="cover-meta-item">
        <label>Généré le</label>
        <span>${now}</span>
      </div>
    </div>
  </div>

  <!-- 1. Synthèse KPI -->
  <h2>📊 Synthèse des émissions</h2>
  <div class="kpi-row">
    <div class="kpi-box scope1">
      <label>Scope 1</label>
      <div class="val">${bilan.scope1.total.toLocaleString("fr-FR")}</div>
      <div class="unit">tCO2eq · ${bilan.scope1.pct}%</div>
    </div>
    <div class="kpi-box scope2">
      <label>Scope 2</label>
      <div class="val">${bilan.scope2.total.toLocaleString("fr-FR")}</div>
      <div class="unit">tCO2eq · ${bilan.scope2.pct}%</div>
    </div>
    <div class="kpi-box scope3">
      <label>Scope 3</label>
      <div class="val">${bilan.scope3.total.toLocaleString("fr-FR")}</div>
      <div class="unit">tCO2eq · ${bilan.scope3.pct}%</div>
    </div>
    <div class="kpi-box total">
      <label>Total Bilan</label>
      <div class="val">${bilan.grandTotal.toLocaleString("fr-FR")}</div>
      <div class="unit">tCO2eq</div>
    </div>
  </div>

  <!-- Barres Scopes -->
  ${[1,2,3].map(s => {
    const sc = bilan[`scope${s}`];
    const color = scopeColor[s];
    return `
    <div class="scope-section">
      <div class="scope-header">
        <span class="scope-name" style="color:${color}">● Scope ${s} — Émissions ${["directes","énergie achetée","indirectes"][s-1]}</span>
        <span class="scope-value">${sc.total.toLocaleString("fr-FR")} tCO2eq (${sc.pct}%)</span>
      </div>
      <div class="scope-bar-bg">
        <div class="scope-bar-fill" style="width:${barPct(sc.total, maxScope)}%;background:${color}"></div>
      </div>
    </div>`;
  }).join("")}

  <!-- 2. Détails Scope 1 -->
  <div class="page-break"></div>
  <h2>🔴 Détail Scope 1 — Émissions directes</h2>
  <table>
    <thead><tr><th>Source</th><th>Type</th><th>Quantité</th><th>tCO2eq</th><th>Part</th></tr></thead>
    <tbody>${detailRows(bilan.scope1.lines, scopeColor[1])}</tbody>
  </table>

  <!-- 3. Détails Scope 2 -->
  <h2>🟡 Détail Scope 2 — Énergie achetée</h2>
  <table>
    <thead><tr><th>Source</th><th>Type</th><th>Quantité</th><th>tCO2eq</th><th>Part</th></tr></thead>
    <tbody>${detailRows(bilan.scope2.lines, scopeColor[2])}</tbody>
  </table>

  <!-- 4. Détails Scope 3 -->
  <h2>🟢 Détail Scope 3 — Émissions indirectes</h2>
  <table>
    <thead><tr><th>Source</th><th>Type</th><th>Quantité</th><th>tCO2eq</th><th>Part</th></tr></thead>
    <tbody>${detailRows(bilan.scope3.lines, scopeColor[3])}</tbody>
  </table>

  <!-- 5. Anomalies -->
  <div class="page-break"></div>
  <h2>⚠️ Analyse qualité des données</h2>
  ${anomalyRows(anomalies)}

  <!-- 6. Recommandations -->
  <h2>💡 Plan d'actions — Réduction des émissions</h2>
  ${suggestionRows(suggestions)}

  <!-- Footer -->
  <div class="report-footer">
    <span>ESG Analyzer Pro — ${bilan.entreprise} — Bilan ${bilan.annee}</span>
    <span>Méthodologie : GHG Protocol Corporate Standard / Base Carbone ADEME 2024</span>
  </div>

</div>
</body>
</html>`;
}

/**
 * Ouvre le rapport dans un nouvel onglet pour impression/PDF
 */
function openReport(bilan, anomalies, suggestions) {
  const html = generateReportHTML(bilan, anomalies, suggestions);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  window.open(url, "_blank");
  // Nettoyer l'URL après 60s
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

/**
 * Télécharge le rapport HTML directement
 */
function downloadReport(bilan, anomalies, suggestions) {
  const html = generateReportHTML(bilan, anomalies, suggestions);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `Rapport_ESG_${bilan.annee}_${bilan.entreprise.replace(/\s+/g,"_")}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Export
window.ReportGenerator = { generateReportHTML, openReport, downloadReport };
