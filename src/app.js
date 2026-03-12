/**
 * ============================================================
 * app.js — Orchestrateur principal ESG Analyzer Pro
 * Initialise Office.js, gère la navigation et connecte
 * tous les modules entre eux.
 * ============================================================
 */

"use strict";

// ─── État global de l'application ────────────────────────────────────────────
const APP_STATE = {
  ready:       false,
  bilanCurrent: null,
  rawData:     null,
  anomalies:   [],
  suggestions: [],
  charts:      {},  // Instances Chart.js pour destruction/recréation
};

// ─── Point d'entrée : attendre Office.js ─────────────────────────────────────
Office.onReady(async (info) => {
  console.log("[App] Office.js prêt :", info.host, info.platform);

  // Vérifier qu'on est bien dans Excel
  if (info.host !== Office.HostType.Excel) {
    showToast("Ce complément nécessite Microsoft Excel.", "error");
    return;
  }

  APP_STATE.ready = true;
  await initApp();
});

// ─── Initialisation de l'interface ───────────────────────────────────────────
async function initApp() {
  // Navigation par onglets
  setupNavigation();

  // Remplir les tableaux de clés d'émission
  populateKeysTable();

  // Récupérer le nom du classeur
  try {
    const name = await ExcelBridge.getWorkbookName();
    document.getElementById("workbook-name").textContent = name;
  } catch {
    document.getElementById("workbook-name").textContent = "Excel connecté";
  }

  // Bouton initialisation classeur
  document.getElementById("btn-init").addEventListener("click", handleInit);

  // Bouton démo
  document.getElementById("btn-demo").addEventListener("click", handleDemo);

  // Bouton lecture données
  document.getElementById("btn-read-data").addEventListener("click", handleReadData);

  // Bouton calcul bilan
  document.getElementById("btn-calc").addEventListener("click", handleCalcBilan);

  // Bouton écriture résultats
  document.getElementById("btn-write-results").addEventListener("click", handleWriteResults);

  // Bouton dashboard Excel natif
  document.getElementById("btn-create-dashboard").addEventListener("click", handleCreateDashboard);

  // Bouton analyse
  document.getElementById("btn-analyse").addEventListener("click", handleAnalyse);

  // Boutons rapport
  document.getElementById("btn-open-report").addEventListener("click", () => handleReport("open"));
  document.getElementById("btn-download-report").addEventListener("click", () => handleReport("download"));

  setStatus("ESG Analyzer Pro prêt ✓");
  console.log("[App] Initialisation terminée");
}

// ─── Navigation par onglets ───────────────────────────────────────────────────
function setupNavigation() {
  const tabs   = document.querySelectorAll(".nav-tab");
  const panels = document.querySelectorAll(".panel");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      const target = document.getElementById(tab.dataset.panel);
      if (target) target.classList.add("active");
    });
  });
}

// ─── Popule les tables de clés d'émission ────────────────────────────────────
function populateKeysTable() {
  const EF = CarbonCalc.EMISSION_FACTORS;
  ["scope1", "scope2", "scope3"].forEach(scope => {
    const tbody = document.getElementById(`keys-${scope}`);
    if (!tbody) return;
    Object.entries(EF[scope]).forEach(([key, ef]) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td style="font-family:var(--font-data);font-size:10px;color:var(--esg-mint)">${key}</td>
        <td style="font-size:10px">${ef.label}</td>
        <td style="font-family:var(--font-data);font-size:10px;color:var(--esg-mist)">${ef.unit}</td>
      `;
      tbody.appendChild(tr);
    });
  });
}

// ─── Handlers des actions ──────────────────────────────────────────────────────

/** 1. Initialiser le classeur Excel */
async function handleInit() {
  setBtnLoading("btn-init", true, "Initialisation…");
  try {
    await ExcelBridge.initWorkbook();
    showToast("✅ Classeur ESG initialisé — feuilles créées !", "success");
    log("collecte", "success", "Classeur initialisé avec les feuilles de collecte");
    setStatus("Classeur initialisé");
  } catch (e) {
    showToast("Erreur lors de l'initialisation", "error");
    log("collecte", "error", e.message);
  }
  setBtnLoading("btn-init", false, "🚀 Initialiser le classeur ESG");
}

/** 2. Charger les données démo */
function handleDemo() {
  APP_STATE.rawData = CarbonCalc.getDemoData();

  // Pré-remplir la config
  document.getElementById("cfg-entreprise").value = APP_STATE.rawData.entreprise;
  document.getElementById("cfg-annee").value       = APP_STATE.rawData.annee;
  document.getElementById("cfg-ca").value          = APP_STATE.rawData.chiffreAffaires;

  log("collecte", "success", `Données démo chargées : ${APP_STATE.rawData.entreprise}`);
  log("collecte", "info", `Scope 1 : ${APP_STATE.rawData.scope1.length} sources`);
  log("collecte", "info", `Scope 2 : ${APP_STATE.rawData.scope2.length} sources`);
  log("collecte", "info", `Scope 3 : ${APP_STATE.rawData.scope3.length} sources`);
  showToast("🎯 Données démo chargées — Calculez le bilan !", "info");
  setStatus("Données démo en mémoire");
}

/** 3. Lire les données depuis Excel */
async function handleReadData() {
  setBtnLoading("btn-read-data", true, "Lecture…");
  try {
    APP_STATE.rawData = await ExcelBridge.readESGData();
    const total = APP_STATE.rawData.scope1.length
                + APP_STATE.rawData.scope2.length
                + APP_STATE.rawData.scope3.length;
    log("collecte", "success", `${total} lignes lues depuis Excel`);
    log("collecte", "info", `Entreprise : ${APP_STATE.rawData.entreprise || "—"}`);
    log("collecte", "info", `Scope 1 : ${APP_STATE.rawData.scope1.length} | Scope 2 : ${APP_STATE.rawData.scope2.length} | Scope 3 : ${APP_STATE.rawData.scope3.length}`);
    showToast(`✅ ${total} lignes lues depuis Excel`, "success");
    setStatus(`${total} lignes ESG collectées`);
  } catch (e) {
    showToast("Erreur lecture Excel", "error");
    log("collecte", "error", e.message);
  }
  setBtnLoading("btn-read-data", false, "📖 Lire les données depuis Excel");
}

/** 4. Calculer le bilan carbone */
async function handleCalcBilan() {
  if (!APP_STATE.rawData) {
    showToast("Chargez d'abord les données (Collecte)", "error");
    return;
  }

  setBtnLoading("btn-calc", true, "Calcul en cours…");

  // Enrichir avec la config UI si modifiée
  APP_STATE.rawData.entreprise      = document.getElementById("cfg-entreprise").value || APP_STATE.rawData.entreprise;
  APP_STATE.rawData.annee           = parseInt(document.getElementById("cfg-annee").value) || APP_STATE.rawData.annee;
  APP_STATE.rawData.chiffreAffaires = parseFloat(document.getElementById("cfg-ca").value) || APP_STATE.rawData.chiffreAffaires;

  try {
    const bilan = CarbonCalc.computeFullBilan(APP_STATE.rawData);
    APP_STATE.bilanCurrent = bilan;

    // Mettre à jour les KPI
    updateBilanUI(bilan);

    // Dessiner les graphiques inline
    drawScopePieChart(bilan);
    drawSourcesBarChart(bilan);
    drawEvolutionChart(bilan);

    // Mettre à jour le dashboard KPI summary
    updateKPISummary(bilan);

    showToast(`✅ Bilan calculé : ${bilan.grandTotal.toLocaleString("fr-FR")} tCO2eq`, "success");
    setStatus(`Bilan ${bilan.annee} calculé`);
  } catch (e) {
    showToast("Erreur calcul bilan", "error");
    console.error(e);
  }

  setBtnLoading("btn-calc", false, "♻️ Calculer le bilan carbone");
}

/** 5. Écrire les résultats dans Excel */
async function handleWriteResults() {
  if (!APP_STATE.bilanCurrent) {
    showToast("Calculez d'abord le bilan carbone", "error");
    return;
  }
  setBtnLoading("btn-write-results", true, "Écriture…");
  try {
    await ExcelBridge.writeResultats(APP_STATE.bilanCurrent);
    showToast("✅ Résultats écrits dans ESG_Resultats", "success");
  } catch (e) {
    showToast("Erreur écriture Excel", "error");
  }
  setBtnLoading("btn-write-results", false, "💾 Écrire dans Excel");
}

/** 6. Créer le dashboard Excel natif */
async function handleCreateDashboard() {
  if (!APP_STATE.bilanCurrent) {
    showToast("Calculez d'abord le bilan carbone", "error");
    return;
  }
  setBtnLoading("btn-create-dashboard", true, "Création…");
  try {
    await ExcelBridge.createDashboard(APP_STATE.bilanCurrent);
    showToast("✅ Dashboard Excel créé avec graphiques natifs !", "success");
  } catch (e) {
    showToast("Erreur création dashboard", "error");
    console.error(e);
  }
  setBtnLoading("btn-create-dashboard", false, "📊 Créer le dashboard Excel natif");
}

/** 7. Analyse anomalies + suggestions */
function handleAnalyse() {
  if (!APP_STATE.bilanCurrent) {
    showToast("Calculez d'abord le bilan carbone", "error");
    return;
  }

  APP_STATE.anomalies   = CarbonCalc.detectAnomalies(APP_STATE.bilanCurrent);
  APP_STATE.suggestions = CarbonCalc.generateSuggestions(APP_STATE.bilanCurrent);

  updateAnomaliesUI(APP_STATE.anomalies);
  updateSuggestionsUI(APP_STATE.suggestions);

  const n = APP_STATE.anomalies.length;
  showToast(`🔍 Analyse terminée : ${n} anomalie(s) détectée(s)`, n > 0 ? "info" : "success");
}

/** 8. Rapport ESG */
function handleReport(action) {
  if (!APP_STATE.bilanCurrent) {
    showToast("Calculez d'abord le bilan carbone", "error");
    return;
  }
  if (action === "open") {
    ReportGenerator.openReport(
      APP_STATE.bilanCurrent,
      APP_STATE.anomalies,
      APP_STATE.suggestions
    );
    showToast("📄 Rapport ouvert dans un nouvel onglet", "info");
  } else {
    ReportGenerator.downloadReport(
      APP_STATE.bilanCurrent,
      APP_STATE.anomalies,
      APP_STATE.suggestions
    );
    showToast("⬇️ Rapport téléchargé", "success");
  }
}

// ─── Mise à jour UI Bilan ─────────────────────────────────────────────────────
function updateBilanUI(bilan) {
  const fmt = v => v.toLocaleString("fr-FR", { minimumFractionDigits: 0 });
  document.getElementById("kpi-s1").textContent = fmt(bilan.scope1.total);
  document.getElementById("kpi-s2").textContent = fmt(bilan.scope2.total);
  document.getElementById("kpi-s3").textContent = fmt(bilan.scope3.total);
  document.getElementById("kpi-total").textContent = fmt(bilan.grandTotal);

  // Progress bars
  const max = Math.max(bilan.scope1.total, bilan.scope2.total, bilan.scope3.total);
  ["s1","s2","s3"].forEach((id, i) => {
    const val = [bilan.scope1.total, bilan.scope2.total, bilan.scope3.total][i];
    const bar = document.getElementById(`bar-${id}`);
    if (bar) bar.style.width = `${Math.round((val / (max||1)) * 100)}%`;
  });

  // Intensité
  document.getElementById("kpi-intensite").textContent =
    `${bilan.intensite} ${bilan.intensiteUnit || ""}`;
  document.getElementById("label-intensite").textContent =
    `Intensité carbone (${bilan.intensiteUnit || "non calculée"})`;

  // Table détail
  const tbody = document.getElementById("table-bilan");
  tbody.innerHTML = "";
  const scopeColor = { scope1: "#E05252", scope2: "#F5A623", scope3: "#2ECC8E" };
  ["scope1","scope2","scope3"].forEach(scope => {
    bilan[scope].lines.forEach(line => {
      const tr = document.createElement("tr");
      const badge = scope.replace("scope","S");
      tr.innerHTML = `
        <td style="font-size:10px">${line.source}</td>
        <td><span class="scope-badge s${scope.slice(-1)}">${badge}</span></td>
        <td style="font-family:var(--font-data);text-align:right;color:${scopeColor[scope]}">${line.tCO2eq.toLocaleString("fr-FR",{minimumFractionDigits:2})}</td>
      `;
      tbody.appendChild(tr);
    });
  });
}

function updateKPISummary(bilan) {
  const el = document.getElementById("kpi-summary");
  el.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase;letter-spacing:.08em">Entreprise</span><br><strong>${bilan.entreprise}</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase;letter-spacing:.08em">Année</span><br><strong>${bilan.annee}</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase;letter-spacing:.08em">Total émissions</span><br><strong style="color:var(--esg-mint)">${bilan.grandTotal.toLocaleString("fr-FR")} tCO2eq</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase;letter-spacing:.08em">Intensité</span><br><strong>${bilan.intensite} ${bilan.intensiteUnit||""}</strong></div>
    </div>
  `;
}

// ─── Graphiques Chart.js ──────────────────────────────────────────────────────
function destroyChart(id) {
  if (APP_STATE.charts[id]) {
    APP_STATE.charts[id].destroy();
    delete APP_STATE.charts[id];
  }
}

function drawScopePieChart(bilan) {
  destroyChart("scopes");
  const canvas = document.getElementById("chart-scopes");
  if (!canvas) return;
  APP_STATE.charts["scopes"] = new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: [`Scope 1 (${bilan.scope1.pct}%)`, `Scope 2 (${bilan.scope2.pct}%)`, `Scope 3 (${bilan.scope3.pct}%)`],
      datasets: [{
        data: [bilan.scope1.total, bilan.scope2.total, bilan.scope3.total],
        backgroundColor: ["#E05252","#F5A623","#2ECC8E"],
        borderColor: "#1C2B3A",
        borderWidth: 2,
        hoverOffset: 8,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: "60%",
      plugins: {
        legend: { position: "bottom", labels: { color: "#8BA5B8", font: { size: 10 }, padding: 12 } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label} : ${ctx.raw.toLocaleString("fr-FR")} tCO2eq` } }
      }
    }
  });
}

function drawSourcesBarChart(bilan) {
  destroyChart("sources");
  const canvas = document.getElementById("chart-sources");
  if (!canvas) return;

  const allLines = [
    ...bilan.scope1.lines.map(l => ({...l, scope: 1})),
    ...bilan.scope2.lines.map(l => ({...l, scope: 2})),
    ...bilan.scope3.lines.map(l => ({...l, scope: 3})),
  ].sort((a,b) => b.tCO2eq - a.tCO2eq).slice(0, 12);

  const colors = { 1: "#E05252", 2: "#F5A623", 3: "#2ECC8E" };
  APP_STATE.charts["sources"] = new Chart(canvas, {
    type: "bar",
    data: {
      labels: allLines.map(l => l.source.substring(0, 18)),
      datasets: [{
        data: allLines.map(l => l.tCO2eq),
        backgroundColor: allLines.map(l => colors[l.scope] + "CC"),
        borderColor: allLines.map(l => colors[l.scope]),
        borderWidth: 1,
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#8BA5B8", font: { size: 9 } }, grid: { color: "rgba(255,255,255,0.05)" } },
        y: { ticks: { color: "#8BA5B8", font: { size: 9 } }, grid: { display: false } }
      }
    }
  });
}

function drawEvolutionChart(bilan) {
  destroyChart("evolution");
  const canvas = document.getElementById("chart-evolution");
  if (!canvas) return;

  // Simulation d'évolution N-3 → N (objectif -4%/an)
  const years = [bilan.annee - 3, bilan.annee - 2, bilan.annee - 1, bilan.annee];
  const coeff = [1.126, 1.071, 1.034, 1.0];
  const totals = coeff.map(c => Math.round(bilan.grandTotal * c));

  APP_STATE.charts["evolution"] = new Chart(canvas, {
    type: "line",
    data: {
      labels: years.map(String),
      datasets: [{
        label: "Total (tCO2eq)",
        data: totals,
        borderColor: "#2ECC8E",
        backgroundColor: "rgba(46,204,142,0.1)",
        fill: true,
        tension: 0.4,
        pointBackgroundColor: "#2ECC8E",
        pointRadius: 5,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: "#8BA5B8", font: { size: 10 } } } },
      scales: {
        x: { ticks: { color: "#8BA5B8", font: { size: 10 } }, grid: { color: "rgba(255,255,255,0.05)" } },
        y: { ticks: { color: "#8BA5B8", font: { size: 10 } }, grid: { color: "rgba(255,255,255,0.05)" } }
      }
    }
  });
}

// ─── Mise à jour UI Analyse ───────────────────────────────────────────────────
function updateAnomaliesUI(anomalies) {
  const el = document.getElementById("anomalies-list");
  if (!anomalies.length) {
    el.innerHTML = `<div class="suggestion-item"><span class="anomaly-icon">✅</span><span>Aucune anomalie détectée</span></div>`;
    return;
  }
  el.innerHTML = anomalies.map(a => `
    <div class="anomaly-item">
      <span class="anomaly-icon">${a.severity === "error" ? "🔴" : a.severity === "warn" ? "🟡" : "🔵"}</span>
      <div class="anomaly-text">
        <strong>${a.scope}</strong>
        ${a.message}
      </div>
    </div>
  `).join("");
}

function updateSuggestionsUI(suggestions) {
  const priorityIcon = { high: "🔥", medium: "💡", low: "💚" };
  const el = document.getElementById("suggestions-list");
  el.innerHTML = suggestions.map(s => `
    <div class="suggestion-item">
      <span class="anomaly-icon">${priorityIcon[s.priority]}</span>
      <div>
        <span class="scope-badge s${s.scope.slice(-1) || "3"}" style="margin-bottom:4px;display:inline-block">${s.scope}</span>
        <strong style="display:block;font-size:12px">${s.action}</strong>
        <span style="color:var(--esg-mist);font-size:10px">${s.detail}</span><br>
        <span style="color:var(--esg-mint);font-size:10px;font-weight:600">↘ ${s.potentiel}</span>
      </div>
    </div>
  `).join("");
}

// ─── Utilitaires UI ──────────────────────────────────────────────────────────

/** Toast notification */
function showToast(message, type = "info", duration = 3500) {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  const icons = { success: "✅", error: "❌", info: "ℹ️" };
  toast.innerHTML = `<span>${icons[type] || "ℹ️"}</span><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateY(8px)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

/** Log collecte */
function log(panelId, type, message) {
  const logEl = document.getElementById(`log-${panelId}`);
  if (!logEl) return;
  const ts = new Date().toLocaleTimeString("fr-FR");
  const entry = document.createElement("div");
  entry.className = `log-entry ${type}`;
  entry.innerHTML = `<span class="log-ts">${ts}</span>${message}`;
  logEl.appendChild(entry);
  logEl.scrollTop = logEl.scrollHeight;
}

/** Bouton loading state */
function setBtnLoading(id, loading, label) {
  const btn = document.getElementById(id);
  if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading
    ? `<span class="spinner"></span> ${label}`
    : label;
}

/** Status footer */
function setStatus(msg) {
  const el = document.getElementById("footer-status");
  if (el) el.textContent = msg;
}
