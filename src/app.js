/**
 * ============================================================
 * app.js — Orchestrateur principal ESG Analyzer Pro
 * v2.0 — Historique réel + Analyse IA Gemini + Chat interactif
 * ============================================================
 */

"use strict";

const APP_STATE = {
  ready:        false,
  bilanCurrent: null,
  rawData:      null,
  anomalies:    [],
  suggestions:  [],
  historique:   [],   // [{ annee, total, scope1, scope2, scope3, variation }]
  charts:       {},
};

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    showToast("Ce complément nécessite Microsoft Excel.", "error"); return;
  }
  APP_STATE.ready = true;
  await initApp();
});

async function initApp() {
  setupNavigation();
  populateKeysTable();

  try {
    const name = await ExcelBridge.getWorkbookName();
    document.getElementById("workbook-name").textContent = name;
  } catch { document.getElementById("workbook-name").textContent = "Excel connecté"; }

  // ── Handlers principaux ───────────────────────────────────────────────────
  document.getElementById("btn-init").addEventListener("click", handleInit);
  document.getElementById("btn-demo").addEventListener("click", handleDemo);
  document.getElementById("btn-read-data").addEventListener("click", handleReadData);
  document.getElementById("btn-calc").addEventListener("click", handleCalcBilan);
  document.getElementById("btn-write-results").addEventListener("click", handleWriteResults);
  document.getElementById("btn-create-dashboard").addEventListener("click", handleCreateDashboard);
  document.getElementById("btn-add-historique").addEventListener("click", handleAddHistorique);
  document.getElementById("btn-clear-historique").addEventListener("click", handleClearHistorique);
  document.getElementById("btn-analyse-ia").addEventListener("click", handleAnalyseIA);
  document.getElementById("btn-chat-send").addEventListener("click", handleChatSend);
  document.getElementById("btn-chat-reset").addEventListener("click", handleChatReset);
  document.getElementById("chat-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleChatSend(); }
  });
  document.querySelectorAll(".chat-suggestion").forEach(btn => {
    btn.addEventListener("click", () => {
      document.getElementById("chat-input").value = btn.dataset.prompt;
      handleChatSend();
    });
  });
  document.getElementById("btn-open-report").addEventListener("click",     () => handleReport("open"));
  document.getElementById("btn-download-report").addEventListener("click", () => handleReport("download"));

  setStatus("ESG Analyzer Pro v2.0 prêt ✓");

  // ── Sync Base Carbone ADEME (arrière-plan, sans bloquer l'UI) ────────────
  if (typeof AdemeSync !== "undefined") {
  AdemeSync.initSync((state) => {
    const el = document.getElementById("ademe-sync-status");
    if (el) {
      const status = AdemeSync.formatSyncStatus(state);
      el.innerHTML = `${status.icon} ${status.text}`;
      el.style.color = status.color;
    }
    if (state.status === "synced") populateKeysTable();
  });
}
}

function setupNavigation() {
  document.querySelectorAll(".nav-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      const target = document.getElementById(tab.dataset.panel);
      if (target) target.classList.add("active");
    });
  });
}

function populateKeysTable() {
  // Utilise AdemeSync (API+cache) > CarbonFactors (base locale) en cascade
  const getFactors = (scope) => {
    if (typeof AdemeSync     !== "undefined") return AdemeSync.getFactorsByScope(scope);
    if (typeof CarbonFactors !== "undefined") return CarbonFactors.getFactorsByScope(scope);
    if (typeof CarbonCalc    !== "undefined" && CarbonCalc.getAvailableFactors)
      return CarbonCalc.getAvailableFactors(scope);
    return {};
  };

  ["scope1", "scope2", "scope3"].forEach(scope => {
    const tbody = document.getElementById(`keys-${scope}`);
    if (!tbody) return;
    tbody.innerHTML = "";
    const factors = getFactors(scope);
    // Grouper par catégorie pour un affichage clair
    const byCategory = {};
    Object.entries(factors).forEach(([key, ef]) => {
      const cat = ef.category || "Autre";
      if (!byCategory[cat]) byCategory[cat] = [];
      byCategory[cat].push({ key, ef });
    });
    Object.entries(byCategory).forEach(([cat, items]) => {
      // Ligne de catégorie
      const trCat = document.createElement("tr");
      trCat.innerHTML = `<td colspan="3" style="font-size:9px;font-weight:700;color:var(--esg-mist);
        padding:6px 4px 2px;text-transform:uppercase;letter-spacing:0.08em">${cat}</td>`;
      tbody.appendChild(trCat);
      // Lignes de facteurs
      items.forEach(({ key, ef }) => {
        const tr = document.createElement("tr");
        const isAPI = ef.source && ef.source.includes("API");
        tr.innerHTML = `
          <td style="font-family:var(--font-data);font-size:9px;color:var(--esg-mint)">${key}${isAPI ? ' <span title="Mis à jour depuis API ADEME" style="color:var(--esg-lime)">↑</span>' : ""}</td>
          <td style="font-size:10px">${ef.label}</td>
          <td style="font-family:var(--font-data);font-size:9px;color:var(--esg-mist)">${ef.unit}</td>`;
        tbody.appendChild(tr);
      });
    });
  });
}

// ── Handlers ──────────────────────────────────────────────────────────────────

async function handleInit() {
  setBtnLoading("btn-init", true, "Initialisation…");
  try {
    await ExcelBridge.initWorkbook();
    showToast("✅ Classeur ESG initialisé !", "success");
    log("collecte", "success", "Feuilles de collecte créées");
  } catch (e) { showToast("Erreur initialisation", "error"); log("collecte", "error", e.message); }
  setBtnLoading("btn-init", false, "🚀 Initialiser le classeur ESG");
}

function handleDemo() {
  APP_STATE.rawData = CarbonCalc.getDemoData();
  document.getElementById("cfg-entreprise").value = APP_STATE.rawData.entreprise;
  document.getElementById("cfg-annee").value       = APP_STATE.rawData.annee;
  document.getElementById("cfg-ca").value          = APP_STATE.rawData.chiffreAffaires;
  log("collecte", "success", `Données démo : ${APP_STATE.rawData.entreprise}`);
  showToast("🎯 Données démo chargées", "info");
}

async function handleReadData() {
  setBtnLoading("btn-read-data", true, "Lecture…");
  try {
    APP_STATE.rawData = await ExcelBridge.readESGData();
    const n = APP_STATE.rawData.scope1.length + APP_STATE.rawData.scope2.length + APP_STATE.rawData.scope3.length;
    log("collecte", "success", `${n} lignes lues`);
    showToast(`✅ ${n} lignes lues depuis Excel`, "success");
  } catch (e) { showToast("Erreur lecture Excel", "error"); log("collecte", "error", e.message); }
  setBtnLoading("btn-read-data", false, "📖 Lire les données depuis Excel");
}

async function handleCalcBilan() {
  if (!APP_STATE.rawData) { showToast("Chargez d'abord les données", "error"); return; }
  setBtnLoading("btn-calc", true, "Calcul en cours…");
  APP_STATE.rawData.entreprise      = document.getElementById("cfg-entreprise").value || APP_STATE.rawData.entreprise;
  APP_STATE.rawData.annee           = parseInt(document.getElementById("cfg-annee").value) || APP_STATE.rawData.annee;
  APP_STATE.rawData.chiffreAffaires = parseFloat(document.getElementById("cfg-ca").value) || APP_STATE.rawData.chiffreAffaires;
  try {
    const bilan = CarbonCalc.computeFullBilan(APP_STATE.rawData);
    APP_STATE.bilanCurrent = bilan;
    updateBilanUI(bilan);
    drawScopePieChart(bilan);
    drawSourcesBarChart(bilan);
    drawEvolutionChart(bilan);
    updateKPISummary(bilan);
    showToast(`✅ ${bilan.grandTotal.toLocaleString("fr-FR")} tCO2eq calculés`, "success");
    setStatus(`Bilan ${bilan.annee} calculé`);
  } catch (e) { showToast("Erreur calcul", "error"); console.error(e); }
  setBtnLoading("btn-calc", false, "♻️ Calculer le bilan carbone");
}

async function handleWriteResults() {
  if (!APP_STATE.bilanCurrent) { showToast("Calculez d'abord le bilan", "error"); return; }
  setBtnLoading("btn-write-results", true, "Écriture…");
  try {
    await ExcelBridge.writeResultats(APP_STATE.bilanCurrent);
    showToast("✅ Résultats écrits dans ESG_Resultats", "success");
  } catch (e) { showToast("Erreur écriture Excel", "error"); }
  setBtnLoading("btn-write-results", false, "💾 Écrire dans Excel");
}

async function handleCreateDashboard() {
  if (!APP_STATE.bilanCurrent) { showToast("Calculez d'abord le bilan", "error"); return; }
  setBtnLoading("btn-create-dashboard", true, "Création…");
  try {
    await ExcelBridge.createDashboard(APP_STATE.bilanCurrent);
    showToast("✅ Dashboard Excel créé !", "success");
  } catch (e) { showToast("Erreur dashboard", "error"); console.error(e); }
  setBtnLoading("btn-create-dashboard", false, "📊 Créer le dashboard Excel natif");
}

// ── Historique réel ───────────────────────────────────────────────────────────

function handleAddHistorique() {
  const annee  = parseInt(document.getElementById("hist-annee").value);
  const total  = parseFloat(document.getElementById("hist-total").value);
  const scope1 = parseFloat(document.getElementById("hist-s1").value) || 0;
  const scope2 = parseFloat(document.getElementById("hist-s2").value) || 0;
  const scope3 = parseFloat(document.getElementById("hist-s3").value) || 0;

  if (!annee || !total || isNaN(annee) || isNaN(total)) {
    showToast("Année et total sont requis", "error"); return;
  }
  if (APP_STATE.historique.find(h => h.annee === annee)) {
    showToast(`Année ${annee} déjà présente`, "error"); return;
  }

  APP_STATE.historique.push({ annee, total, scope1, scope2, scope3, variation: null });
  APP_STATE.historique.sort((a, b) => a.annee - b.annee);
  APP_STATE.historique.forEach((h, i) => {
    h.variation = i === 0 ? null
      : parseFloat(((h.total - APP_STATE.historique[i-1].total) / APP_STATE.historique[i-1].total * 100).toFixed(1));
  });

  ["hist-annee","hist-total","hist-s1","hist-s2","hist-s3"].forEach(id => {
    document.getElementById(id).value = "";
  });

  renderHistoriqueTable();
  if (APP_STATE.bilanCurrent) drawEvolutionChart(APP_STATE.bilanCurrent);
  showToast(`✅ Année ${annee} ajoutée`, "success");
}

function handleClearHistorique() {
  APP_STATE.historique = [];
  renderHistoriqueTable();
  if (APP_STATE.bilanCurrent) drawEvolutionChart(APP_STATE.bilanCurrent);
  showToast("Historique effacé", "info");
}

function renderHistoriqueTable() {
  const tbody = document.getElementById("historique-tbody");
  if (!tbody) return;
  if (!APP_STATE.historique.length) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;color:var(--esg-mist);font-size:10px;padding:12px">Aucune donnée — saisissez vos années précédentes ci-dessus</td></tr>`;
    return;
  }
  tbody.innerHTML = APP_STATE.historique.map(h => {
    const varHtml = h.variation === null ? `<span style="color:var(--esg-mist)">référence</span>`
      : h.variation > 0 ? `<span style="color:var(--esg-danger)">▲ +${h.variation}%</span>`
      : `<span style="color:var(--esg-mint)">▼ ${h.variation}%</span>`;
    return `<tr>
      <td style="font-family:var(--font-data);font-weight:600">${h.annee}</td>
      <td style="font-family:var(--font-data);text-align:right">${h.total.toLocaleString("fr-FR")}</td>
      <td style="text-align:center;font-size:11px">${varHtml}</td>
    </tr>`;
  }).join("");
}

// ── Analyse IA Gemini ─────────────────────────────────────────────────────────


async function handleAnalyseIA() {
  if (!APP_STATE.bilanCurrent) { showToast("Calculez d'abord le bilan", "error"); return; }
  if (!GeminiAI.hasApiKey())   { showToast("Configurez votre clé API Gemini", "error"); return; }

  setBtnLoading("btn-analyse-ia", true, "Analyse IA en cours…");
  document.getElementById("ia-result-zone").innerHTML = `
    <div style="display:flex;align-items:center;gap:8px;color:var(--esg-mist);font-size:12px;padding:16px">
      <span class="spinner"></span> Gemini analyse votre bilan carbone…
    </div>`;

  try {
    const result = await GeminiAI.analyzeBilan(APP_STATE.bilanCurrent, APP_STATE.historique);
    APP_STATE.anomalies   = result.anomalies   || [];
    APP_STATE.suggestions = result.suggestions || [];
    renderIAResult(result);
    showToast(`✅ ${APP_STATE.anomalies.length} anomalie(s) détectée(s) par l'IA`, "success");
    setStatus("Analyse Gemini terminée");
  } catch (e) {
    document.getElementById("ia-result-zone").innerHTML =
      `<div class="anomaly-item"><span class="anomaly-icon">❌</span><div class="anomaly-text"><strong>Erreur API Gemini</strong>${e.message}</div></div>`;
    showToast(e.message, "error");
  }
  setBtnLoading("btn-analyse-ia", false, "✨ Analyser avec Gemini IA");
}

function renderIAResult(result) {
  let html = "";

  if (result.contexte) {
    html += `<div class="card"><div class="card-title">🎯 Synthèse IA</div>
      <p style="font-size:12px;line-height:1.6">${GeminiAI.formatResponseHTML(result.contexte)}</p></div>`;
  }

  if (result.tendance?.commentaire) {
    const traj = result.tendance.trajectoire || "—";
    const trajColor = traj.includes("baisse") ? "var(--esg-mint)"
      : traj.includes("hausse") ? "var(--esg-danger)" : "var(--esg-warning)";
    html += `<div class="card"><div class="card-title">📈 Tendance</div>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <span style="font-family:var(--font-data);font-weight:700;color:${trajColor}">${traj}</span>
        ${result.tendance.alerteCSRD ? `<span style="background:rgba(245,166,35,0.2);color:var(--esg-warning);padding:2px 8px;border-radius:12px;font-size:9px;font-weight:700">⚠️ ALERTE CSRD</span>` : ""}
      </div>
      <p style="font-size:11px;color:var(--esg-mist)">${GeminiAI.formatResponseHTML(result.tendance.commentaire)}</p>
      ${result.tendance.alerteSBTi ? `<p style="font-size:11px;color:var(--esg-warning);margin-top:6px">🎯 SBTi : ${GeminiAI.formatResponseHTML(result.tendance.alerteSBTi)}</p>` : ""}
    </div>`;
  }

  if (result.anomalies?.length) {
    html += `<div class="section-title">⚠️ Anomalies</div>`;
    html += result.anomalies.map(a => `
      <div class="anomaly-item">
        <span class="anomaly-icon">${a.severity==="error"?"🔴":a.severity==="warn"?"🟡":"🔵"}</span>
        <div class="anomaly-text"><strong>${a.titre||a.scope}</strong>${GeminiAI.formatResponseHTML(a.message)}</div>
      </div>`).join("");
  } else {
    html += `<div class="suggestion-item" style="margin-bottom:8px">
      <span class="anomaly-icon">✅</span><span style="font-size:12px">Aucune anomalie critique.</span></div>`;
  }

  if (result.suggestions?.length) {
    html += `<div class="section-title" style="margin-top:var(--gap-md)">💡 Plan d'actions IA</div>`;
    const pIcon = { high:"🔥", medium:"💡", low:"💚" };
    html += result.suggestions.map(s => `
      <div class="suggestion-item" style="flex-direction:column;align-items:flex-start;gap:6px">
        <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
          <span>${pIcon[s.priority]||"💡"}</span>
          <span class="scope-badge s${(s.scope||"").slice(-1)||"3"}">${s.scope}</span>
          ${s.delai?`<span style="font-size:9px;color:var(--esg-mist)">${s.delai}</span>`:""}
          ${s.referentiel?`<span style="font-size:9px;color:var(--esg-mist);font-style:italic">${s.referentiel}</span>`:""}
        </div>
        <strong style="font-size:12px">${s.action}</strong>
        <span style="font-size:11px;color:var(--esg-mist)">${GeminiAI.formatResponseHTML(s.detail)}</span>
        <span style="color:var(--esg-mint);font-size:11px;font-weight:600">↘ ${s.potentiel}</span>
      </div>`).join("");
  }

  document.getElementById("ia-result-zone").innerHTML = html;
}

// ── Chat IA ───────────────────────────────────────────────────────────────────

async function handleChatSend() {
  const input = document.getElementById("chat-input");
  const msg   = input.value.trim();
  if (!msg) return;
  if (!GeminiAI.hasApiKey())   { showToast("Configurez votre clé API (onglet Analyse IA)", "error"); return; }
  if (!APP_STATE.bilanCurrent) { showToast("Calculez d'abord le bilan carbone", "error"); return; }

  input.value = "";
  appendChatMessage("user", msg);
  setChatLoading(true);
  try {
    const response = await GeminiAI.sendChatMessage(msg, APP_STATE.bilanCurrent);
    appendChatMessage("assistant", response);
    updateChatCounter();
  } catch (e) {
    appendChatMessage("error", `❌ ${e.message}`);
    showToast(e.message, "error");
  }
  setChatLoading(false);
}

function handleChatReset() {
  GeminiAI.resetChat();
  document.getElementById("chat-messages").innerHTML = `
    <div class="chat-msg chat-msg--assistant">
      <div class="chat-bubble">Conversation réinitialisée. Posez vos questions ESG.</div>
    </div>`;
  updateChatCounter();
}

function appendChatMessage(role, text) {
  const container = document.getElementById("chat-messages");
  const div = document.createElement("div");
  div.className = `chat-msg chat-msg--${role}`;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble";
  if (role === "assistant") {
    bubble.innerHTML = GeminiAI.formatResponseHTML(text);
  } else if (role === "error") {
    bubble.style.cssText = "background:rgba(224,82,82,0.15);border-color:var(--esg-danger)";
    bubble.textContent = text;
  } else {
    bubble.textContent = text;
  }
  div.appendChild(bubble);
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
}

function setChatLoading(loading) {
  document.getElementById("btn-chat-send").disabled  = loading;
  document.getElementById("chat-input").disabled     = loading;
  if (loading) {
    const container = document.getElementById("chat-messages");
    const div = document.createElement("div");
    div.className = "chat-msg chat-msg--assistant"; div.id = "chat-typing";
    div.innerHTML = `<div class="chat-bubble"><span class="spinner"></span></div>`;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  } else {
    document.getElementById("chat-typing")?.remove();
  }
}

function updateChatCounter() {
  const el = document.getElementById("chat-counter");
  if (el) el.textContent = `${GeminiAI.getChatLength()} échange(s)`;
}

// ── Rapport ───────────────────────────────────────────────────────────────────

function handleReport(action) {
  if (!APP_STATE.bilanCurrent) { showToast("Calculez d'abord le bilan", "error"); return; }
  if (action === "open") {
    ReportGenerator.openReport(APP_STATE.bilanCurrent, APP_STATE.anomalies, APP_STATE.suggestions);
  } else {
    ReportGenerator.downloadReport(APP_STATE.bilanCurrent, APP_STATE.anomalies, APP_STATE.suggestions);
  }
}

// ── UI helpers ────────────────────────────────────────────────────────────────

function updateBilanUI(bilan) {
  const fmt = v => v.toLocaleString("fr-FR");
  document.getElementById("kpi-s1").textContent    = fmt(bilan.scope1.total);
  document.getElementById("kpi-s2").textContent    = fmt(bilan.scope2.total);
  document.getElementById("kpi-s3").textContent    = fmt(bilan.scope3.total);
  document.getElementById("kpi-total").textContent = fmt(bilan.grandTotal);
  const max = Math.max(bilan.scope1.total, bilan.scope2.total, bilan.scope3.total);
  ["s1","s2","s3"].forEach((id, i) => {
    const val = [bilan.scope1.total, bilan.scope2.total, bilan.scope3.total][i];
    document.getElementById(`bar-${id}`)?.style.setProperty("width", `${Math.round(val/(max||1)*100)}%`);
  });
  document.getElementById("kpi-intensite").textContent  = `${bilan.intensite} ${bilan.intensiteUnit||""}`;
  document.getElementById("label-intensite").textContent = `Intensité (${bilan.intensiteUnit||"—"})`;
  const tbody = document.getElementById("table-bilan");
  tbody.innerHTML = "";
  const c = { scope1:"#E05252", scope2:"#F5A623", scope3:"#2ECC8E" };
  ["scope1","scope2","scope3"].forEach(scope => {
    bilan[scope].lines.forEach(line => {
      tbody.insertAdjacentHTML("beforeend", `<tr>
        <td style="font-size:10px">${line.source}</td>
        <td><span class="scope-badge s${scope.slice(-1)}">${scope.replace("scope","S")}</span></td>
        <td style="font-family:var(--font-data);text-align:right;color:${c[scope]}">${line.tCO2eq.toLocaleString("fr-FR",{minimumFractionDigits:2})}</td>
      </tr>`);
    });
  });
}

function updateKPISummary(bilan) {
  document.getElementById("kpi-summary").innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase">Entreprise</span><br><strong>${bilan.entreprise}</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase">Année</span><br><strong>${bilan.annee}</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase">Total</span><br><strong style="color:var(--esg-mint)">${bilan.grandTotal.toLocaleString("fr-FR")} tCO2eq</strong></div>
      <div><span style="color:var(--esg-mist);font-size:9px;text-transform:uppercase">Intensité</span><br><strong>${bilan.intensite} ${bilan.intensiteUnit||""}</strong></div>
    </div>`;
}

// ── Charts ────────────────────────────────────────────────────────────────────

function destroyChart(id) {
  if (APP_STATE.charts[id]) { APP_STATE.charts[id].destroy(); delete APP_STATE.charts[id]; }
}

function drawScopePieChart(bilan) {
  destroyChart("scopes");
  const canvas = document.getElementById("chart-scopes"); if (!canvas) return;
  APP_STATE.charts["scopes"] = new Chart(canvas, {
    type: "doughnut",
    data: { labels: [`S1 (${bilan.scope1.pct}%)`,`S2 (${bilan.scope2.pct}%)`,`S3 (${bilan.scope3.pct}%)`],
      datasets: [{ data: [bilan.scope1.total, bilan.scope2.total, bilan.scope3.total],
        backgroundColor: ["#E05252","#F5A623","#2ECC8E"], borderColor: "#1C2B3A", borderWidth:2, hoverOffset:8 }] },
    options: { responsive:true, maintainAspectRatio:false, cutout:"60%",
      plugins: { legend:{ position:"bottom", labels:{ color:"#8BA5B8", font:{size:10}, padding:12 } },
        tooltip:{ callbacks:{ label: ctx => ` ${ctx.label} : ${ctx.raw.toLocaleString("fr-FR")} tCO2eq` } } } }
  });
}

function drawSourcesBarChart(bilan) {
  destroyChart("sources");
  const canvas = document.getElementById("chart-sources"); if (!canvas) return;
  const allLines = [...bilan.scope1.lines.map(l=>({...l,scope:1})), ...bilan.scope2.lines.map(l=>({...l,scope:2})),
    ...bilan.scope3.lines.map(l=>({...l,scope:3}))].sort((a,b)=>b.tCO2eq-a.tCO2eq).slice(0,12);
  const colors = {1:"#E05252",2:"#F5A623",3:"#2ECC8E"};
  APP_STATE.charts["sources"] = new Chart(canvas, {
    type:"bar", data:{ labels: allLines.map(l=>l.source.substring(0,18)),
      datasets:[{ data: allLines.map(l=>l.tCO2eq), backgroundColor: allLines.map(l=>colors[l.scope]+"CC"),
        borderColor: allLines.map(l=>colors[l.scope]), borderWidth:1, borderRadius:4 }] },
    options:{ indexAxis:"y", responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
      scales:{ x:{ticks:{color:"#8BA5B8",font:{size:9}},grid:{color:"rgba(255,255,255,0.05)"}},
               y:{ticks:{color:"#8BA5B8",font:{size:9}},grid:{display:false}} } }
  });
}

/**
 * Graphique d'évolution — 100% données réelles depuis APP_STATE.historique
 * Affiche un message si pas encore d'historique saisi.
 */
function drawEvolutionChart(bilan) {
  destroyChart("evolution");
  const canvas = document.getElementById("chart-evolution"); if (!canvas) return;

  const allPoints = [
    ...APP_STATE.historique.map(h => ({ annee:h.annee, total:h.total, s1:h.scope1, s2:h.scope2, s3:h.scope3 })),
    { annee:bilan.annee, total:bilan.grandTotal, s1:bilan.scope1.total, s2:bilan.scope2.total, s3:bilan.scope3.total }
  ].filter((p,i,a) => a.findIndex(x=>x.annee===p.annee)===i).sort((a,b)=>a.annee-b.annee);

  if (allPoints.length === 1) {
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0,0,canvas.width,canvas.height);
    ctx.fillStyle="#8BA5B8"; ctx.font="11px sans-serif"; ctx.textAlign="center";
    ctx.fillText("Saisissez vos années précédentes dans l'onglet Dashboard", canvas.width/2, canvas.height/2-8);
    ctx.fillText("pour afficher la vraie courbe d'évolution.", canvas.width/2, canvas.height/2+10);
    return;
  }

  APP_STATE.charts["evolution"] = new Chart(canvas, {
    type:"line",
    data:{ labels: allPoints.map(p=>String(p.annee)),
      datasets:[
        { label:"Total", data:allPoints.map(p=>p.total), borderColor:"#2ECC8E", backgroundColor:"rgba(46,204,142,0.08)",
          fill:true, tension:0.3, pointBackgroundColor:"#2ECC8E", pointRadius:5, borderWidth:2 },
        { label:"Scope 1", data:allPoints.map(p=>p.s1), borderColor:"#E05252", tension:0.3,
          pointBackgroundColor:"#E05252", pointRadius:3, borderWidth:1.5, borderDash:[4,3], backgroundColor:"transparent" },
        { label:"Scope 2", data:allPoints.map(p=>p.s2), borderColor:"#F5A623", tension:0.3,
          pointBackgroundColor:"#F5A623", pointRadius:3, borderWidth:1.5, borderDash:[4,3], backgroundColor:"transparent" },
        { label:"Scope 3", data:allPoints.map(p=>p.s3), borderColor:"#1D8348", tension:0.3,
          pointBackgroundColor:"#1D8348", pointRadius:3, borderWidth:1.5, borderDash:[4,3], backgroundColor:"transparent" },
      ]
    },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:"#8BA5B8", font:{size:9}, boxWidth:12 } },
        tooltip:{ callbacks:{ label: ctx => ` ${ctx.dataset.label} : ${ctx.raw.toLocaleString("fr-FR")} tCO2eq` } } },
      scales:{ x:{ticks:{color:"#8BA5B8",font:{size:10}},grid:{color:"rgba(255,255,255,0.05)"}},
               y:{ticks:{color:"#8BA5B8",font:{size:10}},grid:{color:"rgba(255,255,255,0.05)"}} } }
  });
}

// ── Utilitaires ───────────────────────────────────────────────────────────────

function showToast(message, type="info", duration=3500) {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${{success:"✅",error:"❌",info:"ℹ️"}[type]||"ℹ️"}</span><span>${message}</span>`;
  document.getElementById("toast-container").appendChild(toast);
  setTimeout(() => { toast.style.opacity="0"; toast.style.transform="translateY(8px)";
    toast.style.transition="all 0.3s ease"; setTimeout(()=>toast.remove(),300); }, duration);
}

function log(panelId, type, message) {
  const logEl = document.getElementById(`log-${panelId}`); if (!logEl) return;
  const entry = document.createElement("div");
  entry.className = `log-entry ${type}`;
  entry.innerHTML = `<span class="log-ts">${new Date().toLocaleTimeString("fr-FR")}</span>${message}`;
  logEl.appendChild(entry); logEl.scrollTop = logEl.scrollHeight;
}

function setBtnLoading(id, loading, label) {
  const btn = document.getElementById(id); if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading ? `<span class="spinner"></span> ${label}` : label;
}

function setStatus(msg) {
  const el = document.getElementById("footer-status"); if (el) el.textContent = msg;
}