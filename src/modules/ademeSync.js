/**
 * ============================================================
 * ademeSync.js — Synchronisation Base Carbone ADEME
 *
 * Stratégie :
 *  1. Au démarrage, charge les facteurs depuis carbonFactors.js (local)
 *  2. Tente silencieusement de joindre l'API ADEME (data.ademe.fr)
 *  3. Si disponible : compare les versions, met à jour les facteurs
 *     modifiés en mémoire sans écraser le fichier local
 *  4. Affiche le statut dans l'UI (badge version + date)
 *
 * API utilisée : data.ademe.fr/data-fair (open data, sans clé)
 * Endpoint     : /api/v1/datasets/base-carboner/lines
 * ============================================================
 */

"use strict";

const ADEME_API = {
  base:      "https://data.ademe.fr/data-fair/api/v1/datasets/base-carboner",
  pageSize:  500,          // max par appel
  timeout:   8000,         // ms avant abandon
  cacheKey:  "ademe_sync_cache",
  cacheMaxAge: 7 * 24 * 60 * 60 * 1000, // 7 jours en ms
};

// Mapping colonnes API ADEME → structure interne
// (noms de colonnes Base Carbone V23.6)
const COL = {
  id:       "Identifiant_de_l_element",
  name:     "Nom_base_francais",
  type:     "Type_poste",           // "Poste" | "Élément"
  total:    "Total_poste_non_decompose",
  unit:     "Unité_français",
  category: "Code_de_la_categorie",
  status:   "Statut_de_l_element",  // "Valide générique" etc.
  comment:  "Commentaire_français",
  ghg:      "Type_GES",             // CO2, CH4, N2O…
};

// État de synchronisation
const SYNC_STATE = {
  status:       "idle",     // idle | checking | synced | offline | error
  apiVersion:   null,
  localVersion: null,
  updatedCount: 0,
  lastSync:     null,
  error:        null,
};

// Surcharges en mémoire (facteurs mis à jour depuis l'API)
let _overrides = {};

// ─── API publique du module ───────────────────────────────────────────────────

/**
 * Point d'entrée principal.
 * Lance la vérification en arrière-plan, sans bloquer l'UI.
 * @param {Function} onStatusChange - callback(state) appelé à chaque changement
 */
async function initSync(onStatusChange = () => {}) {
  SYNC_STATE.localVersion = CarbonFactors.ADEME_VERSION;
  SYNC_STATE.status = "checking";
  onStatusChange({ ...SYNC_STATE });

  // Tenter de charger depuis le cache localStorage d'abord
  const cached = loadCache();
  if (cached) {
    _overrides = cached.overrides || {};
    SYNC_STATE.status    = "synced";
    SYNC_STATE.lastSync  = cached.timestamp;
    SYNC_STATE.apiVersion = cached.apiVersion;
    SYNC_STATE.updatedCount = Object.keys(_overrides).length;
    onStatusChange({ ...SYNC_STATE });
    console.log(`[AdemeSync] Cache chargé — ${SYNC_STATE.updatedCount} facteurs en surcharge`);
  }

  // Tenter la sync API en arrière-plan
  try {
    await syncFromAPI(onStatusChange);
  } catch (e) {
    SYNC_STATE.status = "offline";
    SYNC_STATE.error  = e.message;
    onStatusChange({ ...SYNC_STATE });
    console.warn("[AdemeSync] API non joignable — utilisation de la base locale :", e.message);
  }
}

/**
 * Récupère un facteur : surcharge API si disponible, sinon base locale.
 * @param {string} key - clé du facteur (ex: "naturalGas")
 * @returns {{ factor, unit, label, category, scope, source } | null}
 */
function getFactor(key) {
  // 1. Surcharge API en mémoire
  if (_overrides[key]) return _overrides[key];
  // 2. Base locale
  return CarbonFactors.getFactor(key);
}

/**
 * Retourne tous les facteurs d'un scope, avec surcharges appliquées.
 */
function getFactorsByScope(scope) {
  const base = CarbonFactors.getFactorsByScope(scope);
  const result = { ...base };
  for (const [key, val] of Object.entries(_overrides)) {
    if (val.scope === scope) result[key] = val;
  }
  return result;
}

/**
 * Retourne l'état courant de la synchronisation.
 */
function getSyncState() {
  return { ...SYNC_STATE };
}

/**
 * Force une resynchronisation depuis l'API.
 */
async function forceSync(onStatusChange = () => {}) {
  clearCache();
  _overrides = {};
  await syncFromAPI(onStatusChange);
}

// ─── Synchronisation depuis l'API ────────────────────────────────────────────

async function syncFromAPI(onStatusChange) {
  // 1. Vérifier la version disponible
  const meta = await fetchWithTimeout(`${ADEME_API.base}`, {});
  if (!meta?.dataUpdatedAt) throw new Error("Métadonnées API invalides");

  const apiDate    = meta.dataUpdatedAt.substring(0, 10);
  const localDate  = CarbonFactors.ADEME_UPDATED;

  SYNC_STATE.apiVersion = apiDate;

  // Si le cache est plus récent que l'API → pas besoin de re-fetcher
  const cache = loadCache();
  if (cache && cache.apiDate === apiDate) {
    SYNC_STATE.status = "synced";
    onStatusChange({ ...SYNC_STATE });
    console.log("[AdemeSync] API à jour, cache valide.");
    return;
  }

  // 2. Si nouvelle version disponible → télécharger les facteurs critiques
  console.log(`[AdemeSync] Nouvelle version détectée : ${apiDate} (local: ${localDate})`);

  const newOverrides = {};
  let page = 0;
  let total = Infinity;

  while (page * ADEME_API.pageSize < total) {
    const url = `${ADEME_API.base}/lines?size=${ADEME_API.pageSize}&skip=${page * ADEME_API.pageSize}`
      + `&select=${encodeURIComponent(Object.values(COL).join(","))}`
      + `&qs=${encodeURIComponent('Statut_de_l_element:"Valide générique"')}`  // uniquement les facteurs validés
      + `&sort=Identifiant_de_l_element`;

    const data = await fetchWithTimeout(url, {});
    if (!data?.results) break;

    total = data.total || 0;
    processAPIPage(data.results, newOverrides);
    page++;

    // Limiter à 5000 facteurs max (éviter surcharge mémoire)
    if (Object.keys(newOverrides).length > 5000) break;
  }

  // 3. Mettre à jour les surcharges en mémoire
  _overrides = newOverrides;
  SYNC_STATE.status       = "synced";
  SYNC_STATE.updatedCount = Object.keys(newOverrides).length;
  SYNC_STATE.lastSync     = new Date().toISOString();
  SYNC_STATE.error        = null;

  // 4. Sauvegarder en cache
  saveCache({ overrides: newOverrides, apiVersion: apiDate, apiDate, timestamp: SYNC_STATE.lastSync });

  onStatusChange({ ...SYNC_STATE });
  console.log(`[AdemeSync] ✅ Sync terminée — ${SYNC_STATE.updatedCount} facteurs chargés depuis l'API`);
}

/**
 * Traite une page de résultats API et l'injecte dans newOverrides.
 */
function processAPIPage(results, target) {
  for (const row of results) {
    const id     = row[COL.id];
    const name   = row[COL.name];
    const rawUnit   = row[COL.unit] || "";
    const isKwh     = rawUnit.toLowerCase().includes("kwh") && !rawUnit.toLowerCase().includes("mwh");
    const total     = isKwh
      ? parseFloat(row[COL.total]) / 1000   // kWh → MWh
      : parseFloat(row[COL.total]);
    const unit      = normalizeUnit(rawUnit);
    const cat    = row[COL.category] || "";

    if (!id || isNaN(total) || total <= 0) continue;
    if (row[COL.type] !== "Poste") continue; // ignorer les éléments détaillés

    // Retrouver la clé interne correspondant à cet identifiant ADEME
    const internalKey = findKeyById(id);
    if (!internalKey) continue; // facteur non utilisé dans notre outil

    const existing = CarbonFactors.getFactor(internalKey);
    if (!existing) continue;

    // Mise à jour uniquement si valeur différente (arrondi 6 décimales)
    const diff = Math.abs(total - existing.factor) > 1e-6;
    if (diff) {
      target[internalKey] = {
        ...existing,
        factor:  total,
        unit:    unit || existing.unit,
        source:  "ADEME_API_" + (new Date().getFullYear()),
        updated: new Date().toISOString().substring(0, 10),
      };
    }
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Fetch avec timeout.
 */
async function fetchWithTimeout(url, opts = {}) {
  const controller = new AbortController();
  const tid = setTimeout(() => controller.abort(), ADEME_API.timeout);
  try {
    const resp = await fetch(url, { ...opts, signal: controller.signal });
    clearTimeout(tid);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    return await resp.json();
  } catch (e) {
    clearTimeout(tid);
    if (e.name === "AbortError") throw new Error("Timeout réseau");
    throw e;
  }
}

/**
 * Retrouve la clé interne (ex: "naturalGas") depuis un identifiant ADEME (ex: "31600").
 */
function findKeyById(ademeId) {
  const all = {
    ...CarbonFactors.SCOPE1_FACTORS,
    ...CarbonFactors.SCOPE2_FACTORS,
    ...CarbonFactors.SCOPE3_FACTORS,
  };
  for (const [key, val] of Object.entries(all)) {
    if (val.id === ademeId) return key;
  }
  return null;
}

/**
 * Normalise les noms d'unités ADEME vers les unités internes.
 */
function normalizeUnit(rawUnit) {
  if (!rawUnit) return null;
  const u = rawUnit.toLowerCase().trim();
  if (u.includes("mwh"))            return "MWh";
  if (u.includes("kwh"))            return "MWh"; // converti
  if (u.includes("t.km") || u.includes("tonne.km")) return "t.km";
  if (u.includes("tonne") || u === "t") return "t";
  if (u.includes("km"))             return "km";
  if (u === "kg")                   return "kg";
  if (u.includes("m³") || u.includes("m3")) return "m³";
  if (u.includes("litre") || u === "l")     return "L";
  return rawUnit;
}

// ─── Cache localStorage ───────────────────────────────────────────────────────

function saveCache(data) {
  try {
    localStorage.setItem(ADEME_API.cacheKey, JSON.stringify(data));
  } catch (e) {
    console.warn("[AdemeSync] Impossible de sauvegarder le cache :", e);
  }
}

function loadCache() {
  try {
    const raw = localStorage.getItem(ADEME_API.cacheKey);
    if (!raw) return null;
    const data = JSON.parse(raw);
    // Vérifier que le cache n'est pas trop vieux
    if (!data.timestamp) return null;
    const age = Date.now() - new Date(data.timestamp).getTime();
    if (age > ADEME_API.cacheMaxAge) { clearCache(); return null; }
    return data;
  } catch { return null; }
}

function clearCache() {
  try { localStorage.removeItem(ADEME_API.cacheKey); } catch { /* silencieux */ }
}

/**
 * Formate le statut pour affichage dans l'UI.
 */
function formatSyncStatus(state) {
  switch (state.status) {
    case "idle":     return { icon:"⏳", text:"En attente…",             color:"var(--esg-mist)" };
    case "checking": return { icon:"🔄", text:"Vérification ADEME…",     color:"var(--esg-mist)" };
    case "synced": {
      const d = state.lastSync ? new Date(state.lastSync).toLocaleDateString("fr-FR") : "—";
      const upd = state.updatedCount > 0
        ? ` · ${state.updatedCount} facteurs mis à jour`
        : " · Base locale à jour";
      return { icon:"✅", text:`ADEME ${state.apiVersion}${upd} (${d})`, color:"var(--esg-mint)" };
    }
    case "offline":  return { icon:"📴", text:"API hors ligne — base locale V23.6 utilisée", color:"var(--esg-warning)" };
    case "error":    return { icon:"⚠️", text:`Erreur sync : ${state.error}`, color:"var(--esg-danger)" };
    default:         return { icon:"❓", text:state.status, color:"var(--esg-mist)" };
  }
}

window.AdemeSync = {
  initSync,
  getFactor,
  getFactorsByScope,
  getSyncState,
  forceSync,
  formatSyncStatus,
};