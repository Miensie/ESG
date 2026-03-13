/**
 * ============================================================
 * geminiAI.js — ESG Analyzer Pro v2
 * Modèle BYOK (Bring Your Own Key) — Google AI Studio
 *
 * La clé API est :
 *  1. Saisie une seule fois par l'utilisateur
 *  2. Testée immédiatement (appel réel à Gemini)
 *  3. Sauvegardée dans localStorage (persiste entre sessions)
 *  4. Masquée dans l'interface après validation
 *  5. Jamais dans le code source
 * ============================================================
 */

"use strict";

const GEMINI_CONFIG = {
  model:     "gemini-2.5-flash-lite",
  apiBase:   "https://generativelanguage.googleapis.com/v1beta/models",
  maxTokens: 2048,
  storageKey: "esg_gemini_api_key",   // clé localStorage
};

// Clé en mémoire pour la session
let _apiKey = null;

// Historique de conversation du chat
let _chatHistory = [];

// ─── Persistance localStorage ─────────────────────────────────────────────────

/**
 * Charge la clé depuis localStorage si elle existe.
 * À appeler au démarrage de l'app.
 */
function loadSavedKey() {
  try {
    const saved = localStorage.getItem(GEMINI_CONFIG.storageKey);
    if (saved) {
      _apiKey = saved;
      return true;
    }
  } catch (e) {
    console.warn("[GeminiAI] localStorage non disponible :", e);
  }
  return false;
}

/**
 * Sauvegarde la clé dans localStorage après validation réussie.
 */
function saveKey(key) {
  _apiKey = key.trim();
  try {
    localStorage.setItem(GEMINI_CONFIG.storageKey, _apiKey);
  } catch (e) {
    console.warn("[GeminiAI] Impossible de sauvegarder dans localStorage :", e);
  }
}

/**
 * Efface la clé de la mémoire et de localStorage.
 */
function clearKey() {
  _apiKey = null;
  _chatHistory = [];
  try {
    localStorage.removeItem(GEMINI_CONFIG.storageKey);
  } catch (e) { /* silencieux */ }
}

function hasApiKey() {
  return !!_apiKey;
}

/**
 * Retourne les 8 premiers caractères pour affichage masqué.
 * Ex : "AIzaSyAB…" → "AIzaSyAB••••••••••••••"
 */
function getMaskedKey() {
  if (!_apiKey) return null;
  return _apiKey.substring(0, 8) + "•".repeat(16);
}

// ─── Test de connexion ────────────────────────────────────────────────────────

/**
 * Teste la clé avec un appel minimal à Gemini.
 * Lève une erreur explicite si la clé est invalide.
 */
async function testConnection(key) {
  const testKey = key.trim();
  if (!testKey) throw new Error("La clé API est vide.");
  if (!testKey.startsWith("AI")) throw new Error("Format invalide — la clé doit commencer par 'AI'.");

  const url = `${GEMINI_CONFIG.apiBase}/${GEMINI_CONFIG.model}:generateContent?key=${testKey}`;
  const body = {
    contents: [{ role: "user", parts: [{ text: "Réponds uniquement le mot : OK" }] }],
    generationConfig: { maxOutputTokens: 10, temperature: 0 },
  };

  let resp;
  try {
    resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  } catch (e) {
    throw new Error("Impossible de joindre l'API Gemini — vérifiez votre connexion internet.");
  }

  if (resp.status === 400) throw new Error("Clé API invalide ou malformée.");
  if (resp.status === 403) throw new Error("Clé API refusée — vérifiez qu'elle est bien activée sur AI Studio.");
  if (resp.status === 429) throw new Error("Quota API dépassé — réessayez dans quelques instants.");
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Erreur API (HTTP ${resp.status})`);
  }

  const data = await resp.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error("Réponse Gemini vide — réessayez.");

  return true;
}

// ─── Appel API principal ──────────────────────────────────────────────────────

async function callGemini(userMessage, history = [], withSystem = true) {
  if (!_apiKey) throw new Error("Clé API non configurée.");

  const url = `${GEMINI_CONFIG.apiBase}/${GEMINI_CONFIG.model}:generateContent?key=${_apiKey}`;

  const systemInstruction = withSystem ? {
    parts: [{ text: `Tu es un expert ESG senior et consultant en bilan carbone certifié GHG Protocol et Base Carbone ADEME.
Tu analyses des données d'émissions GES pour des entreprises industrielles françaises et internationales.
Tu maîtrises : Scope 1/2/3, ISO 14064, GHG Protocol Corporate Standard, CSRD, SBTi, RE2020, BEGES réglementaire.
Tu réponds toujours en français, de façon précise, structurée et directement actionnable.
Tu fournis des ordres de grandeur chiffrés et des références aux référentiels applicables.
Tu es professionnel et direct, sans formules de politesse superflues.` }]
  } : undefined;

  const body = {
    contents: [...history, { role: "user", parts: [{ text: userMessage }] }],
    generationConfig: { maxOutputTokens: GEMINI_CONFIG.maxTokens, temperature: 0.4, topP: 0.9 },
    ...(systemInstruction ? { systemInstruction } : {}),
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    const msg = err?.error?.message || `HTTP ${resp.status}`;
    if (resp.status === 403) throw new Error("Clé API révoquée ou quota épuisé. Vérifiez votre compte AI Studio.");
    if (resp.status === 429) throw new Error("Limite de requêtes atteinte — attendez quelques secondes.");
    throw new Error(`Erreur Gemini : ${msg}`);
  }

  const data = await resp.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error("Réponse Gemini vide ou bloquée par les filtres de sécurité.");
  return text;
}

// ─── Analyse complète du bilan ────────────────────────────────────────────────

async function analyzeBilan(bilan, historique = []) {
  if (!_apiKey) throw new Error("Clé API non configurée.");

  const bilanSummary = buildBilanSummary(bilan, historique);

  const prompt = `Analyse ce bilan carbone et réponds UNIQUEMENT avec le JSON suivant (pas de markdown, pas de backticks) :

${bilanSummary}

{
  "contexte": "2-3 phrases de synthèse du profil d'émissions",
  "anomalies": [
    {
      "severity": "error|warn|info",
      "scope": "Scope X ou Global",
      "titre": "Titre court",
      "message": "Explication détaillée avec chiffres et comparaisons sectorielles"
    }
  ],
  "suggestions": [
    {
      "priority": "high|medium|low",
      "scope": "Scope X",
      "action": "Titre de l'action",
      "detail": "Description concrète et chiffrée",
      "potentiel": "Ex: -25 à -40%",
      "delai": "Court terme (< 1 an) | Moyen terme (1-3 ans) | Long terme (> 3 ans)",
      "referentiel": "Ex: Base Carbone ADEME, SBTi, CSRD Art.8"
    }
  ],
  "tendance": {
    "commentaire": "Analyse de la tendance si données historiques disponibles, sinon chaîne vide",
    "trajectoire": "En hausse | En baisse | Stable | Données insuffisantes",
    "alerteCSRD": true,
    "alerteSBTi": "Explication si trajectoire non alignée -4.2%/an SBTi, sinon chaîne vide"
  }
}`;

  const responseText = await callGemini(prompt, [], true);

  try {
    const clean = responseText.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    return JSON.parse(clean);
  } catch (e) {
    console.warn("[GeminiAI] Échec parse JSON :", e);
    return {
      contexte: responseText,
      anomalies: [],
      suggestions: [],
      tendance: { commentaire: "", trajectoire: "Données insuffisantes", alerteCSRD: false, alerteSBTi: "" }
    };
  }
}

// ─── Chat interactif ──────────────────────────────────────────────────────────

async function sendChatMessage(userMessage, bilan) {
  if (!_apiKey) throw new Error("Clé API non configurée.");

  let messageToSend = userMessage;

  // Injecter le contexte bilan dans le premier message seulement
  if (_chatHistory.length === 0 && bilan) {
    messageToSend = `Voici le contexte du bilan carbone de l'entreprise (référence pour toute la conversation) :

${buildBilanSummary(bilan, [])}

---
Question : ${userMessage}`;
  }

  const responseText = await callGemini(messageToSend, _chatHistory, true);

  _chatHistory.push(
    { role: "user",  parts: [{ text: messageToSend }] },
    { role: "model", parts: [{ text: responseText }] }
  );

  // Garder max 20 échanges en mémoire
  if (_chatHistory.length > 40) _chatHistory = _chatHistory.slice(-40);

  return responseText;
}

function resetChat() { _chatHistory = []; }
function getChatLength() { return Math.floor(_chatHistory.length / 2); }

// ─── Helpers ──────────────────────────────────────────────────────────────────

function buildBilanSummary(bilan, historique = []) {
  const lines = [
    `ENTREPRISE : ${bilan.entreprise}`,
    `SECTEUR    : ${bilan.secteur || "Non précisé"}`,
    `ANNÉE      : ${bilan.annee}`,
    bilan.chiffreAffaires ? `CA         : ${(bilan.chiffreAffaires / 1e6).toFixed(1)} M€` : null,
    "",
    "BILAN CARBONE :",
    `  Scope 1 (émissions directes)  : ${bilan.scope1.total} tCO2eq (${bilan.scope1.pct}%)`,
    `  Scope 2 (énergie achetée)     : ${bilan.scope2.total} tCO2eq (${bilan.scope2.pct}%)`,
    `  Scope 3 (chaîne de valeur)    : ${bilan.scope3.total} tCO2eq (${bilan.scope3.pct}%)`,
    `  TOTAL                         : ${bilan.grandTotal} tCO2eq`,
    bilan.intensite ? `  Intensité carbone            : ${bilan.intensite} ${bilan.intensiteUnit}` : null,
    "",
    "DÉTAIL SCOPE 1 :",
    ...bilan.scope1.lines.map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
    "DÉTAIL SCOPE 2 :",
    ...bilan.scope2.lines.map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
    "DÉTAIL SCOPE 3 :",
    ...bilan.scope3.lines.map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
  ].filter(l => l !== null);

  if (historique?.length > 0) {
    lines.push("", "HISTORIQUE DES ÉMISSIONS :");
    historique.forEach(h => {
      const v = h.variation !== null ? ` (${h.variation > 0 ? "+" : ""}${h.variation}%)` : " (référence)";
      lines.push(`  ${h.annee} : ${h.total} tCO2eq${v}`);
    });
    if (historique.length >= 1) {
      const oldest = historique[0].total;
      const varTot  = (((bilan.grandTotal - oldest) / oldest) * 100).toFixed(1);
      lines.push(`  Variation ${historique[0].annee}→${bilan.annee} : ${varTot > 0 ? "+" : ""}${varTot}%`);
    }
  }

  return lines.join("\n");
}

function formatResponseHTML(text) {
  return text
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g,     "<em>$1</em>")
    .replace(/^#{1,3}\s+(.+)$/gm, "<strong style='color:var(--esg-mint)'>$1</strong>")
    .replace(/^[-•]\s+(.+)$/gm,   "<span style='display:block;padding-left:12px'>▸ $1</span>")
    .replace(/\n\n/g, "<br><br>")
    .replace(/\n/g,   "<br>");
}

// Export global
window.GeminiAI = {
  loadSavedKey,
  saveKey,
  clearKey,
  hasApiKey,
  getMaskedKey,
  testConnection,
  analyzeBilan,
  sendChatMessage,
  resetChat,
  getChatLength,
  formatResponseHTML,
  buildBilanSummary,
};
