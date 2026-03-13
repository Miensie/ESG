/**
 * ============================================================
 * ESG Analyzer Pro — Backend Node.js
 * Proxy sécurisé Gemini AI
 *
 * La clé API Gemini est stockée dans .env (variable d'env).
 * Les clients appellent /api/ai/* — ils ne voient jamais la clé.
 *
 * DÉPLOIEMENT RECOMMANDÉ :
 *   - Railway.app  (gratuit jusqu'à 500h/mois)
 *   - Render.com   (gratuit plan starter)
 *   - Fly.io       (gratuit plan hobby)
 *   - VPS OVH/Hetzner (~4€/mois)
 * ============================================================
 */

require("dotenv").config();

const express  = require("express");
const cors     = require("cors");
const https    = require("https");
const http     = require("http");
const fs       = require("fs");
const path     = require("path");
const rateLimit = require("express-rate-limit");

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── Variables d'environnement obligatoires ───────────────────────────────────
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "https://miensie.github.io";
const API_SECRET     = process.env.API_SECRET;      // Token optionnel anti-abus

if (!GEMINI_API_KEY) {
  console.error("❌ GEMINI_API_KEY manquante dans .env — serveur arrêté.");
  process.exit(1);
}

// ─── Middleware ───────────────────────────────────────────────────────────────

// CORS : autoriser uniquement votre domaine GitHub Pages + localhost dev
app.use(cors({
  origin: (origin, cb) => {
    const allowed = [
      ALLOWED_ORIGIN,
      "http://localhost:3000",
      "https://localhost:3000",
    ];
    // Autoriser les requêtes sans origin (Postman, tests)
    if (!origin || allowed.some(o => origin.startsWith(o))) {
      cb(null, true);
    } else {
      cb(new Error(`CORS bloqué pour l'origine : ${origin}`));
    }
  },
  methods: ["GET", "POST"],
  allowedHeaders: ["Content-Type", "X-API-Secret"],
}));

app.use(express.json({ limit: "50kb" }));

// ─── Rate limiting — protection anti-abus ────────────────────────────────────
// Max 30 appels IA par IP par heure (ajustable)
const aiLimiter = rateLimit({
  windowMs: 60 * 60 * 1000,   // 1 heure
  max: 30,
  message: { error: "Trop de requêtes — réessayez dans une heure." },
  standardHeaders: true,
  legacyHeaders: false,
});

// Max 5 appels par seconde (protection burst)
const burstLimiter = rateLimit({
  windowMs: 1000,
  max: 5,
  message: { error: "Trop de requêtes simultanées." },
});

// ─── Middleware optionnel : vérification token secret ────────────────────────
function checkSecret(req, res, next) {
  // Si API_SECRET est défini dans .env, le client doit l'envoyer en header
  // Cela permet de bloquer les appels qui ne viennent pas de votre add-in
  if (API_SECRET) {
    const token = req.headers["x-api-secret"];
    if (token !== API_SECRET) {
      return res.status(401).json({ error: "Non autorisé." });
    }
  }
  next();
}

// ─── Sert les fichiers statiques (l'add-in lui-même) ─────────────────────────
// Si votre add-in est hébergé sur GitHub Pages, vous n'avez pas besoin de ça.
// Utile uniquement si vous hébergez tout sur ce même serveur.
const PUBLIC_DIR = path.join(__dirname, "..");
app.use(express.static(PUBLIC_DIR));

// ─── ROUTES ──────────────────────────────────────────────────────────────────

/** Santé du serveur */
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    version: "2.0.0",
    gemini: "configuré",
    timestamp: new Date().toISOString(),
  });
});

/**
 * POST /api/ai/analyze
 * Analyse IA complète du bilan carbone
 * Body : { bilan: {...}, historique: [...] }
 * Retourne : JSON structuré { contexte, anomalies, suggestions, tendance }
 */
app.post("/api/ai/analyze", checkSecret, burstLimiter, aiLimiter, async (req, res) => {
  const { bilan, historique = [] } = req.body;

  if (!bilan || !bilan.grandTotal) {
    return res.status(400).json({ error: "Données bilan manquantes ou invalides." });
  }

  try {
    const prompt = buildAnalyzePrompt(bilan, historique);
    const geminiResponse = await callGemini(prompt);
    const parsed = parseGeminiJSON(geminiResponse);
    res.json({ success: true, result: parsed });
  } catch (e) {
    console.error("[/api/ai/analyze] Erreur :", e.message);
    res.status(500).json({ error: e.message });
  }
});

/**
 * POST /api/ai/chat
 * Message de chat ESG interactif
 * Body : { message: "...", history: [...], bilan: {...} }
 * Retourne : { reply: "..." }
 */
app.post("/api/ai/chat", checkSecret, burstLimiter, aiLimiter, async (req, res) => {
  const { message, history = [], bilan } = req.body;

  if (!message || message.trim().length === 0) {
    return res.status(400).json({ error: "Message vide." });
  }
  if (message.length > 2000) {
    return res.status(400).json({ error: "Message trop long (max 2000 caractères)." });
  }

  try {
    // Premier message : injecter le contexte bilan
    let finalMessage = message;
    if (history.length === 0 && bilan) {
      const ctx = buildBilanContext(bilan);
      finalMessage = `Contexte du bilan carbone (référence pour toute la conversation) :\n\n${ctx}\n\n---\n\nQuestion : ${message}`;
    }

    const geminiResponse = await callGeminiChat(finalMessage, history);
    res.json({ success: true, reply: geminiResponse });
  } catch (e) {
    console.error("[/api/ai/chat] Erreur :", e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── Appels Gemini (côté serveur) ─────────────────────────────────────────────

const GEMINI_MODEL   = "gemini-2.0-flash";
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

const SYSTEM_PROMPT = `Tu es un expert ESG et consultant en bilan carbone certifié GHG Protocol.
Tu analyses des données d'émissions de gaz à effet de serre pour des entreprises industrielles.
Tu connais parfaitement la Base Carbone ADEME, les scopes 1/2/3, ISO 14064, le GHG Protocol Corporate Standard, la CSRD et les trajectoires SBTi.
Tu réponds toujours en français, de façon précise, structurée et actionnable.
Tu fournis des recommandations concrètes avec des ordres de grandeur chiffrés.
Tu mentionnes les référentiels applicables (ADEME, CSRD, SBTi, RE2020, etc.) quand pertinent.`;

async function callGemini(userPrompt) {
  const body = {
    systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
    contents: [{ role: "user", parts: [{ text: userPrompt }] }],
    generationConfig: { maxOutputTokens: 2048, temperature: 0.3, topP: 0.9 },
  };

  return fetchGemini(body);
}

async function callGeminiChat(userMessage, history) {
  // Construire le contenu avec l'historique
  const contents = [
    ...history,
    { role: "user", parts: [{ text: userMessage }] },
  ];

  const body = {
    systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
    contents,
    generationConfig: { maxOutputTokens: 1024, temperature: 0.5, topP: 0.9 },
  };

  return fetchGemini(body);
}

async function fetchGemini(body) {
  const url = `${GEMINI_API_URL}?key=${GEMINI_API_KEY}`;

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const msg = err?.error?.message || `HTTP ${response.status}`;
    if (response.status === 400) throw new Error(`Requête Gemini invalide : ${msg}`);
    if (response.status === 429) throw new Error("Quota Gemini dépassé — réessayez dans quelques secondes.");
    if (response.status === 403) throw new Error("Clé API Gemini invalide ou désactivée.");
    throw new Error(`Erreur Gemini : ${msg}`);
  }

  const data = await response.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error("Réponse Gemini vide ou bloquée par les filtres de sécurité.");
  return text;
}

// ─── Builders de prompts ──────────────────────────────────────────────────────

function buildAnalyzePrompt(bilan, historique) {
  const ctx = buildBilanContext(bilan, historique);
  return `Analyse ce bilan carbone d'entreprise et fournis une réponse JSON structurée.

${ctx}

Réponds UNIQUEMENT avec ce JSON valide (sans backticks, sans markdown) :
{
  "contexte": "2-3 phrases de synthèse du profil d'émissions de cette entreprise",
  "anomalies": [
    {
      "severity": "error|warn|info",
      "scope": "Scope X ou Global",
      "titre": "Titre court",
      "message": "Explication détaillée avec chiffres concrets"
    }
  ],
  "suggestions": [
    {
      "priority": "high|medium|low",
      "scope": "Scope X",
      "action": "Titre de l'action",
      "detail": "Description concrète et actionnable avec exemples",
      "potentiel": "Réduction estimée en %",
      "delai": "Court terme (< 1 an) | Moyen terme (1-3 ans) | Long terme (> 3 ans)",
      "referentiel": "Norme ou référentiel applicable (ADEME, CSRD, SBTi...)"
    }
  ],
  "tendance": {
    "commentaire": "Analyse de la tendance si données historiques disponibles, sinon chaîne vide",
    "trajectoire": "En hausse | En baisse | Stable | Données insuffisantes",
    "alerteCSRD": true,
    "alerteSBTi": "Explication si trajectoire non alignée SBTi, sinon chaîne vide"
  }
}`;
}

function buildBilanContext(bilan, historique = []) {
  const lines = [
    `ENTREPRISE : ${bilan.entreprise || "N/A"}`,
    `SECTEUR    : ${bilan.secteur || "Industrie"}`,
    `ANNÉE      : ${bilan.annee || new Date().getFullYear()}`,
    bilan.chiffreAffaires ? `CA         : ${(bilan.chiffreAffaires / 1e6).toFixed(1)} M€` : null,
    "",
    "BILAN CARBONE :",
    `  Scope 1 : ${bilan.scope1?.total ?? 0} tCO2eq (${bilan.scope1?.pct ?? 0}%)`,
    `  Scope 2 : ${bilan.scope2?.total ?? 0} tCO2eq (${bilan.scope2?.pct ?? 0}%)`,
    `  Scope 3 : ${bilan.scope3?.total ?? 0} tCO2eq (${bilan.scope3?.pct ?? 0}%)`,
    `  TOTAL   : ${bilan.grandTotal ?? 0} tCO2eq`,
    bilan.intensite ? `  Intensité : ${bilan.intensite} ${bilan.intensiteUnit}` : null,
    "",
    "SOURCES SCOPE 1 :",
    ...(bilan.scope1?.lines || []).map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
    "SOURCES SCOPE 2 :",
    ...(bilan.scope2?.lines || []).map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
    "SOURCES SCOPE 3 :",
    ...(bilan.scope3?.lines || []).map(l => `  - ${l.source} : ${l.tCO2eq} tCO2eq`),
  ].filter(l => l !== null);

  if (historique.length > 0) {
    lines.push("", "HISTORIQUE :");
    historique.forEach(h => {
      const v = h.variation !== null ? ` (${h.variation > 0 ? "+" : ""}${h.variation}%)` : "";
      lines.push(`  ${h.annee} : ${h.total} tCO2eq${v}`);
    });
    if (historique.length >= 2) {
      const oldest = historique[0].total;
      const totalVar = (((bilan.grandTotal - oldest) / oldest) * 100).toFixed(1);
      lines.push(`  Variation totale ${historique[0].annee}→${bilan.annee} : ${totalVar > 0 ? "+" : ""}${totalVar}%`);
    }
  }

  return lines.join("\n");
}

function parseGeminiJSON(text) {
  try {
    const clean = text.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    return JSON.parse(clean);
  } catch {
    // Si le JSON est mal formé, retourner une structure minimale avec le texte brut
    return {
      contexte: text,
      anomalies: [],
      suggestions: [],
      tendance: { commentaire: "", trajectoire: "Données insuffisantes", alerteCSRD: false, alerteSBTi: "" },
    };
  }
}

// ─── Gestion des erreurs globale ──────────────────────────────────────────────
app.use((err, req, res, next) => {
  if (err.message && err.message.startsWith("CORS")) {
    return res.status(403).json({ error: err.message });
  }
  console.error("[Erreur globale]", err.message);
  res.status(500).json({ error: "Erreur serveur interne." });
});

// ─── Démarrage (HTTP en dev, HTTPS en prod géré par le proxy de l'hébergeur) ──
app.listen(PORT, () => {
  console.log(`
  ╔═══════════════════════════════════════════════════╗
  ║   ESG Analyzer Pro — Backend v2.0                ║
  ║   Port    : ${PORT}                                   ║
  ║   CORS    : ${ALLOWED_ORIGIN.padEnd(30)}  ║
  ║   Gemini  : ✅ configuré (clé masquée)            ║
  ║   Santé   : GET /api/health                       ║
  ╚═══════════════════════════════════════════════════╝
  `);
});

module.exports = app;
