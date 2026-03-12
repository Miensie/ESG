/**
 * ============================================================
 * MODULE : carbonCalc.js
 * Calcul du bilan carbone selon GHG Protocol / ADEME
 * Scopes 1, 2 et 3
 * ============================================================
 */

"use strict";

// ─── Facteurs d'émission (tCO2eq par unité) ──────────────────────────────────
// Source : Base Carbone ADEME 2024 + GHG Protocol
const EMISSION_FACTORS = {

  // ── SCOPE 1 : Émissions directes ─────────────────────────────────────────
  scope1: {
    // Combustibles (tCO2eq / MWh PCI)
    naturalGas:     { factor: 0.2041,  unit: "MWh",  label: "Gaz naturel" },
    fuelOil:        { factor: 0.2773,  unit: "MWh",  label: "Fioul lourd" },
    diesel:         { factor: 0.2670,  unit: "MWh",  label: "Gazole" },
    lpg:            { factor: 0.2274,  unit: "MWh",  label: "GPL" },
    coal:           { factor: 0.3411,  unit: "MWh",  label: "Charbon" },
    // Procédés industriels (tCO2eq / tonne produit)
    cementProduction: { factor: 0.820, unit: "t",    label: "Production ciment" },
    steelProduction:  { factor: 1.850, unit: "t",    label: "Production acier" },
    // Fluides frigorigènes (tCO2eq / kg)
    r410a:          { factor: 2.088,   unit: "kg",   label: "R-410A (fuite)" },
    r32:            { factor: 0.675,   unit: "kg",   label: "R-32 (fuite)" },
    r404a:          { factor: 3.922,   unit: "kg",   label: "R-404A (fuite)" },
  },

  // ── SCOPE 2 : Énergie achetée ─────────────────────────────────────────────
  scope2: {
    // Électricité (tCO2eq / MWh) — facteurs par pays
    electricityFrance:     { factor: 0.0490, unit: "MWh", label: "Électricité France" },
    electricityEurope:     { factor: 0.2950, unit: "MWh", label: "Électricité Europe moy." },
    electricityUSA:        { factor: 0.3860, unit: "MWh", label: "Électricité USA" },
    electricityChina:      { factor: 0.5810, unit: "MWh", label: "Électricité Chine" },
    electricityWorld:      { factor: 0.4760, unit: "MWh", label: "Électricité monde moy." },
    // Vapeur / chaleur achetée (tCO2eq / MWh)
    districtHeat:          { factor: 0.1670, unit: "MWh", label: "Chaleur réseau" },
    steamPurchased:        { factor: 0.1020, unit: "MWh", label: "Vapeur achetée" },
  },

  // ── SCOPE 3 : Émissions indirectes ───────────────────────────────────────
  scope3: {
    // Transport amont fret (tCO2eq / t.km)
    roadFreightLight:    { factor: 0.000191, unit: "t.km", label: "Fret routier léger (<3,5t)" },
    roadFreightHeavy:    { factor: 0.000096, unit: "t.km", label: "Fret routier lourd (>3,5t)" },
    railFreight:         { factor: 0.000028, unit: "t.km", label: "Fret ferroviaire" },
    seaFreight:          { factor: 0.000011, unit: "t.km", label: "Fret maritime" },
    airFreight:          { factor: 0.000602, unit: "t.km", label: "Fret aérien" },
    // Transport employés (tCO2eq / km.personne)
    businessCar:         { factor: 0.000193, unit: "km",   label: "Voiture essence (dépl. pro)" },
    businessCarDiesel:   { factor: 0.000163, unit: "km",   label: "Voiture diesel (dépl. pro)" },
    businessCarElectric: { factor: 0.000019, unit: "km",   label: "Voiture électrique (dépl. pro)" },
    businessAirShort:    { factor: 0.000258, unit: "km",   label: "Avion court-courrier" },
    businessAirLong:     { factor: 0.000195, unit: "km",   label: "Avion long-courrier" },
    trainTravel:         { factor: 0.000003, unit: "km",   label: "Train (dépl. pro)" },
    // Déchets (tCO2eq / tonne)
    wasteIncineration:   { factor: 0.8540,   unit: "t",    label: "Incinération déchets" },
    wasteLandfill:       { factor: 0.4580,   unit: "t",    label: "Enfouissement déchets" },
    wasteRecycling:      { factor: -0.0830,  unit: "t",    label: "Recyclage (évité)" },
    wasteWater:          { factor: 0.7080,   unit: "t",    label: "Traitement eaux usées" },
    // Achats de biens et services
    steelPurchased:      { factor: 1.850,    unit: "t",    label: "Acier acheté" },
    aluminiumPurchased:  { factor: 8.240,    unit: "t",    label: "Aluminium acheté" },
    plasticPurchased:    { factor: 3.140,    unit: "t",    label: "Plastique acheté" },
    paperPurchased:      { factor: 0.919,    unit: "t",    label: "Papier/carton acheté" },
    concreteUsed:        { factor: 0.130,    unit: "t",    label: "Béton utilisé" },
  }
};

// ─── Calculateur principal ────────────────────────────────────────────────────

/**
 * Calcule les émissions pour une activité donnée
 * @param {string} category  - "scope1" | "scope2" | "scope3"
 * @param {string} type      - clé du facteur (ex: "naturalGas")
 * @param {number} quantity  - quantité dans l'unité du facteur
 * @returns {{ tCO2eq: number, label: string, unit: string, factor: number }}
 */
function calcEmission(category, type, quantity) {
  const scopeFactors = EMISSION_FACTORS[category];
  if (!scopeFactors) throw new Error(`Catégorie inconnue : ${category}`);
  const ef = scopeFactors[type];
  if (!ef) throw new Error(`Type inconnu : ${type} dans ${category}`);

  return {
    tCO2eq: parseFloat((ef.factor * quantity).toFixed(4)),
    label:  ef.label,
    unit:   ef.unit,
    factor: ef.factor,
    quantity
  };
}

/**
 * Calcule le bilan carbone complet depuis un objet de données structuré
 * @param {Object} data - Données ESG collectées depuis Excel
 * @returns {BilanCarbone}
 */
function computeFullBilan(data) {
  const result = {
    scope1: { total: 0, lines: [] },
    scope2: { total: 0, lines: [] },
    scope3: { total: 0, lines: [] },
    grandTotal: 0,
    intensite: 0,
    annee: data.annee || new Date().getFullYear(),
    entreprise: data.entreprise || "—",
    secteur: data.secteur || "Industrie",
  };

  // Calcul Scope 1
  (data.scope1 || []).forEach(item => {
    try {
      const em = calcEmission("scope1", item.type, item.quantity);
      result.scope1.lines.push({ ...em, source: item.source || item.type });
      result.scope1.total += em.tCO2eq;
    } catch (e) {
      console.warn("[Scope1] Ignoré :", item, e.message);
    }
  });

  // Calcul Scope 2
  (data.scope2 || []).forEach(item => {
    try {
      const em = calcEmission("scope2", item.type, item.quantity);
      result.scope2.lines.push({ ...em, source: item.source || item.type });
      result.scope2.total += em.tCO2eq;
    } catch (e) {
      console.warn("[Scope2] Ignoré :", item, e.message);
    }
  });

  // Calcul Scope 3
  (data.scope3 || []).forEach(item => {
    try {
      const em = calcEmission("scope3", item.type, item.quantity);
      result.scope3.lines.push({ ...em, source: item.source || item.type });
      result.scope3.total += em.tCO2eq;
    } catch (e) {
      console.warn("[Scope3] Ignoré :", item, e.message);
    }
  });

  // Totaux
  result.scope1.total = parseFloat(result.scope1.total.toFixed(2));
  result.scope2.total = parseFloat(result.scope2.total.toFixed(2));
  result.scope3.total = parseFloat(result.scope3.total.toFixed(2));
  result.grandTotal = parseFloat(
    (result.scope1.total + result.scope2.total + result.scope3.total).toFixed(2)
  );

  // Intensité carbone (tCO2eq / M€ CA ou tCO2eq / t produit)
  if (data.chiffreAffaires && data.chiffreAffaires > 0) {
    result.intensite = parseFloat(
      (result.grandTotal / (data.chiffreAffaires / 1_000_000)).toFixed(2)
    );
    result.intensiteUnit = "tCO2eq/M€";
  } else if (data.productionTonnes && data.productionTonnes > 0) {
    result.intensite = parseFloat(
      (result.grandTotal / data.productionTonnes).toFixed(4)
    );
    result.intensiteUnit = "tCO2eq/t produit";
  }

  // Répartition en %
  result.scope1.pct = result.grandTotal > 0
    ? ((result.scope1.total / result.grandTotal) * 100).toFixed(1) : "0.0";
  result.scope2.pct = result.grandTotal > 0
    ? ((result.scope2.total / result.grandTotal) * 100).toFixed(1) : "0.0";
  result.scope3.pct = result.grandTotal > 0
    ? ((result.scope3.total / result.grandTotal) * 100).toFixed(1) : "0.0";

  return result;
}

/**
 * Données de démonstration pour tester sans Excel
 */
function getDemoData() {
  return {
    annee: 2024,
    entreprise: "IndustrieCo SAS",
    secteur: "Industrie manufacturière",
    chiffreAffaires: 45_000_000,
    scope1: [
      { type: "naturalGas",    quantity: 8500,  source: "Chaudière principale" },
      { type: "diesel",        quantity: 1200,  source: "Groupe électrogène" },
      { type: "r410a",         quantity: 25,    source: "Climatisation" },
    ],
    scope2: [
      { type: "electricityFrance", quantity: 6800, source: "Site A – production" },
      { type: "electricityFrance", quantity: 1200, source: "Site B – bureaux" },
      { type: "districtHeat",      quantity: 850,  source: "Réseau chaleur" },
    ],
    scope3: [
      { type: "roadFreightHeavy",  quantity: 420000, source: "Livraisons clients" },
      { type: "seaFreight",        quantity: 180000, source: "Import matières" },
      { type: "businessAirShort",  quantity: 85000,  source: "Déplacements pro" },
      { type: "businessCar",       quantity: 124000, source: "Voitures salariés" },
      { type: "wasteIncineration", quantity: 185,    source: "Déchets industriels" },
      { type: "wasteRecycling",    quantity: 95,     source: "Recyclage métal" },
      { type: "steelPurchased",    quantity: 320,    source: "Matières premières" },
    ],
  };
}

/**
 * Détection d'anomalies dans le bilan
 */
function detectAnomalies(bilan, previousBilan = null) {
  const anomalies = [];

  // Scope 3 > 80% sans explications = alerte
  if (parseFloat(bilan.scope3.pct) > 80) {
    anomalies.push({
      severity: "warn",
      scope: "Scope 3",
      message: `Scope 3 représente ${bilan.scope3.pct}% des émissions — vérifier l'exhaustivité des données Scope 1 & 2.`
    });
  }

  // Variation > 20% par rapport à N-1
  if (previousBilan) {
    const variation = ((bilan.grandTotal - previousBilan.grandTotal) / previousBilan.grandTotal) * 100;
    if (Math.abs(variation) > 20) {
      anomalies.push({
        severity: variation > 0 ? "error" : "info",
        scope: "Global",
        message: `Variation de ${variation.toFixed(1)}% vs N-1 — ${
          variation > 0 ? "hausse significative à investiguer" : "baisse à valider"
        }.`
      });
    }
  }

  // Intensité carbone élevée
  if (bilan.intensiteUnit === "tCO2eq/M€" && bilan.intensite > 500) {
    anomalies.push({
      severity: "warn",
      scope: "Intensité",
      message: `Intensité carbone de ${bilan.intensite} tCO2eq/M€ — au-dessus de la médiane sectorielle (500).`
    });
  }

  // Lignes à 0
  const zerosScope1 = bilan.scope1.lines.filter(l => l.tCO2eq === 0).length;
  if (zerosScope1 > 0) {
    anomalies.push({
      severity: "info",
      scope: "Scope 1",
      message: `${zerosScope1} source(s) avec émissions nulles — données manquantes ?`
    });
  }

  return anomalies;
}

/**
 * Suggestions de réduction basées sur les résultats
 */
function generateSuggestions(bilan) {
  const suggestions = [];
  const s1pct = parseFloat(bilan.scope1.pct);
  const s2pct = parseFloat(bilan.scope2.pct);
  const s3pct = parseFloat(bilan.scope3.pct);

  // Scope 1 dominant
  if (s1pct > 40) {
    suggestions.push({
      priority: "high",
      scope: "Scope 1",
      action: "Décarbonation combustibles",
      detail: "Substituer gaz naturel / fioul par biomasse, hydrogène ou chaleur fatale récupérée.",
      potentiel: "–30 à –60% Scope 1"
    });
  }

  // Scope 2 > 20%
  if (s2pct > 20) {
    suggestions.push({
      priority: "high",
      scope: "Scope 2",
      action: "Électricité bas-carbone",
      detail: "Souscrire un contrat PPAVRE ou installer du solaire en autoconsommation.",
      potentiel: "–80 à –100% Scope 2"
    });
  }

  // Transport Scope 3
  const hasHighFreight = bilan.scope3.lines.some(
    l => (l.source || "").toLowerCase().includes("fret") && l.tCO2eq > 50
  );
  if (hasHighFreight) {
    suggestions.push({
      priority: "medium",
      scope: "Scope 3",
      action: "Optimisation logistique",
      detail: "Report modal vers ferroviaire/maritime, optimisation des chargements, mutualisation transport.",
      potentiel: "–20 à –40% émissions transport"
    });
  }

  // Déchets
  const hasIncineration = bilan.scope3.lines.some(l => l.source.includes("Incinér"));
  if (hasIncineration) {
    suggestions.push({
      priority: "medium",
      scope: "Scope 3",
      action: "Plan de réduction déchets",
      detail: "Économie circulaire, valorisation matière, diagnostic déchets pour éviter l'incinération.",
      potentiel: "–15 à –30% déchets"
    });
  }

  // Efficacité énergétique globale
  suggestions.push({
    priority: "low",
    scope: "Scope 1+2",
    action: "Audit énergétique ISO 50001",
    detail: "Mettre en place un système de management de l'énergie pour identifier et réduire les gaspillages.",
    potentiel: "–10 à –25% énergie"
  });

  return suggestions;
}

// Export
window.CarbonCalc = {
  EMISSION_FACTORS,
  calcEmission,
  computeFullBilan,
  getDemoData,
  detectAnomalies,
  generateSuggestions,
};
