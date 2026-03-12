/**
 * ============================================================
 * ESG CALCULATOR — Moteur de calcul bilan carbone
 * Basé sur GHG Protocol & facteurs ADEME 2024
 * ============================================================
 */

"use strict";

// ============================================================
// CONSTANTES : FACTEURS D'ÉMISSION PAR DÉFAUT (kgCO₂e/unité)
// Source : Base Carbone ADEME 2024
// ============================================================
const EMISSION_FACTORS = {
  // Scope 1 — Émissions directes
  scope1: {
    gaz_naturel:     { factor: 0.205,   unit: "MWh",   label: "Gaz naturel" },
    fioul_lourd:     { factor: 2.96,    unit: "L",     label: "Fioul lourd" },
    charbon:         { factor: 0.341,   unit: "kWh",   label: "Charbon" },
    propane:         { factor: 1.52,    unit: "kg",    label: "Propane" },
    process_chimique:{ factor: 1.0,     unit: "tCO₂",  label: "Procédés chimiques" },
    refrigerants:    { factor: 1430,    unit: "kg",    label: "Réfrigérants (R-134a)" },
  },
  // Scope 2 — Énergie achetée
  scope2: {
    electricite_rfe: { factor: 0.0571,  unit: "MWh",   label: "Électricité réseau FR" },
    electricite_eu:  { factor: 0.296,   unit: "MWh",   label: "Électricité mix EU" },
    chaleur_reseau:  { factor: 0.059,   unit: "MWh",   label: "Chaleur réseau" },
    froid_reseau:    { factor: 0.030,   unit: "MWh",   label: "Froid réseau" },
  },
  // Scope 3 — Émissions indirectes
  scope3: {
    transport_route: { factor: 0.0621,  unit: "t·km",  label: "Transport routier (PL)" },
    transport_avion: { factor: 0.158,   unit: "pass·km",label: "Avion (éco court)" },
    transport_train: { factor: 0.00285, unit: "pass·km",label: "Train" },
    dechets_enfouis: { factor: 449,     unit: "t",     label: "Déchets enfouis" },
    dechets_incineres:{ factor: 0.540,  unit: "t",     label: "Déchets incinérés" },
    achats_matieres: { factor: 1.85,    unit: "t",     label: "Achats matières" },
    achats_services: { factor: 0.65,    unit: "k€",    label: "Achats services" },
    eau:             { factor: 0.000344,unit: "m³",    label: "Consommation eau" },
  }
};

// ============================================================
// CLASSE PRINCIPALE : ESGCalculator
// ============================================================
class ESGCalculator {

  constructor() {
    this.results = {
      scope1: 0, scope2: 0, scope3: 0, total: 0,
      details: [],
      intensity: 0,
      metadata: {}
    };
  }

  // ----------------------------------------------------------
  // Calcul principal depuis saisie manuelle
  // ----------------------------------------------------------
  calculateFromManual(inputs, customFactors = {}) {
    const fe = this._mergeFactors(customFactors);
    const details = [];

    // SCOPE 1
    const s1_gaz      = this._calc(inputs.qty_gaz,      fe.gaz_naturel,      "Scope 1", "Gaz naturel");
    const s1_fioul    = this._calc(inputs.qty_fioul,     fe.fioul_lourd,      "Scope 1", "Fioul lourd");
    const s1_process  = this._calc(inputs.qty_process,   fe.process_chimique, "Scope 1", "Procédés");
    details.push(s1_gaz, s1_fioul, s1_process);
    const scope1_total = s1_gaz.tco2e + s1_fioul.tco2e + s1_process.tco2e;

    // SCOPE 2
    const s2_elec     = this._calc(inputs.qty_elec,     fe.electricite_rfe,  "Scope 2", "Électricité");
    const s2_chaleur  = this._calc(inputs.qty_chaleur,  fe.chaleur_reseau,   "Scope 2", "Chaleur réseau");
    details.push(s2_elec, s2_chaleur);
    const scope2_total = s2_elec.tco2e + s2_chaleur.tco2e;

    // SCOPE 3
    const s3_transport= this._calc(inputs.qty_transport, fe.transport_route,  "Scope 3", "Transport");
    const s3_dechets  = this._calc(inputs.qty_dechets,   fe.dechets_enfouis,  "Scope 3", "Déchets enfouis");
    const s3_achats   = this._calc(inputs.qty_achats,    fe.achats_matieres,  "Scope 3", "Achats matières");
    details.push(s3_transport, s3_dechets, s3_achats);
    const scope3_total = s3_transport.tco2e + s3_dechets.tco2e + s3_achats.tco2e;

    const total = scope1_total + scope2_total + scope3_total;

    this.results = {
      scope1: Math.round(scope1_total * 100) / 100,
      scope2: Math.round(scope2_total * 100) / 100,
      scope3: Math.round(scope3_total * 100) / 100,
      total:  Math.round(total * 100) / 100,
      details: details.filter(d => d.qty > 0),
      intensity: inputs.activity_value > 0
        ? Math.round((total / inputs.activity_value) * 1000) / 1000
        : null,
      metadata: {
        year:         inputs.year || new Date().getFullYear(),
        activity_unit: inputs.activity_unit || "unité",
        activity_value: inputs.activity_value || 0,
        calculated_at: new Date().toISOString()
      }
    };

    return this.results;
  }

  // ----------------------------------------------------------
  // Calcul depuis données Excel structurées
  // ----------------------------------------------------------
  calculateFromExcelData(sheetData, customFactors = {}) {
    const fe = this._mergeFactors(customFactors);
    let scope1 = 0, scope2 = 0, scope3 = 0;
    const details = [];

    if (!sheetData || !Array.isArray(sheetData)) {
      throw new Error("Données Excel invalides ou manquantes");
    }

    for (const row of sheetData) {
      if (!row || row.length < 3) continue;
      const source  = String(row[0] || "").toLowerCase().trim();
      const qty     = parseFloat(row[1]) || 0;
      const scope   = String(row[2] || "").toLowerCase().trim();
      const label   = String(row[0] || "Source inconnue");

      if (qty === 0) continue;

      // Détecter le facteur d'émission par mots-clés dans le nom de source
      const factor = this._detectFactor(source, fe);

      if (factor) {
        const tco2e = (qty * factor.value) / 1000; // kg → t
        const detail = { source: label, qty, unit: factor.unit, factor: factor.value, tco2e, scope: scope || factor.scope };
        details.push(detail);

        const scopeKey = scope.replace("scope ", "scope").replace("scope1","s1").replace("scope2","s2").replace("scope3","s3");
        if (scope.includes("1") || factor.scope === "scope1") scope1 += tco2e;
        else if (scope.includes("2") || factor.scope === "scope2") scope2 += tco2e;
        else scope3 += tco2e;
      }
    }

    const total = scope1 + scope2 + scope3;
    this.results = {
      scope1: Math.round(scope1 * 100) / 100,
      scope2: Math.round(scope2 * 100) / 100,
      scope3: Math.round(scope3 * 100) / 100,
      total:  Math.round(total * 100) / 100,
      details,
      intensity: null,
      metadata: { calculated_at: new Date().toISOString() }
    };

    return this.results;
  }

  // ----------------------------------------------------------
  // Détection anomalies dans un jeu de données
  // ----------------------------------------------------------
  detectAnomalies(data, threshold = 20) {
    const anomalies = [];
    if (!data || data.length < 2) return anomalies;

    // Calcul moyenne et écart-type par source
    const bySource = {};
    for (const row of data) {
      const key = String(row[0] || "").trim();
      const val = parseFloat(row[1]) || 0;
      if (!bySource[key]) bySource[key] = [];
      bySource[key].push(val);
    }

    for (const [source, values] of Object.entries(bySource)) {
      if (values.length < 2) continue;
      const mean = values.reduce((a, b) => a + b, 0) / values.length;
      const std  = Math.sqrt(values.map(v => Math.pow(v - mean, 2)).reduce((a, b) => a + b, 0) / values.length);

      for (let i = 0; i < values.length; i++) {
        const pct = mean > 0 ? Math.abs((values[i] - mean) / mean) * 100 : 0;
        if (pct > threshold) {
          const severity = pct > 50 ? "critical" : pct > 30 ? "warning" : "info";
          anomalies.push({
            source,
            value: values[i],
            mean: Math.round(mean * 100) / 100,
            deviation: Math.round(pct * 10) / 10,
            severity,
            index: i,
            message: `Écart de ${Math.round(pct)}% par rapport à la moyenne (${Math.round(mean)})`,
            action: pct > 30 ? "Vérifier les données sources et les processus associés" : "Surveiller l'évolution"
          });
        }
      }
    }

    return anomalies;
  }

  // ----------------------------------------------------------
  // Générer recommandations de réduction
  // ----------------------------------------------------------
  generateRecommendations(results, targetPct = 30, horizonYears = 3) {
    const recos = [];
    if (!results || results.total === 0) return recos;

    const targetReduction = results.total * (targetPct / 100);

    // Recommandations basées sur la répartition des scopes
    const pct1 = results.scope1 / results.total * 100;
    const pct2 = results.scope2 / results.total * 100;
    const pct3 = results.scope3 / results.total * 100;

    // Scope 2 — Toujours actionnable
    if (pct2 > 10) {
      const saving = results.scope2 * 0.6;
      recos.push({
        priority: 1, scope: "Scope 2",
        title: "Passage aux énergies renouvelables",
        description: "Souscrire à des contrats d'électricité verte certifiée (garanties d'origine) ou installer des panneaux solaires.",
        saving_tco2e: Math.round(saving * 10) / 10,
        saving_pct: Math.round(saving / results.total * 100),
        horizon: "1-2 ans",
        investment: "Moyen",
        icon: "⚡"
      });
    }

    // Scope 3 Transport
    if (results.scope3 > results.scope1 * 0.5) {
      const saving = results.scope3 * 0.25;
      recos.push({
        priority: 2, scope: "Scope 3",
        title: "Optimisation logistique & transport",
        description: "Optimiser les tournées, massifier les flux, favoriser le rail et le fret maritime pour les longues distances.",
        saving_tco2e: Math.round(saving * 10) / 10,
        saving_pct: Math.round(saving / results.total * 100),
        horizon: "1-3 ans",
        investment: "Faible",
        icon: "🚛"
      });
    }

    // Scope 1 — Efficacité énergétique
    if (pct1 > 30) {
      const saving = results.scope1 * 0.15;
      recos.push({
        priority: 3, scope: "Scope 1",
        title: "Efficacité énergétique des procédés",
        description: "Audit énergétique, isolation thermique, récupération de chaleur fatale, optimisation des fours et compresseurs.",
        saving_tco2e: Math.round(saving * 10) / 10,
        saving_pct: Math.round(saving / results.total * 100),
        horizon: "2-4 ans",
        investment: "Élevé",
        icon: "🏭"
      });
    }

    // Déchets
    recos.push({
      priority: 4, scope: "Scope 3",
      title: "Plan de réduction des déchets",
      description: "Objectif zéro déchets enfouis (ZLD) : tri sélectif, valorisation matière, économie circulaire avec les fournisseurs.",
      saving_tco2e: Math.round(results.scope3 * 0.08 * 10) / 10,
      saving_pct: Math.round(results.scope3 * 0.08 / results.total * 100),
      horizon: "1-2 ans",
      investment: "Faible",
      icon: "♻️"
    });

    // ISO 50001
    recos.push({
      priority: 5, scope: "Tous",
      title: "Certification ISO 50001 (énergie)",
      description: "Mise en place d'un système de management de l'énergie pour réduire structurellement les consommations.",
      saving_tco2e: Math.round(results.total * 0.05 * 10) / 10,
      saving_pct: 5,
      horizon: "2-3 ans",
      investment: "Moyen",
      icon: "📋"
    });

    return recos.sort((a, b) => b.saving_tco2e - a.saving_tco2e);
  }

  // ----------------------------------------------------------
  // UTILITAIRES PRIVÉS
  // ----------------------------------------------------------
  _calc(qty, factorKey, scopeLabel, label) {
    const q   = parseFloat(qty) || 0;
    const fac = this._getFactorValue(factorKey);
    const tco2e = (q * fac) / 1000; // kg → t
    return {
      label, scope: scopeLabel,
      qty: q, factor: fac,
      unit: EMISSION_FACTORS.scope1[factorKey]?.unit ||
            EMISSION_FACTORS.scope2[factorKey]?.unit ||
            EMISSION_FACTORS.scope3[factorKey]?.unit || "—",
      tco2e: Math.round(tco2e * 1000) / 1000
    };
  }

  _getFactorValue(key) {
    for (const scope of Object.values(EMISSION_FACTORS)) {
      if (scope[key]) return scope[key].factor;
    }
    return 1.0;
  }

  _mergeFactors(custom) {
    return {
      gaz_naturel:     custom.fe_gaz      || EMISSION_FACTORS.scope1.gaz_naturel.factor,
      fioul_lourd:     custom.fe_fioul    || EMISSION_FACTORS.scope1.fioul_lourd.factor,
      process_chimique:custom.fe_process  || EMISSION_FACTORS.scope1.process_chimique.factor,
      electricite_rfe: custom.fe_elec     || EMISSION_FACTORS.scope2.electricite_rfe.factor,
      chaleur_reseau:  custom.fe_chaleur  || EMISSION_FACTORS.scope2.chaleur_reseau.factor,
      transport_route: custom.fe_transport|| EMISSION_FACTORS.scope3.transport_route.factor,
      dechets_enfouis: custom.fe_dechets  || EMISSION_FACTORS.scope3.dechets_enfouis.factor,
      achats_matieres: custom.fe_achats   || EMISSION_FACTORS.scope3.achats_matieres.factor,
    };
  }

  _detectFactor(source, fe) {
    if (source.includes("gaz") || source.includes("gas"))
      return { value: fe.gaz_naturel * 1000, unit: "MWh", scope: "scope1" };
    if (source.includes("fioul") || source.includes("fuel") || source.includes("gazole"))
      return { value: fe.fioul_lourd * 1000, unit: "L", scope: "scope1" };
    if (source.includes("elec") || source.includes("électr"))
      return { value: fe.electricite_rfe * 1000, unit: "MWh", scope: "scope2" };
    if (source.includes("chaleur") || source.includes("vapeur"))
      return { value: fe.chaleur_reseau * 1000, unit: "MWh", scope: "scope2" };
    if (source.includes("transport") || source.includes("camion") || source.includes("fret"))
      return { value: fe.transport_route * 1000, unit: "t·km", scope: "scope3" };
    if (source.includes("déchet") || source.includes("waste"))
      return { value: fe.dechets_enfouis * 1000, unit: "t", scope: "scope3" };
    if (source.includes("achat") || source.includes("matière") || source.includes("mati"))
      return { value: fe.achats_matieres * 1000, unit: "t", scope: "scope3" };
    return null;
  }

  // ----------------------------------------------------------
  // Formatage pour export Excel
  // ----------------------------------------------------------
  toExcelRows() {
    if (!this.results || !this.results.details) return [];
    return this.results.details.map(d => [
      d.label || d.source,
      d.scope,
      d.qty,
      d.unit,
      d.factor,
      d.tco2e
    ]);
  }

  getScopesSummary() {
    return [
      ["Scope 1 — Émissions directes", this.results.scope1, this.results.scope1 / this.results.total * 100],
      ["Scope 2 — Énergie achetée",    this.results.scope2, this.results.scope2 / this.results.total * 100],
      ["Scope 3 — Émissions indirectes",this.results.scope3, this.results.scope3 / this.results.total * 100],
      ["TOTAL",                          this.results.total,  100]
    ];
  }
}

// Export global (utilisé par taskpane.js)
window.ESGCalculator = ESGCalculator;
window.EMISSION_FACTORS = EMISSION_FACTORS;
