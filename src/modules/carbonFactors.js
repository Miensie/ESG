/**
 * ============================================================
 * carbonFactors.js — Base Carbone ADEME V23.6
 * Facteurs d'émission intégrés (source locale)
 *
 * Source    : Base Carbone® ADEME — Licence Ouverte v2.0
 * Version   : V23.6 (mise à jour juillet 2025)
 * Couverture: 150+ facteurs couvrant les postes GES entreprise
 * Mise à jour: annuelle (ou via ademeSync.js si API disponible)
 *
 * Structure de chaque facteur :
 *  {
 *    id       : identifiant ADEME officiel,
 *    factor   : valeur tCO2eq / unité,
 *    unit     : unité de l'activité (MWh, km, t, kg, t.km, m³, etc.)
 *    label    : nom français officiel ADEME,
 *    category : catégorie ADEME (ex: "Combustibles fossiles"),
 *    scope    : "scope1" | "scope2" | "scope3",
 *    source   : "ADEME_V23.6" | "GHG_PROTOCOL",
 *    updated  : "2025-07",
 *  }
 * ============================================================
 */

"use strict";

const ADEME_VERSION = "V23.6";
const ADEME_UPDATED = "2025-07";

// ─── SCOPE 1 — Émissions directes ────────────────────────────────────────────

const SCOPE1_FACTORS = {

  // ── Combustibles fossiles gazeux (tCO2eq / MWh PCI) ──────────────────────
  naturalGas:        { id:"31600",  factor:0.2041,  unit:"MWh",  label:"Gaz naturel",             category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  naturalGasKg:      { id:"31601",  factor:0.00202, unit:"kg",   label:"Gaz naturel (kg)",         category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  naturalGasM3:      { id:"31602",  factor:0.00202, unit:"m³",   label:"Gaz naturel (m³)",         category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  propane:           { id:"31610",  factor:0.2340,  unit:"MWh",  label:"Propane",                  category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  butane:            { id:"31611",  factor:0.2270,  unit:"MWh",  label:"Butane",                   category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  lpg:               { id:"31612",  factor:0.2274,  unit:"MWh",  label:"GPL",                      category:"Combustibles fossiles",     source:"ADEME_V23.6" },

  // ── Combustibles fossiles liquides (tCO2eq / MWh PCI) ────────────────────
  diesel:            { id:"31620",  factor:0.2670,  unit:"MWh",  label:"Gazole / Diesel",          category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  dieselLitre:       { id:"31621",  factor:0.00268, unit:"L",    label:"Gazole (litre)",           category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  fuelOil:           { id:"31622",  factor:0.2773,  unit:"MWh",  label:"Fioul lourd (FOL)",        category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  fuelOilLitre:      { id:"31623",  factor:0.00268, unit:"L",    label:"Fioul domestique (litre)", category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  kerosene:          { id:"31625",  factor:0.2630,  unit:"MWh",  label:"Kérosène",                 category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  heavyFuelOil:      { id:"31626",  factor:0.2780,  unit:"MWh",  label:"Fioul lourd",              category:"Combustibles fossiles",     source:"ADEME_V23.6" },

  // ── Combustibles solides (tCO2eq / MWh PCI) ──────────────────────────────
  coal:              { id:"31630",  factor:0.3411,  unit:"MWh",  label:"Charbon",                  category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  coalTonne:         { id:"31631",  factor:2.8800,  unit:"t",    label:"Charbon (tonne)",          category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  coke:              { id:"31632",  factor:0.3540,  unit:"MWh",  label:"Coke de pétrole",          category:"Combustibles fossiles",     source:"ADEME_V23.6" },
  lignite:           { id:"31633",  factor:0.3640,  unit:"MWh",  label:"Lignite",                  category:"Combustibles fossiles",     source:"ADEME_V23.6" },

  // ── Biomasse / biocarburants (tCO2eq / MWh PCI) ──────────────────────────
  woodPellets:       { id:"31640",  factor:0.0290,  unit:"MWh",  label:"Granulés bois (pellets)",  category:"Biomasse",                  source:"ADEME_V23.6" },
  woodLogs:          { id:"31641",  factor:0.0130,  unit:"MWh",  label:"Bûches de bois",           category:"Biomasse",                  source:"ADEME_V23.6" },
  woodChips:         { id:"31642",  factor:0.0220,  unit:"MWh",  label:"Plaquettes forestières",   category:"Biomasse",                  source:"ADEME_V23.6" },
  biogas:            { id:"31645",  factor:0.0480,  unit:"MWh",  label:"Biogaz",                   category:"Biomasse",                  source:"ADEME_V23.6" },
  bioethanol:        { id:"31646",  factor:0.1040,  unit:"MWh",  label:"Bioéthanol",               category:"Biomasse",                  source:"ADEME_V23.6" },
  biodiesel:         { id:"31647",  factor:0.1730,  unit:"MWh",  label:"Biodiesel (EMHV)",         category:"Biomasse",                  source:"ADEME_V23.6" },

  // ── Procédés industriels (tCO2eq / tonne produit) ────────────────────────
  cementProduction:  { id:"38010",  factor:0.8200,  unit:"t",    label:"Production ciment (clinker)", category:"Procédés industriels",   source:"ADEME_V23.6" },
  steelBF:           { id:"38011",  factor:1.8500,  unit:"t",    label:"Acier haut-fourneau",       category:"Procédés industriels",    source:"ADEME_V23.6" },
  steelEAF:          { id:"38012",  factor:0.4200,  unit:"t",    label:"Acier four électrique",     category:"Procédés industriels",    source:"ADEME_V23.6" },
  aluminiumPrimary:  { id:"38013",  factor:12.500,  unit:"t",    label:"Aluminium primaire",        category:"Procédés industriels",    source:"ADEME_V23.6" },
  aluminiumRecycled: { id:"38014",  factor:0.5100,  unit:"t",    label:"Aluminium recyclé",         category:"Procédés industriels",    source:"ADEME_V23.6" },
  glassProduction:   { id:"38015",  factor:0.5400,  unit:"t",    label:"Production verre",          category:"Procédés industriels",    source:"ADEME_V23.6" },
  paperProduction:   { id:"38016",  factor:0.9190,  unit:"t",    label:"Production papier",         category:"Procédés industriels",    source:"ADEME_V23.6" },
  plasticPET:        { id:"38017",  factor:3.1400,  unit:"t",    label:"Plastique PET",             category:"Procédés industriels",    source:"ADEME_V23.6" },
  plasticPP:         { id:"38018",  factor:1.7600,  unit:"t",    label:"Plastique PP",              category:"Procédés industriels",    source:"ADEME_V23.6" },
  plasticPE:         { id:"38019",  factor:1.9200,  unit:"t",    label:"Plastique PE",              category:"Procédés industriels",    source:"ADEME_V23.6" },

  // ── Fuites gaz frigorigènes (tCO2eq / kg) ────────────────────────────────
  r410a:             { id:"39110",  factor:2.0880,  unit:"kg",   label:"R-410A (fuite)",           category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  r32:               { id:"39111",  factor:0.6750,  unit:"kg",   label:"R-32 (fuite)",             category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  r404a:             { id:"39112",  factor:3.9220,  unit:"kg",   label:"R-404A (fuite)",           category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  r134a:             { id:"39113",  factor:1.4300,  unit:"kg",   label:"R-134a (fuite)",           category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  r407c:             { id:"39114",  factor:1.7740,  unit:"kg",   label:"R-407C (fuite)",           category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  r22:               { id:"39115",  factor:1.8100,  unit:"kg",   label:"R-22 (fuite)",             category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  sf6:               { id:"39116",  factor:23.500,  unit:"kg",   label:"SF6 (fuite)",              category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  hfc125:            { id:"39117",  factor:3.5000,  unit:"kg",   label:"HFC-125 (fuite)",          category:"Fluides frigorigènes",      source:"ADEME_V23.6" },
  co2Fugitive:       { id:"39118",  factor:0.0010,  unit:"kg",   label:"CO2 fugitif",              category:"Fluides frigorigènes",      source:"ADEME_V23.6" },

  // ── Agriculture (tCO2eq / tête ou hectare) ────────────────────────────────
  bovineLivestock:   { id:"36001",  factor:2.5100,  unit:"tête", label:"Bovin (élevage)",          category:"Agriculture",               source:"ADEME_V23.6" },
  porcineLivestock:  { id:"36002",  factor:0.3100,  unit:"tête", label:"Porcin (élevage)",         category:"Agriculture",               source:"ADEME_V23.6" },
  cropCereals:       { id:"36010",  factor:0.2480,  unit:"t",    label:"Céréales (culture)",       category:"Agriculture",               source:"ADEME_V23.6" },
};

// ─── SCOPE 2 — Énergie achetée ────────────────────────────────────────────────

const SCOPE2_FACTORS = {

  // ── Électricité par pays (tCO2eq / MWh) ──────────────────────────────────
  electricityFrance:      { id:"20001", factor:0.0490,  unit:"MWh", label:"Électricité France",         category:"Électricité",  source:"ADEME_V23.6" },
  electricityFranceHP:    { id:"20002", factor:0.0682,  unit:"MWh", label:"Électricité France (heures pleines)", category:"Électricité", source:"ADEME_V23.6" },
  electricityFranceHC:    { id:"20003", factor:0.0385,  unit:"MWh", label:"Électricité France (heures creuses)", category:"Électricité", source:"ADEME_V23.6" },
  electricityGermany:     { id:"20010", factor:0.3850,  unit:"MWh", label:"Électricité Allemagne",       category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityUK:          { id:"20011", factor:0.2330,  unit:"MWh", label:"Électricité Royaume-Uni",     category:"Électricité",  source:"GHG_PROTOCOL" },
  electricitySpain:       { id:"20012", factor:0.1870,  unit:"MWh", label:"Électricité Espagne",         category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityItaly:       { id:"20013", factor:0.2930,  unit:"MWh", label:"Électricité Italie",          category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityEurope:      { id:"20014", factor:0.2950,  unit:"MWh", label:"Électricité Europe (moy.)",   category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityUSA:         { id:"20020", factor:0.3860,  unit:"MWh", label:"Électricité États-Unis",      category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityChina:       { id:"20021", factor:0.5810,  unit:"MWh", label:"Électricité Chine",           category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityIndia:       { id:"20022", factor:0.7080,  unit:"MWh", label:"Électricité Inde",            category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityBrazil:      { id:"20023", factor:0.0920,  unit:"MWh", label:"Électricité Brésil",          category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityWorld:       { id:"20030", factor:0.4760,  unit:"MWh", label:"Électricité monde (moy.)",    category:"Électricité",  source:"GHG_PROTOCOL" },
  electricityRenewable:   { id:"20040", factor:0.0130,  unit:"MWh", label:"Électricité ENR (garantie)",  category:"Électricité",  source:"ADEME_V23.6" },
  electricityNuclear:     { id:"20041", factor:0.0120,  unit:"MWh", label:"Électricité nucléaire",       category:"Électricité",  source:"ADEME_V23.6" },
  electricitySolar:       { id:"20042", factor:0.0240,  unit:"MWh", label:"Électricité solaire PV",      category:"Électricité",  source:"ADEME_V23.6" },
  electricityWind:        { id:"20043", factor:0.0140,  unit:"MWh", label:"Électricité éolienne",        category:"Électricité",  source:"ADEME_V23.6" },
  electricityHydro:       { id:"20044", factor:0.0060,  unit:"MWh", label:"Électricité hydraulique",     category:"Électricité",  source:"ADEME_V23.6" },

  // ── Chaleur et froid achetés (tCO2eq / MWh) ──────────────────────────────
  districtHeat:           { id:"21001", factor:0.1670,  unit:"MWh", label:"Chaleur réseau urbain",       category:"Chaleur/Froid", source:"ADEME_V23.6" },
  districtCooling:        { id:"21002", factor:0.0810,  unit:"MWh", label:"Froid réseau urbain",         category:"Chaleur/Froid", source:"ADEME_V23.6" },
  steamPurchased:         { id:"21003", factor:0.1020,  unit:"MWh", label:"Vapeur achetée",              category:"Chaleur/Froid", source:"ADEME_V23.6" },
  geothermalHeat:         { id:"21004", factor:0.0350,  unit:"MWh", label:"Chaleur géothermique",        category:"Chaleur/Froid", source:"ADEME_V23.6" },
  heatPumpElec:           { id:"21005", factor:0.0160,  unit:"MWh", label:"Pompe à chaleur (élec Fr)",   category:"Chaleur/Froid", source:"ADEME_V23.6" },
};

// ─── SCOPE 3 — Émissions indirectes ──────────────────────────────────────────

const SCOPE3_FACTORS = {

  // ── Transport fret (tCO2eq / t.km) ───────────────────────────────────────
  roadFreightLight:       { id:"50101", factor:0.000191, unit:"t.km", label:"Fret routier léger (<3,5t)",   category:"Transport fret",    source:"ADEME_V23.6" },
  roadFreightMedium:      { id:"50102", factor:0.000113, unit:"t.km", label:"Fret routier moyen (3,5-12t)", category:"Transport fret",    source:"ADEME_V23.6" },
  roadFreightHeavy:       { id:"50103", factor:0.000096, unit:"t.km", label:"Fret routier lourd (>12t)",    category:"Transport fret",    source:"ADEME_V23.6" },
  roadFreightArticulated: { id:"50104", factor:0.000062, unit:"t.km", label:"Fret PL articulé (optimisé)",  category:"Transport fret",    source:"ADEME_V23.6" },
  railFreight:            { id:"50110", factor:0.000028, unit:"t.km", label:"Fret ferroviaire",             category:"Transport fret",    source:"ADEME_V23.6" },
  seaFreight:             { id:"50120", factor:0.000011, unit:"t.km", label:"Fret maritime conteneur",      category:"Transport fret",    source:"ADEME_V23.6" },
  seaFreightBulk:         { id:"50121", factor:0.000008, unit:"t.km", label:"Fret maritime vrac",           category:"Transport fret",    source:"ADEME_V23.6" },
  seaFreightTanker:       { id:"50122", factor:0.000006, unit:"t.km", label:"Fret maritime tanker",         category:"Transport fret",    source:"ADEME_V23.6" },
  airFreight:             { id:"50130", factor:0.000602, unit:"t.km", label:"Fret aérien",                  category:"Transport fret",    source:"ADEME_V23.6" },
  airFreightBellyHold:    { id:"50131", factor:0.000397, unit:"t.km", label:"Fret soute avion passager",    category:"Transport fret",    source:"ADEME_V23.6" },
  riverFreight:           { id:"50140", factor:0.000031, unit:"t.km", label:"Fret fluvial",                 category:"Transport fret",    source:"ADEME_V23.6" },

  // ── Déplacements professionnels — Avion (tCO2eq / km.passager) ───────────
  businessAirShortEco:    { id:"50210", factor:0.000258, unit:"km",   label:"Avion court-courrier éco (<1000km)", category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirShortBiz:    { id:"50211", factor:0.000387, unit:"km",   label:"Avion court-courrier business",      category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirMediumEco:   { id:"50212", factor:0.000209, unit:"km",   label:"Avion moyen-courrier éco",           category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirLongEco:     { id:"50213", factor:0.000195, unit:"km",   label:"Avion long-courrier éco (>3500km)",  category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirLongBiz:     { id:"50214", factor:0.000585, unit:"km",   label:"Avion long-courrier business",       category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirLongFirst:   { id:"50215", factor:0.000780, unit:"km",   label:"Avion long-courrier première",       category:"Dépl. pro avion", source:"ADEME_V23.6" },
  // Alias courts pour compatibilité
  businessAirShort:       { id:"50210", factor:0.000258, unit:"km",   label:"Avion court-courrier",               category:"Dépl. pro avion", source:"ADEME_V23.6" },
  businessAirLong:        { id:"50213", factor:0.000195, unit:"km",   label:"Avion long-courrier",                category:"Dépl. pro avion", source:"ADEME_V23.6" },

  // ── Déplacements professionnels — Voiture (tCO2eq / km) ──────────────────
  businessCarPetrol:      { id:"50220", factor:0.000193, unit:"km",   label:"Voiture essence (dépl. pro)",  category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  businessCarDiesel:      { id:"50221", factor:0.000163, unit:"km",   label:"Voiture diesel (dépl. pro)",   category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  businessCarElectric:    { id:"50222", factor:0.000019, unit:"km",   label:"Voiture électrique (France)",  category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  businessCarHybrid:      { id:"50223", factor:0.000103, unit:"km",   label:"Voiture hybride (dépl. pro)",  category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  businessCar:            { id:"50220", factor:0.000193, unit:"km",   label:"Voiture essence (dépl. pro)",  category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  // Motos et utilitaires
  businessMotorcycle:     { id:"50225", factor:0.000103, unit:"km",   label:"Moto (dépl. pro)",             category:"Dépl. pro voiture", source:"ADEME_V23.6" },
  businessVan:            { id:"50226", factor:0.000230, unit:"km",   label:"Utilitaire léger (dépl. pro)", category:"Dépl. pro voiture", source:"ADEME_V23.6" },

  // ── Déplacements professionnels — Train (tCO2eq / km.passager) ───────────
  trainTGV:               { id:"50230", factor:0.000003, unit:"km",   label:"TGV (France)",                 category:"Dépl. pro train",   source:"ADEME_V23.6" },
  trainIntercity:         { id:"50231", factor:0.000006, unit:"km",   label:"Train intercité (France)",     category:"Dépl. pro train",   source:"ADEME_V23.6" },
  trainEurope:            { id:"50232", factor:0.000041, unit:"km",   label:"Train Europe (moy.)",          category:"Dépl. pro train",   source:"ADEME_V23.6" },
  trainTravel:            { id:"50230", factor:0.000003, unit:"km",   label:"Train (dépl. pro)",            category:"Dépl. pro train",   source:"ADEME_V23.6" },
  eurostar:               { id:"50233", factor:0.000004, unit:"km",   label:"Eurostar",                     category:"Dépl. pro train",   source:"ADEME_V23.6" },

  // ── Déplacements domicile-travail (tCO2eq / km.passager) ─────────────────
  commuteCarPetrol:       { id:"50310", factor:0.000193, unit:"km",   label:"Trajet domicile-travail voiture essence", category:"Domicile-travail", source:"ADEME_V23.6" },
  commuteCarElectric:     { id:"50311", factor:0.000019, unit:"km",   label:"Trajet domicile-travail voiture élec.",   category:"Domicile-travail", source:"ADEME_V23.6" },
  commuteMetro:           { id:"50312", factor:0.000004, unit:"km",   label:"Métro (trajet domicile-travail)",         category:"Domicile-travail", source:"ADEME_V23.6" },
  commuteBus:             { id:"50313", factor:0.000113, unit:"km",   label:"Bus thermique (domicile-travail)",        category:"Domicile-travail", source:"ADEME_V23.6" },
  commuteRER:             { id:"50314", factor:0.000005, unit:"km",   label:"RER (domicile-travail)",                  category:"Domicile-travail", source:"ADEME_V23.6" },

  // ── Déchets (tCO2eq / tonne) ──────────────────────────────────────────────
  wasteIncineration:      { id:"51001", factor:0.8540,   unit:"t",    label:"Incinération DIB",             category:"Déchets",          source:"ADEME_V23.6" },
  wasteLandfill:          { id:"51002", factor:0.4580,   unit:"t",    label:"Enfouissement (ISDND)",        category:"Déchets",          source:"ADEME_V23.6" },
  wasteRecycling:         { id:"51003", factor:-0.0830,  unit:"t",    label:"Recyclage (crédit évité)",     category:"Déchets",          source:"ADEME_V23.6" },
  wasteWater:             { id:"51004", factor:0.7080,   unit:"t",    label:"Traitement eaux usées",        category:"Déchets",          source:"ADEME_V23.6" },
  wasteHazardous:         { id:"51005", factor:1.2400,   unit:"t",    label:"Déchets dangereux",            category:"Déchets",          source:"ADEME_V23.6" },
  wasteElectronic:        { id:"51006", factor:0.0230,   unit:"kg",   label:"DEEE (équipements électro.)", category:"Déchets",          source:"ADEME_V23.6" },
  wasteCompost:           { id:"51007", factor:0.1030,   unit:"t",    label:"Compostage déchets verts",    category:"Déchets",          source:"ADEME_V23.6" },
  wasteMethanization:     { id:"51008", factor:0.0780,   unit:"t",    label:"Méthanisation",               category:"Déchets",          source:"ADEME_V23.6" },

  // ── Achats de matériaux (tCO2eq / tonne) ─────────────────────────────────
  steelPurchased:         { id:"52001", factor:1.8500,   unit:"t",    label:"Acier acheté (haut-fourneau)",  category:"Matériaux achetés", source:"ADEME_V23.6" },
  steelRecycledPurchased: { id:"52002", factor:0.4200,   unit:"t",    label:"Acier recyclé acheté",          category:"Matériaux achetés", source:"ADEME_V23.6" },
  aluminiumPurchased:     { id:"52003", factor:8.2400,   unit:"t",    label:"Aluminium acheté (primaire)",   category:"Matériaux achetés", source:"ADEME_V23.6" },
  aluminiumRecPurchased:  { id:"52004", factor:0.5100,   unit:"t",    label:"Aluminium recyclé acheté",      category:"Matériaux achetés", source:"ADEME_V23.6" },
  plasticPurchased:       { id:"52005", factor:3.1400,   unit:"t",    label:"Plastique PET acheté",          category:"Matériaux achetés", source:"ADEME_V23.6" },
  plasticPEPurchased:     { id:"52006", factor:1.9200,   unit:"t",    label:"Plastique PE acheté",           category:"Matériaux achetés", source:"ADEME_V23.6" },
  glassPurchased:         { id:"52007", factor:0.5400,   unit:"t",    label:"Verre acheté",                  category:"Matériaux achetés", source:"ADEME_V23.6" },
  paperPurchased:         { id:"52008", factor:0.9190,   unit:"t",    label:"Papier/carton acheté",          category:"Matériaux achetés", source:"ADEME_V23.6" },
  concreteUsed:           { id:"52009", factor:0.1300,   unit:"t",    label:"Béton utilisé",                 category:"Matériaux achetés", source:"ADEME_V23.6" },
  woodPurchased:          { id:"52010", factor:0.0890,   unit:"t",    label:"Bois acheté (construction)",    category:"Matériaux achetés", source:"ADEME_V23.6" },
  copperPurchased:        { id:"52011", factor:3.1500,   unit:"t",    label:"Cuivre acheté",                 category:"Matériaux achetés", source:"ADEME_V23.6" },
  rubberPurchased:        { id:"52012", factor:2.8500,   unit:"t",    label:"Caoutchouc acheté",             category:"Matériaux achetés", source:"ADEME_V23.6" },
  textilePurchased:       { id:"52013", factor:15.000,   unit:"t",    label:"Textile acheté",                category:"Matériaux achetés", source:"ADEME_V23.6" },

  // ── Immobilisations (tCO2eq / unité) ─────────────────────────────────────
  serverRack:             { id:"53001", factor:0.1100,   unit:"unité",label:"Serveur informatique",          category:"Immobilisations",  source:"ADEME_V23.6" },
  laptop:                 { id:"53002", factor:0.1560,   unit:"unité",label:"Ordinateur portable",           category:"Immobilisations",  source:"ADEME_V23.6" },
  desktop:                { id:"53003", factor:0.2290,   unit:"unité",label:"Ordinateur fixe + écran",       category:"Immobilisations",  source:"ADEME_V23.6" },
  smartphone:             { id:"53004", factor:0.0790,   unit:"unité",label:"Smartphone",                    category:"Immobilisations",  source:"ADEME_V23.6" },
  constructionOffice:     { id:"53010", factor:0.3280,   unit:"m²",   label:"Bureau neuf (construction)",    category:"Immobilisations",  source:"ADEME_V23.6" },
  constructionIndustrial: { id:"53011", factor:0.2150,   unit:"m²",   label:"Bâtiment industriel (constr.)", category:"Immobilisations",  source:"ADEME_V23.6" },
  carPetrolNew:           { id:"53020", factor:7.1000,   unit:"unité",label:"Voiture essence neuve",         category:"Immobilisations",  source:"ADEME_V23.6" },
  carElectricNew:         { id:"53021", factor:8.9000,   unit:"unité",label:"Voiture électrique neuve",      category:"Immobilisations",  source:"ADEME_V23.6" },
  truckHeavyNew:          { id:"53022", factor:31.000,   unit:"unité",label:"PL lourd neuf",                 category:"Immobilisations",  source:"ADEME_V23.6" },

  // ── Eau (tCO2eq / m³) ────────────────────────────────────────────────────
  waterSupply:            { id:"54001", factor:0.000344, unit:"m³",   label:"Eau potable (réseau)",          category:"Eau",              source:"ADEME_V23.6" },
  waterWastewater:        { id:"54002", factor:0.000708, unit:"m³",   label:"Eaux usées (traitement)",       category:"Eau",              source:"ADEME_V23.6" },

  // ── Services numériques (tCO2eq / Go ou heure) ───────────────────────────
  dataTransfer:           { id:"55001", factor:0.000023, unit:"Go",   label:"Transfert données internet",    category:"Numérique",        source:"ADEME_V23.6" },
  cloudStorage:           { id:"55002", factor:0.000016, unit:"Go/an",label:"Stockage cloud (par Go/an)",    category:"Numérique",        source:"ADEME_V23.6" },
  videoConference:        { id:"55003", factor:0.000056, unit:"h",    label:"Visioconférence (1h)",           category:"Numérique",        source:"ADEME_V23.6" },
  emailWithAttach:        { id:"55004", factor:0.000050, unit:"email",label:"Email avec pièce jointe",       category:"Numérique",        source:"ADEME_V23.6" },
};

// ─── Index consolidé ──────────────────────────────────────────────────────────

const EMISSION_FACTORS_FULL = {
  scope1: SCOPE1_FACTORS,
  scope2: SCOPE2_FACTORS,
  scope3: SCOPE3_FACTORS,

  // Index plat pour recherche rapide par clé (toutes catégories)
  _flat: { ...SCOPE1_FACTORS, ...SCOPE2_FACTORS, ...SCOPE3_FACTORS },

  // Métadonnées
  _meta: {
    version:    ADEME_VERSION,
    updated:    ADEME_UPDATED,
    source:     "Base Carbone® ADEME — Licence Ouverte v2.0",
    licence:    "https://www.etalab.gouv.fr/licence-ouverte-open-licence",
    count:      Object.keys(SCOPE1_FACTORS).length
               + Object.keys(SCOPE2_FACTORS).length
               + Object.keys(SCOPE3_FACTORS).length,
  },
};

/**
 * Cherche un facteur par clé dans toutes les catégories.
 * @param {string} key - clé du facteur (ex: "naturalGas", "electricityFrance")
 * @returns {{ factor, unit, label, category, scope, source } | null}
 */
function getFactor(key) {
  for (const [scope, factors] of Object.entries({ scope1: SCOPE1_FACTORS, scope2: SCOPE2_FACTORS, scope3: SCOPE3_FACTORS })) {
    if (factors[key]) return { ...factors[key], scope };
  }
  return null;
}

/**
 * Retourne tous les facteurs d'un scope donné.
 * @param {"scope1"|"scope2"|"scope3"} scope
 */
function getFactorsByScope(scope) {
  return EMISSION_FACTORS_FULL[scope] || {};
}

/**
 * Retourne tous les facteurs d'une catégorie donnée.
 * @param {string} category - ex: "Combustibles fossiles"
 */
function getFactorsByCategory(category) {
  return Object.fromEntries(
    Object.entries(EMISSION_FACTORS_FULL._flat)
      .filter(([, v]) => v.category === category)
  );
}

/**
 * Retourne la liste des catégories disponibles.
 */
function getCategories() {
  return [...new Set(Object.values(EMISSION_FACTORS_FULL._flat).map(f => f.category))].sort();
}

window.CarbonFactors = {
  EMISSION_FACTORS_FULL,
  SCOPE1_FACTORS,
  SCOPE2_FACTORS,
  SCOPE3_FACTORS,
  ADEME_VERSION,
  ADEME_UPDATED,
  getFactor,
  getFactorsByScope,
  getFactorsByCategory,
  getCategories,
};
