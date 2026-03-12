/**
 * deploy-ghpages.js — Fix ENAMETOOLONG Windows
 * Placer à la racine du projet (même niveau que package.json)
 */

const { execSync } = require("child_process");
const fs   = require("fs");
const path = require("path");

// __dirname = dossier où se trouve CE script = racine du projet
const PROJECT_ROOT = __dirname;
const DIST_DIR     = path.join(PROJECT_ROOT, "dist");
const BRANCH       = "gh-pages";
const COMMIT_MSG   = `deploy: ${new Date().toISOString()}`;

function run(cmd, opts = {}) {
  console.log(`  > ${cmd}`);
  execSync(cmd, { stdio: "inherit", cwd: PROJECT_ROOT, ...opts });
}

function runCapture(cmd) {
  return execSync(cmd, { encoding: "utf8", cwd: PROJECT_ROOT }).trim();
}

// ── 1. Vérifier dist/ AVANT de toucher à git ────────────────
if (!fs.existsSync(DIST_DIR)) {
  console.error("❌  dist/ introuvable. Lancez d'abord : npm run build");
  process.exit(1);
}

const distFiles = fs.readdirSync(DIST_DIR);
if (distFiles.length === 0) {
  console.error("❌  dist/ est vide. Relancez : npm run build");
  process.exit(1);
}

console.log(`✅  dist/ trouvé (${distFiles.length} fichiers)`);

// ── 2. Activer chemins longs ─────────────────────────────────
try { run("git config core.longpaths true"); } catch (_) {}

// ── 3. Remote origin ─────────────────────────────────────────
let remoteUrl;
try {
  remoteUrl = runCapture("git remote get-url origin");
  console.log(`🔗  Remote : ${remoteUrl}`);
} catch (_) {
  console.error("❌  Pas de remote 'origin'. Faites : git remote add origin <URL>");
  process.exit(1);
}

// ── 4. Branche courante (pour revenir après) ─────────────────
let currentBranch = "main";
try { currentBranch = runCapture("git branch --show-current"); } catch (_) {}
console.log(`📌  Branche actuelle : ${currentBranch}`);

// ── 5. Sauvegarder dist/ en mémoire ─────────────────────────
//    On lit tout maintenant, AVANT de changer de branche
console.log("\n📦  Lecture du contenu de dist/...");

function readDirRecursive(dir) {
  const result = {};
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      result[entry.name] = readDirRecursive(fullPath);
    } else {
      result[entry.name] = fs.readFileSync(fullPath);
      console.log(`   ✓ ${entry.name}`);
    }
  }
  return result;
}

const distContents = readDirRecursive(DIST_DIR);

// ── 6. Fonction pour écrire le contenu en mémoire sur disque ─
function writeDirContents(contents, destDir) {
  fs.mkdirSync(destDir, { recursive: true });
  for (const [name, value] of Object.entries(contents)) {
    const dest = path.join(destDir, name);
    if (Buffer.isBuffer(value)) {
      fs.writeFileSync(dest, value);
    } else {
      writeDirContents(value, dest);
    }
  }
}

// ── 7. Basculer sur gh-pages (orpheline) ─────────────────────
console.log(`\n🌿  Création de la branche orpheline ${BRANCH}...`);

try {
  // Stash les changements non commités si besoin
  try { run("git stash --include-untracked"); } catch(_) {}

  run(`git checkout --orphan ${BRANCH}`);

  // Supprimer TOUT (index git) — les fichiers physiques restent
  try { run("git rm -rf . --quiet"); } catch (_) {}

  // Supprimer les fichiers physiques restants (hors .git)
  for (const entry of fs.readdirSync(PROJECT_ROOT, { withFileTypes: true })) {
    if (entry.name === ".git") continue;
    const p = path.join(PROJECT_ROOT, entry.name);
    fs.rmSync(p, { recursive: true, force: true });
  }

} catch (err) {
  console.error("❌  Erreur lors du checkout gh-pages :", err.message);
  try { execSync(`git checkout ${currentBranch}`, { cwd: PROJECT_ROOT, stdio: "pipe" }); } catch(_) {}
  try { execSync("git stash pop", { cwd: PROJECT_ROOT, stdio: "pipe" }); } catch(_) {}
  process.exit(1);
}

// ── 8. Écrire le contenu de dist/ à la racine ────────────────
console.log("\n📁  Écriture des fichiers...");
try {
  writeDirContents(distContents, PROJECT_ROOT);
  fs.writeFileSync(path.join(PROJECT_ROOT, ".nojekyll"), "");
  console.log("   ✓ .nojekyll");
} catch (err) {
  console.error("❌  Erreur écriture fichiers :", err.message);
  try { execSync(`git checkout ${currentBranch}`, { cwd: PROJECT_ROOT, stdio: "pipe" }); } catch(_) {}
  process.exit(1);
}

// ── 9. Git add + commit (fichier par fichier) ─────────────────
console.log("\n📝  Commit...");
try {
  const toAdd = [...Object.keys(distContents), ".nojekyll"];
  for (const f of toAdd) {
    try { run(`git add "${f}"`); } catch (_) {}
  }
  run(`git commit -m "${COMMIT_MSG}"`);
} catch (err) {
  console.error("❌  Erreur commit :", err.message);
  try { execSync(`git checkout ${currentBranch}`, { cwd: PROJECT_ROOT, stdio: "pipe" }); } catch(_) {}
  process.exit(1);
}

// ── 10. Push force ───────────────────────────────────────────
console.log("\n📤  Push...");
try {
  run(`git push origin ${BRANCH} --force`);
} catch (err) {
  console.error("❌  Erreur push :", err.message);
  try { execSync(`git checkout ${currentBranch}`, { cwd: PROJECT_ROOT, stdio: "pipe" }); } catch(_) {}
  process.exit(1);
}

// ── 11. Retour sur la branche principale ─────────────────────
console.log(`\n🔙  Retour sur ${currentBranch}...`);
run(`git checkout ${currentBranch}`);
try { run("git stash pop"); } catch(_) {} // restaurer stash si besoin

console.log(`
✅  Déploiement réussi sur la branche ${BRANCH} !

Prochaines étapes :
  1. GitHub → Settings → Pages → Source : branche "${BRANCH}", dossier "/ (root)"
  2. Votre add-in sera disponible sur :
     https://miensie.github.io/ESG/taskpane.html
  3. Mettez à jour YOUR-DOMAIN dans manifest.xml avec cette URL
`);
