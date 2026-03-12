/**
 * deploy-ghpages.js
 * Alternative à gh-pages qui évite l'erreur ENAMETOOLONG sur Windows
 * Usage : node deploy-ghpages.js
 */

const { execSync } = require("child_process");
const fs   = require("fs");
const path = require("path");

const DIST_DIR   = path.resolve(__dirname, "dist");
const BRANCH     = "gh-pages";
const COMMIT_MSG = `deploy: ${new Date().toISOString()}`;

function run(cmd, opts = {}) {
  console.log(`  > ${cmd}`);
  return execSync(cmd, { stdio: "inherit", ...opts });
}

function runCapture(cmd) {
  return execSync(cmd, { encoding: "utf8" }).trim();
}

// 1. Vérifier que dist/ existe
if (!fs.existsSync(DIST_DIR)) {
  console.error("❌ Le dossier dist/ n'existe pas. Lancez d'abord : npm run build");
  process.exit(1);
}

// 2. Activer le support des chemins longs sur Windows
try {
  run("git config core.longpaths true");
} catch (_) {}

// 3. Récupérer l'URL remote origin
let remoteUrl;
try {
  remoteUrl = runCapture("git remote get-url origin");
  console.log(`🔗 Remote : ${remoteUrl}`);
} catch (_) {
  console.error("❌ Pas de remote 'origin' configuré. Faites : git remote add origin <URL>");
  process.exit(1);
}

// 4. Sauvegarder la branche courante
let currentBranch;
try {
  currentBranch = runCapture("git branch --show-current");
} catch (_) {
  currentBranch = "main";
}

console.log(`\n🚀 Déploiement vers la branche ${BRANCH}...\n`);

// 5. Vérifier si la branche gh-pages existe déjà
let branchExists = false;
try {
  runCapture(`git show-ref --verify refs/heads/${BRANCH}`);
  branchExists = true;
} catch (_) {}

try {
  if (branchExists) {
    // 6a. Basculer sur gh-pages et y copier dist/
    run(`git checkout ${BRANCH}`);
    
    // Supprimer tous les fichiers trackés sauf .git
    try {
      run("git rm -rf . --quiet");
    } catch (_) {}

  } else {
    // 6b. Créer une branche orpheline gh-pages
    run(`git checkout --orphan ${BRANCH}`);
    try { run("git rm -rf . --quiet"); } catch (_) {}
  }

  // 7. Copier les fichiers de dist/ à la racine
  console.log("\n📁 Copie des fichiers dist/ → racine...");
  const files = fs.readdirSync(DIST_DIR);
  for (const file of files) {
    const src  = path.join(DIST_DIR, file);
    const dest = path.join(process.cwd(), file);
    fs.cpSync(src, dest, { recursive: true });
    console.log(`   ✓ ${file}`);
  }

  // 8. Ajouter un .nojekyll pour désactiver Jekyll GitHub Pages
  fs.writeFileSync(".nojekyll", "");
  console.log("   ✓ .nojekyll");

  // 9. Committer fichier par fichier (évite ENAMETOOLONG)
  console.log("\n📝 Commit des fichiers (un par un)...");
  for (const file of [...files, ".nojekyll"]) {
    try {
      run(`git add "${file}"`, { stdio: "pipe" });
    } catch (_) {}
  }

  run(`git commit -m "${COMMIT_MSG}"`);

  // 10. Push
  console.log("\n📤 Push vers origin gh-pages...");
  run(`git push origin ${BRANCH} --force`);

  // 11. Revenir sur la branche d'origine
  run(`git checkout ${currentBranch}`);

  console.log(`\n✅ Déploiement réussi !`);
  console.log(`   Branche : ${BRANCH}`);
  console.log(`   Activez GitHub Pages dans Settings → Pages → Source : ${BRANCH}`);

} catch (err) {
  console.error("\n❌ Erreur pendant le déploiement :", err.message);
  // Tenter de revenir sur la branche d'origine
  try { execSync(`git checkout ${currentBranch}`, { stdio: "pipe" }); } catch (_) {}
  process.exit(1);
}
