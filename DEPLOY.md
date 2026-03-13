# Déploiement du backend ESG Analyzer Pro

## Architecture

```
Client Excel (GitHub Pages)
        │
        │  POST /api/ai/analyze
        │  POST /api/ai/chat
        ▼
Votre backend Node.js          ← clé Gemini ici, invisible pour le client
(Render / Railway / VPS)
        │
        │  API Gemini
        ▼
Google AI Studio (Gemini 2.0 Flash)
```

---

## Option A — Render.com (recommandé, gratuit)

1. Créer un compte sur https://render.com
2. New → Web Service → connecter votre repo GitHub
3. Paramètres :
   - **Root Directory** : `backend`
   - **Build Command** : `npm install`
   - **Start Command** : `node server.js`
   - **Instance Type** : Free
4. Variables d'environnement (onglet Environment) :
   ```
   GEMINI_API_KEY = AIza_votre_cle
   ALLOWED_ORIGIN = https://VOTRE_USER.github.io
   ```
5. Deploy → votre URL sera : `https://esg-analyzer-XXXX.onrender.com`
6. Mettre à jour `geminiAI.js` ligne 11 :
   ```js
   : "https://esg-analyzer-XXXX.onrender.com"
   ```

> ⚠️ Plan gratuit Render : le serveur "dort" après 15 min d'inactivité.
> Premier appel = 30s de délai. Passez au plan Starter ($7/mois) pour éviter ça.

---

## Option B — Railway.app

1. Compte sur https://railway.app
2. New Project → Deploy from GitHub repo
3. Sélectionner le dossier `backend/`
4. Variables d'environnement :
   ```
   GEMINI_API_KEY = AIza_votre_cle
   ALLOWED_ORIGIN = https://VOTRE_USER.github.io
   PORT = 3000
   ```
5. Generate Domain → noter l'URL
6. Gratuit jusqu'à 500h/mois (environ 20 jours)

---

## Option C — VPS (Hetzner/OVH, ~4€/mois)

```bash
# Sur le serveur
git clone https://github.com/VOUS/ESG.git
cd ESG/backend
cp .env.example .env
nano .env          # Remplir GEMINI_API_KEY et ALLOWED_ORIGIN

npm install
npm install -g pm2
pm2 start server.js --name esg-backend
pm2 startup && pm2 save

# HTTPS avec Nginx + Certbot
sudo apt install nginx certbot python3-certbot-nginx
sudo certbot --nginx -d api.votredomaine.fr
```

---

## Test de santé

```bash
curl https://VOTRE_BACKEND/api/health
# Réponse attendue :
# {"status":"ok","version":"2.0.0","gemini":"configuré","timestamp":"..."}
```

---

## Sécurité

- La clé Gemini n'est **jamais** dans le code source ni dans les variables côté client
- Le CORS bloque automatiquement toute origine autre que votre GitHub Pages
- Rate limiting : 30 appels IA / IP / heure, 5 / seconde
- Optionnel : activez `API_SECRET` dans `.env` pour une protection supplémentaire
