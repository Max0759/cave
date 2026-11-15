# Relais IA « Ma Cave » : lecture des étiquettes

L'appli est publique (GitHub Pages) : la clé de l'API Claude ne peut pas y être
écrite. Ce petit relais gratuit (Cloudflare Workers) garde la clé secrète, reçoit
la photo de l'étiquette et renvoie le vin reconnu.

Coût : environ 1 à 2 centimes par photo (API Claude, modèle Claude Opus 5.5).
Cloudflare Workers est gratuit jusqu'à 100 000 appels par jour.

## Installation (une seule fois, ~15 minutes, depuis un ordinateur)

1. **Clé API Claude** : créez un compte sur https://platform.claude.com, ajoutez
   quelques euros de crédit, puis créez une clé API (Settings → API keys).
   Conseil : fixez une limite de dépense mensuelle (par ex. 5 €).
2. **Compte Cloudflare** gratuit : https://dash.cloudflare.com/sign-up
3. **Node.js** installé sur l'ordinateur : https://nodejs.org (version LTS).
4. Dans un terminal, dans ce dossier `worker/` :

   ```bash
   npm install
   npx wrangler login
   npx wrangler secret put ANTHROPIC_API_KEY
   npx wrangler secret put CAVE_CODE
   npx wrangler deploy
   ```

   - `ANTHROPIC_API_KEY` : collez la clé de l'étape 1.
   - `CAVE_CODE` : inventez un code d'accès (une phrase longue, par ex.
     `cave-ucel-2026-tire-bouchon`). Il empêche d'autres personnes d'utiliser
     votre relais.
   - `deploy` affiche l'adresse du relais, du type
     `https://cave-etiquette.<votre-nom>.workers.dev`.

5. **Dans l'appli**, sur chaque téléphone : onglet **Bilan → Lecture des
   étiquettes → Réglages**, collez l'adresse du relais et le code d'accès.

## Utilisation

Bouton **📷 Scanner** (en bas à gauche) → photo de l'étiquette → l'appli
retrouve le vin dans la cave et propose :

- **🍷 Bu (−1)** ou **+ Ajouter** une bouteille ;
- **Compléter la fiche** avec les infos lues (cépages, garde, prix indicatif) ;
- **Nouveau vin** pré-rempli si la bouteille n'est pas encore dans la cave.

La photo peut servir de photo de la fiche (case à cocher).

## Bon à savoir

- L'adresse de l'appli autorisée est dans `wrangler.toml` (`ORIGINE_AUTORISEE`).
- Sans réseau à la cave, utilisez « Chercher à la main » : le scan a besoin
  d'internet.
- La photo est envoyée à l'API Claude uniquement pour la lecture ; seule la
  version réduite est gardée dans votre OneDrive si vous la conservez pour la fiche.
