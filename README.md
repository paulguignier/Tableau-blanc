# 📘 Mémo – Projet TypeScript pour ExcelScript/Automate

## ⚙️ Étapes pour démarrer un nouveau projet

1. 📁 Duplique ce dossier `ModeleProjetTS` et renomme-le selon ton projet.

2. 📦 Ouvre un terminal dans le dossier du projet et installe les dépendances :

```bash
   npm install
```

3. 🛠 Installer TypeScript (si pas encore fait) :

```bash
   npm install typescript --save-dev
```

4. 🏗 Compiler une seule fois :

```bash
   npx tsc
```

   Cela compile tous les fichiers `.ts` de `src/` vers `dist/`.

5. 👀 Compiler automatiquement en continu :

```bash
   npx tsc --watch
```

   > **Pour arrêter :**  
   > Appuyer sur `Ctrl + C`, puis taper `Y`.

6. 🚀 Lancer le projet :

```bash
   node dist/main.js
```

   *(remplacer `main.js` par ton fichier de sortie si besoin)*

7. 🔁 Option : Ajouter des scripts dans `package.json` :

```json
   "scripts": {
     "build": "tsc",
     "watch": "tsc --watch",
     "start": "node dist/main.js"
   }
```

   Utilisation ensuite :

   - Compiler une fois : `npm run build`
   - Surveillance automatique : `npm run watch`
   - Démarrer le projet : `npm run start`

