# Pointage — Carnet de bord d'alternance

Petit outil pour tracker tes activités heure par heure dans la semaine et exporter un Excel à donner au tuteur.

- Grille Lun-Ven × 8h-18h, créneaux d'une heure
- Autocomplete sur ce que t'as déjà tapé
- Import / export `.xlsx` — un onglet par semaine, accumulés dans le même fichier
- Navigation entre semaines, mise en évidence d'aujourd'hui

## En local

```bash
npm install
npm run dev
```

Puis ouvre l'URL affichée (typiquement `http://localhost:5173`).

## Déployer sur GitHub Pages

1. **Crée un repo** GitHub (public ou privé peu importe pour la mécanique, mais Pages gratuit n'est dispo que pour les repos publics ou les comptes Pro).

2. **Push le code** :
   ```bash
   git init
   git add .
   git commit -m "init pointage app"
   git branch -M main
   git remote add origin git@github.com:<ton-user>/<nom-du-repo>.git
   git push -u origin main
   ```

3. **Active GitHub Pages** : dans Settings → Pages → Source = `GitHub Actions`.

4. Le workflow `.github/workflows/deploy.yml` se déclenche au push sur `main`. À la fin, ton site est sur `https://<ton-user>.github.io/<nom-du-repo>/`.

> Le `base` de Vite est calculé automatiquement à partir du nom du repo (variable `GITHUB_REPOSITORY` injectée par Actions). Pas besoin de toucher à `vite.config.js`.

## Format Excel

Chaque semaine est un onglet nommé `S{numéro}-{date du lundi}` (ex : `S19-2026-05-04`).

Layout d'un onglet :

| Heure   | Lundi 04.05 | Mardi 05.05 | Mercredi 06.05 | Jeudi 07.05 | Vendredi 08.05 |
|---------|-------------|-------------|----------------|-------------|----------------|
| 08h-09h | …           | …           | …              | …           | …              |
| …       |             |             |                |             |                |

Quand tu ré-uploades le fichier, l'app détecte les onglets existants, charge leur contenu, et te place sur la dernière semaine trouvée.

## Modifs faciles

Dans `src/App.jsx`, en haut du fichier :

```js
const WORK_HOURS = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17]; // créneaux affichés
const DAY_NAMES = ['LUN', 'MAR', 'MER', 'JEU', 'VEN'];      // jours
```

Pour ajouter le samedi, ajoute `'SAM'` à `DAY_NAMES` et change le `length: 5` en `length: 6` dans `getWeekDays` (et le `for c < 5` dans `handleUpload`, et `WORK_HOURS.length * 5` pour `totalSlots`). Désolé c'est dispatché un peu partout — un refacto vers une constante `NUM_DAYS` se fait en 2 minutes si t'en as besoin.
