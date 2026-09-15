# Showbasen

En fristående webbapp — samma app som byggdes i Claude-chatten, nu redo att köras
utanför den. Datan sparas i webbläsarens `localStorage`, lokalt på den dator/
webbläsare du använder. Ingen server, inget konto, ingen molntjänst.

## Testa lokalt först

Kräver [Node.js](https://nodejs.org) (valfri LTS-version).

```
npm install
npm run dev
```

Öppna länken terminalen visar (oftast http://localhost:5173). Testa att lägga
till ett trick, spara ett uppdrag, ladda om sidan — datan ska ligga kvar.

## Lägg upp på GitHub Pages

1. Skapa ett nytt repo på GitHub, t.ex. `showbasen` (kan vara privat eller
   publikt — koden syns för besökare om det är publikt, men din sparade data
   gör det aldrig, den lever bara i din egen webbläsare).
2. Ladda upp alla filer i den här mappen till repot (via GitHub Desktop, VS
   Code, eller `git`-kommandona nedan).
3. Bygg produktionsversionen:
   ```
   npm run build
   ```
   Det skapar en `dist`-mapp med den färdiga, statiska sajten.
4. Publicera `dist`-mappen som Pages. Enklaste sättet är grenen `gh-pages`:
   ```
   npm install -D gh-pages
   npx gh-pages -d dist
   ```
   Gå sedan till repots **Settings → Pages** och välj grenen `gh-pages` som
   källa (om det inte redan väljs automatiskt).
5. Efter någon minut är appen tillgänglig på
   `https://<ditt-användarnamn>.github.io/showbasen/`.

### Om du hellre kör allt via git manuellt

```
git init
git add .
git commit -m "Showbasen"
git branch -M main
git remote add origin https://github.com/<ditt-användarnamn>/showbasen.git
git push -u origin main
```

Kör sedan byggsteget och `gh-pages`-kommandot ovan.

## Installera som skrivbordsikon (Windows/Mac)

När sidan är live: öppna den i Edge eller Chrome, klicka på
installationsikonen i adressfältet (eller menyn → "Installera appen"). Då får
du ett eget fönster och en ikon, som en vanlig app, fast det fortfarande är
webbsidan som körs.

## Säkerhetskopiera innan du byter dator/webbläsare

`localStorage` följer inte med om du byter dator, webbläsare, eller rensar
webbläsardata. Använd appens egen **Backup**-funktion under Inställningar
(Exportera backup) innan du gör något sådant, och Importera backup för att
återställa.

## Filstruktur

- `src/App.jsx` — hela appen, samma kod som i Claude-chatten
- `src/storage.js` — ersätter Claude-artifactens `window.storage` med en
  `localStorage`-baserad motsvarighet med samma gränssnitt
- `src/main.jsx` — startpunkt som renderar appen
- `vite.config.js` — byggkonfiguration, satt för att fungera på GitHub Pages
