# AssistDMS Add-in - GitHub Pages Deployment

## Deze folder bevat alle bestanden voor GitHub Pages

### Stappen om te deployen:

1. **Maak een nieuwe GitHub repository:**
   - Ga naar https://github.com/new
   - Naam: `assistdms-addin`
   - Public (verplicht voor gratis GitHub Pages)
   - Klik "Create repository"

2. **Upload deze bestanden:**
   - Ga naar je repository
   - Klik "Add file" → "Upload files"
   - Sleep alle bestanden uit deze folder
   - Commit changes

3. **Activeer GitHub Pages:**
   - Ga naar repository Settings
   - Scroll naar "Pages" (linkermenu)
   - Source: Deploy from a branch
   - Branch: main / root
   - Klik Save

4. **Wacht 1-2 minuten**
   - GitHub Pages bouwt je site
   - URL wordt: https://JOUWNAAM.github.io/assistdms-addin/

5. **Update manifest.xml:**
   - Download het aangepaste manifest (wordt gemaakt)
   - Upload het manifest naar Exchange Admin Center

## Bestanden in deze folder:

- `taskpane.html` - Hoofdinterface van de add-in
- `taskpane.js` - JavaScript logica
- `commands.html` - Command functies
- `manifest.xml` - Outlook configuratie (moet nog aangepast!)
- `index.html` - Standalone webapp voor mobiel en pc
- `app.js` - Standalone webapp logica (Azure AD + API)
- `manifest.webmanifest` - PWA installatiebestand
- `sw.js` - Service worker voor snelle app-shell loading
- `icon-*.png` - Iconen

## Let op:

De API moet ook online! Opties:
- Azure App Service (gratis tier)
- Railway.app (gratis tier)  
- Fly.io (gratis tier)
- Ngrok (tijdelijk voor testen)

Of gebruik voor nu localhost en test alleen op je eigen machine.

## Werken op mobiel en pc

Gebruik de standalone webapp via je GitHub Pages URL:

- `https://JOUWNAAM.github.io/assistdms-addin/index.html`

Op mobiel kun je deze installeren als app:

1. Open de URL in Safari of Chrome.
2. Kies `Voeg toe aan beginscherm` of `Install app`.
3. Log in met Microsoft en werk verder zoals op pc.

### Extern tijdschrijfprogramma gebruiken

In `config.js` kun je je bestaande tijdschrijftool invullen:

```js
timeWriterUrl: 'https://jouw-tijdschrijfapp.nl'
```

Daarna verschijnt de knop `Open tijdschrijven` in de header, zowel op mobiel als pc.
