# 🚀 AssistDMS Outlook Add-in - Deployment Instructies

## ✅ Klaar voor deployment naar GitHub Pages!

Alle bestanden zijn aangepast voor: **https://mlcoq.github.io/assistdms-addin/**

---

## 📤 STAP 1: GitHub Repository maken

1. Ga naar: **https://github.com/new**
2. Repository naam: **`assistdms-addin`**
3. ✅ **Public** (verplicht voor gratis GitHub Pages)
4. ❌ **NIET** initialiseren met README (we hebben al bestanden)
5. Klik **"Create repository"**

---

## 📁 STAP 2: Bestanden uploaden

### Optie A: Via GitHub Web Interface (Makkelijkst)

1. In je nieuwe repository, klik **"uploading an existing file"**
2. Sleep **ALLE** bestanden uit de `github-pages-deploy` folder
3. Commit message: "Initial commit - AssistDMS add-in"
4. Klik **"Commit changes"**

### Optie B: Via Git Command Line

```bash
cd "G:\Mijn Drive\Git\AssistDMS\github-pages-deploy"
git init
git add .
git commit -m "Initial commit - AssistDMS add-in"
git branch -M main
git remote add origin https://github.com/mlcoq/assistdms-addin.git
git push -u origin main
```

---

## 🌐 STAP 3: GitHub Pages activeren

1. Ga naar je repository: **https://github.com/mlcoq/assistdms-addin**
2. Klik op **"Settings"** (tabblad bovenaan)
3. Scroll in linkermenu naar **"Pages"**
4. Bij **Source**:
   - Branch: **main**
   - Folder: **/ (root)**
5. Klik **"Save"**
6. ⏱️ Wacht 1-2 minuten...
7. Refresh de pagina - je ziet: **"Your site is live at https://mlcoq.github.io/assistdms-addin/"**

---

## 🔧 STAP 4: API Hosten (BELANGRIJK!)

Je add-in draait nu online, maar de **API** moet ook bereikbaar zijn!

### Voor nu/testen: Ngrok (Tijdelijk)

```powershell
# Download ngrok: https://ngrok.com/download
ngrok http 5001
```

Dan krijg je een URL zoals: `https://abc123.ngrok.io`

**Update dan `config.js` op GitHub:**
```javascript
baseUrl: 'https://abc123.ngrok.io'
```

### Voor productie: Azure/Railway/Fly.io

Kies een hosting platform:
- **Azure App Service** - https://portal.azure.com (gratis tier €0)
- **Railway.app** - https://railway.app (gratis $5/maand)
- **Fly.io** - https://fly.io (gratis tier)

Wil je dat ik hier bij help?

---

## 📧 STAP 5: Installeren in Outlook

### Via Exchange Admin Center (Centralized Deployment)

1. Ga naar **https://admin.microsoft.com**
2. **Settings** → **Integrated apps**
3. **Upload custom apps** → **Upload manifest file**
4. Upload: `github-pages-deploy/manifest.xml`
5. Selecteer wie de add-in krijgt (jezelf, groep, hele organisatie)
6. Deploy

### Of handmatig per gebruiker:

1. Open **Outlook** (nieuwe of klassieke)
2. Ga naar **Get Add-ins**
3. **My add-ins** → **Add from URL**
4. URL: **https://mlcoq.github.io/assistdms-addin/manifest.xml**
5. Add

---

## ✨ Testen

1. **Herstart Outlook**
2. **Open een email**
3. Zoek naar **AssistDMS** knoppen in de ribbon
4. Of klik **... (meer)** → **AssistDMS**

---

## 🐛 Troubleshooting

**Add-in laadt niet:**
- Check of GitHub Pages actief is
- Check of API bereikbaar is (config.js)
- Check browser console (F12) voor errors

**API niet bereikbaar:**
- Is ngrok/je API actief?
- HTTPS verplicht (geen HTTP)
- CORS ingesteld in API?

---

## 📝 Volgende Stappen

- [ ] Bestanden uploaden naar GitHub
- [ ] GitHub Pages activeren
- [ ] API hosten (ngrok of Azure)
- [ ] config.js updaten met API URL
- [ ] Manifest uploaden naar Exchange
- [ ] Testen in Outlook!

**Succes! 🎉**
