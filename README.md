# Libro Mastro — portale personale per il monitoraggio delle finanze

Portale web per caricare mensilmente estratti conto CSV/Excel, categorizzare
entrate e uscite con regole automatiche e monitorare budget e scostamenti.

I dati (movimenti, categorie, regole, budget) restano **solo nel browser**
in cui apri il sito (`localStorage`) — non vengono mai inviati altrove né
salvati nel repository.

## Come pubblicarlo su GitHub Pages

1. **Crea un repository** su [github.com/new](https://github.com/new)
   (va bene pubblico: il codice non contiene tuoi dati personali).

2. **Carica questo progetto** nel repository. Dal terminale, nella cartella
   di questo progetto:

   ```bash
   git init
   git add .
   git commit -m "Prima versione del portale"
   git branch -M main
   git remote add origin https://github.com/TUO-USERNAME/NOME-REPO.git
   git push -u origin main
   ```

   In alternativa, se non usi il terminale, puoi trascinare tutti i file
   e le cartelle di questo progetto direttamente nella pagina web del
   repository appena creato ("Add file" → "Upload files").

3. **Attiva GitHub Pages**: nel repository vai su *Settings → Pages* e,
   sotto "Build and deployment", imposta la sorgente su **GitHub Actions**.
   Il workflow incluso in `.github/workflows/deploy.yml` compila e
   pubblica automaticamente il sito a ogni `push` sul branch `main`.

4. Dopo un paio di minuti il sito sarà online su:
   `https://TUO-USERNAME.github.io/NOME-REPO/`

   Trovi il link esatto anche in *Settings → Pages* una volta completata
   la prima pubblicazione.

## Aggiornare il sito in futuro

Ogni volta che vuoi modificare il portale (nuove categorie, nuove
funzionalità, correzioni), basta aggiornare i file e fare di nuovo:

```bash
git add .
git commit -m "Descrizione della modifica"
git push
```

Il sito si ripubblica automaticamente in 1-2 minuti.

## Sviluppo in locale (facoltativo)

Per provare il sito sul tuo computer prima di pubblicarlo:

```bash
npm install
npm run dev
```
