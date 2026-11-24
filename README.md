# Génitique PAR 5 — application terrain

Application **téléphone** permettant aux travailleurs de remplir sur place le
**PAR 5** (analyse de risques, 16 risques du carnet Génitique de janvier 2025),
de le signer et de le transmettre en PDF.

La fiche et son PDF sont produits sur le téléphone et gardés dans IndexedDB.
Le travailleur peut remplir et enregistrer un ou plusieurs PAR 5 sans réseau et
sans session Microsoft. La connexion (Entra ID, code d'autorisation et PKCE) n'est
demandée que pour l'envoi : au retour du réseau, l'application propose le compte
professionnel, conserve les fiches déjà enregistrées, puis l'API dépose chaque PDF
dans SharePoint et écrit la fiche dans MySQL au nom de ce compte. Ces fiches
apparaissent dans Mes fiches. Si la connexion ou l'envoi échoue, elles restent sur
le téléphone jusqu'à la réussite. Le téléphone n'a pas de secret Microsoft et ne
peut pas écrire dans SharePoint lui-même. Un nouvel envoi de la même fiche écrase
le fichier et met à jour la même ligne.

Le fichier HTML unique posé sur le Bureau ne peut pas déposer dans SharePoint.
C'est la version hébergée (`dist-web/`) qui le fait.

## Pile technique

Vite · React 18 · TypeScript · Tailwind v4 · jsPDF. Pas d'Electron : la cible
est un téléphone.

## Commandes

| Commande | Effet |
| --- | --- |
| `npm i` | Installer les dépendances requises |
| `npm run dev` | Serveur de développement (proxy `/api` vers le port 3001) |
| `npm run api` | API de dépôt SharePoint |
| `npm run test:api` | Contrat de l'API, sans réseau |
| `npm run build` | Contrôle des types puis construction dans `dist/` |
| `npm run publier` | Construit puis copie **`Génitique PAR 5.html`** sur le Bureau |
| `npm run apercu` | Sert `dist/` sur le port 4190 |
| `npm run test:pdf` | Banc d'essai du PDF (relit le PDF produit avec pdf.js) |
| `node tools/construire-web.cjs` | Produit `dist-web/`, la version **installable**, avec l'agent de service |
| `node tools/audit-pdf.cjs` | Audit géométrique du PDF (débordements, chevauchements) |
| `node tools/test-controle.cjs` | Banc d'essai du garde-fou avant transmission |
| `node tools/contraste.cjs` | Vérifie les contrastes de la palette |
| `node tools/faire-icone.cjs` | Regénère les icônes 180/192/512 px |

Copier `.env.example` vers `.env` et y mettre l'application Entra avant le premier dépôt. En développement, `npm run dev` et `npm run api` tournent ensemble : le premier proxifie `/api` vers le second.

## Deux formats de sortie

- **`dist/index.html`** — fichier **unique**, tout inliné (y compris les
  icônes). Se double-clique sur un poste, marche hors ligne, se copie sur une
  clé. C'est ce que `npm run publier` dépose sur le Bureau.
- **`dist-web/`** — dossier à déposer sur un hébergement **HTTPS**. Il porte le
  manifeste, l'agent de service et les icônes en vrais fichiers : c'est le seul
  format qu'Android propose d'« installer » et le seul qui fonctionne hors
  ligne une fois servi par le réseau.

Les deux dossiers sont **régénérés** à chaque construction et ne sont donc pas
versionnés.

## Structure

```
src/
  App.tsx                    coque + barre d'onglets (constante ONGLETS)
  sw.js                      cache hors ligne et reprise du dépôt
  components/                logo, écran d'ouverture, bannière d'installation
  data/
    par5.ts                  modèle des 16 risques + verifierPar5()
    par5-store.ts            fiches (IndexedDB) et reprise de l'ancien localStorage
    par5-db.ts               base locale
    sync-par5.ts             file d'envoi vers l'API
    pdf-par5.ts              construction et envoi du PDF
    photo.ts                 compression et budget mémoire des photos
    theme.ts                 bascule clair/sombre
  pages/par5/                formulaire, mes fiches, définitions
server/                      API : PDF vers SharePoint, fiche écrite dans MySQL
tools/                       construction, publication et bancs d'essai
```

> ⚠️ `src/data/par5.ts` est une **copie** du carnet des 16 risques présent dans
> ToolApp. Si le carnet officiel change, les deux projets doivent être mis à
> jour.
