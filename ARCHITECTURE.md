# Architecture

Deux programmes. Le **téléphone** fait la fiche et le PDF. Le **serveur**
les range dans SharePoint et dans MySQL. Le téléphone ne connaît aucun secret.

```
Téléphone                         Serveur                         Dehors
────────                          ──────                          ──────
Formulaire                        POST /api/par5                  SharePoint (PDF)
   │                              GET  /api/projets               MySQL (fiche)
   ├─ mémoire du téléphone        │
   └─ file d'envoi ──────────────►├─ vérifie le compte Microsoft
                                  ├─ écrit la fiche
                                  ├─ envoie le PDF
                                  └─ relit la liste des projets
```

## Parcours d'une fiche

1. Le travailleur remplit le formulaire. Rien ne part encore.
2. **Enregistrer** écrit la fiche dans la mémoire du téléphone, fabrique le PDF,
   puis la met dans la file d'envoi.
3. S'il n'y a pas de réseau ou pas de compte Microsoft, la fiche reste dans
   la file. On peut en enregistrer plusieurs.
4. Au retour du réseau, l'application demande le compte Microsoft. Les fiches
   déjà enregistrées ne sont pas effacées.
5. Chaque fiche en attente est envoyée au serveur, avec le jeton du compte.
6. Le serveur lance en même temps l'écriture dans MySQL et l'envoi du PDF.
   L'échec de l'un n'annule pas l'autre. Si les deux réussissent, le lien du
   PDF est ajouté à la ligne.
7. **Mes fiches** montre deux pastilles. Gris : en attente. Jaune : en cours.
   Vert : confirmé par la destination. Rouge : échec, avec le message.
   La pastille **Base** du bandeau dit si MySQL répond, à part de l'envoi.
8. La fiche n'est entièrement synchronisée que lorsque les deux pastilles
   sont vertes. Celle qui a échoué reste dans la file et sera réessayée seule.
   Le même identifiant est renvoyé : SharePoint remplace le PDF, MySQL met
   à jour la même ligne.

Les photos restent dans le PDF. Elles ne sont pas copiées dans MySQL.

## Le téléphone (`src/`)

| Fichier | Rôle |
| --- | --- |
| `main.tsx` | Démarre l'application, le thème, et prépare la connexion Microsoft avant d'afficher l'écran. |
| `App.tsx` | L'écran : onglets, bandeau réseau, bouton de connexion. N'empêche jamais de remplir une fiche. |
| `pages/par5/Par5Form.tsx` | Le formulaire : chantier, 16 risques, travailleurs, enregistrement. |
| `pages/par5/MesFiches.tsx` | L'historique de ce téléphone, avec l'état de chaque envoi. |
| `pages/par5/Definitions.tsx` | Le texte d'aide des 16 risques. |
| `components/ChoixProjet.tsx` | Le menu cherchable du numéro de projet. On peut aussi taper à la main. |
| `data/par5.ts` | La forme d'une fiche et la liste des 16 risques. |
| `data/par5-db.ts` | La mémoire du téléphone (IndexedDB) : fiches, PDF, état d'envoi. |
| `data/par5-store.ts` | Enregistrer, lister, reprendre une fiche. |
| `data/pdf-par5.ts` | Fabrique le PDF sur le téléphone. |
| `data/photo.ts` | Réduit les photos pour qu'elles tiennent dans le téléphone. |
| `data/sync-par5.ts` | La file d'envoi. Réessaie toute seule, et n'envoie rien sans jeton quand la connexion est exigée. |
| `data/auth-par5.ts` | La connexion Microsoft du travailleur. Ne sert qu'à appeler notre serveur. |
| `data/projets.ts` | Va chercher la liste des projets et en garde une copie sur le téléphone. |
| `sw.js` | Dans la version hébergée : garde l'application hors ligne et relance l'envoi. iOS ne relance pas tout seul. |
| `data/theme.ts` | Clair ou sombre. |

Chaque fiche a deux états d'envoi, la base et SharePoint : en attente,
en cours, confirmé, ou échec. La fiche est déposée quand les deux sont
confirmés, et partielle quand un seul l'est. Seul le canal encore ouvert
est renvoyé. Un brouillon ne part pas.

## Le serveur (`server/src/`)

| Fichier | Rôle |
| --- | --- |
| `index.js` | Lit `.env`, ouvre MySQL, crée la table si besoin, démarre l'API. |
| `app.js` | Les adresses HTTP. `GET /api/projets` est ouvert : il sert à remplir la fiche. `POST /api/par5` exige le compte Microsoft, puis écrit la base et envoie le PDF l'un sans dépendre de l'autre. `GET /api/health` dit si MySQL répond. |
| `validate.js` | Refuse une fiche mal formée. Le PDF doit vraiment être un PDF, et faire 4 Mo au plus. |
| `entra.js` | Vérifie le jeton Microsoft. Retient le courriel du compte comme auteur. |
| `mysql.js` | Enregistre ou met à jour la fiche. N'efface pas un lien SharePoint déjà connu si le nouvel envoi n'en a pas encore. |
| `sharepoint.js` | Envoie le PDF avec le compte de l'application, pas avec celui du travailleur. |
| `projets.js` | Lit la liste **Projets 2**. Le menu reçoit la colonne **Code projet**, seulement si `_x00c9_tat` vaut **Actif**. |
| `errors.js` | Distingue une erreur à réessayer d'une erreur définitive. |

Deux comptes Microsoft, volontairement séparés :

- **Le travailleur** prouve qui a rempli la fiche. Son jeton n'ouvre pas SharePoint.
- **L'application** (secret seulement dans `.env`, sur le serveur) dépose le PDF
  et lit la liste des projets.

## Ce qui reste sur place

| Endroit | Contenu |
| --- | --- |
| Téléphone | Fiches, PDF, photos, copie de la liste des projets, session Microsoft. |
| SharePoint, dossier `Cursor` | Les PDF, un fichier par fiche. |
| MySQL, table `par5_fiches` | L'en-tête, les réponses, les travailleurs, l'auteur, le lien du PDF. |
| SharePoint, liste **Projets 2** | La source du menu. L'application la lit, elle ne la modifie pas. |
