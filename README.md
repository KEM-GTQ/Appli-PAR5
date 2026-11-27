# PAR 5 — guide simple

Application pour téléphone. Elle sert à remplir une analyse de risques PAR 5
sur le chantier, à la garder, puis à envoyer le PDF.

Le détail du code est dans [ARCHITECTURE.md](ARCHITECTURE.md).

## Sur le chantier

On peut remplir et enregistrer **une ou plusieurs fiches sans Internet et sans
se connecter**.

1. Remplir la fiche. **No de projet** propose les projets actifs de SharePoint.
   Si la liste ne s’affiche pas, on peut taper le nom à la main.
2. **Enregistrer** garde la fiche et son PDF sur le téléphone.
3. Quand le réseau revient, l’application demande le **compte Microsoft du
   travail**. Les fiches déjà enregistrées ne sont pas perdues.
4. Une fois connecté, chaque PDF part dans SharePoint et la fiche est écrite
   dans la base, au nom de ce compte. On les retrouve dans **Mes fiches**.
5. Si la connexion ou l’envoi échoue, les fiches **restent sur le téléphone**
   jusqu’à ce que ça marche.

**Enregistrer le PDF** ouvre aussi l’enregistrement ou le partage du fichier
sur l’appareil. **Exemple** remplit une fiche de démonstration.

Le téléphone n’écrit pas lui-même dans SharePoint. C’est le serveur qui le fait,
après la connexion Microsoft.

## Lancer l’application

1. Copier `.env.example` vers `.env` et y mettre les accès (base, SharePoint,
   compte Microsoft). Ne pas partager ce fichier.
2. Dans un terminal : `npm i`
3. Dans un terminal : `npm run dev`
4. Dans un autre : `npm run api`

Ouvrir http://localhost:5173/Appli-PAR5/

| Commande | À quoi ça sert |
| --- | --- |
| `npm run dev` | L’application, pour développer |
| `npm run api` | L’envoi vers SharePoint et la base |
| `npm run test:api` | Vérifie l’envoi sans toucher au vrai SharePoint |
| `npm run build` | Prépare la version à distribuer |
| `npm run publier` | Pose un fichier HTML unique sur le Bureau |

Le fichier du Bureau marche hors ligne, mais **il n’envoie rien** à SharePoint.
Seule la version hébergée (`dist-web/`, en HTTPS) envoie les PDF et peut
s’installer sur le téléphone.

## Où vont les informations

- **PDF** : dossier SharePoint `Cursor`, sur le site PowerPlatform-Dev.
  Renvoyer la même fiche remplace le PDF, elle n’en crée pas un deuxième.
- **Fiche** (sans les photos) : base MySQL `par5`.
- **Liste des projets** : liste SharePoint **Projets 2** (titre affiché
  « Liste des projets »), sur `genitique.sharepoint.com`. Le menu montre la
  colonne **Code projet**, seulement si **État** (`_x00c9_tat`) vaut **Actif**.

Pour que cette liste s’affiche, un administrateur SharePoint doit autoriser
l’application à **lire** le site `https://genitique.sharepoint.com`
(permission *Sites.Selected*, rôle *read*, application
`83c11de8-30e6-47fb-b241-eb3582f8e930`). Tant que ce n’est pas fait, on tape
le projet à la main.

La connexion Microsoft du téléphone doit aussi avoir cette adresse de retour
déclarée : `http://localhost:5173/Appli-PAR5/`

## Les 16 risques

La liste des risques est dans `src/data/par5.ts`. C’est une copie du carnet
Génitique. Si le carnet officiel change, il faut mettre à jour cette copie.
