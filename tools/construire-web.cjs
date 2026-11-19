/*
 * FABRIQUE LA VERSION INSTALLABLE (dist-web/).
 *
 * C'est la construction Vite avec l'agent de service : cache de l'application
 * et reprise du dépôt SharePoint quand Chromium retrouve le réseau.
 * Le fichier HTML unique du Bureau (`npm run build`) n'a pas cet agent —
 * une page ouverte depuis le disque ne peut pas en enregistrer.
 *
 *   node tools/construire-web.cjs
 */
'use strict'
const { spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const RACINE = path.resolve(__dirname, '..')
const WEB = path.join(RACINE, 'dist-web')

const resultat = spawnSync('npx', ['vite', 'build'], {
  cwd: RACINE,
  stdio: 'inherit',
  env: { ...process.env, PAR5_PWA: '1' },
  shell: true
})

if (resultat.status !== 0) process.exit(resultat.status || 1)

const ko = (fichier) => Math.round(fs.statSync(path.join(WEB, fichier)).size / 1024)
console.log('\ndist-web/ construit\n')
for (const fichier of fs.readdirSync(WEB).sort()) {
  const chemin = path.join(WEB, fichier)
  if (!fs.statSync(chemin).isFile()) continue
  console.log('  ' + String(ko(fichier)).padStart(5) + ' Ko  ' + fichier)
}
console.log("\nÀ déposer tel quel sur un hébergement HTTPS. L'API écoute à part (npm run api).\n")
