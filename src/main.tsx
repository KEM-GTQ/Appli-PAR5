import React from 'react'
import ReactDOM from 'react-dom/client'
import { registerSW } from 'virtual:pwa-register'
import './main.css'
import { App } from './App'
import { jetonApi, preparerAuth } from './data/auth-par5'
import { brancherJeton } from './data/sync-par5'
import { appliquerTheme, lireTheme } from './data/theme'

/* L'agent de service n'existe que dans la version hébergée. Le fichier unique
   du Bureau n'en a pas, et une page en file:// ne peut pas en enregistrer. */
if (__PWA__ && 'serviceWorker' in navigator && location.protocol.startsWith('http')) {
  registerSW({ immediate: true })
}

/* Le thème est posé AVANT le premier rendu : sinon l'app clignote en sombre
   une fraction de seconde avant de passer en clair. */
appliquerTheme(lireTheme())
brancherJeton(jetonApi)

void preparerAuth().finally(() => {
  ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  )
})
