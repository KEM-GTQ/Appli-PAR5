import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo
} from '@azure/msal-browser'

/**
 * Connexion du travailleur avec son compte Microsoft.
 *
 * Le cache MSAL est dans localStorage : un téléphone déjà connecté rouvre
 * l'application sans réseau. Le jeton sert uniquement à appeler notre API.
 * Il n'autorise pas le téléphone à écrire dans SharePoint.
 */

const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID || ''
const tenantId = import.meta.env.VITE_ENTRA_TENANT_ID || ''
const porteeApi = import.meta.env.VITE_ENTRA_API_SCOPE || ''

export type Jeton = { token: string; auteur: string } | { pause: string; reauth: boolean }

let application: PublicClientApplication | null = null
let preparation: Promise<string | null> | null = null
let redirection = false

function portees(): string[] {
  return porteeApi ? [porteeApi] : ['User.Read']
}

export function authRequise(): boolean {
  return Boolean(clientId && tenantId) && location.protocol !== 'file:'
}

function uriRetour(): string {
  return new URL(import.meta.env.BASE_URL || '/', window.location.origin).href
}

function compteDe(info: AccountInfo | null): { nom: string; info: AccountInfo } | null {
  if (!info) return null
  const nom = info.name || info.username || 'Compte Microsoft'
  return { nom, info }
}

export function compteActif(): { nom: string } | null {
  if (!application) return null
  const info = application.getActiveAccount() || application.getAllAccounts()[0] || null
  const compte = compteDe(info)
  return compte ? { nom: compte.nom } : null
}

export function preparerAuth(): Promise<string | null> {
  if (!authRequise()) return Promise.resolve(null)
  if (!preparation) {
    application = new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: uriRetour()
      },
      cache: { cacheLocation: 'localStorage' }
    })
    const msal = application
    preparation = msal
      .initialize()
      .then(() => msal.handleRedirectPromise())
      .then((resultat) => {
        const info = resultat?.account || msal.getAllAccounts()[0] || null
        if (info) msal.setActiveAccount(info)
        return null
      })
      .catch(() => 'La connexion Microsoft a échoué. Réessayez lorsque le réseau est là.')
  }
  return preparation
}

export async function connecter(): Promise<void> {
  if (!authRequise() || redirection || !navigator.onLine) return
  redirection = true
  try {
    await preparerAuth()
    await application!.loginRedirect({ scopes: portees() })
  } catch {
    redirection = false
  }
}

export async function jetonApi(): Promise<Jeton> {
  if (!authRequise()) return { token: '', auteur: '' }
  await preparerAuth()
  const info = application!.getActiveAccount() || application!.getAllAccounts()[0] || null
  if (!info) {
    return {
      pause: navigator.onLine
        ? 'Connexion Microsoft requise pour envoyer la fiche. Elle reste sur ce téléphone.'
        : 'Enregistrée sur ce téléphone. La connexion Microsoft sera demandée au retour du réseau.',
      reauth: true
    }
  }
  application!.setActiveAccount(info)
  try {
    const resultat = await application!.acquireTokenSilent({ account: info, scopes: portees() })
    /* Sans scope d'API dédié, le jeton d'identité a pour audience l'application elle-même. */
    const token = porteeApi ? resultat.accessToken : resultat.idToken
    if (!token) return { pause: 'Session Microsoft incomplète. Reconnectez-vous.', reauth: true }
    return { token, auteur: compteActif()?.nom || info.name || info.username || '' }
  } catch (erreur) {
    if (erreur instanceof InteractionRequiredAuthError) {
      if (!navigator.onLine) {
        return {
          pause: 'Session expirée. Reconnectez-vous dès que le réseau revient.',
          reauth: true
        }
      }
      if (!redirection) {
        redirection = true
        await application!.acquireTokenRedirect({ account: info, scopes: portees() })
      }
      return { pause: 'Session expirée. Reconnectez-vous pour envoyer la fiche.', reauth: true }
    }
    if (!navigator.onLine) {
      return { pause: 'Hors réseau. La fiche partira au retour de la connexion.', reauth: false }
    }
    return { pause: 'Session Microsoft indisponible. La fiche reste sur ce téléphone.', reauth: true }
  }
}
