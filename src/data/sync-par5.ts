/**
 * File d'envoi vers l'API.
 *
 * Le téléphone ne parle qu'à cette API. Le secret Microsoft reste sur le
 * serveur, qui dépose le PDF dans SharePoint. Tant que le serveur ne répond
 * pas, la fiche et son PDF restent ici ; le même id est renvoyé, donc un
 * second essai écrase le fichier au lieu de le dupliquer.
 *
 * Chromium peut finir l'envoi après la fermeture (Background Sync). iOS non :
 * l'application réessaie à la prochaine ouverture, toutes les 30 secondes, et
 * quand le navigateur annonce le retour du réseau.
 */
import type { Par5Enregistre } from './par5-store'
import { fichesADeposer, lireFiche, patchFiche } from './par5-db'

const VERROU = 'gtq-par5-envoi'
const BALISE = 'par5-sync'

let chaine = Promise.resolve()

function avecVerrou<T>(tache: () => Promise<T>): Promise<T> {
  if (typeof navigator !== 'undefined' && navigator.locks?.request) {
    /* Le verrou tient jusqu'à la fin de la promesse et rend sa valeur.
       Le type du DOM emballe cette promesse une fois de trop. */
    return navigator.locks.request(VERROU, () => tache()).then((valeur) => valeur as T)
  }
  const cours = chaine.then(() => tache(), () => tache())
  chaine = cours.then(
    () => undefined,
    () => undefined
  )
  return cours
}

function entetes(): HeadersInit {
  const entete: Record<string, string> = { 'Content-Type': 'application/json' }
  const cle = import.meta.env.VITE_API_KEY
  if (cle) entete['X-API-Key'] = cle
  return entete
}

function adresse(chemin: string): string {
  const base = (import.meta.env.VITE_API_URL || '').replace(/\/$/, '')
  return `${base}${chemin}`
}

function enBase64(data: ArrayBuffer): string {
  const octets = new Uint8Array(data)
  let binaire = ''
  const pas = 0x8000
  for (let i = 0; i < octets.length; i += pas) {
    const morceau = octets.subarray(i, i + pas)
    binaire += String.fromCharCode(...(morceau as unknown as number[]))
  }
  return btoa(binaire)
}

export async function prevenir(): Promise<void> {
  if (typeof window !== 'undefined') {
    window.dispatchEvent(new CustomEvent(BALISE))
    return
  }
  const portée = self as unknown as {
    clients?: {
      matchAll: (filtre: { includeUncontrolled: boolean; type: 'window' }) => Promise<Array<{ postMessage: (message: unknown) => void }>>
    }
  }
  if (!portée.clients?.matchAll) return
  const ouverts = await portée.clients.matchAll({ includeUncontrolled: true, type: 'window' })
  for (const client of ouverts) client.postMessage({ type: BALISE })
}

export async function inscrireSyncArrierePlan(): Promise<void> {
  if (typeof navigator === 'undefined' || !navigator.serviceWorker) return
  try {
    const inscription = await navigator.serviceWorker.ready
    if ('sync' in inscription) {
      await (inscription as ServiceWorkerRegistration & { sync: { register: (t: string) => Promise<void> } }).sync.register(BALISE)
    }
  } catch {
    /* Firefox et iOS n'ont pas Background Sync. L'app ouverte réessaie seule. */
  }
}

async function envoyerUne(fiche: Awaited<ReturnType<typeof fichesADeposer>>[number]): Promise<boolean> {
  if (!fiche.pdf) {
    await patchFiche(fiche.id, {
      statut: 'erreur',
      reessayable: false,
      erreur: 'Le PDF manque sur cet appareil.'
    })
    await prevenir()
    return false
  }

  await patchFiche(fiche.id, { statut: 'envoi', erreur: null })
  await prevenir()

  try {
    const reponse = await fetch(adresse('/api/par5'), {
      method: 'POST',
      headers: entetes(),
      cache: 'no-store',
      body: JSON.stringify({
        id: fiche.id,
        noProjet: fiche.donnees.noProjet,
        installation: fiche.donnees.installation,
        emplacement: fiche.donnees.emplacement,
        tache: fiche.donnees.tache,
        dateISO: fiche.dateISO,
        majISO: fiche.majISO,
        pdfBase64: enBase64(fiche.pdf)
      })
    })
    let corps: { error?: string; retryable?: boolean; sharePointUrl?: string | null; syncedAt?: string } = {}
    try {
      corps = await reponse.json()
    } catch {
      corps = {}
    }
    if (reponse.ok) {
      await patchFiche(fiche.id, {
        statut: 'depose',
        reessayable: false,
        erreur: null,
        sharePointUrl: corps.sharePointUrl || null,
        deposeISO: corps.syncedAt || new Date().toISOString()
      })
      await prevenir()
      return true
    }
    const reessayable = typeof corps.retryable === 'boolean' ? corps.retryable : reponse.status >= 500
    await patchFiche(fiche.id, {
      statut: 'erreur',
      reessayable,
      erreur: typeof corps.error === 'string' ? corps.error : `Dépôt échoué (${reponse.status}).`
    })
    await prevenir()
    return false
  } catch {
    await patchFiche(fiche.id, {
      statut: 'attente',
      reessayable: true,
      erreur: 'En attente d’une connexion.'
    })
    await prevenir()
    return false
  }
}

async function vider(): Promise<{ remaining: number; offline: boolean }> {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    const enAttente = await fichesADeposer()
    return { remaining: enAttente.length, offline: enAttente.length > 0 }
  }

  const fiches = await fichesADeposer()
  let remaining = 0
  for (const fiche of fiches) {
    const deposee = await envoyerUne(fiche)
    if (!deposee) remaining += 1
  }
  return { remaining, offline: false }
}

export async function flushPending(): Promise<{ remaining: number; offline?: boolean }> {
  try {
    return await avecVerrou(() => vider())
  } catch (erreur) {
    console.error(`Sync échouée : ${erreur instanceof Error ? erreur.message : String(erreur)}`)
    return { remaining: 1 }
  }
}

/**
 * Pose le PDF sur la fiche déjà enregistrée et tente le dépôt.
 * Le texte retourné est celui qu'on affiche sous le bouton d'envoi.
 */
export async function deposerPar5(
  fiche: Par5Enregistre,
  pdf: ArrayBuffer
): Promise<{ texte: string; alerte: boolean }> {
  await avecVerrou(() =>
    patchFiche(fiche.id, {
      pdf,
      nomPdf: `PAR5-${fiche.id}.pdf`,
      statut: 'attente',
      reessayable: true,
      erreur: null
    })
  )
  await prevenir()
  void inscrireSyncArrierePlan()

  if (location.protocol === 'file:') {
    return {
      texte: 'Cette copie ouverte depuis un fichier ne dépose pas dans SharePoint.',
      alerte: true
    }
  }

  const bilan = await flushPending()
  const etat = await lireFiche(fiche.id)
  if (etat?.statut === 'depose') return { texte: 'PDF déposé dans SharePoint.', alerte: false }
  if (etat?.statut === 'erreur') {
    return { texte: etat.erreur || 'Dépôt SharePoint impossible.', alerte: true }
  }
  if (bilan.offline) {
    return {
      texte: 'Hors réseau. Le PDF partira dans SharePoint au retour de la connexion.',
      alerte: true
    }
  }
  return {
    texte: 'PDF conservé sur ce téléphone. Il sera déposé dans SharePoint dès que le serveur répond.',
    alerte: false
  }
}

export async function reessayerPar5(id: string): Promise<void> {
  await avecVerrou(() => patchFiche(id, { statut: 'attente', reessayable: true, erreur: null }))
  await prevenir()
  await flushPending()
}
