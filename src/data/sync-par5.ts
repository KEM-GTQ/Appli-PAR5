/**
 * File d'envoi vers l'API.
 *
 * Le téléphone ne parle qu'à cette API. Le secret Microsoft reste sur le
 * serveur. La base et SharePoint partent ensemble, et l'échec de l'un
 * laisse l'autre aboutir. Tant que les deux ne sont pas confirmés, la fiche
 * et son PDF restent ici. Le même id est renvoyé : un second essai remplace
 * le fichier et la ligne au lieu d'en créer d'autres.
 *
 * Chromium peut finir l'envoi après la fermeture (Background Sync). iOS non :
 * l'application réessaie à la prochaine ouverture, toutes les 30 secondes, et
 * quand le navigateur annonce le retour du réseau.
 */
import type { Par5Enregistre } from './par5-store'
import {
  canalOuvert,
  fichesADeposer,
  lireFiche,
  patchFiche,
  statutDesEtats,
  type FicheStockee
} from './par5-db'

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

const BALISE_REAUTH = 'par5-reauth'

type FournisseurJeton = () => Promise<{ token: string; auteur: string } | { pause: string; reauth: boolean }>

let fournisseurJeton: FournisseurJeton | null = null

/** Branché par la page, jamais par l'agent de service : lui n'a pas de fenêtre pour renouveler la session. */
export function brancherJeton(fn: FournisseurJeton): void {
  fournisseurJeton = fn
}

function authConfiguree(): boolean {
  return Boolean(import.meta.env.VITE_ENTRA_CLIENT_ID && import.meta.env.VITE_ENTRA_TENANT_ID)
}

async function entetes(): Promise<{ entete: HeadersInit; auteur: string } | { pause: string; reauth: boolean }> {
  const entete: Record<string, string> = { 'Content-Type': 'application/json' }
  const cle = import.meta.env.VITE_API_KEY
  if (cle) entete['X-API-Key'] = cle
  if (!authConfiguree() || location.protocol === 'file:') return { entete, auteur: '' }
  if (!fournisseurJeton) {
    return { pause: 'Session Microsoft requise pour envoyer la fiche.', reauth: true }
  }
  const jeton = await fournisseurJeton()
  if ('pause' in jeton) return jeton
  if (jeton.token) entete.Authorization = `Bearer ${jeton.token}`
  return { entete, auteur: jeton.auteur }
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

type CanalReponse = {
  tente?: boolean
  ok?: boolean
  error?: string
  retryable?: boolean
  url?: string | null
}

function resumeEnvoi(fiche: FicheStockee, patch: Partial<FicheStockee>): Partial<FicheStockee> {
  const suivant: FicheStockee = { ...fiche, ...patch, id: fiche.id }
  const statut = statutDesEtats(suivant)
  const erreur = [suivant.baseErreur, suivant.sharePointErreur].filter((texte) => texte).join(' ') || null
  const reessayable =
    canalOuvert(suivant.etatBase, suivant.baseReessayable) ||
    canalOuvert(suivant.etatSharePoint, suivant.sharePointReessayable)
  return {
    ...patch,
    statut,
    erreur,
    reessayable,
    deposeISO: statut === 'depose' ? suivant.deposeISO || new Date().toISOString() : suivant.deposeISO
  }
}

async function noter(fiche: FicheStockee, patch: Partial<FicheStockee>): Promise<FicheStockee | null> {
  const suivante = await patchFiche(fiche.id, resumeEnvoi(fiche, patch))
  await prevenir()
  return suivante
}

async function envoyerUne(fiche: FicheStockee): Promise<boolean> {
  let veutBase = fiche.etatBase !== 'ok'
  let veutSharePoint = fiche.etatSharePoint !== 'ok'
  if (!veutBase && !veutSharePoint) return true

  if (veutSharePoint && !fiche.pdf) {
    const sansPdf = await noter(fiche, {
      etatSharePoint: 'erreur',
      sharePointErreur: 'Le PDF manque sur cet appareil.',
      sharePointReessayable: false
    })
    veutSharePoint = false
    if (!veutBase || !sansPdf) return false
    fiche = sansPdf
  }

  const courant = await noter(fiche, {
    ...(veutBase ? { etatBase: 'envoi' as const, baseErreur: null } : {}),
    ...(veutSharePoint ? { etatSharePoint: 'envoi' as const, sharePointErreur: null } : {})
  })
  if (courant) fiche = courant

  let entete: HeadersInit
  let auteur = ''
  try {
    const obtenu = await entetes()
    if ('pause' in obtenu) {
      await noter(fiche, {
        ...(veutBase ? { etatBase: 'attente' as const, baseErreur: obtenu.pause, baseReessayable: true } : {}),
        ...(veutSharePoint
          ? { etatSharePoint: 'attente' as const, sharePointErreur: obtenu.pause, sharePointReessayable: true }
          : {})
      })
      if (obtenu.reauth && typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent(BALISE_REAUTH))
      }
      return false
    }
    entete = obtenu.entete
    auteur = obtenu.auteur
  } catch {
    const texte = 'En attente d’une connexion.'
    await noter(fiche, {
      ...(veutBase ? { etatBase: 'attente' as const, baseErreur: texte, baseReessayable: true } : {}),
      ...(veutSharePoint
        ? { etatSharePoint: 'attente' as const, sharePointErreur: texte, sharePointReessayable: true }
        : {})
    })
    return false
  }

  const corpsEnvoi: Record<string, unknown> = {
    id: fiche.id,
    noProjet: fiche.donnees.noProjet,
    installation: fiche.donnees.installation,
    emplacement: fiche.donnees.emplacement,
    tache: fiche.donnees.tache,
    dateISO: fiche.dateISO,
    majISO: fiche.majISO,
    lignes: fiche.donnees.lignes.map((l) => ({
      risqueId: l.risqueId,
      present: l.present,
      specification: l.specification,
      moyens: l.moyens,
      responsable: l.responsable,
      realise: l.realise
    })),
    travailleurs: fiche.donnees.travailleurs,
    tenter: { base: veutBase, sharePoint: veutSharePoint }
  }
  if (fiche.sharePointUrl) corpsEnvoi.sharePointUrl = fiche.sharePointUrl
  if (veutSharePoint && fiche.pdf) corpsEnvoi.pdfBase64 = enBase64(fiche.pdf)

  try {
    const reponse = await fetch(adresse('/api/par5'), {
      method: 'POST',
      headers: entete,
      cache: 'no-store',
      body: JSON.stringify(corpsEnvoi)
    })
    let corps: {
      id?: string
      error?: string
      retryable?: boolean
      code?: string
      sharePointUrl?: string | null
      syncedAt?: string
      database?: CanalReponse
      sharePoint?: CanalReponse
    } = {}
    try {
      corps = await reponse.json()
    } catch {
      corps = {}
    }
    if (corps.code === 'auth' || reponse.status === 401) {
      const texte =
        typeof corps.error === 'string' ? corps.error : 'Session expirée. Reconnectez-vous pour envoyer la fiche.'
      await noter(fiche, {
        ...(veutBase ? { etatBase: 'attente' as const, baseErreur: texte, baseReessayable: true } : {}),
        ...(veutSharePoint
          ? { etatSharePoint: 'attente' as const, sharePointErreur: texte, sharePointReessayable: true }
          : {})
      })
      if (typeof window !== 'undefined') window.dispatchEvent(new CustomEvent(BALISE_REAUTH))
      return false
    }

    const repli = typeof corps.error === 'string' ? corps.error : `Envoi échoué (${reponse.status}).`
    const repliReessayable = typeof corps.retryable === 'boolean' ? corps.retryable : reponse.status >= 500
    const patch: Partial<FicheStockee> = {}
    if (veutBase) {
      if (corps.database?.ok === true && corps.id === fiche.id) {
        patch.etatBase = 'ok'
        patch.baseErreur = null
        patch.baseReessayable = true
      } else {
        patch.etatBase = 'erreur'
        patch.baseErreur = corps.database?.error || repli
        patch.baseReessayable = corps.database ? corps.database.retryable !== false : repliReessayable
      }
    }
    if (veutSharePoint) {
      if (corps.sharePoint?.ok === true && corps.id === fiche.id) {
        patch.etatSharePoint = 'ok'
        patch.sharePointErreur = null
        patch.sharePointReessayable = true
        patch.sharePointUrl = corps.sharePoint.url || corps.sharePointUrl || fiche.sharePointUrl
      } else {
        patch.etatSharePoint = 'erreur'
        patch.sharePointErreur = corps.sharePoint?.error || repli
        patch.sharePointReessayable = corps.sharePoint ? corps.sharePoint.retryable !== false : repliReessayable
      }
    }
    if (auteur && (patch.etatBase === 'ok' || patch.etatSharePoint === 'ok')) patch.auteur = auteur
    const suivante = await noter(fiche, patch)
    return suivante?.etatBase === 'ok' && suivante.etatSharePoint === 'ok'
  } catch {
    const texte = 'En attente d’une connexion.'
    await noter(fiche, {
      ...(veutBase ? { etatBase: 'attente' as const, baseErreur: texte, baseReessayable: true } : {}),
      ...(veutSharePoint
        ? { etatSharePoint: 'attente' as const, sharePointErreur: texte, sharePointReessayable: true }
        : {})
    })
    return false
  }
}

async function vider(forcer = false): Promise<{ remaining: number; offline: boolean }> {
  /* L'agent de service ne peut pas renouveler une session Microsoft. Il laisse la file à la page. */
  if (authConfiguree() && typeof window === 'undefined') {
    const enAttente = await fichesADeposer()
    return { remaining: enAttente.length, offline: false }
  }
  if (!forcer && typeof navigator !== 'undefined' && navigator.onLine === false) {
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

export async function aDesFichesADeposer(): Promise<boolean> {
  const fiches = await fichesADeposer()
  return fiches.length > 0
}

export async function flushPending(forcer = false): Promise<{ remaining: number; offline?: boolean }> {
  try {
    return await avecVerrou(() => vider(forcer))
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
      etatBase: 'attente',
      etatSharePoint: 'attente',
      baseErreur: null,
      sharePointErreur: null,
      baseReessayable: true,
      sharePointReessayable: true,
      reessayable: true,
      erreur: null,
      sharePointUrl: null,
      deposeISO: null
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
  if (etat?.etatBase === 'ok' && etat.etatSharePoint === 'ok') {
    return { texte: 'Fiche enregistrée dans la base et PDF déposé dans SharePoint.', alerte: false }
  }
  const morceaux: string[] = []
  if (etat?.etatBase === 'ok') morceaux.push('La base a enregistré la fiche.')
  else if (etat?.baseErreur) morceaux.push(etat.baseErreur)
  if (etat?.etatSharePoint === 'ok') morceaux.push('Le PDF est dans SharePoint.')
  else if (etat?.sharePointErreur) morceaux.push(etat.sharePointErreur)
  if (bilan.offline) {
    morceaux.push('Hors réseau. La fiche et le PDF partiront au retour de la connexion.')
  }
  if (morceaux.length) return { texte: morceaux.join(' '), alerte: true }
  return {
    texte: 'Fiche conservée sur ce téléphone. L’envoi reprendra dès que le serveur répond.',
    alerte: false
  }
}

export async function reessayerPar5(id: string): Promise<void> {
  await avecVerrou(async () => {
    const fiche = await lireFiche(id)
    if (!fiche) return
    await patchFiche(
      id,
      resumeEnvoi(fiche, {
        ...(fiche.etatBase !== 'ok' ? { etatBase: 'attente' as const, baseErreur: null, baseReessayable: true } : {}),
        ...(fiche.etatSharePoint !== 'ok'
          ? { etatSharePoint: 'attente' as const, sharePointErreur: null, sharePointReessayable: true }
          : {})
      })
    )
  })
  await prevenir()
  await flushPending()
}
