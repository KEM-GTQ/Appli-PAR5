/**
 * Enregistrement des PAR 5 (analyses de risques remplies).
 *
 * Les fiches vivent dans IndexedDB. Au premier lancement, ce qui était encore
 * dans localStorage (clé `gtq-par5`) est repris, puis la clé est retirée.
 */
import {
  RISQUES,
  ligneVide,
  type DonneesPar5,
  type LignePar5,
  type Presence,
  type Travailleur
} from './par5'
import {
  lireFiche,
  mettreFiche,
  supprimerFiche,
  toutesLesFiches,
  type FicheStockee,
  type StatutSync
} from './par5-db'

export interface Par5Enregistre {
  id: string
  /**
   * La date de L'ANALYSE — celle du premier enregistrement, et elle ne bouge
   * plus.
   *
   * ⚠️ Défaut corrigé : chaque sauvegarde la remplaçait par la date du jour.
   * Rouvrir une fiche d'il y a trois semaines pour corriger une faute de frappe
   * la redatait donc à aujourd'hui — dans la liste ET dans l'en-tête du PDF.
   * Pour un document de sécurité qu'on ressort des mois plus tard, la date de
   * l'analyse est une donnée, pas un horodatage de fichier.
   */
  dateISO: string
  /** Dernière modification — pour savoir si la fiche a été retouchée. */
  majISO?: string
  donnees: DonneesPar5
  /** Absent sur une fiche pas encore passée par la file d'envoi. */
  statut?: StatutSync
  erreur?: string | null
  sharePointUrl?: string | null
  reessayable?: boolean
  /** Compte Microsoft associé au moment du dépôt. */
  auteur?: string | null
}

const CLE = 'gtq-par5'
export const SCHEMA_PAR5 = 1

const uid = (): string => 'par5-' + Math.random().toString(36).slice(2, 10)

function resume(fiche: FicheStockee): Par5Enregistre {
  return {
    id: fiche.id,
    dateISO: fiche.dateISO,
    majISO: fiche.majISO,
    donnees: fiche.donnees,
    statut: fiche.statut,
    erreur: fiche.erreur,
    sharePointUrl: fiche.sharePointUrl,
    reessayable: fiche.reessayable,
    auteur: fiche.auteur
  }
}

let migration: Promise<void> | null = null

/** Reprend les fiches de l'ancien localStorage, une seule fois par onglet. */
function assurerMigration(): Promise<void> {
  if (!migration) {
    migration = migrerAncienStockage().catch((erreur) => {
      migration = null
      throw erreur
    })
  }
  return migration
}

async function migrerAncienStockage(): Promise<void> {
  let brut: string | null = null
  try {
    brut = localStorage.getItem(CLE)
  } catch {
    return
  }
  if (!brut) return

  let arr: unknown[] = []
  try {
    const json = JSON.parse(brut) as unknown
    if (Array.isArray(json)) arr = json
    else if (json && typeof json === 'object' && Array.isArray((json as { par5?: unknown }).par5)) {
      arr = (json as { par5: unknown[] }).par5
    }
  } catch {
    return
  }

  const deja = new Set((await toutesLesFiches()).map((f) => f.id))
  for (const item of arr) {
    const fiche = normaliserPar5(item)
    if (!fiche || deja.has(fiche.id)) continue
    await mettreFiche({
      id: fiche.id,
      dateISO: fiche.dateISO,
      majISO: fiche.majISO,
      donnees: fiche.donnees,
      pdf: null,
      nomPdf: null,
      statut: 'brouillon',
      reessayable: false,
      erreur: null,
      sharePointUrl: null,
      deposeISO: null
    })
  }
  try {
    localStorage.removeItem(CLE)
  } catch {
    /* La copie IndexedDB est faite. Si la clé reste, le prochain lancement la retrouvera déjà. */
  }
}

/** Ouvre le stockage. Échoue si le navigateur bloque IndexedDB. */
export async function ouvrirStockage(): Promise<void> {
  await assurerMigration()
  await toutesLesFiches()
}
const str = (v: unknown, d = ''): string => (typeof v === 'string' ? v : d)

function normaliserLigne(risqueId: string, x: unknown): LignePar5 {
  const o = (x && typeof x === 'object' ? x : {}) as Record<string, unknown>
  const present: Presence = o.present === 'oui' ? 'oui' : o.present === 'non' ? 'non' : null
  return {
    risqueId,
    present,
    specification: str(o.specification),
    moyens: str(o.moyens),
    responsable: str(o.responsable),
    realise: o.realise === true,
    /* On n'accepte QUE des data URL d'image : une chaine quelconque venue
       d'un stockage abime finirait dans un <img src> et dans le PDF. */
    photo:
      typeof o.photo === 'string' && o.photo.startsWith('data:image/') ? o.photo : undefined
  }
}

function normaliserDonnees(x: unknown): DonneesPar5 {
  const o = (x && typeof x === 'object' ? x : {}) as Record<string, unknown>
  // Index des lignes stockées par risqueId, pour retrouver l'ordre officiel des RISQUES.
  const parId = new Map<string, unknown>()
  if (Array.isArray(o.lignes))
    for (const l of o.lignes) {
      const rid = (l && typeof l === 'object' ? (l as Record<string, unknown>).risqueId : '') as string
      if (typeof rid === 'string') parId.set(rid, l)
    }
  const lignes = RISQUES.map((r) =>
    parId.has(r.id) ? normaliserLigne(r.id, parId.get(r.id)) : ligneVide(r.id)
  )
  const travailleurs: Travailleur[] = Array.isArray(o.travailleurs)
    ? o.travailleurs.map((t) => {
        const to = (t && typeof t === 'object' ? t : {}) as Record<string, unknown>
        return { nom: str(to.nom), initiales: str(to.initiales), date: str(to.date) }
      })
    : [{ nom: '', initiales: '', date: '' }]
  return {
    noProjet: str(o.noProjet),
    installation: str(o.installation),
    emplacement: str(o.emplacement),
    tache: str(o.tache),
    lignes,
    travailleurs: travailleurs.length ? travailleurs : [{ nom: '', initiales: '', date: '' }]
  }
}

export function normaliserPar5(x: unknown): Par5Enregistre | null {
  if (!x || typeof x !== 'object') return null
  const o = x as Record<string, unknown>
  const id = str(o.id)
  if (!id) return null
  const dateISO =
    typeof o.dateISO === 'string' && !isNaN(Date.parse(o.dateISO)) ? o.dateISO : new Date().toISOString()
  const majISO = typeof o.majISO === 'string' && !isNaN(Date.parse(o.majISO)) ? o.majISO : undefined
  return { id, dateISO, majISO, donnees: normaliserDonnees(o.donnees) }
}

export async function listerPar5(): Promise<Par5Enregistre[]> {
  await assurerMigration()
  const fiches = await toutesLesFiches()
  return fiches.map(resume).sort((a, b) => b.dateISO.localeCompare(a.dateISO))
}

export async function enregistrerPar5(
  donnees: DonneesPar5,
  id?: string
): Promise<Par5Enregistre & { persiste: boolean }> {
  await assurerMigration()
  const rid = id ?? uid()
  const existante = await lireFiche(rid)
  const maintenant = new Date().toISOString()

  /* On CONSERVE la date de l'analyse quand la fiche existe déjà ; seule la
     date de modification suit. Un enregistrement local n'envoie rien :
     le PDF précédent ne correspond plus au texte, on le retire de la file. */
  const rec: FicheStockee = {
    id: rid,
    dateISO: existante?.dateISO ?? maintenant,
    majISO: maintenant,
    donnees,
    pdf: null,
    nomPdf: null,
    statut: 'brouillon',
    reessayable: false,
    erreur: null,
    sharePointUrl: existante?.sharePointUrl ?? null,
    deposeISO: existante?.deposeISO ?? null,
    auteur: existante?.auteur ?? null
  }
  try {
    await mettreFiche(rec)
    return { ...resume(rec), persiste: true }
  } catch {
    console.error('[par5-store] Écriture IndexedDB échouée — non persisté.')
    return { ...resume(rec), persiste: false }
  }
}

/** La fiche a-t-elle été retouchée après le jour de l'analyse ? */
export function retouchee(p: Par5Enregistre): boolean {
  if (!p.majISO) return false
  return p.majISO.slice(0, 10) !== p.dateISO.slice(0, 10)
}

/**
 * REPRENDRE LA FICHE DE LA VEILLE.
 *
 * Une equipe fait la meme tache trois jours de suite : les seize reponses sont
 * identiques, seules les dates changent. Aujourd'hui il faut tout recocher —
 * donc on recoche vite, donc on ne relit plus. Le but n'est pas de gagner du
 * temps, c'est de rendre a l'analyse une chance d'etre RELUE plutot que
 * retapee.
 *
 * Ce qui est repris : le chantier, les seize reponses, les moyens, les
 * responsables. Ce qui est REMIS A ZERO :
 *  · les dates de signature — on ne recopie pas la signature d'hier ;
 *  · « moyens mis en place » — c'est un constat du jour, il se refait ;
 *  · les photos — elles montraient le chantier d'hier, et les garder
 *    consommerait le budget memoire pour rien.
 */
export function dupliquerPar5(source: Par5Enregistre, dateJour: string): DonneesPar5 {
  const d = source.donnees
  return {
    noProjet: d.noProjet,
    installation: d.installation,
    emplacement: d.emplacement,
    tache: d.tache,
    lignes: d.lignes.map((l) => ({
      risqueId: l.risqueId,
      present: l.present,
      specification: l.specification,
      moyens: l.moyens,
      responsable: l.responsable,
      realise: false,
      photo: undefined
    })),
    travailleurs: d.travailleurs
      .filter((t) => t.nom.trim() || t.initiales.trim())
      .map((t) => ({ nom: t.nom, initiales: t.initiales, date: dateJour }))
  }
}

/**
 * Poids total des photos deja stockees, en caracteres — pour refuser une
 * piece jointe AVANT qu'elle ne fasse echouer l'enregistrement de la fiche.
 * `exclureId` ecarte la fiche en cours d'edition, dont les photos vont etre
 * reecrites et non ajoutees.
 */
export async function poidsPhotos(exclureId?: string): Promise<number> {
  await assurerMigration()
  let n = 0
  for (const p of await toutesLesFiches()) {
    if (p.id === exclureId) continue
    for (const l of p.donnees.lignes) n += l.photo ? l.photo.length : 0
  }
  return n
}

export async function supprimerPar5(id: string): Promise<boolean> {
  try {
    await supprimerFiche(id)
    return true
  } catch {
    return false
  }
}
