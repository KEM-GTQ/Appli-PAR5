/**
 * Fiches sur l'appareil, dans IndexedDB.
 *
 * localStorage plafonne autour de 5 Mo et stocke en texte : un PDF, même
 * d'une page, n'y tient pas à côté des photos. IndexedDB garde le PDF en
 * binaire, et le même identifiant sert au nouvel envoi — SharePoint écrase
 * le fichier au lieu d'en créer un second.
 */
import { openDB, type DBSchema, type IDBPDatabase } from 'idb'
import type { DonneesPar5 } from './par5'

export type StatutSync = 'brouillon' | 'attente' | 'envoi' | 'partiel' | 'depose' | 'erreur'

/** Gris en attente, jaune pendant l'envoi, vert confirmé, rouge en échec. */
export type EtatEnvoi = 'attente' | 'envoi' | 'ok' | 'erreur'

export interface FicheStockee {
  id: string
  dateISO: string
  majISO?: string
  donnees: DonneesPar5
  pdf: ArrayBuffer | null
  nomPdf: string | null
  statut: StatutSync
  reessayable: boolean
  erreur: string | null
  sharePointUrl: string | null
  deposeISO: string | null
  /** Compte Microsoft qui a déposé la fiche. Vide tant que l'envoi n'a pas abouti. */
  auteur?: string | null
  /** Envoi de la fiche vers MySQL. Absent sur les fiches enregistrées avant cette séparation. */
  etatBase?: EtatEnvoi
  baseErreur?: string | null
  baseReessayable?: boolean
  /** Envoi du PDF vers SharePoint. */
  etatSharePoint?: EtatEnvoi
  sharePointErreur?: string | null
  sharePointReessayable?: boolean
}

interface BasePar5 extends DBSchema {
  fiches: {
    key: string
    value: FicheStockee
    indexes: { 'par-statut': StatutSync }
  }
}

const NOM = 'gtq-par5'
const MAGASIN = 'fiches'

let promesse: Promise<IDBPDatabase<BasePar5>> | null = null

export function ouvrirBase(): Promise<IDBPDatabase<BasePar5>> {
  if (!promesse) {
    promesse = openDB<BasePar5>(NOM, 1, {
      upgrade(db) {
        const magasin = db.createObjectStore(MAGASIN, { keyPath: 'id' })
        magasin.createIndex('par-statut', 'statut')
      }
    }).catch((erreur) => {
      promesse = null
      throw erreur
    })
  }
  return promesse
}

export async function mettreFiche(fiche: FicheStockee): Promise<void> {
  const db = await ouvrirBase()
  await db.put(MAGASIN, fiche)
}

export function canalOuvert(etat: EtatEnvoi | undefined, reessayable: boolean | undefined): boolean {
  if (etat === 'ok') return false
  if (etat === 'erreur' && reessayable === false) return false
  return true
}

export function statutDesEtats(fiche: Pick<FicheStockee, 'statut' | 'etatBase' | 'etatSharePoint'>): StatutSync {
  if (fiche.statut === 'brouillon') return 'brouillon'
  if (fiche.etatBase === 'ok' && fiche.etatSharePoint === 'ok') return 'depose'
  if (fiche.etatBase === 'envoi' || fiche.etatSharePoint === 'envoi') return 'envoi'
  if (fiche.etatBase === 'ok' || fiche.etatSharePoint === 'ok') return 'partiel'
  if (fiche.etatBase === 'erreur' || fiche.etatSharePoint === 'erreur') return 'erreur'
  return 'attente'
}

/**
 * Les fiches d'avant la séparation n'ont qu'un statut global.
 * Une fiche déposée est confirmée des deux côtés. Une erreur connue
 * concerne la base : le PDF, lui, n'avait pas été envoyé.
 */
export function avecEtats(fiche: FicheStockee): FicheStockee {
  if (fiche.etatBase && fiche.etatSharePoint) return fiche
  if (fiche.statut === 'brouillon') {
    return {
      ...fiche,
      etatBase: 'attente',
      etatSharePoint: 'attente',
      baseErreur: null,
      sharePointErreur: null,
      baseReessayable: false,
      sharePointReessayable: false
    }
  }
  if (fiche.statut === 'depose') {
    return {
      ...fiche,
      etatBase: 'ok',
      etatSharePoint: 'ok',
      baseErreur: null,
      sharePointErreur: null,
      baseReessayable: true,
      sharePointReessayable: true
    }
  }
  const baseConnue = fiche.statut === 'erreur' || (fiche.statut === 'attente' && Boolean(fiche.erreur))
  const suivante: FicheStockee = {
    ...fiche,
    etatBase: baseConnue ? 'erreur' : 'attente',
    etatSharePoint: fiche.sharePointUrl ? 'ok' : 'attente',
    baseErreur: baseConnue ? fiche.erreur : null,
    sharePointErreur: null,
    baseReessayable: fiche.reessayable !== false,
    sharePointReessayable: true
  }
  suivante.statut = statutDesEtats(suivante)
  return suivante
}

async function ranger(fiche: FicheStockee): Promise<FicheStockee> {
  const complete = avecEtats(fiche)
  if (complete !== fiche) {
    const db = await ouvrirBase()
    await db.put(MAGASIN, complete)
  }
  return complete
}

export async function lireFiche(id: string): Promise<FicheStockee | undefined> {
  const db = await ouvrirBase()
  const fiche = await db.get(MAGASIN, id)
  return fiche ? ranger(fiche) : undefined
}

export async function patchFiche(
  id: string,
  patch: Partial<FicheStockee>
): Promise<FicheStockee | null> {
  const db = await ouvrirBase()
  const actuelle = await db.get(MAGASIN, id)
  if (!actuelle) return null
  const suivante: FicheStockee = { ...actuelle, ...patch, id }
  await db.put(MAGASIN, suivante)
  return suivante
}

export async function supprimerFiche(id: string): Promise<void> {
  const db = await ouvrirBase()
  await db.delete(MAGASIN, id)
}

export async function toutesLesFiches(): Promise<FicheStockee[]> {
  const db = await ouvrirBase()
  const toutes = await db.getAll(MAGASIN)
  return Promise.all(toutes.map((fiche) => ranger(fiche)))
}

/** Une fiche encore incomplète. Le canal déjà confirmé n'est pas renvoyé. */
export function aEncoreUnEnvoi(fiche: FicheStockee): boolean {
  if (fiche.statut === 'brouillon') return false
  return (
    canalOuvert(fiche.etatBase, fiche.baseReessayable) ||
    canalOuvert(fiche.etatSharePoint, fiche.sharePointReessayable)
  )
}

/** Celles qu'un échec réseau doit renvoyer. Un brouillon n'est pas dans la file. */
export async function fichesADeposer(): Promise<FicheStockee[]> {
  const toutes = await toutesLesFiches()
  return toutes.filter(aEncoreUnEnvoi).sort((a, b) => a.dateISO.localeCompare(b.dateISO))
}
