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

export type StatutSync = 'brouillon' | 'attente' | 'envoi' | 'depose' | 'erreur'

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

export async function lireFiche(id: string): Promise<FicheStockee | undefined> {
  const db = await ouvrirBase()
  return db.get(MAGASIN, id)
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
  return db.getAll(MAGASIN)
}

/** Celles qu'un échec réseau doit renvoyer. Un brouillon n'est pas dans la file. */
export async function fichesADeposer(): Promise<FicheStockee[]> {
  const toutes = await toutesLesFiches()
  return toutes
    .filter(
      (f) =>
        f.statut === 'attente' ||
        f.statut === 'envoi' ||
        (f.statut === 'erreur' && f.reessayable)
    )
    .sort((a, b) => a.dateISO.localeCompare(b.dateISO))
}
