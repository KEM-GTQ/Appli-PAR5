/** Projet actif lu dans la liste SharePoint « Projets 2 ». */
export interface Projet {
  numero: string
  phase: string
}

const CLE = 'gtq-par5-projets-actifs'

function normaliser(liste: unknown): Projet[] {
  if (!Array.isArray(liste)) return []
  const vus = new Set<string>()
  const projets: Projet[] = []
  for (const item of liste) {
    if (!item || typeof item !== 'object') continue
    const brut = item as { numero?: unknown; phase?: unknown }
    const numero = typeof brut.numero === 'string' ? brut.numero.trim().slice(0, 255) : ''
    if (!numero || vus.has(numero)) continue
    vus.add(numero)
    const phase = typeof brut.phase === 'string' ? brut.phase.trim().slice(0, 80) : ''
    projets.push({ numero, phase })
  }
  return projets
}

/** Dernière liste gardée sur le téléphone, y compris hors réseau. */
export function projetsEnCache(): Projet[] {
  try {
    const brut = localStorage.getItem(CLE)
    if (!brut) return []
    const json = JSON.parse(brut) as { projets?: unknown }
    return normaliser(json.projets)
  } catch {
    return []
  }
}

function garder(projets: Projet[]): void {
  try {
    localStorage.setItem(CLE, JSON.stringify({ projets }))
  } catch {
    /* La liste reste utilisable pour cette ouverture. */
  }
}

/** Recharge la liste quand le réseau répond. Sinon, rend le cache. */
export async function actualiserProjets(): Promise<Projet[]> {
  const cache = projetsEnCache()
  if (location.protocol === 'file:' || navigator.onLine === false) return cache
  const base = (import.meta.env.VITE_API_URL || '').replace(/\/$/, '')
  const reponse = await fetch(`${base}/api/projets`, { cache: 'no-store' })
  if (!reponse.ok) throw new Error('liste')
  const corps = (await reponse.json()) as { projets?: unknown }
  const projets = normaliser(corps.projets)
  garder(projets)
  return projets
}
