import { useEffect, useState } from 'react'
import { FileText, Trash2, Plus, Share2, Download, CopyPlus } from 'lucide-react'
import { listerPar5, supprimerPar5, type Par5Enregistre, retouchee } from '@/data/par5-store'
import { construirePdfPar5, envoyerPdf, peutChoisirDossier } from '@/data/pdf-par5'
import { reessayerPar5 } from '@/data/sync-par5'
import type { EtatEnvoi } from '@/data/par5-db'
import type { Ouverture } from './Par5Form'

const COULEUR: Record<EtatEnvoi, string> = {
  attente: 'bg-neutral-400',
  envoi: 'bg-amber-400',
  ok: 'bg-emerald-500',
  erreur: 'bg-red-500'
}

const LIBELLE_ETAT: Record<EtatEnvoi, string> = {
  attente: 'en attente',
  envoi: 'en cours',
  ok: 'confirmé',
  erreur: 'échoué'
}

function messagesTransmission(f: Par5Enregistre): Array<{ ton: 'erreur' | 'attente'; texte: string }> {
  const lignes: Array<{ ton: 'erreur' | 'attente'; texte: string }> = []
  if (f.etatBase && f.etatBase !== 'ok' && f.baseErreur) {
    lignes.push({ ton: f.etatBase === 'erreur' ? 'erreur' : 'attente', texte: f.baseErreur })
  }
  if (f.etatSharePoint && f.etatSharePoint !== 'ok' && f.sharePointErreur && f.sharePointErreur !== f.baseErreur) {
    lignes.push({
      ton: f.etatSharePoint === 'erreur' ? 'erreur' : 'attente',
      texte: f.sharePointErreur
    })
  }
  return lignes
}

function peutReessayer(f: Par5Enregistre): boolean {
  if (!f.statut || f.statut === 'brouillon' || f.statut === 'depose' || f.statut === 'envoi') return false
  return f.etatBase !== 'ok' || f.etatSharePoint !== 'ok'
}

function Pastille({
  etat,
  libelle,
  detail
}: {
  etat: EtatEnvoi
  libelle: string
  detail?: string | null
}) {
  return (
    <span className="inline-flex items-center gap-1.5 text-[12px] text-gtq-texte2" title={detail || LIBELLE_ETAT[etat]}>
      <span aria-hidden="true" className={`inline-block h-3 w-3 rounded-full ${COULEUR[etat]}`} />
      {libelle}
      <span className="sr-only">
        {' '}
        : {LIBELLE_ETAT[etat]}
        {detail ? `. ${detail}` : ''}
      </span>
    </span>
  )
}

/**
 * Les PAR 5 de ce téléphone. La fiche reste sur l'appareil ; quand elle a été
 * envoyée, la ligne dit si le PDF est arrivé dans SharePoint.
 */
export function MesFiches({
  onReprendre,
  onDupliquer,
  onNouveau
}: {
  onReprendre: (o: Ouverture) => void
  onDupliquer: (f: Par5Enregistre) => void
  onNouveau: () => void
}) {
  const [fiches, setFiches] = useState<Par5Enregistre[]>([])
  /* Sur un poste le bouton enregistre où l’on veut ; sur un téléphone il partage. */
  const surPoste = peutChoisirDossier()

  useEffect(() => {
    let vivant = true
    const charger = (): void => {
      void listerPar5().then((liste) => {
        if (vivant) setFiches(liste)
      })
    }
    charger()
    window.addEventListener('par5-sync', charger)
    return () => {
      vivant = false
      window.removeEventListener('par5-sync', charger)
    }
  }, [])

  async function envoyer(f: Par5Enregistre): Promise<void> {
    await envoyerPdf(construirePdfPar5(f.donnees, f.dateISO))
  }

  async function supprimer(id: string): Promise<void> {
    await supprimerPar5(id)
    setFiches(await listerPar5())
  }

  if (!fiches.length) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-4 px-8 text-center">
        <FileText size={40} className="text-gtq-texte2/60" aria-hidden="true" />
        <p className="text-[15px] text-gtq-texte2">
          Aucune fiche enregistrée sur ce téléphone.
        </p>
        <button
          type="button"
          onClick={onNouveau}
          className="inline-flex items-center gap-2 rounded-xl bg-gtq-bleu px-4 py-3 text-[15px] font-semibold encre-pastille"
        >
          <Plus size={18} aria-hidden="true" /> Remplir un PAR 5
        </button>
      </div>
    )
  }

  return (
    <ul className="space-y-2.5 px-4 py-4">
      {fiches.map((f) => {
        const presents = f.donnees.lignes.filter((l) => l.present === 'oui').length
        const photos = f.donnees.lignes.filter((l) => l.photo).length
        const date = new Date(f.dateISO)
        const titre = f.donnees.tache || f.donnees.installation || 'Fiche sans titre'
        return (
          <li
            key={f.id}
            className="rounded-xl border border-gtq-filet bg-gtq-surface p-3"
          >
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={() => onReprendre({ donnees: f.donnees, id: f.id })}
                className="min-w-0 flex-1 text-left"
              >
                <div className="truncate text-[15px] font-medium text-gtq-texte">{titre}</div>
                <div className="mt-0.5 text-[13px] text-gtq-texte2">
                  {f.donnees.noProjet ? `${f.donnees.noProjet} · ` : ''}
                  {date.toLocaleDateString('fr-CA')} ·{' '}
                  <span className={presents ? 'text-gtq-orange' : ''}>
                    {presents} risque{presents > 1 ? 's' : ''}
                  </span>
                  {photos > 0 && ` · ${photos} photo${photos > 1 ? 's' : ''}`}
                  {/* La date de l'analyse ne bouge pas ; une retouche se dit à part. */}
                  {retouchee(f) && f.majISO && (
                    <span className="text-gtq-texte2">
                      {' '}
                      · retouchée le {new Date(f.majISO).toLocaleDateString('fr-CA')}
                    </span>
                  )}
                </div>
              </button>
              <button
                type="button"
                aria-label={
                  surPoste
                    ? `Enregistrer le PDF de la fiche ${titre}`
                    : `Envoyer la fiche ${titre} en PDF`
                }
                onClick={() => void envoyer(f)}
                className="shrink-0 rounded-lg p-2.5 text-gtq-bleu"
              >
                {surPoste ? (
                  <Download size={18} aria-hidden="true" />
                ) : (
                  <Share2 size={18} aria-hidden="true" />
                )}
              </button>
              <button
                type="button"
                aria-label={`Supprimer la fiche ${titre}`}
                onClick={() => void supprimer(f.id)}
                className="shrink-0 rounded-lg p-2.5 text-gtq-texte2"
              >
                <Trash2 size={18} aria-hidden="true" />
              </button>
            </div>
            {f.statut !== 'brouillon' && (
              <div className="mt-2">
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
                  <Pastille etat={f.etatBase || 'attente'} libelle="Base" detail={f.baseErreur} />
                  <Pastille
                    etat={f.etatSharePoint || 'attente'}
                    libelle="SharePoint"
                    detail={f.sharePointErreur}
                  />
                  {f.etatSharePoint === 'ok' && f.sharePointUrl && (
                    <a
                      href={f.sharePointUrl}
                      target="_blank"
                      rel="noreferrer"
                      className="text-[12px] text-gtq-vert underline"
                    >
                      Ouvrir
                    </a>
                  )}
                  {f.statut === 'depose' && f.auteur && (
                    <span className="text-[12px] text-gtq-texte2">{f.auteur}</span>
                  )}
                </div>
                {messagesTransmission(f).map((ligne) => (
                  <p
                    key={ligne.texte}
                    className={`mt-1 text-[12px] ${ligne.ton === 'erreur' ? 'text-red-500' : 'text-gtq-texte2'}`}
                  >
                    {ligne.texte}
                  </p>
                ))}
                {peutReessayer(f) && (
                  <button
                    type="button"
                    className="mt-1 text-[12px] text-gtq-texte underline"
                    onClick={() => void reessayerPar5(f.id)}
                  >
                    Réessayer
                  </button>
                )}
              </div>
            )}

            {/*
             * REPRENDRE POUR AUJOURD'HUI.
             *
             * Une équipe fait la même tâche trois jours de suite : les seize
             * réponses sont identiques, seules les dates changent. Recocher
             * seize cases pour la troisième fois, c'est recocher vite — donc ne
             * plus relire. Le bouton ne fait pas gagner du temps, il rend à
             * l'analyse une chance d'être RELUE.
             *
             * Il est en toutes lettres et non en icône : c'est le geste qu'on
             * fait en arrivant sur le chantier, il doit se trouver sans
             * chercher, avec des gants.
             */}
            <button
              type="button"
              onClick={() => onDupliquer(f)}
              className="mt-2.5 flex min-h-[44px] w-full items-center justify-center gap-2 rounded-lg border border-gtq-filet px-3 py-2.5 text-[14px] font-medium text-gtq-texte"
            >
              <CopyPlus size={16} aria-hidden="true" /> Reprendre pour aujourd’hui
            </button>
          </li>
        )
      })}
    </ul>
  )
}
