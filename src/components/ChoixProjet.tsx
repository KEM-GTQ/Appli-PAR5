import { useEffect, useId, useMemo, useState } from 'react'
import { actualiserProjets, projetsEnCache, type Projet } from '@/data/projets'

const VISIBLES_MAX = 40

interface Props {
  id: string
  valeur: string
  onChoisir: (numero: string) => void
}

function correspond(projet: Projet, texte: string, valeur: string): boolean {
  const recherche = texte.trim().toLocaleLowerCase('fr')
  if (!recherche || recherche === valeur.trim().toLocaleLowerCase('fr')) return true
  return (
    projet.numero.toLocaleLowerCase('fr').includes(recherche) ||
    projet.phase.toLocaleLowerCase('fr').includes(recherche)
  )
}

/**
 * Menu cherchable des numéros de projet. La saisie libre reste possible :
 * une fiche déjà enregistrée, ou un téléphone sans liste en cache, ne doit
 * pas perdre son numéro.
 */
export function ChoixProjet({ id, valeur, onChoisir }: Props) {
  const listeId = useId()
  const [recherche, setRecherche] = useState(valeur)
  const [ouvert, setOuvert] = useState(false)
  const [projets, setProjets] = useState<Projet[]>(() => projetsEnCache())
  const [note, setNote] = useState('')
  const [index, setIndex] = useState(0)

  useEffect(() => {
    setRecherche(valeur)
  }, [valeur])

  useEffect(() => {
    let vivant = true
    void actualiserProjets()
      .then((liste) => {
        if (!vivant) return
        setProjets(liste)
        if (!liste.length) setNote('Aucun projet chargé. Le numéro peut être saisi.')
      })
      .catch(() => {
        if (!vivant) return
        if (!projetsEnCache().length) setNote('Liste des projets indisponible. Le numéro peut être saisi.')
      })
    return () => {
      vivant = false
    }
  }, [])

  const visibles = useMemo(
    () => projets.filter((projet) => correspond(projet, recherche, valeur)).slice(0, VISIBLES_MAX),
    [projets, recherche, valeur]
  )
  const trop = projets.filter((projet) => correspond(projet, recherche, valeur)).length > VISIBLES_MAX

  function choisir(numero: string): void {
    onChoisir(numero)
    setRecherche(numero)
    setOuvert(false)
  }

  function quitter(): void {
    setOuvert(false)
    const suivant = recherche.trim().slice(0, 255)
    if (suivant !== valeur) onChoisir(suivant)
  }

  return (
    <div>
      <input
        id={id}
        type="text"
        role="combobox"
        aria-expanded={ouvert}
        aria-controls={listeId}
        aria-autocomplete="list"
        autoComplete="off"
        value={recherche}
        placeholder="Rechercher un projet"
        onFocus={() => setOuvert(true)}
        onChange={(event) => {
          setRecherche(event.target.value)
          setOuvert(true)
          setIndex(0)
        }}
        onBlur={quitter}
        onKeyDown={(event) => {
          if (event.key === 'ArrowDown') {
            event.preventDefault()
            if (!ouvert) {
              setOuvert(true)
              setIndex(0)
              return
            }
            setIndex((courant) => Math.min(visibles.length - 1, courant + 1))
          } else if (event.key === 'ArrowUp') {
            event.preventDefault()
            setIndex((courant) => Math.max(0, courant - 1))
          } else if (event.key === 'Enter' && ouvert && visibles[index]) {
            event.preventDefault()
            choisir(visibles[index].numero)
          } else if (event.key === 'Escape') {
            setOuvert(false)
            setRecherche(valeur)
          }
        }}
      />
      {ouvert && (
        <ul
          id={listeId}
          role="listbox"
          className="mt-1 max-h-64 overflow-y-auto rounded-[10px] border border-gtq-filet bg-gtq-surface"
        >
          {visibles.length === 0 ? (
            <li className="px-3 py-3 text-[14px] text-gtq-texte2">
              Aucun projet correspondant. Le numéro saisi est conservé.
            </li>
          ) : (
            visibles.map((projet, i) => (
              <li key={projet.numero} role="presentation">
                <button
                  type="button"
                  role="option"
                  aria-selected={projet.numero === valeur || i === index}
                  onMouseDown={(event) => {
                    event.preventDefault()
                    choisir(projet.numero)
                  }}
                  className={`flex min-h-11 w-full items-baseline gap-2 px-3 py-2.5 text-left text-[16px] ${
                    i === index ? 'bg-gtq-bleu/10' : ''
                  }`}
                >
                  <span className="font-medium text-gtq-texte">{projet.numero}</span>
                  {projet.phase && <span className="text-[13px] text-gtq-texte2">{projet.phase}</span>}
                </button>
              </li>
            ))
          )}
          {trop && (
            <li className="px-3 py-2 text-[13px] text-gtq-texte2">Affinez la recherche pour voir la suite.</li>
          )}
        </ul>
      )}
      {note && !ouvert && <p className="mt-1.5 text-[13px] text-gtq-texte2">{note}</p>}
    </div>
  )
}
