import { useEffect, useRef, useState } from 'react'
import { ShieldAlert, FolderOpen, BookOpen, Sun, Moon } from 'lucide-react'
import { donneesPar5Exemple } from './data/par5'
import { SplashGtq, HAUTEUR_BARRE } from './components/SplashGtq'
import { Logo } from './components/Logo'
import { Installer } from './components/Installer'
import { Par5Form, type Ouverture } from './pages/par5/Par5Form'
import { MesFiches } from './pages/par5/MesFiches'
import { Definitions } from './pages/par5/Definitions'
import { dupliquerPar5, ouvrirStockage, type Par5Enregistre } from './data/par5-store'
import { authRequise, compteActif, connecter, preparerAuth } from './data/auth-par5'
import { aDesFichesADeposer, flushPending, inscrireSyncArrierePlan } from './data/sync-par5'
import { appliquerTheme, lireTheme, type Theme } from './data/theme'

/**
 * Coque de l'application terrain.
 *
 * Un onglet = une entrée dans ONGLETS. Le premier est le PAR 5 ; les ajouts
 * suivants n'ont qu'à s'inscrire ici.
 */
const ONGLETS = [
  { id: 'par5', libelle: 'PAR 5', Icone: ShieldAlert },
  { id: 'fiches', libelle: 'Mes fiches', Icone: FolderOpen },
  { id: 'definitions', libelle: 'Définitions', Icone: BookOpen }
] as const

type OngletId = (typeof ONGLETS)[number]['id']

function aujourdhui(): string {
  const d = new Date()
  const p = (n: number): string => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`
}

export function App() {
  const [splash, setSplash] = useState(true)
  const [theme, setTheme] = useState<Theme>(lireTheme)
  const [onglet, setOnglet] = useState<OngletId>('par5')
  /** Fiche à rouvrir — ou données d'une fiche dupliquée, qui arrivent sans id. */
  const [aReprendre, setAReprendre] = useState<Ouverture | null>(null)
  /** Y a-t-il du travail non enregistré dans le formulaire ? */
  const [enCours, setEnCours] = useState(false)
  /** Ligne d'état du dépôt. Vide quand tout va bien : l'écran reste celui du PAR 5. */
  const [annonce, setAnnonce] = useState<string | null>(null)
  const [connecte, setConnecte] = useState(() => navigator.onLine)
  const [compte, setCompte] = useState<string | null>(null)
  const [reauth, setReauth] = useState(false)
  const stockageBloque = useRef(false)
  const enCoursRef = useRef(false)
  const reauthRef = useRef(false)
  enCoursRef.current = enCours
  reauthRef.current = reauth

  function basculerTheme(): void {
    const t: Theme = theme === 'sombre' ? 'clair' : 'sombre'
    setTheme(t)
    appliquerTheme(t)
  }

  function reprendre(o: Ouverture): void {
    setAReprendre(o)
    setOnglet('par5')
  }

  function dupliquer(f: Par5Enregistre): void {
    reprendre({ donnees: dupliquerPar5(f, aujourdhui()) })
  }

  function remplirExemple(): void {
    if (
      enCours &&
      !window.confirm(
        'Remplacer la fiche en cours par l’exemple ?\n\nCe qui n’est pas enregistré sera perdu.'
      )
    )
      return
    reprendre({ donnees: donneesPar5Exemple(aujourdhui()), exemple: true })
  }

  /**
   * LE LOGO REDÉMARRE L'APPLICATION — même geste dans tous les outils Génitique.
   *
   * C'est un vrai rechargement (`location.reload`), pas une remise à zéro des
   * états : on veut retrouver l'application exactement comme au premier
   * lancement, écran d'ouverture compris. Sur un fichier HTML unique posé sur
   * le Bureau, c'est aussi la façon la plus sûre de repartir propre quand
   * quelque chose s'est mal passé.
   *
   * ⚠️ Mais on ne l'emporte JAMAIS en silence sur une fiche à moitié remplie.
   * Un PAR 5 se remplit debout, souvent avec des gants : le logo est en haut à
   * gauche, exactement là où le pouce se pose en tenant le téléphone. Sans
   * cette question, le geste le plus naturel du monde effacerait vingt minutes
   * de travail.
   */
  useEffect(() => {
    let annule = false
    const stockage = stockageBloque

    async function sante(): Promise<void> {
      if (stockage.current || !navigator.onLine || location.protocol === 'file:') return
      try {
        const base = (import.meta.env.VITE_API_URL || '').replace(/\/$/, '')
        const reponse = await fetch(`${base}/api/health`, { cache: 'no-store' })
        if (!reponse.ok) throw new Error('sante')
        const corps = (await reponse.json()) as { sharePointConfigured?: boolean }
        if (annule || stockage.current) return
        if (!corps.sharePointConfigured) {
          setAnnonce('SharePoint n’est pas configuré. Les fiches restent sur ce téléphone.')
        } else if (!reauthRef.current) {
          setAnnonce(null)
        }
      } catch {
        if (!annule && !stockage.current) {
          setAnnonce('Le serveur ne répond pas. Les fiches restent sur ce téléphone.')
        }
      }
    }

    function enLigne(): void {
      setConnecte(true)
      if (stockage.current) return
      void sante()
      void inscrireSyncArrierePlan()
      if (authRequise() && !compteActif()) {
        setReauth(true)
        void aDesFichesADeposer().then((enAttente) => {
          if (!enAttente) {
            setAnnonce(null)
            return
          }
          setAnnonce('Connexion au réseau. Connectez-vous pour envoyer les fiches enregistrées sur ce téléphone.')
          if (!enCoursRef.current) void connecter()
        })
        return
      }
      setAnnonce(null)
      void flushPending(true)
    }

    function horsLigne(): void {
      setConnecte(false)
      if (stockage.current) return
      setAnnonce(
        authRequise() && !compteActif()
          ? 'Hors réseau. Vous pouvez remplir des PAR 5. La connexion Microsoft sera demandée au retour du réseau.'
          : 'Hors réseau. Les fiches restent sur ce téléphone et partent au retour de la connexion.'
      )
    }

    function messageAgent(event: MessageEvent): void {
      if (event.data?.type === 'par5-sync') window.dispatchEvent(new CustomEvent('par5-sync'))
    }

    function demanderReauth(): void {
      setReauth(true)
      const actif = compteActif()
      setCompte(actif?.nom ?? null)
      if (!actif) {
        setAnnonce('Connectez-vous avec votre compte Microsoft pour envoyer les fiches enregistrées sur ce téléphone.')
        if (navigator.onLine && !enCoursRef.current) void connecter()
        return
      }
      setAnnonce('Session Microsoft expirée. Les fiches restent sur ce téléphone jusqu’à la reconnexion.')
    }

    void ouvrirStockage().catch(() => {
      stockage.current = true
      if (!annule) {
        setAnnonce('Ce navigateur bloque le stockage : les fiches ne peuvent pas être gardées hors ligne.')
      }
    })

    void preparerAuth().then((erreurAuth) => {
      if (annule) return
      const actif = compteActif()
      setCompte(actif?.nom ?? null)
      if (erreurAuth) setAnnonce(erreurAuth)
      if (authRequise() && !actif) {
        setReauth(true)
        if (!navigator.onLine) {
          setAnnonce('Hors réseau. Vous pouvez remplir des PAR 5. La connexion Microsoft sera demandée au retour du réseau.')
          return
        }
        if (erreurAuth) return
        void aDesFichesADeposer().then((enAttente) => {
          if (annule) return
          if (!enAttente) return
          setAnnonce('Connectez-vous pour envoyer les fiches enregistrées sur ce téléphone.')
          if (!enCoursRef.current) void connecter()
        })
        return
      }
      if (location.protocol === 'file:') return
      if (navigator.onLine) {
        void sante()
        void inscrireSyncArrierePlan()
        void flushPending()
      } else {
        horsLigne()
      }
    })

    window.addEventListener('online', enLigne)
    window.addEventListener('offline', horsLigne)
    window.addEventListener('par5-reauth', demanderReauth)
    navigator.serviceWorker?.addEventListener('message', messageAgent)
    const timer = window.setInterval(() => {
      if (navigator.onLine) void sante()
      if (navigator.onLine) void flushPending()
    }, 30_000)

    return () => {
      annule = true
      window.removeEventListener('online', enLigne)
      window.removeEventListener('offline', horsLigne)
      window.removeEventListener('par5-reauth', demanderReauth)
      navigator.serviceWorker?.removeEventListener('message', messageAgent)
      window.clearInterval(timer)
    }
  }, [])

  function redemarrer(): void {
    if (
      enCours &&
      !window.confirm(
        'Redémarrer l’application ?\n\nLa fiche en cours n’est pas enregistrée : elle sera perdue.'
      )
    )
      return
    window.location.reload()
  }

  return (
    <div className="relative flex h-full flex-col bg-gtq-fond">
      {/* Barre du haut — même noir que l'écran d'ouverture, qui s'y fond. */}
      <header className="safe-haut shrink-0 bg-gtq-noir">
        <div
          className="flex items-center gap-3 px-4"
          style={{ height: HAUTEUR_BARRE }}
        >
          <button
            type="button"
            onClick={redemarrer}
            aria-label="Redémarrer l’application"
            title="Redémarrer l’application"
            className="-ml-2 flex h-11 items-center rounded-lg px-2"
          >
            <Logo className="h-7 w-auto" />
          </button>
          <span className="h-5 w-px bg-white/15" />
          {/* Le titre cède la place avant la version : sur un téléphone de
              chantier, savoir QUELLE build on a dans les mains est la seule
              façon de diagnostiquer quoi que ce soit à distance. */}
          <div className="flex min-w-0 items-baseline gap-1.5">
            <h1 className="truncate text-[15px] font-semibold tracking-tight text-white">
              {ONGLETS.find((o) => o.id === onglet)?.libelle}
            </h1>
            <span className="sur-noir-2 shrink-0 text-[11px] font-medium tabular-nums">
              v{__VERSION__}
            </span>
          </div>

          {/* La barre reste noire dans les deux thèmes : ses encres sont donc
              posées en blanc littéral, pas avec les jetons qui basculent. */}
          <div className="ml-auto flex items-center gap-1">
            {compte && (
              <span className="max-w-16 truncate px-1 text-[12px] font-medium text-white/70" title={compte}>
                {compte}
              </span>
            )}
            {reauth && (
              <button
                type="button"
                onClick={() => void connecter()}
                disabled={!connecte}
                className="flex h-9 items-center rounded-full px-2.5 text-[12px] font-semibold text-amber-300 disabled:opacity-50"
              >
                Connexion
              </button>
            )}
            <button
              type="button"
              onClick={remplirExemple}
              className="flex h-9 items-center rounded-full px-2.5 text-[12px] font-semibold text-white/90"
            >
              Exemple
            </button>
            <span
              role="status"
              aria-live="polite"
              className="flex items-center gap-1.5 px-1.5 text-[12px] font-medium text-white/80"
            >
              <span
                aria-hidden="true"
                className={`h-2 w-2 rounded-full ${connecte ? 'bg-emerald-400' : 'bg-amber-400'}`}
              />
              {connecte ? 'En ligne' : 'Hors ligne'}
            </span>
          <button
            type="button"
            onClick={basculerTheme}
            aria-label={theme === 'sombre' ? 'Passer en mode clair' : 'Passer en mode sombre'}
            className="-mr-2 flex h-11 w-11 items-center justify-center rounded-full text-white/75"
          >
            {theme === 'sombre' ? (
              <Sun size={20} aria-hidden="true" />
            ) : (
              <Moon size={20} aria-hidden="true" />
            )}
          </button>
          </div>
        </div>
        <div className="barre-gtq" />
      </header>

      {/* L’invitation à poser l’icône sur l’écran d’accueil. Elle ne s’affiche
          que là où elle a un sens : hébergée, pas déjà installée, pas masquée. */}
      <Installer />

      {annonce && (
        <p
          role="status"
          className="shrink-0 border-b border-gtq-filet bg-gtq-surface px-4 py-2 text-center text-[13px] text-gtq-texte"
        >
          {annonce}
        </p>
      )}

      <main className="min-h-0 flex-1 overflow-y-auto">
        {onglet === 'par5' ? (
          <Par5Form
            reprise={aReprendre}
            onRepriseConsommee={() => setAReprendre(null)}
            onEnregistre={() => setOnglet('fiches')}
            onEnCours={setEnCours}
          />
        ) : onglet === 'fiches' ? (
          <MesFiches
            onReprendre={reprendre}
            onDupliquer={dupliquer}
            onNouveau={() => setOnglet('par5')}
          />
        ) : (
          <Definitions />
        )}
      </main>

      {/* Barre d'onglets — cibles de 56 px, atteignables au pouce. */}
      <nav className="safe-bas shrink-0 border-t border-gtq-filet bg-gtq-noir">
        <div className="flex">
          {ONGLETS.map(({ id, libelle, Icone }) => {
            const actif = id === onglet
            return (
              <button
                key={id}
                type="button"
                onClick={() => setOnglet(id)}
                aria-current={actif ? 'page' : undefined}
                className="flex flex-1 flex-col items-center justify-center gap-1 py-2.5"
              >
                <Icone
                  size={22}
                  strokeWidth={actif ? 2.2 : 1.7}
                  className={actif ? 'sur-noir-accent' : 'sur-noir-2'}
                  aria-hidden="true"
                />
                <span
                  className={`text-[11px] font-medium ${actif ? 'sur-noir' : 'sur-noir-2'}`}
                >
                  {libelle}
                </span>
              </button>
            )
          })}
        </div>
      </nav>

      {splash && <SplashGtq onFini={() => setSplash(false)} />}
    </div>
  )
}
