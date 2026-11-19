/** Remplace `virtual:pwa-register` dans la construction en fichier unique, qui n'a pas d'agent de service. */
export function registerSW(): (recharger?: boolean) => Promise<void> {
  return () => Promise.resolve()
}
