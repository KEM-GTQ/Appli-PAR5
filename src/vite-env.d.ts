/// <reference types="vite/client" />

/** Injecté à la compilation depuis le champ `version` de package.json. */
declare const __VERSION__: string

/** Vrai dans `npm run dev` et dans la construction installable. Faux pour le fichier HTML unique. */
declare const __PWA__: boolean

interface ImportMetaEnv {
  readonly VITE_API_URL?: string
  readonly VITE_API_KEY?: string
}

declare module 'virtual:pwa-register' {
  export function registerSW(options?: { immediate?: boolean }): (recharger?: boolean) => Promise<void>
}

declare module '*.png' {
  const src: string
  export default src
}
