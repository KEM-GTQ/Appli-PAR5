import { copyFileSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "fs";
import { resolve } from "path";
import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { viteSingleFile } from "vite-plugin-singlefile";
import { VitePWA } from "vite-plugin-pwa";

/**
 * Application terrain Génitique — compilée en UN SEUL fichier HTML autonome
 * (JS, CSS, logo et icônes inlinés). Elle s'ouvre sur un téléphone sans réseau.
 *
 *   npm run build     → dist/index.html
 *   npm run publier   → copie aussi vers le Bureau
 */

/**
 * `vite-plugin-singlefile` inline le JS, le CSS et les images référencées
 * depuis le CODE, mais laisse sortir en fichiers séparés celles référencées par
 * des balises <link> de l'index (les icônes du raccourci). Sans ce
 * post-traitement, le « fichier unique » pointe vers des images absentes.
 */
function inlinerIcones(): Plugin {
  return {
    name: "inliner-icones",
    closeBundle() {
      const dist = resolve(__dirname, "dist");
      const indexPath = resolve(dist, "index.html");
      let html = readFileSync(indexPath, "utf8");

      const restants: string[] = [];
      for (const fichier of readdirSync(dist)) {
        if (!/\.(png|jpe?g|svg|webp)$/i.test(fichier)) continue;
        const ref = `./${fichier}`;
        if (!html.includes(ref)) {
          restants.push(fichier);
          continue;
        }
        const type = fichier.endsWith(".svg") ? "image/svg+xml" : "image/png";
        const b64 = readFileSync(resolve(dist, fichier)).toString("base64");
        html = html.split(ref).join(`data:${type};base64,${b64}`);
        rmSync(resolve(dist, fichier));
      }

      writeFileSync(indexPath, html);
      if (restants.length) {
        console.warn(
          `[inliner-icones] non référencés, laissés en place : ${restants.join(", ")}`,
        );
      }
    },
  };
}

/**
 * Le numéro de version affiché dans la barre du haut vient d'ICI, jamais d'une
 * chaîne recopiée dans le code : deux sources finissent toujours par diverger,
 * et une version qui ment est pire que pas de version du tout.
 */
const version = JSON.parse(
  readFileSync(resolve(__dirname, "package.json"), "utf8"),
).version as string;

/** Les icônes du manifeste doivent être de vrais fichiers. On les recopie depuis src/assets à chaque lancement. */
function copierIconesPubliques(): void {
  const pub = resolve(__dirname, "public");
  mkdirSync(pub, { recursive: true });
  for (const nom of ["icone-180.png", "icone-192.png", "icone-512.png"]) {
    copyFileSync(resolve(__dirname, "src/assets", nom), resolve(pub, nom));
  }
}

export default defineConfig(({ command }) => {
  const pwa = command === "serve" || process.env.PAR5_PWA === "1";
  const fichierUnique = command === "build" && !pwa;
  if (pwa) copierIconesPubliques();

  /* Le fichier unique reste à /Appli-PAR5/. La version installable se dépose
     dans n'importe quel dossier : ses chemins sont relatifs. */
  const base = pwa && command === "build" ? "./" : "/Appli-PAR5/";

  return {
    base,
    publicDir: pwa ? "public" : false,
    define: {
      __VERSION__: JSON.stringify(version),
      __PWA__: JSON.stringify(pwa),
    },
    resolve: {
      alias: {
        "@": resolve(__dirname, "src"),
        ...(pwa
          ? {}
          : { "virtual:pwa-register": resolve(__dirname, "src/pwa-vide.ts") }),
      },
    },
    plugins: [
      react(),
      tailwindcss(),
      pwa &&
        VitePWA({
          strategies: "injectManifest",
          srcDir: "src",
          filename: "sw.js",
          registerType: "autoUpdate",
          injectRegister: false,
          includeAssets: ["icone-180.png", "icone-192.png", "icone-512.png"],
          manifest: {
            name: "Génitique PAR 5",
            short_name: "PAR 5",
            description: "Analyse de risques PAR 5 — Génitique",
            lang: "fr-CA",
            dir: "ltr",
            start_url: base,
            scope: base,
            id: base,
            display: "standalone",
            orientation: "portrait",
            background_color: "#000000",
            theme_color: "#000000",
            icons: [
              { src: "icone-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
              { src: "icone-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
              { src: "icone-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
            ],
          },
          injectManifest: {
            globPatterns: ["**/*.{js,css,html,ico,png,svg,webmanifest}"],
            maximumFileSizeToCacheInBytes: 4 * 1024 * 1024,
          },
          devOptions: {
            enabled: command === "serve",
            type: "module",
            navigateFallback: "index.html",
          },
        }),
      fichierUnique && viteSingleFile(),
      fichierUnique && inlinerIcones(),
    ].filter(Boolean),
    server: {
      proxy: {
        "/api": { target: "http://127.0.0.1:3001", changeOrigin: true },
      },
    },
    build: {
      outDir: pwa && command === "build" ? "dist-web" : "dist",
      emptyOutDir: true,
      assetsInlineLimit: fichierUnique ? 100_000_000 : 4096,
      chunkSizeWarningLimit: 3_000,
    },
  };
});
