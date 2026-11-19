/*
 * Agent de service du PAR 5.
 *
 * Le cache sert l'application sans réseau. Les lectures /api ne sont pas
 * réécrites vers la page : un dépôt de PDF doit atteindre le serveur.
 * Background Sync reprend la file quand Chromium retrouve le réseau,
 * même si l'application a été fermée. iOS ignore cet événement.
 */
import { clientsClaim } from "workbox-core";
import { cleanupOutdatedCaches, createHandlerBoundToURL, precacheAndRoute } from "workbox-precaching";
import { NavigationRoute, registerRoute } from "workbox-routing";
import { flushPending } from "./data/sync-par5";

self.skipWaiting();
clientsClaim();
precacheAndRoute(self.__WB_MANIFEST);
cleanupOutdatedCaches();

try {
  registerRoute(
    new NavigationRoute(createHandlerBoundToURL(`${import.meta.env.BASE_URL}index.html`), {
      denylist: [/^\/api/],
    }),
  );
} catch (error) {
  console.warn(`Route de navigation ignorée : ${error.message}`);
}

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("sync", (event) => {
  if (event.tag !== "par5-sync") return;
  event.waitUntil(
    flushPending().then((result) => {
      if (result?.offline) throw new Error("Des PAR 5 attendent une connexion.");
    }),
  );
});
