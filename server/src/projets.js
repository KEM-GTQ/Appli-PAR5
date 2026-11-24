import { HttpError } from "./errors.js";
import { getAccessToken } from "./sharepoint.js";

const CHAMP_PROJET = "Projet";
const CHAMP_ETAT = "État";
const ETAT_ACTIF = "Actif";
const CACHE_MS = 10 * 60 * 1000;
const PAGE_MAX = 100;
const NOM_MAX = 255;

let cache = { key: "", jusqua: 0, projets: [] };
let enCours = null;

export function viderCacheProjets() {
  cache = { key: "", jusqua: 0, projets: [] };
  enCours = null;
}

/** Adresse Graph du site. Une liste à la racine du locataire n'a pas de /sites/. */
export function urlGraphDuSite(urlSite) {
  let url;
  try {
    url = new URL(String(urlSite || "").trim());
  } catch {
    return null;
  }
  if (url.protocol !== "https:" || !url.hostname.endsWith(".sharepoint.com")) return null;
  const chemin = decodeURIComponent(url.pathname);
  if (chemin.includes("..")) return null;
  const site = chemin.match(/^\/(sites|teams)\/[^/]+/i);
  if (site) return `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${site[0]}`;
  return `https://graph.microsoft.com/v1.0/sites/${url.hostname}`;
}

function texte(valeur) {
  if (valeur == null) return "";
  if (typeof valeur === "string" || typeof valeur === "number") return String(valeur).trim();
  if (typeof valeur === "object") {
    return String(valeur.Label || valeur.lookupValue || valeur.Value || "").trim();
  }
  return "";
}

/**
 * Garde les projets actifs, une fois chacun. Le texte affiché est la colonne Projet.
 */
export function projetsDepuisItems(items, champs) {
  const nomChamp = champs?.nom || CHAMP_PROJET;
  const etatChamp = champs?.etat || CHAMP_ETAT;
  const actif = (champs?.actif || ETAT_ACTIF).trim().toLocaleLowerCase("fr");
  const vus = new Set();
  const projets = [];

  for (const item of items || []) {
    const fields = item?.fields || {};
    if (texte(fields[etatChamp]).toLocaleLowerCase("fr") !== actif) continue;
    const nom = texte(fields[nomChamp]).slice(0, NOM_MAX);
    const cle = nom.toLocaleLowerCase("fr");
    if (!nom || vus.has(cle)) continue;
    vus.add(cle);
    projets.push({ numero: nom, phase: "" });
  }

  return projets.sort((a, b) =>
    a.numero.localeCompare(b.numero, "fr", { numeric: true, sensitivity: "base" }),
  );
}

async function lireJson(fetchImpl, url, token) {
  const adresse = new URL(url);
  if (adresse.protocol !== "https:" || adresse.hostname !== "graph.microsoft.com") {
    throw new HttpError("Réponse SharePoint inattendue.", 502, true);
  }
  const response = await fetchImpl(adresse.href, { headers: { Authorization: `Bearer ${token}` } });
  const text = await response.text();
  let json = {};
  if (text) {
    try {
      json = JSON.parse(text);
    } catch {
      json = {};
    }
  }
  if (!response.ok) {
    const message = String(json.error?.message || "La liste des projets est indisponible.").slice(0, 300);
    throw new HttpError(message, 502, response.status === 429 || response.status >= 500);
  }
  return json;
}

function trouverColonne(colonnes, voulu) {
  const cle = voulu.trim().toLocaleLowerCase("fr");
  const colonne = colonnes.find(
    (c) =>
      String(c.displayName || "").trim().toLocaleLowerCase("fr") === cle ||
      String(c.name || "").trim().toLocaleLowerCase("fr") === cle,
  );
  if (!colonne?.name) throw new HttpError(`La colonne « ${voulu} » est introuvable.`, 502, true);
  return colonne.name;
}

async function charger(env, fetchImpl) {
  const urlSite = String(env.SHAREPOINT_PROJETS_SITE_URL || "").trim();
  const nomListe = String(env.SHAREPOINT_LISTE_PROJETS || "").trim();
  if (!urlSite || !nomListe) {
    throw new HttpError("Le site de la liste des projets n'est pas configuré.", 503, false);
  }
  const vouluProjet = String(env.SHAREPOINT_CHAMP_PROJET || CHAMP_PROJET).trim() || CHAMP_PROJET;
  const vouluEtat = String(env.SHAREPOINT_CHAMP_ETAT || CHAMP_ETAT).trim() || CHAMP_ETAT;
  const vouluActif = String(env.SHAREPOINT_ETAT_ACTIF || ETAT_ACTIF).trim() || ETAT_ACTIF;
  const cle = `${urlSite}:${nomListe}:${vouluProjet}:${vouluEtat}:${vouluActif}`;
  if (cache.key === cle && cache.jusqua > Date.now()) return cache.projets;

  const siteUrl = urlGraphDuSite(urlSite);
  if (!siteUrl) throw new HttpError("L'adresse du site des projets n'est pas valide.", 503, false);

  const token = await getAccessToken(env, fetchImpl);
  const site = await lireJson(fetchImpl, `${siteUrl}?$select=id`, token);
  if (!site.id) throw new HttpError("Le site des projets n'a pas pu être ouvert.", 502, true);

  let listeId = "";
  let listesUrl = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName&$top=200`;
  for (let page = 0; page < 10 && listesUrl && !listeId; page += 1) {
    const listes = await lireJson(fetchImpl, listesUrl, token);
    listeId = (listes.value || []).find((liste) => liste.displayName === nomListe)?.id || "";
    listesUrl = listes["@odata.nextLink"] || "";
  }
  if (!listeId) throw new HttpError(`La liste « ${nomListe} » est introuvable.`, 502, true);

  const baseListe = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(listeId)}`;
  const colonnes = await lireJson(fetchImpl, `${baseListe}/columns?$select=name,displayName&$top=200`, token);
  const champNom = trouverColonne(colonnes.value || [], vouluProjet);
  const champEtat = trouverColonne(colonnes.value || [], vouluEtat);
  const select = [champNom, champEtat].map(encodeURIComponent).join(",");

  let itemsUrl = `${baseListe}/items?$expand=fields($select=${select})&$top=200`;
  const items = [];
  for (let page = 0; page < PAGE_MAX && itemsUrl; page += 1) {
    const lot = await lireJson(fetchImpl, itemsUrl, token);
    items.push(...(lot.value || []));
    itemsUrl = lot["@odata.nextLink"] || "";
  }

  const projets = projetsDepuisItems(items, { nom: champNom, etat: champEtat, actif: vouluActif });
  cache = { key: cle, jusqua: Date.now() + CACHE_MS, projets };
  return projets;
}

export function listerProjets(env, fetchImpl = fetch) {
  if (!enCours) {
    enCours = charger(env, fetchImpl).finally(() => {
      enCours = null;
    });
  }
  return enCours;
}
