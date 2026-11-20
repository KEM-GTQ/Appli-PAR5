import { HttpError } from "./errors.js";

const FILE_NAME_RE = /^PAR5-par5-[a-z0-9]{8}\.pdf$/;
let tokenCache = { key: "", token: "", expiresAt: 0 };

export function nomPdf(id) {
  const nom = `PAR5-${id}.pdf`;
  if (!FILE_NAME_RE.test(nom)) {
    throw new HttpError("Nom de PDF refusé.", 500, false);
  }
  return nom;
}

export function isSharePointConfigured(env) {
  return Boolean(
    env.GRAPH_TENANT_ID &&
      env.GRAPH_CLIENT_ID &&
      env.GRAPH_CLIENT_SECRET &&
      (env.SHAREPOINT_SITE_URL || env.SHAREPOINT_SITE_ID),
  );
}

export function parseSharePointLocation(value) {
  if (typeof value !== "string" || !value.trim()) return null;
  let url;
  try {
    url = new URL(value.trim());
  } catch {
    return null;
  }
  if (!url.hostname.endsWith(".sharepoint.com")) return null;

  const idPath = url.searchParams.get("id");
  const rawPath = decodeURIComponent(idPath || url.pathname);
  const siteMatch = rawPath.match(/^\/(sites|teams)\/[^/]+/i);
  if (!siteMatch || rawPath.includes("..")) return null;

  const sitePath = siteMatch[0];
  const remainder = rawPath
    .slice(sitePath.length)
    .replace(/\/Forms\/AllItems\.aspx$/i, "")
    .split("/")
    .map((segment) => segment.trim())
    .filter(Boolean);
  const library = remainder[0] && remainder[0].toLowerCase() !== "forms" ? remainder[0] : "";
  const folder = library ? remainder.slice(1).join("/") : "";

  return { hostname: url.hostname, sitePath, library, folder };
}

export function clearTokenCache() {
  tokenCache = { key: "", token: "", expiresAt: 0 };
}

export function missingSettings(env) {
  const missing = [];
  for (const key of ["GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET"]) {
    if (!env[key]) missing.push(key);
  }
  if (!env.SHAREPOINT_SITE_URL && !env.SHAREPOINT_SITE_ID) missing.push("SHAREPOINT_SITE_URL");
  return missing;
}

export function buildContentUrl(env, fileName, resolvedSiteId) {
  if (!FILE_NAME_RE.test(fileName)) {
    throw new HttpError("Nom de PDF refusé.", 500, false);
  }

  const parsedSite = parseSharePointLocation(env.SHAREPOINT_SITE_URL);
  const folder = env.SHAREPOINT_FOLDER || parsedSite?.folder || "Cursor";
  const segments = folder
    .split("/")
    .map((segment) => segment.trim())
    .filter((segment) => segment && segment !== "." && segment !== "..")
    .map(encodeURIComponent);
  const itemPath = [...segments, encodeURIComponent(fileName)].join("/");

  if (env.SHAREPOINT_DRIVE_ID) {
    return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(env.SHAREPOINT_DRIVE_ID)}/root:/${itemPath}:/content`;
  }

  const siteId = String(resolvedSiteId || env.SHAREPOINT_SITE_ID || "").trim();
  if (!siteId || /[\s/?#\\]/.test(siteId)) {
    throw new HttpError("SHAREPOINT_SITE_ID n'est pas valide.", 500, false);
  }
  return `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${itemPath}:/content`;
}

export async function resolveSiteId(env, fetchImpl, token) {
  const parsed = parseSharePointLocation(env.SHAREPOINT_SITE_URL);
  if (!parsed) return String(env.SHAREPOINT_SITE_ID || "").trim();

  const response = await fetchImpl(`https://graph.microsoft.com/v1.0/sites/${parsed.hostname}:${parsed.sitePath}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const json = await readJson(response);
  if (!response.ok || typeof json.id !== "string") {
    const message = String(json.error?.message || "Le site SharePoint n'a pas pu être ouvert.").slice(0, 300);
    const retryable = response.status === 429 || response.status >= 500;
    throw new HttpError(message, 502, retryable);
  }
  return json.id;
}

async function readJson(response) {
  const text = await response.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text.slice(0, 300) };
  }
}

export async function getAccessToken(env, fetchImpl = fetch) {
  const key = `${env.GRAPH_TENANT_ID}:${env.GRAPH_CLIENT_ID}`;
  if (
    fetchImpl === fetch &&
    tokenCache.token &&
    tokenCache.key === key &&
    tokenCache.expiresAt > Date.now() + 60_000
  ) {
    return tokenCache.token;
  }

  const body = new URLSearchParams({
    client_id: env.GRAPH_CLIENT_ID,
    client_secret: env.GRAPH_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const response = await fetchImpl(
    `https://login.microsoftonline.com/${encodeURIComponent(env.GRAPH_TENANT_ID)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  const json = await readJson(response);
  if (!response.ok || !json.access_token) {
    const message = String(json.error_description || json.error || "Jeton Microsoft refusé.").slice(0, 300);
    throw new HttpError(message, 502, response.status === 429 || response.status >= 500);
  }

  if (fetchImpl === fetch) {
    tokenCache = {
      key,
      token: json.access_token,
      expiresAt: Date.now() + Number(json.expires_in || 3600) * 1000,
    };
  }
  return json.access_token;
}

export async function uploadPdf(env, fileName, buffer, fetchImpl = fetch) {
  if (buffer.length > 4 * 1024 * 1024) {
    throw new HttpError("Le PDF dépasse 4 Mo, limite de l'envoi simple Graph.", 413, false);
  }

  const token = await getAccessToken(env, fetchImpl);
  const siteId = await resolveSiteId(env, fetchImpl, token);
  const response = await fetchImpl(buildContentUrl(env, fileName, siteId), {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/pdf",
    },
    body: buffer,
  });
  const json = await readJson(response);
  if (!response.ok) {
    const message = String(json.error?.message || json.raw || "Dépôt SharePoint impossible.").slice(0, 300);
    const retryable = response.status === 429 || response.status >= 500;
    throw new HttpError(message, 502, retryable);
  }

  return {
    webUrl: typeof json.webUrl === "string" ? json.webUrl : "",
    itemId: typeof json.id === "string" ? json.id : "",
  };
}
