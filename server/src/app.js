import crypto from "node:crypto";
import express from "express";
import { HttpError } from "./errors.js";
import { enregistrerFiche as ecrireMysql, isDatabaseConfigured } from "./mysql.js";
import {
  isSharePointConfigured,
  missingSettings,
  uploadPdf as uploadPdfToSharePoint,
} from "./sharepoint.js";
import { validateFiche, ValidationError } from "./validate.js";
import { isEntraConfigured, verifierEntra } from "./entra.js";
import { listerProjets as lireProjetsSharePoint } from "./projets.js";

function asyncRoute(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

function apiKeyOk(expected, provided) {
  const left = Buffer.from(expected);
  const right = Buffer.from(provided);
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function originesAutorisees(env) {
  return String(env.CORS_ORIGIN || "")
    .split(",")
    .map((origine) => origine.trim())
    .filter(Boolean);
}

export function createApp({
  env = process.env,
  uploadPdf = uploadPdfToSharePoint,
  pool = null,
  enregistrerFiche = null,
  verifierJeton = null,
  listerProjets = null,
} = {}) {
  const app = express();
  app.disable("x-powered-by");
  app.set("trust proxy", 1);
  app.use(express.json({ limit: "9mb" }));
  app.use((req, res, next) => {
    res.set("X-Content-Type-Options", "nosniff");
    res.set("Referrer-Policy", "no-referrer");
    if (req.path.startsWith("/api")) res.set("Cache-Control", "no-store");

    const origine = req.get("Origin");
    const permises = originesAutorisees(env);
    if (origine && permises.includes(origine)) {
      res.set("Access-Control-Allow-Origin", origine);
      res.set("Vary", "Origin");
      res.set("Access-Control-Allow-Headers", "Content-Type, X-API-Key, Authorization");
      res.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    }
    if (req.method === "OPTIONS") {
      res.status(204).end();
      return;
    }
    next();
  });

  const verifier = verifierJeton || (isEntraConfigured(env) ? verifierEntra(env) : null);

  function requireUser(req, res, next) {
    if (!verifier) return next();
    const entete = req.get("Authorization") || "";
    const token = entete.startsWith("Bearer ") ? entete.slice("Bearer ".length).trim() : "";
    if (!token) {
      res.status(401).json({ error: "Connexion Microsoft requise.", retryable: true, code: "auth" });
      return;
    }
    Promise.resolve()
      .then(() => verifier(token))
      .then((compte) => {
        req.user = compte;
        next();
      })
      .catch((error) => {
        console.error(`Jeton refusé : ${error.message}`);
        if (!res.headersSent) {
          res.status(401).json({ error: "Session Microsoft expirée.", retryable: true, code: "auth" });
        }
      });
  }

  function requireApiKey(req, res, next) {
    if (!env.API_KEY) return next();
    const provided = req.get("X-API-Key") || "";
    if (!apiKeyOk(env.API_KEY, provided)) {
      res.status(401).json({ error: "Non autorisé.", retryable: false });
      return;
    }
    next();
  }

  app.get("/api/health", asyncRoute(async (req, res) => {
    let databaseOk = false;
    if (pool && isDatabaseConfigured(env)) {
      try {
        await pool.query("SELECT 1");
        databaseOk = true;
      } catch (error) {
        console.error(`MySQL injoignable : ${error.message}`);
      }
    }
    res.json({
      ok: true,
      databaseConfigured: isDatabaseConfigured(env),
      databaseOk,
      sharePointConfigured: isSharePointConfigured(env),
      authRequired: Boolean(verifier),
    });
  }));

  function detailErreur(error) {
    const morceaux = [];
    if (error?.code) morceaux.push(error.code);
    if (error?.message) morceaux.push(error.message);
    if (Array.isArray(error?.errors)) {
      for (const cause of error.errors) {
        if (cause?.code || cause?.message) morceaux.push(cause.code || cause.message);
      }
    }
    return morceaux.filter(Boolean).join(" — ") || "erreur inconnue";
  }

  function messagePour(error, defaut) {
    if (error instanceof HttpError && error.message) return error.message;
    const code = error?.code || error?.errors?.find((cause) => cause?.code)?.code;
    if (typeof code === "string" && /^[A-Z0-9_]{2,40}$/.test(code)) return `${defaut} (${code})`;
    return defaut;
  }

  function echecCanal(error, defaut) {
    console.error(detailErreur(error));
    return {
      tente: true,
      ok: false,
      error: messagePour(error, defaut),
      retryable: error instanceof HttpError ? error.retryable : true,
    };
  }

  async function ecrireEnBase(fiche, uploaded) {
    const ecrire = enregistrerFiche || ((ficheEcrite, fichier) => ecrireMysql(pool, ficheEcrite, fichier));
    try {
      await ecrire(fiche, uploaded);
    } catch (error) {
      if (error instanceof HttpError) throw error;
      console.error(`Écriture MySQL échouée pour ${fiche.id} : ${detailErreur(error)}`);
      throw new HttpError(messagePour(error, "L'enregistrement dans la base a échoué."), 502, true);
    }
  }

  /* Un essai précédent a déjà le lien. On le réécrit avec la fiche,
     sans renvoyer le PDF. */
  function lienDejaConnu(body) {
    const url = body?.sharePointUrl;
    if (typeof url !== "string") return null;
    const trimmed = url.trim();
    if (!trimmed.startsWith("https://") || trimmed.length > 2048) return null;
    return { webUrl: trimmed, itemId: null };
  }

  function operationsDemandees(body) {
    const brut = body?.tenter;
    return {
      base: brut?.base !== false,
      sharePoint: brut?.sharePoint !== false,
    };
  }

  /* La liste sert à remplir la fiche, y compris avant la connexion Microsoft.
     Le dépôt, lui, reste protégé. */
  app.get("/api/projets", asyncRoute(async (req, res) => {
    if (!isSharePointConfigured(env)) {
      throw new HttpError("SharePoint n'est pas configuré.", 503, true);
    }
    const lire = listerProjets || lireProjetsSharePoint;
    const projets = await lire(env);
    res.json({ projets });
  }));

  app.post("/api/par5", requireUser, requireApiKey, asyncRoute(async (req, res) => {
    const tenter = operationsDemandees(req.body);
    if (!tenter.base && !tenter.sharePoint) {
      throw new HttpError("Rien à envoyer.", 400, false);
    }

    /* Un PDF illisible empêche SharePoint. Il n'empêche pas d'écrire la fiche. */
    let fiche;
    let erreurPdf = null;
    try {
      fiche = validateFiche(req.body, { pdfRequis: tenter.sharePoint });
    } catch (error) {
      const pdfEnCause = error instanceof ValidationError && /pdf|fichier/i.test(error.message);
      if (!pdfEnCause || !tenter.base || !tenter.sharePoint) throw error;
      erreurPdf = error;
      fiche = validateFiche(req.body, { pdfRequis: false });
    }

    const auteur = req.user?.preferred_username || req.user?.name || "";
    fiche.auteur = auteur.slice(0, 320);
    const lienConnu = lienDejaConnu(req.body);

    async function operationBase() {
      if (!isDatabaseConfigured(env)) {
        throw new HttpError("La base n'est pas configurée.", 503, true);
      }
      await ecrireEnBase(fiche, lienConnu);
    }

    async function operationSharePoint() {
      if (erreurPdf) throw erreurPdf;
      const missing = missingSettings(env);
      if (missing.length) {
        console.error(`Dépôt non configuré. Manque : ${missing.join(", ")}`);
        throw new HttpError("Le dépôt SharePoint n'est pas configuré.", 503, true);
      }
      if (!fiche.pdf) throw new HttpError("Le PDF est requis.", 400, false);
      return uploadPdf(env, fiche.pdfName, fiche.pdf);
    }

    const [baseSettled, shareSettled] = await Promise.allSettled([
      tenter.base ? operationBase() : Promise.resolve(null),
      tenter.sharePoint ? operationSharePoint() : Promise.resolve(null),
    ]);

    let database = tenter.base
      ? baseSettled.status === "fulfilled"
        ? { tente: true, ok: true }
        : echecCanal(baseSettled.reason, "L'enregistrement dans la base a échoué.")
      : { tente: false, ok: false };

    let sharePoint = { tente: false, ok: false };
    let uploaded = null;
    if (tenter.sharePoint) {
      if (shareSettled.status === "fulfilled") {
        uploaded = shareSettled.value;
        if (!uploaded?.webUrl) console.warn(`Le dépôt de ${fiche.id} n'a pas renvoyé d'adresse.`);
        sharePoint = { tente: true, ok: true, url: uploaded?.webUrl || null };
      } else {
        sharePoint = echecCanal(shareSettled.reason, "Dépôt SharePoint impossible.");
      }
    }

    /* Le lien est écrit après les deux essais. Son échec ne retire pas
       une fiche déjà confirmée, ni un PDF déjà déposé. */
    if (sharePoint.ok && uploaded?.webUrl && isDatabaseConfigured(env) && (database.ok || !tenter.base)) {
      try {
        await ecrireEnBase(fiche, uploaded);
      } catch (error) {
        console.error(`Lien SharePoint non écrit pour ${fiche.id} : ${detailErreur(error)}`);
      }
    }

    const tentes = [database, sharePoint].filter((canal) => canal.tente);
    const reussis = tentes.filter((canal) => canal.ok).length;
    const status = reussis === 0 ? 502 : reussis === tentes.length ? 201 : 207;
    const syncedAt = database.ok && sharePoint.ok ? new Date().toISOString() : null;
    console.log(
      `PAR 5 ${fiche.id} : base ${database.tente ? (database.ok ? "ok" : "échec") : "laissée"}, SharePoint ${sharePoint.tente ? (sharePoint.ok ? "ok" : "échec") : "laissé"}.`,
    );
    res.status(status).json({
      id: fiche.id,
      database,
      sharePoint,
      sharePointUrl: sharePoint.url || lienConnu?.webUrl || null,
      syncedAt,
    });
  }));

  app.use((req, res) => {
    res.status(404).json({ error: "Introuvable.", retryable: false });
  });

  app.use((err, req, res, next) => {
    if (res.headersSent) return next(err);
    if (err?.type === "entity.too.large") {
      res.status(413).json({ error: "La fiche est trop lourde.", retryable: false });
      return;
    }
    if (err instanceof SyntaxError || err?.type === "entity.parse.failed") {
      res.status(400).json({ error: "Le corps doit être du JSON.", retryable: false });
      return;
    }
    if (err instanceof HttpError) {
      res.status(err.status).json({ error: err.message, retryable: err.retryable });
      return;
    }
    console.error(err);
    res.status(500).json({ error: "Dépôt échoué.", retryable: true });
  });

  return app;
}
