import crypto from "node:crypto";
import express from "express";
import { HttpError } from "./errors.js";
import { enregistrerFiche as ecrireMysql, isDatabaseConfigured } from "./mysql.js";
import {
  isSharePointConfigured,
  missingSettings,
  uploadPdf as uploadPdfToSharePoint,
} from "./sharepoint.js";
import { validateFiche } from "./validate.js";
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

  async function ecrireEnBase(fiche, uploaded) {
    const ecrire = enregistrerFiche || ((ficheEcrite, fichier) => ecrireMysql(pool, ficheEcrite, fichier));
    try {
      await ecrire(fiche, uploaded);
    } catch (error) {
      if (error instanceof HttpError) throw error;
      console.error(`Écriture MySQL échouée pour ${fiche.id} : ${error.message}`);
      throw new HttpError("L'enregistrement dans la base a échoué.", 502, true);
    }
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
    const fiche = validateFiche(req.body);
    const auteur = req.user?.preferred_username || req.user?.name || "";
    fiche.auteur = auteur.slice(0, 320);
    const missing = missingSettings(env);
    if (missing.length) {
      console.error(`Dépôt non configuré. Manque : ${missing.join(", ")}`);
      throw new HttpError("Le dépôt SharePoint n'est pas configuré.", 503, true);
    }

    /* La fiche est écrite avant le PDF. Si la connexion tombe pendant l'envoi
       vers SharePoint, la ligne existe déjà et le prochain essai la complète. */
    if (isDatabaseConfigured(env)) await ecrireEnBase(fiche, null);

    let uploaded;
    try {
      uploaded = await uploadPdf(env, fiche.pdfName, fiche.pdf);
    } catch (error) {
      if (error instanceof HttpError) throw error;
      console.error(`Dépôt SharePoint échoué : ${error.message}`);
      throw new HttpError("Dépôt SharePoint impossible.", 502, true);
    }

    if (!uploaded.webUrl) {
      console.warn(`Le dépôt de ${fiche.id} n'a pas renvoyé d'adresse.`);
    }

    if (isDatabaseConfigured(env)) await ecrireEnBase(fiche, uploaded);

    const syncedAt = new Date().toISOString();
    console.log(
      `PDF ${fiche.pdfName} déposé dans SharePoint${isDatabaseConfigured(env) ? " et écrit dans MySQL" : ""}.`,
    );
    res.status(201).json({
      id: fiche.id,
      sharePointUrl: uploaded.webUrl || null,
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
