import crypto from "node:crypto";
import express from "express";
import { HttpError } from "./errors.js";
import { enregistrerFiche, isDatabaseConfigured } from "./mysql.js";
import {
  isSharePointConfigured,
  missingSettings,
  uploadPdf as uploadPdfToSharePoint,
} from "./sharepoint.js";
import { validateFiche } from "./validate.js";

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

export function createApp({ env = process.env, uploadPdf = uploadPdfToSharePoint } = {}) {
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
      res.set("Access-Control-Allow-Headers", "Content-Type, X-API-Key");
      res.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    }
    if (req.method === "OPTIONS") {
      res.status(204).end();
      return;
    }
    next();
  });

  function requireApiKey(req, res, next) {
    if (!env.API_KEY) return next();
    const provided = req.get("X-API-Key") || "";
    if (!apiKeyOk(env.API_KEY, provided)) {
      res.status(401).json({ error: "Non autorisé.", retryable: false });
      return;
    }
    next();
  }

  app.get("/api/health", (req, res) => {
    res.json({
      ok: true,
      databaseConfigured: isDatabaseConfigured(env),
      databaseOk: false,
      sharePointConfigured: isSharePointConfigured(env),
    });
  });

  app.post("/api/par5", requireApiKey, asyncRoute(async (req, res) => {
    const fiche = validateFiche(req.body);
    const missing = missingSettings(env);
    if (missing.length) {
      console.error(`Dépôt non configuré. Manque : ${missing.join(", ")}`);
      throw new HttpError("Le dépôt SharePoint n'est pas configuré.", 503, true);
    }

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

    if (isDatabaseConfigured(env)) {
      await enregistrerFiche();
      console.log(`MySQL est renseigné pour ${fiche.id}, l'écriture n'est pas encore branchée.`);
    }

    const syncedAt = new Date().toISOString();
    console.log(`PDF ${fiche.pdfName} déposé dans SharePoint. Écriture MySQL inactive.`);
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
