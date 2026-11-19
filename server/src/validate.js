import { HttpError } from "./errors.js";
import { nomPdf } from "./sharepoint.js";

const ID_RE = /^par5-[a-z0-9]{8}$/;
const MAX_PDF_BYTES = 4 * 1024 * 1024;

export class ValidationError extends HttpError {
  constructor(message) {
    super(message, 400, false);
    this.name = "ValidationError";
  }
}

function optionalString(value, name, max) {
  if (value == null) return "";
  if (typeof value !== "string") throw new ValidationError(`${name} doit être du texte.`);
  const trimmed = value.trim();
  if (trimmed.length > max) throw new ValidationError(`${name} est trop long.`);
  return trimmed;
}

function requiredDate(value, name) {
  if (typeof value !== "string" || Number.isNaN(Date.parse(value))) {
    throw new ValidationError(`${name} doit être une date.`);
  }
  return new Date(value).toISOString();
}

function decodePdf(value) {
  if (typeof value !== "string" || !value.trim()) {
    throw new ValidationError("Le PDF est requis.");
  }
  if (value.length > 8_000_000) throw new ValidationError("Le PDF est trop lourd.");
  const cleaned = value.replace(/^data:application\/pdf;base64,/i, "");
  const buffer = Buffer.from(cleaned, "base64");
  if (buffer.length < 5 || buffer.subarray(0, 4).toString("utf8") !== "%PDF") {
    throw new ValidationError("Le fichier n'est pas un PDF.");
  }
  if (buffer.length > MAX_PDF_BYTES) {
    throw new ValidationError("Le PDF doit faire 4 Mo ou moins.");
  }
  return buffer;
}

export function validateFiche(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    throw new ValidationError("La fiche doit être un objet JSON.");
  }

  const id = optionalString(body.id, "id", 16).toLowerCase();
  if (!ID_RE.test(id)) throw new ValidationError("id invalide.");

  return {
    id,
    noProjet: optionalString(body.noProjet, "noProjet", 80),
    installation: optionalString(body.installation, "installation", 200),
    emplacement: optionalString(body.emplacement, "emplacement", 200),
    tache: optionalString(body.tache, "tache", 300),
    dateISO: requiredDate(body.dateISO, "dateISO"),
    majISO: body.majISO ? requiredDate(body.majISO, "majISO") : null,
    pdfName: nomPdf(id),
    pdf: decodePdf(body.pdfBase64),
  };
}
