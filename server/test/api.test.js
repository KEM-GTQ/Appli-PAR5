import assert from "node:assert/strict";
import test from "node:test";
import { createApp } from "../src/app.js";
import { buildContentUrl, parseSharePointLocation } from "../src/sharepoint.js";
import { validateFiche } from "../src/validate.js";

const configuredEnv = {
  GRAPH_TENANT_ID: "tenant-id",
  GRAPH_CLIENT_ID: "client-id",
  GRAPH_CLIENT_SECRET: "client-secret",
  SHAREPOINT_SITE_ID: "genitique.sharepoint.com,site-guid,web-guid",
  SHAREPOINT_FOLDER: "PAR5",
};

function pdfBase64() {
  return Buffer.from("%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF\n").toString("base64");
}

function fiche(overrides = {}) {
  return {
    id: "par5-ab12cd34",
    noProjet: "HQC26-075",
    installation: "Poste Nord",
    emplacement: "Salle des batteries",
    tache: "Remplacement de cellule",
    dateISO: "2026-11-19T15:04:05.000Z",
    pdfBase64: pdfBase64(),
    ...overrides,
  };
}

async function withServer(app, run) {
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  try {
    return await run(base);
  } finally {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
}

test("construit l'adresse Graph avec l'identifiant de site, pas le chemin à deux-points", () => {
  const url = buildContentUrl(configuredEnv, "PAR5-par5-ab12cd34.pdf");
  assert.equal(
    url,
    "https://graph.microsoft.com/v1.0/sites/genitique.sharepoint.com,site-guid,web-guid/drive/root:/PAR5/PAR5-par5-ab12cd34.pdf:/content",
  );
});

test("lit le dossier dans un lien de bibliothèque, et SHAREPOINT_FOLDER gagne", () => {
  const lieu = parseSharePointLocation(
    "https://genitique.sharepoint.com/sites/PowerPlatform-Dev/Documents%20partages/Forms/AllItems.aspx?id=%2Fsites%2FPowerPlatform%2DDev%2FDocuments%20partages%2FCursor",
  );
  assert.equal(lieu.folder, "Cursor");
  const url = buildContentUrl(
    {
      SHAREPOINT_SITE_URL:
        "https://genitique.sharepoint.com/sites/PowerPlatform-Dev/Documents%20partages/Forms/AllItems.aspx?id=%2Fsites%2FPowerPlatform%2DDev%2FDocuments%20partages%2FCursor",
      SHAREPOINT_FOLDER: "PAR5",
    },
    "PAR5-par5-ab12cd34.pdf",
    "genitique.sharepoint.com,site-guid,web-guid",
  );
  assert.match(url, /\/root:\/PAR5\/PAR5-par5-ab12cd34\.pdf:\/content$/);
});

test("refuse un nom de fichier qui n'est pas celui de la fiche", () => {
  assert.throws(() => buildContentUrl(configuredEnv, "../secret.pdf"), /Nom de PDF refusé/);
});

test("la validation ignore le nom de fichier envoyé par le téléphone", () => {
  const parsed = validateFiche(fiche({ pdfName: "../../autre.pdf" }));
  assert.equal(parsed.pdfName, "PAR5-par5-ab12cd34.pdf");
  assert.equal(parsed.pdf.subarray(0, 4).toString("utf8"), "%PDF");
});

test("dépose le PDF sans MySQL et réutilise le même nom", async () => {
  const appels = [];
  const app = createApp({
    env: configuredEnv,
    uploadPdf: async (env, fileName, buffer) => {
      appels.push({ fileName, octets: buffer.length });
      return { webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" };
    },
  });

  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(response.status, 201);
    const body = await response.json();
    assert.equal(body.id, "par5-ab12cd34");
    assert.equal(body.sharePointUrl, "https://genitique.sharepoint.com/fiche");
    assert.equal(appels.length, 1);
    assert.equal(appels[0].fileName, "PAR5-par5-ab12cd34.pdf");
  });
});

test("un MySQL renseigné ne bloque pas le dépôt", async () => {
  const app = createApp({
    env: { ...configuredEnv, MYSQL_HOST: "db.example", MYSQL_USER: "par5_app", MYSQL_DATABASE: "par5" },
    uploadPdf: async () => ({ webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" }),
  });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(response.status, 201);
  });
});

test("répond 503 réessayable tant que SharePoint n'est pas configuré", async () => {
  const app = createApp({ env: {}, uploadPdf: async () => { throw new Error("ne doit pas être appelé"); } });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(response.status, 503);
    const body = await response.json();
    assert.equal(body.retryable, true);
  });
});

test("refuse une fiche sans PDF", async () => {
  const app = createApp({ env: configuredEnv, uploadPdf: async () => ({ webUrl: "", itemId: "" }) });
  await withServer(app, async (base) => {
    const { pdfBase64, ...sansPdf } = fiche();
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(sansPdf),
    });
    assert.equal(response.status, 400);
    assert.equal(pdfBase64.startsWith("JVBE"), true);
  });
});

test("la santé signale SharePoint sans prétendre que MySQL écrit", async () => {
  const app = createApp({ env: configuredEnv });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/health`);
    const body = await response.json();
    assert.equal(body.ok, true);
    assert.equal(body.sharePointConfigured, true);
    assert.equal(body.databaseConfigured, false);
    assert.equal(body.databaseOk, false);
  });
});
