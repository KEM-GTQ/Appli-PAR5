import assert from "node:assert/strict";
import test from "node:test";
import { createApp } from "../src/app.js";
import { enregistrerFiche } from "../src/mysql.js";
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

test("l'upsert MySQL réutilise l'id et n'enregistre pas le PDF", async () => {
  const requetes = [];
  const pool = {
    async query(sql, params) {
      requetes.push({ sql, params });
      return [[]];
    },
  };
  await enregistrerFiche(
    pool,
    {
      id: "par5-ab12cd34",
      noProjet: "HQC26-075",
      installation: "Poste",
      emplacement: "Salle",
      tache: "Cellule",
      dateISO: "2026-11-20T15:00:00.000Z",
      majISO: null,
      pdfName: "PAR5-par5-ab12cd34.pdf",
      lignes: [{ risqueId: "chute", present: "oui", specification: "", moyens: "Harnais", responsable: "AM", realise: true, photo: "data:image/jpeg;ne-pas-garder" }],
      travailleurs: [{ nom: "Alex Martin", initiales: "AM", date: "2026-11-20" }],
    },
    { webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" },
  );
  assert.match(requetes[0].sql, /ON DUPLICATE KEY UPDATE/);
  assert.match(requetes[0].sql, /COALESCE\(VALUES\(sharepoint_url\), sharepoint_url\)/);
  assert.equal(requetes[0].params[0], "par5-ab12cd34");
  assert.equal(requetes[0].params[10].includes("ne-pas-garder"), false);
  assert.equal(requetes[0].params[10].includes("Harnais"), true);
});

test("écrit la fiche dans MySQL avant et après le dépôt SharePoint", async () => {
  const etapes = [];
  const app = createApp({
    env: { ...configuredEnv, MYSQL_HOST: "localhost", MYSQL_USER: "root", MYSQL_DATABASE: "par5" },
    uploadPdf: async () => {
      etapes.push("pdf");
      return { webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" };
    },
    enregistrerFiche: async (recue, fichier) => {
      etapes.push(fichier?.webUrl || "base");
      assert.equal(recue.lignes.length, 1);
      return { ecrit: true };
    },
  });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(
        fiche({
          lignes: [{ risqueId: "chute", present: "oui", moyens: "Harnais", responsable: "AM", realise: true }],
          travailleurs: [{ nom: "Alex Martin", initiales: "AM", date: "2026-11-20" }],
        }),
      ),
    });
    assert.equal(response.status, 201);
    assert.deepEqual(etapes, ["base", "pdf", "https://genitique.sharepoint.com/fiche"]);
  });
});

test("un échec MySQL laisse la fiche réessayable sans déposer le PDF", async () => {
  let depots = 0;
  const app = createApp({
    env: { ...configuredEnv, MYSQL_HOST: "localhost", MYSQL_USER: "root", MYSQL_DATABASE: "par5" },
    uploadPdf: async () => {
      depots += 1;
      return { webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" };
    },
    enregistrerFiche: async () => {
      throw new Error("connexion refusée");
    },
  });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(response.status, 502);
    const body = await response.json();
    assert.equal(body.retryable, true);
    assert.equal(depots, 0);
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

test("refuse le dépôt tant que le jeton Entra manque", async () => {
  let depots = 0;
  const app = createApp({
    env: { ...configuredEnv, ENTRA_TENANT_ID: "tenant", ENTRA_API_AUDIENCE: "api://par5" },
    verifierJeton: async () => {
      throw new Error("signature");
    },
    uploadPdf: async () => {
      depots += 1;
      return { webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" };
    },
  });
  await withServer(app, async (base) => {
    const sans = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(sans.status, 401);
    const corps = await sans.json();
    assert.equal(corps.code, "auth");
    assert.equal(corps.retryable, true);
    const mauvais = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer mauvais" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(mauvais.status, 401);
    assert.equal(depots, 0);
  });
});

test("retient le compte Entra qui dépose la fiche", async () => {
  const auteurs = [];
  const app = createApp({
    env: {
      ...configuredEnv,
      ENTRA_TENANT_ID: "tenant",
      ENTRA_API_AUDIENCE: "api://par5",
      MYSQL_HOST: "localhost",
      MYSQL_USER: "root",
      MYSQL_DATABASE: "par5",
    },
    verifierJeton: async (token) => {
      assert.equal(token, "jeton-alex");
      return { preferred_username: "alex@genitique.com", name: "Alex Martin", oid: "oid-1" };
    },
    uploadPdf: async () => ({ webUrl: "https://genitique.sharepoint.com/fiche", itemId: "item-1" }),
    enregistrerFiche: async (recue) => {
      auteurs.push(recue.auteur);
      return { ecrit: true };
    },
  });
  await withServer(app, async (base) => {
    const response = await fetch(`${base}/api/par5`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer jeton-alex" },
      body: JSON.stringify(fiche()),
    });
    assert.equal(response.status, 201);
    assert.deepEqual(auteurs, ["alex@genitique.com", "alex@genitique.com"]);
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
    assert.equal(body.authRequired, false);
  });
});
