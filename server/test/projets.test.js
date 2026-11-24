import assert from "node:assert/strict";
import test from "node:test";
import { createApp } from "../src/app.js";
import { projetsDepuisItems, urlGraphDuSite, viderCacheProjets } from "../src/projets.js";

test("ne garde que les projets dont l'état est Actif", () => {
  const projets = projetsDepuisItems(
    [
      { fields: { Projet: "Poste Nord", Etat: "Actif" } },
      { fields: { Projet: " poste nord ", Etat: "Actif" } },
      { fields: { Projet: "Ancien chantier", Etat: "Fermé" } },
      { fields: { Projet: "Ligne B", Etat: { Label: "actif" } } },
      { fields: { Projet: "   ", Etat: "Actif" } },
    ],
    { nom: "Projet", etat: "Etat", actif: "Actif" },
  );
  assert.deepEqual(projets, [
    { numero: "Ligne B", phase: "" },
    { numero: "Poste Nord", phase: "" },
  ]);
});

test("la liste Projets 2 est lue sur le site racine, pas sur PowerPlatform-Dev", () => {
  assert.equal(
    urlGraphDuSite("https://genitique.sharepoint.com/Lists/Projets%202/AllItems.aspx?origin=createList"),
    "https://graph.microsoft.com/v1.0/sites/genitique.sharepoint.com",
  );
  assert.equal(
    urlGraphDuSite(
      "https://genitique.sharepoint.com/sites/PowerPlatform-Dev/Documents%20partages/Forms/AllItems.aspx",
    ),
    "https://graph.microsoft.com/v1.0/sites/genitique.sharepoint.com:/sites/PowerPlatform-Dev",
  );
});

test("le menu des projets ne lit pas le site où sont déposés les PDF", async () => {
  viderCacheProjets();
  const app = createApp({
    env: {
      GRAPH_TENANT_ID: "tenant-id",
      GRAPH_CLIENT_ID: "client-id",
      GRAPH_CLIENT_SECRET: "client-secret",
      SHAREPOINT_SITE_ID: "genitique.sharepoint.com,site-guid,web-guid",
    },
  });
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  try {
    const response = await fetch(`http://127.0.0.1:${server.address().port}/api/projets`);
    const body = await response.json();
    assert.equal(response.status, 503);
    assert.match(body.error, /liste des projets/);
  } finally {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
});
