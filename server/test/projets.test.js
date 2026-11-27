import assert from "node:assert/strict";
import test from "node:test";
import { createApp } from "../src/app.js";
import { projetsDepuisItems, trouverListe, urlGraphDuSite, viderCacheProjets } from "../src/projets.js";

test("ne garde que les codes projet dont l'état _x00c9_tat est Actif", () => {
  const projets = projetsDepuisItems(
    [
      { fields: { Code_x0020_projet: "HQC26-075", _x00c9_tat: "Actif" } },
      { fields: { Code_x0020_projet: " hqc26-075 ", _x00c9_tat: "Actif" } },
      { fields: { Code_x0020_projet: "EEQ22-039", _x00c9_tat: "Fermé" } },
      { fields: { Code_x0020_projet: "LHI21-123", _x00c9_tat: { Label: "actif" } } },
      { fields: { Code_x0020_projet: "   ", _x00c9_tat: "Actif" } },
    ],
    { nom: "Code_x0020_projet", etat: "_x00c9_tat", actif: "Actif" },
  );
  assert.deepEqual(projets, [
    { numero: "HQC26-075", phase: "" },
    { numero: "LHI21-123", phase: "" },
  ]);
});

test("retrouve la liste par son nom interne ou son adresse, pas seulement par son titre", () => {
  const listes = [
    { id: "autre", displayName: "Fiches projets", name: "FichesProjets", webUrl: "https://genitique.sharepoint.com/Lists/Fiches%20projets" },
    {
      id: "bonne",
      displayName: "Liste des projets",
      name: "Projets 2",
      webUrl: "https://genitique.sharepoint.com/Lists/Projets%202",
    },
  ];
  assert.equal(trouverListe(listes, "Projets 2")?.id, "bonne");
  assert.equal(trouverListe(listes, "Liste des projets")?.id, "bonne");
  assert.equal(trouverListe(listes, "Absente"), null);
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
