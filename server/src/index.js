import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";
import { createApp } from "./app.js";
import { isDatabaseConfigured } from "./mysql.js";
import { isSharePointConfigured } from "./sharepoint.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
dotenv.config({ path: path.join(repoRoot, ".env") });

const port = Number(process.env.PORT || 3001);
console.log(
  `MySQL ${isDatabaseConfigured(process.env) ? "renseigné, écriture inactive" : "non configuré"}; SharePoint ${isSharePointConfigured(process.env) ? "configuré" : "non configuré"}`,
);

const server = createApp({ env: process.env }).listen(port, () => {
  console.log(`API à l'écoute sur http://localhost:${port}`);
});

server.on("error", (error) => {
  console.error(error.message);
  process.exitCode = 1;
});
