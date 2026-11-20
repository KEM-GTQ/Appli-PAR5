import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";
import { createApp } from "./app.js";
import { createPool, ensureSchema, isDatabaseConfigured } from "./mysql.js";
import { isSharePointConfigured } from "./sharepoint.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
dotenv.config({ path: path.join(repoRoot, ".env") });

const port = Number(process.env.PORT || 3001);
const pool = isDatabaseConfigured(process.env) ? createPool(process.env) : null;

if (pool) {
  try {
    await ensureSchema(pool);
    console.log("MySQL connecté, table par5_fiches prête.");
  } catch (error) {
    console.error(`MySQL injoignable : ${error.message}`);
  }
} else {
  console.log("MySQL non configuré.");
}
console.log(`SharePoint ${isSharePointConfigured(process.env) ? "configuré" : "non configuré"}.`);

const server = createApp({ env: process.env, pool }).listen(port, () => {
  console.log(`API à l'écoute sur http://localhost:${port}`);
});

server.on("error", (error) => {
  console.error(error.message);
  process.exitCode = 1;
});
