import mysql from "mysql2/promise";

/**
 * Écriture des PAR 5 dans MySQL.
 *
 * Le PDF reste dans SharePoint. La ligne garde l'en-tête, les 16 réponses
 * (sans les photos, qui sont dans le PDF) et l'adresse du fichier.
 * Le même id met à jour la ligne au lieu d'en créer une seconde.
 */

export function isDatabaseConfigured(env) {
  return Boolean(env.MYSQL_HOST && env.MYSQL_USER && env.MYSQL_DATABASE);
}

export function createPool(env) {
  return mysql.createPool({
    host: env.MYSQL_HOST,
    port: Number(env.MYSQL_PORT || 3306),
    user: env.MYSQL_USER,
    password: env.MYSQL_PASSWORD || "",
    database: env.MYSQL_DATABASE,
    waitForConnections: true,
    connectionLimit: 4,
    enableKeepAlive: true,
  });
}

const CREER_TABLE = `
CREATE TABLE IF NOT EXISTS par5_fiches (
  id VARCHAR(16) NOT NULL,
  no_projet VARCHAR(255) NOT NULL,
  installation VARCHAR(200) NOT NULL,
  emplacement VARCHAR(200) NOT NULL,
  tache VARCHAR(300) NOT NULL,
  date_analyse DATETIME(3) NOT NULL,
  maj DATETIME(3) NULL,
  depose_le DATETIME(3) NOT NULL,
  nom_pdf VARCHAR(255) NOT NULL,
  sharepoint_url VARCHAR(2048) NULL,
  sharepoint_item_id VARCHAR(128) NULL,
  lignes JSON NOT NULL,
  travailleurs JSON NOT NULL,
  auteur VARCHAR(320) NULL,
  PRIMARY KEY (id)
)`;

async function aColonne(pool, nom) {
  const [rows] = await pool.query(
    `SELECT 1 AS ok FROM information_schema.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'par5_fiches' AND COLUMN_NAME = ?
     LIMIT 1`,
    [nom],
  );
  return rows.length > 0;
}

export async function ensureSchema(pool) {
  await pool.query(CREER_TABLE);
  if (!(await aColonne(pool, "lignes"))) {
    await pool.query("ALTER TABLE par5_fiches ADD COLUMN lignes JSON NOT NULL");
  }
  if (!(await aColonne(pool, "travailleurs"))) {
    await pool.query("ALTER TABLE par5_fiches ADD COLUMN travailleurs JSON NOT NULL");
  }
  if (!(await aColonne(pool, "auteur"))) {
    await pool.query("ALTER TABLE par5_fiches ADD COLUMN auteur VARCHAR(320) NULL");
  }
  await pool.query("ALTER TABLE par5_fiches MODIFY no_projet VARCHAR(255) NOT NULL");
}

function versMysql(iso) {
  return new Date(iso).toISOString().slice(0, 23).replace("T", " ");
}

export async function enregistrerFiche(pool, fiche, uploaded) {
  if (!pool) throw new Error("MySQL n'est pas connecté.");
  const lignes = (fiche.lignes || []).map((l) => ({
    risqueId: l.risqueId,
    present: l.present,
    specification: l.specification,
    moyens: l.moyens,
    responsable: l.responsable,
    realise: Boolean(l.realise),
  }));
  await pool.query(
    `INSERT INTO par5_fiches (
       id, no_projet, installation, emplacement, tache,
       date_analyse, maj, depose_le, nom_pdf,
       sharepoint_url, sharepoint_item_id, lignes, travailleurs, auteur
     ) VALUES (?, ?, ?, ?, ?, ?, ?, UTC_TIMESTAMP(3), ?, ?, ?, ?, ?, ?)
     ON DUPLICATE KEY UPDATE
       no_projet = VALUES(no_projet),
       installation = VALUES(installation),
       emplacement = VALUES(emplacement),
       tache = VALUES(tache),
       date_analyse = VALUES(date_analyse),
       maj = VALUES(maj),
       depose_le = VALUES(depose_le),
       nom_pdf = VALUES(nom_pdf),
       sharepoint_url = COALESCE(VALUES(sharepoint_url), sharepoint_url),
       sharepoint_item_id = COALESCE(VALUES(sharepoint_item_id), sharepoint_item_id),
       lignes = VALUES(lignes),
       travailleurs = VALUES(travailleurs),
       auteur = COALESCE(VALUES(auteur), auteur)`,
    [
      fiche.id,
      fiche.noProjet,
      fiche.installation,
      fiche.emplacement,
      fiche.tache,
      versMysql(fiche.dateISO),
      fiche.majISO ? versMysql(fiche.majISO) : null,
      fiche.pdfName,
      uploaded?.webUrl || null,
      uploaded?.itemId || null,
      JSON.stringify(lignes),
      JSON.stringify(fiche.travailleurs || []),
      fiche.auteur || null,
    ],
  );
  return { ecrit: true };
}
