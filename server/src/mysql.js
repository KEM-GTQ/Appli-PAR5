/**
 * Emplacement prévu pour MySQL.
 *
 * L'écriture n'est pas branchée : le PDF part dans SharePoint que MySQL
 * soit là ou non. Quand la base existera, `enregistrerFiche` fera l'upsert
 * et `server/schema.sql` sera le script à approuver avant de le lancer.
 * Le PDF lui-même ne va pas dans la base — seulement l'adresse SharePoint.
 */

export function isDatabaseConfigured(env) {
  return Boolean(env.MYSQL_HOST && env.MYSQL_USER && env.MYSQL_DATABASE);
}

export async function enregistrerFiche() {
  return { ecrit: false };
}
