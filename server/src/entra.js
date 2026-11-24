import { createRemoteJWKSet, jwtVerify } from "jose";

/**
 * Vérifie le jeton de l'utilisateur. Le téléphone n'obtient jamais un jeton
 * Graph : seulement un jeton pour cette API. Le dépôt SharePoint reste
 * un identifiant d'application, côté serveur.
 */

const clefs = new Map();

export function isEntraConfigured(env) {
  return Boolean(env.ENTRA_TENANT_ID && env.ENTRA_API_AUDIENCE);
}

export function verifierEntra(env) {
  const tenant = env.ENTRA_TENANT_ID;
  const audience = String(env.ENTRA_API_AUDIENCE)
    .split(",")
    .map((valeur) => valeur.trim())
    .filter(Boolean);
  const issuer = `https://login.microsoftonline.com/${tenant}/v2.0`;
  let jwks = clefs.get(tenant);
  if (!jwks) {
    jwks = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys`),
    );
    clefs.set(tenant, jwks);
  }

  return async function verifier(token) {
    const { payload } = await jwtVerify(token, jwks, {
      issuer,
      audience,
      algorithms: ["RS256"],
    });
    return {
      oid: typeof payload.oid === "string" ? payload.oid : "",
      name: typeof payload.name === "string" ? payload.name : "",
      preferred_username:
        typeof payload.preferred_username === "string" ? payload.preferred_username : "",
    };
  };
}
