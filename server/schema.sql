-- Même table que l'API crée au démarrage quand MySQL est configuré.
-- Le PDF n'est pas stocké ici. La ligne garde l'en-tête, les réponses et l'adresse SharePoint.

CREATE TABLE IF NOT EXISTS par5_fiches (
  id VARCHAR(16) NOT NULL,
  no_projet VARCHAR(80) NOT NULL,
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
  PRIMARY KEY (id)
);

-- Compte dédié, mot de passe hors du dépôt git.
-- SELECT, INSERT, UPDATE sur par5_fiches seulement. Pas de DELETE, pas de DDL.
--
-- CREATE USER 'par5_app'@'%' IDENTIFIED BY '<défini hors du dépôt>';
-- GRANT SELECT, INSERT, UPDATE ON par5_fiches TO 'par5_app'@'%';
