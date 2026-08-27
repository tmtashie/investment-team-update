const fs = require("fs");
const path = require("path");

function createBackupService({
  BACKUPS_DIR,
  DATA_FILE,
  TASKS_FILE,
  AI_UPDATE_PROPOSALS_FILE,
  COMPANY_DOCUMENTS_FILE,
  DATA_SCHEMA_VERSION,
  ensureDataFile,
  readJsonFile,
  writeJsonFile,
  readMetadata,
  writeMetadata
}) {
  function createBackupSnapshot(reason = "manual") {
    ensureDataFile();
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const backup = {
      schemaVersion: DATA_SCHEMA_VERSION,
      reason,
      createdAt: new Date().toISOString(),
      metadata: readMetadata(),
      investments: readJsonFile(DATA_FILE, []),
      tasks: readJsonFile(TASKS_FILE, []),
      aiUpdateProposals: AI_UPDATE_PROPOSALS_FILE
        ? readJsonFile(AI_UPDATE_PROPOSALS_FILE, [])
        : [],
      companyDocuments: readJsonFile(COMPANY_DOCUMENTS_FILE, [])
    };
    const fileName = `bvb-backup-${timestamp}-${String(reason)
      .replace(/[^a-z0-9_-]+/gi, "-")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 40) || "snapshot"}.json`;
    const filePath = path.join(BACKUPS_DIR, fileName);
    fs.writeFileSync(filePath, JSON.stringify(backup, null, 2) + "\n", "utf8");
    writeMetadata({
      lastBackupAt: backup.createdAt,
      lastBackupReason: reason
    });
    return { fileName, filePath, backup };
  }

  function restoreFromBackupPayload(payload) {
    const investments = Array.isArray(payload.investments) ? payload.investments : [];
    const tasks = Array.isArray(payload.tasks) ? payload.tasks : [];
    const aiUpdateProposals = Array.isArray(payload.aiUpdateProposals)
      ? payload.aiUpdateProposals
      : [];
    const companyDocuments = Array.isArray(payload.companyDocuments) ? payload.companyDocuments : [];
    writeJsonFile(DATA_FILE, investments);
    writeJsonFile(TASKS_FILE, tasks);
    if (AI_UPDATE_PROPOSALS_FILE) {
      writeJsonFile(AI_UPDATE_PROPOSALS_FILE, aiUpdateProposals);
    }
    writeJsonFile(COMPANY_DOCUMENTS_FILE, companyDocuments);
    writeMetadata({
      schemaVersion: DATA_SCHEMA_VERSION,
      lastMigrationAt: new Date().toISOString()
    });
  }

  return {
    createBackupSnapshot,
    restoreFromBackupPayload
  };
}

module.exports = {
  createBackupService
};
