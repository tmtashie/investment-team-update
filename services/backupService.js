const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const MAX_BACKUP_SNAPSHOTS = 50;
const BACKUP_FILE_PATTERN = /^bvb-backup-.+\.json$/;

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
  function buildBackupPayload(reason = "manual") {
    return {
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
  }

  function buildBackupFileName(backup) {
    const timestamp = String(backup.createdAt || new Date().toISOString()).replace(/[:.]/g, "-");
    return `bvb-backup-${timestamp}-${String(backup.reason || "snapshot")
      .replace(/[^a-z0-9_-]+/gi, "-")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 40) || "snapshot"}-${crypto.randomUUID().slice(0, 8)}.json`;
  }

  function readValidBackupSnapshot(filePath) {
    try {
      const parsed = JSON.parse(fs.readFileSync(filePath, "utf8"));
      const isValid = Boolean(
        parsed &&
          typeof parsed === "object" &&
          parsed.schemaVersion &&
          parsed.createdAt &&
          Array.isArray(parsed.investments) &&
          Array.isArray(parsed.tasks) &&
          Array.isArray(parsed.companyDocuments)
      );
      if (!isValid) {
        return null;
      }
      const createdAtMs = new Date(parsed.createdAt).getTime();
      if (!Number.isFinite(createdAtMs)) {
        return null;
      }
      return { parsed, createdAtMs };
    } catch (error) {
      return null;
    }
  }

  function pruneOldBackupSnapshots(protectedFileName) {
    let entries;
    try {
      entries = fs.readdirSync(BACKUPS_DIR, { withFileTypes: true });
    } catch (error) {
      console.warn("[backup-retention] Could not read backups directory:", error.message);
      return;
    }

    const validBackups = entries
      .filter((entry) => entry.isFile() && BACKUP_FILE_PATTERN.test(entry.name))
      .map((entry) => {
        const filePath = path.join(BACKUPS_DIR, entry.name);
        const snapshot = readValidBackupSnapshot(filePath);
        if (!snapshot) {
          return null;
        }
        return {
          name: entry.name,
          filePath,
          createdAtMs: snapshot.createdAtMs
        };
      })
      .filter(Boolean)
      .sort((left, right) => right.createdAtMs - left.createdAtMs);

    const staleBackups = validBackups
      .filter((backup) => backup.name !== protectedFileName)
      .slice(Math.max(0, MAX_BACKUP_SNAPSHOTS - 1));

    staleBackups.forEach((backup) => {
      try {
        fs.unlinkSync(backup.filePath);
        console.log("[backup-retention] Pruned old backup snapshot:", backup.name);
      } catch (error) {
        console.warn("[backup-retention] Could not prune backup snapshot:", backup.name, error.message);
      }
    });
  }

  function writeBackupSnapshotAtomically(filePath, backup) {
    const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
    try {
      fs.writeFileSync(tempPath, JSON.stringify(backup, null, 2) + "\n", "utf8");
      fs.renameSync(tempPath, filePath);
    } catch (error) {
      try {
        if (fs.existsSync(tempPath)) {
          fs.unlinkSync(tempPath);
        }
      } catch (cleanupError) {
        console.warn("[backup] Could not remove incomplete backup temp file:", cleanupError.message);
      }
      throw error;
    }
  }

  function createBackupSnapshot(reason = "manual") {
    ensureDataFile();
    const backup = buildBackupPayload(reason);
    const fileName = buildBackupFileName(backup);
    const filePath = path.join(BACKUPS_DIR, fileName);
    writeBackupSnapshotAtomically(filePath, backup);
    try {
      pruneOldBackupSnapshots(fileName);
    } catch (error) {
      console.warn("[backup-retention] Backup pruning failed:", error.message);
    }
    writeMetadata({
      lastBackupAt: backup.createdAt,
      lastBackupReason: reason
    });
    return { fileName, filePath, backup };
  }

  function createBackupExportPayload(reason = "manual-export") {
    ensureDataFile();
    const backup = buildBackupPayload(reason);
    const fileName = buildBackupFileName(backup);
    return { fileName, backup };
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
    createBackupExportPayload,
    restoreFromBackupPayload
  };
}

module.exports = {
  createBackupService
};
