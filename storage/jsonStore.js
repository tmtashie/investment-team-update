const fs = require("fs");

function createJsonStore({
  DATA_DIR,
  DATA_FILE,
  TASKS_FILE,
  COMPANY_DOCUMENTS_FILE,
  METADATA_FILE,
  BACKUPS_DIR,
  UPLOADS_DIR,
  DATA_SCHEMA_VERSION
}) {
  function ensureDataFile() {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }

    if (!fs.existsSync(UPLOADS_DIR)) {
      fs.mkdirSync(UPLOADS_DIR, { recursive: true });
    }

    if (!fs.existsSync(BACKUPS_DIR)) {
      fs.mkdirSync(BACKUPS_DIR, { recursive: true });
    }

    if (!fs.existsSync(DATA_FILE)) {
      fs.writeFileSync(DATA_FILE, "[]\n", "utf8");
    }

    if (!fs.existsSync(TASKS_FILE)) {
      fs.writeFileSync(TASKS_FILE, "[]\n", "utf8");
    }

    if (!fs.existsSync(COMPANY_DOCUMENTS_FILE)) {
      fs.writeFileSync(COMPANY_DOCUMENTS_FILE, "[]\n", "utf8");
    }

    if (!fs.existsSync(METADATA_FILE)) {
      fs.writeFileSync(
        METADATA_FILE,
        JSON.stringify(
          {
            schemaVersion: DATA_SCHEMA_VERSION,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          },
          null,
          2
        ) + "\n",
        "utf8"
      );
    }
  }

  function readJsonFile(filePath, fallback) {
    ensureDataFile();
    try {
      const raw = fs.readFileSync(filePath, "utf8");
      return JSON.parse(raw);
    } catch (error) {
      return fallback;
    }
  }

  function writeJsonFile(filePath, value) {
    ensureDataFile();
    fs.writeFileSync(filePath, JSON.stringify(value, null, 2) + "\n", "utf8");
  }

  return {
    ensureDataFile,
    readJsonFile,
    writeJsonFile
  };
}

module.exports = {
  createJsonStore
};
