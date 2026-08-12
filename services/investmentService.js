function createInvestmentService({
  DATA_FILE,
  readJsonFile,
  writeJsonFile,
  writeMetadata,
  normalizeInvestment,
  createBackupSnapshot,
  findLatestByPositionKey,
  getInvestmentPositionKey,
  normalizeCompanyKey,
  syncNextStepReminderTasks
}) {
  function writeInvestments(investments) {
    writeJsonFile(DATA_FILE, investments);
  }

  function readInvestments() {
    const parsed = readJsonFile(DATA_FILE, []);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map(normalizeInvestment);
    const changed = JSON.stringify(parsed) !== JSON.stringify(normalized);
    if (changed) {
      writeInvestments(normalized);
      writeMetadata({ lastMigrationAt: new Date().toISOString() });
    }

    return normalized;
  }

  function saveInvestment(entry) {
    const investments = readInvestments();
    createBackupSnapshot("before-investment-create");
    const normalizedEntry = normalizeInvestment({
      ...entry,
      updatedAt: new Date().toISOString()
    });
    const latestMatch = findLatestByPositionKey(
      getInvestmentPositionKey(normalizedEntry),
      investments
    );

    if (latestMatch && latestMatch.company) {
      normalizedEntry.company = latestMatch.company;
    }

    investments.unshift(normalizedEntry);
    writeInvestments(investments);
    syncNextStepReminderTasks(investments);
  }

  function updateInvestment(id, updates) {
    const investments = readInvestments();
    const index = investments.findIndex((investment) => investment.id === id);

    if (index === -1) {
      return null;
    }

    createBackupSnapshot("before-investment-update");
    const merged = normalizeInvestment({
      ...investments[index],
      ...updates,
      id: investments[index].id,
      companyKey: normalizeCompanyKey(updates.company || investments[index].company),
      createdAt: investments[index].createdAt,
      submittedBy: investments[index].submittedBy,
      updatedAt: new Date().toISOString()
    });

    const latestMatch = findLatestByPositionKey(
      getInvestmentPositionKey(merged),
      investments.filter((investment) => investment.id !== id)
    );

    if (latestMatch && latestMatch.company) {
      merged.company = latestMatch.company;
    }

    investments[index] = merged;
    writeInvestments(investments);
    syncNextStepReminderTasks(investments);
    return merged;
  }

  function deleteInvestment(id) {
    const investments = readInvestments();
    const remaining = investments.filter((investment) => investment.id !== id);

    if (remaining.length === investments.length) {
      return false;
    }

    createBackupSnapshot("before-investment-delete");
    writeInvestments(remaining);
    syncNextStepReminderTasks(remaining);
    return true;
  }

  return {
    readInvestments,
    writeInvestments,
    saveInvestment,
    updateInvestment,
    deleteInvestment
  };
}

module.exports = {
  createInvestmentService
};
