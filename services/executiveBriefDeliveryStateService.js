function createExecutiveBriefDeliveryStateService({ STATE_FILE, readJsonFile, writeJsonFile }) {
  function readState() {
    const value = readJsonFile(STATE_FILE, []);
    return Array.isArray(value) ? value : [];
  }

  function writeState(entries) {
    writeJsonFile(STATE_FILE, entries);
  }

  function reserve(entry) {
    const entries = readState();
    if (entries.some((item) => item.idempotencyKey === entry.idempotencyKey)) {
      return { reserved: false, reason: "duplicate" };
    }
    if (entries.filter((item) => item.briefingDate === entry.briefingDate).length >= 2) {
      return { reserved: false, reason: "daily-limit" };
    }
    entries.unshift({ ...entry, status: "pending" });
    writeState(entries);
    return { reserved: true };
  }

  function update(idempotencyKey, patch) {
    const entries = readState();
    const index = entries.findIndex((item) => item.idempotencyKey === idempotencyKey);
    if (index === -1) {
      return null;
    }
    entries[index] = { ...entries[index], ...patch };
    writeState(entries);
    return entries[index];
  }

  return {
    markFailedClosed: (key, failedAt) => update(key, { status: "failed-closed", failedAt }),
    markSent: (key, sentAt, providerMessageId) => update(key, {
      status: "sent",
      sentAt,
      providerMessageId: String(providerMessageId || "").slice(0, 200)
    }),
    readState,
    reserve
  };
}

module.exports = { createExecutiveBriefDeliveryStateService };
