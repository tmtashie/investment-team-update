"use strict";

const { DatabaseSync, constants } = require("node:sqlite");

const REQUIRED_COLUMNS = Object.freeze({
  chat: ["guid"],
  chat_handle_join: ["chat_id", "handle_id"],
  chat_message_join: ["chat_id", "message_id"],
  handle: ["id"],
  message: [
    "guid",
    "handle_id",
    "is_from_me",
    "date",
    "text",
    "associated_message_type",
    "item_type",
    "is_empty",
    "date_edited",
    "date_retracted",
    "thread_originator_guid"
  ]
});

const WRITE_ACTIONS = new Set([
  constants.SQLITE_ALTER_TABLE,
  constants.SQLITE_ANALYZE,
  constants.SQLITE_ATTACH,
  constants.SQLITE_CREATE_INDEX,
  constants.SQLITE_CREATE_TABLE,
  constants.SQLITE_CREATE_TEMP_INDEX,
  constants.SQLITE_CREATE_TEMP_TABLE,
  constants.SQLITE_CREATE_TEMP_TRIGGER,
  constants.SQLITE_CREATE_TEMP_VIEW,
  constants.SQLITE_CREATE_TRIGGER,
  constants.SQLITE_CREATE_VIEW,
  constants.SQLITE_CREATE_VTABLE,
  constants.SQLITE_DELETE,
  constants.SQLITE_DETACH,
  constants.SQLITE_DROP_INDEX,
  constants.SQLITE_DROP_TABLE,
  constants.SQLITE_DROP_TEMP_INDEX,
  constants.SQLITE_DROP_TEMP_TABLE,
  constants.SQLITE_DROP_TEMP_TRIGGER,
  constants.SQLITE_DROP_TEMP_VIEW,
  constants.SQLITE_DROP_TRIGGER,
  constants.SQLITE_DROP_VIEW,
  constants.SQLITE_DROP_VTABLE,
  constants.SQLITE_INSERT,
  constants.SQLITE_REINDEX,
  constants.SQLITE_SAVEPOINT,
  constants.SQLITE_TRANSACTION,
  constants.SQLITE_UPDATE
]);

function databaseError() {
  const error = new Error("The Messages database is unavailable or incompatible.");
  error.code = "MESSAGES_DATABASE_UNAVAILABLE";
  return error;
}

function validateSchema(database) {
  for (const [table, requiredColumns] of Object.entries(REQUIRED_COLUMNS)) {
    let columns;
    try {
      columns = database.prepare(`PRAGMA table_info(${table})`).all().map((column) => column.name);
    } catch {
      throw databaseError();
    }
    if (requiredColumns.some((column) => !columns.includes(column))) throw databaseError();
  }
}

function openReadOnlyMessagesDatabase(databasePath) {
  let database;
  try {
    database = new DatabaseSync(databasePath, {
      open: true,
      readOnly: true,
      allowExtension: false,
      timeout: 1000
    });
    database.enableDefensive(true);
    database.exec("PRAGMA query_only = ON");
    validateSchema(database);
    database.setAuthorizer((actionCode, argumentOne, argumentTwo) => {
      if (WRITE_ACTIONS.has(actionCode)) return constants.SQLITE_DENY;
      if (actionCode === constants.SQLITE_PRAGMA && !(argumentOne === "query_only" && argumentTwo === null)) {
        return constants.SQLITE_DENY;
      }
      return constants.SQLITE_OK;
    });
    return database;
  } catch {
    if (database) {
      try { database.close(); } catch {}
    }
    throw databaseError();
  }
}

module.exports = {
  openReadOnlyMessagesDatabase,
  validateSchema
};
