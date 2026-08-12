function createTaskService({
  TASKS_FILE,
  readJsonFile,
  writeJsonFile,
  writeMetadata,
  normalizeTask,
  createBackupSnapshot
}) {
  function writeTasks(tasks) {
    writeJsonFile(TASKS_FILE, tasks);
  }

  function readTasks() {
    const parsed = readJsonFile(TASKS_FILE, []);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map(normalizeTask);
    const changed = JSON.stringify(parsed) !== JSON.stringify(normalized);
    if (changed) {
      writeTasks(normalized);
      writeMetadata({ lastMigrationAt: new Date().toISOString() });
    }

    return normalized;
  }

  function saveTask(entry) {
    const tasks = readTasks();
    createBackupSnapshot("before-task-create");
    const normalized = normalizeTask({
      ...entry,
      updatedAt: new Date().toISOString()
    });
    tasks.unshift(normalized);
    writeTasks(tasks);
    return normalized;
  }

  function updateTask(id, updates) {
    const tasks = readTasks();
    const index = tasks.findIndex((task) => task.id === id);

    if (index === -1) {
      return null;
    }

    createBackupSnapshot("before-task-update");
    const merged = normalizeTask({
      ...tasks[index],
      ...updates,
      id: tasks[index].id,
      createdAt: tasks[index].createdAt,
      createdBy: tasks[index].createdBy,
      completedAt:
        updates.status === "Completed" && !tasks[index].completedAt
          ? new Date().toISOString()
          : updates.status && updates.status !== "Completed"
            ? ""
            : tasks[index].completedAt,
      updatedAt: new Date().toISOString()
    });

    tasks[index] = merged;
    writeTasks(tasks);
    return merged;
  }

  function deleteTask(id) {
    const tasks = readTasks();
    const remaining = tasks.filter((task) => task.id !== id);

    if (remaining.length === tasks.length) {
      return false;
    }

    createBackupSnapshot("before-task-delete");
    writeTasks(remaining);
    return true;
  }

  return {
    readTasks,
    writeTasks,
    saveTask,
    updateTask,
    deleteTask
  };
}

module.exports = {
  createTaskService
};
