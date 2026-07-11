(function attachImportBackup(root, factory) {
  const api = factory(root.TeamsTypingHelperConfig, root.TeamsTypingHelperExportData);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperImportBackup = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createImportBackupApi(configApi, exportApi) {
  "use strict";

  function createImportBackup(config, date = new Date()) {
    const normalized = configApi.normalizeConfig(config);
    return {
      createdAt: date.toISOString(),
      phrases: normalized.phrases,
      settings: exportApi.configToSettings(normalized),
      config: normalized
    };
  }

  function storageGet(storageArea, keys) {
    return new Promise((resolve) => storageArea.get(keys, (result) => resolve(result || {})));
  }

  function storageSet(storageArea, values) {
    return new Promise((resolve, reject) => {
      storageArea.set(values, () => {
        const error = chrome.runtime && chrome.runtime.lastError;
        if (error) {
          reject(new Error(error.message));
          return;
        }
        resolve();
      });
    });
  }

  async function createImportBackupInStorage(storageArea, config) {
    const backup = createImportBackup(config);
    await storageSet(storageArea, { [configApi.IMPORT_BACKUP_KEY]: backup });
    return backup;
  }

  async function readImportBackup(storageArea) {
    const result = await storageGet(storageArea, configApi.IMPORT_BACKUP_KEY);
    return result[configApi.IMPORT_BACKUP_KEY] || null;
  }

  async function restoreImportBackup(storageArea) {
    const backup = await readImportBackup(storageArea);
    if (!backup || !backup.config) {
      throw new Error("当前没有可恢复的导入备份。");
    }
    await storageSet(storageArea, { [configApi.STORAGE_KEY]: configApi.normalizeConfig(backup.config) });
    return configApi.normalizeConfig(backup.config);
  }

  return {
    createImportBackup,
    createImportBackupInStorage,
    readImportBackup,
    restoreImportBackup
  };
});
