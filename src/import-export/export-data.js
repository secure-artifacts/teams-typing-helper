(function attachExportData(root, factory) {
  const api = factory(root.TeamsTypingHelperConfig);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperExportData = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createExportDataApi(configApi) {
  "use strict";

  const CURRENT_FORMAT_VERSION = 1;

  function pad(value) {
    return String(value).padStart(2, "0");
  }

  function formatFilenameDate(date) {
    return [
      date.getFullYear(),
      pad(date.getMonth() + 1),
      pad(date.getDate())
    ].join("-") + "-" + [
      pad(date.getHours()),
      pad(date.getMinutes()),
      pad(date.getSeconds())
    ].join("");
  }

  function isValidIsoDate(value) {
    return typeof value === "string" && !Number.isNaN(Date.parse(value));
  }

  function buildExportFilename(date = new Date()) {
    return `teams-typing-helper-backup-${formatFilenameDate(date)}.json`;
  }

  function configToSettings(config) {
    const normalized = configApi.normalizeConfig(config);
    return {
      toolbarEnabled: normalized.enabled,
      compactMode: normalized.compactMode,
      buttonsPerRow: normalized.columns,
      insertMode: normalized.insertMode,
      toolbarPosition: normalized.toolbar.position,
      toolbarCollapsed: normalized.toolbar.collapsed
    };
  }

  function buildExportData(config, date = new Date()) {
    const normalized = configApi.normalizeConfig(config);
    const exportedAt = date.toISOString();
    return {
      app: configApi.APP_NAME,
      formatVersion: CURRENT_FORMAT_VERSION,
      exportedAt,
      phrases: normalized.phrases.map((phrase, index) => ({
        id: phrase.id,
        name: phrase.name,
        text: phrase.text,
        order: index,
        enabled: phrase.enabled !== false,
        createdAt: isValidIsoDate(phrase.createdAt) ? phrase.createdAt : exportedAt,
        updatedAt: isValidIsoDate(phrase.updatedAt) ? phrase.updatedAt : exportedAt
      })),
      settings: configToSettings(normalized)
    };
  }

  function downloadExportFile(exportData, deps = {}) {
    const doc = deps.document || document;
    const urlApi = deps.URL || URL;
    const BlobCtor = deps.Blob || Blob;
    const date = exportData && exportData.exportedAt ? new Date(exportData.exportedAt) : new Date();
    const filename = buildExportFilename(date);
    const json = JSON.stringify(exportData, null, 2);
    const blob = new BlobCtor([json], { type: "application/json;charset=utf-8" });
    const downloadUrl = urlApi.createObjectURL(blob);
    try {
      const link = doc.createElement("a");
      link.href = downloadUrl;
      link.download = filename;
      link.style.display = "none";
      (doc.body || doc.documentElement).appendChild(link);
      link.click();
      link.remove();
    } finally {
      urlApi.revokeObjectURL(downloadUrl);
    }
    return { filename, json };
  }

  return {
    CURRENT_FORMAT_VERSION,
    buildExportData,
    buildExportFilename,
    configToSettings,
    downloadExportFile
  };
});
