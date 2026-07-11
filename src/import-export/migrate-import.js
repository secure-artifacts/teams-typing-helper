(function attachMigrateImport(root, factory) {
  const api = factory(root.TeamsTypingHelperConfig);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperMigrateImport = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createMigrateImportApi(configApi) {
  "use strict";

  const CURRENT_FORMAT_VERSION = 1;

  function isPlainObject(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return false;
    }
    const proto = Object.getPrototypeOf(value);
    return proto === Object.prototype || proto === null;
  }

  function migrateImportData(rawData) {
    if (!isPlainObject(rawData)) {
      throw new Error("备份文件根数据必须是对象。");
    }

    if (rawData.app !== undefined && rawData.app !== configApi.APP_NAME) {
      throw new Error("这不是 Teams 辅助打字插件的备份文件。");
    }

    if (rawData.formatVersion === undefined) {
      if (Array.isArray(rawData.phrases)) {
        return {
          legacy: true,
          legacyMessage: "检测到旧版备份文件，导入时会自动升级格式。",
          formatVersion: CURRENT_FORMAT_VERSION,
          exportedAt: typeof rawData.exportedAt === "string" ? rawData.exportedAt : null,
          phrases: rawData.phrases,
          settings: isPlainObject(rawData.settings) ? rawData.settings : null
        };
      }

      if (isPlainObject(rawData.config) && Array.isArray(rawData.config.phrases)) {
        return {
          legacy: true,
          legacyMessage: "检测到旧版备份文件，导入时会自动升级格式。",
          formatVersion: CURRENT_FORMAT_VERSION,
          exportedAt: typeof rawData.exportedAt === "string" ? rawData.exportedAt : null,
          phrases: rawData.config.phrases,
          settings: {
            toolbarEnabled: rawData.config.enabled,
            compactMode: rawData.config.compactMode,
            buttonsPerRow: rawData.config.columns,
            insertMode: rawData.config.insertMode,
            toolbarPosition: rawData.config.toolbar && rawData.config.toolbar.position,
            toolbarCollapsed: rawData.config.toolbar && rawData.config.toolbar.collapsed
          }
        };
      }

      throw new Error("备份文件缺少 formatVersion。");
    }

    if (typeof rawData.formatVersion !== "number") {
      throw new Error("备份文件 formatVersion 必须是数字。");
    }
    if (rawData.formatVersion > CURRENT_FORMAT_VERSION) {
      throw new Error("该备份文件来自更高版本的插件，当前版本无法安全导入。");
    }
    if (!Array.isArray(rawData.phrases)) {
      throw new Error("备份文件缺少 phrases 数组。");
    }

    return {
      legacy: false,
      formatVersion: rawData.formatVersion,
      exportedAt: typeof rawData.exportedAt === "string" ? rawData.exportedAt : null,
      phrases: rawData.phrases,
      settings: isPlainObject(rawData.settings) ? rawData.settings : null
    };
  }

  return {
    CURRENT_FORMAT_VERSION,
    isPlainObject,
    migrateImportData
  };
});
