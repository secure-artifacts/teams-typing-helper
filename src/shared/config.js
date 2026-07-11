(function attachConfig(root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperConfig = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createConfigApi() {
  "use strict";

  const STORAGE_KEY = "teamsTypingHelperConfig";
  const IMPORT_BACKUP_KEY = "teamsTypingHelperImportBackup";
  const APP_NAME = "teams-typing-helper";
  const EXPORT_VERSION = 1;

  const INSERT_MODES = ["cursor", "append", "replace"];

  const DEFAULT_CONFIG = {
    version: 1,
    enabled: true,
    insertMode: "cursor",
    compactMode: false,
    columns: 3,
    toolbar: {
      collapsed: false,
      position: null
    },
    phrases: [
      {
        id: "default-greeting",
        name: "问候",
        text: "您好，我现在看一下，稍后回复您。"
      },
      {
        id: "default-thanks",
        name: "感谢",
        text: "谢谢，我已经收到。"
      },
      {
        id: "default-followup",
        name: "稍后处理",
        text: "我需要再确认一下，确认后马上回复您。"
      }
    ]
  };

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function asString(value) {
    return typeof value === "string" ? value : "";
  }

  function createId(prefix) {
    const safePrefix = asString(prefix).trim() || "phrase";
    if (typeof crypto !== "undefined" && crypto.randomUUID) {
      return `${safePrefix}-${crypto.randomUUID()}`;
    }
    return `${safePrefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function normalizePhrase(raw, index) {
    const source = raw && typeof raw === "object" ? raw : {};
    const text = asString(source.text);
    let name = asString(source.name).trim();
    if (!name && text) {
      name = text.replace(/\s+/g, " ").trim().slice(0, 24);
    }
    if (!name) {
      name = `词组 ${index + 1}`;
    }

    return {
      id: asString(source.id).trim() || createId("phrase"),
      name,
      text,
      order: Number.isFinite(source.order) ? Math.max(0, Math.round(source.order)) : index,
      enabled: source.enabled !== false,
      createdAt: isValidIsoDate(source.createdAt) ? source.createdAt : null,
      updatedAt: isValidIsoDate(source.updatedAt) ? source.updatedAt : null
    };
  }

  function isValidIsoDate(value) {
    return typeof value === "string" && !Number.isNaN(Date.parse(value));
  }

  function normalizeToolbar(rawToolbar) {
    const toolbar = rawToolbar && typeof rawToolbar === "object" ? rawToolbar : {};
    const position = toolbar.position && typeof toolbar.position === "object"
      ? {
        left: Number.isFinite(toolbar.position.left) ? toolbar.position.left : null,
        top: Number.isFinite(toolbar.position.top) ? toolbar.position.top : null
      }
      : null;

    return {
      collapsed: Boolean(toolbar.collapsed),
      position: position && Number.isFinite(position.left) && Number.isFinite(position.top)
        ? position
        : null
    };
  }

  function normalizeConfig(raw) {
    if (!raw || typeof raw !== "object") {
      return clone(DEFAULT_CONFIG);
    }

    const phrasesSource = Array.isArray(raw.phrases) ? raw.phrases : DEFAULT_CONFIG.phrases;
    const phrases = phrasesSource
      .map((phrase, index) => normalizePhrase(phrase, index))
      .filter((phrase) => phrase.name.trim() || phrase.text);

    const insertMode = INSERT_MODES.includes(raw.insertMode) ? raw.insertMode : DEFAULT_CONFIG.insertMode;
    const columns = Number.isFinite(raw.columns)
      ? Math.max(1, Math.min(10, Math.round(raw.columns)))
      : DEFAULT_CONFIG.columns;

    return {
      version: 1,
      enabled: typeof raw.enabled === "boolean" ? raw.enabled : DEFAULT_CONFIG.enabled,
      insertMode,
      compactMode: Boolean(raw.compactMode),
      columns,
      toolbar: normalizeToolbar(raw.toolbar),
      phrases: phrases.map((phrase, index) => ({ ...phrase, order: index }))
    };
  }

  function buildExportPayload(config) {
    return {
      app: APP_NAME,
      exportVersion: EXPORT_VERSION,
      exportedAt: new Date().toISOString(),
      config: normalizeConfig(config)
    };
  }

  function parseImportedConfig(text) {
    let payload;
    try {
      payload = JSON.parse(text);
    } catch (error) {
      throw new Error("导入失败：文件不是有效的 JSON。");
    }

    const maybeConfig = payload && typeof payload === "object" && payload.config
      ? payload.config
      : payload;

    const config = normalizeConfig(maybeConfig);
    if (!Array.isArray(config.phrases)) {
      throw new Error("导入失败：配置中缺少词组列表。");
    }
    return config;
  }

  return {
    APP_NAME,
    STORAGE_KEY,
    IMPORT_BACKUP_KEY,
    DEFAULT_CONFIG,
    INSERT_MODES,
    clone,
    createId,
    normalizeConfig,
    normalizePhrase,
    buildExportPayload,
    parseImportedConfig
  };
});
