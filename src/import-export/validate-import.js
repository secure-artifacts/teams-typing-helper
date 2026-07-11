(function attachValidateImport(root, factory) {
  const api = factory(root.TeamsTypingHelperMigrateImport);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperValidateImport = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createValidateImportApi(migrateApi) {
  "use strict";

  const MAX_FILE_SIZE = 5 * 1024 * 1024;
  const MAX_PHRASES = 1000;
  const MAX_NAME_LENGTH = 100;
  const MAX_TEXT_LENGTH = 10000;

  function parseJson(text) {
    try {
      return JSON.parse(text);
    } catch (error) {
      throw new Error("JSON 文件内容损坏，无法解析。");
    }
  }

  function validatePhrase(rawPhrase, index) {
    if (!migrateApi.isPlainObject(rawPhrase)) {
      return { valid: false, reason: `第 ${index + 1} 个词组必须是对象。` };
    }
    if (typeof rawPhrase.name !== "string") {
      return { valid: false, reason: `第 ${index + 1} 个词组缺少按钮名称。` };
    }
    if (typeof rawPhrase.text !== "string") {
      return { valid: false, reason: `第 ${index + 1} 个词组缺少完整文字。` };
    }
    const name = rawPhrase.name.trim();
    const text = rawPhrase.text;
    if (!name) {
      return { valid: false, reason: `第 ${index + 1} 个词组缺少按钮名称。` };
    }
    if (!text.trim()) {
      return { valid: false, reason: `第 ${index + 1} 个词组文字不能为空。` };
    }
    if (name.length > MAX_NAME_LENGTH) {
      return { valid: false, reason: `第 ${index + 1} 个词组按钮名称超过 ${MAX_NAME_LENGTH} 个字符。` };
    }
    if (text.length > MAX_TEXT_LENGTH) {
      return { valid: false, reason: `第 ${index + 1} 个词组文字超过 ${MAX_TEXT_LENGTH} 个字符。` };
    }
    return { valid: true, phrase: rawPhrase };
  }

  function validateImportData(rawData, fileInfo = {}) {
    const migrated = migrateApi.migrateImportData(rawData);
    if (!Array.isArray(migrated.phrases)) {
      throw new Error("备份文件缺少 phrases 数组。");
    }
    if (migrated.phrases.length > MAX_PHRASES) {
      throw new Error("词组数量超过 1000 个。");
    }

    const validPhrases = [];
    const invalidPhrases = [];
    migrated.phrases.forEach((rawPhrase, index) => {
      const result = validatePhrase(rawPhrase, index);
      if (result.valid) {
        validPhrases.push({ raw: result.phrase, sourceIndex: index });
      } else {
        invalidPhrases.push({ raw: rawPhrase, sourceIndex: index, reason: result.reason });
      }
    });

    if (migrated.phrases.length > 0 && validPhrases.length === 0) {
      throw new Error("文件中没有可导入的有效词组。");
    }

    return {
      fileName: fileInfo.name || "",
      fileSize: Number.isFinite(fileInfo.size) ? fileInfo.size : null,
      exportedAt: migrated.exportedAt,
      formatVersion: migrated.formatVersion,
      legacy: migrated.legacy,
      legacyMessage: migrated.legacyMessage || "",
      totalCount: migrated.phrases.length,
      validPhrases,
      invalidPhrases,
      settings: migrated.settings,
      containsSettings: Boolean(migrated.settings)
    };
  }

  function validateImportText(text, fileInfo = {}) {
    if (Number.isFinite(fileInfo.size) && fileInfo.size > MAX_FILE_SIZE) {
      throw new Error("文件超过 5 MB，无法导入。");
    }
    return validateImportData(parseJson(text), fileInfo);
  }

  async function readImportFile(file) {
    if (!file) {
      throw new Error("未选择导入文件。");
    }
    if (file.size > MAX_FILE_SIZE) {
      throw new Error("文件超过 5 MB，无法导入。");
    }
    const text = await file.text();
    return validateImportText(text, { name: file.name, size: file.size });
  }

  return {
    MAX_FILE_SIZE,
    MAX_PHRASES,
    MAX_NAME_LENGTH,
    MAX_TEXT_LENGTH,
    parseJson,
    readImportFile,
    validateImportData,
    validateImportText,
    validatePhrase
  };
});
