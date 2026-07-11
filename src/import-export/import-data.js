(function attachImportData(root, factory) {
  const api = factory(root.TeamsTypingHelperConfig, root.TeamsTypingHelperExportData);
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.TeamsTypingHelperImportData = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createImportDataApi(configApi, exportApi) {
  "use strict";

  function isValidIsoDate(value) {
    return typeof value === "string" && !Number.isNaN(Date.parse(value));
  }

  function normalizeTextNewlines(value) {
    return String(value).replace(/\r\n/g, "\n");
  }

  function createUniquePhraseId(usedIds = new Set()) {
    let id;
    do {
      id = configApi.createId("phrase");
    } while (usedIds.has(id));
    usedIds.add(id);
    return id;
  }

  function normalizeImportedPhrase(rawPhrase, index, usedIds = new Set(), now = new Date().toISOString()) {
    return {
      id: createUniquePhraseId(usedIds),
      name: String(rawPhrase.name).trim(),
      text: String(rawPhrase.text),
      order: index,
      enabled: rawPhrase.enabled !== false,
      createdAt: isValidIsoDate(rawPhrase.createdAt) ? rawPhrase.createdAt : now,
      updatedAt: now
    };
  }

  function createPhraseFingerprint(phrase) {
    const normalizedName = phrase.name
      .trim()
      .replace(/\s+/g, " ")
      .toLocaleLowerCase();
    const normalizedText = normalizeTextNewlines(phrase.text).trim();
    return `${normalizedName}\u0000${normalizedText}`;
  }

  function detectDuplicatePhrases(importedPhrases, existingPhrases) {
    const existingFingerprints = new Set(existingPhrases.map(createPhraseFingerprint));
    return importedPhrases.map((phrase) => existingFingerprints.has(createPhraseFingerprint(phrase)));
  }

  function sanitizeImportedSettings(rawSettings, currentConfig) {
    const current = configApi.normalizeConfig(currentConfig);
    const settings = rawSettings && typeof rawSettings === "object" ? rawSettings : {};
    const result = {};

    if (typeof settings.toolbarEnabled === "boolean") {
      result.enabled = settings.toolbarEnabled;
    }
    if (typeof settings.compactMode === "boolean") {
      result.compactMode = settings.compactMode;
    }
    if (Number.isInteger(settings.buttonsPerRow)) {
      result.columns = Math.max(1, Math.min(10, settings.buttonsPerRow));
    }
    if (configApi.INSERT_MODES.includes(settings.insertMode)) {
      result.insertMode = settings.insertMode;
    }

    const toolbar = {};
    if (settings.toolbarPosition && typeof settings.toolbarPosition === "object") {
      const left = Number(settings.toolbarPosition.left);
      const top = Number(settings.toolbarPosition.top);
      if (Number.isFinite(left) && Number.isFinite(top)) {
        toolbar.position = { left, top };
      }
    }
    if (typeof settings.toolbarCollapsed === "boolean") {
      toolbar.collapsed = settings.toolbarCollapsed;
    }
    if (Object.keys(toolbar).length) {
      result.toolbar = { ...current.toolbar, ...toolbar };
    }

    return result;
  }

  function applySettings(config, rawSettings) {
    const settings = sanitizeImportedSettings(rawSettings, config);
    return configApi.normalizeConfig({
      ...config,
      ...settings,
      toolbar: {
        ...config.toolbar,
        ...(settings.toolbar || {})
      }
    });
  }

  function buildImportPreview(validationResult, currentConfig) {
    const current = configApi.normalizeConfig(currentConfig);
    const usedIds = new Set(current.phrases.map((phrase) => phrase.id));
    const now = new Date().toISOString();
    const normalizedPhrases = validationResult.validPhrases.map((item, index) => (
      normalizeImportedPhrase(item.raw, index, usedIds, now)
    ));
    const duplicateFlags = detectDuplicatePhrases(normalizedPhrases, current.phrases);
    const duplicateCount = duplicateFlags.filter(Boolean).length;
    const importablePhrases = normalizedPhrases.filter((phrase, index) => !duplicateFlags[index]);

    return {
      ...validationResult,
      normalizedPhrases,
      duplicateFlags,
      duplicateCount,
      finalImportCount: importablePhrases.length,
      importablePhrases,
      previewRows: normalizedPhrases.slice(0, 10).map((phrase, index) => ({
        name: phrase.name,
        text: phrase.text,
        valid: true,
        duplicate: duplicateFlags[index],
        reason: duplicateFlags[index] ? "重复词组，将跳过" : ""
      })),
      hiddenPreviewCount: Math.max(0, normalizedPhrases.length - 10)
    };
  }

  function renumberPhrases(phrases) {
    return phrases.map((phrase, index) => ({ ...phrase, order: index }));
  }

  function mergeImportedPhrases(currentConfig, preview, includeSettings) {
    let config = configApi.normalizeConfig(currentConfig);
    const phrases = renumberPhrases([
      ...config.phrases,
      ...preview.importablePhrases
    ]);
    config = configApi.normalizeConfig({ ...config, phrases });
    if (includeSettings && preview.containsSettings) {
      config = applySettings(config, preview.settings);
    }
    return {
      config,
      importedCount: preview.importablePhrases.length,
      skippedDuplicateCount: preview.duplicateCount,
      ignoredInvalidCount: preview.invalidPhrases.length
    };
  }

  function replaceImportedPhrases(currentConfig, preview, includeSettings) {
    let config = configApi.normalizeConfig({
      ...currentConfig,
      phrases: renumberPhrases(preview.normalizedPhrases)
    });
    if (includeSettings && preview.containsSettings) {
      config = applySettings(config, preview.settings);
    }
    return {
      config,
      importedCount: preview.normalizedPhrases.length,
      skippedDuplicateCount: 0,
      ignoredInvalidCount: preview.invalidPhrases.length
    };
  }

  function summarizeImportResult(result) {
    return `成功导入 ${result.importedCount} 个词组，跳过 ${result.skippedDuplicateCount} 个重复词组，忽略 ${result.ignoredInvalidCount} 个无效词组。`;
  }

  return {
    applySettings,
    buildImportPreview,
    createPhraseFingerprint,
    createUniquePhraseId,
    detectDuplicatePhrases,
    mergeImportedPhrases,
    normalizeImportedPhrase,
    replaceImportedPhrases,
    sanitizeImportedSettings,
    summarizeImportResult
  };
});
