const test = require("node:test");
const assert = require("node:assert/strict");

const configApi = require("../../src/shared/config.js");
global.TeamsTypingHelperConfig = configApi;
const exportApi = require("../../src/import-export/export-data.js");
global.TeamsTypingHelperExportData = exportApi;
const migrateApi = require("../../src/import-export/migrate-import.js");
global.TeamsTypingHelperMigrateImport = migrateApi;
const validateApi = require("../../src/import-export/validate-import.js");
global.TeamsTypingHelperValidateImport = validateApi;
const importApi = require("../../src/import-export/import-data.js");
global.TeamsTypingHelperImportData = importApi;
const backupApi = require("../../src/import-export/import-backup.js");

function config(phrases = [], extra = {}) {
  return configApi.normalizeConfig({
    enabled: true,
    insertMode: "cursor",
    compactMode: false,
    columns: 3,
    toolbar: { collapsed: false, position: null },
    phrases,
    ...extra
  });
}

function backup(phrases = [], settings = {}) {
  return {
    app: "teams-typing-helper",
    formatVersion: 1,
    exportedAt: "2026-07-11T21:15:30.000Z",
    phrases,
    settings
  };
}

function validate(payload, file = { name: "backup.json", size: 100 }) {
  return validateApi.validateImportText(JSON.stringify(payload), file);
}

test("exports Chinese phrases with the current JSON shape", () => {
  const data = exportApi.buildExportData(config([{ id: "a", name: "问候", text: "你好，请问现在方便吗？" }]), new Date("2026-07-11T21:15:30.000Z"));
  assert.equal(data.app, "teams-typing-helper");
  assert.equal(data.formatVersion, 1);
  assert.equal(data.phrases[0].name, "问候");
  assert.equal(data.settings.buttonsPerRow, 3);
});

test("exports multiline unicode phrases without losing content", () => {
  const text = "Hola\nПривет\n😀";
  const json = JSON.stringify(exportApi.buildExportData(config([{ name: "多语言", text }]), new Date()), null, 2);
  assert.match(json, /Hola\\n/);
  assert.match(json, /Привет/);
  assert.match(json, /😀/);
});

test("exports an empty phrase list", () => {
  const data = exportApi.buildExportData(config([]));
  assert.deepEqual(data.phrases, []);
});

test("exported JSON can be validated for import", () => {
  const data = exportApi.buildExportData(config([{ name: "A", text: "one" }]));
  const validation = validate(data);
  assert.equal(validation.validPhrases.length, 1);
});

test("builds the required backup filename", () => {
  assert.equal(exportApi.buildExportFilename(new Date("2026-07-11T23:15:30")), "teams-typing-helper-backup-2026-07-11-231530.json");
});

test("merge import appends non-duplicate phrases", () => {
  const current = config([{ name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "two" }])), current);
  const result = importApi.mergeImportedPhrases(current, preview, false);
  assert.deepEqual(result.config.phrases.map((phrase) => phrase.name), ["A", "B"]);
});

test("replace import replaces all phrases", () => {
  const current = config([{ name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "two" }])), current);
  const result = importApi.replaceImportedPhrases(current, preview, false);
  assert.deepEqual(result.config.phrases.map((phrase) => phrase.name), ["B"]);
});

test("merge skips exact duplicate name and text", () => {
  const current = config([{ name: " A ", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ name: "a", text: "one" }])), current);
  assert.equal(preview.duplicateCount, 1);
  assert.equal(importApi.mergeImportedPhrases(current, preview, false).importedCount, 0);
});

test("same name with different text is allowed", () => {
  const current = config([{ name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ name: "A", text: "two" }])), current);
  assert.equal(preview.finalImportCount, 1);
});

test("different name with same text is allowed", () => {
  const current = config([{ name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "one" }])), current);
  assert.equal(preview.finalImportCount, 1);
});

test("id conflicts generate new ids", () => {
  const current = config([{ id: "same", name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validate(backup([{ id: "same", name: "B", text: "two" }])), current);
  assert.notEqual(preview.normalizedPhrases[0].id, "same");
});

test("damaged JSON reports a clear error", () => {
  assert.throws(() => validateApi.validateImportText("{bad"), /JSON 文件内容损坏/);
});

test("wrong app name is rejected", () => {
  assert.throws(() => validate({ app: "wrong-app", formatVersion: 1, phrases: [] }), /Teams 辅助打字/);
});

test("missing phrases is rejected", () => {
  assert.throws(() => validate({ app: "teams-typing-helper", formatVersion: 1 }), /phrases 数组/);
});

test("blank name is invalid", () => {
  assert.throws(() => validate(backup([{ name: " ", text: "text" }])), /没有可导入/);
});

test("blank text is invalid", () => {
  assert.throws(() => validate(backup([{ name: "A", text: "   " }])), /没有可导入/);
});

test("overlong phrase fields are invalid", () => {
  assert.throws(() => validate(backup([{ name: "x".repeat(101), text: "ok" }])), /没有可导入/);
  assert.throws(() => validate(backup([{ name: "A", text: "x".repeat(10001) }])), /没有可导入/);
});

test("more than 1000 phrases is rejected", () => {
  assert.throws(() => validate(backup(Array.from({ length: 1001 }, (_, index) => ({ name: `P${index}`, text: "x" })))), /1000/);
});

test("files larger than five megabytes are rejected before parsing", () => {
  assert.throws(() => validateApi.validateImportText("{}", { name: "big.json", size: 5 * 1024 * 1024 + 1 }), /5 MB/);
});

test("future format versions are rejected", () => {
  assert.throws(() => validate({ app: "teams-typing-helper", formatVersion: 999, phrases: [] }), /更高版本/);
});

test("legacy backups without formatVersion migrate", () => {
  const result = validateApi.validateImportData({ phrases: [{ name: "Old", text: "ok" }] });
  assert.equal(result.legacy, true);
  assert.equal(result.validPhrases.length, 1);
});

test("legacy-looking backup with wrong app is rejected", () => {
  assert.throws(() => validateApi.validateImportData({ app: "wrong", phrases: [{ name: "Old", text: "ok" }] }), /Teams 辅助打字/);
});

test("malicious html remains plain text in preview data", () => {
  const preview = importApi.buildImportPreview(validate(backup([{ name: "<img onerror=alert(1)>", text: "<script>alert(1)</script>" }])), config([]));
  assert.equal(preview.previewRows[0].name, "<img onerror=alert(1)>");
  assert.equal(preview.previewRows[0].text, "<script>alert(1)</script>");
});

test("proto fields do not pollute Object prototype", () => {
  const text = '{"app":"teams-typing-helper","formatVersion":1,"phrases":[{"name":"A","text":"B","__proto__":{"polluted":true}}]}';
  validateApi.validateImportText(text);
  assert.equal({}.polluted, undefined);
});

test("settings stay unchanged when not included", () => {
  const current = config([{ name: "A", text: "one" }], { insertMode: "append", columns: 4 });
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "two" }], { insertMode: "replace", buttonsPerRow: 8 })), current);
  const result = importApi.mergeImportedPhrases(current, preview, false);
  assert.equal(result.config.insertMode, "append");
  assert.equal(result.config.columns, 4);
});

test("settings import only whitelisted fields", () => {
  const current = config([], { insertMode: "cursor", columns: 3 });
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "two" }], { insertMode: "replace", buttonsPerRow: 8, unknown: "<b>x</b>" })), current);
  const result = importApi.mergeImportedPhrases(current, preview, true);
  assert.equal(result.config.insertMode, "replace");
  assert.equal(result.config.columns, 8);
  assert.equal(result.config.unknown, undefined);
});

test("toolbar position settings are sanitized", () => {
  const settings = importApi.sanitizeImportedSettings({ toolbarPosition: { left: 10, top: 20 }, toolbarCollapsed: true }, config([]));
  assert.deepEqual(settings.toolbar, { collapsed: true, position: { left: 10, top: 20 } });
});

test("invalid toolbar position is ignored", () => {
  const settings = importApi.sanitizeImportedSettings({ toolbarPosition: { left: "bad", top: 20 } }, config([]));
  assert.equal(settings.toolbar, undefined);
});

test("import backup payload contains phrases and settings", () => {
  const backupPayload = backupApi.createImportBackup(config([{ name: "A", text: "one" }]), new Date("2026-07-11T21:00:00Z"));
  assert.equal(backupPayload.createdAt, "2026-07-11T21:00:00.000Z");
  assert.equal(backupPayload.phrases.length, 1);
  assert.equal(backupPayload.settings.insertMode, "cursor");
});

test("restore recent backup through storage mock", async () => {
  const store = {};
  global.chrome = { runtime: { lastError: null } };
  const storage = {
    get(keys, cb) {
      cb(typeof keys === "string" ? { [keys]: store[keys] } : store);
    },
    set(values, cb) {
      Object.assign(store, values);
      cb();
    }
  };
  const oldConfig = config([{ name: "Old", text: "one" }]);
  await backupApi.createImportBackupInStorage(storage, oldConfig);
  await storage.set({ [configApi.STORAGE_KEY]: config([{ name: "New", text: "two" }]) }, () => {});
  const restored = await backupApi.restoreImportBackup(storage);
  assert.equal(restored.phrases[0].name, "Old");
});

test("merge summary reports imported duplicate and invalid counts", () => {
  const validation = validate(backup([{ name: "A", text: "one" }, { name: "B", text: "two" }, { name: "", text: "" }]));
  const current = config([{ name: "A", text: "one" }]);
  const preview = importApi.buildImportPreview(validation, current);
  const result = importApi.mergeImportedPhrases(current, preview, false);
  assert.match(importApi.summarizeImportResult(result), /成功导入 1 个词组，跳过 1 个重复词组，忽略 1 个无效词组/);
});

test("orders are recalculated after import", () => {
  const preview = importApi.buildImportPreview(validate(backup([{ name: "B", text: "two", order: 99 }])), config([{ name: "A", text: "one", order: 55 }]));
  const result = importApi.mergeImportedPhrases(config([{ name: "A", text: "one", order: 55 }]), preview, false);
  assert.deepEqual(result.config.phrases.map((phrase) => phrase.order), [0, 1]);
});

test("empty phrase backup can be replaced", () => {
  const preview = importApi.buildImportPreview(validate(backup([])), config([{ name: "A", text: "one" }]));
  const result = importApi.replaceImportedPhrases(config([{ name: "A", text: "one" }]), preview, false);
  assert.deepEqual(result.config.phrases, []);
});

test("same file content can be validated repeatedly", () => {
  const payload = JSON.stringify(backup([{ name: "A", text: "one" }]));
  assert.equal(validateApi.validateImportText(payload).validPhrases.length, 1);
  assert.equal(validateApi.validateImportText(payload).validPhrases.length, 1);
});
