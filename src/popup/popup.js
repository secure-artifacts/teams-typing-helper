(function attachPopup(root) {
  "use strict";

  const configApi = root.TeamsTypingHelperConfig;
  const exportApi = root.TeamsTypingHelperExportData;
  const validateApi = root.TeamsTypingHelperValidateImport;
  const importApi = root.TeamsTypingHelperImportData;
  const backupApi = root.TeamsTypingHelperImportBackup;
  const documentRef = root.document;

  const state = {
    config: configApi.normalizeConfig(null),
    editingId: "",
    pendingImport: null,
    activeActionMenu: null,
    activeMenuTrigger: null,
    activeMenuPhraseId: "",
    ignoreMenuScrollCloseUntil: 0,
    deleteCandidate: null,
    lastFocusedElement: null,
    hasBackup: false,
    busy: false
  };

  const elements = {
    status: documentRef.getElementById("status"),
    formTitle: documentRef.getElementById("form-title"),
    form: documentRef.getElementById("phrase-form"),
    id: documentRef.getElementById("phrase-id"),
    name: documentRef.getElementById("phrase-name"),
    text: documentRef.getElementById("phrase-text"),
    savePhrase: documentRef.getElementById("save-phrase"),
    cancelEdit: documentRef.getElementById("cancel-edit"),
    list: documentRef.getElementById("phrase-list"),
    phraseCount: documentRef.getElementById("phrase-count"),
    floatingLayer: documentRef.getElementById("tth-floating-layer"),
    enabled: documentRef.getElementById("enabled"),
    compactMode: documentRef.getElementById("compact-mode"),
    insertMode: documentRef.getElementById("insert-mode"),
    columns: documentRef.getElementById("columns"),
    exportConfig: documentRef.getElementById("export-config"),
    importTrigger: documentRef.getElementById("import-trigger"),
    importConfig: documentRef.getElementById("import-config"),
    restoreBackup: documentRef.getElementById("restore-backup"),
    undoImport: documentRef.getElementById("undo-import"),
    backupHint: documentRef.getElementById("backup-hint"),
    resetDefaults: documentRef.getElementById("reset-defaults"),
    importPreview: documentRef.getElementById("import-preview"),
    closeImportPreview: documentRef.getElementById("close-import-preview"),
    importSummary: documentRef.getElementById("import-summary"),
    importWarning: documentRef.getElementById("import-warning"),
    importSettings: documentRef.getElementById("import-settings"),
    phrasePreviewList: documentRef.getElementById("phrase-preview-list"),
    previewMore: documentRef.getElementById("preview-more"),
    confirmImport: documentRef.getElementById("confirm-import"),
    cancelImport: documentRef.getElementById("cancel-import"),
    deleteConfirm: documentRef.getElementById("delete-confirm"),
    cancelDelete: documentRef.getElementById("cancel-delete"),
    confirmDelete: documentRef.getElementById("confirm-delete")
  };

  function hasChromeStorage() {
    return Boolean(root.chrome && root.chrome.storage && root.chrome.storage.local);
  }

  function storageGet(keys) {
    return new Promise((resolve) => {
      if (!hasChromeStorage()) {
        resolve({});
        return;
      }
      root.chrome.storage.local.get(keys, (result) => {
        resolve(result || {});
      });
    });
  }

  function storageSetValues(values) {
    return new Promise((resolve, reject) => {
      if (!hasChromeStorage()) {
        if (values[configApi.STORAGE_KEY]) {
          state.config = configApi.normalizeConfig(values[configApi.STORAGE_KEY]);
        }
        if (values[configApi.IMPORT_BACKUP_KEY]) {
          state.hasBackup = true;
        }
        resolve();
        return;
      }
      root.chrome.storage.local.set(values, () => {
        const error = root.chrome.runtime && root.chrome.runtime.lastError;
        if (error) {
          reject(new Error(error.message));
          return;
        }
        resolve();
      });
    });
  }

  async function readLatestConfig() {
    const stored = await storageGet(configApi.STORAGE_KEY);
    return configApi.normalizeConfig(stored[configApi.STORAGE_KEY]);
  }

  async function storageSet(config) {
    await storageSetValues({ [configApi.STORAGE_KEY]: configApi.normalizeConfig(config) });
  }

  function statusTone(message) {
    if (/失败|错误|无法|未找到|请填写|没有可恢复/.test(message || "")) {
      return "error";
    }
    if (/警告|确认|覆盖|替换/.test(message || "")) {
      return "warning";
    }
    return "success";
  }

  function showStatus(message, keepVisible) {
    root.clearTimeout(showStatus.timer);
    elements.status.textContent = message || "";
    elements.status.dataset.tone = statusTone(message);
    elements.status.classList.toggle("is-visible", Boolean(message));
    if (message) {
      showStatus.timer = root.setTimeout(() => {
        elements.status.classList.remove("is-visible");
      }, 2500);
    }
  }

  function setBusy(isBusy, label) {
    state.busy = isBusy;
    [elements.exportConfig, elements.importTrigger, elements.restoreBackup, elements.undoImport, elements.confirmImport].forEach((button) => {
      if (button) {
        button.disabled = isBusy || (button === elements.restoreBackup && !state.hasBackup);
      }
    });
    if (label) {
      showStatus(label, true);
    }
  }

  async function saveConfig(message) {
    state.config = configApi.normalizeConfig(state.config);
    await storageSet(state.config);
    render();
    showStatus(message || "已保存。");
  }

  function resetForm() {
    state.editingId = "";
    elements.id.value = "";
    elements.name.value = "";
    elements.text.value = "";
    elements.formTitle.textContent = "新建词组";
    elements.savePhrase.textContent = "＋ 添加词组";
    elements.cancelEdit.hidden = true;
  }

  function movePhrase(index, direction) {
    const nextIndex = index + direction;
    if (nextIndex < 0 || nextIndex >= state.config.phrases.length) {
      return;
    }
    const phrases = state.config.phrases.slice();
    const [phrase] = phrases.splice(index, 1);
    phrases.splice(nextIndex, 0, phrase);
    state.config.phrases = phrases;
    saveConfig("顺序已更新。");
  }

  function editPhrase(phrase) {
    closePhraseMenu();
    state.editingId = phrase.id;
    elements.id.value = phrase.id;
    elements.name.value = phrase.name;
    elements.text.value = phrase.text;
    elements.formTitle.textContent = "编辑词组";
    elements.savePhrase.textContent = "✓ 保存修改";
    elements.cancelEdit.hidden = false;
    elements.name.focus();
  }

  function openDeleteConfirm(phrase, trigger) {
    closeActionMenu({ restoreFocus: false });
    state.deleteCandidate = phrase;
    state.lastFocusedElement = trigger || documentRef.activeElement;
    elements.deleteConfirm.hidden = false;
    elements.cancelDelete.focus();
  }

  function closeDeleteConfirm() {
    state.deleteCandidate = null;
    elements.deleteConfirm.hidden = true;
    if (state.lastFocusedElement && typeof state.lastFocusedElement.focus === "function") {
      state.lastFocusedElement.focus();
    }
    state.lastFocusedElement = null;
  }

  async function confirmDeletePhrase() {
    const phrase = state.deleteCandidate;
    if (!phrase) {
      return;
    }
    state.config.phrases = state.config.phrases.filter((item) => item.id !== phrase.id);
    if (state.editingId === phrase.id) {
      resetForm();
    }
    closeDeleteConfirm();
    await saveConfig("词组已删除。");
  }

  function createSvgIcon(name) {
    const svg = documentRef.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("focusable", "false");
    const path = documentRef.createElementNS("http://www.w3.org/2000/svg", "path");
    const paths = {
      dots: "M12 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4Zm0 2a2 2 0 1 0 0 4 2 2 0 0 0 0-4Zm0 6a2 2 0 1 0 0 4 2 2 0 0 0 0-4Z",
      up: "M12 5 5.5 11.5l1.4 1.4 4.1-4.1V20h2V8.8l4.1 4.1 1.4-1.4L12 5Z",
      down: "M11 4v11.2l-4.1-4.1-1.4 1.4L12 19l6.5-6.5-1.4-1.4-4.1 4.1V4h-2Z",
      edit: "m5 16.6-.7 3.1 3.1-.7 10-10-2.4-2.4-10 10ZM16.2 4.4l2.4 2.4 1-1a1.7 1.7 0 0 0-2.4-2.4l-1 1Z",
      trash: "M8 7h8l-.7 12.2A2 2 0 0 1 13.3 21h-2.6a2 2 0 0 1-2-1.8L8 7Zm1-3h6l1 2h4v2H4V6h4l1-2Zm1 6h2v8h-2v-8Zm4 0h2v8h-2v-8Z",
      bubble: "M4 5.5A3.5 3.5 0 0 1 7.5 2h9A3.5 3.5 0 0 1 20 5.5v6A3.5 3.5 0 0 1 16.5 15H11l-5.5 5v-5A3.5 3.5 0 0 1 2 11.5v-6h2Z",
      download: "M11 4h2v8.2l3.2-3.2 1.4 1.4L12 16l-5.6-5.6L7.8 9l3.2 3.2V4Zm-6 13h14v3H5v-3Z",
      upload: "M11 20v-8.2L7.8 15 6.4 13.6 12 8l5.6 5.6-1.4 1.4-3.2-3.2V20h-2ZM5 4h14v3H5V4Z",
      history: "M12 5a7 7 0 1 1-6.3 4H3l4-4 4 4H8a5 5 0 1 0 4-2V5Zm-1 4h2v4l3 1.8-1 1.7-4-2.4V9Z"
    };
    path.setAttribute("d", paths[name] || paths.dots);
    path.setAttribute("fill", "currentColor");
    svg.appendChild(path);
    return svg;
  }

  function prependButtonIcon(button, iconName, className) {
    const icon = createSvgIcon(iconName);
    icon.classList.add(className || "button-icon");
    button.prepend(icon);
  }

  function createMenuItem(iconName, label, title, disabled, danger, onClick) {
    const button = documentRef.createElement("button");
    button.type = "button";
    button.className = `tth-action-menu-item menu-item${danger ? " tth-action-menu-item-danger is-danger" : ""}`;
    button.setAttribute("role", "menuitem");
    button.setAttribute("aria-label", title);
    button.title = title;
    button.disabled = disabled;
    const icon = createSvgIcon(iconName);
    icon.classList.add("menu-icon");
    const text = documentRef.createElement("span");
    text.textContent = label;
    button.append(icon, text);
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      onClick(button);
    });
    return button;
  }

  function getFloatingLayer() {
    if (elements.floatingLayer) {
      return elements.floatingLayer;
    }
    const layer = documentRef.createElement("div");
    layer.id = "tth-floating-layer";
    documentRef.body.appendChild(layer);
    elements.floatingLayer = layer;
    return layer;
  }

  function positionActionMenu(menu, triggerButton) {
    const viewportPadding = 8;
    const gap = 6;
    const triggerRect = triggerButton.getBoundingClientRect();
    const menuWidth = menu.offsetWidth;
    const menuHeight = menu.offsetHeight;
    const spaceBelow = root.innerHeight - triggerRect.bottom - viewportPadding;
    const spaceAbove = triggerRect.top - viewportPadding;
    const shouldOpenUpward = spaceBelow < menuHeight && spaceAbove > spaceBelow;
    let top = shouldOpenUpward
      ? triggerRect.top - menuHeight - gap
      : triggerRect.bottom + gap;
    let left = triggerRect.right - menuWidth;

    left = Math.max(
      viewportPadding,
      Math.min(left, root.innerWidth - menuWidth - viewportPadding)
    );
    top = Math.max(
      viewportPadding,
      Math.min(top, root.innerHeight - menuHeight - viewportPadding)
    );

    menu.style.left = `${Math.round(left)}px`;
    menu.style.top = `${Math.round(top)}px`;
    menu.dataset.placement = shouldOpenUpward ? "top" : "bottom";
  }

  function getEnabledMenuItems(menu) {
    return Array.from(menu.querySelectorAll("[role='menuitem']")).filter((item) => !item.disabled);
  }

  function focusMenuItem(menu, direction) {
    const items = getEnabledMenuItems(menu);
    if (!items.length) {
      return;
    }
    const activeIndex = items.indexOf(documentRef.activeElement);
    const nextIndex = activeIndex === -1
      ? (direction > 0 ? 0 : items.length - 1)
      : (activeIndex + direction + items.length) % items.length;
    items[nextIndex].focus();
  }

  function closeActionMenu(options = {}) {
    const restoreFocus = options.restoreFocus !== false;
    const trigger = state.activeMenuTrigger;
    if (state.activeActionMenu) {
      state.activeActionMenu.remove();
    }
    state.activeActionMenu = null;
    state.activeMenuTrigger = null;
    state.activeMenuPhraseId = "";
    if (trigger) {
      trigger.setAttribute("aria-expanded", "false");
      if (restoreFocus && typeof trigger.focus === "function" && documentRef.contains(trigger)) {
        trigger.focus();
      }
    }
  }

  function handleActionMenuKeydown(event) {
    if (!state.activeActionMenu) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      closeActionMenu();
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      focusMenuItem(state.activeActionMenu, 1);
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      focusMenuItem(state.activeActionMenu, -1);
      return;
    }
    if (event.key === "Home") {
      event.preventDefault();
      const items = getEnabledMenuItems(state.activeActionMenu);
      if (items[0]) {
        items[0].focus();
      }
      return;
    }
    if (event.key === "End") {
      event.preventDefault();
      const items = getEnabledMenuItems(state.activeActionMenu);
      if (items[items.length - 1]) {
        items[items.length - 1].focus();
      }
    }
  }

  function openActionMenu(phrase, index, phraseCount, triggerButton) {
    if (state.activeMenuTrigger === triggerButton) {
      closeActionMenu();
      return;
    }

    closeActionMenu({ restoreFocus: false });
    const menu = documentRef.createElement("div");
    menu.className = "tth-action-menu phrase-menu";
    menu.setAttribute("role", "menu");
    menu.style.visibility = "hidden";
    menu.append(
      createMenuItem("up", "上移", "上移词组", index === 0, false, () => {
        closeActionMenu({ restoreFocus: false });
        movePhrase(index, -1);
      }),
      createMenuItem("down", "下移", "下移词组", index === phraseCount - 1, false, () => {
        closeActionMenu({ restoreFocus: false });
        movePhrase(index, 1);
      }),
      createMenuItem("edit", "编辑", "编辑词组", false, false, () => {
        closeActionMenu({ restoreFocus: false });
        editPhrase(phrase);
      }),
      createMenuItem("trash", "删除", "删除词组", false, true, () => {
        openDeleteConfirm(phrase, triggerButton);
      })
    );

    getFloatingLayer().appendChild(menu);
    state.activeActionMenu = menu;
    state.activeMenuTrigger = triggerButton;
    state.activeMenuPhraseId = phrase.id;
    state.ignoreMenuScrollCloseUntil = Date.now() + 150;
    triggerButton.setAttribute("aria-expanded", "true");
    positionActionMenu(menu, triggerButton);
    menu.style.visibility = "visible";
    const firstEnabledItem = getEnabledMenuItems(menu)[0];
    if (firstEnabledItem) {
      firstEnabledItem.focus();
    }
  }

  function closePhraseMenu() {
    closeActionMenu();
  }

  function closeActionMenuAfterScroll() {
    if (Date.now() < state.ignoreMenuScrollCloseUntil) {
      return;
    }
    closeActionMenu({ restoreFocus: false });
  }

  function renderPhraseList() {
    closeActionMenu({ restoreFocus: false });
    elements.list.innerHTML = "";
    const phraseCount = state.config.phrases.length;
    elements.phraseCount.textContent = `${phraseCount} 个`;

    if (!phraseCount) {
      const empty = documentRef.createElement("div");
      empty.className = "empty-state";
      const icon = documentRef.createElement("div");
      icon.className = "empty-icon";
      icon.appendChild(createSvgIcon("bubble"));
      const title = documentRef.createElement("h3");
      title.textContent = "还没有词组";
      const copy = documentRef.createElement("p");
      copy.textContent = "创建一个常用回复，之后就能一键插入 Teams。";
      const button = documentRef.createElement("button");
      button.type = "button";
      button.className = "button-primary";
      button.textContent = "添加第一个词组";
      button.addEventListener("click", () => elements.name.focus());
      empty.append(icon, title, copy, button);
      elements.list.appendChild(empty);
      return;
    }

    state.config.phrases.forEach((phrase, index) => {
      const row = documentRef.createElement("article");
      row.className = "phrase-row";

      const header = documentRef.createElement("div");
      header.className = "phrase-card-header";

      const main = documentRef.createElement("div");
      main.className = "phrase-main";

      const name = documentRef.createElement("div");
      name.className = "phrase-name";
      name.textContent = phrase.name;
      name.title = phrase.name;

      const text = documentRef.createElement("div");
      text.className = "phrase-text";
      text.textContent = phrase.text;
      text.title = phrase.text;

      main.append(name, text);

      const menuButton = documentRef.createElement("button");
      menuButton.type = "button";
      menuButton.className = "phrase-menu-button";
      menuButton.title = "打开词组操作菜单";
      menuButton.setAttribute("aria-label", "打开词组操作菜单");
      menuButton.setAttribute("aria-haspopup", "menu");
      menuButton.setAttribute("aria-expanded", "false");
      menuButton.appendChild(createSvgIcon("dots"));
      menuButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openActionMenu(phrase, index, phraseCount, menuButton);
      });

      header.append(main, menuButton);

      const meta = documentRef.createElement("div");
      meta.className = "phrase-meta";
      meta.textContent = "可在工具栏中使用";

      row.append(header, meta);

      elements.list.appendChild(row);
    });
  }

  function renderSettings() {
    elements.enabled.checked = state.config.enabled;
    elements.compactMode.checked = state.config.compactMode;
    elements.insertMode.value = state.config.insertMode;
    elements.columns.value = String(state.config.columns);
  }

  function renderBackupState() {
    elements.restoreBackup.disabled = state.busy || !state.hasBackup;
    if (!state.hasBackup) {
      elements.restoreBackup.title = "当前没有可恢复的导入备份。";
    } else {
      elements.restoreBackup.title = "恢复最近一次导入前的数据";
    }
  }

  function decorateStaticIcons() {
    documentRef.querySelectorAll(".icon-text-button[data-icon]").forEach((button) => {
      if (button.querySelector("svg")) {
        return;
      }
      prependButtonIcon(button, button.dataset.icon);
    });
  }

  function render() {
    renderPhraseList();
    renderSettings();
    renderBackupState();
  }

  function bindForm() {
    elements.form.addEventListener("submit", (event) => {
      event.preventDefault();
      const name = elements.name.value.trim();
      const text = elements.text.value;
      if (!name || !text.trim()) {
        showStatus("请填写按钮名称和完整文字。");
        return;
      }

      if (state.editingId) {
        state.config.phrases = state.config.phrases.map((phrase) => (
          phrase.id === state.editingId ? { ...phrase, name, text, updatedAt: new Date().toISOString() } : phrase
        ));
        resetForm();
        saveConfig("词组已修改。");
        return;
      }

      const now = new Date().toISOString();
      state.config.phrases.push({
        id: configApi.createId("phrase"),
        name,
        text,
        order: state.config.phrases.length,
        enabled: true,
        createdAt: now,
        updatedAt: now
      });
      resetForm();
      saveConfig("词组已添加。");
    });

    elements.cancelEdit.addEventListener("click", resetForm);
  }

  function bindGlobalUi() {
    documentRef.addEventListener("click", (event) => {
      if (!state.activeActionMenu) {
        return;
      }
      if (
        state.activeActionMenu.contains(event.target) ||
        (state.activeMenuTrigger && state.activeMenuTrigger.contains(event.target))
      ) {
        return;
      }
      closeActionMenu({ restoreFocus: false });
    });
    documentRef.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") {
        handleActionMenuKeydown(event);
        return;
      }
      if (!elements.deleteConfirm.hidden) {
        closeDeleteConfirm();
        return;
      }
      if (!elements.importPreview.hidden) {
        closeImportPreview();
        return;
      }
      if (state.activeActionMenu) {
        closeActionMenu();
      }
    });
    root.addEventListener("scroll", closeActionMenuAfterScroll, true);
    root.addEventListener("resize", () => closeActionMenu({ restoreFocus: false }));
    elements.deleteConfirm.addEventListener("click", (event) => {
      if (event.target === elements.deleteConfirm) {
        closeDeleteConfirm();
      }
    });
    elements.importPreview.addEventListener("click", (event) => {
      if (event.target === elements.importPreview) {
        closeImportPreview();
      }
    });
    elements.cancelDelete.addEventListener("click", closeDeleteConfirm);
    elements.confirmDelete.addEventListener("click", confirmDeletePhrase);
  }

  function bindSettings() {
    elements.enabled.addEventListener("change", () => {
      state.config.enabled = elements.enabled.checked;
      saveConfig(state.config.enabled ? "页面工具栏已开启。" : "页面工具栏已关闭。");
    });
    elements.compactMode.addEventListener("change", () => {
      state.config.compactMode = elements.compactMode.checked;
      saveConfig("显示模式已更新。");
    });
    elements.insertMode.addEventListener("change", () => {
      state.config.insertMode = elements.insertMode.value;
      saveConfig("插入方式已更新。");
    });
    elements.columns.addEventListener("change", () => {
      state.config.columns = Number(elements.columns.value);
      saveConfig("每行按钮数已更新。");
    });
  }

  function getSelectedImportMode() {
    const selected = documentRef.querySelector("input[name='import-mode']:checked");
    return selected ? selected.value : "merge";
  }

  function setSelectedImportMode(mode) {
    const input = documentRef.querySelector(`input[name='import-mode'][value='${mode}']`);
    if (input) {
      input.checked = true;
    }
    updateImportModeUi();
  }

  function updateImportModeUi() {
    const mode = getSelectedImportMode();
    elements.importWarning.hidden = mode !== "replace";
    elements.confirmImport.textContent = mode === "replace" ? "确认替换" : "确认导入";
    if (mode === "replace" && state.pendingImport?.preview?.containsSettings) {
      elements.importSettings.checked = true;
    }
  }

  function createSummaryItem(label, value) {
    const item = documentRef.createElement("div");
    item.className = "summary-item";
    const strong = documentRef.createElement("strong");
    strong.textContent = label;
    const span = documentRef.createElement("span");
    span.textContent = String(value);
    item.append(strong, span);
    return item;
  }

  function renderImportPreview(preview) {
    elements.importSummary.innerHTML = "";
    const fileSize = preview.fileSize === null ? "未知" : `${Math.round(preview.fileSize / 1024)} KB`;
    elements.importSummary.append(
      createSummaryItem("文件名称", preview.fileName || "未命名文件"),
      createSummaryItem("文件大小", fileSize),
      createSummaryItem("导出时间", preview.exportedAt || "未提供"),
      createSummaryItem("格式版本", preview.legacy ? "旧版" : preview.formatVersion),
      createSummaryItem("文件词组数", preview.totalCount),
      createSummaryItem("有效词组", preview.validPhrases.length),
      createSummaryItem("无效词组", preview.invalidPhrases.length),
      createSummaryItem("重复词组", preview.duplicateCount),
      createSummaryItem("预计导入", preview.finalImportCount),
      createSummaryItem("包含设置", preview.containsSettings ? "是" : "否")
    );

    elements.phrasePreviewList.innerHTML = "";
    const rows = [
      ...preview.previewRows,
      ...preview.invalidPhrases.map((item) => ({
        name: `第 ${item.sourceIndex + 1} 个词组`,
        text: item.reason,
        valid: false,
        duplicate: false,
        reason: item.reason
      }))
    ].slice(0, 10);

    rows.forEach((row) => {
      const item = documentRef.createElement("article");
      item.className = `preview-row${row.duplicate ? " is-duplicate" : ""}`;
      const name = documentRef.createElement("div");
      name.className = "preview-name";
      name.textContent = row.name;
      const text = documentRef.createElement("div");
      text.className = "preview-text";
      text.textContent = row.text;
      const stateText = documentRef.createElement("div");
      stateText.className = "preview-state";
      stateText.textContent = row.valid ? (row.duplicate ? "重复，将跳过" : "有效") : `无效：${row.reason}`;
      item.append(name, text, stateText);
      elements.phrasePreviewList.appendChild(item);
    });

    const hiddenCount = Math.max(0, preview.normalizedPhrases.length + preview.invalidPhrases.length - 10);
    elements.previewMore.textContent = hiddenCount > 0 ? `还有 ${hiddenCount} 个词组未在预览中显示。` : (preview.legacyMessage || "");
    elements.importSettings.checked = false;
    elements.importSettings.disabled = !preview.containsSettings;
    setSelectedImportMode("merge");
    state.lastFocusedElement = documentRef.activeElement;
    elements.importPreview.hidden = false;
    elements.closeImportPreview.focus();
  }

  function closeImportPreview() {
    state.pendingImport = null;
    elements.importPreview.hidden = true;
    elements.importConfig.value = "";
    if (state.lastFocusedElement && typeof state.lastFocusedElement.focus === "function") {
      state.lastFocusedElement.focus();
    }
    state.lastFocusedElement = null;
  }

  async function writeConfigWithRollback(nextConfig, previousConfig) {
    await storageSet(nextConfig);
    const stored = await readLatestConfig();
    if (JSON.stringify(stored) !== JSON.stringify(configApi.normalizeConfig(nextConfig))) {
      await storageSet(previousConfig);
      throw new Error("写入浏览器存储后验证失败，已恢复原数据。");
    }
    return stored;
  }

  async function exportPhrases() {
    setBusy(true, "正在导出……");
    try {
      const latestConfig = await readLatestConfig();
      const exportData = exportApi.buildExportData(latestConfig);
      exportApi.downloadExportFile(exportData, { document: documentRef, URL: root.URL, Blob: root.Blob });
      showStatus("词组已成功导出。");
    } catch (error) {
      showStatus(error.message || "导出失败。");
    } finally {
      setBusy(false);
      renderBackupState();
    }
  }

  async function handleImportFile(file) {
    if (!file) {
      return;
    }
    setBusy(true, "正在读取文件……");
    try {
      const validation = await validateApi.readImportFile(file);
      const latestConfig = await readLatestConfig();
      const preview = importApi.buildImportPreview(validation, latestConfig);
      state.pendingImport = { validation, preview };
      renderImportPreview(preview);
      showStatus("请确认导入预览。", true);
    } catch (error) {
      showStatus(error.message || "导入失败。", true);
      elements.importConfig.value = "";
    } finally {
      setBusy(false);
      renderBackupState();
    }
  }

  async function confirmImport() {
    if (!state.pendingImport || state.busy) {
      return;
    }
    const mode = getSelectedImportMode();
    const includeSettings = elements.importSettings.checked;
    setBusy(true, "正在导入……");
    try {
      const latestConfig = await readLatestConfig();
      const preview = importApi.buildImportPreview(state.pendingImport.validation, latestConfig);
      const operation = mode === "replace"
        ? importApi.replaceImportedPhrases(latestConfig, preview, includeSettings)
        : importApi.mergeImportedPhrases(latestConfig, preview, includeSettings);
      await backupApi.createImportBackupInStorage(root.chrome.storage.local, latestConfig);
      state.hasBackup = true;
      const stored = await writeConfigWithRollback(operation.config, latestConfig);
      state.config = stored;
      resetForm();
      closeImportPreview();
      render();
      elements.undoImport.hidden = false;
      showStatus(importApi.summarizeImportResult(operation), true);
    } catch (error) {
      showStatus(error.message || "导入失败。", true);
    } finally {
      setBusy(false);
      renderBackupState();
    }
  }

  async function restoreBackup(message) {
    if (!state.hasBackup || state.busy) {
      showStatus("当前没有可恢复的导入备份。");
      return;
    }
    setBusy(true, "正在恢复备份……");
    try {
      const restored = await backupApi.restoreImportBackup(root.chrome.storage.local);
      state.config = restored;
      resetForm();
      render();
      showStatus(message || "已恢复导入前的数据。", true);
    } catch (error) {
      showStatus(error.message || "恢复失败。", true);
    } finally {
      setBusy(false);
      renderBackupState();
    }
  }

  function bindBackup() {
    elements.exportConfig.addEventListener("click", exportPhrases);
    elements.importTrigger.addEventListener("click", () => {
      if (!state.busy) {
        elements.importConfig.click();
      }
    });
    elements.importConfig.addEventListener("change", async () => {
      const file = elements.importConfig.files && elements.importConfig.files[0];
      await handleImportFile(file);
    });
    documentRef.querySelectorAll("input[name='import-mode']").forEach((input) => {
      input.addEventListener("change", updateImportModeUi);
    });
    elements.closeImportPreview.addEventListener("click", closeImportPreview);
    elements.cancelImport.addEventListener("click", closeImportPreview);
    elements.confirmImport.addEventListener("click", confirmImport);
    elements.restoreBackup.addEventListener("click", () => restoreBackup("已恢复最近导入前的数据。"));
    elements.undoImport.addEventListener("click", () => restoreBackup("已恢复导入前的数据。"));

    elements.resetDefaults.addEventListener("click", async () => {
      if (!root.confirm("恢复默认设置会覆盖当前词组和显示设置。继续？")) {
        return;
      }
      state.config = configApi.clone(configApi.DEFAULT_CONFIG);
      resetForm();
      await saveConfig("已恢复默认设置。");
    });
  }

  async function init() {
    const stored = await storageGet([configApi.STORAGE_KEY, configApi.IMPORT_BACKUP_KEY]);
    state.config = configApi.normalizeConfig(stored[configApi.STORAGE_KEY]);
    state.hasBackup = Boolean(stored[configApi.IMPORT_BACKUP_KEY]);
    decorateStaticIcons();
    bindForm();
    bindSettings();
    bindBackup();
    bindGlobalUi();
    render();
  }

  init();
})(typeof globalThis !== "undefined" ? globalThis : this);
