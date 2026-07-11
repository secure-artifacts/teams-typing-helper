(function attachContentScript(root) {
  "use strict";

  if (root.top !== root || !root.document) {
    return;
  }

  const configApi = root.TeamsTypingHelperConfig;
  if (!configApi) {
    return;
  }

  const INSERT_EVENT = "tth:insert:v1";
  const DIAGNOSE_EVENT = "tth:diagnose:v1";
  const RESULT_EVENT = "tth:result:v1";
  const INSERT_ATTR = "data-tth-insert-payload";
  const DIAGNOSE_ATTR = "data-tth-diagnose-payload";
  const RESULT_ATTR = "data-tth-result-payload";

  const state = {
    config: configApi.normalizeConfig(null),
    rootElement: null,
    statusElement: null,
    messageElement: null,
    observer: null,
    drag: null,
    debugOpen: false,
    statusTimer: null,
    requestCounter: 0
  };

  function hasChromeStorage() {
    return Boolean(root.chrome && root.chrome.storage && root.chrome.storage.local);
  }

  function storageGet() {
    return new Promise((resolve) => {
      if (!hasChromeStorage()) {
        resolve({});
        return;
      }
      root.chrome.storage.local.get(configApi.STORAGE_KEY, (result) => {
        resolve(result || {});
      });
    });
  }

  function storageSet(config) {
    return new Promise((resolve) => {
      if (!hasChromeStorage()) {
        resolve();
        return;
      }
      root.chrome.storage.local.set({ [configApi.STORAGE_KEY]: configApi.normalizeConfig(config) }, resolve);
    });
  }

  async function loadConfig() {
    const result = await storageGet();
    state.config = configApi.normalizeConfig(result[configApi.STORAGE_KEY]);
  }

  async function saveConfig(partial) {
    state.config = configApi.normalizeConfig({
      ...state.config,
      ...partial,
      toolbar: {
        ...state.config.toolbar,
        ...(partial.toolbar || {})
      }
    });
    await storageSet(state.config);
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  }

  function resultText(result) {
    const map = {
      INSERT_OK: "已填入，Teams 100ms 后仍保留文字。",
      INSERT_REVERTED: "普通编辑器插入验证未通过。请点“诊断输入框”并查看控制台。",
      WRONG_EDITOR: "疑似插入到了错误输入框。请点“诊断输入框”。",
      NO_EDITOR: "未找到消息输入框。请先点击 Teams 消息输入框，再点词组。",
      RANGE_INVALID: "保存的光标范围失效，插入未成功。请重新点击输入框。",
      EXEC_COMMAND_FAILED: "浏览器编辑命令和备用插入都未改变输入框内容。",
      EDITOR_IN_FRAME: "当前焦点可能在 iframe 中，请打开诊断日志确认。",
      DIAGNOSTIC_OK: "诊断结果已输出到页面控制台。",
      CKEDITOR_MODEL_INSERT_OK: "文字已通过 Teams 编辑器模型插入。",
      NO_CKEDITOR_INSTANCE: "检测到 Teams 编辑器，但无法取得 CKEditor 实例。",
      CKEDITOR_INSTANCE_MISSING: "检测到 Teams 编辑器，但无法取得 CKEditor 实例。",
      CKEDITOR_NOT_READY: "Teams 编辑器尚未准备完成，请稍后重试。",
      CKEDITOR_MODEL_NOT_CHANGED: "已调用 Teams 编辑器模型，但内容没有发生变化。",
      CKEDITOR_MODEL_EXCEPTION: result?.diagnostics?.insertionResult?.message || "Teams 编辑器模型插入发生异常。"
    };
    return result && (map[result.code] || result.message || result.code) || "";
  }

  function setStatus(message, ok) {
    if (state.statusElement) {
      state.statusElement.classList.toggle("tth-ok", ok === true);
      state.statusElement.classList.toggle("tth-error", ok === false);
      state.statusElement.title = message || "";
    }
    if (state.messageElement && message) {
      state.messageElement.textContent = message;
      state.messageElement.classList.toggle("tth-ok", ok === true);
      state.messageElement.classList.toggle("tth-error", ok === false);
      state.messageElement.hidden = false;
      clearTimeout(state.statusTimer);
      state.statusTimer = setTimeout(() => {
        if (state.messageElement) {
          state.messageElement.hidden = true;
        }
      }, ok === false ? 7000 : 2600);
    }
  }

  function nextRequestId(prefix) {
    state.requestCounter += 1;
    return `${prefix}-${Date.now()}-${state.requestCounter}`;
  }

  function dispatchBridgeEvent(attr, eventName, payload) {
    const target = root.document.documentElement;
    target.setAttribute(attr, JSON.stringify(payload));
    root.document.dispatchEvent(new Event(eventName, { bubbles: true, composed: true }));
  }

  function requestInsert(text, mode, source) {
    const payload = {
      id: nextRequestId("insert"),
      text: typeof text === "string" ? text : String(text || ""),
      mode: mode === "append" || mode === "replace" ? mode : "cursor",
      source: source || "phrase-button",
      requestedAt: Date.now()
    };
    setStatus("正在插入并验证...", null);
    dispatchBridgeEvent(INSERT_ATTR, INSERT_EVENT, payload);
  }

  function requestDiagnostic(source) {
    const payload = {
      id: nextRequestId("diagnose"),
      source: source || "toolbar",
      requestedAt: Date.now()
    };
    setStatus("正在输出诊断日志...", null);
    dispatchBridgeEvent(DIAGNOSE_ATTR, DIAGNOSE_EVENT, payload);
  }

  function handleBridgeResult() {
    const target = root.document.documentElement;
    const raw = target.getAttribute(RESULT_ATTR);
    target.removeAttribute(RESULT_ATTR);
    if (!raw) {
      return;
    }
    let result = null;
    try {
      result = JSON.parse(raw);
    } catch (error) {
      setStatus("无法读取插入结果。", false);
      return;
    }
    root.__teamsTypingHelperLastToolbarResult = result;
    setStatus(resultText(result), result.ok);
  }

  function applyPosition() {
    const element = state.rootElement;
    if (!element) {
      return;
    }

    const position = state.config.toolbar.position;
    if (position) {
      element.style.left = `${clamp(position.left, 8, root.innerWidth - 48)}px`;
      element.style.top = `${clamp(position.top, 8, root.innerHeight - 36)}px`;
      element.style.right = "auto";
      element.style.bottom = "auto";
      return;
    }

    element.style.right = "20px";
    element.style.left = "auto";
    element.style.bottom = "96px";
    element.style.top = "auto";
  }

  function createSvgIcon(name) {
    const svg = root.document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("focusable", "false");
    const path = root.document.createElementNS("http://www.w3.org/2000/svg", "path");
    const paths = {
      chevron: "m7.4 8.6 1.4-1.4 3.2 3.2 3.2-3.2 1.4 1.4-4.6 4.6-4.6-4.6Z",
      dots: "M12 8a2 2 0 1 0 0-4 2 2 0 0 0 0 4Zm0 2a2 2 0 1 0 0 4 2 2 0 0 0 0-4Zm0 6a2 2 0 1 0 0 4 2 2 0 0 0 0-4Z",
      bolt: "m13 2-8 12h6l-1 8 8-12h-6l1-8Z",
      test: "M11 3h2v6h6v2h-6v10h-2V11H5V9h6V3Z",
      search: "M10.5 4a6.5 6.5 0 0 1 5.2 10.4l4 4-1.4 1.4-4-4A6.5 6.5 0 1 1 10.5 4Zm0 2a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9Z"
    };
    path.setAttribute("d", paths[name] || paths.dots);
    path.setAttribute("fill", "currentColor");
    svg.appendChild(path);
    return svg;
  }

  function createButton(label, title, className) {
    const button = root.document.createElement("button");
    button.type = "button";
    button.className = className;
    button.textContent = label;
    button.title = title || label;
    button.addEventListener("mousedown", (event) => event.preventDefault());
    return button;
  }

  function bindPointerAction(button, action) {
    button.addEventListener("pointerdown", (event) => {
      if (event.button !== undefined && event.button !== 0) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      action();
    });
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
  }

  function renderPhrases(container) {
    container.innerHTML = "";
    container.style.setProperty("--tth-columns", String(state.config.columns));

    if (!state.config.phrases.length) {
      const empty = root.document.createElement("div");
      empty.className = "tth-empty";
      empty.textContent = "还没有词组，请在扩展弹窗中添加。";
      container.appendChild(empty);
      return;
    }

    state.config.phrases.forEach((phrase) => {
      const button = createButton(phrase.name, `${phrase.name}\n${phrase.text || ""}`.trim(), "tth-phrase-button");
      bindPointerAction(button, () => requestInsert(phrase.text, state.config.insertMode, "phrase-button"));
      container.appendChild(button);
    });
  }

  function startDrag(event) {
    if (state.drag || event.button !== 0 || (event.target.closest && event.target.closest("button"))) {
      return;
    }
    const rect = state.rootElement.getBoundingClientRect();
    state.drag = {
      pointerId: event.pointerId ?? "mouse",
      captureElement: event.currentTarget,
      dx: event.clientX - rect.left,
      dy: event.clientY - rect.top
    };
    if (event.pointerId !== undefined && event.currentTarget.setPointerCapture) {
      event.currentTarget.setPointerCapture(event.pointerId);
    }
    root.document.addEventListener("pointermove", moveDrag, true);
    root.document.addEventListener("pointerup", endDrag, true);
    root.document.addEventListener("pointercancel", endDrag, true);
    root.document.addEventListener("mousemove", moveDrag, true);
    root.document.addEventListener("mouseup", endDrag, true);
    event.preventDefault();
  }

  function moveDrag(event) {
    if (!state.drag) {
      return;
    }
    if (event.pointerId !== undefined && state.drag.pointerId !== event.pointerId) {
      return;
    }
    const rect = state.rootElement.getBoundingClientRect();
    const left = clamp(event.clientX - state.drag.dx, 8, root.innerWidth - rect.width - 8);
    const top = clamp(event.clientY - state.drag.dy, 8, root.innerHeight - rect.height - 8);
    state.rootElement.style.left = `${left}px`;
    state.rootElement.style.top = `${top}px`;
    state.rootElement.style.right = "auto";
    state.rootElement.style.bottom = "auto";
  }

  async function endDrag(event) {
    if (!state.drag) {
      return;
    }
    if (event.pointerId !== undefined && state.drag.pointerId !== event.pointerId) {
      return;
    }
    if (event.pointerId !== undefined && state.drag.captureElement && state.drag.captureElement.releasePointerCapture) {
      try {
        state.drag.captureElement.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Pointer capture may already be gone after a Teams re-render.
      }
    }
    root.document.removeEventListener("mousemove", moveDrag, true);
    root.document.removeEventListener("mouseup", endDrag, true);
    root.document.removeEventListener("pointermove", moveDrag, true);
    root.document.removeEventListener("pointerup", endDrag, true);
    root.document.removeEventListener("pointercancel", endDrag, true);
    state.drag = null;
    const rect = state.rootElement.getBoundingClientRect();
    await saveConfig({ toolbar: { position: { left: rect.left, top: rect.top } } });
  }

  function buildToolbar() {
    const element = root.document.createElement("section");
    element.id = "teams-typing-helper-toolbar";
    element.className = "tth-root";
    element.dataset.tthIgnore = "true";
    element.setAttribute("aria-label", "Teams 辅助打字词组工具栏");

    const header = root.document.createElement("div");
    header.className = "tth-header";
    header.title = "拖动可移动位置，双击可恢复默认位置";

    const toggleButton = createButton("", "折叠或展开工具栏", "tth-icon-button tth-toggle-button");
    toggleButton.setAttribute("aria-label", "折叠或展开工具栏");
    toggleButton.setAttribute("aria-expanded", String(!state.config.toolbar.collapsed));
    const chevron = createSvgIcon("chevron");
    chevron.classList.add("tth-chevron");
    toggleButton.appendChild(chevron);
    toggleButton.addEventListener("click", async () => {
      await saveConfig({ toolbar: { collapsed: !state.config.toolbar.collapsed } });
      renderToolbar();
    });

    const appIcon = root.document.createElement("span");
    appIcon.className = "tth-app-icon";
    appIcon.setAttribute("aria-hidden", "true");
    appIcon.appendChild(createSvgIcon("bolt"));

    const title = root.document.createElement("div");
    title.className = "tth-title";
    title.textContent = "Teams 辅助打字";

    const moreButton = createButton("", "打开诊断工具", "tth-icon-button tth-more-button");
    moreButton.setAttribute("aria-label", "打开诊断工具");
    moreButton.setAttribute("aria-expanded", String(state.debugOpen));
    moreButton.appendChild(createSvgIcon("dots"));

    const status = root.document.createElement("span");
    status.className = "tth-status";
    status.title = "等待识别 Teams 输入框";
    state.statusElement = status;

    header.append(toggleButton, appIcon, title, moreButton, status);
    header.addEventListener("pointerdown", startDrag);
    header.addEventListener("pointermove", moveDrag);
    header.addEventListener("pointerup", endDrag);
    header.addEventListener("pointercancel", endDrag);
    header.addEventListener("mousedown", startDrag);
    header.addEventListener("dblclick", async () => {
      await saveConfig({ toolbar: { position: null } });
      applyPosition();
    });

    const body = root.document.createElement("div");
    body.className = "tth-body";

    const message = root.document.createElement("div");
    message.className = "tth-toast";
    message.hidden = true;
    state.messageElement = message;

    const tools = root.document.createElement("div");
    tools.className = "tth-tools";
    tools.hidden = !state.debugOpen;
    const toolsTitle = root.document.createElement("div");
    toolsTitle.className = "tth-tools-title";
    toolsTitle.textContent = "诊断工具";
    const toolButtons = root.document.createElement("div");
    toolButtons.className = "tth-tool-buttons";
    const testButton = createButton("测试插入：你好", "向当前 Teams 输入框插入“你好”", "tth-tool-button tth-test-button");
    testButton.prepend(createSvgIcon("test"));
    bindPointerAction(testButton, () => requestInsert("你好", "cursor", "test-button"));
    const diagnoseButton = createButton("诊断输入框", "在页面控制台输出输入框诊断信息", "tth-tool-button tth-diagnose-button");
    diagnoseButton.prepend(createSvgIcon("search"));
    bindPointerAction(diagnoseButton, () => requestDiagnostic("diagnose-button"));
    toolButtons.append(testButton, diagnoseButton);
    tools.append(toolsTitle, toolButtons);
    moreButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      state.debugOpen = !state.debugOpen;
      tools.hidden = !state.debugOpen;
      moreButton.setAttribute("aria-expanded", String(state.debugOpen));
    });

    const buttons = root.document.createElement("div");
    buttons.className = "tth-buttons";
    renderPhrases(buttons);

    body.append(message, tools, buttons);
    element.append(header, body);
    return element;
  }

  function removeToolbar() {
    if (state.rootElement) {
      state.rootElement.remove();
      state.rootElement = null;
      state.statusElement = null;
      state.messageElement = null;
    }
  }

  function renderToolbar() {
    if (!state.config.enabled) {
      removeToolbar();
      return;
    }

    const oldElement = state.rootElement;
    const element = buildToolbar();
    element.classList.toggle("tth-collapsed", state.config.toolbar.collapsed);
    element.classList.toggle("tth-compact", state.config.compactMode);

    if (oldElement) {
      oldElement.replaceWith(element);
    } else {
      (root.document.body || root.document.documentElement).appendChild(element);
    }
    state.rootElement = element;
    applyPosition();
  }

  function debounce(fn, delay) {
    let timer = null;
    return function debounced() {
      clearTimeout(timer);
      timer = setTimeout(fn, delay);
    };
  }

  function observePage() {
    if (!root.MutationObserver || state.observer) {
      return;
    }
    const refresh = debounce(() => {
      if (state.config.enabled && state.rootElement && !root.document.documentElement.contains(state.rootElement)) {
        renderToolbar();
      }
      if (state.config.enabled && !state.config.toolbar.position) {
        applyPosition();
      }
    }, 200);
    state.observer = new root.MutationObserver(refresh);
    state.observer.observe(root.document.documentElement, { childList: true, subtree: true });
  }

  function bindStorageChanges() {
    if (!root.chrome || !root.chrome.storage || !root.chrome.storage.onChanged) {
      return;
    }
    root.chrome.storage.onChanged.addListener((changes, areaName) => {
      if (areaName !== "local" || !changes[configApi.STORAGE_KEY]) {
        return;
      }
      state.config = configApi.normalizeConfig(changes[configApi.STORAGE_KEY].newValue);
      renderToolbar();
    });
  }

  async function init() {
    await loadConfig();
    renderToolbar();
    observePage();
    bindStorageChanges();
    root.document.addEventListener(RESULT_EVENT, handleBridgeResult, true);
    root.document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && state.debugOpen) {
        state.debugOpen = false;
        renderToolbar();
      }
    }, true);
    root.addEventListener("resize", debounce(() => {
      if (!state.config.toolbar.position) {
        applyPosition();
      }
    }, 150));
  }

  init();
})(typeof globalThis !== "undefined" ? globalThis : this);
