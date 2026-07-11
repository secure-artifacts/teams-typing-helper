(function attachMainBridge(root) {
  "use strict";

  const INSERT_EVENT = "tth:insert:v1";
  const DIAGNOSE_EVENT = "tth:diagnose:v1";
  const RESULT_EVENT = "tth:result:v1";
  const INSERT_ATTR = "data-tth-insert-payload";
  const DIAGNOSE_ATTR = "data-tth-diagnose-payload";
  const RESULT_ATTR = "data-tth-result-payload";
  const INTERNAL_SELECTOR = "[data-tth-ignore='true']";
  const EDITABLE_SELECTOR = [
    "textarea",
    "input",
    "[contenteditable='true']",
    "[contenteditable='plaintext-only']",
    "[role='textbox']"
  ].join(",");
  const TEXT_INPUT_TYPES = new Set(["", "text", "search", "email", "url", "tel"]);
  const POSITIVE_LABEL_RE = /(message|chat|reply|compose|type|send|new post|announcement|post|输入|消息|聊天|回复|撰写|键入|发送|公告)/i;
  const NEGATIVE_LABEL_RE = /(search|filter|find|contact|people|settings|command|查找|搜索|筛选|联系人|设置|命令)/i;

  const state = {
    lastEditor: null,
    lastRange: null,
    lastEditorDocument: null,
    lastUpdatedAt: 0,
    lastDiagnostics: null
  };

  function getDoc() {
    return root.document;
  }

  function isElement(node) {
    return Boolean(node && node.nodeType === Node.ELEMENT_NODE);
  }

  function isTextNode(node) {
    return Boolean(node && node.nodeType === Node.TEXT_NODE);
  }

  function getElementFromNode(node) {
    if (!node) {
      return null;
    }
    if (isElement(node)) {
      return node;
    }
    if (isTextNode(node) && node.parentElement) {
      return node.parentElement;
    }
    return node.parentElement || null;
  }

  function safeAttr(element, name) {
    if (!isElement(element)) {
      return "";
    }
    return element.getAttribute(name) || "";
  }

  function getNodeRoot(node) {
    return node && node.getRootNode ? node.getRootNode() : null;
  }

  function getParentAcrossRoots(node) {
    if (!node) {
      return null;
    }
    if (node.parentElement) {
      return node.parentElement;
    }
    if (node.parentNode && node.parentNode.nodeType !== Node.DOCUMENT_FRAGMENT_NODE) {
      return node.parentNode;
    }
    const rootNode = getNodeRoot(node);
    return rootNode && rootNode.host ? rootNode.host : null;
  }

  function closestAcrossRoots(node, predicate) {
    let current = getElementFromNode(node);
    while (current) {
      if (isElement(current) && predicate(current)) {
        return current;
      }
      current = getParentAcrossRoots(current);
    }
    return null;
  }

  function containsNode(container, node) {
    if (!container || !node) {
      return false;
    }
    if (container === node) {
      return true;
    }
    const directNode = isTextNode(node) ? node.parentNode : node;
    if (directNode && container.contains && container.contains(directNode)) {
      return true;
    }

    let current = directNode;
    while (current) {
      if (current === container) {
        return true;
      }
      current = getParentAcrossRoots(current);
    }
    return false;
  }

  function isInsideToolbar(node) {
    return Boolean(closestAcrossRoots(node, (element) => element.matches && element.matches(INTERNAL_SELECTOR)));
  }

  function isTextInput(element) {
    if (!isElement(element)) {
      return false;
    }
    const tagName = element.tagName.toLowerCase();
    if (tagName === "textarea") {
      return true;
    }
    if (tagName !== "input") {
      return false;
    }
    return TEXT_INPUT_TYPES.has((element.getAttribute("type") || element.type || "").toLowerCase());
  }

  function isContentEditableElement(element) {
    if (!isElement(element)) {
      return false;
    }
    const attr = element.getAttribute("contenteditable");
    if (attr === null) {
      return false;
    }
    const normalized = attr.toLowerCase();
    return normalized === "" || normalized === "true" || normalized === "plaintext-only";
  }

  function isEditableElement(element) {
    return isTextInput(element) || isContentEditableElement(element);
  }

  function isCkEditorElement(element) {
    return Boolean(
      isElement(element) &&
      (
        element.ckeditorInstance ||
        element.classList?.contains("ck-editor__editable") ||
        element.classList?.contains("ck-content") ||
        safeAttr(element, "data-tid") === "ckeditor"
      )
    );
  }

  function getCkEditorInstance(element) {
    return element && element.ckeditorInstance ? element.ckeditorInstance : null;
  }

  function closestEditable(node) {
    return closestAcrossRoots(node, (element) => isEditableElement(element));
  }

  function isDisabledOrReadonly(element) {
    return Boolean(
      element.disabled ||
      element.readOnly ||
      safeAttr(element, "aria-disabled") === "true" ||
      safeAttr(element, "readonly")
    );
  }

  function isAriaHidden(element) {
    return Boolean(closestAcrossRoots(element, (node) => safeAttr(node, "aria-hidden") === "true"));
  }

  function getWindowForElement(element) {
    return element && element.ownerDocument && element.ownerDocument.defaultView
      ? element.ownerDocument.defaultView
      : root;
  }

  function isVisible(element) {
    if (!isElement(element) || !element.isConnected || element.hidden || isAriaHidden(element)) {
      return false;
    }
    const win = getWindowForElement(element);
    const style = win.getComputedStyle ? win.getComputedStyle(element) : null;
    if (style && (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0)) {
      return false;
    }
    const rect = element.getBoundingClientRect ? element.getBoundingClientRect() : null;
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      return false;
    }
    const viewportWidth = win.innerWidth || root.innerWidth || 0;
    const viewportHeight = win.innerHeight || root.innerHeight || 0;
    if (viewportWidth && (rect.right < 0 || rect.left > viewportWidth)) {
      return false;
    }
    if (viewportHeight && (rect.bottom < 0 || rect.top > viewportHeight + 80)) {
      return false;
    }
    return true;
  }

  function getLabel(element) {
    return [
      safeAttr(element, "aria-label"),
      safeAttr(element, "aria-placeholder"),
      safeAttr(element, "placeholder"),
      safeAttr(element, "title"),
      safeAttr(element, "data-tid"),
      safeAttr(element, "data-testid"),
      safeAttr(element, "name")
    ].filter(Boolean).join(" ");
  }

  function isExcluded(element) {
    const label = getLabel(element);
    return (
      !isEditableElement(element) ||
      isInsideToolbar(element) ||
      isDisabledOrReadonly(element) ||
      !isVisible(element) ||
      NEGATIVE_LABEL_RE.test(label)
    );
  }

  function getSelection(doc) {
    try {
      return doc && doc.getSelection ? doc.getSelection() : null;
    } catch (error) {
      return null;
    }
  }

  function getSelectionEditor(doc) {
    const selection = getSelection(doc);
    if (!selection || selection.rangeCount === 0) {
      return null;
    }
    return closestEditable(selection.anchorNode) || closestEditable(selection.focusNode);
  }

  function getActiveEditor(doc) {
    return closestEditable(doc && doc.activeElement);
  }

  function collectElementsDeep(rootNode, selector, output) {
    if (!rootNode || !rootNode.querySelectorAll) {
      return;
    }
    rootNode.querySelectorAll(selector).forEach((element) => output.add(element));
    rootNode.querySelectorAll("*").forEach((element) => {
      if (element.shadowRoot) {
        collectElementsDeep(element.shadowRoot, selector, output);
      }
    });
  }

  function collectCandidateElements(doc) {
    const output = new Set();
    const selectionEditor = getSelectionEditor(doc);
    const activeEditor = getActiveEditor(doc);
    if (selectionEditor) {
      output.add(selectionEditor);
    }
    if (activeEditor) {
      output.add(activeEditor);
    }
    if (state.lastEditor && state.lastEditor.isConnected) {
      output.add(state.lastEditor);
    }
    collectElementsDeep(doc, EDITABLE_SELECTOR, output);
    return Array.from(output).filter((element) => isEditableElement(element));
  }

  function rectSummary(element) {
    const rect = element && element.getBoundingClientRect ? element.getBoundingClientRect() : null;
    if (!rect) {
      return null;
    }
    return {
      x: Math.round(rect.x),
      y: Math.round(rect.y),
      width: Math.round(rect.width),
      height: Math.round(rect.height),
      top: Math.round(rect.top),
      bottom: Math.round(rect.bottom)
    };
  }

  function scoreCandidate(element, doc) {
    const activeEditor = getActiveEditor(doc);
    const selectionEditor = getSelectionEditor(doc);
    const label = getLabel(element);
    const rect = element.getBoundingClientRect();
    const win = getWindowForElement(element);
    let score = 0;

    if (element === selectionEditor) {
      score += 1200;
    }
    if (element === activeEditor || containsNode(element, doc.activeElement)) {
      score += 1000;
    }
    if (element === state.lastEditor) {
      score += 800;
    }
    if (safeAttr(element, "role") === "textbox") {
      score += 140;
    }
    if (isContentEditableElement(element)) {
      score += 130;
    }
    if (POSITIVE_LABEL_RE.test(label)) {
      score += 100;
    }
    if (NEGATIVE_LABEL_RE.test(label)) {
      score -= 900;
    }
    if (rect.width >= 180 && rect.height >= 20) {
      score += 30;
    }
    if ((win.innerHeight || 0) && rect.top > win.innerHeight * 0.45) {
      score += 40;
    }
    score += Math.max(0, Math.min(30, Math.round(((win.innerHeight || 0) - rect.bottom) / -10)));
    return score;
  }

  function candidateRow(element, doc, selectedElement) {
    const rect = rectSummary(element);
    const win = getWindowForElement(element);
    return {
      tagName: element.tagName,
      role: safeAttr(element, "role"),
      contenteditable: safeAttr(element, "contenteditable"),
      ariaLabel: safeAttr(element, "aria-label"),
      dataTid: safeAttr(element, "data-tid") || safeAttr(element, "data-testid"),
      className: typeof element.className === "string" ? element.className.slice(0, 160) : "",
      boundingClientRect: rect,
      visible: isVisible(element),
      insideToolbar: isInsideToolbar(element),
      distanceFromBottom: rect && win.innerHeight ? Math.round(win.innerHeight - rect.bottom) : null,
      score: isExcluded(element) ? -9999 : scoreCandidate(element, doc),
      selected: element === selectedElement
    };
  }

  function findBestEditor() {
    const doc = getDoc();
    const selectionEditor = getSelectionEditor(doc);
    if (selectionEditor && !isExcluded(selectionEditor)) {
      const candidates = collectCandidateElements(doc);
      return buildFindResult(selectionEditor, candidates);
    }

    const activeEditor = getActiveEditor(doc);
    if (activeEditor && !isExcluded(activeEditor)) {
      const candidates = collectCandidateElements(doc);
      return buildFindResult(activeEditor, candidates);
    }

    if (state.lastEditor && state.lastEditor.isConnected && !isExcluded(state.lastEditor)) {
      const candidates = collectCandidateElements(doc);
      return buildFindResult(state.lastEditor, candidates);
    }

    const candidates = collectCandidateElements(doc);
    const scored = candidates
      .filter((element) => !isExcluded(element))
      .map((element) => ({ element, score: scoreCandidate(element, doc) }))
      .sort((a, b) => b.score - a.score);
    return buildFindResult(scored.length ? scored[0].element : null, candidates);
  }

  function buildFindResult(editor, candidates) {
    const doc = getDoc();
    return {
      editor,
      candidates,
      rows: candidates.map((candidate) => candidateRow(candidate, doc, editor))
    };
  }

  function snapshotEditor(editor) {
    if (!editor) {
      return { text: "", value: "", html: "" };
    }
    return {
      text: editor.innerText || editor.textContent || "",
      value: isTextInput(editor) ? editor.value || "" : "",
      html: isContentEditableElement(editor) ? (editor.innerHTML || "").slice(0, 500) : ""
    };
  }

  function getCkEditorData(editorInstance) {
    if (!editorInstance || typeof editorInstance.getData !== "function") {
      return null;
    }
    try {
      return editorInstance.getData();
    } catch (error) {
      return null;
    }
  }

  function getCkEditorSelectionInfo(editorInstance) {
    const modelSelection = editorInstance?.model?.document?.selection;
    const firstPosition = modelSelection?.getFirstPosition?.();
    return firstPosition
      ? {
        rootName: firstPosition.root?.rootName || null,
        path: Array.from(firstPosition.path || []),
        isCollapsed: Boolean(modelSelection.isCollapsed)
      }
      : null;
  }

  function getCkEditorInfo(domEditable) {
    const editorInstance = getCkEditorInstance(domEditable);
    return {
      isCkEditorElement: isCkEditorElement(domEditable),
      hasCkeditorInstance: Boolean(editorInstance),
      ckeditorState: editorInstance?.state || null,
      hasModel: Boolean(editorInstance?.model),
      hasEditingView: Boolean(editorInstance?.editing?.view),
      hasGetData: typeof editorInstance?.getData === "function",
      beforeEditorData: getCkEditorData(editorInstance),
      modelSelectionPosition: getCkEditorSelectionInfo(editorInstance),
      modelSelectionIsCollapsed: Boolean(editorInstance?.model?.document?.selection?.isCollapsed)
    };
  }

  function logCkEditorInstance(domEditable) {
    const editorInstance = getCkEditorInstance(domEditable);
    console.log("[Teams Typing Helper] CKEditor instance", {
      exists: Boolean(editorInstance),
      state: editorInstance?.state,
      hasModel: Boolean(editorInstance?.model),
      hasEditing: Boolean(editorInstance?.editing),
      hasGetData: typeof editorInstance?.getData === "function"
    });
  }

  function snapshotChanged(before, after, insertedText) {
    if (before.text !== after.text || before.value !== after.value || before.html !== after.html) {
      return true;
    }
    if (insertedText && after.text.includes(insertedText) && !before.text.includes(insertedText)) {
      return true;
    }
    if (insertedText && after.value.includes(insertedText) && !before.value.includes(insertedText)) {
      return true;
    }
    return false;
  }

  function contentContainsText(value, text) {
    if (!text) {
      return true;
    }
    return typeof value === "string" && value.includes(text);
  }

  function getSnapshotContent(snapshot) {
    return snapshot.value || snapshot.text || "";
  }

  function createInputEvent(editor, inputType, text) {
    const win = getWindowForElement(editor);
    if (typeof win.InputEvent === "function") {
      try {
        return new win.InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType,
          data: text
        });
      } catch (error) {
        return new win.Event("input", { bubbles: true, composed: true });
      }
    }
    return new win.Event("input", { bubbles: true, composed: true });
  }

  function dispatchInput(editor, inputType, text) {
    editor.dispatchEvent(createInputEvent(editor, inputType, text));
    editor.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
  }

  function focusEditor(editor) {
    if (!editor || typeof editor.focus !== "function") {
      return;
    }
    try {
      editor.focus({ preventScroll: true });
    } catch (error) {
      editor.focus();
    }
  }

  function setSelectionToEnd(editor) {
    const doc = editor.ownerDocument;
    const selection = getSelection(doc);
    if (!selection || !doc.createRange) {
      return { ok: false, reason: "no-selection" };
    }
    const range = doc.createRange();
    range.selectNodeContents(editor);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
    return { ok: true, range };
  }

  function selectEditorContents(editor) {
    const doc = editor.ownerDocument;
    const selection = getSelection(doc);
    if (!selection || !doc.createRange) {
      return { ok: false, reason: "no-selection" };
    }
    const range = doc.createRange();
    range.selectNodeContents(editor);
    selection.removeAllRanges();
    selection.addRange(range);
    return { ok: true, range };
  }

  function currentSelectionInside(editor) {
    const selection = getSelection(editor.ownerDocument);
    return Boolean(selection && selection.rangeCount > 0 && containsNode(editor, selection.anchorNode));
  }

  function restoreContentSelection(editor, mode) {
    if (mode === "append") {
      return { ...setSelectionToEnd(editor), rangeInvalid: false };
    }
    if (mode === "replace") {
      return { ...selectEditorContents(editor), rangeInvalid: false };
    }

    const selection = getSelection(editor.ownerDocument);
    if (state.lastRange && state.lastEditor === editor && state.lastEditorDocument === editor.ownerDocument) {
      try {
        if (containsNode(editor, state.lastRange.startContainer) && containsNode(editor, state.lastRange.endContainer)) {
          selection.removeAllRanges();
          selection.addRange(state.lastRange.cloneRange());
          return { ok: true, rangeInvalid: false };
        }
      } catch (error) {
        return { ...setSelectionToEnd(editor), rangeInvalid: true };
      }
      return { ...setSelectionToEnd(editor), rangeInvalid: true };
    }

    if (currentSelectionInside(editor)) {
      return { ok: true, rangeInvalid: false };
    }
    return { ...setSelectionToEnd(editor), rangeInvalid: false };
  }

  function fallbackRangeInsert(editor, text) {
    const doc = editor.ownerDocument;
    const selection = getSelection(doc);
    if (!selection) {
      return false;
    }
    if (!selection.rangeCount || !containsNode(editor, selection.anchorNode)) {
      setSelectionToEnd(editor);
    }
    const range = selection.rangeCount ? selection.getRangeAt(0) : null;
    if (!range) {
      return false;
    }
    range.deleteContents();
    const textNode = doc.createTextNode(text);
    range.insertNode(textNode);
    range.setStartAfter(textNode);
    range.setEndAfter(textNode);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
    return true;
  }

  function getNativeValueSetter(editor) {
    const win = getWindowForElement(editor);
    const tagName = editor.tagName.toLowerCase();
    const prototype = tagName === "textarea" ? win.HTMLTextAreaElement.prototype : win.HTMLInputElement.prototype;
    const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
    return descriptor && descriptor.set;
  }

  function insertIntoTextInput(editor, text, mode) {
    focusEditor(editor);
    const before = snapshotEditor(editor);
    const original = editor.value || "";
    let start = Number.isFinite(editor.selectionStart) ? editor.selectionStart : original.length;
    let end = Number.isFinite(editor.selectionEnd) ? editor.selectionEnd : start;
    if (mode === "append") {
      start = original.length;
      end = original.length;
    } else if (mode === "replace") {
      start = 0;
      end = original.length;
    }

    const nextValue = `${original.slice(0, start)}${text}${original.slice(end)}`;
    const setter = getNativeValueSetter(editor);
    if (setter) {
      setter.call(editor, nextValue);
    } else {
      editor.value = nextValue;
    }
    if (editor.setSelectionRange) {
      const caret = start + text.length;
      editor.setSelectionRange(caret, caret);
    }
    dispatchInput(editor, mode === "replace" ? "insertReplacementText" : "insertText", text);
    const immediate = snapshotEditor(editor);
    return {
      before,
      immediate,
      execCommandRan: false,
      fallbackRan: false,
      insertedNow: snapshotChanged(before, immediate, text),
      rangeInvalid: false
    };
  }

  function insertViaCkEditor5(domEditable, text) {
    const editorInstance = getCkEditorInstance(domEditable);
    const before = snapshotEditor(domEditable);
    logCkEditorInstance(domEditable);

    if (!editorInstance) {
      return {
        ok: false,
        code: "NO_CKEDITOR_INSTANCE",
        insertionMethod: "CKEDITOR_MODEL",
        before,
        immediate: before,
        beforeEditorData: null,
        afterEditorData: null,
        modelSelectionPositionBefore: null,
        modelSelectionPositionAfter: null,
        insertedNow: false
      };
    }

    if (editorInstance.state && editorInstance.state !== "ready") {
      return {
        ok: false,
        code: "CKEDITOR_NOT_READY",
        state: editorInstance.state,
        insertionMethod: "CKEDITOR_MODEL",
        before,
        immediate: before,
        beforeEditorData: getCkEditorData(editorInstance),
        afterEditorData: getCkEditorData(editorInstance),
        modelSelectionPositionBefore: getCkEditorSelectionInfo(editorInstance),
        modelSelectionPositionAfter: getCkEditorSelectionInfo(editorInstance),
        insertedNow: false
      };
    }

    const beforeEditorData = getCkEditorData(editorInstance);
    const modelSelectionPositionBefore = getCkEditorSelectionInfo(editorInstance);

    try {
      editorInstance.editing?.view?.focus?.();

      editorInstance.model.change((writer) => {
        const selection = editorInstance.model.document.selection;
        const attributes = Object.fromEntries(selection.getAttributes());
        const textNode = writer.createText(String(text), attributes);
        editorInstance.model.insertContent(textNode);
      });

      editorInstance.editing?.view?.focus?.();

      const afterEditorData = getCkEditorData(editorInstance);
      const immediate = snapshotEditor(domEditable);
      const changed = beforeEditorData === null || afterEditorData === null || beforeEditorData !== afterEditorData;

      return {
        ok: changed,
        code: changed ? "CKEDITOR_MODEL_INSERT_OK" : "CKEDITOR_MODEL_NOT_CHANGED",
        insertionMethod: "CKEDITOR_MODEL",
        before,
        immediate,
        beforeEditorData,
        afterEditorData,
        modelSelectionPositionBefore,
        modelSelectionPositionAfter: getCkEditorSelectionInfo(editorInstance),
        insertedNow: changed
      };
    } catch (error) {
      console.error("[Teams Typing Helper] CKEditor model insertion failed", error);
      return {
        ok: false,
        code: "CKEDITOR_MODEL_EXCEPTION",
        message: error?.message || String(error),
        insertionMethod: "CKEDITOR_MODEL",
        before,
        immediate: snapshotEditor(domEditable),
        beforeEditorData,
        afterEditorData: getCkEditorData(editorInstance),
        modelSelectionPositionBefore,
        modelSelectionPositionAfter: getCkEditorSelectionInfo(editorInstance),
        insertedNow: false
      };
    }
  }

  function insertIntoContentEditable(editor, text, mode) {
    const before = snapshotEditor(editor);
    focusEditor(editor);
    const selectionState = restoreContentSelection(editor, mode);
    let execCommandRan = false;
    let execCommandResult = false;

    try {
      if (editor.ownerDocument && typeof editor.ownerDocument.execCommand === "function") {
        execCommandRan = true;
        execCommandResult = editor.ownerDocument.execCommand("insertText", false, text);
      }
    } catch (error) {
      execCommandResult = false;
    }

    let afterExec = snapshotEditor(editor);
    let fallbackRan = false;
    if (!snapshotChanged(before, afterExec, text)) {
      fallbackRan = fallbackRangeInsert(editor, text);
      if (fallbackRan) {
        dispatchInput(editor, mode === "replace" ? "insertReplacementText" : "insertText", text);
      }
    } else {
      dispatchInput(editor, mode === "replace" ? "insertReplacementText" : "insertText", text);
    }

    const immediate = snapshotEditor(editor);
    return {
      before,
      immediate,
      execCommandRan,
      execCommandResult,
      fallbackRan,
      insertedNow: snapshotChanged(before, immediate, text),
      rangeInvalid: Boolean(selectionState.rangeInvalid)
    };
  }

  function updateMemoryFromEditor(editor) {
    if (!editor || !isEditableElement(editor) || !editor.isConnected || isInsideToolbar(editor)) {
      return;
    }
    state.lastEditor = editor;
    state.lastEditorDocument = editor.ownerDocument;
    state.lastUpdatedAt = Date.now();
    if (isContentEditableElement(editor)) {
      const selection = getSelection(editor.ownerDocument);
      if (selection && selection.rangeCount > 0 && containsNode(editor, selection.anchorNode)) {
        try {
          state.lastRange = selection.getRangeAt(0).cloneRange();
        } catch (error) {
          state.lastRange = null;
        }
      }
    }
  }

  function rememberFromEvent(event) {
    if (event && isInsideToolbar(event.target)) {
      return;
    }
    const doc = getDoc();
    const selectionEditor = getSelectionEditor(doc);
    const eventEditor = event ? closestEditable(event.target) : null;
    const editor = selectionEditor || eventEditor || getActiveEditor(doc);
    if (editor && !isExcluded(editor)) {
      updateMemoryFromEditor(editor);
    }
  }

  function buildDiagnostics(selectedEditor, rows, extra) {
    const doc = getDoc();
    const selection = getSelection(doc);
    const active = doc.activeElement;
    const selectedRoot = selectedEditor ? getNodeRoot(selectedEditor) : null;
    const inShadowRoot = Boolean(selectedRoot && selectedRoot.host);
    const activeFrame = active && active.tagName === "IFRAME";
    const ckEditorInfo = getCkEditorInfo(selectedEditor);
    return {
      activeElement: active ? {
        tagName: active.tagName,
        role: safeAttr(active, "role"),
        ariaLabel: safeAttr(active, "aria-label"),
        dataTid: safeAttr(active, "data-tid"),
        className: typeof active.className === "string" ? active.className.slice(0, 160) : ""
      } : null,
      selection: selection ? {
        rangeCount: selection.rangeCount,
        anchorNode: selection.anchorNode ? selection.anchorNode.nodeName : null,
        focusNode: selection.focusNode ? selection.focusNode.nodeName : null,
        text: String(selection).slice(0, 120)
      } : null,
      counts: {
        contenteditable: doc.querySelectorAll("[contenteditable='true'],[contenteditable='plaintext-only']").length,
        roleTextbox: doc.querySelectorAll("[role='textbox']").length,
        candidates: rows.length
      },
      selectedEditor: selectedEditor ? candidateRow(selectedEditor, doc, selectedEditor) : null,
      ownerDocumentLocation: selectedEditor && selectedEditor.ownerDocument ? String(selectedEditor.ownerDocument.location) : String(doc.location),
      editorInFrame: root.top !== root || activeFrame,
      editorInShadowRoot: inShadowRoot,
      isCkEditorElement: ckEditorInfo.isCkEditorElement,
      hasCkeditorInstance: ckEditorInfo.hasCkeditorInstance,
      ckeditorState: ckEditorInfo.ckeditorState,
      hasModel: ckEditorInfo.hasModel,
      hasEditingView: ckEditorInfo.hasEditingView,
      beforeEditorData: extra?.beforeEditorData ?? ckEditorInfo.beforeEditorData,
      afterEditorData: extra?.afterEditorData ?? null,
      afterDelayEditorData: extra?.afterDelayEditorData ?? null,
      modelSelectionPosition: extra?.modelSelectionPositionAfter ?? ckEditorInfo.modelSelectionPosition,
      modelSelectionIsCollapsed: ckEditorInfo.modelSelectionIsCollapsed,
      insertionMethod: extra?.insertionMethod ?? null,
      insertionResult: extra?.insertionResult ?? null,
      lastEditor: state.lastEditor ? candidateRow(state.lastEditor, doc, selectedEditor) : null,
      lastRange: state.lastRange ? {
        startContainer: state.lastRange.startContainer ? state.lastRange.startContainer.nodeName : null,
        endContainer: state.lastRange.endContainer ? state.lastRange.endContainer.nodeName : null,
        collapsed: state.lastRange.collapsed
      } : null,
      lastUpdatedAt: state.lastUpdatedAt,
      candidates: rows,
      ...extra
    };
  }

  function emitResult(result) {
    const payload = {
      id: result.id || "",
      ok: Boolean(result.ok),
      code: result.code || "UNKNOWN",
      message: result.message || "",
      timestamp: Date.now(),
      frame: {
        isTop: root.top === root,
        location: String(getDoc().location)
      },
      diagnostics: result.diagnostics || null
    };
    root.__teamsTypingHelperLastResult = payload;
    getDoc().documentElement.setAttribute(RESULT_ATTR, JSON.stringify(payload));
    getDoc().dispatchEvent(new Event(RESULT_EVENT, { bubbles: true, composed: true }));
  }

  function finalizeInsertResult(payload, editor, rows, insertResult) {
    const afterDelay = snapshotEditor(editor);
    const ckEditorInstance = getCkEditorInstance(editor);
    const afterDelayEditorData = getCkEditorData(ckEditorInstance);
    const isCkEditorMethod = insertResult.insertionMethod === "CKEDITOR_MODEL";
    const stillInserted = isCkEditorMethod
      ? (
        contentContainsText(afterDelayEditorData, payload.text) ||
        contentContainsText(afterDelay.text, payload.text) ||
        contentContainsText(afterDelay.html, payload.text)
      )
      : snapshotChanged(insertResult.before, afterDelay, payload.text);
    const code = isCkEditorMethod
      ? (
        insertResult.ok && stillInserted
          ? "CKEDITOR_MODEL_INSERT_OK"
          : insertResult.code
      )
      : (
        !insertResult.insertedNow
          ? (insertResult.rangeInvalid ? "RANGE_INVALID" : "EXEC_COMMAND_FAILED")
          : (stillInserted ? "INSERT_OK" : "INSERT_REVERTED")
      );
    const ok = isCkEditorMethod ? code === "CKEDITOR_MODEL_INSERT_OK" : code === "INSERT_OK";
    const insertionResult = {
      code,
      ok,
      insertionMethod: insertResult.insertionMethod || "DOM_FALLBACK",
      beforeEditorData: insertResult.beforeEditorData ?? null,
      afterEditorData: insertResult.afterEditorData ?? null,
      afterDelayEditorData,
      stillPresentAfterDelay: stillInserted,
      modelSelectionPositionBefore: insertResult.modelSelectionPositionBefore ?? null,
      modelSelectionPositionAfter: insertResult.modelSelectionPositionAfter ?? null,
      message: insertResult.message || null,
      state: insertResult.state || null
    };
    const diagnostics = buildDiagnostics(editor, rows, {
      execCommandRan: insertResult.execCommandRan,
      execCommandResult: insertResult.execCommandResult,
      fallbackRan: insertResult.fallbackRan,
      before: insertResult.before,
      immediate: insertResult.immediate,
      afterDelay,
      rangeInvalid: insertResult.rangeInvalid,
      beforeEditorData: insertResult.beforeEditorData ?? null,
      afterEditorData: insertResult.afterEditorData ?? null,
      afterDelayEditorData,
      modelSelectionPositionAfter: insertResult.modelSelectionPositionAfter ?? null,
      insertionMethod: insertionResult.insertionMethod,
      insertionResult
    });
    state.lastDiagnostics = diagnostics;
    emitResult({
      id: payload.id,
      ok,
      code,
      message: resultMessage(code),
      diagnostics
    });
  }

  function resultMessage(code) {
    const messages = {
      INSERT_OK: "文字已插入，100ms 后仍然存在。",
      INSERT_REVERTED: "普通编辑器插入验证未通过。",
      WRONG_EDITOR: "疑似选中了错误的输入框。",
      NO_EDITOR: "没有找到当前可用的消息编辑器。",
      RANGE_INVALID: "保存的光标范围已失效，备用插入仍未成功。",
      EXEC_COMMAND_FAILED: "execCommand 和备用 Range 插入都没有改变编辑器内容。",
      EDITOR_IN_FRAME: "检测到当前焦点可能位于 iframe，请查看诊断日志。",
      DIAGNOSTIC_OK: "诊断结果已输出到页面控制台。",
      CKEDITOR_MODEL_INSERT_OK: "文字已通过 Teams 编辑器模型插入。",
      NO_CKEDITOR_INSTANCE: "检测到 Teams 编辑器，但无法取得 CKEditor 实例。",
      CKEDITOR_INSTANCE_MISSING: "检测到 Teams 编辑器，但无法取得 CKEditor 实例。",
      CKEDITOR_NOT_READY: "Teams 编辑器尚未准备完成，请稍后重试。",
      CKEDITOR_MODEL_NOT_CHANGED: "已调用 Teams 编辑器模型，但内容没有发生变化。",
      CKEDITOR_MODEL_EXCEPTION: "Teams 编辑器模型插入发生异常。"
    };
    return messages[code] || code;
  }

  function insertFromPayload(payload) {
    const findResult = findBestEditor();
    const editor = findResult.editor;
    const rows = findResult.rows;
    if (!editor) {
      emitResult({
        id: payload.id,
        ok: false,
        code: getDoc().activeElement && getDoc().activeElement.tagName === "IFRAME" ? "EDITOR_IN_FRAME" : "NO_EDITOR",
        message: resultMessage(getDoc().activeElement && getDoc().activeElement.tagName === "IFRAME" ? "EDITOR_IN_FRAME" : "NO_EDITOR"),
        diagnostics: buildDiagnostics(null, rows, {})
      });
      return;
    }

    updateMemoryFromEditor(editor);
    const text = typeof payload.text === "string" ? payload.text : String(payload.text || "");
    const mode = payload.mode === "append" || payload.mode === "replace" ? payload.mode : "cursor";
    let insertResult;
    if (getCkEditorInstance(editor)) {
      insertResult = insertViaCkEditor5(editor, text);
    } else if (isCkEditorElement(editor)) {
      insertResult = insertViaCkEditor5(editor, text);
    } else {
      insertResult = isTextInput(editor)
        ? insertIntoTextInput(editor, text, mode)
        : insertIntoContentEditable(editor, text, mode);
    }
    updateMemoryFromEditor(editor);
    const verificationDelayMs = insertResult.insertionMethod === "CKEDITOR_MODEL" ? 500 : 100;
    root.setTimeout(() => finalizeInsertResult(payload, editor, rows, insertResult), verificationDelayMs);
  }

  function runDiagnostics(id) {
    const findResult = findBestEditor();
    const diagnostics = buildDiagnostics(findResult.editor, findResult.rows, {
      execCommandAvailable: Boolean(getDoc().execCommand),
      bridgeWorld: "MAIN",
      insertionMethod: null,
      insertionResult: null
    });
    state.lastDiagnostics = diagnostics;
    try {
      console.group("[Teams Typing Helper] Input diagnostics");
      console.log("activeElement", getDoc().activeElement);
      console.log("selection", getSelection(getDoc()));
      console.log("selectedEditor", findResult.editor);
      console.log("lastEditor", state.lastEditor);
      console.log("lastRange", state.lastRange);
      console.table(findResult.rows);
      console.log("diagnostics", diagnostics);
      console.groupEnd();
    } catch (error) {
      console.log("[Teams Typing Helper] Diagnostics", diagnostics);
    }
    emitResult({
      id,
      ok: Boolean(findResult.editor),
      code: "DIAGNOSTIC_OK",
      message: resultMessage("DIAGNOSTIC_OK"),
      diagnostics
    });
  }

  function readPayload(attr) {
    const doc = getDoc();
    const raw = doc.documentElement.getAttribute(attr);
    doc.documentElement.removeAttribute(attr);
    if (!raw) {
      return {};
    }
    try {
      return JSON.parse(raw);
    } catch (error) {
      return {};
    }
  }

  function handleInsertEvent() {
    insertFromPayload(readPayload(INSERT_ATTR));
  }

  function handleDiagnoseEvent() {
    const payload = readPayload(DIAGNOSE_ATTR);
    runDiagnostics(payload.id || "");
  }

  function bindMemoryListeners() {
    ["focusin", "selectionchange", "keyup", "mouseup", "input", "beforeinput"].forEach((eventName) => {
      getDoc().addEventListener(eventName, rememberFromEvent, true);
    });
  }

  bindMemoryListeners();
  getDoc().addEventListener(INSERT_EVENT, handleInsertEvent, true);
  getDoc().addEventListener(DIAGNOSE_EVENT, handleDiagnoseEvent, true);

  root.TeamsTypingHelperMainBridge = {
    diagnose: runDiagnostics,
    insertText(text, mode) {
      insertFromPayload({ id: "manual", text, mode: mode || "cursor" });
    },
    findBestEditor,
    getLastDiagnostics() {
      return state.lastDiagnostics;
    }
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
