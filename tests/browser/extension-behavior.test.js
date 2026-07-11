const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { loadPlaywright } = require("./playwright-loader.js");

const playwright = loadPlaywright();
const projectRoot = path.resolve(__dirname, "../..");
const configPath = path.join(projectRoot, "src/shared/config.js");
const mainBridgePath = path.join(projectRoot, "src/content/main-bridge.js");
const contentScriptPath = path.join(projectRoot, "src/content/content-script.js");
const toolbarCssPath = path.join(projectRoot, "src/content/toolbar.css");
const popupPath = path.join(projectRoot, "src/popup/popup.html");
const STORAGE_KEY = "teamsTypingHelperConfig";

async function withBrowser(t, callback) {
  if (!playwright) {
    t.skip("Playwright is not installed in this workspace.");
    return;
  }

  let browser;
  try {
    browser = await playwright.chromium.launch({ headless: true });
  } catch (error) {
    t.skip(`Chromium is not available: ${error.message}`);
    return;
  }

  try {
    await callback(browser);
  } finally {
    await browser.close();
  }
}

function buildConfig(overrides = {}) {
  return {
    version: 1,
    enabled: true,
    insertMode: "cursor",
    compactMode: false,
    columns: 2,
    toolbar: { collapsed: false, position: null },
    phrases: [{ id: "p1", name: "短语", text: "HELLO" }],
    ...overrides
  };
}

async function installChromeMock(page, config) {
  await page.addInitScript(({ storageKey, initialConfig }) => {
    const listeners = [];
    const store = { [storageKey]: initialConfig };
    window.chrome = {
      runtime: { lastError: null },
      storage: {
        local: {
          get(key, callback) {
            const result = {};
            if (typeof key === "string") {
              result[key] = store[key];
            } else if (Array.isArray(key)) {
              key.forEach((item) => {
                result[item] = store[item];
              });
            } else {
              Object.assign(result, store);
            }
            callback(result);
          },
          set(values, callback) {
            Object.entries(values).forEach(([key, value]) => {
              const oldValue = store[key];
              store[key] = value;
              const changes = { [key]: { oldValue, newValue: value } };
              listeners.forEach((listener) => listener(changes, "local"));
            });
            if (callback) {
              callback();
            }
          }
        },
        onChanged: {
          addListener(listener) {
            listeners.push(listener);
          }
        }
      }
    };
    window.__tthStore = store;
  }, { storageKey: STORAGE_KEY, initialConfig: config });
}

function scenarioHtml(editorHtml) {
  return `<!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { margin: 0; font-family: sans-serif; }
          .layout { min-height: 100vh; display: grid; grid-template-rows: 1fr auto; }
          .messages { padding: 24px; }
          .composer { padding: 20px 24px; border-top: 1px solid #ddd; }
          [contenteditable="true"] { min-height: 42px; width: 620px; border: 1px solid #888; padding: 8px; }
          .hidden-box { display: none; }
          .search-box { position: absolute; top: 12px; left: 12px; width: 200px; min-height: 24px; }
        </style>
      </head>
      <body>
        <div class="layout">
          <div class="messages">
            <div class="search-box" contenteditable="true" role="textbox" aria-label="Search messages"></div>
            <div class="hidden-box" contenteditable="true" role="textbox" aria-label="Hidden message box"></div>
          </div>
          <div class="composer">${editorHtml}</div>
        </div>
        <script>
          window.__events = [];
          document.addEventListener("beforeinput", (event) => window.__events.push(event.type + ":" + event.inputType), true);
          document.addEventListener("input", (event) => window.__events.push(event.type + ":" + event.inputType), true);
          document.addEventListener("change", (event) => window.__events.push(event.type), true);
        </script>
      </body>
    </html>`;
}

async function loadTeamsFixture(page, url, editorHtml, config = buildConfig()) {
  await installChromeMock(page, config);
  const fulfillFixture = (route) => route.fulfill({
    status: 200,
    contentType: "text/html",
    body: scenarioHtml(editorHtml)
  });
  await page.route("https://teams.microsoft.com/**", fulfillFixture);
  await page.route("https://teams.live.com/**", fulfillFixture);
  await page.route("https://teams.cloud.microsoft/**", fulfillFixture);
  await page.goto(url);
  await injectExtensionScripts(page);
}

async function injectExtensionScripts(page) {
  await page.addStyleTag({ path: toolbarCssPath });
  await page.addScriptTag({ path: mainBridgePath });
  await page.addScriptTag({ path: configPath });
  await page.addScriptTag({ path: contentScriptPath });
  await page.waitForSelector("#teams-typing-helper-toolbar");
}

async function clickPhrase(page) {
  await page.evaluate(() => {
    window.__teamsTypingHelperLastResult = null;
    window.__teamsTypingHelperLastToolbarResult = null;
  });
  await page.locator(".tth-phrase-button").first().click();
  await page.waitForFunction(() => window.__teamsTypingHelperLastResult);
}

async function lastResult(page) {
  await page.waitForFunction(() => window.__teamsTypingHelperLastResult);
  return page.evaluate(() => window.__teamsTypingHelperLastResult);
}

test("private chat: inserts phrase into an empty Teams-like contenteditable editor", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=user@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.locator("#editor").click();
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "HELLO");
    const events = await page.evaluate(() => window.__events);
    assert.ok(events.some((event) => event.startsWith("input")));
    assert.ok(events.includes("change"));
    assert.equal((await lastResult(page)).code, "INSERT_OK");
  });
});

test("group chat: inserts at caret position without replacing existing text", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/19:group/thread.v2",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a message">ABC</div>`
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      const editor = document.querySelector("#editor");
      const range = document.createRange();
      range.setStart(editor.firstChild, 1);
      range.setEnd(editor.firstChild, 1);
      const selection = getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    });
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "AHELLOBC");
  });
});

test("channel page: replace mode replaces all editor content", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/channel/19:channel-id/team-name?groupId=1&tenantId=2",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Start a new post">OLD TEXT</div>`,
      buildConfig({ insertMode: "replace" })
    );

    await page.locator("#editor").click();
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "HELLO");
  });
});

test("append mode appends to the existing editor text", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/v2/",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Message chat">ABC</div>`,
      buildConfig({ insertMode: "append" })
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      const editor = document.querySelector("#editor");
      const range = document.createRange();
      range.setStart(editor.firstChild, 1);
      range.collapse(true);
      const selection = getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    });
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "ABCHELLO");
  });
});

test("selected text is replaced at the current selection", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/19:group/thread.v2",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a message">ABCDE</div>`
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      const editor = document.querySelector("#editor");
      const range = document.createRange();
      range.setStart(editor.firstChild, 1);
      range.setEnd(editor.firstChild, 4);
      const selection = getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    });
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "AHELLOE");
    assert.equal((await lastResult(page)).code, "INSERT_OK");
  });
});

test("manual beforeinput preventDefault does not stop fallback insertion", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=blocked@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.evaluate(() => {
      document.addEventListener("beforeinput", (event) => event.preventDefault(), true);
    });
    await page.locator("#editor").click();
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "HELLO");
    assert.equal((await lastResult(page)).code, "INSERT_OK");
  });
});

test("test insert button inserts the fixed nihao text", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=test-button@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      window.__teamsTypingHelperLastResult = null;
    });
    await page.locator(".tth-more-button").click();
    await page.waitForSelector(".tth-tools:not([hidden])");
    await page.locator(".tth-test-button").click();
    await page.waitForFunction(() => window.__teamsTypingHelperLastResult);

    assert.equal(await page.locator("#editor").innerText(), "\u4f60\u597d");
    assert.equal((await lastResult(page)).code, "INSERT_OK");
  });
});

test("diagnostic button produces a diagnostic result", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=diagnose@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      window.__teamsTypingHelperLastResult = null;
    });
    await page.locator(".tth-more-button").click();
    await page.waitForSelector(".tth-tools:not([hidden])");
    await page.locator(".tth-diagnose-button").click();
    await page.waitForFunction(() => window.__teamsTypingHelperLastResult);
    const result = await lastResult(page);

    assert.equal(result.code, "DIAGNOSTIC_OK");
    assert.ok(result.diagnostics.counts.candidates >= 1);
  });
});

test("ckeditor element without instance does not use DOM fallback insertion", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.live.com/v2/",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="\u952e\u5165\u6d88\u606f" data-tid="ckeditor">
        <p class="ck-placeholder" data-placeholder="\u952e\u5165\u6d88\u606f"><br></p>
      </div>`
    );

    await page.locator("#editor").click();
    await page.evaluate(() => {
      const editor = document.querySelector("#editor");
      const placeholder = editor.querySelector("p");
      editor.focus();
      const range = document.createRange();
      range.setStart(placeholder, 0);
      range.setEnd(placeholder, 0);
      const selection = getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event("selectionchange", { bubbles: true }));
    });
    await clickPhrase(page);
    const result = await lastResult(page);

    assert.doesNotMatch(await page.locator("#editor").innerText(), /HELLO/);
    assert.equal(result.code, "NO_CKEDITOR_INSTANCE");
    assert.equal(result.diagnostics.insertionMethod, "CKEDITOR_MODEL");
    assert.equal(result.diagnostics.selectedEditor.tagName, "DIV");
    assert.equal(result.diagnostics.selectedEditor.dataTid, "ckeditor");
  });
});

test("ckeditor instance uses model insertion and verifies getData after delay", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.live.com/v2/",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="\u952e\u5165\u6d88\u606f" data-tid="ckeditor" class="ck ck-content ck-editor__editable">
        <p class="ck-placeholder" data-placeholder="\u952e\u5165\u6d88\u606f"><br></p>
      </div>`
    );

    await page.evaluate(() => {
      const editorElement = document.querySelector("#editor");
      let editorData = "<p>OLD</p>";
      window.__execCommandCalled = false;
      document.execCommand = () => {
        window.__execCommandCalled = true;
        return false;
      };
      editorElement.ckeditorInstance = {
        state: "ready",
        getData() {
          return editorData;
        },
        editing: {
          view: {
            focus() {
              window.__ckEditorFocused = (window.__ckEditorFocused || 0) + 1;
            }
          }
        },
        model: {
          document: {
            selection: {
              isCollapsed: true,
              getAttributes() {
                return [["bold", true]];
              },
              getFirstPosition() {
                return {
                  root: { rootName: "main" },
                  path: [0, 0]
                };
              }
            }
          },
          change(callback) {
            const writer = {
              createText(text, attributes) {
                return { text, attributes };
              }
            };
            callback(writer);
          },
          insertContent(textNode) {
            window.__ckInsertedNode = textNode;
            editorData = `<p>OLD${textNode.text}</p>`;
            editorElement.innerHTML = `<p>OLD${textNode.text}</p>`;
          }
        }
      };
      editorElement.focus();
      const placeholder = editorElement.querySelector("p");
      const range = document.createRange();
      range.setStart(placeholder, 0);
      range.setEnd(placeholder, 0);
      const selection = getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event("selectionchange", { bubbles: true }));
    });

    await clickPhrase(page);
    const result = await lastResult(page);

    assert.equal(result.code, "CKEDITOR_MODEL_INSERT_OK");
    assert.equal(result.diagnostics.insertionMethod, "CKEDITOR_MODEL");
    assert.equal(result.diagnostics.insertionResult.stillPresentAfterDelay, true);
    assert.match(result.diagnostics.afterEditorData, /HELLO/);
    assert.match(result.diagnostics.afterDelayEditorData, /HELLO/);
    assert.deepEqual(result.diagnostics.insertionResult.modelSelectionPositionBefore.path, [0, 0]);
    assert.equal(await page.evaluate(() => window.__execCommandCalled), false);
    assert.equal(await page.evaluate(() => window.__ckInsertedNode.attributes.bold), true);
    assert.match(await page.locator("#editor").innerText(), /HELLO/);
  });
});

test("does not insert into the wrong contenteditable when several editors exist", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/19:group/thread.v2",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a message"></div>`
    );

    await page.locator("#editor").click();
    await clickPhrase(page);

    assert.equal(await page.locator("#editor").innerText(), "HELLO");
    assert.equal(await page.locator(".search-box").innerText(), "");
  });
});

test("continues to work after switching chats and editor re-render", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=a@example.com",
      `<div id="editor-a" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.locator("#editor-a").click();
    await clickPhrase(page);
    assert.equal(await page.locator("#editor-a").innerText(), "HELLO");

    await page.evaluate(() => {
      document.querySelector(".composer").innerHTML = `<div id="editor-b" contenteditable="true" role="textbox" aria-label="Type a message"></div>`;
    });
    await page.locator("#editor-b").click();
    await clickPhrase(page);
    assert.equal(await page.locator("#editor-b").innerText(), "HELLO");

    await page.evaluate(() => {
      document.querySelector(".composer").innerHTML = `<div id="editor-c" contenteditable="true" role="textbox" aria-label="Reply"></div>`;
    });
    await page.locator("#editor-c").click();
    await clickPhrase(page);
    assert.equal(await page.locator("#editor-c").innerText(), "HELLO");
  });
});

test("continues to work after page refresh", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=refresh@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    await page.locator("#editor").click();
    await clickPhrase(page);
    assert.equal(await page.locator("#editor").innerText(), "HELLO");

    await page.reload();
    await injectExtensionScripts(page);
    await page.locator("#editor").click();
    await clickPhrase(page);
    assert.equal(await page.locator("#editor").innerText(), "HELLO");
  });
});

test("toolbar can be dragged and stores its position", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=a@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    const header = page.locator(".tth-header");
    const before = await page.locator("#teams-typing-helper-toolbar").boundingBox();
    await header.evaluate((element) => {
      const rect = element.getBoundingClientRect();
      const startX = rect.left + rect.width / 2;
      const startY = rect.top + rect.height / 2;
      element.dispatchEvent(new MouseEvent("mousedown", {
        bubbles: true,
        button: 0,
        clientX: startX,
        clientY: startY
      }));
      document.dispatchEvent(new MouseEvent("mousemove", {
        bubbles: true,
        button: 0,
        clientX: startX + 120,
        clientY: startY + 70
      }));
      document.dispatchEvent(new MouseEvent("mouseup", {
        bubbles: true,
        button: 0,
        clientX: startX + 120,
        clientY: startY + 70
      }));
    });

    const storedPosition = await page.evaluate(() => window.__tthStore.teamsTypingHelperConfig.toolbar.position);
    assert.ok(Number.isFinite(storedPosition.left));
    assert.ok(Number.isFinite(storedPosition.top));
    assert.notEqual(Math.round(storedPosition.left), Math.round(before.x));
  });
});

test("toolbar rerenders when phrases change in storage", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=refresh-toolbar@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`,
      buildConfig({ phrases: [{ id: "a", name: "A", text: "one" }] })
    );

    assert.deepEqual(await page.locator(".tth-phrase-button").evaluateAll((nodes) => nodes.map((node) => node.textContent)), ["A"]);
    await page.evaluate((storageKey) => {
      chrome.storage.local.set({
        [storageKey]: {
          version: 1,
          enabled: true,
          insertMode: "cursor",
          compactMode: false,
          columns: 2,
          toolbar: { collapsed: false, position: null },
          phrases: [{ id: "b", name: "B", text: "two" }]
        }
      });
    }, STORAGE_KEY);

    await page.waitForFunction(() => Array.from(document.querySelectorAll(".tth-phrase-button")).some((node) => node.textContent === "B"));
    assert.deepEqual(await page.locator(".tth-phrase-button").evaluateAll((nodes) => nodes.map((node) => node.textContent)), ["B"]);
  });
});

test("toolbar keeps diagnostic tools behind the more button", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage({ viewport: { width: 900, height: 700 } });
    await loadTeamsFixture(
      page,
      "https://teams.microsoft.com/l/chat/0/0?users=toolbar-ui@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`
    );

    assert.equal(await page.locator(".tth-tools").isHidden(), true);
    await page.locator(".tth-more-button").click();
    await page.waitForSelector(".tth-tools:not([hidden])");
    assert.equal(await page.locator(".tth-tools").isVisible(), true);
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => document.querySelector(".tth-tools")?.hidden === true);
    assert.equal(await page.locator("#teams-typing-helper-toolbar").count(), 1);
  });
});

test("popup can add, edit, delete, reorder, export and import phrases", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await installChromeMock(page, buildConfig({
      phrases: [
        { id: "a", name: "A", text: "one" },
        { id: "b", name: "B", text: "two" }
      ]
    }));
    await page.goto(`file:///${popupPath.replace(/\\/g, "/")}`);

    await page.locator("#phrase-name").fill("C");
    await page.locator("#phrase-text").fill("three");
    await page.locator("#save-phrase").click();
    await assertPhraseNames(page, ["A", "B", "C"]);

    await clickPhraseMenuAction(page, "C", "编辑词组");
    await page.locator("#phrase-name").fill("C2");
    await page.locator("#phrase-text").fill("three edited");
    await page.locator("#save-phrase").click();
    await assertPhraseNames(page, ["A", "B", "C2"]);

    await clickPhraseMenuAction(page, "C2", "上移词组");
    await assertPhraseNames(page, ["A", "C2", "B"]);

    await clickPhraseMenuAction(page, "B", "删除词组");
    await page.waitForSelector("#delete-confirm:not([hidden])");
    assert.equal(await page.evaluate(() => document.activeElement.id), "cancel-delete");
    await page.locator("#confirm-delete").click();
    await assertPhraseNames(page, ["A", "C2"]);

    const exported = await page.evaluate(() => JSON.stringify(TeamsTypingHelperConfig.buildExportPayload(window.__tthStore.teamsTypingHelperConfig)));
    assert.match(exported, /C2/);

    const imported = await page.evaluate((jsonText) => TeamsTypingHelperConfig.parseImportedConfig(jsonText), exported);
    assert.deepEqual(imported.phrases.map((phrase) => phrase.name), ["A", "C2"]);

    const savedConfig = await page.evaluate(() => window.__tthStore.teamsTypingHelperConfig);
    const teamsPage = await browser.newPage();
    await loadTeamsFixture(
      teamsPage,
      "https://teams.microsoft.com/l/chat/0/0?users=persist@example.com",
      `<div id="editor" contenteditable="true" role="textbox" aria-label="Type a new message"></div>`,
      savedConfig
    );
    const toolbarNames = await teamsPage.locator(".tth-phrase-button").evaluateAll((nodes) => nodes.map((node) => node.textContent));
    assert.deepEqual(toolbarNames, ["A", "C2"]);
  });
});

test("popup UI contains long phrases, menus and delete confirmation at compact width", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage({ viewport: { width: 360, height: 760 } });
    await installChromeMock(page, buildConfig({
      phrases: [{
        id: "long",
        name: "Very long phrase name that should not break the card layout",
        text: "Line one with a lot of text. Line two with a lot of text. Line three with a lot of text. Line four should be clipped."
      }]
    }));
    await page.goto(`file:///${popupPath.replace(/\\/g, "/")}`);

    for (const width of [360, 400, 420, 520]) {
      await page.setViewportSize({ width, height: 760 });
      const fitsViewport = await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth);
      assert.equal(fitsViewport, true, `popup overflowed at ${width}px`);
      const controlsFit = await page.locator("input:not([type='hidden']), textarea, select").evaluateAll((nodes) => nodes.every((node) => {
        if (node.offsetParent === null) {
          return true;
        }
        const rect = node.getBoundingClientRect();
        const parentRect = node.parentElement.getBoundingClientRect();
        return rect.left >= parentRect.left - 0.5 && rect.right <= parentRect.right + 0.5;
      }));
      assert.equal(controlsFit, true, `form controls overflowed at ${width}px`);
    }

    assert.equal(await page.locator(".phrase-text").evaluate((node) => getComputedStyle(node).webkitLineClamp), "3");

    await page.locator(".phrase-row").getByLabel("打开词组操作菜单").click();
    const menu = page.locator("#tth-floating-layer .phrase-menu");
    await menu.waitFor();
    assert.equal(await menu.evaluate((node) => node.parentElement.id), "tth-floating-layer");
    assert.equal(await page.locator(".phrase-row .phrase-menu").count(), 0);
    assert.deepEqual(await menu.getByRole("menuitem").evaluateAll((nodes) => nodes.map((node) => node.textContent.trim())), ["上移", "下移", "编辑", "删除"]);
    assert.equal(await menu.getByRole("menuitem").evaluateAll((nodes) => nodes.every((node) => {
      const rect = node.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    })), true);
    const menuBox = await menu.boundingBox();
    const viewport = page.viewportSize();
    assert.ok(menuBox.x >= 0);
    assert.ok(menuBox.x + menuBox.width <= viewport.width);
    assert.ok(menuBox.y >= 0);
    assert.ok(menuBox.y + menuBox.height <= viewport.height);

    await menu.getByRole("menuitem", { name: "删除词组" }).click();
    await page.waitForSelector("#delete-confirm:not([hidden])");
    assert.equal(await page.evaluate(() => document.activeElement.id), "cancel-delete");
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => document.querySelector("#delete-confirm")?.hidden === true);
    await assertPhraseNames(page, ["Very long phrase name that should not break the card layout"]);
  });
});

test("popup action menu floats above cards and repositions near viewport edges", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage({ viewport: { width: 400, height: 420 } });
    await installChromeMock(page, buildConfig({
      phrases: Array.from({ length: 8 }, (_, index) => ({
        id: `p${index}`,
        name: `Phrase ${index + 1}`,
        text: `Text ${index + 1}`
      }))
    }));
    await page.goto(`file:///${popupPath.replace(/\\/g, "/")}`);

    await openPhraseMenu(page, "Phrase 1");
    assert.equal(await page.locator("#tth-floating-layer .phrase-menu").count(), 1);
    await assertMenuInsideViewport(page);
    assert.equal(await page.getByRole("menuitem", { name: "上移词组" }).isDisabled(), true);
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => document.querySelectorAll(".phrase-menu").length === 0);
    assert.equal(await page.evaluate(() => document.activeElement?.getAttribute("aria-label")), "打开词组操作菜单");

    await openPhraseMenu(page, "Phrase 4");
    await assertMenuInsideViewport(page);
    await page.keyboard.press("ArrowDown");
    assert.equal(await page.evaluate(() => document.activeElement?.getAttribute("aria-label")), "下移词组");

    await page.keyboard.press("Escape");
    await page.waitForFunction(() => document.querySelectorAll(".phrase-menu").length === 0);
    await page.locator(".phrase-row").filter({ hasText: "Phrase 8" }).getByLabel("打开词组操作菜单").evaluate((button) => {
      button.scrollIntoView({ block: "end" });
    });
    await openPhraseMenu(page, "Phrase 8");
    const menu = page.locator("#tth-floating-layer .phrase-menu");
    assert.equal(await menu.evaluate((node) => node.dataset.placement), "top");
    await assertMenuInsideViewport(page);
    assert.equal(await page.getByRole("menuitem", { name: "下移词组" }).isDisabled(), true);

    await page.waitForTimeout(180);
    await page.evaluate(() => window.dispatchEvent(new Event("scroll")));
    await page.waitForFunction(() => document.querySelectorAll(".phrase-menu").length === 0);

    await openPhraseMenu(page, "Phrase 2");
    await openPhraseMenu(page, "Phrase 3");
    assert.equal(await page.locator("#tth-floating-layer .phrase-menu").count(), 1);

    await page.mouse.click(8, 8);
    await page.waitForFunction(() => document.querySelectorAll(".phrase-menu").length === 0);
  });
});

test("popup imports from file only after preview confirmation and can undo", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await installChromeMock(page, buildConfig({
      phrases: [{ id: "a", name: "A", text: "one" }]
    }));
    await page.goto(`file:///${popupPath.replace(/\\/g, "/")}`);

    const payload = {
      app: "teams-typing-helper",
      formatVersion: 1,
      exportedAt: "2026-07-11T21:15:30.000Z",
      phrases: [
        { id: "a", name: "A", text: "one" },
        { id: "b", name: "<img src=x onerror='window.__bad=1'>", text: "<script>window.__bad=1</script>" },
        { id: "c", name: "Hola", text: "Привет 😀\nLinea dos" }
      ],
      settings: { insertMode: "replace", buttonsPerRow: 8, unknown: "<b>x</b>" }
    };

    await page.locator("#import-config").setInputFiles({
      name: "backup.json",
      mimeType: "application/json",
      buffer: Buffer.from(JSON.stringify(payload), "utf8")
    });
    await page.waitForSelector("#import-preview:not([hidden])");

    assert.equal(await page.evaluate(() => window.__bad), undefined);
    assert.match(await page.locator(".preview-name").nth(1).textContent(), /img src=x/);
    assert.match(await page.locator("#import-summary").textContent(), /重复词组/);

    await page.locator("#confirm-import").click();
    await page.waitForFunction(() => window.__tthStore.teamsTypingHelperConfig.phrases.length === 3);
    let stored = await page.evaluate(() => window.__tthStore.teamsTypingHelperConfig);
    assert.deepEqual(stored.phrases.map((phrase) => phrase.name), ["A", "<img src=x onerror='window.__bad=1'>", "Hola"]);
    assert.equal(stored.insertMode, "cursor");

    await page.locator("#undo-import").click();
    await page.waitForFunction(() => window.__tthStore.teamsTypingHelperConfig.phrases.length === 1);
    stored = await page.evaluate(() => window.__tthStore.teamsTypingHelperConfig);
    assert.deepEqual(stored.phrases.map((phrase) => phrase.name), ["A"]);
  });
});

test("popup can select the same import file twice", async (t) => {
  await withBrowser(t, async (browser) => {
    const page = await browser.newPage();
    await installChromeMock(page, buildConfig({ phrases: [] }));
    await page.goto(`file:///${popupPath.replace(/\\/g, "/")}`);
    const file = {
      name: "same.json",
      mimeType: "application/json",
      buffer: Buffer.from(JSON.stringify({
        app: "teams-typing-helper",
        formatVersion: 1,
        phrases: [{ name: "Same", text: "value" }]
      }), "utf8")
    };

    await page.locator("#import-config").setInputFiles(file);
    await page.waitForSelector("#import-preview:not([hidden])");
    await page.locator("#cancel-import").click();
    await page.locator("#import-config").setInputFiles(file);
    await page.waitForSelector("#import-preview:not([hidden])");
    assert.match(await page.locator("#import-summary").textContent(), /Same|有效词组/);
  });
});

async function clickPhraseMenuAction(page, phraseName, actionLabel) {
  await openPhraseMenu(page, phraseName);
  await page.locator("#tth-floating-layer .phrase-menu").waitFor();
  await page.getByRole("menuitem", { name: actionLabel }).click();
}

async function openPhraseMenu(page, phraseName) {
  const row = page.locator(".phrase-row").filter({
    has: page.locator(".phrase-name", { hasText: phraseName })
  }).first();
  await row.getByLabel("打开词组操作菜单").click();
  await page.locator("#tth-floating-layer .phrase-menu").waitFor();
}

async function assertMenuInsideViewport(page) {
  const box = await page.locator("#tth-floating-layer .phrase-menu").boundingBox();
  const viewport = page.viewportSize();
  assert.ok(box.x >= 0, "menu exceeded left edge");
  assert.ok(box.y >= 0, "menu exceeded top edge");
  assert.ok(box.x + box.width <= viewport.width, "menu exceeded right edge");
  assert.ok(box.y + box.height <= viewport.height, "menu exceeded bottom edge");
}

async function assertPhraseNames(page, expected) {
  const names = await page.locator(".phrase-name").evaluateAll((nodes) => nodes.map((node) => node.textContent));
  assert.deepEqual(names, expected);
}
