# Teams 辅助打字插件

适用于 Microsoft Teams 网页版的 Chrome / Edge Manifest V3 扩展。用户可以维护常用词组，并在 Teams 页面通过固定工具栏把词组插入当前消息输入框。扩展不会自动发送消息，不读取聊天记录，不上传任何数据。

## 功能

- 新建、编辑、删除常用词组。
- 上移、下移调整词组按钮顺序。
- 插入方式：光标位置插入、追加到末尾、替换全部内容。
- 页面工具栏支持开启/关闭、折叠/展开、拖动位置、紧凑模式、每行按钮数。
- 导入、导出 JSON 配置，支持恢复默认设置。
- 词组和设置保存到 `chrome.storage.local`。
- 页面工具栏提供“测试插入：你好”和“诊断输入框”按钮。

支持 Teams 地址：

- `https://teams.live.com/v2/*`
- `https://teams.microsoft.com/*`
- `https://teams.cloud.microsoft/*`

## 插入机制

扩展拆成两个执行环境：

- `src/content/content-script.js` 运行在 `ISOLATED` world，负责工具栏、配置读取、按钮交互和结果提示。
- `src/content/main-bridge.js` 运行在 `MAIN` world，负责识别 Teams 真实编辑器、保存/恢复光标、执行插入和输出诊断。

词组按钮使用 `pointerdown` 同步触发插入请求。content script 会把待插入文字临时写入 `document.documentElement` 的 data 属性，并同步派发 DOM Event；MAIN bridge 监听该事件后立即在页面主环境中执行插入。

如果识别到 Teams 的 CKEditor 5 输入框，扩展会直接读取：

```js
domEditable.ckeditorInstance
```

并通过 CKEditor Model 插入：

```js
editor.model.change(writer => {
  const textNode = writer.createText(text, attributes);
  editor.model.insertContent(textNode);
});
```

这条路径不会调用 `execCommand`，不会使用 `Range.insertNode()`，也不会伪造 `input` 事件。只有当元素不是 CKEditor 时，才会使用普通 textarea/contenteditable 备用方案。

## 安装

### Chrome

1. 打开 `chrome://extensions/`。
2. 开启右上角“开发者模式”。
3. 点击“加载已解压的扩展程序”。
4. 选择目录：`D:\redownload\ruanjian\teams-typing-helper`。

### Edge

1. 打开 `edge://extensions/`。
2. 开启“开发人员模式”。
3. 点击“加载解压缩的扩展”。
4. 选择目录：`D:\redownload\ruanjian\teams-typing-helper`。

## 使用

1. 点击浏览器工具栏中的扩展图标。
2. 在弹窗中添加或修改词组。
3. 打开 Teams 网页版聊天、群聊或频道页面。
4. 先点击 Teams 消息输入框，让光标出现在输入框内。
5. 点击页面工具栏里的“测试插入：你好”或任意词组按钮。

插入完成后，工具栏会显示结果，例如 `CKEDITOR_MODEL_INSERT_OK`、`NO_CKEDITOR_INSTANCE`、`CKEDITOR_NOT_READY`、`NO_EDITOR` 等对应提示。

## 导入导出

弹窗中的“备份与恢复”区域提供：

- `导出词组`：读取浏览器存储中的最新词组和设置，生成本地 JSON 文件。
- `导入词组`：选择 `.json` 文件后先显示预览，不会立即覆盖当前数据。
- `恢复最近备份`：恢复最近一次正式导入前自动保存的数据。
- `撤销本次导入`：导入成功后出现，用于恢复导入前的数据。

导出文件名格式：

```text
teams-typing-helper-backup-YYYY-MM-DD-HHmmss.json
```

导出 JSON 结构：

```json
{
  "app": "teams-typing-helper",
  "formatVersion": 1,
  "exportedAt": "2026-07-11T21:15:30.000Z",
  "phrases": [
    {
      "id": "phrase_xxxxx",
      "name": "问候",
      "text": "你好，请问现在方便吗？",
      "order": 0,
      "enabled": true,
      "createdAt": "2026-07-11T20:00:00.000Z",
      "updatedAt": "2026-07-11T20:00:00.000Z"
    }
  ],
  "settings": {
    "toolbarEnabled": true,
    "compactMode": false,
    "buttonsPerRow": 3,
    "insertMode": "cursor",
    "toolbarPosition": null,
    "toolbarCollapsed": false
  }
}
```

导入支持两种方式：

- 合并导入：保留现有词组，导入的新词组追加到后面；完全重复的名称和文字会跳过；默认不导入设置。
- 替换全部：删除当前词组，使用导入文件中的词组；默认同时导入设置；替换前会自动创建本地备份。

导入会严格验证文件：`app`、`formatVersion`、`phrases`、字段类型、名称/文字长度、最多 1000 个词组、最大 5 MB。只导入白名单设置字段，未知字段会忽略。预览使用纯文本渲染，不执行导入文件中的 HTML。

## 诊断日志

在 Teams 页面点击工具栏中的“诊断输入框”，然后打开页面控制台查看：

- `document.activeElement`
- 当前 Selection 和 anchorNode
- contenteditable / role=textbox 数量
- 候选编辑器表格
- 最终选择的 editor
- editor 是否位于 iframe 或 open Shadow Root
- lastEditor / lastRange
- CKEditor 实例是否存在、`state`、`model`、`editing.view`
- CKEditor `getData()` 插入前、插入后、500ms 后的数据
- CKEditor Model selection 的安全序列化位置
- 使用的插入方法和插入结果

如果真实 Teams 中仍然无法插入，请把控制台里 `[Teams Typing Helper] Input diagnostics` 这组日志发回来。

## 隐私和安全

- 权限仅 `storage`。
- 不读取联系人、群组成员或聊天历史。
- 不上传聊天内容或词组。
- 不在后台自动输入内容。
- 没有自动发送消息功能。
- 只有用户主动点击工具栏按钮时才会插入文字。

## 已知限制

- 自动化测试使用 Teams-like 本地页面模拟 Teams 编辑器行为，没有登录真实 Teams 账号，也不会向真实联系人发送消息。
- 真实 Teams 可能把编辑器放在 iframe 或 Shadow DOM 中。当前 bridge 会检测 open Shadow Root，并在诊断中标出 iframe 迹象；如果日志证明输入框确实在 iframe 中，再继续实现 frame 路由。
- 当前优先修复 `https://teams.live.com/v2/` 的 CKEditor 5 输入框；其他 Teams 域名仍保留普通编辑器备用逻辑。

## 开发和测试

校验：

```powershell
node scripts/validate-extension.js
```

测试：

```powershell
npm.cmd test
```

或：

```powershell
node --test tests/unit/*.test.js tests/browser/*.test.js
```

## 项目结构

```text
teams-typing-helper/
  manifest.json
  package.json
  README.md
  scripts/
    validate-extension.js
  src/
    shared/
      config.js
    content/
      content-script.js
      main-bridge.js
      toolbar.css
    popup/
      popup.html
      popup.css
      popup.js
  tests/
    unit/
      config.test.js
      manifest.test.js
    browser/
      extension-behavior.test.js
      playwright-loader.js
```
