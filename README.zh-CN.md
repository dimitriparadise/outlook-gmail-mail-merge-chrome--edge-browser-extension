# Mail Merge Draft Helper 中文说明

这是一个 Manifest V3 浏览器插件，用来根据 CSV 名单生成个性化 Outlook 或 Gmail 邮件草稿。插件在弹窗界面中运行：用户可以上传 CSV 文件，或直接粘贴 CSV 文本；然后选择 Outlook Mode 或 Gmail Mode，填写邮件主题、正文、可选 CC/BCC；最后按顺序打开单封草稿，或按指定范围批量打开草稿。

## 主要功能

- 支持上传 CSV 文件，也支持直接粘贴 CSV 内容。
- CSV 必须包含 `Email` 或 `email` 列，作为主收件人 To。
- 支持模板变量，例如 `{{Name}}`、`{{Course}}`、`{{DueDate}}`。
- 点击 `Generate Drafts` 后，会根据 CSV header 自动生成变量按钮。
- 支持个性化 To、CC、BCC、Subject 和 Body。
- 支持 Outlook Mode 和 Gmail Mode。
- 支持一次打开下一封草稿，也支持按范围批量打开草稿。
- 使用 `chrome.storage.local` 保存当前内容和进度，关闭插件弹窗后可以继续。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `manifest.json` | 定义插件信息、权限、弹窗入口和邮件服务域名权限。 |
| `popup.html` | 定义插件弹窗界面，包括 CSV 输入、变量按钮、模板、预览和打开草稿按钮。 |
| `popup.css` | 定义弹窗样式。 |
| `popup.js` | 实现 CSV 解析、变量替换、状态保存、预览生成和打开邮件草稿。 |

## 使用流程

1. 上传 CSV 文件，或把 CSV 内容粘贴到输入框。
2. 选择 `Outlook Mode` 或 `Gmail Mode`。
3. 点击 `Generate Drafts`。
4. 插件会读取 CSV，并显示第一封邮件的预览。
5. 如果 CSV 有更多列，插件会生成对应变量按钮，例如 `{{Name}}`、`{{Course}}`。
6. 可以把变量插入 Subject、CC、BCC 或 Body。
7. 点击 `Open Next Draft` 打开下一封草稿。
8. 或者填写 `From` / `To`，点击 `Open Range` 一次打开某个范围内的草稿。

范围从 1 开始计数。例如 CSV 有 30 行，想一次打开第 6 到第 20 封，就填写：

```text
From: 6
To: 20
```

## CSV 格式

CSV 必须有 header row，并且至少有一行数据。必须包含 `Email` 或 `email` 列。

示例：

```csv
Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com
```

CSV 中的任何 header 都可以作为变量使用：

```text
{{Name}}
{{Course}}
{{Section}}
{{DueDate}}
{{CcEmail}}
{{BccEmail}}
```

示例模板：

```text
Subject:
Reminder for {{Course}}

Body:
Hi {{Name}},

This is a quick reminder about {{Course}} section {{Section}}, due {{DueDate}}.
```

## CC 和 BCC

CC 和 BCC 是可选的，可以留空。

可以填写固定邮箱：

```text
CC: ta@example.com
BCC: archive@example.com
```

也可以填写 CSV 变量：

```text
CC: {{CcEmail}}
BCC: {{BccEmail}}
```

注意：Outlook Web 的 compose deeplink 对 CC/BCC 支持不稳定，可能会忽略 `cc` 和 `bcc` 参数。因此，在 Outlook Mode 中，只要生成的草稿包含 CC 或 BCC，插件会改用标准 `mailto:` 打开草稿。

Gmail Mode 不需要 `mailto:`，会直接使用 Gmail compose URL，并支持 To、CC、BCC、Subject 和 Body。

## Compose Mode

插件有两种模式：

- `Outlook Mode`：打开 Outlook Web 草稿。没有 CC/BCC 时使用 Outlook Web deeplink；有 CC/BCC 时使用 `mailto:`，因为 Outlook Web 可能忽略 CC/BCC。
- `Gmail Mode`：直接打开 Gmail compose URL，不依赖 `mailto:`。Gmail Mode 支持 To、CC、BCC、Subject 和 Body。

如果你希望 CC/BCC 不依赖 Chrome 的 `mailto:` 设置，可以使用 Gmail Mode。

## 在 Chrome 中设置 Mailto

如果在 Outlook Mode 中使用 CC/BCC 时草稿打开到了错误的邮件应用，或者没有打开，请设置 Chrome 的 `mailto:` handler。Gmail Mode 不需要这个设置。

1. 打开 Chrome。
2. 进入 `chrome://settings/handlers`。
3. 打开 `Sites can ask to handle protocols`。
4. 在 Chrome 中打开你想使用的邮件服务，例如 Gmail 或 Outlook Web。
5. 如果地址栏右侧出现 protocol handler 图标，通常是双菱形图标，点击它。
6. 选择 `Allow`，然后点击 `Done`。
7. 回到 `chrome://settings/handlers`，确认该邮件服务已经成为 email 或 `mailto:` handler。

如果没有看到双菱形图标，可以尝试：

- 在 `chrome://settings/handlers` 删除旧的 blocked/default email handlers。
- 刷新 Gmail 或 Outlook Web 页面。
- 重启 Chrome。
- 在 macOS 或 Windows 中，把 Chrome 或你的邮件应用设置为系统默认邮件应用。

## 本地安装插件

1. 打开 Chrome 或 Edge。
2. 进入扩展页面：
   - Chrome: `chrome://extensions`
   - Edge: `edge://extensions`
3. 打开 Developer mode。
4. 点击 `Load unpacked`。
5. 选择这个文件夹。
6. 点击浏览器工具栏中的插件图标，打开弹窗。

每次修改代码后，需要在扩展页面点击 reload，浏览器才会加载最新版本。

## 权限说明

插件使用以下权限：

- `storage`：保存 CSV 文本、模板、生成结果和打开进度。
- `tabs`：打开 Outlook、Gmail 或 mailto 草稿。
- Gmail 和 Outlook host permissions：允许打开对应的 compose 页面。

## 实现说明

- CSV 解析在 `popup.js` 中本地完成。
- CSV parser 支持带引号的字段和转义引号，但不是完整的企业级 CSV 解析器。
- 上传文件后，插件会把文件内容复制到 textarea，因为插件弹窗关闭后无法恢复文件 input。
- Compose 链接使用 `encodeURIComponent()` 编码，避免空格和换行被错误处理。
- Outlook Mode 中有 CC/BCC 的草稿使用 `mailto:`，因为 Outlook Web deeplink 可能忽略 CC/BCC。
- Gmail Mode 使用 Gmail compose URL，不依赖 `mailto:`。
- 插件只打开草稿，不会自动发送邮件。

## 当前限制

- 没有 `Email` 或 `email` 的行会被忽略。
- 不支持附件。
- 不支持富文本编辑器。
- 不检查重复收件人。
- 一次打开很大的范围时，浏览器可能会限制或拦截大量新 tab。
