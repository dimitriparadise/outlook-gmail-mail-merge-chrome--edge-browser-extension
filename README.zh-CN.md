# Mail Merge Draft Helper 中文说明

Mail Merge Draft Helper 是一个轻量级 Chrome 插件，用来根据 CSV 名单快速生成个性化 Gmail 或 Outlook Web 邮件草稿。它适合教师、行政人员、小团队和任何需要批量准备相似邮件、但仍希望逐封检查内容的人使用。

插件完全在浏览器弹窗中运行。你可以上传或粘贴 CSV，编写可复用的邮件主题和正文模板，逐封预览个性化结果，然后在 Gmail 或 Outlook Web 中打开对应草稿。

## 核心功能

- 根据 CSV 每一行生成个性化 Gmail 或 Outlook Web 草稿。
- 支持 `{{Name}}`、`{{Course}}` 等模板变量，也支持任意 CSV 列名。
- 支持逐封预览生成后的邮件内容。
- 支持打开当前预览中的草稿，或按范围批量打开草稿。
- 支持可选 CC 和 BCC，也可以从 CSV 列中读取 CC/BCC。
- 在打开草稿前提示未知模板变量、明显非法邮箱、重复收件人，以及 To 与 CC/BCC 重叠。
- 使用 `chrome.storage.local` 在本地保存进度。
- 可以随时点击 `Clear Saved Data` 清除本地保存的数据。
- 包含实验性的 Auto-Send 模式，并会在按钮上明确显示 `Open & Send...`。

## 使用流程

1. 上传 CSV 文件，或直接粘贴 CSV 文本。
2. 选择 Gmail Mode 或 Outlook Mode。
3. 编辑 Subject、Body、CC、BCC 模板。
4. 点击 `Generate Drafts`。
5. 使用预览按钮逐封检查邮件内容。
6. 点击 `Open Preview Draft` 打开当前预览中的草稿，或填写 From/To 后点击 `Open Selected Drafts` 批量打开。

勾选 Auto-Send 后，单封按钮会变成 `Open & Send Preview Email`，范围按钮会变成 `Open & Send Selected Emails`。

## CSV 格式

CSV 必须包含 header row，并且至少有一行数据。必须包含 `Email` 或 `email` 列。

```csv
Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com
```

任何 CSV header 都可以作为模板变量使用：

```text
Subject:
Reminder for {{Course}}

Body:
Hi {{Name}},

This is a quick reminder about {{Course}} section {{Section}}, due {{DueDate}}.

CC:
{{CcEmail}}

BCC:
{{BccEmail}}
```

注意：header row 和第一行数据之间必须有真实换行。文本框里的自动视觉换行不等于 CSV row break。

## Gmail 和 Outlook 模式

`Gmail Mode` 会直接打开 Gmail compose URL，支持 To、CC、BCC、Subject 和 Body。

`Outlook Mode` 会打开 Outlook Web compose 链接。如果生成的邮件包含 CC 或 BCC，插件会改用标准 `mailto:`，因为某些 Outlook Web 环境可能忽略 deeplink 中的抄送/密送参数。

如果 Outlook Mode 中存在 CC 或 BCC，不能同时使用 Auto-Send。此时请使用 Gmail Mode，或关闭 Auto-Send。

## 隐私说明

Mail Merge Draft Helper 不运行后端服务，也不会把你的 CSV 数据发送到第三方服务器。

插件会把 CSV 文本、模板、生成结果和打开进度保存在 Chrome extension storage 中，这样关闭并重新打开弹窗后可以继续工作。你可以使用 `Clear Saved Data` 删除这些本地保存的数据。

插件使用以下权限：

- `storage`：保存本地草稿生成状态。
- `tabs`：打开 Gmail、Outlook 或 `mailto:` 草稿标签页。
- Gmail 和 Outlook host permissions：用于实验性的 Auto-Send 功能识别由插件打开的 compose 窗口。

## 当前限制

- 当前基于 compose URL 的流程不支持附件。
- 不支持富文本邮件编辑器。
- 一次打开特别大的范围时，浏览器仍然可能限制或拦截大量新 tab。
- Auto-Send 是实验性功能，使用前请仔细检查草稿内容。

## 本地开发和测试

1. 打开 Chrome 或 Edge。
2. 进入 `chrome://extensions` 或 `edge://extensions`。
3. 打开 Developer mode。
4. 点击 `Load unpacked`。
5. 选择这个文件夹。
6. 点击浏览器工具栏中的插件图标打开弹窗。

运行测试：

```bash
node --test
```

运行语法检查：

```bash
node --check popup.js && node --check content.js
```

## 发布到 Chrome Web Store

发布前建议准备：

- 可以上传的生产版本 zip 包。
- 商店展示名称、简短简介、详细描述、分类、语言和截图。
- 128x128 插件图标，以及用于提升展示效果的宣传图片。
- 如果数据披露要求需要，准备隐私政策 URL。
- 对 `storage`、`tabs` 和 host permissions 的清晰权限说明。
- 在 Chrome Developer Dashboard 中完成 privacy practices 和 limited-use certification。

发布流程：

1. 创建或登录 Chrome Web Store developer account。
2. 打开 [Chrome Developer Dashboard](https://chrome.google.com/webstore/devconsole/)。
3. 点击 `Add new item`。
4. 上传插件 zip 文件。
5. 填写 store listing 信息。
6. 填写 privacy practices。
7. 设置 distribution、visibility 和发布地区。
8. 提交 Chrome Web Store 审核。

官方参考：

- [Publish in the Chrome Web Store](https://developer.chrome.com/docs/webstore/publish)
- [Chrome Web Store Developer Program Policies](https://developer.chrome.com/docs/webstore/program-policies)
- [Privacy disclosure requirements](https://developer.chrome.com/docs/webstore/program-policies/user-data-faq)
