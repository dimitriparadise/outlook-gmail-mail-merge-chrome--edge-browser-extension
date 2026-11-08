# Mail Merge Draft Helper 中文说明

[English](README.md)

Mail Merge Draft Helper 是一个轻量的 Chrome / Edge 扩展，可以把一份 CSV 名单变成一封封个性化的 Outlook Web 或 Gmail 邮件。适合老师、行政人员和小团队：要发很多封内容相似的邮件，但每封发出去之前还想自己看一眼。

所有操作都在扩展弹窗里、在你自己的浏览器本地完成。你载入收件人名单，写一份带变量（比如 `{{FirstName}}`）的模板，逐封预览，然后在 Outlook Web 或 Gmail 里打开草稿，可以一封一封打开，也可以按批次打开。

<p align="center"><img src="docs/images/02-popup-overview.png" alt="扩展弹窗" width="360"></p>

## 目录

- [安装](#安装)
- [分步教程](#分步教程)
- [邮件方案（导入 / 导出）](#邮件方案导入--导出)
- [自动发送（实验功能）](#自动发送实验功能)
- [大批量发送建议](#大批量发送建议)
- [CSV 格式](#csv-格式)
- [常见问题](#常见问题)
- [隐私和权限](#隐私和权限)
- [当前限制](#当前限制)
- [开发](#开发)

## 安装

扩展还没有上架 Chrome 应用商店，所以要以"未打包扩展"的方式安装，大约一分钟。

1. 打开 [最新版本（Releases）](https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/releases/latest)，下载 `mail-merge-draft-helper-<版本号>.zip`。
2. 解压到一个会长期保留的文件夹，比如 `Documents/mail-merge-draft-helper`。之后不要删掉这个文件夹，Chrome 是从这里加载扩展的。
3. 在 Chrome 地址栏打开 `chrome://extensions`（Edge 打开 `edge://extensions`）。
4. 打开 **Developer mode（开发者模式）**：Chrome 在右上角，Edge 在左侧栏。
5. 点击 **Load unpacked（加载已解压的扩展程序）**，选择刚才解压出来、里面有 `manifest.json` 的那个文件夹。

![chrome://extensions 页面上的开发者模式和 Load unpacked](docs/images/01-load-unpacked.png)

6. 点击工具栏上的拼图图标，把 **Mail Merge Draft Helper** 固定（📌）在工具栏上，方便一键打开。

**更新版本：** 下载新的 zip，把同一个文件夹里的文件替换掉，然后在 `chrome://extensions` 里点这个扩展卡片上的刷新按钮（↻）。

## 分步教程

弹窗里从上到下是三张编号卡片（收件人、模板、设置），下面是 **Generate Drafts** 按钮、预览区和打开草稿的按钮。

### 第 1 步：载入收件人名单

点击 **1. Recipient CSV Data** 里的虚线框，选择 CSV 文件，或者直接把文件拖进去。右上角会显示读入了多少行。也可以点 **Or paste / view raw CSV text**，直接粘贴 CSV 文本。

![载入 CSV 后显示 "3 rows loaded"](docs/images/03-load-csv.png)

CSV 第一行必须是表头，并且要有一列邮箱（列名可以是 `Email`、`E-mail`、`Mail`、`邮箱`、`To` 等）。其他每一列都可以当作模板变量。可以先用 [`examples/sample-recipients.csv`](examples/sample-recipients.csv) 试一试，格式说明见 [CSV 格式](#csv-格式)。

> 如果名单、主题、正文都已经准备好了，可以直接点 **Import Plan** 一次性导入，见 [邮件方案](#邮件方案导入--导出)。

### 第 2 步：写邮件模板

在 **2. Email Template** 里填写主题（Subject）和正文（Body）。**Click to insert variable** 区域会为每一列 CSV 生成一个按钮，点一下就会在光标处插入对应的变量。发送时每个 `{{变量}}` 都会换成这位收件人那一行的值。

![模板卡片和变量按钮](docs/images/05-template.png)

- 需要抄送时，点 **+ Add CC / BCC**。抄送栏里也可以用变量，比如 `{{CcEmail}}`。
- **Preset** 下拉菜单里有一个示例模板，可以在它的基础上改。
- 正文是纯文本：换行会保留，但不能加粗、改颜色或插图片。

### 第 3 步：选择邮箱服务和打开速度

在 **3. Compose Settings** 里：

- **Service Mode：** `Outlook Web`（outlook.office.com，学校或工作账号）或 `Gmail`。
- **Batch Opening Speed：** 批量打开时两封草稿之间的间隔。名单超过几十人时，建议选 `2.0s (Safe / Large)`。
- **Auto-Send Drafts：** 在测试成功之前先不要勾，见 [自动发送](#自动发送实验功能)。

![设置卡片](docs/images/06-settings.png)

### 第 4 步：生成并预览

点击 **Generate Drafts**，扩展会为每一行生成一封邮件并显示预览。用 **◀ Previous / Next ▶** 多翻几封检查：名字、主题、正文都应该已经替换好了。

![生成后的预览和翻页按钮](docs/images/07-preview.png)

如果有问题，会出现红色的检查提示，比如变量名写错（`{{Coursee}}`）、邮箱格式不对或重复、同一个地址同时出现在收件人和抄送里。改好 CSV 或模板后，再点一次 **Generate Drafts**。

### 第 5 步：打开草稿

- **Open Preview Draft**：打开当前正在预览的这一封。
- 想一次打开多封：在 **From / To** 里填行号（从 1 开始），再点 **Open Selected Drafts**。每封草稿会在后台单独打开一个标签页，间隔按第 3 步的设置。

![Open Preview Draft 和 From / To 范围](docs/images/08-open-drafts.png)

每个标签页都是正常的 Outlook 或 Gmail 写信窗口，收件人、主题、正文都已填好。检查后由你自己点 **Send** 发送。扩展会记住进度，关掉弹窗再打开时，范围会从上次停下的地方开始。

## 邮件方案（导入 / 导出）

邮件方案是一个 `.json` 文件，可以一次性准备好整个弹窗：收件人名单、主题、正文、CC、BCC、Outlook/Gmail 模式和打开速度。适合重复发同类邮件，或者由别人帮你准备好再交给你发送。

![Import Plan 和 Export Plan 按钮](docs/images/04-import-plan-bar.png)

- **Import Plan**（或把 `.json` 文件拖到 CSV 上传区）：填好所有字段并载入名单。检查模板后点 **Generate Drafts**。
- **Export Plan**：把当前设置保存为 `mail-merge-plan-YYYY-MM-DD.json`。文件里包含收件人名单，请妥善保管。
- 方案文件不会保存自动发送设置，导入后自动发送一律是关闭的。

![导入方案后的提示](docs/images/04-import-plan-status.png)

方案文件可以手写或修改。`csv` 和 `body` 可以写成一整段文字，也可以写成按行排列的数组：

```json
{
  "type": "mail-merge-plan",
  "version": 1,
  "name": "Sample event invite",
  "csv": ["FirstName,Email", "Ada,ada@example.com"],
  "subject": "You're invited, {{FirstName}}!",
  "body": ["Hi {{FirstName}},", "", "Hope to see you there!"],
  "settings": { "composeMode": "outlook", "batchDelay": 2000 }
}
```

`composeMode` 可以是 `outlook` 或 `gmail`；`batchDelay` 可以是 `500`、`1000` 或 `2000`（毫秒）。完整示例见 [`examples/sample-plan.json`](examples/sample-plan.json)。

## 自动发送（实验功能）

勾选 **Auto-Send Drafts** 后，扩展在打开每个写信标签页后还会自动点 **Send**，然后尝试关闭这个标签页。按钮文字会变成 **Open & Send Preview Email** 和 **Open & Send Selected Emails**，提醒你邮件会直接发出去。

![勾选自动发送](docs/images/09-auto-send.png)

![自动发送模式下的按钮](docs/images/09-auto-send-buttons.png)

使用前请注意：

- **先给自己发一封测试邮件。** 把第一行改成你自己的邮箱，只发第 1 行，确认收到了，并且写信标签页没有停在草稿状态。
- 它只会操作由扩展打开的写信标签页，而且只点对应那封草稿的发送按钮。
- 只认识写着 "Send" 或"发送"的按钮。如果你的 Outlook 或 Gmail 是其他语言，15 秒后会放弃，草稿留给你手动发送。
- Outlook 模式下不能和 CC/BCC 同时使用（带抄送的草稿走 `mailto:`）。请改用 Gmail 模式，或关掉自动发送。
- 扩展更新后，要在 `chrome://extensions` 里刷新一下，新版本才会生效。

## 大批量发送建议

- **每批 40–50 人左右。** 可以载入整份名单，先发第 1–50 行，过一段时间再发 51–100 行；也可以每批单独准备一个 CSV 或方案文件。
- **批次之间隔一段时间。** 短时间内发出大量邮件，可能会被邮箱服务限流或标记。
- **先发给自己测试**（第 1 行放自己的邮箱），用自动发送之前尤其要测。
- **打开速度选 `2.0s (Safe / Large)`**，让浏览器和 Outlook 来得及处理。
- **名单只保留需要的列。** 模板用不到的成绩、学号等敏感信息不要放进去。

## CSV 格式

CSV 需要一行表头、至少一行数据和一列邮箱。任何表头都可以当变量用：

```csv
FirstName,LastName,Email,Course,CcEmail
Ada,Lovelace,ada.lovelace@example.com,ISOM 210,ta@example.com
Grace,Hopper,grace.hopper@example.com,ISOM 340,ta@example.com
```

```text
Subject:  Reminder for {{Course}}
Body:     Hi {{FirstName}},

          This is a quick reminder about {{Course}}.
CC:       {{CcEmail}}
```

- 内容里有逗号或换行时，要用英文双引号括起来。
- 变量名区分大小写：`{{firstname}}` 对不上 `FirstName` 这一列。
- 直接粘贴时，表头和第一行数据之间必须是真正的换行；文本框里的自动折行不算。
- 用 Excel 准备名单时，选择 **文件 → 另存为 → CSV UTF-8**，中文和其他语言的名字才不会乱码。

## 常见问题

| 问题 | 解决办法 |
|---|---|
| 提示 "CSV needs a header row and at least one data row" | 确认第一行是表头，每位收件人单独一行。 |
| 检查提示 "Unknown template variable" | 变量名和 CSV 表头没有完全一致（拼写、大小写）。 |
| 邮件里直接出现了 `{{变量}}` | 同上，或者这一行对应的格子是空的。 |
| Outlook 打开了，但收件人或主题是空的 | 先在同一个浏览器里登录 Outlook Web，再重新打开草稿。 |
| 自动发送没有点发送 | 在 `chrome://extensions` 刷新扩展，检查发送按钮的语言，并参考 [自动发送](#自动发送实验功能)。 |
| 标签页打开太快，或有几封没打开 | 速度选 `2.0s (Safe / Large)`，每次少发一些。 |
| 想全部重来 | 点 **Clear Data**，会清除保存的名单、模板和进度。 |

## 隐私和权限

Mail Merge Draft Helper 没有服务器，不会把你的数据发到任何地方。CSV、模板、生成的草稿和进度保存在浏览器的扩展存储里，这样关掉弹窗再打开也不会丢。点 **Clear Data** 可以清除。导出的方案文件只会保存在浏览器的下载位置。

需要的权限：

- `storage`：在本地保存弹窗状态。
- Gmail 和 Outlook 的网站权限：让可选的自动发送脚本找到扩展打开的写信窗口。

扩展不会读取你的邮箱内容，也不使用 Gmail API 或 Microsoft Graph。完整隐私政策见 [store-assets/privacy-policy.md](store-assets/privacy-policy.md)。

## 当前限制

- 不支持附件（草稿是通过写信链接打开的）。
- 正文只支持纯文本，没有富文本格式。
- 正文太长时，可能超出写信链接能携带的长度。
- 自动发送是实验功能，依赖 Outlook / Gmail 页面的结构。

## 开发

1. 克隆这个仓库。
2. 用 **Load unpacked** 直接加载仓库文件夹本身（见 [安装](#安装)）。
3. 修改代码后，在扩展卡片上点刷新。

```bash
node --test
```

```bash
./scripts/build_release.sh
```

`build_release.sh` 会把扩展文件复制到 `dist/mail-merge-draft-helper/`，并按 `manifest.json` 里的版本号生成 `dist/mail-merge-draft-helper-<版本号>.zip`。

上架 Chrome 应用商店的资料（商店描述、截图、审核说明、检查清单）在 [`store-assets/`](store-assets/) 里。官方文档：

- [Publish in the Chrome Web Store](https://developer.chrome.com/docs/webstore/publish)
- [Chrome Web Store Developer Program Policies](https://developer.chrome.com/docs/webstore/program-policies)
- [Privacy disclosure requirements](https://developer.chrome.com/docs/webstore/program-policies/user-data-faq)
