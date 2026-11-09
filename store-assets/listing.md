# Chrome Web Store Listing (0.2.2)

Everything needed to fill in the Chrome Developer Dashboard for Mail Merge Draft Helper. Copy each block into the matching field.

| Dashboard field | Value |
|---|---|
| Package | `dist/mail-merge-draft-helper-0.2.2.zip` (or the zip on the [v0.2.2 release](https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/releases)) |
| Name | Mail Merge Draft Helper (from `_locales/*/messages.json`, same in every language) |
| Summary | Read from the package (`_locales/<lang>/messages.json` → `extDescription`); shown below for reference |
| Category | Productivity → Communication (or keep the current category) |
| Default language | English |
| Other listing languages | Chinese (Simplified), Spanish |
| Homepage URL | https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension |
| Support URL | https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/issues |
| Privacy policy URL | https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/blob/main/store-assets/privacy-policy.md |
| Screenshots (1280 × 800) | `store-assets/screenshots/<en, zh, es>/01-overview.png` … `05-private.png` |
| Small promo tile (440 × 280) | `store-assets/promo/<en, zh, es>/small-promo-tile-440x280.png` |
| Mature content | No |

---

## English (default)

**Summary** (from the package, 118 characters)

Personalized Outlook and Gmail emails from Excel or CSV lists. Preview, check, and send in batches. No sign-in needed.

**Description**

```text
Mail Merge Draft Helper turns an Excel or CSV list into personalized Outlook or Gmail emails, without signing in, without giving any app access to your mailbox, and without uploading your list anywhere.

It is built for teachers, administrators, and small teams who send many similar emails but still want to check them before they go out.

HOW IT WORKS
1. Recipients: load an Excel (.xlsx) or CSV file. Pick the email column and only the columns your message needs; grades, IDs, and everything else stay out.
2. Email: write one subject and body. Type {{ to insert a column, and add defaults such as {{First Name|there}}. Misspelled variables are flagged as you type.
3. Review: read every personalized email. Review notes point out names that may not match the address, unusual email domains, repeated addresses, and empty values, and you can fix them right there.
4. Send: send yourself a [TEST] copy, then open the drafts in batches (50 at a time by default). The extension remembers where you stopped.

New here? Click "Try it with sample data" to walk through all four steps with a made-up list.

MORE FEATURES
• Outlook on the web (Microsoft 365 work or school accounts) and Gmail
• Optional CC and BCC, including values from your list
• Save or load a whole mailing as an email plan file
• Optional Auto-Send (experimental) clicks Send in the drafts it opens
• Interface in English, Chinese, and Spanish

PRIVATE BY DESIGN
• No account, no sign-in, nothing to authorize in Microsoft 365 or Google
• Your list and templates are processed in your browser and kept only in Chrome's local extension storage
• No servers, no tracking, no analytics
• Clear Data removes everything at any time

GOOD TO KNOW
• Attachments and rich-text formatting are not supported, because drafts open through compose links
• Auto-Send is experimental: send yourself a test first

Free. Source code, step-by-step guide, and help: https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension
```

**Screenshots:** `store-assets/screenshots/en/` (5 images) · **Promo tile:** `store-assets/promo/en/small-promo-tile-440x280.png`

---

## Chinese (Simplified) / 中文（简体）

**Summary** (from the package)

从 Excel 或 CSV 名单生成个性化的 Outlook 和 Gmail 邮件。逐封预览、自动检查、分批发送，无需登录。

**Description**

```text
Mail Merge Draft Helper 可以把 Excel 或 CSV 名单变成一封封个性化的 Outlook 或 Gmail 邮件。无需登录，不需要授权任何应用访问你的邮箱，名单也不会上传到任何地方。

它适合老师、行政人员和小团队：要发很多封内容相似的邮件，但发出去之前还想自己检查一遍。

使用步骤
1. 名单：载入 Excel（.xlsx）或 CSV 文件。选好邮箱列，只载入邮件需要的列；成绩、学号等其他列都不会载入。
2. 内容：写一份主题和正文。输入 {{ 就能插入列名，还可以设默认值，比如 {{First Name|同学}}。变量写错会当场提示。
3. 检查：逐封查看个性化邮件。收件人检查会标出名字可能和邮箱对不上、邮箱后缀与众不同、同一邮箱出现多次、字段为空等情况，可以直接在弹窗里修改。
4. 发送：先给自己发一封 [TEST] 测试，再分批打开草稿（默认每批 50 封）。扩展会记住你发到了哪里。

第一次用？点"先用示例数据试一试"，用虚构的名单把四个步骤走一遍。

更多功能
• 支持 Outlook 网页版（Microsoft 365 学校或工作账号）和 Gmail
• 可选抄送、密送，也可以从名单中读取
• 把整次群发保存为邮件方案文件，随时导入
• 可选的自动发送（实验功能），会在打开的草稿里自动点击发送
• 界面支持 English、中文、Español

为隐私而设计
• 无需注册账号，无需登录，无需授权 Microsoft 365 或 Google
• 名单和模板只在你的浏览器里处理，只保存在 Chrome 的本地扩展存储中
• 没有服务器，没有跟踪，没有统计
• 随时可以用"清除数据"删除全部内容

须知
• 草稿通过写信链接打开，因此不支持附件和富文本格式
• 自动发送是实验功能，请先给自己发一封测试

免费使用。源代码、分步教程和帮助：https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension/blob/main/README.zh-CN.md
```

**Screenshots:** `store-assets/screenshots/zh/` · **Promo tile:** `store-assets/promo/zh/small-promo-tile-440x280.png`

---

## Spanish / Español

**Summary** (from the package)

Correos personalizados de Outlook y Gmail desde listas de Excel o CSV. Revisa y envía por lotes, sin iniciar sesión.

**Description**

```text
Mail Merge Draft Helper convierte una lista de Excel o CSV en correos personalizados de Outlook o Gmail, sin iniciar sesión, sin dar acceso a tu buzón a ninguna aplicación y sin subir tu lista a ningún sitio.

Está pensado para docentes, personal administrativo y equipos pequeños que envían muchos correos parecidos pero quieren revisarlos antes de enviarlos.

CÓMO FUNCIONA
1. Destinatarios: carga un archivo de Excel (.xlsx) o CSV. Elige la columna de correo y solo las columnas que necesita tu mensaje; las notas, los identificadores y lo demás quedan fuera.
2. Correo: escribe un asunto y un mensaje. Escribe {{ para insertar una columna y añade valores por defecto como {{First Name|estudiante}}. Las variables mal escritas se señalan mientras escribes.
3. Revisar: lee cada correo personalizado. Los avisos de revisión señalan nombres que quizá no coinciden con la dirección, dominios poco comunes, direcciones repetidas y valores vacíos, y puedes corregirlos ahí mismo.
4. Enviar: envíate una copia [TEST] y luego abre los borradores por lotes (50 cada vez por defecto). La extensión recuerda dónde te quedaste.

¿Primera vez? Pulsa "Pruébalo con datos de ejemplo" para recorrer los cuatro pasos con una lista inventada.

MÁS FUNCIONES
• Outlook en la web (cuentas de trabajo o centro educativo de Microsoft 365) y Gmail
• CC y CCO opcionales, también con valores de tu lista
• Guarda o carga un envío completo como archivo de plan
• Envío automático opcional (experimental): pulsa Enviar en los borradores que abre
• Interfaz en inglés, chino y español

PRIVADO POR DISEÑO
• Sin cuenta, sin iniciar sesión, nada que autorizar en Microsoft 365 ni Google
• Tu lista y tus plantillas se procesan en tu navegador y solo se guardan en el almacenamiento local de la extensión en Chrome
• Sin servidores, sin seguimiento, sin analíticas
• "Borrar datos" elimina todo en cualquier momento

TEN EN CUENTA
• No admite archivos adjuntos ni formato enriquecido, porque los borradores se abren mediante enlaces de redacción
• El envío automático es experimental: envíate una prueba primero

Gratis. Código fuente, guía paso a paso y ayuda: https://github.com/dimitriparadise/outlook-gmail-mail-merge-chrome--edge-browser-extension
```

**Screenshots:** `store-assets/screenshots/es/` · **Promo tile:** `store-assets/promo/es/small-promo-tile-440x280.png`

---

## Privacy practices tab

**Single purpose**

```text
Mail Merge Draft Helper creates personalized Outlook on the web or Gmail email drafts from a recipient list (Excel or CSV) and helps the user review, open, and optionally send those messages.
```

**Permission justifications**

`storage`

```text
Saves the popup's state locally (the recipient list the user loaded, the templates, generated drafts, sending progress, and settings) so the user can close and reopen the popup without losing work. Nothing is sent off the device.
```

Host permissions (`https://mail.google.com/*`, `https://outlook.office.com/*`, `https://outlook.live.com/*`)

```text
Used only by the optional Auto-Send feature, which is off by default. When the user turns it on, the content script finds the compose window that the extension itself opened (marked with mailMergeAutoSend=true in its URL) and clicks its Send button. It does not read the mailbox or any other email.
```

**Remote code:** No, I am not using remote code. (All JavaScript is in the package; no external scripts or eval.)

**Data usage:** do not tick any data type. The extension does not collect or transmit user data: lists, templates, and drafts stay in the browser's local extension storage, and the only network activity is opening Outlook or Gmail compose pages that the user asks for.

Tick all three certifications:

- I do not sell or transfer user data to third parties, outside of the approved use cases.
- I do not use or transfer user data for purposes that are unrelated to my item's single purpose.
- I do not use or transfer user data to determine creditworthiness or for lending purposes.
