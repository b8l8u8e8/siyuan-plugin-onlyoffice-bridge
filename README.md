# 📝 Office Editor

**Office Editor** is a SiYuan Note plugin that allows you to **preview and edit Office documents directly** within SiYuan—supporting Word, Excel, PowerPoint, PDF, and 20+ other formats, powered by ONLYOFFICE.

🌍 Documentation:
[中文 README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_zh_CN.md) ｜ [English README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README.md)

---

## 🚀 Key Features

📄 **Extensive Format Support**

Preview and edit over 20 formats including docx, xlsx, pptx, pdf, odt, csv, rtf, txt, md, and more.

✏️ **Auto-save & Sync**

Changes made in ONLYOFFICE are **automatically synced back to SiYuan assets**—no manual uploading or saving required.

📱 **Versatile Viewing Modes**

Full support for both previewing and editing. Documents can be opened in popups, embedded into the current page, or opened in a new tab.

---

## ✨ Usage

### 📌 Quick Start

1️⃣ Deploy the ONLYOFFICE + Bridge service on a public server → Refer to the [Docker Deployment Guide](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER_zh_CN.md).

2️⃣ Install this plugin from the SiYuan Marketplace.

3️⃣ Open plugin settings and enter your **Bridge Address**.

4️⃣ Right-click any Office attachment link in your document → Select **Preview** or **Edit**.

### 📌 Opening Methods

- **Popup Mode**: Right-click and select "Preview (Popup)" or "Edit (Popup)" to open in a floating window.
- **Embedded Mode**: Right-click and select "Preview (Embed)" or "Edit (Embed)" to insert the editor directly into your document.
- **Tab Mode**: Right-click and select "Preview (New Tab)" or "Edit (New Tab)" to open in a standard SiYuan tab.

---

## ⚙️ Plugin Settings

| Setting | Description |
|--------|------|
| **Bridge Address** | Required. The public URL of your Bridge service, e.g., `http://123.45.67.89:6789` |
| **ONLYOFFICE Address** | Optional. The browser-accessible URL for ONLYOFFICE, e.g., `http://123.45.67.89:7070`. Leave blank to use the Bridge server's default configuration. |
| **Bridge Secret** | Optional. Must match the `BRIDGE_SECRET` configured on your server. Leave blank if not enabled. |

---

## ⚠️ Important Notes

- **PDF files** are restricted to preview only; editing is not supported.
- Edited files are **automatically written back** to SiYuan assets. Please ensure the save process is complete before closing the editor.
- **Conflict Prevention**: Ensure the file is not being accessed by other local programs (e.g., Microsoft Office, WPS) during the save process to avoid synchronization failure.

---

## 📸 Video Demo (Source: GitHub)

<video controls width="600">
  <source src="https://github.com/user-attachments/assets/dc9e24f1-21d3-4b4e-96e5-5c2013c26ef9" type="video/mp4">
  Your browser does not support the video tag.
</video>

## 📖 More Information

- 🐳 Deployment: [Docker Deployment Guide](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER_zh_CN.md) 
- 🐞 Feedback: [GitHub Issues](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/issues)
- 📄 License: [MIT License](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/LICENSE)

---

## 📄 License

MIT
