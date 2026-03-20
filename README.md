# 📝 Office Editor 

**Office Editor** is a SiYuan Note plugin that allows you to **preview and edit Office documents directly inside SiYuan** — including Word, Excel, PowerPoint, PDF, and 20+ formats, powered by ONLYOFFICE.

🌍 Document Language:  
[中文 README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_zh_CN.md) ｜ [English README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README.md)

---

## 🚀 Plugin Features

📄 **Rich Format Support**

Preview and edit docx, xlsx, pptx, pdf, odt, csv, rtf, txt, md, and 20+ other formats.

✏️ **Automatic Save Back**

Content edited in ONLYOFFICE will **automatically sync back to SiYuan attachments** — no manual action required.

📱 **Cross-Platform Support**

Supports preview + editing, and documents can be opened in popup, embedded, or tab mode.

---

## ✨ Usage

### 📌 Quick Start

1️⃣ Deploy ONLYOFFICE + Bridge service → See the [Docker Deployment Guide](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER.md)

2️⃣ Install this plugin from the SiYuan marketplace

3️⃣ Open the plugin settings and fill in the **Bridge address**

4️⃣ Right-click an Office attachment link in a document → Choose **Preview** or **Edit**

### 📌 Open Modes

- **Popup Mode**: Right-click menu → “Preview (Popup)” or “Edit (Popup)” to open in a popup
- **Embedded Mode**: Right-click menu → “Preview (Embed in Current Page)” or “Edit (Embed in Current Page)” to embed into the current document
- **Tab Mode**: Right-click menu → “Preview (New Tab)” or “Edit (New Tab)” to open in a SiYuan tab

---

## ⚙️ Plugin Settings

| Setting                | Description                                                  |
| ---------------------- | ------------------------------------------------------------ |
| **Bridge Address**     | Required. Default `http://127.0.0.1:27689` (same device only). For multi-device access, use a browser-reachable LAN address; for internet access, use reverse proxy |
| **ONLYOFFICE Address** | Optional. Default `http://127.0.0.1:27670` (same device only). For multi-device access, use a browser-reachable LAN address; for internet access, use reverse proxy. If empty, the Bridge server configuration will be used |
| **Bridge Secret**      | Optional. Must match the server `BRIDGE_SECRET`. Leave empty to disable |

---

## ⚠️ Notes

- **PDF files** support preview only, editing is not supported
- Edited files will **automatically sync back** to SiYuan attachments. Please ensure saving is completed before closing the editor
- When saving, the file **must not be occupied by other programs** such as local Office or WPS, otherwise saving may fail

---

## 📸 Video Demo (Hosted on GitHub, loading may require a proxy)

<video controls width="600">
  <source src="https://github.com/user-attachments/assets/dc9e24f1-21d3-4b4e-96e5-5c2013c26ef9" type="video/mp4">
  Your browser does not support the video tag.
</video>

---

## ☕ Support the Author

If you find this project helpful, feel free to support the author ❤️  
Your support will encourage me to **continue maintaining and improving** this tool.

<div align="center">
    <a href="https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge">
        <img src="https://img.shields.io/github/stars/b8l8u8e8/siyuan-plugin-onlyoffice-bridge?style=for-the-badge&color=ffd700&label=Give%20a%20Star" alt="Github Star">
    </a>
</div>

<div align="center" style="margin-top: 40px;">
    <div style="display: flex; justify-content: center; align-items: center; gap: 30px;">
        <div style="text-align: center;">
            <img src="https://github.com/user-attachments/assets/81d0a064-b760-4e97-9c9b-bf83f6cafc8a" 
                 style="height: 280px; width: auto; border-radius: 10px; border: 2px solid #07c160;">
            <br/>
            <b style="color: #07c160; margin-top: 10px; display: block;">WeChat Pay</b>
        </div>
        <div style="text-align: center;">
            <img src="https://github.com/user-attachments/assets/9e1988d0-4016-4b8d-9ea6-ce8ff714ee17" 
                 style="height: 280px; width: auto; border-radius: 10px; border: 2px solid #1677ff;">
            <br/>
            <b style="color: #1677ff; margin-top: 10px; display: block;">Alipay</b>
        </div>
    </div>
    <p style="margin-top: 20px;"><i>Your support is the biggest motivation for continuous iteration 🙏</i></p>
</div>

---

## 📖 More Information

- 🐳 Deployment Guide：[Docker Deployment Guide](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER.md) 
- 🐞 Issue Tracker：[GitHub Issues](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/issues)
- 📄 License：[MIT License](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/LICENSE)
- 🧾 Changelog：[CHANGELOG.md](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/CHANGELOG.md)
- 💖 Sponsor List: [Sponsor List](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/sponsor-list.md)
