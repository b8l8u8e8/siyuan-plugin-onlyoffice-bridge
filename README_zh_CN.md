# 📝 Office 编辑器（Office Editor）

**Office 编辑器** 是一款思源笔记插件，让你在思源中**直接预览和编辑 Office 文档**——Word、Excel、PowerPoint、PDF 等 20+ 格式，基于 ONLYOFFICE 实现。

🌍 文档语言：
[中文 README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_zh_CN.md) ｜ [English README](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README.md)

---

## 🚀 插件优势

📄 **丰富的格式支持**

预览和编辑 docx、xlsx、pptx、pdf、odt、csv、rtf、txt、md 等 20+ 格式。

✏️ **自动保存回写**

在 ONLYOFFICE 中编辑的内容会**自动同步回思源附件**——无需手动操作。

📱 **全平台覆盖**

支持预览 + 编辑，可在弹窗、嵌入当前页或标签页中打开文档。

---

## ✨ 使用方法

### 📌 快速开始

1️⃣ 部署 ONLYOFFICE + Bridge 服务→ 参考 [Docker 部署指南](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER.md)

2️⃣ 在思源集市安装本插件

3️⃣ 打开插件设置，填写 **Bridge 地址**

4️⃣ 在文档中右键点击 Office 附件链接 → 选择**预览**或**编辑**

### 📌 打开方式

- **弹窗模式**：右键菜单选择「预览（弹窗）」或「编辑（弹窗）」，在弹窗中打开
- **嵌入模式**：右键菜单选择「预览（嵌入当前页）」或「编辑（嵌入当前页）」，嵌入到当前文档中
- **标签页模式**：右键菜单选择「预览（新标签页）」或「编辑（新标签页）」，在思源标签页中打开

---

## ⚙️ 插件设置

| 设置项 | 说明 |
|--------|------|
| **Bridge 地址** | 必填。默认 `http://127.0.0.1:27689`（仅当前设备可用）；跨设备请改为浏览器可访问的内网地址，公网请走反向代理 |
| **ONLYOFFICE 地址** | 可选。默认 `http://127.0.0.1:27670`（仅当前设备可用）；跨设备请改为浏览器可访问的内网地址，公网请走反向代理。留空则使用 Bridge 服务端配置 |
| **Bridge 密钥** | 可选。需与服务端 `BRIDGE_SECRET` 一致，留空则不启用 |

---

## ⚠️ 注意事项

- **PDF 文件**仅支持预览，不支持编辑
- 编辑后的文件会**自动回写**到思源附件，关闭编辑器前请确保保存完毕
- 保存时，对应文件请**不要有其他程序占用**，例如本机的office程序、WPS，会导致保存失败

---

## 📸 视频演示（视频源自 GitHub，加载可能需“魔法”）

<video controls width="600">
  <source src="https://github.com/user-attachments/assets/dc9e24f1-21d3-4b4e-96e5-5c2013c26ef9" type="video/mp4">
  Your browser does not support the video tag.
</video>

---

## ☕ 支持作者

如果你觉得这个项目对你有帮助，欢迎支持作者 ❤️  
你的支持将激励我 **持续维护与优化**，打造更好用的工具。

<div align="center">
    <a href="https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge">
        <img src="https://img.shields.io/github/stars/b8l8u8e8/siyuan-plugin-onlyoffice-bridge?style=for-the-badge&color=ffd700&label=%E7%BB%99%E4%B8%AAStar%E5%90%A7" alt="Github Star">
    </a>
</div>


<div align="center" style="margin-top: 40px;">
    <div style="display: flex; justify-content: center; align-items: center; gap: 30px;">
        <div style="text-align: center;">
            <img src="https://github.com/user-attachments/assets/81d0a064-b760-4e97-9c9b-bf83f6cafc8a" 
                 style="height: 280px; width: auto; border-radius: 10px; border: 2px solid #07c160;">
            <br/>
            <b style="color: #07c160; margin-top: 10px; display: block;">微信支付</b>
        </div>
        <div style="text-align: center;">
            <img src="https://github.com/user-attachments/assets/9e1988d0-4016-4b8d-9ea6-ce8ff714ee17" 
                 style="height: 280px; width: auto; border-radius: 10px; border: 2px solid #1677ff;">
            <br/>
            <b style="color: #1677ff; margin-top: 10px; display: block;">支付宝</b>
        </div>
    </div>
    <p style="margin-top: 20px;"><i>你的支持，是我持续迭代的最大动力 🙏</i></p>
</div>

---

## 📖 更多信息

- 🐳 部署指南：[Docker 部署指南](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/README_DOCKER.md) 
- 🐞 问题反馈：[GitHub Issues](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/issues)
- 📄 开源协议：[MIT License](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/LICENSE)
- 🧾 更新日志：[CHANGELOG.md](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/CHANGELOG.md)
- 💖 赞助列表：[Sponsor List](https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge/blob/main/sponsor-list.md)

