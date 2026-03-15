# 思源笔记 ONLYOFFICE Bridge

在思源中直接预览和编辑 Office 文档（docx/xlsx/pptx/pdf 等），通过 ONLYOFFICE + Bridge 实现。

适用场景：
- ONLYOFFICE + Bridge 部署在公网服务器
- 思源桌面版或 Docker 网页版可能在私网

浏览器承担中转角色，因此 Bridge 不需要直连私网中的思源。

## 本次修复点

- 新增插件设置：**ONLYOFFICE 地址（可选）**
- 上传失败 `HTTP 404` 给出更明确提示
- Bridge 支持反向代理子路径（例如 `/bridge/upload`）
- Bridge 拆分 ONLYOFFICE 内外网地址：
  - `ONLYOFFICE_INTERNAL_URL`：Bridge 服务端连通用
  - `ONLYOFFICE_PUBLIC_URL`：浏览器加载 `api.js` 用

## 架构（Push 模式）

1. 插件从思源读取附件（浏览器 -> 思源）
2. 插件上传文档到 Bridge（`POST /upload`）
3. ONLYOFFICE 从 Bridge 读取文档（`GET /proxy/<asset>`）
4. 用户在 ONLYOFFICE 中保存
5. ONLYOFFICE 回调 Bridge（`POST /callback`）
6. 插件从 Bridge 拉取已保存版本（`GET /saved`）并写回思源

## 快速部署（公网服务器）

仓库内 `docker-compose.example.yml` 已更新，可直接参考。

Bridge 关键环境变量：

- `ONLYOFFICE_INTERNAL_URL=http://onlyoffice:80`
- `ONLYOFFICE_PUBLIC_URL=http://你的公网IP:8080`
- `BRIDGE_URL=http://你的公网IP:6789`
- `BRIDGE_SECRET=`（可选）

启动：

```bash
docker compose up -d
```

## 插件设置

在思源插件设置中填写：

- **Bridge 地址**（必填）
  - 例如：`http://你的公网IP:6789`
- **ONLYOFFICE 地址（可选）**
  - 例如：`http://你的公网IP:8080`
  - 留空则使用 Bridge 服务端配置
- **Bridge 密钥（可选）**
  - 需与服务端 `BRIDGE_SECRET` 一致

## 反向代理 / 子路径部署

如果你把 Bridge 暴露为子路径（例如 `https://example.com/bridge`）：

1. 插件 Bridge 地址填 `https://example.com/bridge`
2. Bridge 推荐设置 `BRIDGE_BASE_PATH=/bridge`
3. 或者 `BRIDGE_URL=https://example.com/bridge`

当前版本同时兼容根路径和子路径：
- `/upload` 与 `/bridge/upload`
- `/editor` 与 `/bridge/editor`
- 其余端点同理

## Bridge 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `BRIDGE_PORT` | `6789` | Bridge 监听端口 |
| `ONLYOFFICE_INTERNAL_URL` | `ONLYOFFICE_URL` 或 `http://127.0.0.1:8080` | Bridge 服务端访问 ONLYOFFICE 的地址 |
| `ONLYOFFICE_PUBLIC_URL` | 空 | 浏览器可访问的 ONLYOFFICE 地址（用于加载 `api.js`） |
| `BRIDGE_URL` | 空 | Bridge 对外地址（用于生成 callback/proxy 链接） |
| `BRIDGE_BASE_PATH` | 从 `BRIDGE_URL` path 推断 | 可选：反向代理子路径前缀 |
| `BRIDGE_SECRET` | 空 | 共享密钥 |
| `SIYUAN_URL` | 空 | 可选：Bridge 可直连思源时使用 |
| `SIYUAN_TOKEN` | 空 | 可选：思源 token |

## API 端点

- `GET /health`
- `POST /upload?asset=<path>`
- `GET /proxy/<path>`
- `GET /editor`
- `POST /callback`
- `GET /saved?asset=<path>`
- `POST /cleanup?asset=<path>`

这些端点在子路径模式下同样可用。

## 常见问题

### 上传失败：`Bridge returned HTTP 404`

通常是以下原因之一：

1. Bridge 地址填成了 ONLYOFFICE（`:8080`），而不是 Bridge（`:6789`）
2. 反向代理子路径未正确转发
3. 浏览器无法访问 Bridge

先检查：

```bash
curl http://你的Bridge地址/health
curl http://你的Bridge地址/health?detail=true
```

### 编辑器页面加载失败

Docker 场景下不要只配置 `ONLYOFFICE_INTERNAL_URL=http://onlyoffice:80`。
还需要设置 `ONLYOFFICE_PUBLIC_URL`，或者在插件里填写 ONLYOFFICE 地址。

### 编辑后没有回写到思源

- 检查 bridge 日志
- 确保 ONLYOFFICE 能访问 callback 地址
- 确保 `BRIDGE_URL` / 子路径与实际公网入口一致

## 安全建议

- 生产环境建议设置 `BRIDGE_SECRET`
- 建议通过 HTTPS 访问
- Bridge 已对附件路径做安全校验

## 许可证

MIT
