# 🐳 Office 编辑器 Bridge — Docker 部署指南

本文档介绍如何在公网服务器上部署 **ONLYOFFICE + Bridge** 服务，供思源笔记 Office 编辑器插件使用。

---

## 📋 前置要求

- 一台拥有**公网 IP** 的服务器
- 已安装 **Docker** 和 **Docker Compose**
- 确保端口 `7070`（ONLYOFFICE）和 `6789`（Bridge）对外可访问

> 如果端口已被其他服务占用，可以在 `docker-compose.yml` 中修改端口映射后再启动。

```bash
# 检查 Docker 是否已安装
docker -v
docker compose version
```

如果未安装 Docker，可参考以下命令：

```bash
# Linux (Ubuntu/Debian) 一键安装 Docker
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER

# 安装完后需要重新登录终端才能生效
# 或访问 https://docs.docker.com/get-docker/
```

---

## 🚀 快速启动

### 1. 获取源码文件

将仓库克隆到服务器上。如果不使用 Git，也可以直接下载仓库的 ZIP 压缩包，解压后放到当前目录，效果一样。

```bash
git clone https://github.com/b8l8u8e8/siyuan-plugin-onlyoffice-bridge.git
cd siyuan-plugin-onlyoffice-bridge
```

### 2. 修改配置

编辑 `docker-compose.yml`，将 `YOUR_SERVER_IP` **替换为你服务器的公网 IP**：

```yaml
services:
  onlyoffice:
    image: onlyoffice/documentserver:latest
    ports:
      - "7070:80"                    # 宿主机端口:容器端口
    volumes:
      - ./server/onlyoffice-bootstrap.sh:/app/onlyoffice-bootstrap.sh:ro
    entrypoint:
      - /bin/bash
      - /app/onlyoffice-bootstrap.sh
    environment:
      - JWT_ENABLED=false            
      - ALLOW_PRIVATE_IP_ADDRESS=true
      - ALLOW_META_IP_ADDRESS=true
      - MAX_DOWNLOAD_BYTES=536870912 # ONLYOFFICE 最大下载文件大小（字节），默认 512MB
    restart: unless-stopped

  bridge:
    build: ./server
    ports:
      - "6789:6789"                  # 宿主机端口:容器端口
    environment:
      - BRIDGE_PORT=6789
      - ONLYOFFICE_INTERNAL_URL=http://onlyoffice:80
      - ONLYOFFICE_PUBLIC_URL=http://127.0.0.1:7070   
      - BRIDGE_URL=http://bridge:6789
      - BRIDGE_SECRET=               # 可选：填写密钥后插件设置也需对应填写
    depends_on:
      - onlyoffice
    restart: unless-stopped
```

### 3. 启动服务

```bash
docker compose up -d
```

首次启动需要**构建 Bridge 镜像 + 拉取 ONLYOFFICE 镜像**，可能需要几分钟，请耐心等待。

> ONLYOFFICE 服务启动较慢（约 1-2 分钟），启动后才能正常使用。

### 4. 验证服务是否正常

请配置反向代理，将本机 `127.0.0.1` 的 `6789` 与 `7070` 端口映射到公网访问。

```bash
# 检查 Bridge 是否正常运行
curl http://你的公网IP:6789/health

# 查看详细连接状态（包括 ONLYOFFICE 连接是否正常）
curl http://你的公网IP:6789/health?detail=true
```

如果返回类似 `{"status":"ok"}` 的内容，说明服务已正常运行。

### 5. 配置思源插件

在思源笔记中打开 Office 编辑器插件设置，填写以下内容：

| 插件设置项 | 填什么 | 说明 |
|------------|--------|------|
| **Bridge 地址** | `http://你的公网IP:6789` | 必填，指向 Bridge 服务 |
| **ONLYOFFICE 地址** | `http://你的公网IP:7070` | 可选，留空也行（Bridge 会自动提供） |
| **Bridge 密钥** | 与 `BRIDGE_SECRET` 一致 | 可选，留空则不启用 |

配置完成后，在文档中右键点击 Office 附件即可预览或编辑。

---

## 🔧 环境变量详解

以下是 `docker-compose.yml` 中 Bridge 服务可配置的**所有环境变量**：

### Bridge 核心配置

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `BRIDGE_PORT` | `6789` | Bridge 服务监听的端口号 |
| `BRIDGE_URL` | 空 | Bridge 对外访问地址，用于生成 ONLYOFFICE 的回调链接。Docker 部署中通常填 `http://bridge:6789`（容器间通信） |
| `BRIDGE_BASE_PATH` | 从 `BRIDGE_URL` 推断 | 反向代理子路径前缀，例如 `/bridge`。通常不需要手动设置 |
| `BRIDGE_SECRET` | 空 | 共享密钥。设置后，插件设置中也需填写相同的密钥才能连接 |

### ONLYOFFICE 连接配置

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `ONLYOFFICE_INTERNAL_URL` | `http://127.0.0.1:7070` | Bridge **内部**访问 ONLYOFFICE 的地址。Docker 部署中填 `http://onlyoffice:80`（容器间通信） |
| `ONLYOFFICE_PUBLIC_URL` | 空 | **浏览器**访问 ONLYOFFICE 的地址（用于加载编辑器 `api.js`）。填你的公网地址，如 `http://123.45.67.89:7070` |

### 文件大小限制

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `MAX_FILE_MB` | `512` | Bridge 允许上传的最大文件大小（MB） |
| `MAX_CHUNK_MB` | `8` | 分片上传时每个分片的最大大小（MB） |

### ONLYOFFICE 容器配置

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `JWT_ENABLED` | `false` | 是否启用 ONLYOFFICE JWT 验证。生产环境建议开启 |
| `ALLOW_PRIVATE_IP_ADDRESS` | `true` | 允许 ONLYOFFICE 访问内网 IP（Docker 容器间通信需要） |
| `ALLOW_META_IP_ADDRESS` | `true` | 允许 ONLYOFFICE 访问元数据 IP（如 `169.254.x.x`） |
| `MAX_DOWNLOAD_BYTES` | `536870912` | ONLYOFFICE 最大下载文件大小，单位字节（默认 512MB） |

---

## 🔄 更新应用

```bash
# 拉取最新源码
# 如果不使用 Git，也可以直接下载仓库的最新压缩包，解压后覆盖当前目录中的源码。
#
# 如果拉取时提示冲突（conflict），说明：
# 你本地修改过的文件，刚好也在云端被更新了。
# 为了防止你的修改被直接覆盖，Git 会要求你先确认如何处理。
# 这是正常现象，不是报错。
# 根据提示处理完成后，再次执行 git pull 即可。
git pull

# 停止服务
docker compose down

# 重新构建 Bridge 镜像并启动
docker compose build bridge
docker compose up -d
```

---

## 🔧 常用命令

```bash
# 启动服务
docker compose up -d

# 查看所有容器状态
docker compose ps

# 查看 Bridge 日志
docker compose logs -f bridge

# 查看 ONLYOFFICE 日志
docker compose logs -f onlyoffice

# 重启所有服务
docker compose restart

# 停止所有服务
docker compose down

# 重新构建 Bridge 镜像
docker compose build bridge && docker compose up -d bridge
```

---

## 🐛 常见问题

### 上传失败：`Bridge returned HTTP 404`

**可能原因**：
1. Bridge 地址填错了——填成了 ONLYOFFICE 的地址（`:7070`），应该填 Bridge 的地址（`:6789`）
3. 浏览器无法访问 Bridge 服务

**排查方法**：
```bash
curl http://你的Bridge地址/health
curl http://你的Bridge地址/health?detail=true
```

### 编辑器页面加载失败

Docker 部署下需要同时配置：
- `ONLYOFFICE_INTERNAL_URL`：容器内通信地址（如 `http://onlyoffice:80`）
- `ONLYOFFICE_PUBLIC_URL`：浏览器可访问的公网地址（如 `http://123.45.67.89:7070`）

或者在插件设置中手动填写 ONLYOFFICE 地址。

### 编辑后内容没有回写到思源

1. 检查 Bridge 日志：`docker compose logs bridge`
2. 确保 ONLYOFFICE 能访问 Bridge 的 callback 地址
3. 确保 `BRIDGE_URL` 填写正确（Docker 部署中为 `http://bridge:6789`）

### ONLYOFFICE 启动后一直连不上

ONLYOFFICE 服务启动较慢，首次启动可能需要 1-2 分钟。可以通过以下命令查看启动进度：

```bash
docker compose logs -f onlyoffice
```

看到类似 `Starting ONLYOFFICE Document Server...` 后面出现 `Running...` 就表示启动完成。

---

## 🔒 安全建议

- 生产环境建议设置 `BRIDGE_SECRET`，防止未授权访问

---

## 🔌 API 端点参考

| 端点 | 方法 | 说明 |
|------|------|------|
| `/health` | GET | 健康检查（加 `?detail=true` 获取连接详情） |
| `/upload?asset=<path>` | POST | 插件上传文件到 Bridge |
| `/proxy/<path>` | GET | ONLYOFFICE 从 Bridge 获取文件 |
| `/oo/<path>` | GET | 代理 ONLYOFFICE 静态资源（api.js 等） |
| `/editor` | GET | 编辑器 HTML 页面 |
| `/callback` | POST | ONLYOFFICE 保存回调 |
| `/saved?asset=<path>` | GET | 插件拉取已保存的文件 |
| `/cleanup?asset=<path>` | POST | 清理 Bridge 内存中的文件缓存 |
