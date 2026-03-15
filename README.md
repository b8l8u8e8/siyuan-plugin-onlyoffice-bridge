# ONLYOFFICE Bridge for SiYuan

Preview and edit Office files (docx/xlsx/pptx/pdf, etc.) in SiYuan using ONLYOFFICE.

This plugin is designed for:
- ONLYOFFICE + Bridge on a **public server**
- SiYuan desktop or Docker web on a **private network**

The browser relays file sync, so Bridge does not need direct access to private-network SiYuan.

## What changed in this version

- Added plugin setting: **ONLYOFFICE URL (optional)**
- Better error hint for upload `HTTP 404`
- Bridge now supports reverse-proxy sub-paths (example: `/bridge/upload`)
- Bridge supports separate ONLYOFFICE internal/public URLs:
  - `ONLYOFFICE_INTERNAL_URL` for bridge-side connectivity
  - `ONLYOFFICE_PUBLIC_URL` for browser-side `api.js` loading

## Architecture (push model)

1. Plugin reads attachment from SiYuan (browser -> SiYuan)
2. Plugin uploads file to Bridge (`POST /upload`)
3. ONLYOFFICE reads file from Bridge (`GET /proxy/<asset>`)
4. User saves in ONLYOFFICE
5. ONLYOFFICE callback -> Bridge (`POST /callback`)
6. Plugin pulls saved file (`GET /saved`) and writes back to SiYuan

## Quick deployment (public server)

Use `docker-compose.example.yml` in this repo.

Key environment variables for `bridge`:

- `ONLYOFFICE_INTERNAL_URL=http://onlyoffice:80`
- `ONLYOFFICE_PUBLIC_URL=http://YOUR_SERVER_IP:8080`
- `BRIDGE_URL=http://YOUR_SERVER_IP:6789`
- `BRIDGE_SECRET=` (optional)

Run:

```bash
docker compose up -d
```

## Plugin settings

Open plugin settings in SiYuan and set:

- **Bridge URL** (required)
  - Example: `http://YOUR_SERVER_IP:6789`
- **ONLYOFFICE URL (optional)**
  - Example: `http://YOUR_SERVER_IP:8080`
  - If empty, bridge server config is used.
- **Bridge secret (optional)**
  - Must match `BRIDGE_SECRET` on server.

## Reverse proxy / sub-path deployment

If Bridge is exposed as a sub-path (example `https://example.com/bridge`):

1. Set plugin **Bridge URL** to `https://example.com/bridge`
2. Set bridge env `BRIDGE_BASE_PATH=/bridge` (recommended)
3. Or set `BRIDGE_URL=https://example.com/bridge`

Bridge now accepts both root and prefixed endpoints:
- `/upload` and `/bridge/upload`
- `/editor` and `/bridge/editor`
- etc.

## Bridge environment variables

| Variable | Default | Description |
|---|---|---|
| `BRIDGE_PORT` | `6789` | Listen port |
| `ONLYOFFICE_INTERNAL_URL` | `ONLYOFFICE_URL` or `http://127.0.0.1:8080` | Bridge-side ONLYOFFICE URL |
| `ONLYOFFICE_PUBLIC_URL` | empty | Browser-side ONLYOFFICE URL for loading `api.js` |
| `BRIDGE_URL` | empty | External bridge URL used to generate callback/proxy links |
| `BRIDGE_BASE_PATH` | inferred from `BRIDGE_URL` path | Optional reverse-proxy path prefix |
| `BRIDGE_SECRET` | empty | Shared secret |
| `SIYUAN_URL` | empty | Optional direct SiYuan URL (co-located setup) |
| `SIYUAN_TOKEN` | empty | Optional SiYuan token |

## API endpoints

- `GET /health`
- `POST /upload?asset=<path>`
- `GET /proxy/<path>`
- `GET /editor`
- `POST /callback`
- `GET /saved?asset=<path>`
- `POST /cleanup?asset=<path>`

All endpoints also work under configured/prefixed base path.

## Troubleshooting

### Upload failed: `Bridge returned HTTP 404`

Usually one of these:

1. Bridge URL points to ONLYOFFICE (`:8080`) instead of Bridge (`:6789`)
2. Reverse proxy sub-path is not configured correctly
3. Bridge service is not reachable from your browser

Checks:

```bash
curl http://YOUR_BRIDGE/health
curl http://YOUR_BRIDGE/health?detail=true
```

### ONLYOFFICE editor page cannot load

If using Docker, do not rely on `ONLYOFFICE_INTERNAL_URL` for browser loading.
Set `ONLYOFFICE_PUBLIC_URL` to a browser-reachable address (or fill plugin ONLYOFFICE URL).

### Edits are not saved back

- Check bridge logs
- Ensure callback URL is externally reachable by ONLYOFFICE
- Keep `BRIDGE_URL` / base path consistent with actual public entry

## Security notes

- Set `BRIDGE_SECRET` in production
- Use HTTPS behind reverse proxy
- Asset paths are validated to prevent traversal

## License

MIT
