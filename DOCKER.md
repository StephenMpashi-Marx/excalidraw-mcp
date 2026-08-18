# Excalidraw MCP Server — Docker Setup

Run the Excalidraw MCP server in a Docker container using Colima (or Docker Desktop).

## Prerequisites

- [Colima](https://github.com/abiosoft/colima) or Docker Desktop
- Docker CLI (`brew install docker`)
- Node.js and pnpm (for building from source)

## Quick Start

### 1. Start Colima

```bash
colima start
```

### 2. Build the project

```bash
pnpm install && pnpm run build
```

### 3. Build the Docker image

```bash
docker build -t excalidraw-mcp .
```

### 4. Run the container

```bash
docker run -d --name excalidraw-mcp -p 3001:3001 excalidraw-mcp
```

The MCP server will be available at `http://localhost:3001/mcp`.

## Verify it works

```bash
curl -X POST http://localhost:3001/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0.0"}}}'
```

You should receive a response containing `serverInfo: { name: "Excalidraw", version: "1.0.0" }`.

## Cursor IDE Configuration

Add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "Excalidraw": {
      "url": "http://localhost:3001/mcp"
    }
  }
}
```

Reload Cursor after updating the config.

## Container Management

```bash
docker logs excalidraw-mcp       # View logs
docker stop excalidraw-mcp       # Stop the server
docker start excalidraw-mcp      # Restart the server
docker rm -f excalidraw-mcp      # Remove the container
```

## Rebuilding after code changes

```bash
pnpm run build
docker rm -f excalidraw-mcp
docker build -t excalidraw-mcp .
docker run -d --name excalidraw-mcp -p 3001:3001 excalidraw-mcp
```

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `3001` | HTTP port the server listens on |

To use a custom port:

```bash
docker run -d --name excalidraw-mcp -p 8080:8080 -e PORT=8080 excalidraw-mcp
```

## Troubleshooting

**Port already in use:**

```bash
lsof -ti:3001 | xargs kill -9
```

**Colima not running:**

```bash
colima status
colima start
```

**Container won't start:**

```bash
docker logs excalidraw-mcp
```
