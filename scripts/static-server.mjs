#!/usr/bin/env node
import fs from "node:fs";
import http from "node:http";
import path from "node:path";

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".mjs": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".ttf": "font/ttf",
  ".otf": "font/otf",
};

function usage() {
  console.log(`Usage:
  node scripts/static-server.mjs [root] [--port 3010] [--host 127.0.0.1] [--no-cors]

Examples:
  node scripts/static-server.mjs public/packages --port 3010
  node scripts/static-server.mjs dist --host 0.0.0.0 --port 8080
`);
}

function parseArgs(argv) {
  const options = {
    root: "public/packages",
    port: Number(process.env.STATIC_SERVER_PORT || 3010),
    host: process.env.STATIC_SERVER_HOST || "127.0.0.1",
    cors: true,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      usage();
      process.exit(0);
    }
    if (arg === "--port" || arg === "-p") {
      options.port = Number(argv[++i]);
      continue;
    }
    if (arg === "--host" || arg === "-H") {
      options.host = argv[++i];
      continue;
    }
    if (arg === "--no-cors") {
      options.cors = false;
      continue;
    }
    if (!arg.startsWith("-")) {
      options.root = arg;
    }
  }

  return {
    ...options,
    root: path.resolve(process.cwd(), options.root),
  };
}

function contentType(filePath) {
  return MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream";
}

function writePlain(res, status, body, options) {
  const headers = {
    "Content-Type": "text/plain; charset=utf-8",
    "Cache-Control": "no-cache",
  };
  if (options.cors) {
    headers["Access-Control-Allow-Origin"] = "*";
    headers["Access-Control-Allow-Methods"] = "GET,HEAD,OPTIONS";
    headers["Access-Control-Allow-Headers"] = "*";
  }
  res.writeHead(status, headers);
  res.end(body);
}

function createServer(options) {
  return http.createServer((req, res) => {
    if (req.method === "OPTIONS") {
      writePlain(res, 204, "", options);
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      writePlain(res, 405, "Method Not Allowed", options);
      return;
    }

    const urlPath = (req.url || "/").split("?")[0];
    let pathname;
    try {
      pathname = decodeURIComponent(urlPath);
    } catch {
      writePlain(res, 400, "Bad Request", options);
      return;
    }

    let filePath = path.resolve(options.root, pathname.replace(/^\/+/, ""));
    if (filePath !== options.root && !filePath.startsWith(options.root + path.sep)) {
      writePlain(res, 403, "Forbidden", options);
      return;
    }

    fs.stat(filePath, (statError, stat) => {
      if (!statError && stat.isDirectory()) {
        filePath = path.join(filePath, "index.html");
      }

      fs.stat(filePath, (fileError, fileStat) => {
        if (fileError || !fileStat.isFile()) {
          writePlain(res, 404, "Not Found", options);
          return;
        }

        const headers = {
          "Content-Type": contentType(filePath),
          "Content-Length": String(fileStat.size),
          "Cache-Control": "no-cache",
        };
        if (options.cors) {
          headers["Access-Control-Allow-Origin"] = "*";
          headers["Access-Control-Allow-Methods"] = "GET,HEAD,OPTIONS";
          headers["Access-Control-Allow-Headers"] = "*";
        }

        res.writeHead(200, headers);
        if (req.method === "HEAD") {
          res.end();
          return;
        }
        fs.createReadStream(filePath).pipe(res);
      });
    });
  });
}

const options = parseArgs(process.argv.slice(2));

if (!Number.isFinite(options.port) || options.port <= 0) {
  console.error("[static-server] invalid port");
  process.exit(1);
}

if (!fs.existsSync(options.root) || !fs.statSync(options.root).isDirectory()) {
  console.error(`[static-server] root is not a directory: ${options.root}`);
  process.exit(1);
}

const server = createServer(options);

server.on("error", (error) => {
  if (error && error.code === "EADDRINUSE") {
    console.error(`[static-server] port is already in use: ${options.port}`);
    process.exit(1);
  }
  throw error;
});

server.listen(options.port, options.host, () => {
  console.log(`[static-server] root ${options.root}`);
  console.log(`[static-server] 用于测试cdn加载：资源填入 http://${options.host}:${options.port}`);
});
