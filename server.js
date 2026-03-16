"use strict";

const fs = require("fs");
const path = require("path");
const http = require("http");
const { translateDocumentWithAI } = require("./translate-service");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 8002);
const HOST = process.env.HOST || "127.0.0.1";

loadEnvFile(path.join(ROOT, ".env.local"));

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".ttf": "font/ttf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
};

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === "OPTIONS" && req.url === "/api/translate") {
      sendCors(res, 204);
      return;
    }

    if (req.method === "POST" && req.url === "/api/translate") {
      await handleTranslate(req, res);
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      sendJson(res, 405, { error: "Method not allowed" });
      return;
    }

    serveStatic(req, res);
  } catch (error) {
    console.error(error);
    sendJson(res, 500, { error: error.message || "Internal server error" });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`mySUNI Weekly Formatter server listening on http://${HOST}:${PORT}`);
});

async function handleTranslate(req, res) {
  const body = await readJsonBody(req);
  const document = await translateDocumentWithAI(body, process.env.GEMINI_API_KEY);
  sendJson(res, 200, { document }, true);
}

function serveStatic(req, res) {
  const requestPath = req.url === "/" ? "/index.html" : decodeURIComponent(req.url.split("?")[0]);
  const safePath = path.normalize(path.join(ROOT, requestPath));

  if (!safePath.startsWith(ROOT)) {
    sendJson(res, 403, { error: "Forbidden" });
    return;
  }

  fs.readFile(safePath, (error, data) => {
    if (error) {
      if (error.code === "ENOENT") {
        sendJson(res, 404, { error: "Not found" });
        return;
      }
      sendJson(res, 500, { error: error.message });
      return;
    }

    const ext = path.extname(safePath).toLowerCase();
    res.writeHead(200, { "Content-Type": MIME_TYPES[ext] || "application/octet-stream" });
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    res.end(data);
  });
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch (error) {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, status, payload, cors = false) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  };
  if (cors) {
    headers["Access-Control-Allow-Origin"] = "*";
    headers["Access-Control-Allow-Headers"] = "Content-Type";
    headers["Access-Control-Allow-Methods"] = "POST, OPTIONS";
  }
  res.writeHead(status, headers);
  res.end(JSON.stringify(payload));
}

function sendCors(res, status) {
  res.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Cache-Control": "no-store"
  });
  res.end();
}

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) {
    return;
  }

  const content = fs.readFileSync(filePath, "utf8");
  content.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      return;
    }
    const separator = trimmed.indexOf("=");
    if (separator < 0) {
      return;
    }
    const key = trimmed.slice(0, separator).trim();
    const value = trimmed.slice(separator + 1).trim();
    if (key && !(key in process.env)) {
      process.env[key] = value;
    }
  });
}
