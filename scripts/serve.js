// Dependency-free HTTPS static server for the add-in.
// Office add-ins must be served over HTTPS with a certificate the OS trusts;
// run `npx office-addin-dev-certs install` once to create + trust the certs,
// then `npm start`.
//
// Flags: --http (plain HTTP, for quick browser checks only), --port <n>

import https from "node:https";
import http from "node:http";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const args = process.argv.slice(2);
const useHttp = args.includes("--http");
const port = Number(args[args.indexOf("--port") + 1]) || 3000;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".xml": "application/xml",
  ".json": "application/json"
};

function handler(req, res) {
  const urlPath = decodeURIComponent(new URL(req.url, "http://x").pathname);
  let file = path.normalize(path.join(ROOT, urlPath));
  if (!file.startsWith(ROOT)) {
    res.writeHead(403).end("Forbidden");
    return;
  }
  if (urlPath === "/") file = path.join(ROOT, "src", "taskpane", "taskpane.html");
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404).end(`Not found: ${urlPath}`);
      return;
    }
    res.writeHead(200, {
      "Content-Type": MIME[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-cache"
    });
    res.end(data);
  });
}

if (useHttp) {
  http.createServer(handler).listen(port, () => {
    console.log(`SlideCharts dev server (HTTP — browser checks only): http://localhost:${port}`);
    console.log("Note: PowerPoint requires HTTPS; run without --http for sideloading.");
  });
} else {
  const crt = path.join(CERT_DIR, "localhost.crt");
  const key = path.join(CERT_DIR, "localhost.key");
  if (!fs.existsSync(crt) || !fs.existsSync(key)) {
    console.error(
      "Dev certificates not found.\n" +
      `Expected: ${crt}\n\n` +
      "Run this once to create and trust them (one OS prompt):\n" +
      "  npx office-addin-dev-certs install\n\n" +
      "Or run with --http for a quick browser-only check."
    );
    process.exit(1);
  }
  https
    .createServer({ cert: fs.readFileSync(crt), key: fs.readFileSync(key) }, handler)
    .listen(port, () => {
      console.log(`SlideCharts dev server: https://localhost:${port}`);
      console.log(`Taskpane: https://localhost:${port}/src/taskpane/taskpane.html`);
    });
}
