#!/usr/bin/env node
const express = require("express");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { createProxyMiddleware } = require("http-proxy-middleware");

const app = express();

// Serve static add-in files from project root.
app.use(express.static(process.cwd()));

// Proxy Ollama through same origin to avoid mixed content + CORS.
app.use(
  "/ollama",
  createProxyMiddleware({
    target: "http://localhost:11434",
    changeOrigin: true,
    pathRewrite: { "^/ollama": "" }
  })
);

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const certPath = path.join(certDir, "localhost.crt");
const keyPath = path.join(certDir, "localhost.key");

const options = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath)
};

https.createServer(options, app).listen(3000, () => {
  console.log("Add-in dev server running at https://localhost:3000");
});
