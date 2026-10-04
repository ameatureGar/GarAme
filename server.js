const fs = require("fs");
const path = require("path");
const express = require("express");
const { createBillingServer } = require("./scripts/server");

for (const file of [
  path.join(__dirname, "data", "invoice-meta.json"),
  path.join(__dirname, "assets", "letterhead.jpg"),
  path.join(__dirname, "assets", "letterhead.png"),
  path.join(__dirname, "public", "index.html"),
  path.join(__dirname, "public", "styles.css"),
  path.join(__dirname, "public", "app.js"),
]) {
  fs.accessSync(file);
}

const app = express();
const { server } = createBillingServer();
const handle = server.listeners("request")[0];

app.use((req, res) => {
  handle(req, res);
});

module.exports = app;
