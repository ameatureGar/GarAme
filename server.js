const express = require("express");
const { createBillingServer } = require("./scripts/server");

const app = express();
const { server } = createBillingServer();
const handle = server.listeners("request")[0];

app.use((req, res) => {
  handle(req, res);
});

module.exports = app;
