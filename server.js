
// server.js
const express = require("express");
const bodyParser = require("body-parser");

const app = express();
app.use(bodyParser.json());

// Microsoft Graph webhook callback
app.post("/api/callback", (req, res) => {
  // Step 1: Handle validation (Graph sends ?validationToken=... on subscription creation)
  if (req.query && req.query.validationToken) {
    console.log("Validation request received");
    res.status(200).send(req.query.validationToken);
    return;
  }

  // Step 2: Handle notifications
  if (req.body && req.body.value) {
    console.log("Received notifications:", JSON.stringify(req.body.value, null, 2));
    // TODO: store in DB or forward to frontend
  }

  res.sendStatus(202);
});

// Start server
const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Callback API running on http://localhost:${PORT}`));
