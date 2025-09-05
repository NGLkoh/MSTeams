// callback.ts
import express from "express";
import bodyParser from "body-parser";

const app = express();
app.use(bodyParser.json());

interface CallbackRequest extends express.Request {
  query: {
    validationToken?: string;
  };
  body: {
    value?: any[];
  };
}

// Microsoft Graph webhook callback
app.post("/api/callback", (req: CallbackRequest, res: express.Response) => {
  // Step 1: Handle validation
  if (req.query && req.query.validationToken) {
    console.log("Validation request received");
    res.status(200).send(req.query.validationToken);
    return;
  }

  // Step 2: Handle notifications
  if (req.body && req.body.value) {
    console.log("Received Graph notifications:", JSON.stringify(req.body.value, null, 2));
    // TODO: Save in DB or forward to frontend
  }

  res.sendStatus(202);
});

// Start server
const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Callback API running on http://localhost:${PORT}`));

// Export to make it a module
export default app;