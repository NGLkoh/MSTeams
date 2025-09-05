const express = require('express');
const next = require('next');
const bodyParser = require('body-parser');
const interviewRouter = require('./interview'); // Adjust path if needed
require('dotenv').config();

const port = process.env.PORT || 3000;
const dev = process.env.NODE_ENV !== 'production';
const app = next({ dev });
const handle = app.getRequestHandler();

app.prepare().then(() => {
  const server = express();

  // Middleware
  server.use(express.json());
  server.use(bodyParser.json()); // For webhook parsing

  // Microsoft Graph webhook callback endpoint
  server.post("/api/callback", (req, res) => {
    console.log('Webhook callback received:', {
      method: req.method,
      url: req.url,
      query: req.query,
      headers: req.headers,
      body: req.body
    });

    // Step 1: Handle validation (Graph sends ?validationToken=... on subscription creation)
    if (req.query && req.query.validationToken) {
      console.log("Validation request received, token:", req.query.validationToken);
      return res.status(200).send(req.query.validationToken);
    }

    // Step 2: Handle notifications
    if (req.body && req.body.value) {
      console.log("Received notifications:", JSON.stringify(req.body.value, null, 2));
      
      // Process each notification
      req.body.value.forEach(notification => {
        console.log('Processing notification:', {
          subscriptionId: notification.subscriptionId,
          changeType: notification.changeType,
          resource: notification.resource,
          resourceData: notification.resourceData,
          clientState: notification.clientState
        });
        
        // TODO: Process the notification
        // - Update your local calendar cache
        // - Notify connected clients via WebSocket/SSE
        // - Store in database
        // - Forward to frontend
      });
    }

    // Always return 202 to acknowledge receipt
    res.sendStatus(202);
  });

  // Handle GET requests to callback (for testing/debugging)
  server.get("/api/callback", (req, res) => {
    console.log('GET request to callback endpoint');
    
    if (req.query && req.query.validationToken) {
      console.log("Validation token via GET:", req.query.validationToken);
      return res.status(200).send(req.query.validationToken);
    }
    
    res.json({ 
      status: 'Webhook callback endpoint is active',
      timestamp: new Date().toISOString(),
      method: 'GET'
    });
  });

  // Interview router
  server.use('/api/interview', interviewRouter);

  // Health check endpoint
  server.get('/api/health', (req, res) => {
    res.json({ 
      status: 'healthy', 
      timestamp: new Date().toISOString(),
      port: port,
      environment: process.env.NODE_ENV || 'development'
    });
  });

  // Webhook status endpoint for debugging
  server.get('/api/webhook/status', (req, res) => {
    res.json({
      callbackUrl: `${req.protocol}://${req.get('host')}/api/callback`,
      ready: true,
      timestamp: new Date().toISOString()
    });
  });

  // Handle all other requests with Next.js
  server.all('*', (req, res) => {
    return handle(req, res);
  });

  // Error handling middleware
  server.use((err, req, res, next) => {
    console.error('Server error:', err);
    
    if (req.url.startsWith('/api/callback')) {
      // For webhook endpoints, always return success to avoid retries
      return res.sendStatus(202);
    }
    
    res.status(500).json({ error: 'Internal server error' });
  });

  server.listen(port, (err) => {
    if (err) throw err;
    console.log(`> Ready on http://localhost:${port}`);
    console.log(`> Webhook callback available at: http://localhost:${port}/api/callback`);
    console.log(`> Health check at: http://localhost:${port}/api/health`);
    console.log(`> Environment: ${process.env.NODE_ENV || 'development'}`);
  });

  // Graceful shutdown
  process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    server.close(() => {
      console.log('Process terminated');
    });
  });

  process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    server.close(() => {
      console.log('Process terminated');
    });
  });
});