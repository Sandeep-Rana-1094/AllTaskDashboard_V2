import express from "express";
import { createServer } from "http";
import { WebSocketServer, WebSocket } from "ws";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const server = createServer(app);
  const wss = new WebSocketServer({ server });

  const PORT = 3000;

  // Presence tracking
  // NOTE: In-memory state like Maps will be reset on Vercel/serverless environments.
  // For persistent presence tracking across instances, use a database (e.g., Redis, MongoDB).
  const clients = new Map<WebSocket, { email: string; role: string }>();
  const dailyLogins = new Map<string, { role: string; lastSeen: string }>();

  app.use(express.json());

  function getAllUsers() {
    const liveEmails = new Set(Array.from(clients.values()).map(c => c.email));
    return Array.from(dailyLogins.entries()).map(([email, info]) => ({
      email,
      role: info.role,
      lastSeen: info.lastSeen,
      isLive: liveEmails.has(email)
    }));
  }

  app.post("/api/presence/check-in", (req, res) => {
    const { email, role } = req.body;
    if (email && role) {
      dailyLogins.set(email, { role, lastSeen: new Date().toISOString() });
      broadcastPresence();
      res.json({ success: true });
    } else {
      res.status(400).json({ error: "Missing email or role" });
    }
  });

  app.get("/api/presence/state", (req, res) => {
    const allUsers = getAllUsers();
    res.json({
      liveUsers: allUsers.filter(u => u.isLive),
      dailyUsers: allUsers
    });
  });

  wss.on("connection", (ws) => {
    ws.on("message", (message) => {
      try {
        const data = JSON.parse(message.toString());
        if (data.type === "identify") {
          clients.set(ws, { email: data.email, role: data.role });
          dailyLogins.set(data.email, { role: data.role, lastSeen: new Date().toISOString() });
          broadcastPresence();
        }
      } catch (e) {
        console.error("Failed to parse message", e);
      }
    });

    ws.on("close", () => {
      clients.delete(ws);
      broadcastPresence();
    });
  });

  function broadcastPresence() {
    const allUsers = getAllUsers();
    const message = JSON.stringify({ 
      type: "presence", 
      liveUsers: allUsers.filter(u => u.isLive),
      dailyUsers: allUsers
    });
    wss.clients.forEach((client) => {
      if (client.readyState === WebSocket.OPEN) {
        client.send(message);
      }
    });
  }

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  server.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
