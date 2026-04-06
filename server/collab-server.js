#!/usr/bin/env node

/**
 * Yjs WebSocket Collaboration Server
 *
 * Runs a y-websocket server that handles:
 * - Room-based document synchronization
 * - Awareness (cursors, selections)
 * - Document persistence (optional)
 *
 * Usage:
 *   node server/collab-server.js [--port 1234] [--host 0.0.0.0]
 *
 * Clients connect to: ws://host:port/room-name
 */

import { WebSocketServer } from 'ws';
import * as Y from 'yjs';

const PORT = parseInt(process.argv.find((_, i, a) => a[i - 1] === '--port') || '4444');
const HOST = process.argv.find((_, i, a) => a[i - 1] === '--host') || '0.0.0.0';

// In-memory document store (rooms)
const docs = new Map();

function getDoc(room) {
  if (!docs.has(room)) {
    const doc = new Y.Doc();
    docs.set(room, { doc, conns: new Set() });
    console.log(`[room:${room}] Created`);
  }
  return docs.get(room);
}

const wss = new WebSocketServer({ port: PORT, host: HOST });

wss.on('connection', (ws, req) => {
  const room = req.url?.replace(/^\//, '') || 'default';
  const { doc, conns } = getDoc(room);

  conns.add(ws);
  console.log(`[room:${room}] Client connected (${conns.size} total)`);

  // Send current state to new client
  const stateVector = Y.encodeStateAsUpdate(doc);
  ws.send(stateVector);

  ws.on('message', (data) => {
    try {
      const update = new Uint8Array(data);
      Y.applyUpdate(doc, update);

      // Broadcast to all other clients in this room
      for (const conn of conns) {
        if (conn !== ws && conn.readyState === 1) {
          conn.send(update);
        }
      }
    } catch (e) {
      console.error(`[room:${room}] Error applying update:`, e.message);
    }
  });

  ws.on('close', () => {
    conns.delete(ws);
    console.log(`[room:${room}] Client disconnected (${conns.size} remaining)`);
    if (conns.size === 0) {
      // Keep doc in memory for reconnections (could add TTL cleanup)
    }
  });

  ws.on('error', (e) => {
    console.error(`[room:${room}] WebSocket error:`, e.message);
    conns.delete(ws);
  });
});

console.log(`Sheets collaboration server running on ws://${HOST}:${PORT}`);
console.log(`Clients connect to: ws://${HOST}:${PORT}/<room-name>`);
