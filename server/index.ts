import https from 'node:https';
import express from 'express';
import { createApp } from './app';
import { loadTls } from './tls';

const app = createApp();
if (process.env.NODE_ENV === 'production' || import.meta.url.includes('/dist-server/')) {
  app.use(express.static('dist'));
  const port = Number(process.env.PORT || 3001);
  https
    .createServer(loadTls(), app)
    .listen(port, 'localhost', () => console.log(`WordAgent: https://localhost:${port}`));
} else {
  const port = Number(process.env.API_PORT || 4318);
  app.listen(port, '127.0.0.1', () => console.log(`WordAgent API: http://127.0.0.1:${port}`));
}
