import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';

// Serve only the compiled app. Never expose input/, output/, or the repo.
const host = '127.0.0.1';
const port = 4173;
const server = createServer(async (req, res) => {
  if (req.url !== '/' && req.url !== '/index.html') {
    res.writeHead(404); res.end('Not found'); return;
  }
  try {
    const html = await readFile(new URL('../index.html', import.meta.url));
    res.writeHead(200, {'Content-Type':'text/html; charset=utf-8', 'Cache-Control':'no-store'});
    res.end(html);
  } catch {
    res.writeHead(500); res.end('Run npm run build first.');
  }
});
server.on('error', error => { console.error(error.message); process.exitCode = 1; });
server.listen(port, host, () => console.log(`Local preview: http://${host}:${port} (Ctrl+C to stop)`));
