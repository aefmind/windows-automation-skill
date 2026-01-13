// windows-desktop-automation helper
// Usage:
//   node scripts/send-command.cjs '{"action":"explore"}'
//   echo '{"action":"explore"}' | node scripts/send-command.cjs
//
// This script communicates with MainAgentService HTTP server and sends one JSON command.
// If server is not running, it starts it automatically.

const http = require('http');
const { spawn } = require('child_process');
const path = require('path');

// Resolve agent path relative to this script (using Release build)
const AGENT_PATH = path.resolve(__dirname, '..', 'src', 'src', 'MainAgentService', 'bin', 'Release', 'net9.0-windows', 'MainAgentService.exe');

function readStdin() {
  return new Promise((resolve) => {
    let data = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', (chunk) => (data += chunk));
    process.stdin.on('end', () => resolve(data.trim()));
    if (process.stdin.isTTY) resolve(null);
  });
}

// Function to check if service is running
function isServiceRunning() {
  return new Promise((resolve) => {
    const req = http.get('http://localhost:5000/health', (res) => {
      resolve(res.statusCode === 200);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(1000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

// Function to start service if not running
function startService() {
  return new Promise((resolve, reject) => {
    console.error('Starting MainAgentService...');
    const serviceProcess = spawn(AGENT_PATH, ['--http-server'], {
      detached: true,
      stdio: 'ignore',
      windowsHide: true
    });
    serviceProcess.unref();

    // Wait for service to start
    let attempts = 0;
    const checkInterval = setInterval(async () => {
      if (await isServiceRunning()) {
        clearInterval(checkInterval);
        resolve();
      } else if (attempts++ > 10) {
        clearInterval(checkInterval);
        reject(new Error('Failed to start service'));
      }
    }, 500);
  });
}

(async () => {
  const arg = process.argv.slice(2).join(' ').trim();
  const stdinData = await readStdin();
  const payload = arg || stdinData;

  if (!payload) {
    console.error('No JSON payload provided');
    process.exit(1);
  }

  try {
    // Ensure service is running
    if (!(await isServiceRunning())) {
      await startService();
    }

    // Send command via HTTP
    console.error('Sending command...');
    const postData = payload;
    const options = {
      hostname: 'localhost',
      port: 5000,
      path: '/command',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const response = await new Promise((resolve, reject) => {
      const req = http.request(options, (res) => {
        let body = '';
        res.on('data', (chunk) => body += chunk);
        res.on('end', () => resolve(body));
      });
      req.on('error', reject);
      req.write(postData);
      req.end();
    });

    process.stdout.write(response);
    process.exit(0);
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
})();
