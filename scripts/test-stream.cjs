/**
 * Windows Desktop Automation - Vision Stream Test
 * 
 * Tests WebSocket-based real-time screenshot streaming.
 * 
 * Usage:
 *   node scripts/test-stream.cjs [options]
 * 
 * Options:
 *   --fps <number>      Frames per second (default: 5)
 *   --quality <number>  JPEG quality 1-100 (default: 70)
 *   --duration <number> Test duration in seconds (default: 5)
 *   --save              Save received frames to disk
 *   --help              Show this help
 */

const path = require('path');
const fs = require('fs');
const http = require('http');

// ==================== CONFIGURATION ====================

const DEFAULTS = {
  wsUrl: 'ws://localhost:5001/vision/stream',
  fps: 5,
  quality: 70,
  duration: 5,
  saveFrames: false,
  outputDir: path.join(__dirname, '../test-output/stream'),
};

// ==================== LOGGING ====================

const c = { reset: '\x1b[0m', bold: '\x1b[1m', red: '\x1b[31m', green: '\x1b[32m', yellow: '\x1b[33m', blue: '\x1b[34m', cyan: '\x1b[36m', gray: '\x1b[90m' };

const log = {
  info: (msg) => console.log(`${c.blue}[INFO]${c.reset} ${msg}`),
  ok: (msg) => console.log(`${c.green}[OK]${c.reset} ${msg}`),
  fail: (msg) => console.log(`${c.red}[FAIL]${c.reset} ${msg}`),
  warn: (msg) => console.log(`${c.yellow}[WARN]${c.reset} ${msg}`),
  header: (msg) => console.log(`\n${c.cyan}${c.bold}=== ${msg} ===${c.reset}\n`),
  detail: (msg) => console.log(`${c.gray}       ${msg}${c.reset}`),
};

// ==================== HELPERS ====================

function calcStats(frameTimes) {
  if (!frameTimes.length) return { avg: 0, min: 0, max: 0 };
  const sum = frameTimes.reduce((a, b) => a + b, 0);
  return {
    avg: sum / frameTimes.length,
    min: Math.min(...frameTimes),
    max: Math.max(...frameTimes),
  };
}

function getWebSocket() {
  if (typeof globalThis.WebSocket !== 'undefined') return globalThis.WebSocket;
  try { return require('ws'); } catch { return null; }
}

function checkBridge() {
  return new Promise((resolve) => {
    const req = http.get('http://localhost:5001/health', (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(['ok', 'success'].includes(JSON.parse(data).status)); }
        catch { resolve(false); }
      });
    });
    req.on('error', () => resolve(false));
    req.setTimeout(3000, () => { req.destroy(); resolve(false); });
  });
}

function parseArgs(args) {
  const config = { ...DEFAULTS };
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--fps') config.fps = parseInt(args[++i]) || DEFAULTS.fps;
    else if (arg === '--quality') config.quality = parseInt(args[++i]) || DEFAULTS.quality;
    else if (arg === '--duration') config.duration = parseInt(args[++i]) || DEFAULTS.duration;
    else if (arg === '--save') config.saveFrames = true;
    else if (arg === '--help') {
      console.log(`
Vision Stream Test - Tests WebSocket screenshot streaming

Usage: node test-stream.cjs [options]

Options:
  --fps <n>       Frames per second (default: ${DEFAULTS.fps})
  --quality <n>   JPEG quality 1-100 (default: ${DEFAULTS.quality})
  --duration <n>  Duration in seconds (default: ${DEFAULTS.duration})
  --save          Save frames to disk
  --help          Show this help
`);
      process.exit(0);
    }
  }
  return config;
}

// ==================== TEST ====================

async function testVisionStream(config) {
  log.header('VISION STREAM TEST');
  log.info(`URL: ${config.wsUrl} | FPS: ${config.fps} | Quality: ${config.quality} | Duration: ${config.duration}s`);

  const WebSocket = getWebSocket();
  if (!WebSocket) {
    log.fail('WebSocket not available. Install: npm install ws');
    return { success: false };
  }

  if (config.saveFrames) {
    fs.mkdirSync(config.outputDir, { recursive: true });
    log.info(`Saving frames to: ${config.outputDir}`);
  }

  return new Promise((resolve) => {
    let frameCount = 0, totalBytes = 0, startTime = null, lastFrameTime = null;
    const frameTimes = [];
    const url = `${config.wsUrl}?fps=${config.fps}&quality=${config.quality}`;

    log.info('Connecting...');
    const ws = new WebSocket(url);

    ws.on('open', () => {
      log.ok('Connected');
      startTime = lastFrameTime = Date.now();
      setTimeout(() => { log.info('Closing...'); ws.close(); }, config.duration * 1000);
    });

    ws.on('message', (data) => {
      const now = Date.now();
      frameCount++;
      const frameData = data.toString();
      totalBytes += frameData.length;

      if (lastFrameTime) frameTimes.push(now - lastFrameTime);
      lastFrameTime = now;

      // Progress every 5 frames
      if (frameCount % 5 === 1) {
        const elapsed = (now - startTime) / 1000;
        log.detail(`Frame ${frameCount}: ${(frameData.length / 1024).toFixed(1)} KB | FPS: ${(frameCount / elapsed).toFixed(1)}`);
      }

      // Save if requested
      if (config.saveFrames) {
        const filename = path.join(config.outputDir, `frame_${String(frameCount).padStart(4, '0')}.jpg`);
        try { fs.writeFileSync(filename, Buffer.from(frameData, 'base64')); }
        catch (e) { log.warn(`Save failed: ${e.message}`); }
      }
    });

    ws.on('error', (err) => {
      log.fail(`Error: ${err.message}`);
      resolve({ success: false, error: err.message });
    });

    ws.on('close', () => {
      const totalTime = (Date.now() - (startTime || Date.now())) / 1000;
      const avgFps = frameCount / totalTime;
      const avgFrameSize = frameCount ? totalBytes / frameCount : 0;
      const timing = calcStats(frameTimes);

      log.header('RESULTS');
      console.log(`  Frames: ${frameCount} | Time: ${totalTime.toFixed(2)}s | FPS: ${avgFps.toFixed(2)} (target: ${config.fps})`);
      console.log(`  Data: ${(totalBytes / 1024 / 1024).toFixed(2)} MB | Avg frame: ${(avgFrameSize / 1024).toFixed(1)} KB`);
      console.log(`  Timing: avg=${timing.avg.toFixed(1)}ms, min=${timing.min}ms, max=${timing.max}ms`);

      const success = frameCount > 0 && avgFps > config.fps * 0.5;
      (success ? log.ok : log.fail)(`Test ${success ? 'PASSED' : 'FAILED'}`);
      resolve({ success, frameCount, totalTime, avgFps, avgFrameSize, timing });
    });
  });
}

// ==================== MAIN ====================

async function main() {
  const config = parseArgs(process.argv.slice(2));

  log.info('Checking bridge...');
  if (!await checkBridge()) {
    log.fail('Bridge not responding at localhost:5001');
    log.info('Start with: scripts/start-all.ps1');
    process.exit(1);
  }
  log.ok('Bridge running');

  const result = await testVisionStream(config);
  process.exit(result.success ? 0 : 1);
}

main().catch(err => { log.fail(`Fatal: ${err.message}`); process.exit(1); });
