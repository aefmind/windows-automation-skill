/**
 * Windows Desktop Automation - Vision Stream Test
 * 
 * Tests the WebSocket-based real-time screenshot streaming.
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
 * 
 * Prerequisites:
 *   - Python bridge must be running on port 5001
 *   - npm install ws (or use Node 22+ with built-in WebSocket)
 */

const path = require('path');
const fs = require('fs');

// ==================== CONFIGURATION ====================

const DEFAULT_CONFIG = {
  wsUrl: 'ws://localhost:5001/vision/stream',
  fps: 5,
  quality: 70,
  duration: 5,
  saveFrames: false,
  outputDir: path.join(__dirname, '../test-output/stream'),
};

// ==================== COLORS ====================

const colors = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
  gray: '\x1b[90m',
};

const log = {
  info: (msg) => console.log(`${colors.blue}[INFO]${colors.reset} ${msg}`),
  success: (msg) => console.log(`${colors.green}[OK]${colors.reset} ${msg}`),
  fail: (msg) => console.log(`${colors.red}[FAIL]${colors.reset} ${msg}`),
  warn: (msg) => console.log(`${colors.yellow}[WARN]${colors.reset} ${msg}`),
  header: (msg) => console.log(`\n${colors.cyan}${colors.bright}=== ${msg} ===${colors.reset}\n`),
  detail: (msg) => console.log(`${colors.gray}       ${msg}${colors.reset}`),
};

// ==================== WEBSOCKET TEST ====================

async function testVisionStream(config) {
  log.header('VISION STREAM TEST');
  log.info(`URL: ${config.wsUrl}`);
  log.info(`FPS: ${config.fps}`);
  log.info(`Quality: ${config.quality}`);
  log.info(`Duration: ${config.duration}s`);
  log.info(`Save frames: ${config.saveFrames}`);

  // Build URL with query params
  const url = `${config.wsUrl}?fps=${config.fps}&quality=${config.quality}`;

  // Try to use built-in WebSocket (Node 22+) or require 'ws' package
  let WebSocket;
  try {
    // Node 22+ has built-in WebSocket
    if (typeof globalThis.WebSocket !== 'undefined') {
      WebSocket = globalThis.WebSocket;
    } else {
      WebSocket = require('ws');
    }
  } catch (e) {
    log.fail('WebSocket not available. Install ws package: npm install ws');
    process.exit(1);
  }

  // Create output directory if saving frames
  if (config.saveFrames) {
    if (!fs.existsSync(config.outputDir)) {
      fs.mkdirSync(config.outputDir, { recursive: true });
    }
    log.info(`Output directory: ${config.outputDir}`);
  }

  return new Promise((resolve) => {
    let frameCount = 0;
    let totalBytes = 0;
    let startTime = null;
    let lastFrameTime = null;
    const frameTimes = [];

    log.info('Connecting to WebSocket...');

    const ws = new WebSocket(url);

    ws.on('open', () => {
      log.success('Connected to vision stream');
      startTime = Date.now();
      lastFrameTime = startTime;

      // Set timeout to close connection after duration
      setTimeout(() => {
        log.info('Duration complete, closing connection...');
        ws.close();
      }, config.duration * 1000);
    });

    ws.on('message', (data) => {
      const now = Date.now();
      frameCount++;

      // Data is base64 JPEG string
      const frameData = data.toString();
      const frameSize = frameData.length;
      totalBytes += frameSize;

      // Calculate frame timing
      if (lastFrameTime) {
        const elapsed = now - lastFrameTime;
        frameTimes.push(elapsed);
      }
      lastFrameTime = now;

      // Progress indicator
      if (frameCount % 5 === 0 || frameCount === 1) {
        const elapsed = (now - startTime) / 1000;
        const currentFps = frameCount / elapsed;
        log.detail(`Frame ${frameCount}: ${(frameSize / 1024).toFixed(1)} KB | Avg FPS: ${currentFps.toFixed(1)}`);
      }

      // Save frame if requested
      if (config.saveFrames) {
        const filename = path.join(config.outputDir, `frame_${String(frameCount).padStart(4, '0')}.jpg`);
        try {
          const buffer = Buffer.from(frameData, 'base64');
          fs.writeFileSync(filename, buffer);
        } catch (e) {
          log.warn(`Failed to save frame: ${e.message}`);
        }
      }
    });

    ws.on('error', (error) => {
      log.fail(`WebSocket error: ${error.message}`);
      resolve({
        success: false,
        error: error.message,
      });
    });

    ws.on('close', () => {
      const totalTime = (Date.now() - (startTime || Date.now())) / 1000;
      
      // Calculate statistics
      const avgFps = frameCount / totalTime;
      const avgFrameSize = frameCount > 0 ? totalBytes / frameCount : 0;
      const avgFrameTime = frameTimes.length > 0 
        ? frameTimes.reduce((a, b) => a + b, 0) / frameTimes.length 
        : 0;
      const minFrameTime = frameTimes.length > 0 ? Math.min(...frameTimes) : 0;
      const maxFrameTime = frameTimes.length > 0 ? Math.max(...frameTimes) : 0;

      log.header('RESULTS');
      console.log(`  Total frames:     ${frameCount}`);
      console.log(`  Total time:       ${totalTime.toFixed(2)}s`);
      console.log(`  Average FPS:      ${avgFps.toFixed(2)}`);
      console.log(`  Target FPS:       ${config.fps}`);
      console.log(`  FPS accuracy:     ${((avgFps / config.fps) * 100).toFixed(1)}%`);
      console.log(`  Total data:       ${(totalBytes / 1024 / 1024).toFixed(2)} MB`);
      console.log(`  Avg frame size:   ${(avgFrameSize / 1024).toFixed(1)} KB`);
      console.log(`  Frame timing:`);
      console.log(`    - Average:      ${avgFrameTime.toFixed(1)}ms`);
      console.log(`    - Min:          ${minFrameTime}ms`);
      console.log(`    - Max:          ${maxFrameTime}ms`);

      const success = frameCount > 0 && avgFps > (config.fps * 0.5);
      if (success) {
        log.success('Vision stream test PASSED');
      } else {
        log.fail('Vision stream test FAILED');
      }

      resolve({
        success,
        frameCount,
        totalTime,
        avgFps,
        avgFrameSize,
        frameTimes: {
          avg: avgFrameTime,
          min: minFrameTime,
          max: maxFrameTime,
        },
      });
    });
  });
}

// ==================== CHECK BRIDGE ====================

async function checkBridge() {
  const http = require('http');
  
  return new Promise((resolve) => {
    const req = http.get('http://localhost:5001/health', (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          resolve(json.status === 'ok' || json.status === 'success');
        } catch {
          resolve(false);
        }
      });
    });
    
    req.on('error', () => resolve(false));
    req.setTimeout(3000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

// ==================== MAIN ====================

async function main() {
  const args = process.argv.slice(2);
  const config = { ...DEFAULT_CONFIG };

  // Parse arguments
  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--fps':
        config.fps = parseInt(args[++i]) || DEFAULT_CONFIG.fps;
        break;
      case '--quality':
        config.quality = parseInt(args[++i]) || DEFAULT_CONFIG.quality;
        break;
      case '--duration':
        config.duration = parseInt(args[++i]) || DEFAULT_CONFIG.duration;
        break;
      case '--save':
        config.saveFrames = true;
        break;
      case '--help':
        console.log(`
Vision Stream Test - Tests WebSocket screenshot streaming

Usage:
  node test-stream.cjs [options]

Options:
  --fps <number>      Frames per second (default: ${DEFAULT_CONFIG.fps})
  --quality <number>  JPEG quality 1-100 (default: ${DEFAULT_CONFIG.quality})
  --duration <number> Test duration in seconds (default: ${DEFAULT_CONFIG.duration})
  --save              Save received frames to disk
  --help              Show this help

Prerequisites:
  - Python bridge must be running on port 5001
  - Install ws package: npm install ws
`);
        process.exit(0);
    }
  }

  // Check if bridge is running
  log.info('Checking Python bridge...');
  const bridgeOk = await checkBridge();
  if (!bridgeOk) {
    log.fail('Python bridge not responding at localhost:5001');
    log.info('Start the bridge with: scripts/start-all.ps1');
    process.exit(1);
  }
  log.success('Python bridge is running');

  // Run test
  const result = await testVisionStream(config);
  process.exit(result.success ? 0 : 1);
}

main().catch(err => {
  log.fail(`Fatal error: ${err.message}`);
  process.exit(1);
});
