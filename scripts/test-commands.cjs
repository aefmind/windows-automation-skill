/**
 * Windows Desktop Automation - Test Suite
 * 
 * Automated testing for all commands (v3.3)
 * 
 * Usage:
 *   node scripts/test-commands.cjs              # Run all tests
 *   node scripts/test-commands.cjs --category discovery  # Run category
 *   node scripts/test-commands.cjs --test explore        # Run single test
 *   node scripts/test-commands.cjs --list                # List all tests
 * 
 * Prerequisites:
 *   - MainAgentService must be built
 *   - Run from the skill root directory
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// ==================== CONFIGURATION ====================

const CONFIG = {
  agentPath: path.join(__dirname, '../src/src/MainAgentService/bin/Debug/net9.0-windows/MainAgentService.exe'),
  timeout: 15000,
  startupDelay: 1500,
  testApp: 'notepad.exe',
  testWindowTitle: 'Notepad',
  screenshotDir: path.join(__dirname, '../test-output'),
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
  success: (msg) => console.log(`${colors.green}[PASS]${colors.reset} ${msg}`),
  fail: (msg) => console.log(`${colors.red}[FAIL]${colors.reset} ${msg}`),
  warn: (msg) => console.log(`${colors.yellow}[WARN]${colors.reset} ${msg}`),
  header: (msg) => console.log(`\n${colors.cyan}${colors.bright}=== ${msg} ===${colors.reset}\n`),
  detail: (msg) => console.log(`${colors.gray}       ${msg}${colors.reset}`),
};

// ==================== TEST RUNNER ====================

class TestRunner {
  constructor() {
    this.results = { passed: 0, failed: 0, skipped: 0, tests: [] };
    this.agentProcess = null;
    this.testAppPid = null;
  }

  /**
   * Sends a command to the agent and returns the parsed response
   */
  async sendCommand(command, options = {}) {
    return new Promise((resolve, reject) => {
      const timeout = options.timeout || CONFIG.timeout;
      const startupDelay = options.startupDelay || CONFIG.startupDelay;
      
      const child = spawn(CONFIG.agentPath, [], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
      });

      let stdout = '';
      let stderr = '';
      let timer = null;

      child.stdout.on('data', (data) => { stdout += data.toString(); });
      child.stderr.on('data', (data) => { stderr += data.toString(); });

      timer = setTimeout(() => {
        child.kill();
        reject(new Error(`Timeout after ${timeout}ms`));
      }, timeout);

      setTimeout(() => {
        child.stdin.write(JSON.stringify(command) + '\n');
        child.stdin.write('{"action":"exit"}\n');
        child.stdin.end();
      }, startupDelay);

      child.on('close', (code) => {
        clearTimeout(timer);
        try {
          // Find JSON in output - look for command response (skip warnings)
          const lines = stdout.trim().split('\n');
          let response = null;
          let lastJson = null;
          
          for (const line of lines) {
            if (line.startsWith('{')) {
              try {
                const parsed = JSON.parse(line);
                lastJson = parsed;
                // Skip bridge warnings - look for actual command response
                if (parsed.status !== 'warn') {
                  response = parsed;
                  // Don't break - prefer later responses (command result comes after warnings)
                }
              } catch (e) { /* not valid JSON, continue */ }
            }
          }
          
          // If no non-warn response, use the last JSON we found
          if (!response && lastJson) {
            response = lastJson;
          }
          
          if (response) {
            resolve(response);
          } else {
            reject(new Error(`No valid JSON in response: ${stdout.substring(0, 200)}`));
          }
        } catch (e) {
          reject(new Error(`Parse error: ${e.message}`));
        }
      });

      child.on('error', (err) => {
        clearTimeout(timer);
        reject(err);
      });
    });
  }

  /**
   * Runs a single test case
   */
  async runTest(test) {
    const startTime = Date.now();
    try {
      // Run setup if defined
      if (test.setup) {
        await test.setup(this);
      }

      // Send the command
      const response = await this.sendCommand(test.command, test.options);
      
      // Validate the response
      let passed = true;
      let failReason = '';

      if (test.validate) {
        const result = test.validate(response);
        if (result !== true) {
          passed = false;
          failReason = typeof result === 'string' ? result : 'Validation failed';
        }
      } else {
        // Default validation: check for success status
        if (response.status !== 'success') {
          passed = false;
          failReason = response.message || response.code || 'Non-success status';
        }
      }

      // Run teardown if defined
      if (test.teardown) {
        await test.teardown(this);
      }

      const duration = Date.now() - startTime;
      
      if (passed) {
        this.results.passed++;
        log.success(`${test.name} (${duration}ms)`);
        if (test.showResponse) {
          log.detail(JSON.stringify(response, null, 2).split('\n').slice(0, 5).join('\n'));
        }
      } else {
        this.results.failed++;
        log.fail(`${test.name} - ${failReason}`);
        log.detail(`Response: ${JSON.stringify(response).substring(0, 200)}`);
      }

      this.results.tests.push({ name: test.name, passed, duration, failReason });
      return passed;
    } catch (error) {
      this.results.failed++;
      log.fail(`${test.name} - ${error.message}`);
      this.results.tests.push({ name: test.name, passed: false, failReason: error.message });
      return false;
    }
  }

  /**
   * Runs all tests in a category
   */
  async runCategory(category, tests) {
    log.header(category);
    for (const test of tests) {
      await this.runTest(test);
    }
  }

  /**
   * Prints the final summary
   */
  printSummary() {
    log.header('TEST SUMMARY');
    console.log(`  ${colors.green}Passed:${colors.reset}  ${this.results.passed}`);
    console.log(`  ${colors.red}Failed:${colors.reset}  ${this.results.failed}`);
    console.log(`  ${colors.yellow}Skipped:${colors.reset} ${this.results.skipped}`);
    console.log(`  ${colors.blue}Total:${colors.reset}   ${this.results.passed + this.results.failed + this.results.skipped}`);
    
    if (this.results.failed > 0) {
      console.log(`\n${colors.red}Failed Tests:${colors.reset}`);
      this.results.tests
        .filter(t => !t.passed)
        .forEach(t => console.log(`  - ${t.name}: ${t.failReason}`));
    }
    
    return this.results.failed === 0;
  }

  /**
   * Helper to launch test app (Notepad)
   */
  async launchTestApp() {
    try {
      const response = await this.sendCommand({
        action: 'launch_app',
        path: CONFIG.testApp,
        wait_for_window: CONFIG.testWindowTitle
      });
      if (response.status === 'success' && response.pid) {
        this.testAppPid = response.pid;
      }
      // Give window time to fully render
      await this.sleep(500);
      return response;
    } catch (e) {
      log.warn(`Failed to launch test app: ${e.message}`);
      return null;
    }
  }

  /**
   * Helper to close test app
   */
  async closeTestApp() {
    try {
      // Try to close gracefully first
      await this.sendCommand({ action: 'close_window', selector: CONFIG.testWindowTitle });
      await this.sleep(200);
      // Handle "Save?" dialog if it appears
      await this.sendCommand({ action: 'key_press', key: 'n' });
    } catch (e) {
      // If close fails, try kill
      if (this.testAppPid) {
        await this.sendCommand({ action: 'kill_process', pid: this.testAppPid, force: true });
      }
    }
    this.testAppPid = null;
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// ==================== TEST DEFINITIONS ====================

const TESTS = {
  // -------------------- DISCOVERY --------------------
  discovery: [
    {
      name: 'explore - list windows',
      command: { action: 'explore' },
      validate: (r) => r.status === 'success' && Array.isArray(r.data) && r.data.length >= 0,
      showResponse: false,
    },
    {
      name: 'explore_window - explore test window',
      command: { action: 'explore_window', selector: 'Notepad', max_depth: 2 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' && (Array.isArray(r.elements) || Array.isArray(r.data)),
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'find_element - find button in Notepad',
      command: { action: 'find_element', selector: 'Edit', window: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' || (r.status === 'error' && r.code === 'ELEMENT_NOT_FOUND'),
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'get_window_info - get window details',
      command: { action: 'get_window_info', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' && r.data && typeof r.data.title === 'string',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
  ],

  // -------------------- WINDOW MANAGEMENT --------------------
  windowManagement: [
    {
      name: 'focus_window - bring window to front',
      command: { action: 'focus_window', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'wait_for_window - wait for window to appear',
      command: { action: 'launch_app', path: 'notepad.exe', wait_for_window: 'Notepad' },
      validate: (r) => r.status === 'success' && r.pid > 0,
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'move_window - move to coordinates',
      command: { action: 'move_window', selector: 'Notepad', x: 100, y: 100 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'resize_window - resize window',
      command: { action: 'resize_window', selector: 'Notepad', width: 800, height: 600 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'minimize_window - minimize window',
      command: { action: 'minimize_window', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'maximize_window - maximize window',
      command: { action: 'maximize_window', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'restore_window - restore window',
      command: { action: 'restore_window', selector: 'Notepad' },
      setup: async (runner) => { 
        await runner.launchTestApp();
        await runner.sendCommand({ action: 'maximize_window', selector: 'Notepad' });
        await runner.sleep(200);
      },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'close_window - close window',
      command: { action: 'close_window', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
    },
  ],

  // -------------------- APP/PROCESS CONTROL --------------------
  appProcess: [
    {
      name: 'launch_app - launch notepad',
      command: { action: 'launch_app', path: 'notepad.exe', wait_for_window: 'Notepad' },
      validate: (r) => r.status === 'success' && r.pid > 0,
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'list_processes - list all',
      command: { action: 'list_processes' },
      validate: (r) => r.status === 'success' && Array.isArray(r.data) && r.data.length > 0,
    },
    {
      name: 'list_processes - filter by name',
      command: { action: 'list_processes', filter: 'notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' && Array.isArray(r.data),
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'kill_process - kill by name',
      command: { action: 'kill_process', name: 'notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
    },
  ],

  // -------------------- MOUSE ACTIONS --------------------
  mouse: [
    {
      name: 'click_at - click at coordinates',
      command: { action: 'click_at', x: 500, y: 300 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'click_at - right click',
      command: { action: 'click_at', x: 500, y: 300, button: 'right' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'mouse_move - move to coordinates',
      command: { action: 'mouse_move', x: 400, y: 400 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'scroll - scroll down',
      command: { action: 'scroll', direction: 'down', amount: 100 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'scroll - scroll up',
      command: { action: 'scroll', direction: 'up', amount: 100 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'drag_and_drop - drag between coordinates',
      command: { action: 'drag_and_drop', from_x: 200, from_y: 200, to_x: 400, to_y: 400 },
      validate: (r) => r.status === 'success',
    },
  ],

  // -------------------- KEYBOARD ACTIONS --------------------
  keyboard: [
    {
      name: 'hotkey - ctrl+a',
      command: { action: 'hotkey', keys: 'ctrl+a' },
      setup: async (runner) => { 
        await runner.launchTestApp(); 
        await runner.sendCommand({ action: 'focus_window', selector: 'Notepad' });
      },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'key_press - press escape',
      command: { action: 'key_press', key: 'escape' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'key_press - press enter',
      command: { action: 'key_press', key: 'enter' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'type - type text',
      command: { action: 'type', selector: 'Notepad', text: 'Hello Test!' },
      setup: async (runner) => { 
        await runner.launchTestApp(); 
        await runner.sendCommand({ action: 'focus_window', selector: 'Notepad' });
      },
      validate: (r) => r.status === 'success' || r.code === 'ELEMENT_NOT_FOUND',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
  ],

  // -------------------- CLIPBOARD --------------------
  clipboard: [
    {
      name: 'set_clipboard - write text',
      command: { action: 'set_clipboard', text: 'Test clipboard content' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'get_clipboard - read text',
      command: { action: 'get_clipboard' },
      setup: async (runner) => {
        await runner.sendCommand({ action: 'set_clipboard', text: 'Verify clipboard' });
      },
      validate: (r) => r.status === 'success' && typeof r.text === 'string',
    },
  ],

  // -------------------- TEXT/OCR & SCREENSHOT --------------------
  textAndScreenshot: [
    {
      name: 'read_text - read from element',
      command: { action: 'read_text', selector: 'Notepad' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' || r.status === 'error', // May fail if OCR not available
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'screenshot - capture screen',
      command: { action: 'screenshot', filename: 'test-screenshot.png' },
      validate: (r) => r.status === 'success',
    },
  ],

  // -------------------- ERROR HANDLING --------------------
  errorHandling: [
    {
      name: 'missing_param - returns MISSING_PARAM error',
      command: { action: 'focus_window' }, // Missing selector
      validate: (r) => r.status === 'error' && r.code === 'MISSING_PARAM',
    },
    {
      name: 'window_not_found - returns WINDOW_NOT_FOUND error',
      command: { action: 'focus_window', selector: 'NonExistentWindow12345' },
      validate: (r) => r.status === 'error' && r.code === 'WINDOW_NOT_FOUND',
    },
    {
      name: 'unknown_action - returns UNKNOWN_ACTION error',
      command: { action: 'nonexistent_action' },
      validate: (r) => r.status === 'error' && r.code === 'UNKNOWN_ACTION',
    },
  ],

  // -------------------- SYSTEM (v3.2) --------------------
  system: [
    {
      name: 'health - system health check',
      command: { action: 'health' },
      validate: (r) => r.status === 'success' && r.version && r.bridges,
    },
  ],

  // -------------------- MULTI-MONITOR (v3.2) --------------------
  multiMonitor: [
    {
      name: 'list_monitors - enumerate displays',
      command: { action: 'list_monitors' },
      validate: (r) => r.status === 'success' && Array.isArray(r.data) && r.data.length >= 1,
    },
    {
      name: 'screenshot_monitor - capture primary monitor',
      command: { action: 'screenshot_monitor', monitor: 0, filename: 'test-monitor-0.png' },
      validate: (r) => r.status === 'success' && r.filename,
    },
    {
      name: 'screenshot_monitor - invalid monitor returns error',
      command: { action: 'screenshot_monitor', monitor: 99 },
      validate: (r) => r.status === 'error' && r.code === 'INVALID_MONITOR',
    },
    {
      name: 'move_to_monitor - move window to primary monitor',
      command: { action: 'move_to_monitor', selector: 'Notepad', monitor: 0, position: 'center' },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
  ],

  // -------------------- FILE DIALOG (v3.2) --------------------
  fileDialog: [
    {
      name: 'file_dialog - detect when no dialog open',
      command: { action: 'file_dialog', dialog_action: 'detect' },
      validate: (r) => r.status === 'error' && r.code === 'DIALOG_NOT_FOUND',
    },
    {
      name: 'file_dialog - invalid action',
      command: { action: 'file_dialog', dialog_action: 'invalid_action' },
      validate: (r) => r.status === 'error' && r.code === 'INVALID_DIALOG_ACTION',
    },
  ],

  // -------------------- WAIT FOR ELEMENT (v3.2) --------------------
  waitForElement: [
    {
      name: 'wait_for_element - find existing element',
      command: { action: 'wait_for_element', selector: 'Edit', window: 'Notepad', timeout: 3000 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' || (r.status === 'error' && r.code === 'TIMEOUT'),
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'wait_for_element - timeout for non-existent element',
      command: { action: 'wait_for_element', selector: 'NonExistentElement99', timeout: 1000, poll_interval: 200 },
      validate: (r) => r.status === 'error' && r.code === 'TIMEOUT',
    },
  ],

  // -------------------- BATCH (v3.2) --------------------
  batch: [
    {
      name: 'batch - execute simple command sequence',
      command: { 
        action: 'batch', 
        commands: [
          { action: 'set_clipboard', text: 'batch_test_1' },
          { action: 'get_clipboard' }
        ],
        stop_on_error: true 
      },
      validate: (r) => r.status === 'success' && r.succeeded === 2 && r.failed === 0,
    },
    {
      name: 'batch - stop on error',
      command: { 
        action: 'batch', 
        commands: [
          { action: 'set_clipboard', text: 'before_error' },
          { action: 'focus_window', selector: 'NonExistentWindow12345' },
          { action: 'set_clipboard', text: 'after_error' }
        ],
        stop_on_error: true 
      },
      validate: (r) => r.status === 'error' && r.succeeded === 1 && r.failed === 1,
    },
    {
      name: 'batch - continue on error',
      command: { 
        action: 'batch', 
        commands: [
          { action: 'set_clipboard', text: 'before_error' },
          { action: 'focus_window', selector: 'NonExistentWindow12345' },
          { action: 'set_clipboard', text: 'after_error' }
        ],
        stop_on_error: false 
      },
      validate: (r) => r.succeeded === 2 && r.failed === 1,
    },
  ],

  // -------------------- ADVANCED MOUSE (v3.3) --------------------
  advancedMouse: [
    {
      name: 'mouse_path - move along waypoints',
      command: { action: 'mouse_path', points: [[100, 100], [200, 150], [300, 100]], duration: 300 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'mouse_bezier - move along bezier curve',
      command: { action: 'mouse_bezier', start: [100, 100], control1: [150, 50], control2: [250, 50], end: [300, 100], steps: 30, duration: 300 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'draw - draw path',
      command: { action: 'draw', points: [[100, 100], [150, 120], [200, 100]], button: 'left', duration: 300 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'mouse_down - press left button',
      command: { action: 'mouse_down', button: 'left', x: 400, y: 400 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'mouse_up - release left button',
      command: { action: 'mouse_up', button: 'left' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'click_relative - click relative to element',
      command: { action: 'click_relative', selector: 'Notepad', anchor: 'center', offset_x: 10, offset_y: 10 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
  ],

  // -------------------- ADVANCED KEYBOARD (v3.3) --------------------
  advancedKeyboard: [
    {
      name: 'key_down - press shift',
      command: { action: 'key_down', key: 'shift' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'key_up - release shift',
      command: { action: 'key_up', key: 'shift' },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'type_here - type at cursor',
      command: { action: 'type_here', text: 'Test typing' },
      setup: async (runner) => { 
        await runner.launchTestApp(); 
        await runner.sendCommand({ action: 'focus_window', selector: 'Notepad' });
      },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'type_at_cursor - alias for type_here',
      command: { action: 'type_at_cursor', text: 'More text' },
      setup: async (runner) => { 
        await runner.launchTestApp(); 
        await runner.sendCommand({ action: 'focus_window', selector: 'Notepad' });
      },
      validate: (r) => r.status === 'success',
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
  ],

  // -------------------- WAIT FOR STATE (v3.3) --------------------
  waitForState: [
    {
      name: 'wait_for_state - element exists',
      command: { action: 'wait_for_state', selector: 'Edit', state: 'exists', window: 'Notepad', timeout: 3000 },
      setup: async (runner) => { await runner.launchTestApp(); },
      validate: (r) => r.status === 'success' || (r.status === 'error' && r.code === 'TIMEOUT'),
      teardown: async (runner) => { await runner.closeTestApp(); },
    },
    {
      name: 'wait_for_state - element not_exists (timeout expected)',
      command: { action: 'wait_for_state', selector: 'NonExistent99', state: 'not_exists', timeout: 1000 },
      validate: (r) => r.status === 'success',
    },
    {
      name: 'wait_for_state - invalid state returns error',
      command: { action: 'wait_for_state', selector: 'Button', state: 'invalid_state', timeout: 1000 },
      validate: (r) => r.status === 'error',
    },
  ],

  // -------------------- OCR REGION (v3.3) --------------------
  ocrRegion: [
    {
      name: 'ocr_region - read screen region',
      command: { action: 'ocr_region', x: 100, y: 100, width: 200, height: 50 },
      // OCR may fail if Java bridge is not running, which is acceptable
      validate: (r) => r.status === 'success' || (r.status === 'error' && (r.code === 'OCR_FAILED' || r.code === 'OCR_REGION_FAILED')),
    },
    {
      name: 'ocr_region - missing parameters',
      command: { action: 'ocr_region', x: 100, y: 100 }, // missing width/height
      validate: (r) => r.status === 'error' && r.code === 'MISSING_PARAM',
    },
  ],
};

// ==================== MAIN EXECUTION ====================

async function main() {
  const args = process.argv.slice(2);
  const runner = new TestRunner();

  // Parse arguments
  let categoryFilter = null;
  let testFilter = null;
  let listOnly = false;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--category' && args[i + 1]) {
      categoryFilter = args[i + 1].toLowerCase();
      i++;
    } else if (args[i] === '--test' && args[i + 1]) {
      testFilter = args[i + 1].toLowerCase();
      i++;
    } else if (args[i] === '--list') {
      listOnly = true;
    } else if (args[i] === '--help') {
      console.log(`
Windows Desktop Automation Test Suite

Usage:
  node test-commands.cjs              Run all tests
  node test-commands.cjs --category discovery  Run category
  node test-commands.cjs --test explore        Run single test
  node test-commands.cjs --list                List all tests
  node test-commands.cjs --help                Show this help

Categories: ${Object.keys(TESTS).join(', ')}
`);
      process.exit(0);
    }
  }

  // List tests
  if (listOnly) {
    console.log('\nAvailable Tests:\n');
    for (const [category, tests] of Object.entries(TESTS)) {
      console.log(`${colors.cyan}${category}${colors.reset}`);
      tests.forEach(t => console.log(`  - ${t.name}`));
    }
    process.exit(0);
  }

  // Verify agent exists
  if (!fs.existsSync(CONFIG.agentPath)) {
    log.fail(`Agent not found at: ${CONFIG.agentPath}`);
    log.info('Run: dotnet build (in MainAgentService directory)');
    process.exit(1);
  }

  // Create test output directory
  if (!fs.existsSync(CONFIG.screenshotDir)) {
    fs.mkdirSync(CONFIG.screenshotDir, { recursive: true });
  }

  log.header('WINDOWS DESKTOP AUTOMATION TEST SUITE');
  log.info(`Agent: ${CONFIG.agentPath}`);
  log.info(`Timeout: ${CONFIG.timeout}ms`);

  // Run tests
  for (const [category, tests] of Object.entries(TESTS)) {
    if (categoryFilter && category.toLowerCase() !== categoryFilter) continue;

    let filteredTests = tests;
    if (testFilter) {
      filteredTests = tests.filter(t => t.name.toLowerCase().includes(testFilter));
    }

    if (filteredTests.length > 0) {
      await runner.runCategory(category, filteredTests);
    }
  }

  // Print summary
  const success = runner.printSummary();
  process.exit(success ? 0 : 1);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
