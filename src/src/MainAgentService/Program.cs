using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MainAgentService
{
    class Program
    {
        private static UIA3Automation? _automation;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
        private const string PythonBridgeUrl = "http://localhost:5001";

        // === v3.5 Performance & Diagnostics ===
        private static bool _debugMode = false;
        private static readonly Dictionary<string, CommandMetrics> _metrics = new();
        private static readonly Dictionary<string, AutomationElement?> _elementCache = new();
        private static DateTime _cacheTimestamp = DateTime.MinValue;
        private static readonly TimeSpan _cacheExpiry = TimeSpan.FromSeconds(5);
        private static readonly Dictionary<string, EventSubscription> _subscriptions = new();
        private static readonly object _metricsLock = new object();

        private static async Task RunHttpServer()
        {
            var builder = WebApplication.CreateBuilder();
            var app = builder.Build();

            app.MapPost("/command", async (HttpContext context) =>
            {
                try
                {
                    // Read raw body and parse with Newtonsoft.Json for compatibility
                    using var reader = new StreamReader(context.Request.Body, Encoding.UTF8);
                    var body = await reader.ReadToEndAsync();
                    var command = JObject.Parse(body);

                    var action = command["action"]?.ToString();
                    if (action == null) return Results.BadRequest("Missing action");

                    // Capture console output with UTF-8 encoding
                    using var stringWriter = new StringWriter();
                    var originalOut = Console.Out;
                    Console.SetOut(stringWriter);

                    try
                    {
                        await ExecuteSingleCommand(action, command);
                        var output = stringWriter.ToString().Trim();
                        return Results.Text(output, "application/json", Encoding.UTF8);
                    }
                    finally
                    {
                        Console.SetOut(originalOut);
                    }
                }
                catch (JsonReaderException ex)
                {
                    return Results.Json(new { status = "error", code = "INVALID_JSON", message = ex.Message });
                }
                catch (Exception ex)
                {
                    return Results.Json(new { status = "error", message = ex.Message });
                }
            });

            app.MapGet("/health", () => Results.Ok(new { status = "ok", mode = "http", version = "4.3" }));

            Console.WriteLine("Main Agent Service HTTP server running on http://localhost:5000");
            
            // Auto-cleanup old screenshots on startup
            AutoCleanupScreenshots();

            try
            {
                await app.RunAsync("http://localhost:5000");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start HTTP server: {ex.Message}");
            }
        }

        // === v3.6 Human-like Behavior ===
        private static HumanModeSettings _humanMode = new HumanModeSettings();
        private static readonly Random _random = new Random();
        private static DateTime _sessionStart = DateTime.Now;

        // === v4.3 Temporary Screenshots ===
        private static readonly string _tempScreenshotDir = Path.Combine(Path.GetTempPath(), "windows-desktop-automation", "screenshots");
        private static readonly TimeSpan _screenshotAutoCleanupAge = TimeSpan.FromMinutes(30);

        // Keyboard layout for typo simulation
        private static readonly Dictionary<char, char[]> _keyboardAdjacent = new()
        {
            ['a'] = new[] { 'q', 'w', 's', 'z' },
            ['b'] = new[] { 'v', 'g', 'h', 'n' },
            ['c'] = new[] { 'x', 'd', 'f', 'v' },
            ['d'] = new[] { 's', 'e', 'r', 'f', 'c', 'x' },
            ['e'] = new[] { 'w', 's', 'd', 'r' },
            ['f'] = new[] { 'd', 'r', 't', 'g', 'v', 'c' },
            ['g'] = new[] { 'f', 't', 'y', 'h', 'b', 'v' },
            ['h'] = new[] { 'g', 'y', 'u', 'j', 'n', 'b' },
            ['i'] = new[] { 'u', 'j', 'k', 'o' },
            ['j'] = new[] { 'h', 'u', 'i', 'k', 'm', 'n' },
            ['k'] = new[] { 'j', 'i', 'o', 'l', 'm' },
            ['l'] = new[] { 'k', 'o', 'p' },
            ['m'] = new[] { 'n', 'j', 'k' },
            ['n'] = new[] { 'b', 'h', 'j', 'm' },
            ['o'] = new[] { 'i', 'k', 'l', 'p' },
            ['p'] = new[] { 'o', 'l' },
            ['q'] = new[] { 'w', 'a' },
            ['r'] = new[] { 'e', 'd', 'f', 't' },
            ['s'] = new[] { 'a', 'w', 'e', 'd', 'x', 'z' },
            ['t'] = new[] { 'r', 'f', 'g', 'y' },
            ['u'] = new[] { 'y', 'h', 'j', 'i' },
            ['v'] = new[] { 'c', 'f', 'g', 'b' },
            ['w'] = new[] { 'q', 'a', 's', 'e' },
            ['x'] = new[] { 'z', 's', 'd', 'c' },
            ['y'] = new[] { 't', 'g', 'h', 'u' },
            ['z'] = new[] { 'a', 's', 'x' },
        };

        static async Task Main(string[] args)
        {
            // Singleton pattern to prevent multiple instances
            using var mutex = new Mutex(true, "MainAgentServiceMutex", out bool createdNew);
            if (!createdNew)
            {
                Console.WriteLine("Service already running");
                return;
            }

            _automation = new UIA3Automation();

            // Check for HTTP server mode
            if (args.Length > 0 && args[0] == "--http-server")
            {
                Console.WriteLine("Starting Main Agent Service in HTTP mode...");
                await RunHttpServer();
                return;
            }

            // Default stdin mode
            Console.WriteLine("Main Agent Service Started in stdin mode. Waiting for commands...");

            // Auto-cleanup old screenshots on startup
            AutoCleanupScreenshots();

            await CheckBridges();

            while (true)
            {
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input)) continue;

                try
                {
                    var command = JObject.Parse(input);
                    var action = command["action"]?.ToString();

                    if (action == null) continue;

                    // Handle special commands that need main loop context
                    switch (action)
                    {
                        case "exit":
                            return;
                        case "batch":
                            await HandleBatch(command);
                            break;
                        default:
                            // Delegate all other commands to the unified dispatcher
                            await ExecuteSingleCommand(action, command);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(new { status = "error", message = ex.Message }));
                }
            }
        }

        /// <summary>
        /// Helper: Checks if a bridge endpoint is available
        /// </summary>
        private static async Task<bool> IsBridgeOnline(string url, string endpoint = "health")
        {
            try
            {
                var response = await _httpClient.GetAsync($"{url}/{endpoint}");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        private static async Task CheckBridges()
        {
            var pyOnline = await IsBridgeOnline(PythonBridgeUrl, "explore");
            Console.WriteLine(JsonConvert.SerializeObject(new { status = "info", bridge = "Python", active = pyOnline }));
        }

        private static void HandleExplore()
        {
            if (_automation == null) return;
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren();
            var result = windows.Select(w => new
            {
                Title = SafeGet(() => w.Name, "Unknown"),
                Type = w.ControlType.ToString(),
                AutomationId = SafeGet(() => w.AutomationId, ""),
                ClassName = SafeGet(() => w.ClassName, ""),
                ProcessId = SafeGet(() => w.Properties.ProcessId.Value, 0),
                IsEnabled = SafeGet(() => w.IsEnabled, false),
                Rect = w.BoundingRectangle
            }).Where(w => !string.IsNullOrWhiteSpace(w.Title) && w.Title != "Unknown").ToList();

            WriteSuccess(new { status = "success", count = result.Count, data = result });
        }

        private static T SafeGet<T>(Func<T> getter, T defaultValue)
        {
            try { return getter(); } catch { return defaultValue; }
        }

        // ==================== RESPONSE HELPERS ====================

        /// <summary>
        /// Outputs a success response with the given data
        /// </summary>
        private static void WriteSuccess(object data)
        {
            Console.WriteLine(JsonConvert.SerializeObject(data));
        }

        /// <summary>
        /// Outputs a success response for an action
        /// </summary>
        private static void WriteSuccess(string action, object? additionalData = null)
        {
            var response = new Dictionary<string, object?> { ["status"] = "success", ["action"] = action };
            if (additionalData != null)
            {
                foreach (var prop in JObject.FromObject(additionalData).Properties())
                {
                    response[prop.Name] = prop.Value?.ToObject<object>();
                }
            }
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }

        /// <summary>
        /// Outputs an error response with code and message
        /// </summary>
        private static void WriteError(string code, string message)
        {
            Console.WriteLine(JsonConvert.SerializeObject(new { status = "error", code, message }));
        }

        /// <summary>
        /// Outputs a missing parameter error
        /// </summary>
        private static void WriteMissingParam(string paramName)
        {
            WriteError("MISSING_PARAM", $"Missing '{paramName}' parameter");
        }

        /// <summary>
        /// Outputs a window not found error
        /// </summary>
        private static void WriteWindowNotFound(string selector)
        {
            WriteError("WINDOW_NOT_FOUND", $"Window '{selector}' not found");
        }

        /// <summary>
        /// Outputs an element not found error
        /// </summary>
        private static void WriteElementNotFound(string selector)
        {
            WriteError("ELEMENT_NOT_FOUND", $"Element '{selector}' not found");
        }

        // ==================== INPUT HELPERS ====================

        /// <summary>
        /// Parses a button string to MouseButton enum
        /// </summary>
        private static MouseButton ParseMouseButton(string? button)
        {
            return button?.ToLower() switch
            {
                "right" => MouseButton.Right,
                "middle" => MouseButton.Middle,
                _ => MouseButton.Left
            };
        }

        /// <summary>
        /// Finds a button in a dialog by trying AutomationId first, then by name fallbacks
        /// </summary>
        private static AutomationElement? FindDialogButton(Window dialog, string automationId, params string[] nameFallbacks)
        {
            var button = dialog.FindFirstDescendant(cf => cf.ByAutomationId(automationId));
            if (button != null) return button;

            foreach (var name in nameFallbacks)
            {
                button = dialog.FindFirstDescendant(cf => cf.ByName(name));
                if (button != null) return button;
            }

            return null;
        }

        // ==================== STA THREAD HELPER ====================

        /// <summary>
        /// Runs an action on an STA thread (required for clipboard operations)
        /// </summary>
        private static void RunOnSTAThread(Action action)
        {
            var thread = new Thread(() => action());
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        /// <summary>
        /// Runs a function on an STA thread and returns the result
        /// </summary>
        private static T? RunOnSTAThread<T>(Func<T> func)
        {
            T? result = default;
            var thread = new Thread(() => result = func());
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        private static async Task HandleMouseAction(JObject command, string actionType)
        {
            var selector = command["selector"]?.ToString();
            var imageHint = command["image_hint"]?.ToString();

            // 1. Try FlaUI
            try
            {
                var element = FindElement(selector);
                if (element != null)
                {
                    switch (actionType)
                    {
                        case "click": element.Click(); break;
                        case "double_click": element.DoubleClick(); break;
                        case "right_click": element.RightClick(); break;
                        case "long_press": 
                            Mouse.MoveTo(element.GetClickablePoint());
                            Mouse.Down(MouseButton.Left);
                            await Task.Delay(2000);
                            Mouse.Up(MouseButton.Left);
                            break;
                    }
                    WriteSuccess(actionType, new { engine = "FlaUI" });
                    return;
                }
            }
            catch (Exception ex) { Console.Error.WriteLine($"FlaUI Error: {ex.Message}"); }

            // 2. Fallback to Python (pywinauto)
            if (await CallBridge(PythonBridgeUrl, actionType, command)) return;

            WriteError($"{actionType.ToUpper()}_FAILED", $"All engines failed to {actionType}");
        }

        private static async Task HandleType(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var text = command["text"]?.ToString() ?? "";

            var element = FindElement(selector);
            if (element != null)
            {
                element.Focus();
                Keyboard.Type(text);
                WriteSuccess("type", new { engine = "FlaUI" });
                return;
            }

            if (await CallBridge(PythonBridgeUrl, "type", command)) return;

            WriteError("TYPE_FAILED", "Failed to type text - element not found");
        }

        private static async Task HandleScroll(JObject command)
        {
            var direction = command["direction"]?.ToString() ?? "down";
            var amount = command["amount"]?.Value<int>() ?? 100;

            Mouse.Scroll(amount * (direction == "up" ? 1 : -1));
            WriteSuccess("scroll", new { direction, amount });
            await Task.CompletedTask;
        }

        private static void HandleScreenshot(JObject command)
        {
            var fileName = command["filename"]?.ToString();
            var useTemp = command["temp"]?.Value<bool>() ?? true; // Default to temp directory
            var desktop = _automation?.GetDesktop();
            if (desktop == null)
            {
                WriteError("SCREENSHOT_FAILED", "Automation not initialized");
                return;
            }

            try
            {
                string fullPath;
                bool isTemporary;

                if (fileName != null && Path.IsPathRooted(fileName))
                {
                    // Absolute path provided - user wants permanent location
                    fullPath = fileName;
                    isTemporary = false;
                }
                else
                {
                    // Use temp directory (default behavior)
                    Directory.CreateDirectory(_tempScreenshotDir);
                    var name = fileName ?? $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    fullPath = Path.Combine(_tempScreenshotDir, name);
                    isTemporary = useTemp;
                }

                var image = FlaUI.Core.Capturing.Capture.Screen();
                image.ToFile(fullPath);
                WriteSuccess("screenshot", new { path = fullPath, temporary = isTemporary });
            }
            catch (Exception ex)
            {
                WriteError("SCREENSHOT_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Cleans up old temporary screenshots
        /// </summary>
        private static void HandleCleanupScreenshots(JObject command)
        {
            var maxAgeMinutes = command["max_age_minutes"]?.Value<int>() ?? 30;
            var cutoff = DateTime.Now.AddMinutes(-maxAgeMinutes);
            int deleted = 0;
            int failed = 0;
            var errors = new List<string>();

            try
            {
                if (Directory.Exists(_tempScreenshotDir))
                {
                    foreach (var file in Directory.GetFiles(_tempScreenshotDir, "*.png"))
                    {
                        try
                        {
                            if (File.GetCreationTime(file) < cutoff)
                            {
                                File.Delete(file);
                                deleted++;
                            }
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            errors.Add($"{Path.GetFileName(file)}: {ex.Message}");
                        }
                    }
                }

                WriteSuccess("cleanup_screenshots", new
                {
                    deleted,
                    failed,
                    directory = _tempScreenshotDir,
                    max_age_minutes = maxAgeMinutes,
                    errors = errors.Count > 0 ? errors : null
                });
            }
            catch (Exception ex)
            {
                WriteError("CLEANUP_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Performs auto-cleanup of old screenshots on startup
        /// </summary>
        private static void AutoCleanupScreenshots()
        {
            try
            {
                if (!Directory.Exists(_tempScreenshotDir)) return;

                var cutoff = DateTime.Now - _screenshotAutoCleanupAge;
                int deleted = 0;

                foreach (var file in Directory.GetFiles(_tempScreenshotDir, "*.png"))
                {
                    try
                    {
                        if (File.GetCreationTime(file) < cutoff)
                        {
                            File.Delete(file);
                            deleted++;
                        }
                    }
                    catch
                    {
                        // Silently ignore cleanup errors on startup
                    }
                }

                if (deleted > 0 && _debugMode)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(new
                    {
                        status = "info",
                        action = "auto_cleanup",
                        deleted,
                        directory = _tempScreenshotDir
                    }));
                }
            }
            catch
            {
                // Silently ignore cleanup errors on startup
            }
        }

        // ==================== TIER 1 COMMANDS ====================

        /// <summary>
        /// Brings a window to the foreground by selector (Name or AutomationId)
        /// Uses robust Win32 API technique to force focus even when Windows blocks it
        /// </summary>
        private static void HandleFocusWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var window = FindWindow(selector);
            if (window != null)
            {
                try
                {
                    var hwnd = window.Properties.NativeWindowHandle.Value;
                    ForceSetForegroundWindow(hwnd);
                    
                    // Small delay to ensure window is ready for input
                    Thread.Sleep(100);
                    
                    WriteSuccess("focus_window", new { window = SafeGet(() => window.Title, "Unknown") });
                }
                catch (Exception ex)
                {
                    WriteError("FOCUS_FAILED", $"Could not focus window: {ex.Message}");
                }
            }
            else
            {
                WriteWindowNotFound(selector);
            }
        }

        /// <summary>
        /// Forces a window to the foreground using Win32 API tricks
        /// This bypasses Windows restrictions on stealing focus
        /// </summary>
        private static void ForceSetForegroundWindow(IntPtr hwnd)
        {
            // Get the current foreground window's thread
            IntPtr foregroundHwnd = GetForegroundWindow();
            uint foregroundThread = GetWindowThreadProcessId(foregroundHwnd, IntPtr.Zero);
            uint currentThread = GetCurrentThreadId();

            // Attach to the foreground thread to allow stealing focus
            if (foregroundThread != currentThread)
            {
                AttachThreadInput(currentThread, foregroundThread, true);
            }

            try
            {
                // Restore the window if minimized
                ShowWindow(hwnd, SW_RESTORE);
                
                // Bring window to foreground
                SetForegroundWindow(hwnd);
                
                // Use FlaUI's method as backup
                // This handles additional focus logic
            }
            finally
            {
                // Detach from foreground thread
                if (foregroundThread != currentThread)
                {
                    AttachThreadInput(currentThread, foregroundThread, false);
                }
            }
        }

        /// <summary>
        /// Closes a window by selector
        /// </summary>
        private static void HandleCloseWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var window = FindWindow(selector);
            if (window != null)
            {
                var title = SafeGet(() => window.Title, "Unknown");
                window.Close();
                WriteSuccess("close_window", new { window = title });
            }
            else
            {
                WriteWindowNotFound(selector);
            }
        }

        /// <summary>
        /// Launches an application by path or name
        /// </summary>
        private static void HandleLaunchApp(JObject command)
        {
            var path = command["path"]?.ToString();
            var args = command["args"]?.ToString() ?? "";
            var waitForWindow = command["wait_for_window"]?.ToString();

            if (string.IsNullOrEmpty(path))
            {
                WriteMissingParam("path");
                return;
            }

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = path,
                    Arguments = args,
                    UseShellExecute = true
                };
                var process = Process.Start(startInfo);

                if (!string.IsNullOrEmpty(waitForWindow) && process != null)
                {
                    // Wait up to 10 seconds for the window to appear
                    var timeout = TimeSpan.FromSeconds(10);
                    var window = Retry.WhileNull(
                        () => FindWindow(waitForWindow),
                        timeout,
                        TimeSpan.FromMilliseconds(500)
                    ).Result;

                    if (window != null)
                    {
                        WriteSuccess("launch_app", new { path, pid = process.Id, window_found = true });
                        return;
                    }
                }

                WriteSuccess("launch_app", new { path, pid = process?.Id });
            }
            catch (Exception ex)
            {
                WriteError("LAUNCH_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Sends a hotkey combination (e.g., Ctrl+C, Alt+F4)
        /// </summary>
        private static void HandleHotkey(JObject command)
        {
            var keys = command["keys"]?.ToString();
            if (string.IsNullOrEmpty(keys))
            {
                WriteError("MISSING_PARAM", "Missing 'keys' parameter (e.g., 'ctrl+c', 'alt+f4')");
                return;
            }

            try
            {
                var keyList = ParseHotkey(keys);
                if (keyList.Count == 0)
                {
                    WriteError("INVALID_KEYS", $"Could not parse keys: {keys}");
                    return;
                }

                Keyboard.TypeSimultaneously(keyList.ToArray());
                WriteSuccess("hotkey", new { keys });
            }
            catch (Exception ex)
            {
                WriteError("HOTKEY_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Explores child elements inside a specific window
        /// </summary>
        private static void HandleExploreWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var maxDepth = command["max_depth"]?.Value<int>() ?? 2;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            var elements = ExploreElementTree(window, maxDepth, 0);
            WriteSuccess(new { 
                status = "success", 
                window = SafeGet(() => window.Title, "Unknown"),
                element_count = elements.Count,
                data = elements 
            });
        }

        // ==================== TIER 2 COMMANDS ====================

        /// <summary>
        /// Waits for a window to appear with timeout
        /// </summary>
        private static async Task HandleWaitForWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var timeoutMs = command["timeout"]?.Value<int>() ?? 10000;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var timeout = TimeSpan.FromMilliseconds(timeoutMs);
            var window = Retry.WhileNull(
                () => FindWindow(selector),
                timeout,
                TimeSpan.FromMilliseconds(250)
            ).Result;

            if (window != null)
            {
                WriteSuccess("wait_for_window", new { window = SafeGet(() => window.Title, "Unknown") });
            }
            else
            {
                WriteError("TIMEOUT", $"Window '{selector}' did not appear within {timeoutMs}ms");
            }
            await Task.CompletedTask;
        }

        /// <summary>
        /// Clicks at absolute screen coordinates
        /// </summary>
        private static void HandleClickAt(JObject command)
        {
            var x = command["x"]?.Value<int>();
            var y = command["y"]?.Value<int>();
            var button = command["button"]?.ToString() ?? "left";

            if (!x.HasValue || !y.HasValue)
            {
                WriteMissingParam("x' or 'y");
                return;
            }

            var point = new System.Drawing.Point(x.Value, y.Value);
            Mouse.MoveTo(point);

            switch (button.ToLower())
            {
                case "right":
                    Mouse.RightClick();
                    break;
                case "middle":
                    Mouse.Click(MouseButton.Middle);
                    break;
                default:
                    Mouse.LeftClick();
                    break;
            }

            WriteSuccess("click_at", new { x, y, button });
        }

        /// <summary>
        /// Gets the current clipboard text content
        /// </summary>
        private static void HandleGetClipboard()
        {
            try
            {
                var text = RunOnSTAThread(() => 
                    System.Windows.Forms.Clipboard.ContainsText() 
                        ? System.Windows.Forms.Clipboard.GetText() 
                        : null);
                WriteSuccess("get_clipboard", new { text = text ?? "" });
            }
            catch (Exception ex)
            {
                WriteError("CLIPBOARD_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Sets text to the clipboard
        /// </summary>
        private static void HandleSetClipboard(JObject command)
        {
            var text = command["text"]?.ToString() ?? "";
            try
            {
                RunOnSTAThread(() => System.Windows.Forms.Clipboard.SetText(text, System.Windows.Forms.TextDataFormat.UnicodeText));
                WriteSuccess("set_clipboard");
            }
            catch (Exception ex)
            {
                WriteError("CLIPBOARD_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Sends a single key press (Enter, Escape, Tab, F1-F12, etc.)
        /// </summary>
        private static void HandleKeyPress(JObject command)
        {
            var key = command["key"]?.ToString();
            if (string.IsNullOrEmpty(key))
            {
                WriteMissingParam("key");
                return;
            }

            try
            {
                var vk = ParseSingleKey(key);
                if (vk == null)
                {
                    WriteError("INVALID_KEY", $"Unknown key: {key}");
                    return;
                }

                Keyboard.Press(vk.Value);
                Keyboard.Release(vk.Value);
                WriteSuccess("key_press", new { key });
            }
            catch (Exception ex)
            {
                WriteError("KEY_PRESS_FAILED", ex.Message);
            }
        }

        // ==================== TIER 3 COMMANDS ====================

        /// <summary>
        /// Moves a window to specified screen coordinates
        /// </summary>
        private static void HandleMoveWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var x = command["x"]?.Value<int>();
            var y = command["y"]?.Value<int>();

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }
            if (!x.HasValue || !y.HasValue)
            {
                WriteMissingParam("x' or 'y");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            try
            {
                window.Move(x.Value, y.Value);
                WriteSuccess("move_window", new { 
                    window = SafeGet(() => window.Title, "Unknown"),
                    x = x.Value,
                    y = y.Value
                });
            }
            catch (Exception ex)
            {
                WriteError("MOVE_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Resizes a window to specified dimensions
        /// </summary>
        private static void HandleResizeWindow(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var width = command["width"]?.Value<int>();
            var height = command["height"]?.Value<int>();

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }
            if (!width.HasValue || !height.HasValue)
            {
                WriteMissingParam("width' or 'height");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            try
            {
                // Get current position to maintain it
                var rect = window.BoundingRectangle;
                // FlaUI doesn't have direct resize, so we use Win32 API via pattern
                var transformPattern = window.Patterns.Transform.PatternOrDefault;
                if (transformPattern != null && transformPattern.CanResize)
                {
                    transformPattern.Resize(width.Value, height.Value);
                    WriteSuccess("resize_window", new { 
                        window = SafeGet(() => window.Title, "Unknown"),
                        width = width.Value,
                        height = height.Value
                    });
                }
                else
                {
                    // Fallback: Use Win32 API
                    SetWindowPos(window.Properties.NativeWindowHandle.Value, IntPtr.Zero, 
                        (int)rect.X, (int)rect.Y, width.Value, height.Value, 0x0004); // SWP_NOZORDER
                    WriteSuccess("resize_window", new { 
                        window = SafeGet(() => window.Title, "Unknown"),
                        width = width.Value,
                        height = height.Value,
                        method = "Win32"
                    });
                }
            }
            catch (Exception ex)
            {
                WriteError("RESIZE_FAILED", ex.Message);
            }
        }

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        /// <summary>
        /// Unified handler for window state changes (minimize, maximize, restore)
        /// </summary>
        private static void HandleWindowState(JObject command, FlaUI.Core.Definitions.WindowVisualState targetState, string actionName)
        {
            var selector = command["selector"]?.ToString();
            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            try
            {
                var windowPattern = window.Patterns.Window.PatternOrDefault;
                if (windowPattern != null)
                {
                    windowPattern.SetWindowVisualState(targetState);
                    WriteSuccess(actionName, new { window = SafeGet(() => window.Title, "Unknown") });
                }
                else
                {
                    WriteError("NOT_SUPPORTED", $"Window does not support {actionName.Replace("_", " ")}");
                }
            }
            catch (Exception ex)
            {
                WriteError($"{actionName.ToUpper().Replace("_", "")}_FAILED", ex.Message);
            }
        }

        private static void HandleMinimizeWindow(JObject command) =>
            HandleWindowState(command, FlaUI.Core.Definitions.WindowVisualState.Minimized, "minimize_window");

        private static void HandleMaximizeWindow(JObject command) =>
            HandleWindowState(command, FlaUI.Core.Definitions.WindowVisualState.Maximized, "maximize_window");

        private static void HandleRestoreWindow(JObject command) =>
            HandleWindowState(command, FlaUI.Core.Definitions.WindowVisualState.Normal, "restore_window");

        /// <summary>
        /// Lists running processes, optionally filtered by name
        /// </summary>
        private static void HandleListProcesses(JObject command)
        {
            var filter = command["filter"]?.ToString();
            var showWindows = command["show_windows"]?.Value<bool>() ?? false;

            try
            {
                var processes = Process.GetProcesses();
                
                if (!string.IsNullOrEmpty(filter))
                {
                    processes = processes.Where(p => 
                        p.ProcessName.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToArray();
                }

                var result = processes.Select(p => {
                    var info = new Dictionary<string, object?>
                    {
                        ["name"] = p.ProcessName,
                        ["pid"] = p.Id,
                        ["memory_mb"] = Math.Round(p.WorkingSet64 / 1024.0 / 1024.0, 2)
                    };
                    
                    try
                    {
                        info["responding"] = p.Responding;
                    }
                    catch { info["responding"] = null; }

                    if (showWindows)
                    {
                        try
                        {
                            info["main_window"] = string.IsNullOrEmpty(p.MainWindowTitle) ? null : p.MainWindowTitle;
                        }
                        catch { info["main_window"] = null; }
                    }

                    return info;
                }).ToList();

                WriteSuccess("list_processes", new { count = result.Count, data = result });
            }
            catch (Exception ex)
            {
                WriteError("LIST_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Kills a process by PID or name
        /// </summary>
        private static void HandleKillProcess(JObject command)
        {
            var pid = command["pid"]?.Value<int>();
            var name = command["name"]?.ToString();
            var force = command["force"]?.Value<bool>() ?? false;

            if (!pid.HasValue && string.IsNullOrEmpty(name))
            {
                WriteMissingParam("pid' or 'name");
                return;
            }

            try
            {
                Process[] targetProcesses;
                
                if (pid.HasValue)
                {
                    targetProcesses = new[] { Process.GetProcessById(pid.Value) };
                }
                else
                {
                    targetProcesses = Process.GetProcessesByName(name!);
                }

                if (targetProcesses.Length == 0)
                {
                    WriteError("PROCESS_NOT_FOUND", "No matching process found");
                    return;
                }

                var killed = new List<object>();
                foreach (var proc in targetProcesses)
                {
                    try
                    {
                        var procInfo = new { name = proc.ProcessName, pid = proc.Id };
                        
                        if (force)
                        {
                            proc.Kill();
                        }
                        else
                        {
                            proc.CloseMainWindow();
                            if (!proc.WaitForExit(3000))
                            {
                                proc.Kill();
                            }
                        }
                        
                        killed.Add(procInfo);
                    }
                    catch { /* Skip processes we can't kill */ }
                }

                WriteSuccess("kill_process", new { killed_count = killed.Count, killed });
            }
            catch (Exception ex)
            {
                WriteError("KILL_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Reads text from screen using OCR via Python bridge (WinRT native OCR)
        /// </summary>
        private static async Task HandleReadText(JObject command)
        {
            var region = command["region"]; // Optional: {x, y, width, height}
            var selector = command["selector"]?.ToString();

            try
            {
                // First try to get text from UI element directly
                if (!string.IsNullOrEmpty(selector))
                {
                    var element = FindElement(selector);
                    if (element != null)
                    {
                        // Try to get text from various patterns
                        var textPattern = element.Patterns.Text.PatternOrDefault;
                        if (textPattern != null)
                        {
                            var text = textPattern.DocumentRange.GetText(-1);
                            WriteSuccess("read_text", new { method = "TextPattern", text });
                            return;
                        }

                        // Try Value pattern
                        var valuePattern = element.Patterns.Value.PatternOrDefault;
                        if (valuePattern != null)
                        {
                            var text = valuePattern.Value;
                            WriteSuccess("read_text", new { method = "ValuePattern", text });
                            return;
                        }

                        // Try Name property
                        var name = SafeGet(() => element.Name, "");
                        if (!string.IsNullOrEmpty(name))
                        {
                            WriteSuccess("read_text", new { method = "Name", text = name });
                            return;
                        }
                    }
                }

                // Fallback to OCR via Python bridge (WinRT native OCR)
                var ocrRequest = new JObject { ["action"] = "ocr" };
                if (region != null)
                {
                    ocrRequest["region"] = region;
                }

                var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/ocr",
                    new StringContent(ocrRequest.ToString(), Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadAsStringAsync();
                    var parsed = JObject.Parse(result);
                    WriteSuccess("read_text", new { method = "OCR", text = parsed["text"]?.ToString() ?? "" });
                }
                else
                {
                    WriteError("OCR_FAILED", "OCR bridge returned error");
                }
            }
            catch (Exception ex)
            {
                WriteError("READ_TEXT_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Gets detailed information about a window
        /// </summary>
        private static void HandleGetWindowInfo(JObject command)
        {
            var selector = command["selector"]?.ToString();
            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            try
            {
                var rect = window.BoundingRectangle;
                var windowPattern = window.Patterns.Window.PatternOrDefault;
                
                var info = new Dictionary<string, object?>
                {
                    ["title"] = SafeGet(() => window.Title, ""),
                    ["automation_id"] = SafeGet(() => window.AutomationId, ""),
                    ["class_name"] = SafeGet(() => window.ClassName, ""),
                    ["process_id"] = SafeGet(() => window.Properties.ProcessId.Value, 0),
                    ["handle"] = SafeGet(() => window.Properties.NativeWindowHandle.Value.ToString(), ""),
                    ["is_enabled"] = SafeGet(() => window.IsEnabled, false),
                    ["is_offscreen"] = SafeGet(() => window.IsOffscreen, false),
                    ["bounds"] = new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height }
                };

                if (windowPattern != null)
                {
                    info["is_modal"] = SafeGet(() => windowPattern.IsModal, false);
                    info["is_topmost"] = SafeGet(() => windowPattern.IsTopmost, false);
                    info["can_minimize"] = SafeGet(() => windowPattern.CanMinimize, false);
                    info["can_maximize"] = SafeGet(() => windowPattern.CanMaximize, false);
                    info["window_state"] = SafeGet(() => windowPattern.WindowVisualState.ToString(), "Unknown");
                }

                WriteSuccess("get_window_info", new { data = info });
            }
            catch (Exception ex)
            {
                WriteError("INFO_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Finds a specific element and returns detailed info
        /// </summary>
        private static void HandleFindElement(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var windowSelector = command["window"]?.ToString();
            var controlType = command["control_type"]?.ToString();

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var searchRoot = GetSearchScope(windowSelector, selector);
            if (searchRoot == null) return;

            try
            {
                // Build search condition
                var cf = new ConditionFactory(new UIA3PropertyLibrary());
                AutomationElement? element = null;

                // Search by name or automation ID
                element = searchRoot.FindFirstDescendant(c => 
                    c.ByName(selector).Or(c.ByAutomationId(selector)));

                // If not found, try partial match
                if (element == null)
                {
                    var allElements = searchRoot.FindAllDescendants();
                    element = allElements.FirstOrDefault(e => 
                    {
                        var name = SafeGet(() => e.Name, "") ?? "";
                        var autoId = SafeGet(() => e.AutomationId, "") ?? "";
                        return name.Contains(selector, StringComparison.OrdinalIgnoreCase) ||
                               autoId.Contains(selector, StringComparison.OrdinalIgnoreCase);
                    });
                }

                if (element == null)
                {
                    WriteElementNotFound(selector);
                    return;
                }

                var rect = element.BoundingRectangle;
                var info = new Dictionary<string, object?>
                {
                    ["name"] = SafeGet(() => element.Name, ""),
                    ["automation_id"] = SafeGet(() => element.AutomationId, ""),
                    ["class_name"] = SafeGet(() => element.ClassName, ""),
                    ["control_type"] = element.ControlType.ToString(),
                    ["is_enabled"] = SafeGet(() => element.IsEnabled, false),
                    ["is_offscreen"] = SafeGet(() => element.IsOffscreen, false),
                    ["bounds"] = new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height },
                    ["clickable_point"] = SafeGet(() => {
                        var pt = element.GetClickablePoint();
                        return new { x = pt.X, y = pt.Y };
                    }, (object?)null)
                };

                WriteSuccess("find_element", new { data = info });
            }
            catch (Exception ex)
            {
                WriteError("FIND_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Performs drag and drop between two elements or coordinates
        /// </summary>
        private static void HandleDragAndDrop(JObject command)
        {
            var fromSelector = command["from_selector"]?.ToString();
            var toSelector = command["to_selector"]?.ToString();
            var fromX = command["from_x"]?.Value<int>();
            var fromY = command["from_y"]?.Value<int>();
            var toX = command["to_x"]?.Value<int>();
            var toY = command["to_y"]?.Value<int>();

            System.Drawing.Point startPoint;
            System.Drawing.Point endPoint;

            try
            {
                // Determine start point
                if (!string.IsNullOrEmpty(fromSelector))
                {
                    var fromElement = FindElement(fromSelector);
                    if (fromElement == null)
                    {
                        WriteError("ELEMENT_NOT_FOUND", $"Source element '{fromSelector}' not found");
                        return;
                    }
                    startPoint = fromElement.GetClickablePoint();
                }
                else if (fromX.HasValue && fromY.HasValue)
                {
                    startPoint = new System.Drawing.Point(fromX.Value, fromY.Value);
                }
                else
                {
                    WriteError("MISSING_PARAM", "Missing 'from_selector' or 'from_x/from_y' parameters");
                    return;
                }

                // Determine end point
                if (!string.IsNullOrEmpty(toSelector))
                {
                    var toElement = FindElement(toSelector);
                    if (toElement == null)
                    {
                        WriteError("ELEMENT_NOT_FOUND", $"Target element '{toSelector}' not found");
                        return;
                    }
                    endPoint = toElement.GetClickablePoint();
                }
                else if (toX.HasValue && toY.HasValue)
                {
                    endPoint = new System.Drawing.Point(toX.Value, toY.Value);
                }
                else
                {
                    WriteError("MISSING_PARAM", "Missing 'to_selector' or 'to_x/to_y' parameters");
                    return;
                }

                // Perform drag and drop
                Mouse.MoveTo(startPoint);
                Mouse.Down(MouseButton.Left);
                Thread.Sleep(100);
                Mouse.MoveTo(endPoint);
                Thread.Sleep(100);
                Mouse.Up(MouseButton.Left);

                WriteSuccess("drag_and_drop", new { 
                    from = new { x = startPoint.X, y = startPoint.Y },
                    to = new { x = endPoint.X, y = endPoint.Y }
                });
            }
            catch (Exception ex)
            {
                WriteError("DRAG_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Returns system health status, bridge availability, and diagnostics
        /// </summary>
        private static async Task HandleHealth()
        {
            var pythonBridgeOnline = await IsBridgeOnline(PythonBridgeUrl);

            // Get system info
            var currentProcess = Process.GetCurrentProcess();
            var windowCount = 0;
            try
            {
                if (_automation != null)
                {
                    var desktop = _automation.GetDesktop();
                    windowCount = desktop.FindAllChildren().Length;
                }
            }
            catch { }

            var health = new
            {
                status = "success",
                action = "health",
                version = "4.3",
                uptime_seconds = (DateTime.Now - currentProcess.StartTime).TotalSeconds,
                memory_mb = Math.Round(currentProcess.WorkingSet64 / 1024.0 / 1024.0, 2),
                engines = new
                {
                    flaui = new { online = true, primary = true },
                    pywinauto = new { online = pythonBridgeOnline, url = PythonBridgeUrl, includes = "vision, ocr, context" }
                },
                system = new
                {
                    os_version = Environment.OSVersion.ToString(),
                    machine_name = Environment.MachineName,
                    processor_count = Environment.ProcessorCount,
                    is_64bit = Environment.Is64BitOperatingSystem,
                    visible_windows = windowCount
                },
                timestamp = DateTime.UtcNow.ToString("o")
            };

            WriteSuccess(health);
        }

        // ==================== MULTI-MONITOR SUPPORT ====================

        /// <summary>
        /// Lists all monitors with their bounds and properties
        /// </summary>
        private static void HandleListMonitors()
        {
            var monitors = new List<object>();
            var allScreens = Screen.AllScreens;

            for (int i = 0; i < allScreens.Length; i++)
            {
                var screen = allScreens[i];
                monitors.Add(new
                {
                    index = i,
                    name = screen.DeviceName,
                    is_primary = screen.Primary,
                    bounds = new
                    {
                        x = screen.Bounds.X,
                        y = screen.Bounds.Y,
                        width = screen.Bounds.Width,
                        height = screen.Bounds.Height
                    },
                    working_area = new
                    {
                        x = screen.WorkingArea.X,
                        y = screen.WorkingArea.Y,
                        width = screen.WorkingArea.Width,
                        height = screen.WorkingArea.Height
                    },
                    bits_per_pixel = screen.BitsPerPixel
                });
            }

            WriteSuccess("list_monitors", new { count = monitors.Count, monitors });
        }

        /// <summary>
        /// Takes a screenshot of a specific monitor
        /// </summary>
        private static void HandleScreenshotMonitor(JObject command)
        {
            var monitorIndex = command["monitor"]?.Value<int>() ?? 0;
            var fileName = command["filename"]?.ToString();
            var useTemp = command["temp"]?.Value<bool>() ?? true; // Default to temp directory

            var allScreens = Screen.AllScreens;
            if (monitorIndex < 0 || monitorIndex >= allScreens.Length)
            {
                WriteError("INVALID_MONITOR", $"Monitor index {monitorIndex} is out of range. Available: 0-{allScreens.Length - 1}");
                return;
            }

            var screen = allScreens[monitorIndex];
            var bounds = screen.Bounds;

            try
            {
                string fullPath;
                bool isTemporary;

                if (fileName != null && Path.IsPathRooted(fileName))
                {
                    // Absolute path provided - user wants permanent location
                    fullPath = fileName;
                    isTemporary = false;
                }
                else
                {
                    // Use temp directory (default behavior)
                    Directory.CreateDirectory(_tempScreenshotDir);
                    var name = fileName ?? $"monitor_{monitorIndex}_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    fullPath = Path.Combine(_tempScreenshotDir, name);
                    isTemporary = useTemp;
                }

                using (var bitmap = new Bitmap(bounds.Width, bounds.Height))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(bounds.Location, System.Drawing.Point.Empty, bounds.Size);
                    }
                    bitmap.Save(fullPath, ImageFormat.Png);
                }

                WriteSuccess("screenshot_monitor", new
                {
                    monitor = monitorIndex,
                    path = fullPath,
                    temporary = isTemporary,
                    bounds = new { x = bounds.X, y = bounds.Y, width = bounds.Width, height = bounds.Height }
                });
            }
            catch (Exception ex)
            {
                WriteError("SCREENSHOT_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Moves a window to a specific monitor
        /// </summary>
        private static void HandleMoveToMonitor(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var monitorIndex = command["monitor"]?.Value<int>() ?? 0;
            var position = command["position"]?.ToString() ?? "center";

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var allScreens = Screen.AllScreens;
            if (monitorIndex < 0 || monitorIndex >= allScreens.Length)
            {
                WriteError("INVALID_MONITOR", $"Monitor index {monitorIndex} is out of range. Available: 0-{allScreens.Length - 1}");
                return;
            }

            var window = FindWindow(selector);
            if (window == null)
            {
                WriteWindowNotFound(selector);
                return;
            }

            try
            {
                var screen = allScreens[monitorIndex];
                var workingArea = screen.WorkingArea;
                var windowBounds = window.BoundingRectangle;

                int newX, newY;

                switch (position.ToLower())
                {
                    case "topleft":
                        newX = workingArea.X;
                        newY = workingArea.Y;
                        break;
                    case "topright":
                        newX = workingArea.Right - (int)windowBounds.Width;
                        newY = workingArea.Y;
                        break;
                    case "bottomleft":
                        newX = workingArea.X;
                        newY = workingArea.Bottom - (int)windowBounds.Height;
                        break;
                    case "bottomright":
                        newX = workingArea.Right - (int)windowBounds.Width;
                        newY = workingArea.Bottom - (int)windowBounds.Height;
                        break;
                    case "center":
                    default:
                        newX = workingArea.X + (workingArea.Width - (int)windowBounds.Width) / 2;
                        newY = workingArea.Y + (workingArea.Height - (int)windowBounds.Height) / 2;
                        break;
                }

                window.Move(newX, newY);

                WriteSuccess("move_to_monitor", new
                {
                    window = SafeGet(() => window.Title, "Unknown"),
                    monitor = monitorIndex,
                    position,
                    new_location = new { x = newX, y = newY }
                });
            }
            catch (Exception ex)
            {
                WriteError("MOVE_FAILED", ex.Message);
            }
        }

        // ==================== FILE DIALOG AUTOMATION ====================

        /// <summary>
        /// Helper: Selects all text and types new content
        /// </summary>
        private static void SelectAllAndType(string text)
        {
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
            Thread.Sleep(50);
            Keyboard.Type(text);
        }

        /// <summary>
        /// Automates file dialogs (Open, Save As, Browse for Folder)
        /// </summary>
        private static void HandleFileDialog(JObject command)
        {
            var dialogAction = command["dialog_action"]?.ToString() ?? "detect";
            var path = command["path"]?.ToString();
            var filename = command["filename"]?.ToString();

            // Find any open file dialog
            var dialog = FindFileDialog();
            if (dialog == null)
            {
                WriteError("DIALOG_NOT_FOUND", "No file dialog found. Ensure a file dialog is open.");
                return;
            }

            try
            {
                switch (dialogAction.ToLower())
                {
                    case "detect":
                        HandleFileDialogDetect(dialog);
                        break;
                    case "set_path":
                        HandleFileDialogSetPath(dialog, path);
                        break;
                    case "set_filename":
                        HandleFileDialogSetFilename(dialog, filename);
                        break;
                    case "confirm":
                        HandleFileDialogConfirm(dialog);
                        break;
                    case "cancel":
                        HandleFileDialogCancel(dialog);
                        break;
                    default:
                        WriteError("INVALID_ACTION", $"Unknown dialog_action: {dialogAction}. Use: detect, set_path, set_filename, confirm, cancel");
                        break;
                }
            }
            catch (Exception ex)
            {
                WriteError("DIALOG_FAILED", ex.Message);
            }
        }

        private static Window? FindFileDialog()
        {
            if (_automation == null) return null;
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren();

            // Common file dialog class names
            var dialogClassNames = new[] { "#32770", "Alternate Modal Top Most" };

            foreach (var win in windows)
            {
                var className = SafeGet(() => win.ClassName, "");
                var title = SafeGet(() => win.Name, "");

                // Check for common file dialog patterns
                if (dialogClassNames.Contains(className) ||
                    title.Contains("Open", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("Save", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("Browse", StringComparison.OrdinalIgnoreCase))
                {
                    // Verify it has file dialog elements
                    var hasFileElements = win.FindFirstDescendant(cf => 
                        cf.ByAutomationId("1148").Or(cf.ByAutomationId("1001")).Or(cf.ByName("File name:"))) != null;
                    
                    if (hasFileElements || className == "#32770")
                    {
                        return win.AsWindow();
                    }
                }
            }

            return null;
        }

        private static void HandleFileDialogDetect(Window dialog)
        {
            var title = SafeGet(() => dialog.Title, "Unknown");
            var className = SafeGet(() => dialog.ClassName, "");

            // Try to get current path from address bar
            var addressBar = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1148"));
            var currentPath = addressBar != null ? SafeGet(() => addressBar.Name, "") : "";

            // Try to get filename from filename box
            var filenameBox = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1001"));
            var currentFilename = "";
            if (filenameBox != null)
            {
                var valuePattern = filenameBox.Patterns.Value.PatternOrDefault;
                currentFilename = valuePattern != null ? SafeGet(() => valuePattern.Value, "") : "";
            }

            // Determine dialog type
            var dialogType = "unknown";
            if (title.Contains("Open", StringComparison.OrdinalIgnoreCase))
                dialogType = "open";
            else if (title.Contains("Save", StringComparison.OrdinalIgnoreCase))
                dialogType = "save";
            else if (title.Contains("Browse", StringComparison.OrdinalIgnoreCase) || title.Contains("Folder", StringComparison.OrdinalIgnoreCase))
                dialogType = "folder";

            WriteSuccess("file_dialog", new
            {
                dialog_action = "detect",
                dialog_type = dialogType,
                title,
                current_path = currentPath,
                current_filename = currentFilename
            });
        }

        private static void HandleFileDialogSetPath(Window dialog, string? path)
        {
            if (string.IsNullOrEmpty(path))
            {
                WriteMissingParam("path");
                return;
            }

            // Focus the address bar and type path
            var addressBar = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1148"));
            if (addressBar == null)
            {
                // Try clicking the breadcrumb to activate edit mode
                var breadcrumb = dialog.FindFirstDescendant(cf => cf.ByClassName("Breadcrumb Parent"));
                if (breadcrumb != null)
                {
                    breadcrumb.Click();
                    Thread.Sleep(200);
                    addressBar = dialog.FindFirstDescendant(cf => cf.ByClassName("Edit"));
                }
            }

            if (addressBar != null)
            {
                addressBar.Focus();
                SelectAllAndType(path);
                Keyboard.Press(VirtualKeyShort.ENTER);
                Keyboard.Release(VirtualKeyShort.ENTER);
                Thread.Sleep(300);

                WriteSuccess("file_dialog", new { dialog_action = "set_path", path });
            }
            else
            {
                WriteError("ELEMENT_NOT_FOUND", "Could not find address bar in file dialog");
            }
        }

        private static void HandleFileDialogSetFilename(Window dialog, string? filename)
        {
            if (string.IsNullOrEmpty(filename))
            {
                WriteMissingParam("filename");
                return;
            }

            var filenameBox = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1001"));
            if (filenameBox == null)
            {
                filenameBox = dialog.FindFirstDescendant(cf => cf.ByName("File name:"));
            }

            if (filenameBox != null)
            {
                filenameBox.Focus();
                SelectAllAndType(filename);

                WriteSuccess("file_dialog", new { dialog_action = "set_filename", filename });
            }
            else
            {
                WriteError("ELEMENT_NOT_FOUND", "Could not find filename input in file dialog");
            }
        }

        private static void HandleFileDialogConfirm(Window dialog)
        {
            // Look for Open/Save button (usually AutomationId "1" or name contains Open/Save)
            var confirmButton = FindDialogButton(dialog, "1", "Open", "Save");

            if (confirmButton != null)
            {
                confirmButton.Click();
                WriteSuccess("file_dialog", new { dialog_action = "confirm" });
            }
            else
            {
                WriteError("ELEMENT_NOT_FOUND", "Could not find confirm button in file dialog");
            }
        }

        private static void HandleFileDialogCancel(Window dialog)
        {
            var cancelButton = FindDialogButton(dialog, "2", "Cancel");

            if (cancelButton != null)
            {
                cancelButton.Click();
                WriteSuccess("file_dialog", new { dialog_action = "cancel" });
            }
            else
            {
                WriteError("ELEMENT_NOT_FOUND", "Could not find cancel button in file dialog");
            }
        }

        // ==================== BATCH COMMAND EXECUTION ====================

        /// <summary>
        /// Executes multiple commands in sequence with configurable error handling
        /// </summary>
        private static async Task HandleBatch(JObject command)
        {
            var commands = command["commands"] as JArray;
            var stopOnError = command["stop_on_error"]?.Value<bool>() ?? true;
            var delayBetweenMs = command["delay_between"]?.Value<int>() ?? 100;

            if (commands == null || commands.Count == 0)
            {
                WriteError("MISSING_PARAM", "Missing or empty 'commands' array");
                return;
            }

            var results = new List<object>();
            var succeeded = 0;
            var failed = 0;
            var startTime = DateTime.Now;

            for (int i = 0; i < commands.Count; i++)
            {
                var cmd = commands[i] as JObject;
                if (cmd == null)
                {
                    results.Add(new { index = i, status = "error", message = "Invalid command format" });
                    failed++;
                    if (stopOnError) break;
                    continue;
                }

                var action = cmd["action"]?.ToString();
                if (string.IsNullOrEmpty(action))
                {
                    results.Add(new { index = i, status = "error", message = "Missing action" });
                    failed++;
                    if (stopOnError) break;
                    continue;
                }

                try
                {
                    // Capture console output for this command
                    var originalOut = Console.Out;
                    using var sw = new StringWriter();
                    Console.SetOut(sw);

                    // Execute the command
                    await ExecuteSingleCommand(action, cmd);

                    Console.SetOut(originalOut);
                    var output = sw.ToString().Trim();

                    // Parse result
                    JObject? resultObj = null;
                    try
                    {
                        if (!string.IsNullOrEmpty(output))
                        {
                            // Get the last line (actual result)
                            var lines = output.Split('\n');
                            var lastLine = lines.LastOrDefault(l => l.Contains("\"status\""))?.Trim() ?? output;
                            resultObj = JObject.Parse(lastLine);
                        }
                    }
                    catch { }

                    var status = resultObj?["status"]?.ToString() ?? "unknown";
                    if (status == "success")
                    {
                        succeeded++;
                        results.Add(new { index = i, action, status = "success", result = resultObj });
                    }
                    else
                    {
                        failed++;
                        results.Add(new { index = i, action, status = "error", result = resultObj });
                        if (stopOnError) break;
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    results.Add(new { index = i, action, status = "error", message = ex.Message });
                    if (stopOnError) break;
                }

                // Delay between commands
                if (i < commands.Count - 1 && delayBetweenMs > 0)
                {
                    await Task.Delay(delayBetweenMs);
                }
            }

            var duration = (DateTime.Now - startTime).TotalMilliseconds;

            var batchStatus = failed == 0 ? "success" :
                            (stopOnError && failed > 0 ? "partial" : "completed_with_errors");

            WriteSuccess(new
            {
                status = batchStatus,
                action = "batch",
                total_commands = commands.Count,
                succeeded,
                failed,
                duration_ms = Math.Round(duration, 2),
                results
            });
        }

        /// <summary>
        /// Executes a single command (used by batch processing)
        /// </summary>
        private static async Task ExecuteSingleCommand(string action, JObject command)
        {
            switch (action)
            {
                case "click":
                case "double_click":
                case "right_click":
                case "long_press":
                    await HandleMouseAction(command, action);
                    break;
                case "scroll":
                    await HandleScroll(command);
                    break;
                case "explore":
                    HandleExplore();
                    break;
                case "type":
                    await HandleType(command);
                    break;
                case "screenshot":
                    HandleScreenshot(command);
                    break;
                case "cleanup_screenshots":
                    HandleCleanupScreenshots(command);
                    break;
                case "focus_window":
                    HandleFocusWindow(command);
                    break;
                case "close_window":
                    HandleCloseWindow(command);
                    break;
                case "launch_app":
                    HandleLaunchApp(command);
                    break;
                case "hotkey":
                    HandleHotkey(command);
                    break;
                case "explore_window":
                    HandleExploreWindow(command);
                    break;
                case "wait_for_window":
                    await HandleWaitForWindow(command);
                    break;
                case "click_at":
                    HandleClickAt(command);
                    break;
                case "get_clipboard":
                    HandleGetClipboard();
                    break;
                case "set_clipboard":
                    HandleSetClipboard(command);
                    break;
                case "key_press":
                    HandleKeyPress(command);
                    break;
                case "move_window":
                    HandleMoveWindow(command);
                    break;
                case "resize_window":
                    HandleResizeWindow(command);
                    break;
                case "minimize_window":
                    HandleMinimizeWindow(command);
                    break;
                case "maximize_window":
                    HandleMaximizeWindow(command);
                    break;
                case "restore_window":
                    HandleRestoreWindow(command);
                    break;
                case "list_processes":
                    HandleListProcesses(command);
                    break;
                case "kill_process":
                    HandleKillProcess(command);
                    break;
                case "read_text":
                    await HandleReadText(command);
                    break;
                case "get_window_info":
                    HandleGetWindowInfo(command);
                    break;
                case "find_element":
                    HandleFindElement(command);
                    break;
                case "drag_and_drop":
                    HandleDragAndDrop(command);
                    break;
                case "mouse_move":
                    HandleMouseMove(command);
                    break;
                case "wait_for_element":
                    await HandleWaitForElement(command);
                    break;
                case "list_monitors":
                    HandleListMonitors();
                    break;
                case "screenshot_monitor":
                    HandleScreenshotMonitor(command);
                    break;
                case "move_to_monitor":
                    HandleMoveToMonitor(command);
                    break;
                case "file_dialog":
                    HandleFileDialog(command);
                    break;
                // v3.3 commands
                case "mouse_path":
                    await HandleMousePath(command);
                    break;
                case "mouse_bezier":
                    await HandleMouseBezier(command);
                    break;
                case "draw":
                    await HandleDraw(command);
                    break;
                case "mouse_down":
                    HandleMouseDown(command);
                    break;
                case "mouse_up":
                    HandleMouseUp(command);
                    break;
                case "key_down":
                    HandleKeyDown(command);
                    break;
                case "key_up":
                    HandleKeyUp(command);
                    break;
                case "type_here":
                case "type_at_cursor":
                    HandleTypeHere(command);
                    break;
                case "wait_for_state":
                    await HandleWaitForState(command);
                    break;
                case "ocr_region":
                    await HandleOcrRegion(command);
                    break;
                case "click_relative":
                case "tap_relative":
                    HandleClickRelative(command);
                    break;
                // v3.4 commands
                case "get_cursor_position":
                    HandleGetCursorPosition();
                    break;
                case "get_element_bounds":
                    HandleGetElementBounds(command);
                    break;
                case "element_screenshot":
                    HandleElementScreenshot(command);
                    break;
                case "hover":
                    await HandleHover(command);
                    break;
                case "mouse_move_eased":
                    await HandleMouseMoveEased(command);
                    break;
                case "swipe":
                    await HandleSwipe(command);
                    break;
                case "wait_for_color":
                    await HandleWaitForColor(command);
                    break;
                case "wait_for_idle":
                    await HandleWaitForIdle(command);
                    break;
                case "wait_for_text":
                    await HandleWaitForText(command);
                    break;
                // v3.5 commands
                case "subscribe":
                    HandleSubscribe(command);
                    break;
                case "unsubscribe":
                    HandleUnsubscribe(command);
                    break;
                case "get_subscriptions":
                    HandleGetSubscriptions();
                    break;
                case "get_metrics":
                    HandleGetMetrics();
                    break;
                case "clear_metrics":
                    HandleClearMetrics();
                    break;
                case "set_debug_mode":
                    HandleSetDebugMode(command);
                    break;
                case "get_cache_stats":
                    HandleGetCacheStats();
                    break;
                case "clear_cache":
                    HandleClearCache();
                    break;
                // v3.6 commands
                case "set_human_mode":
                    HandleSetHumanMode(command);
                    break;
                case "get_human_mode":
                    HandleGetHumanMode();
                    break;
                case "human_click":
                    await HandleHumanClick(command);
                    break;
                case "human_type":
                    await HandleHumanType(command);
                    break;
                case "human_move":
                    await HandleHumanMove(command);
                    break;
                // v4.0 Token-Efficient Commands
                case "element_exists":
                    await HandleElementExists(command);
                    break;
                case "get_interactive_elements":
                    HandleGetInteractiveElements(command);
                    break;
                case "get_window_summary":
                    HandleGetWindowSummary(command);
                    break;
                case "describe_element":
                    await HandleDescribeElement(command);
                    break;
                case "get_element_brief":
                    HandleGetElementBrief(command);
                    break;
                case "fuzzy_find_element":
                    await HandleFuzzyFindElement(command);
                    break;
                case "vision_click":
                    await HandleVisionClick(command);
                    break;
                // v4.0 Smart Fallback Commands
                case "smart_click":
                    await HandleSmartClick(command);
                    break;
                case "smart_type":
                    await HandleSmartType(command);
                    break;
                // v4.3 Agent Vision Commands - proxy to Python bridge
                case "vision_screenshot":
                    await HandleVisionScreenshot(command);
                    break;
                case "vision_screenshot_region":
                    await HandleVisionScreenshotRegion(command);
                    break;
                case "vision_config":
                    await HandleVisionConfig(command);
                    break;
                case "vision_analyze_smart":
                    await HandleVisionAnalyzeSmart(command);
                    break;
                case "vision_screenshot_cache_stats":
                    await HandleVisionScreenshotCacheStats();
                    break;
                case "vision_screenshot_cache_clear":
                    await HandleVisionScreenshotCacheClear();
                    break;
                case "vision_stream":
                    HandleVisionStream(command);
                    break;
                // Utility commands
                case "health":
                    await HandleHealth();
                    break;
                default:
                    WriteError("UNKNOWN_ACTION", $"Unknown action: {action}");
                    break;
            }
        }

        /// <summary>
        /// Gets the search scope based on optional window selector
        /// </summary>
        private static AutomationElement? GetSearchScope(string? windowSelector, string? selector)
        {
            if (!string.IsNullOrEmpty(windowSelector))
            {
                var window = FindWindow(windowSelector);
                if (window == null)
                {
                    WriteWindowNotFound(windowSelector);
                    return null;
                }
                return window;
            }
            return _automation?.GetDesktop();
        }

        /// <summary>
        /// Waits for an element to appear with configurable timeout and polling
        /// </summary>
        private static async Task HandleWaitForElement(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var windowSelector = command["window"]?.ToString();
            var timeoutMs = command["timeout"]?.Value<int>() ?? 10000;
            var pollMs = command["poll_interval"]?.Value<int>() ?? 250;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var searchRoot = GetSearchScope(windowSelector, selector);
            if (searchRoot == null) return;

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            AutomationElement? element = null;

            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                element = FindElementInScope(searchRoot, selector);
                if (element != null)
                {
                    var rect = element.BoundingRectangle;
                    WriteSuccess("wait_for_element", new
                    {
                        found = true,
                        elapsed_ms = stopwatch.ElapsedMilliseconds,
                        element = new
                        {
                            name = SafeGet(() => element.Name, ""),
                            automation_id = SafeGet(() => element.AutomationId, ""),
                            control_type = element.ControlType.ToString(),
                            bounds = new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height }
                        }
                    });
                    return;
                }

                await Task.Delay(pollMs);
            }

            WriteError("TIMEOUT", $"Element '{selector}' did not appear within {timeoutMs}ms");
        }

        /// <summary>
        /// Finds element within a specific scope (window or desktop)
        /// </summary>
        private static AutomationElement? FindElementInScope(AutomationElement scope, string selector)
        {
            // Try exact match first
            var element = scope.FindFirstDescendant(cf => cf.ByAutomationId(selector).Or(cf.ByName(selector)));
            if (element != null) return element;

            // Fallback to partial match
            var allElements = scope.FindAllDescendants();
            return allElements.FirstOrDefault(e =>
            {
                var name = SafeGet(() => e.Name, "") ?? "";
                var autoId = SafeGet(() => e.AutomationId, "") ?? "";
                return name.Contains(selector, StringComparison.OrdinalIgnoreCase) ||
                       autoId.Contains(selector, StringComparison.OrdinalIgnoreCase);
            });
        }

        /// <summary>
        /// Moves the mouse cursor to specified coordinates
        /// </summary>
        private static void HandleMouseMove(JObject command)
        {
            var x = command["x"]?.Value<int>();
            var y = command["y"]?.Value<int>();
            var selector = command["selector"]?.ToString();

            System.Drawing.Point targetPoint;

            if (!string.IsNullOrEmpty(selector))
            {
                var element = FindElement(selector);
                if (element == null)
                {
                    WriteElementNotFound(selector);
                    return;
                }
                targetPoint = element.GetClickablePoint();
            }
            else if (x.HasValue && y.HasValue)
            {
                targetPoint = new System.Drawing.Point(x.Value, y.Value);
            }
            else
            {
                WriteError("MISSING_PARAM", "Missing 'selector' or 'x/y' parameters");
                return;
            }

            try
            {
                Mouse.MoveTo(targetPoint);
                WriteSuccess("mouse_move", new { x = targetPoint.X, y = targetPoint.Y });
            }
            catch (Exception ex)
            {
                WriteError("MOVE_FAILED", ex.Message);
            }
        }

        // ==================== v3.3 MOUSE PATH & BEZIER ====================

        /// <summary>
        /// Moves mouse along an array of points with optional duration
        /// </summary>
        private static async Task HandleMousePath(JObject command)
        {
            var pointsArray = command["points"] as JArray;
            var durationMs = command["duration"]?.Value<int>() ?? 500;
            var steps = command["steps"]?.Value<int>();

            if (pointsArray == null || pointsArray.Count < 2)
            {
                WriteError("MISSING_PARAM", "Missing or invalid 'points' array (need at least 2 points)");
                return;
            }

            try
            {
                var points = new List<System.Drawing.Point>();
                foreach (var pt in pointsArray)
                {
                    if (pt is JArray arr && arr.Count >= 2)
                    {
                        points.Add(new System.Drawing.Point(arr[0].Value<int>(), arr[1].Value<int>()));
                    }
                    else if (pt is JObject obj)
                    {
                        points.Add(new System.Drawing.Point(obj["x"]!.Value<int>(), obj["y"]!.Value<int>()));
                    }
                }

                if (points.Count < 2)
                {
                    WriteError("INVALID_PARAM", "Could not parse at least 2 valid points");
                    return;
                }

                // Calculate total path length for timing
                var totalSteps = steps ?? Math.Max(points.Count * 10, 20);
                var delayPerStep = Math.Max(1, durationMs / totalSteps);

                // Interpolate between points
                var interpolatedPoints = InterpolatePoints(points, totalSteps);

                foreach (var point in interpolatedPoints)
                {
                    Mouse.MoveTo(point);
                    await Task.Delay(delayPerStep);
                }

                WriteSuccess("mouse_path", new
                {
                    points_count = points.Count,
                    total_steps = interpolatedPoints.Count,
                    duration_ms = durationMs
                });
            }
            catch (Exception ex)
            {
                WriteError("MOUSE_PATH_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Interpolates points along a path for smooth movement
        /// </summary>
        private static List<System.Drawing.Point> InterpolatePoints(List<System.Drawing.Point> waypoints, int totalSteps)
        {
            var result = new List<System.Drawing.Point>();
            if (waypoints.Count < 2) return waypoints;

            // Calculate segment lengths
            var segmentLengths = new List<double>();
            double totalLength = 0;
            for (int i = 0; i < waypoints.Count - 1; i++)
            {
                var dx = waypoints[i + 1].X - waypoints[i].X;
                var dy = waypoints[i + 1].Y - waypoints[i].Y;
                var length = Math.Sqrt(dx * dx + dy * dy);
                segmentLengths.Add(length);
                totalLength += length;
            }

            // Generate interpolated points
            for (int step = 0; step <= totalSteps; step++)
            {
                double t = (double)step / totalSteps;
                double targetDistance = t * totalLength;

                // Find which segment we're in
                double accumulatedDistance = 0;
                for (int i = 0; i < segmentLengths.Count; i++)
                {
                    if (accumulatedDistance + segmentLengths[i] >= targetDistance || i == segmentLengths.Count - 1)
                    {
                        double segmentT = segmentLengths[i] > 0
                            ? (targetDistance - accumulatedDistance) / segmentLengths[i]
                            : 0;
                        segmentT = Math.Max(0, Math.Min(1, segmentT));

                        int x = (int)(waypoints[i].X + (waypoints[i + 1].X - waypoints[i].X) * segmentT);
                        int y = (int)(waypoints[i].Y + (waypoints[i + 1].Y - waypoints[i].Y) * segmentT);
                        result.Add(new System.Drawing.Point(x, y));
                        break;
                    }
                    accumulatedDistance += segmentLengths[i];
                }
            }

            return result;
        }

        /// <summary>
        /// Moves mouse along a cubic bezier curve
        /// </summary>
        private static async Task HandleMouseBezier(JObject command)
        {
            // Parse control points
            var start = ParsePoint(command["start"]);
            var control1 = ParsePoint(command["control1"]);
            var control2 = ParsePoint(command["control2"]);
            var end = ParsePoint(command["end"]);
            var steps = command["steps"]?.Value<int>() ?? 50;
            var durationMs = command["duration"]?.Value<int>() ?? 500;

            if (start == null || end == null)
            {
                WriteError("MISSING_PARAM", "Missing 'start' or 'end' point");
                return;
            }

            // If control points not specified, use linear interpolation
            control1 ??= start;
            control2 ??= end;

            try
            {
                var delayPerStep = Math.Max(1, durationMs / steps);

                for (int i = 0; i <= steps; i++)
                {
                    double t = (double)i / steps;
                    var point = CalculateBezierPoint(t, start.Value, control1.Value, control2.Value, end.Value);
                    Mouse.MoveTo(point);
                    await Task.Delay(delayPerStep);
                }

                WriteSuccess("mouse_bezier", new
                {
                    start = new { x = start.Value.X, y = start.Value.Y },
                    end = new { x = end.Value.X, y = end.Value.Y },
                    steps,
                    duration_ms = durationMs
                });
            }
            catch (Exception ex)
            {
                WriteError("MOUSE_BEZIER_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Parses a point from JSON (array or object)
        /// </summary>
        private static System.Drawing.Point? ParsePoint(JToken? token)
        {
            if (token == null) return null;

            if (token is JArray arr && arr.Count >= 2)
            {
                return new System.Drawing.Point(arr[0].Value<int>(), arr[1].Value<int>());
            }
            else if (token is JObject obj)
            {
                var x = obj["x"]?.Value<int>();
                var y = obj["y"]?.Value<int>();
                if (x.HasValue && y.HasValue)
                {
                    return new System.Drawing.Point(x.Value, y.Value);
                }
            }
            return null;
        }

        /// <summary>
        /// Calculates a point on a cubic bezier curve at parameter t
        /// </summary>
        private static System.Drawing.Point CalculateBezierPoint(double t, System.Drawing.Point p0, System.Drawing.Point p1, System.Drawing.Point p2, System.Drawing.Point p3)
        {
            double u = 1 - t;
            double tt = t * t;
            double uu = u * u;
            double uuu = uu * u;
            double ttt = tt * t;

            double x = uuu * p0.X + 3 * uu * t * p1.X + 3 * u * tt * p2.X + ttt * p3.X;
            double y = uuu * p0.Y + 3 * uu * t * p1.Y + 3 * u * tt * p2.Y + ttt * p3.Y;

            return new System.Drawing.Point((int)x, (int)y);
        }

        /// <summary>
        /// Performs drawing: mouse down, move along path, mouse up
        /// </summary>
        private static async Task HandleDraw(JObject command)
        {
            var pointsArray = command["points"] as JArray;
            var pathArray = command["path"] as JArray; // Alias for points
            var button = command["button"]?.ToString()?.ToLower() ?? "left";
            var durationMs = command["duration"]?.Value<int>() ?? 500;

            var actualPoints = pointsArray ?? pathArray;

            if (actualPoints == null || actualPoints.Count < 2)
            {
                WriteError("MISSING_PARAM", "Missing or invalid 'points' or 'path' array (need at least 2 points)");
                return;
            }

            try
            {
                var points = new List<System.Drawing.Point>();
                foreach (var pt in actualPoints)
                {
                    var point = ParsePoint(pt);
                    if (point.HasValue)
                    {
                        points.Add(point.Value);
                    }
                }

                if (points.Count < 2)
                {
                    WriteError("INVALID_PARAM", "Could not parse at least 2 valid points");
                    return;
                }

                // Move to first point
                Mouse.MoveTo(points[0]);
                await Task.Delay(50);

                // Press mouse button
                var mouseButton = ParseMouseButton(button);
                Mouse.Down(mouseButton);

                // Move along path
                var totalSteps = Math.Max(points.Count * 10, 20);
                var delayPerStep = Math.Max(1, durationMs / totalSteps);
                var interpolatedPoints = InterpolatePoints(points, totalSteps);

                foreach (var point in interpolatedPoints)
                {
                    Mouse.MoveTo(point);
                    await Task.Delay(delayPerStep);
                }

                // Release mouse button
                Mouse.Up(mouseButton);

                WriteSuccess("draw", new
                {
                    points_count = points.Count,
                    button,
                    duration_ms = durationMs
                });
            }
            catch (Exception ex)
            {
                WriteError("DRAW_FAILED", ex.Message);
            }
        }

        // ==================== v3.3 KEY/MOUSE HOLD & RELEASE ====================

        /// <summary>
        /// Presses and holds a mouse button
        /// </summary>
        private static void HandleMouseDown(JObject command)
        {
            var button = command["button"]?.ToString()?.ToLower() ?? "left";
            var x = command["x"]?.Value<int>();
            var y = command["y"]?.Value<int>();

            try
            {
                // Move to position if specified
                if (x.HasValue && y.HasValue)
                {
                    Mouse.MoveTo(new System.Drawing.Point(x.Value, y.Value));
                }

                var mouseButton = ParseMouseButton(button);

                Mouse.Down(mouseButton);
                WriteSuccess("mouse_down", new { button, x, y });
            }
            catch (Exception ex)
            {
                WriteError("MOUSE_DOWN_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Releases a held mouse button
        /// </summary>
        private static void HandleMouseUp(JObject command)
        {
            var button = command["button"]?.ToString()?.ToLower() ?? "left";

            try
            {
                var mouseButton = ParseMouseButton(button);

                Mouse.Up(mouseButton);
                WriteSuccess("mouse_up", new { button });
            }
            catch (Exception ex)
            {
                WriteError("MOUSE_UP_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Presses and holds a keyboard key
        /// </summary>
        private static void HandleKeyDown(JObject command)
        {
            var key = command["key"]?.ToString();
            if (string.IsNullOrEmpty(key))
            {
                WriteMissingParam("key");
                return;
            }

            try
            {
                var vk = ParseSingleKey(key);
                if (vk == null)
                {
                    WriteError("INVALID_KEY", $"Unknown key: {key}");
                    return;
                }

                Keyboard.Press(vk.Value);
                WriteSuccess("key_down", new { key });
            }
            catch (Exception ex)
            {
                WriteError("KEY_DOWN_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Releases a held keyboard key
        /// </summary>
        private static void HandleKeyUp(JObject command)
        {
            var key = command["key"]?.ToString();
            if (string.IsNullOrEmpty(key))
            {
                WriteMissingParam("key");
                return;
            }

            try
            {
                var vk = ParseSingleKey(key);
                if (vk == null)
                {
                    WriteError("INVALID_KEY", $"Unknown key: {key}");
                    return;
                }

                Keyboard.Release(vk.Value);
                WriteSuccess("key_up", new { key });
            }
            catch (Exception ex)
            {
                WriteError("KEY_UP_FAILED", ex.Message);
            }
        }

        // ==================== v3.3 TYPE AT CURSOR ====================

        /// <summary>
        /// Types text at the current cursor position (no selector needed)
        /// </summary>
        private static void HandleTypeHere(JObject command)
        {
            var text = command["text"]?.ToString();
            if (string.IsNullOrEmpty(text))
            {
                WriteMissingParam("text");
                return;
            }

            try
            {
                Keyboard.Type(text);
                WriteSuccess("type_here", new { text_length = text.Length });
            }
            catch (Exception ex)
            {
                WriteError("TYPE_HERE_FAILED", ex.Message);
            }
        }

        // ==================== v3.3 ELEMENT STATE MONITORING ====================

        /// <summary>
        /// Waits for an element to reach a specific state (enabled/disabled/visible/hidden)
        /// </summary>
        private static async Task HandleWaitForState(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var windowSelector = command["window"]?.ToString();
            var state = command["state"]?.ToString()?.ToLower() ?? "enabled";
            var timeoutMs = command["timeout"]?.Value<int>() ?? 10000;
            var pollMs = command["poll_interval"]?.Value<int>() ?? 250;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var validStates = new[] { "enabled", "disabled", "visible", "hidden", "exists", "not_exists" };
            if (!validStates.Contains(state))
            {
                WriteError("INVALID_STATE", $"Invalid state: {state}. Valid: {string.Join(", ", validStates)}");
                return;
            }

            var searchRoot = GetSearchScope(windowSelector, selector);
            if (searchRoot == null) return;

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                var element = FindElementInScope(searchRoot, selector);
                bool conditionMet = false;

                switch (state)
                {
                    case "enabled":
                        conditionMet = element != null && SafeGet(() => element.IsEnabled, false);
                        break;
                    case "disabled":
                        conditionMet = element != null && !SafeGet(() => element.IsEnabled, true);
                        break;
                    case "visible":
                        conditionMet = element != null && !SafeGet(() => element.IsOffscreen, true);
                        break;
                    case "hidden":
                        conditionMet = element != null && SafeGet(() => element.IsOffscreen, false);
                        break;
                    case "exists":
                        conditionMet = element != null;
                        break;
                    case "not_exists":
                        conditionMet = element == null;
                        break;
                }

                if (conditionMet)
                {
                    var response = new Dictionary<string, object?>
                    {
                        ["status"] = "success",
                        ["action"] = "wait_for_state",
                        ["state"] = state,
                        ["elapsed_ms"] = stopwatch.ElapsedMilliseconds
                    };

                    if (element != null && state != "not_exists")
                    {
                        var rect = element.BoundingRectangle;
                        response["element"] = new
                        {
                            name = SafeGet(() => element.Name, ""),
                            automation_id = SafeGet(() => element.AutomationId, ""),
                            is_enabled = SafeGet(() => element.IsEnabled, false),
                            is_offscreen = SafeGet(() => element.IsOffscreen, false),
                            bounds = new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height }
                        };
                    }

                    WriteSuccess(response);
                    return;
                }

                await Task.Delay(pollMs);
            }

            WriteError("TIMEOUT", $"Element '{selector}' did not reach state '{state}' within {timeoutMs}ms");
        }

        // ==================== v3.3 OCR WITH REGION ====================

        /// <summary>
        /// Performs OCR on a specific screen region using Python bridge (WinRT native OCR)
        /// </summary>
        private static async Task HandleOcrRegion(JObject command)
        {
            var x = command["x"]?.Value<int>();
            var y = command["y"]?.Value<int>();
            var width = command["width"]?.Value<int>();
            var height = command["height"]?.Value<int>();

            if (!x.HasValue || !y.HasValue || !width.HasValue || !height.HasValue)
            {
                WriteError("MISSING_PARAM", "Missing required parameters: x, y, width, height");
                return;
            }

            try
            {
                // Build OCR request with region for Python bridge
                var ocrRequest = new JObject
                {
                    ["x"] = x.Value,
                    ["y"] = y.Value,
                    ["width"] = width.Value,
                    ["height"] = height.Value
                };

                var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/ocr_region",
                    new StringContent(ocrRequest.ToString(), Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadAsStringAsync();
                    var parsed = JObject.Parse(result);
                    WriteSuccess("ocr_region", new
                    {
                        region = new { x = x.Value, y = y.Value, width = width.Value, height = height.Value },
                        text = parsed["text"]?.ToString() ?? ""
                    });
                }
                else
                {
                    WriteError("OCR_FAILED", "OCR bridge returned error");
                }
            }
            catch (HttpRequestException)
            {
                WriteError("OCR_UNAVAILABLE", "Python bridge is not available. Start it with: python bridge.py");
            }
            catch (Exception ex)
            {
                WriteError("OCR_REGION_FAILED", ex.Message);
            }
        }

        // ==================== v3.3 RELATIVE CLICK ====================

        /// <summary>
        /// Clicks at a position relative to an element or window
        /// </summary>
        private static void HandleClickRelative(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var offsetX = command["offset_x"]?.Value<int>() ?? command["x"]?.Value<int>() ?? 0;
            var offsetY = command["offset_y"]?.Value<int>() ?? command["y"]?.Value<int>() ?? 0;
            var button = command["button"]?.ToString()?.ToLower() ?? "left";
            var anchor = command["anchor"]?.ToString()?.ToLower() ?? "center";

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            try
            {
                // Try to find as element first, then as window
                AutomationElement? element = FindElement(selector);
                if (element == null)
                {
                    WriteElementNotFound(selector);
                    return;
                }

                var rect = element.BoundingRectangle;

                // Calculate anchor point
                int anchorX, anchorY;
                switch (anchor)
                {
                    case "topleft":
                        anchorX = (int)rect.X;
                        anchorY = (int)rect.Y;
                        break;
                    case "topright":
                        anchorX = (int)(rect.X + rect.Width);
                        anchorY = (int)rect.Y;
                        break;
                    case "bottomleft":
                        anchorX = (int)rect.X;
                        anchorY = (int)(rect.Y + rect.Height);
                        break;
                    case "bottomright":
                        anchorX = (int)(rect.X + rect.Width);
                        anchorY = (int)(rect.Y + rect.Height);
                        break;
                    case "center":
                    default:
                        anchorX = (int)(rect.X + rect.Width / 2);
                        anchorY = (int)(rect.Y + rect.Height / 2);
                        break;
                }

                var targetX = anchorX + offsetX;
                var targetY = anchorY + offsetY;

                Mouse.MoveTo(new System.Drawing.Point(targetX, targetY));

                var mouseButton = ParseMouseButton(button);
                Mouse.Click(mouseButton);

                WriteSuccess("click_relative", new
                {
                    selector,
                    anchor,
                    offset_x = offsetX,
                    offset_y = offsetY,
                    clicked_at = new { x = targetX, y = targetY },
                    button
                });
            }
            catch (Exception ex)
            {
                WriteError("CLICK_RELATIVE_FAILED", ex.Message);
            }
        }

        // ==================== CORE HELPER METHODS ====================

        /// <summary>
        /// Finds a window by selector (supports partial/contains matching on Name or AutomationId)
        /// </summary>
        private static Window? FindWindow(string? selector)
        {
            if (string.IsNullOrEmpty(selector) || _automation == null) return null;
            var desktop = _automation.GetDesktop();
            
            // Try exact match first
            var element = desktop.FindFirstChild(cf => 
                cf.ByAutomationId(selector).Or(cf.ByName(selector)));
            
            if (element != null)
                return element.AsWindow();

            // Fallback to contains/partial match
            var windows = desktop.FindAllChildren();
            foreach (var win in windows)
            {
                var name = SafeGet(() => win.Name, "") ?? "";
                var autoId = SafeGet(() => win.AutomationId, "") ?? "";
                
                if (name.Contains(selector, StringComparison.OrdinalIgnoreCase) ||
                    autoId.Contains(selector, StringComparison.OrdinalIgnoreCase))
                {
                    return win.AsWindow();
                }
            }
            
            return null;
        }

        /// <summary>
        /// Finds an element by selector (supports partial/contains matching)
        /// </summary>
        private static AutomationElement? FindElement(string? selector)
        {
            if (string.IsNullOrEmpty(selector) || _automation == null) return null;
            var desktop = _automation.GetDesktop();
            
            // Try exact match first
            var element = desktop.FindFirstDescendant(cf => cf.ByAutomationId(selector).Or(cf.ByName(selector)));
            if (element != null) return element;

            // Fallback to contains/partial match on windows
            var windows = desktop.FindAllChildren();
            foreach (var win in windows)
            {
                var name = SafeGet(() => win.Name, "") ?? "";
                var autoId = SafeGet(() => win.AutomationId, "") ?? "";
                
                if (name.Contains(selector, StringComparison.OrdinalIgnoreCase) ||
                    autoId.Contains(selector, StringComparison.OrdinalIgnoreCase))
                {
                    return win;
                }
            }
            
            return null;
        }

        /// <summary>
        /// Recursively explores element tree up to maxDepth
        /// </summary>
        private static List<object> ExploreElementTree(AutomationElement parent, int maxDepth, int currentDepth)
        {
            var result = new List<object>();
            if (currentDepth >= maxDepth) return result;

            var children = parent.FindAllChildren();
            foreach (var child in children)
            {
                var elementInfo = new Dictionary<string, object?>
                {
                    ["Name"] = SafeGet(() => child.Name, ""),
                    ["Type"] = child.ControlType.ToString(),
                    ["AutomationId"] = SafeGet(() => child.AutomationId, ""),
                    ["ClassName"] = SafeGet(() => child.ClassName, ""),
                    ["IsEnabled"] = SafeGet(() => child.IsEnabled, false),
                    ["Rect"] = SafeGet(() => child.BoundingRectangle.ToString(), "")
                };

                if (currentDepth + 1 < maxDepth)
                {
                    var subChildren = ExploreElementTree(child, maxDepth, currentDepth + 1);
                    if (subChildren.Count > 0)
                    {
                        elementInfo["Children"] = subChildren;
                    }
                }

                result.Add(elementInfo);
            }
            return result;
        }

        /// <summary>
        /// Parses hotkey string (e.g., "ctrl+shift+a") into VirtualKeyShort array
        /// </summary>
        private static List<VirtualKeyShort> ParseHotkey(string keys)
        {
            var result = new List<VirtualKeyShort>();
            var parts = keys.ToLower().Split('+');

            foreach (var part in parts)
            {
                var key = part.Trim();
                var vk = ParseSingleKey(key);
                if (vk.HasValue)
                {
                    result.Add(vk.Value);
                }
            }

            return result;
        }

        /// <summary>
        /// Parses a single key name to VirtualKeyShort
        /// </summary>
        private static VirtualKeyShort? ParseSingleKey(string key)
        {
            return key.ToLower() switch
            {
                // Modifiers
                "ctrl" or "control" => VirtualKeyShort.CONTROL,
                "alt" => VirtualKeyShort.ALT,
                "shift" => VirtualKeyShort.SHIFT,
                "win" or "windows" or "lwin" => VirtualKeyShort.LWIN,
                "rwin" => VirtualKeyShort.RWIN,

                // Function keys
                "f1" => VirtualKeyShort.F1,
                "f2" => VirtualKeyShort.F2,
                "f3" => VirtualKeyShort.F3,
                "f4" => VirtualKeyShort.F4,
                "f5" => VirtualKeyShort.F5,
                "f6" => VirtualKeyShort.F6,
                "f7" => VirtualKeyShort.F7,
                "f8" => VirtualKeyShort.F8,
                "f9" => VirtualKeyShort.F9,
                "f10" => VirtualKeyShort.F10,
                "f11" => VirtualKeyShort.F11,
                "f12" => VirtualKeyShort.F12,

                // Navigation
                "enter" or "return" => VirtualKeyShort.ENTER,
                "tab" => VirtualKeyShort.TAB,
                "escape" or "esc" => VirtualKeyShort.ESCAPE,
                "space" or "spacebar" => VirtualKeyShort.SPACE,
                "backspace" or "back" => VirtualKeyShort.BACK,
                "delete" or "del" => VirtualKeyShort.DELETE,
                "insert" or "ins" => VirtualKeyShort.INSERT,
                "home" => VirtualKeyShort.HOME,
                "end" => VirtualKeyShort.END,
                "pageup" or "pgup" => VirtualKeyShort.PRIOR,
                "pagedown" or "pgdn" => VirtualKeyShort.NEXT,

                // Arrow keys
                "up" or "arrowup" => VirtualKeyShort.UP,
                "down" or "arrowdown" => VirtualKeyShort.DOWN,
                "left" or "arrowleft" => VirtualKeyShort.LEFT,
                "right" or "arrowright" => VirtualKeyShort.RIGHT,

                // Letters (A-Z)
                "a" => VirtualKeyShort.KEY_A,
                "b" => VirtualKeyShort.KEY_B,
                "c" => VirtualKeyShort.KEY_C,
                "d" => VirtualKeyShort.KEY_D,
                "e" => VirtualKeyShort.KEY_E,
                "f" => VirtualKeyShort.KEY_F,
                "g" => VirtualKeyShort.KEY_G,
                "h" => VirtualKeyShort.KEY_H,
                "i" => VirtualKeyShort.KEY_I,
                "j" => VirtualKeyShort.KEY_J,
                "k" => VirtualKeyShort.KEY_K,
                "l" => VirtualKeyShort.KEY_L,
                "m" => VirtualKeyShort.KEY_M,
                "n" => VirtualKeyShort.KEY_N,
                "o" => VirtualKeyShort.KEY_O,
                "p" => VirtualKeyShort.KEY_P,
                "q" => VirtualKeyShort.KEY_Q,
                "r" => VirtualKeyShort.KEY_R,
                "s" => VirtualKeyShort.KEY_S,
                "t" => VirtualKeyShort.KEY_T,
                "u" => VirtualKeyShort.KEY_U,
                "v" => VirtualKeyShort.KEY_V,
                "w" => VirtualKeyShort.KEY_W,
                "x" => VirtualKeyShort.KEY_X,
                "y" => VirtualKeyShort.KEY_Y,
                "z" => VirtualKeyShort.KEY_Z,

                // Numbers
                "0" => VirtualKeyShort.KEY_0,
                "1" => VirtualKeyShort.KEY_1,
                "2" => VirtualKeyShort.KEY_2,
                "3" => VirtualKeyShort.KEY_3,
                "4" => VirtualKeyShort.KEY_4,
                "5" => VirtualKeyShort.KEY_5,
                "6" => VirtualKeyShort.KEY_6,
                "7" => VirtualKeyShort.KEY_7,
                "8" => VirtualKeyShort.KEY_8,
                "9" => VirtualKeyShort.KEY_9,

                // Special
                "printscreen" or "prtsc" => VirtualKeyShort.SNAPSHOT,
                "pause" => VirtualKeyShort.PAUSE,
                "capslock" or "caps" => VirtualKeyShort.CAPITAL,
                "numlock" => VirtualKeyShort.NUMLOCK,
                "scrolllock" => VirtualKeyShort.SCROLL,

                _ => null
            };
        }

        private static async Task<bool> CallBridge(string baseUrl, string action, JObject data)
        {
            try
            {
                var response = await _httpClient.PostAsync($"{baseUrl}/{action}", 
                    new StringContent(data.ToString(), Encoding.UTF8, "application/json"));
                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(new { status = "success", engine = "pywinauto", action }));
                    return true;
                }
            }
            catch { /* Ignore and continue fallback */ }
            return false;
        }

        // ==================== v3.4 COMMAND HANDLERS ====================

        /// <summary>
        /// Get current cursor/mouse position
        /// </summary>
        private static void HandleGetCursorPosition()
        {
            var pos = System.Windows.Forms.Cursor.Position;
            WriteSuccess("get_cursor_position", new { x = pos.X, y = pos.Y });
        }

        /// <summary>
        /// Get element bounding box
        /// </summary>
        private static void HandleGetElementBounds(JObject command)
        {
            var selector = command["selector"]?.ToString();
            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var element = FindElement(selector);
            if (element == null)
            {
                WriteElementNotFound(selector);
                return;
            }

            var rect = element.BoundingRectangle;
            WriteSuccess("get_element_bounds", new
            {
                selector,
                bounds = new
                {
                    x = (int)rect.X,
                    y = (int)rect.Y,
                    width = (int)rect.Width,
                    height = (int)rect.Height,
                    center_x = (int)(rect.X + rect.Width / 2),
                    center_y = (int)(rect.Y + rect.Height / 2)
                }
            });
        }

        /// <summary>
        /// Screenshot a specific element only
        /// </summary>
        private static void HandleElementScreenshot(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var fileName = command["filename"]?.ToString();
            var useTemp = command["temp"]?.Value<bool>() ?? true; // Default to temp directory

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var element = FindElement(selector);
            if (element == null)
            {
                WriteElementNotFound(selector);
                return;
            }

            try
            {
                string fullPath;
                bool isTemporary;

                if (fileName != null && Path.IsPathRooted(fileName))
                {
                    // Absolute path provided - user wants permanent location
                    fullPath = fileName;
                    isTemporary = false;
                }
                else
                {
                    // Use temp directory (default behavior)
                    Directory.CreateDirectory(_tempScreenshotDir);
                    var name = fileName ?? $"element_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                    fullPath = Path.Combine(_tempScreenshotDir, name);
                    isTemporary = useTemp;
                }

                var rect = element.BoundingRectangle;
                using (var bitmap = new Bitmap((int)rect.Width, (int)rect.Height))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen((int)rect.X, (int)rect.Y, 0, 0, bitmap.Size);
                    }
                    bitmap.Save(fullPath, ImageFormat.Png);
                }
                WriteSuccess("element_screenshot", new { selector, path = fullPath, temporary = isTemporary, bounds = rect.ToString() });
            }
            catch (Exception ex)
            {
                WriteError("ELEMENT_SCREENSHOT_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Hover over an element for a specified duration
        /// </summary>
        private static async Task HandleHover(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var duration = command["duration"]?.ToObject<int>() ?? 1000;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var element = FindElement(selector);
            if (element == null)
            {
                WriteElementNotFound(selector);
                return;
            }

            try
            {
                var point = element.GetClickablePoint();
                Mouse.MoveTo(point);
                await Task.Delay(duration);
                WriteSuccess("hover", new { selector, duration });
            }
            catch (Exception ex)
            {
                WriteError("HOVER_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Mouse move with easing functions
        /// </summary>
        private static async Task HandleMouseMoveEased(JObject command)
        {
            var x = command["x"]?.ToObject<int>() ?? 0;
            var y = command["y"]?.ToObject<int>() ?? 0;
            var duration = command["duration"]?.ToObject<int>() ?? 500;
            var easingType = command["easing"]?.ToString()?.ToLower() ?? "linear";
            var steps = command["steps"]?.ToObject<int>() ?? 50;

            var startPos = System.Windows.Forms.Cursor.Position;
            var startX = startPos.X;
            var startY = startPos.Y;

            try
            {
                for (int i = 0; i <= steps; i++)
                {
                    double t = (double)i / steps;
                    double easedT = ApplyEasing(t, easingType);
                    
                    int currentX = (int)(startX + (x - startX) * easedT);
                    int currentY = (int)(startY + (y - startY) * easedT);
                    
                    Mouse.MoveTo(new System.Drawing.Point(currentX, currentY));
                    await Task.Delay(duration / steps);
                }
                WriteSuccess("mouse_move_eased", new { x, y, easing = easingType, duration });
            }
            catch (Exception ex)
            {
                WriteError("MOUSE_MOVE_EASED_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Apply easing function to progress value
        /// </summary>
        private static double ApplyEasing(double t, string easingType)
        {
            return easingType switch
            {
                "ease-in" or "easein" => t * t,
                "ease-out" or "easeout" => 1 - (1 - t) * (1 - t),
                "ease-in-out" or "easeinout" => t < 0.5 ? 2 * t * t : 1 - Math.Pow(-2 * t + 2, 2) / 2,
                "bounce" => BounceEasing(t),
                "elastic" => ElasticEasing(t),
                "cubic" => t * t * t,
                "quart" => t * t * t * t,
                "sine" => 1 - Math.Cos(t * Math.PI / 2),
                _ => t // linear
            };
        }

        private static double BounceEasing(double t)
        {
            if (t < 1 / 2.75)
                return 7.5625 * t * t;
            else if (t < 2 / 2.75)
                return 7.5625 * (t -= 1.5 / 2.75) * t + 0.75;
            else if (t < 2.5 / 2.75)
                return 7.5625 * (t -= 2.25 / 2.75) * t + 0.9375;
            else
                return 7.5625 * (t -= 2.625 / 2.75) * t + 0.984375;
        }

        private static double ElasticEasing(double t)
        {
            if (t == 0 || t == 1) return t;
            return -Math.Pow(2, 10 * t - 10) * Math.Sin((t * 10 - 10.75) * (2 * Math.PI / 3));
        }

        /// <summary>
        /// Swipe gesture (touch-style drag)
        /// </summary>
        private static async Task HandleSwipe(JObject command)
        {
            var startX = command["start_x"]?.ToObject<int>() ?? 0;
            var startY = command["start_y"]?.ToObject<int>() ?? 0;
            var endX = command["end_x"]?.ToObject<int>() ?? 0;
            var endY = command["end_y"]?.ToObject<int>() ?? 0;
            var duration = command["duration"]?.ToObject<int>() ?? 300;
            var steps = command["steps"]?.ToObject<int>() ?? 20;

            try
            {
                Mouse.MoveTo(new System.Drawing.Point(startX, startY));
                Mouse.Down(MouseButton.Left);

                for (int i = 1; i <= steps; i++)
                {
                    double t = (double)i / steps;
                    // Use ease-out for natural swipe deceleration
                    double easedT = 1 - (1 - t) * (1 - t);
                    int currentX = (int)(startX + (endX - startX) * easedT);
                    int currentY = (int)(startY + (endY - startY) * easedT);
                    Mouse.MoveTo(new System.Drawing.Point(currentX, currentY));
                    await Task.Delay(duration / steps);
                }

                Mouse.Up(MouseButton.Left);
                WriteSuccess("swipe", new { start_x = startX, start_y = startY, end_x = endX, end_y = endY, duration });
            }
            catch (Exception ex)
            {
                WriteError("SWIPE_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Wait for a specific color at a pixel or region
        /// </summary>
        private static async Task HandleWaitForColor(JObject command)
        {
            var x = command["x"]?.ToObject<int>() ?? 0;
            var y = command["y"]?.ToObject<int>() ?? 0;
            var expectedColor = command["color"]?.ToString() ?? "#FFFFFF";
            var timeout = command["timeout"]?.ToObject<int>() ?? 10000;
            var tolerance = command["tolerance"]?.ToObject<int>() ?? 10;

            // Parse color
            Color targetColor;
            try
            {
                targetColor = ColorTranslator.FromHtml(expectedColor);
            }
            catch
            {
                WriteError("INVALID_COLOR", $"Cannot parse color: {expectedColor}");
                return;
            }

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeout)
            {
                try
                {
                    using (var bitmap = new Bitmap(1, 1))
                    {
                        using (var g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(x, y, 0, 0, new Size(1, 1));
                        }
                        var pixelColor = bitmap.GetPixel(0, 0);
                        
                        if (ColorMatch(pixelColor, targetColor, tolerance))
                        {
                            WriteSuccess("wait_for_color", new
                            {
                                x, y,
                                expected_color = expectedColor,
                                found_color = ColorTranslator.ToHtml(pixelColor),
                                elapsed_ms = stopwatch.ElapsedMilliseconds
                            });
                            return;
                        }
                    }
                }
                catch { }
                await Task.Delay(100);
            }

            WriteError("TIMEOUT", $"Color {expectedColor} not found at ({x},{y}) within {timeout}ms");
        }

        private static bool ColorMatch(Color c1, Color c2, int tolerance)
        {
            return Math.Abs(c1.R - c2.R) <= tolerance &&
                   Math.Abs(c1.G - c2.G) <= tolerance &&
                   Math.Abs(c1.B - c2.B) <= tolerance;
        }

        /// <summary>
        /// Wait for UI to become idle/responsive
        /// </summary>
        private static async Task HandleWaitForIdle(JObject command)
        {
            var timeout = command["timeout"]?.ToObject<int>() ?? 10000;
            var checkInterval = command["interval"]?.ToObject<int>() ?? 200;

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                while (stopwatch.ElapsedMilliseconds < timeout)
                {
                    // Check if the foreground window is responding
                    var hwnd = GetForegroundWindow();
                    if (hwnd != IntPtr.Zero)
                    {
                        // SendMessageTimeout to check if window is responding
                        IntPtr result;
                        var responded = SendMessageTimeout(hwnd, 0, IntPtr.Zero, IntPtr.Zero, 
                            SMTO_ABORTIFHUNG, 100, out result);
                        
                        if (responded != IntPtr.Zero)
                        {
                            WriteSuccess("wait_for_idle", new { elapsed_ms = stopwatch.ElapsedMilliseconds });
                            return;
                        }
                    }
                    await Task.Delay(checkInterval);
                }
                WriteError("TIMEOUT", $"UI did not become idle within {timeout}ms");
            }
            catch (Exception ex)
            {
                WriteError("WAIT_FOR_IDLE_FAILED", ex.Message);
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, 
            uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
        
        private const uint SMTO_ABORTIFHUNG = 0x0002;

        /// <summary>
        /// Wait for specific text to appear via OCR (using Python bridge WinRT OCR)
        /// </summary>
        private static async Task HandleWaitForText(JObject command)
        {
            var text = command["text"]?.ToString();
            var timeout = command["timeout"]?.ToObject<int>() ?? 10000;
            var region = command["region"];
            var caseSensitive = command["case_sensitive"]?.ToObject<bool>() ?? false;

            if (string.IsNullOrEmpty(text))
            {
                WriteMissingParam("text");
                return;
            }

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            while (stopwatch.ElapsedMilliseconds < timeout)
            {
                try
                {
                    // Call Python bridge for OCR (WinRT native)
                    var ocrRequest = new JObject();
                    if (region != null)
                    {
                        ocrRequest["x"] = region["x"];
                        ocrRequest["y"] = region["y"];
                        ocrRequest["width"] = region["width"];
                        ocrRequest["height"] = region["height"];
                    }

                    var endpoint = region != null ? "/vision/ocr_region" : "/vision/ocr";
                    var response = await _httpClient.PostAsync($"{PythonBridgeUrl}{endpoint}",
                        new StringContent(ocrRequest.ToString(), Encoding.UTF8, "application/json"));

                    if (response.IsSuccessStatusCode)
                    {
                        var result = await response.Content.ReadAsStringAsync();
                        var json = JObject.Parse(result);
                        var ocrText = json["text"]?.ToString() ?? "";

                        bool found = caseSensitive
                            ? ocrText.Contains(text)
                            : ocrText.Contains(text, StringComparison.OrdinalIgnoreCase);

                        if (found)
                        {
                            WriteSuccess("wait_for_text", new
                            {
                                text,
                                found = true,
                                elapsed_ms = stopwatch.ElapsedMilliseconds
                            });
                            return;
                        }
                    }
                }
                catch { }
                await Task.Delay(500);
            }

            WriteError("TIMEOUT", $"Text '{text}' not found within {timeout}ms");
        }

        // ==================== v3.5 COMMAND HANDLERS ====================

        /// <summary>
        /// Subscribe to element events
        /// </summary>
        private static void HandleSubscribe(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var eventType = command["event"]?.ToString() ?? "PropertyChanged";
            var subscriptionId = command["subscription_id"]?.ToString() ?? Guid.NewGuid().ToString();

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var element = FindElement(selector);
            if (element == null)
            {
                WriteElementNotFound(selector);
                return;
            }

            try
            {
                var subscription = new EventSubscription
                {
                    Id = subscriptionId,
                    Selector = selector,
                    EventType = eventType,
                    CreatedAt = DateTime.Now
                };

                _subscriptions[subscriptionId] = subscription;
                DebugLog($"Subscribed to {eventType} on {selector}");

                WriteSuccess("subscribe", new
                {
                    subscription_id = subscriptionId,
                    selector,
                    event_type = eventType
                });
            }
            catch (Exception ex)
            {
                WriteError("SUBSCRIBE_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Unsubscribe from events
        /// </summary>
        private static void HandleUnsubscribe(JObject command)
        {
            var subscriptionId = command["subscription_id"]?.ToString();

            if (string.IsNullOrEmpty(subscriptionId))
            {
                WriteMissingParam("subscription_id");
                return;
            }

            if (_subscriptions.Remove(subscriptionId))
            {
                WriteSuccess("unsubscribe", new { subscription_id = subscriptionId });
            }
            else
            {
                WriteError("SUBSCRIPTION_NOT_FOUND", $"No subscription with id {subscriptionId}");
            }
        }

        /// <summary>
        /// Get all active subscriptions
        /// </summary>
        private static void HandleGetSubscriptions()
        {
            var subs = _subscriptions.Values.Select(s => new
            {
                subscription_id = s.Id,
                selector = s.Selector,
                event_type = s.EventType,
                created_at = s.CreatedAt
            }).ToList();

            WriteSuccess("get_subscriptions", new { count = subs.Count, subscriptions = subs });
        }

        /// <summary>
        /// Get performance metrics
        /// </summary>
        private static void HandleGetMetrics()
        {
            lock (_metricsLock)
            {
                var metricsData = _metrics.Select(kv => new
                {
                    command = kv.Key,
                    call_count = kv.Value.CallCount,
                    total_ms = kv.Value.TotalMs,
                    avg_ms = kv.Value.CallCount > 0 ? kv.Value.TotalMs / kv.Value.CallCount : 0,
                    min_ms = kv.Value.MinMs,
                    max_ms = kv.Value.MaxMs,
                    last_called = kv.Value.LastCalled
                }).ToList();

                WriteSuccess("get_metrics", new { command_count = metricsData.Count, metrics = metricsData });
            }
        }

        /// <summary>
        /// Clear performance metrics
        /// </summary>
        private static void HandleClearMetrics()
        {
            lock (_metricsLock)
            {
                _metrics.Clear();
            }
            WriteSuccess("clear_metrics", new { message = "Metrics cleared" });
        }

        /// <summary>
        /// Set debug mode
        /// </summary>
        private static void HandleSetDebugMode(JObject command)
        {
            var enabled = command["enabled"]?.ToObject<bool>() ?? false;
            _debugMode = enabled;
            WriteSuccess("set_debug_mode", new { debug_mode = _debugMode });
        }

        private static void DebugLog(string message)
        {
            if (_debugMode)
            {
                Console.Error.WriteLine($"[DEBUG {DateTime.Now:HH:mm:ss.fff}] {message}");
            }
        }

        /// <summary>
        /// Get element cache statistics
        /// </summary>
        private static void HandleGetCacheStats()
        {
            WriteSuccess("get_cache_stats", new
            {
                cached_elements = _elementCache.Count,
                cache_age_ms = (DateTime.Now - _cacheTimestamp).TotalMilliseconds,
                cache_expiry_ms = _cacheExpiry.TotalMilliseconds,
                is_valid = (DateTime.Now - _cacheTimestamp) < _cacheExpiry
            });
        }

        /// <summary>
        /// Clear element cache
        /// </summary>
        private static void HandleClearCache()
        {
            _elementCache.Clear();
            _cacheTimestamp = DateTime.MinValue;
            WriteSuccess("clear_cache", new { message = "Element cache cleared" });
        }

        // ==================== v3.6 HUMAN-LIKE COMMAND HANDLERS ====================

        /// <summary>
        /// Configure human-like behavior mode
        /// </summary>
        private static void HandleSetHumanMode(JObject command)
        {
            _humanMode.Enabled = command["enabled"]?.ToObject<bool>() ?? _humanMode.Enabled;
            _humanMode.MinPauseMs = command["min_pause_ms"]?.ToObject<int>() ?? _humanMode.MinPauseMs;
            _humanMode.MaxPauseMs = command["max_pause_ms"]?.ToObject<int>() ?? _humanMode.MaxPauseMs;
            _humanMode.MouseJitter = command["mouse_jitter"]?.ToObject<double>() ?? _humanMode.MouseJitter;
            _humanMode.TypingErrorRate = command["typing_error_rate"]?.ToObject<double>() ?? _humanMode.TypingErrorRate;
            _humanMode.TypingMinDelayMs = command["typing_min_delay_ms"]?.ToObject<int>() ?? _humanMode.TypingMinDelayMs;
            _humanMode.TypingMaxDelayMs = command["typing_max_delay_ms"]?.ToObject<int>() ?? _humanMode.TypingMaxDelayMs;
            _humanMode.FatigueEnabled = command["fatigue_enabled"]?.ToObject<bool>() ?? _humanMode.FatigueEnabled;
            _humanMode.FatigueMultiplier = command["fatigue_multiplier"]?.ToObject<double>() ?? _humanMode.FatigueMultiplier;
            _humanMode.ThinkingDelayEnabled = command["thinking_delay_enabled"]?.ToObject<bool>() ?? _humanMode.ThinkingDelayEnabled;
            _humanMode.ThinkingMinMs = command["thinking_min_ms"]?.ToObject<int>() ?? _humanMode.ThinkingMinMs;
            _humanMode.ThinkingMaxMs = command["thinking_max_ms"]?.ToObject<int>() ?? _humanMode.ThinkingMaxMs;

            if (command["reset_session"]?.ToObject<bool>() == true)
            {
                _sessionStart = DateTime.Now;
            }

            WriteSuccess("set_human_mode", new
            {
                enabled = _humanMode.Enabled,
                min_pause_ms = _humanMode.MinPauseMs,
                max_pause_ms = _humanMode.MaxPauseMs,
                mouse_jitter = _humanMode.MouseJitter,
                typing_error_rate = _humanMode.TypingErrorRate,
                fatigue_enabled = _humanMode.FatigueEnabled,
                session_duration_minutes = (DateTime.Now - _sessionStart).TotalMinutes
            });
        }

        /// <summary>
        /// Get current human mode configuration
        /// </summary>
        private static void HandleGetHumanMode()
        {
            var sessionDuration = DateTime.Now - _sessionStart;
            var fatigueFactor = _humanMode.FatigueEnabled 
                ? 1.0 + (sessionDuration.TotalMinutes / 60.0) * _humanMode.FatigueMultiplier 
                : 1.0;

            WriteSuccess("get_human_mode", new
            {
                enabled = _humanMode.Enabled,
                min_pause_ms = _humanMode.MinPauseMs,
                max_pause_ms = _humanMode.MaxPauseMs,
                mouse_jitter = _humanMode.MouseJitter,
                typing_error_rate = _humanMode.TypingErrorRate,
                typing_min_delay_ms = _humanMode.TypingMinDelayMs,
                typing_max_delay_ms = _humanMode.TypingMaxDelayMs,
                fatigue_enabled = _humanMode.FatigueEnabled,
                fatigue_multiplier = _humanMode.FatigueMultiplier,
                current_fatigue_factor = fatigueFactor,
                thinking_delay_enabled = _humanMode.ThinkingDelayEnabled,
                thinking_min_ms = _humanMode.ThinkingMinMs,
                thinking_max_ms = _humanMode.ThinkingMaxMs,
                session_start = _sessionStart,
                session_duration_minutes = sessionDuration.TotalMinutes
            });
        }

        /// <summary>
        /// Human-like click with natural mouse movement
        /// </summary>
        private static async Task HandleHumanClick(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var button = command["button"]?.ToString()?.ToLower() ?? "left";
            var doubleClick = command["double_click"]?.ToObject<bool>() ?? false;

            if (string.IsNullOrEmpty(selector))
            {
                WriteMissingParam("selector");
                return;
            }

            var element = FindElement(selector);
            if (element == null)
            {
                WriteElementNotFound(selector);
                return;
            }

            try
            {
                // Thinking delay before action
                if (_humanMode.ThinkingDelayEnabled)
                {
                    var thinkDelay = _random.Next(_humanMode.ThinkingMinMs, _humanMode.ThinkingMaxMs);
                    thinkDelay = ApplyFatigue(thinkDelay);
                    await Task.Delay(thinkDelay);
                }

                // Get target position with slight randomness
                var rect = element.BoundingRectangle;
                var targetX = (int)(rect.X + rect.Width / 2 + (_random.NextDouble() - 0.5) * rect.Width * 0.3);
                var targetY = (int)(rect.Y + rect.Height / 2 + (_random.NextDouble() - 0.5) * rect.Height * 0.3);

                // Human-like mouse movement
                await HumanMouseMoveTo(targetX, targetY);

                // Pause before clicking
                var preClickPause = _random.Next(_humanMode.MinPauseMs, _humanMode.MaxPauseMs);
                await Task.Delay(ApplyFatigue(preClickPause));

                // Perform click
                var mouseButton = ParseMouseButton(button);
                if (doubleClick)
                {
                    Mouse.DoubleClick(mouseButton);
                }
                else
                {
                    Mouse.Click(mouseButton);
                }

                // Post-click pause
                var postClickPause = _random.Next(50, 150);
                await Task.Delay(ApplyFatigue(postClickPause));

                WriteSuccess("human_click", new { selector, button, double_click = doubleClick });
            }
            catch (Exception ex)
            {
                WriteError("HUMAN_CLICK_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Human-like mouse movement using bezier curves with jitter
        /// </summary>
        private static async Task HumanMouseMoveTo(int targetX, int targetY)
        {
            var startPos = System.Windows.Forms.Cursor.Position;
            var startX = startPos.X;
            var startY = startPos.Y;

            // Calculate number of steps based on distance
            var distance = Math.Sqrt(Math.Pow(targetX - startX, 2) + Math.Pow(targetY - startY, 2));
            var steps = Math.Max(20, (int)(distance / 10));

            // Generate random control points for bezier curve
            var ctrl1X = startX + (targetX - startX) * (0.3 + _random.NextDouble() * 0.2);
            var ctrl1Y = startY + (_random.NextDouble() - 0.5) * 100;
            var ctrl2X = startX + (targetX - startX) * (0.6 + _random.NextDouble() * 0.2);
            var ctrl2Y = targetY + (_random.NextDouble() - 0.5) * 100;

            for (int i = 0; i <= steps; i++)
            {
                double t = (double)i / steps;
                
                // Cubic bezier calculation
                double x = BezierPoint(t, startX, ctrl1X, ctrl2X, targetX);
                double y = BezierPoint(t, startY, ctrl1Y, ctrl2Y, targetY);

                // Add micro-jitter
                x += (_random.NextDouble() - 0.5) * _humanMode.MouseJitter;
                y += (_random.NextDouble() - 0.5) * _humanMode.MouseJitter;

                Mouse.MoveTo(new System.Drawing.Point((int)x, (int)y));

                // Variable delay between movements
                var delay = _random.Next(5, 15);
                await Task.Delay(delay);
            }

            // Final position correction
            Mouse.MoveTo(new System.Drawing.Point(targetX, targetY));
        }

        private static double BezierPoint(double t, double p0, double p1, double p2, double p3)
        {
            double u = 1 - t;
            return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3;
        }

        /// <summary>
        /// Human-like typing with variable speed and occasional typos
        /// </summary>
        private static async Task HandleHumanType(JObject command)
        {
            var selector = command["selector"]?.ToString();
            var text = command["text"]?.ToString() ?? "";
            var correctErrors = command["correct_errors"]?.ToObject<bool>() ?? true;

            if (string.IsNullOrEmpty(text))
            {
                WriteMissingParam("text");
                return;
            }

            try
            {
                // Focus element if selector provided
                if (!string.IsNullOrEmpty(selector))
                {
                    var element = FindElement(selector);
                    if (element == null)
                    {
                        WriteElementNotFound(selector);
                        return;
                    }
                    element.Focus();
                    await Task.Delay(_random.Next(100, 300));
                }

                // Thinking delay
                if (_humanMode.ThinkingDelayEnabled)
                {
                    await Task.Delay(ApplyFatigue(_random.Next(_humanMode.ThinkingMinMs, _humanMode.ThinkingMaxMs)));
                }

                int errorsIntroduced = 0;
                int errorsCorrected = 0;

                foreach (char c in text)
                {
                    // Decide if we make a typo
                    bool makeTypo = _random.NextDouble() < _humanMode.TypingErrorRate;

                    if (makeTypo && _humanMode.Enabled)
                    {
                        // Type wrong character
                        var wrongChar = GetAdjacentKey(c);
                        Keyboard.Type(wrongChar.ToString());
                        errorsIntroduced++;

                        // Wait before noticing the error
                        await Task.Delay(ApplyFatigue(_random.Next(150, 400)));

                        if (correctErrors)
                        {
                            // Delete and retype
                            Keyboard.Press(VirtualKeyShort.BACK);
                            await Task.Delay(ApplyFatigue(_random.Next(50, 150)));
                            errorsCorrected++;
                        }
                    }

                    // Type the correct character
                    Keyboard.Type(c.ToString());

                    // Variable delay between keystrokes
                    var delay = _random.Next(_humanMode.TypingMinDelayMs, _humanMode.TypingMaxDelayMs);
                    
                    // Extra pause after punctuation
                    if (".!?;:,".Contains(c))
                    {
                        delay += _random.Next(100, 400);
                    }
                    // Occasional pause (simulating thinking)
                    else if (c == ' ' && _random.NextDouble() < 0.1)
                    {
                        delay += _random.Next(200, 500);
                    }

                    await Task.Delay(ApplyFatigue(delay));
                }

                WriteSuccess("human_type", new
                {
                    text_length = text.Length,
                    errors_introduced = errorsIntroduced,
                    errors_corrected = errorsCorrected,
                    selector
                });
            }
            catch (Exception ex)
            {
                WriteError("HUMAN_TYPE_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Human-like mouse movement to coordinates
        /// </summary>
        private static async Task HandleHumanMove(JObject command)
        {
            var x = command["x"]?.ToObject<int>();
            var y = command["y"]?.ToObject<int>();
            var selector = command["selector"]?.ToString();

            int targetX, targetY;

            if (!string.IsNullOrEmpty(selector))
            {
                var element = FindElement(selector);
                if (element == null)
                {
                    WriteElementNotFound(selector);
                    return;
                }
                var point = element.GetClickablePoint();
                targetX = (int)point.X;
                targetY = (int)point.Y;
            }
            else if (x.HasValue && y.HasValue)
            {
                targetX = x.Value;
                targetY = y.Value;
            }
            else
            {
                WriteError("MISSING_PARAM", "Either 'selector' or 'x' and 'y' required");
                return;
            }

            try
            {
                await HumanMouseMoveTo(targetX, targetY);
                WriteSuccess("human_move", new { x = targetX, y = targetY, selector });
            }
            catch (Exception ex)
            {
                WriteError("HUMAN_MOVE_FAILED", ex.Message);
            }
        }

        /// <summary>
        /// Get adjacent key for typo simulation
        /// </summary>
        private static char GetAdjacentKey(char c)
        {
            var lower = char.ToLower(c);
            if (_keyboardAdjacent.TryGetValue(lower, out var adjacent) && adjacent.Length > 0)
            {
                var result = adjacent[_random.Next(adjacent.Length)];
                return char.IsUpper(c) ? char.ToUpper(result) : result;
            }
            // Return a random nearby letter
            return (char)('a' + _random.Next(26));
        }

        /// <summary>
        /// Apply fatigue multiplier to delay
        /// </summary>
        private static int ApplyFatigue(int baseDelayMs)
        {
            if (!_humanMode.FatigueEnabled) return baseDelayMs;
            
            var sessionMinutes = (DateTime.Now - _sessionStart).TotalMinutes;
            var fatigueFactor = 1.0 + (sessionMinutes / 60.0) * _humanMode.FatigueMultiplier;
            return (int)(baseDelayMs * fatigueFactor);
        }

    // ==================== v4.0 TOKEN-EFFICIENT COMMANDS ====================

    /// <summary>
    /// Checks if an element exists. Returns only boolean for minimal token usage (~30 tokens).
    /// </summary>
    private static async Task HandleElementExists(JObject command)
    {
        var selector = command["selector"]?.ToString();
        var windowSelector = command["window"]?.ToString();
        var timeout = command["timeout"]?.Value<int>() ?? 0; // 0 = immediate check

        if (string.IsNullOrEmpty(selector))
        {
            WriteMissingParam("selector");
            return;
        }

        var searchRoot = GetSearchScope(windowSelector, selector);
        if (searchRoot == null)
        {
            WriteSuccess("element_exists", new { exists = false });
            return;
        }

        AutomationElement? element = null;

        if (timeout > 0)
        {
            // Wait for element with timeout
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeout)
            {
                element = FindElementInScope(searchRoot, selector);
                if (element != null) break;
                await Task.Delay(100);
            }
        }
        else
        {
            element = FindElementInScope(searchRoot, selector);
        }

        WriteSuccess("element_exists", new { exists = element != null });
    }

    /// <summary>
    /// Gets only interactive (clickable/typeable) elements. Reduces response from ~2000 to ~500 tokens.
    /// </summary>
    private static void HandleGetInteractiveElements(JObject command)
    {
        var windowSelector = command["window"]?.ToString();
        var maxCount = command["max_count"]?.Value<int>() ?? 50;

        if (_automation == null)
        {
            WriteError("NOT_INITIALIZED", "Automation not initialized");
            return;
        }

        AutomationElement searchRoot;
        if (!string.IsNullOrEmpty(windowSelector))
        {
            var window = FindWindow(windowSelector);
            if (window == null)
            {
                WriteWindowNotFound(windowSelector);
                return;
            }
            searchRoot = window;
        }
        else
        {
            searchRoot = _automation.GetDesktop();
        }

        try
        {
            var interactiveTypes = new HashSet<FlaUI.Core.Definitions.ControlType>
            {
                FlaUI.Core.Definitions.ControlType.Button,
                FlaUI.Core.Definitions.ControlType.Edit,
                FlaUI.Core.Definitions.ControlType.ComboBox,
                FlaUI.Core.Definitions.ControlType.CheckBox,
                FlaUI.Core.Definitions.ControlType.RadioButton,
                FlaUI.Core.Definitions.ControlType.Hyperlink,
                FlaUI.Core.Definitions.ControlType.MenuItem,
                FlaUI.Core.Definitions.ControlType.Tab,
                FlaUI.Core.Definitions.ControlType.TabItem,
                FlaUI.Core.Definitions.ControlType.ListItem,
                FlaUI.Core.Definitions.ControlType.TreeItem,
                FlaUI.Core.Definitions.ControlType.Slider
            };

            var allElements = searchRoot.FindAllDescendants();
            var interactiveElements = allElements
                .Where(e => interactiveTypes.Contains(e.ControlType) && SafeGet(() => e.IsEnabled, false))
                .Take(maxCount)
                .Select(e =>
                {
                    var rect = e.BoundingRectangle;
                    return new
                    {
                        name = SafeGet(() => e.Name, ""),
                        type = e.ControlType.ToString(),
                        id = SafeGet(() => e.AutomationId, ""),
                        x = (int)rect.X,
                        y = (int)rect.Y,
                        w = (int)rect.Width,
                        h = (int)rect.Height
                    };
                })
                .ToList();

            WriteSuccess("get_interactive_elements", new { count = interactiveElements.Count, elements = interactiveElements });
        }
        catch (Exception ex)
        {
            WriteError("QUERY_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Gets a structured window summary. Provides key info without full tree dump (~300 tokens).
    /// </summary>
    private static void HandleGetWindowSummary(JObject command)
    {
        var selector = command["selector"]?.ToString();

        if (string.IsNullOrEmpty(selector))
        {
            WriteMissingParam("selector");
            return;
        }

        var window = FindWindow(selector);
        if (window == null)
        {
            WriteWindowNotFound(selector);
            return;
        }

        try
        {
            var rect = window.BoundingRectangle;
            var allElements = window.FindAllDescendants();

            // Count by type
            var typeCounts = allElements
                .GroupBy(e => e.ControlType.ToString())
                .ToDictionary(g => g.Key, g => g.Count());

            // Get key interactive elements (first 10 buttons/links/inputs)
            var keyElements = allElements
                .Where(e =>
                {
                    var ct = e.ControlType;
                    return ct == FlaUI.Core.Definitions.ControlType.Button ||
                           ct == FlaUI.Core.Definitions.ControlType.Edit ||
                           ct == FlaUI.Core.Definitions.ControlType.Hyperlink;
                })
                .Take(10)
                .Select(e => new
                {
                    name = SafeGet(() => e.Name, ""),
                    type = e.ControlType.ToString().Replace("ControlType.", ""),
                    id = SafeGet(() => e.AutomationId, "")
                })
                .Where(e => !string.IsNullOrEmpty(e.name) || !string.IsNullOrEmpty(e.id))
                .ToList();

            var summary = new
            {
                title = SafeGet(() => window.Title, ""),
                bounds = new { x = (int)rect.X, y = (int)rect.Y, w = (int)rect.Width, h = (int)rect.Height },
                element_count = allElements.Length,
                types = typeCounts,
                key_elements = keyElements
            };

            WriteSuccess("get_window_summary", new { data = summary });
        }
        catch (Exception ex)
        {
            WriteError("SUMMARY_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Describes an element at given bounds using vision (OCR + detection).
    /// Useful for unknown/canvas elements (~100 tokens).
    /// </summary>
    private static async Task HandleDescribeElement(JObject command)
    {
        var x = command["x"]?.Value<int>();
        var y = command["y"]?.Value<int>();
        var width = command["width"]?.Value<int>() ?? 100;
        var height = command["height"]?.Value<int>() ?? 50;

        if (!x.HasValue || !y.HasValue)
        {
            WriteMissingParam("x' or 'y");
            return;
        }

        try
        {
            // Call vision service to analyze region
            var regionRequest = new JObject
            {
                ["action"] = "describe_region",
                ["x"] = x.Value,
                ["y"] = y.Value,
                ["width"] = width,
                ["height"] = height
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/analyze",
                new StringContent(regionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                var parsed = JObject.Parse(result);
                
                // Extract just the essential info
                var fullText = parsed["full_text"]?.ToString() ?? "";
                // Truncate to first 100 chars for token efficiency
                var truncatedText = fullText.Length > 100 ? fullText.Substring(0, 100) + "..." : fullText;
                
                // Get element type from first detected element
                var elements = parsed["elements"] as JArray;
                var elementType = elements?.Count > 0 ? elements[0]["type"]?.ToString() ?? "unknown" : "unknown";
                
                WriteSuccess("describe_element", new
                {
                    x = x.Value,
                    y = y.Value,
                    text = truncatedText,
                    element_type = elementType,
                    element_count = parsed["element_count"]?.Value<int>() ?? 0
                });
            }
            else
            {
                // Fallback to basic OCR
                WriteSuccess("describe_element", new
                {
                    x = x.Value,
                    y = y.Value,
                    text = "",
                    element_type = "unknown",
                    note = "Vision service unavailable"
                });
            }
        }
        catch (Exception ex)
        {
            WriteError("DESCRIBE_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Gets minimal element info - just name, ID, and bounds (~100 tokens).
    /// </summary>
    private static void HandleGetElementBrief(JObject command)
    {
        var selector = command["selector"]?.ToString();
        var windowSelector = command["window"]?.ToString();

        if (string.IsNullOrEmpty(selector))
        {
            WriteMissingParam("selector");
            return;
        }

        var searchRoot = GetSearchScope(windowSelector, selector);
        if (searchRoot == null) return;

        var element = FindElementInScope(searchRoot, selector);
        if (element == null)
        {
            WriteElementNotFound(selector);
            return;
        }

        var rect = element.BoundingRectangle;
        WriteSuccess("get_element_brief", new
        {
            name = SafeGet(() => element.Name, ""),
            id = SafeGet(() => element.AutomationId, ""),
            type = element.ControlType.ToString(),
            x = (int)rect.X,
            y = (int)rect.Y,
            w = (int)rect.Width,
            h = (int)rect.Height
        });
    }

    /// <summary>
    /// Finds elements using fuzzy text matching. Works with partial names and typos.
    /// </summary>
    private static async Task HandleFuzzyFindElement(JObject command)
    {
        var text = command["text"]?.ToString();
        var windowSelector = command["window"]?.ToString();
        var threshold = command["threshold"]?.Value<double>() ?? 0.6;
        var maxResults = command["max_results"]?.Value<int>() ?? 5;

        if (string.IsNullOrEmpty(text))
        {
            WriteMissingParam("text");
            return;
        }

        AutomationElement searchRoot;
        if (!string.IsNullOrEmpty(windowSelector))
        {
            var window = FindWindow(windowSelector);
            if (window == null)
            {
                WriteWindowNotFound(windowSelector);
                return;
            }
            searchRoot = window;
        }
        else if (_automation != null)
        {
            searchRoot = _automation.GetDesktop();
        }
        else
        {
            WriteError("NOT_INITIALIZED", "Automation not initialized");
            return;
        }

        try
        {
            var allElements = searchRoot.FindAllDescendants();
            var searchLower = text.ToLower();

            var matches = allElements
                .Select(e =>
                {
                    var name = SafeGet(() => e.Name, "") ?? "";
                    var autoId = SafeGet(() => e.AutomationId, "") ?? "";
                    
                    // Calculate similarity score
                    var nameScore = CalculateSimilarity(name.ToLower(), searchLower);
                    var idScore = CalculateSimilarity(autoId.ToLower(), searchLower);
                    var score = Math.Max(nameScore, idScore);

                    return new { element = e, name, autoId, score };
                })
                .Where(m => m.score >= threshold)
                .OrderByDescending(m => m.score)
                .Take(maxResults)
                .Select(m =>
                {
                    var rect = m.element.BoundingRectangle;
                    return new
                    {
                        name = m.name,
                        id = m.autoId,
                        type = m.element.ControlType.ToString(),
                        score = Math.Round(m.score, 2),
                        x = (int)rect.X,
                        y = (int)rect.Y
                    };
                })
                .ToList();

            WriteSuccess("fuzzy_find_element", new { count = matches.Count, matches });
        }
        catch (Exception ex)
        {
            WriteError("FUZZY_FIND_FAILED", ex.Message);
        }

        await Task.CompletedTask;
    }

    /// <summary>
    /// Clicks on text using vision (for Flutter/Electron/Canvas apps).
    /// Delegates to Python vision service.
    /// </summary>
    private static async Task HandleVisionClick(JObject command)
    {
        var text = command["text"]?.ToString();
        var caseSensitive = command["case_sensitive"]?.Value<bool>() ?? false;

        if (string.IsNullOrEmpty(text))
        {
            WriteMissingParam("text");
            return;
        }

        try
        {
            var visionRequest = new JObject
            {
                ["text"] = text,
                ["case_sensitive"] = caseSensitive
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/click_text",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result); // Pass through the JSON response
            }
            else
            {
                WriteError("VISION_CLICK_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_CLICK_FAILED", ex.Message);
        }
    }

    // ==================== AGENT VISION COMMANDS (v4.3) ====================
    // These commands proxy to Python bridge agent vision endpoints
    // for AI agents that perform their own vision analysis

    /// <summary>
    /// Takes an optimized screenshot for agent vision mode.
    /// Proxies to Python bridge /vision/screenshot_cached endpoint.
    /// </summary>
    private static async Task HandleVisionScreenshot(JObject command)
    {
        try
        {
            var visionRequest = new JObject
            {
                ["use_cache"] = command["use_cache"]?.Value<bool>() ?? true,
                ["jpeg_quality"] = command["jpeg_quality"]?.Value<int>() ?? 75,
                ["max_width"] = command["max_width"]?.Value<int>() ?? 1920,
                ["max_height"] = command["max_height"]?.Value<int>() ?? 1080,
                ["include_thumbnail"] = command["include_thumbnail"]?.Value<bool>() ?? false
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/screenshot_cached",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result); // Pass through the JSON response
            }
            else
            {
                WriteError("VISION_SCREENSHOT_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_SCREENSHOT_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Takes an optimized screenshot of a specific screen region.
    /// Proxies to Python bridge /vision/screenshot_region endpoint.
    /// </summary>
    private static async Task HandleVisionScreenshotRegion(JObject command)
    {
        var x = command["x"]?.Value<int>();
        var y = command["y"]?.Value<int>();
        var width = command["width"]?.Value<int>();
        var height = command["height"]?.Value<int>();

        if (x == null || y == null || width == null || height == null)
        {
            WriteMissingParam("x, y, width, height");
            return;
        }

        try
        {
            var visionRequest = new JObject
            {
                ["x"] = x,
                ["y"] = y,
                ["width"] = width,
                ["height"] = height,
                ["jpeg_quality"] = command["jpeg_quality"]?.Value<int>() ?? 75
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/screenshot_region",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result); // Pass through the JSON response
            }
            else
            {
                WriteError("VISION_SCREENSHOT_REGION_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_SCREENSHOT_REGION_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Gets or sets vision configuration.
    /// Proxies to Python bridge /vision/config endpoint.
    /// </summary>
    private static async Task HandleVisionConfig(JObject command)
    {
        try
        {
            // Check if this is a GET (no config params) or POST (with config params)
            var mode = command["mode"]?.ToString();
            var jpegQuality = command["jpeg_quality"]?.Value<int>();
            var maxWidth = command["max_width"]?.Value<int>();
            var maxHeight = command["max_height"]?.Value<int>();

            if (mode == null && jpegQuality == null && maxWidth == null && maxHeight == null)
            {
                // GET - retrieve current config
                var response = await _httpClient.GetAsync($"{PythonBridgeUrl}/vision/config");
                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(result);
                }
                else
                {
                    WriteError("VISION_CONFIG_FAILED", $"Vision service returned: {response.StatusCode}");
                }
            }
            else
            {
                // POST - update config
                var visionRequest = new JObject();
                if (mode != null) visionRequest["mode"] = mode;
                if (jpegQuality != null) visionRequest["jpeg_quality"] = jpegQuality;
                if (maxWidth != null) visionRequest["max_width"] = maxWidth;
                if (maxHeight != null) visionRequest["max_height"] = maxHeight;

                var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/config",
                    new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(result);
                }
                else
                {
                    WriteError("VISION_CONFIG_FAILED", $"Vision service returned: {response.StatusCode}");
                }
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_CONFIG_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Smart analysis endpoint that respects vision mode configuration.
    /// Proxies to Python bridge /vision/analyze_or_screenshot endpoint.
    /// </summary>
    private static async Task HandleVisionAnalyzeSmart(JObject command)
    {
        try
        {
            var visionRequest = new JObject
            {
                ["use_cache"] = command["use_cache"]?.Value<bool>() ?? true,
                ["jpeg_quality"] = command["jpeg_quality"]?.Value<int>() ?? 75,
                ["force_screenshot"] = command["force_screenshot"]?.Value<bool>() ?? false
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/analyze_or_screenshot",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result);
            }
            else
            {
                WriteError("VISION_ANALYZE_SMART_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_ANALYZE_SMART_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Gets screenshot cache statistics.
    /// Proxies to Python bridge /vision/screenshot_cache/stats endpoint.
    /// </summary>
    private static async Task HandleVisionScreenshotCacheStats()
    {
        try
        {
            var response = await _httpClient.GetAsync($"{PythonBridgeUrl}/vision/screenshot_cache/stats");
            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result);
            }
            else
            {
                WriteError("VISION_SCREENSHOT_CACHE_STATS_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_SCREENSHOT_CACHE_STATS_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Clears the screenshot cache.
    /// Proxies to Python bridge /vision/screenshot_cache/clear endpoint.
    /// </summary>
    private static async Task HandleVisionScreenshotCacheClear()
    {
        try
        {
            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/screenshot_cache/clear",
                new StringContent("{}", Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                Console.WriteLine(result);
            }
            else
            {
                WriteError("VISION_SCREENSHOT_CACHE_CLEAR_FAILED", $"Vision service returned: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            WriteError("VISION_SCREENSHOT_CACHE_CLEAR_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Returns WebSocket URL for real-time screenshot streaming.
    /// </summary>
    private static void HandleVisionStream(JsonElement command)
    {
        try
        {
            // Extract parameters with defaults
            int fps = GetIntOrDefault(command, "fps", 5);
            int quality = GetIntOrDefault(command, "quality", 70);
            int maxWidth = GetIntOrDefault(command, "max_width", 1920);
            int maxHeight = GetIntOrDefault(command, "max_height", 1080);

            var wsUrl = $"ws://localhost:5001/vision/stream?fps={fps}&quality={quality}&max_width={maxWidth}&max_height={maxHeight}";

            Console.WriteLine(JsonSerializer.Serialize(new
            {
                status = "success",
                websocket_url = wsUrl,
                fps,
                quality,
                max_width = maxWidth,
                max_height = maxHeight,
                instructions = "Connect to WebSocket URL to receive JPEG frames as base64."
            }, JsonOptions));
        }
        catch (Exception ex)
        {
            WriteError("VISION_STREAM_FAILED", ex.Message);
        }
    }

    private static int GetIntOrDefault(JsonElement command, string property, int defaultValue)
    {
        return command.TryGetProperty(property, out var element) ? element.GetInt32() : defaultValue;
    }

    // ==================== SMART FALLBACK CHAIN (v4.0) ====================

    /// <summary>
    /// Result of a find operation with fallback chain
    /// </summary>
    private class FindResult
    {
        public bool Found { get; set; }
        public string Method { get; set; } = "";
        public AutomationElement? Element { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Tries to find an element using FlaUI first, then vision as fallback.
    /// Returns a unified result indicating which method succeeded.
    /// </summary>
    private static async Task<FindResult> TryFindWithVisionFallback(
        string selectorOrText, 
        string? windowSelector = null,
        bool useVisionFallback = true)
    {
        var result = new FindResult();
        
        // 1. Try FlaUI first (fast, reliable for standard apps)
        try
        {
            AutomationElement? searchRoot = null;
            if (!string.IsNullOrEmpty(windowSelector))
            {
                searchRoot = FindWindow(windowSelector);
            }
            searchRoot ??= _automation?.GetDesktop();

            if (searchRoot != null)
            {
                var element = FindElementInScope(searchRoot, selectorOrText);
                if (element != null)
                {
                    result.Found = true;
                    result.Method = "flaui";
                    result.Element = element;
                    var clickPoint = element.GetClickablePoint();
                    result.X = (int)clickPoint.X;
                    result.Y = (int)clickPoint.Y;
                    return result;
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[SmartFallback] FlaUI lookup failed: {ex.Message}");
        }

        // 2. Try Vision fallback (for Flutter/Electron/Canvas apps)
        if (useVisionFallback)
        {
            try
            {
                var visionRequest = new JObject
                {
                    ["text"] = selectorOrText,
                    ["case_sensitive"] = false
                };

                var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/find_text",
                    new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var visionResult = JObject.Parse(content);
                    
                    if (visionResult["status"]?.ToString() == "success" && visionResult["found"]?.Value<bool>() == true)
                    {
                        result.Found = true;
                        result.Method = "vision";
                        result.X = visionResult["x"]?.Value<int>() ?? 0;
                        result.Y = visionResult["y"]?.Value<int>() ?? 0;
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[SmartFallback] Vision lookup failed: {ex.Message}");
                result.ErrorMessage = ex.Message;
            }
        }

        result.Found = false;
        result.ErrorMessage ??= $"Element '{selectorOrText}' not found with any method";
        return result;
    }

    /// <summary>
    /// Smart click: Tries FlaUI click first, falls back to vision click.
    /// Works with both standard Win32 apps and modern Flutter/Electron apps.
    /// v4.3: Adds agent_fallback option to return screenshot when all methods fail.
    /// </summary>
    private static async Task HandleSmartClick(JObject command)
    {
        var selectorOrText = command["selector"]?.ToString() ?? command["text"]?.ToString();
        var windowSelector = command["window"]?.ToString();
        var useVisionFallback = command["vision_fallback"]?.Value<bool>() ?? true;
        var useAgentFallback = command["agent_fallback"]?.Value<bool>() ?? false;
        var jpegQuality = command["jpeg_quality"]?.Value<int>() ?? 75;

        if (string.IsNullOrEmpty(selectorOrText))
        {
            WriteMissingParam("selector or text");
            return;
        }

        var findResult = await TryFindWithVisionFallback(selectorOrText, windowSelector, useVisionFallback);

        if (!findResult.Found)
        {
            // v4.3: If agent_fallback is enabled, return screenshot instead of just error
            if (useAgentFallback)
            {
                await WriteSmartClickFailureWithScreenshot(selectorOrText, findResult.ErrorMessage, jpegQuality);
            }
            else
            {
                WriteError("SMART_CLICK_FAILED", findResult.ErrorMessage ?? "Element not found");
            }
            return;
        }

        try
        {
            if (findResult.Method == "flaui" && findResult.Element != null)
            {
                // Use FlaUI native click
                findResult.Element.Click();
            }
            else
            {
                // Use coordinate-based click from vision
                Mouse.Click(new System.Drawing.Point(findResult.X, findResult.Y));
            }

            WriteSuccess("smart_click", new
            {
                method_used = findResult.Method,
                x = findResult.X,
                y = findResult.Y
            });
        }
        catch (Exception ex)
        {
            WriteError("SMART_CLICK_FAILED", ex.Message);
        }
    }

    /// <summary>
    /// Smart type: Finds element using smart fallback, then types text.
    /// Works with both standard Win32 apps and modern Flutter/Electron apps.
    /// v4.3: Adds agent_fallback option to return screenshot when all methods fail.
    /// </summary>
    private static async Task HandleSmartType(JObject command)
    {
        var selectorOrText = command["selector"]?.ToString() ?? command["text"]?.ToString();
        var windowSelector = command["window"]?.ToString();
        var textToType = command["input"]?.ToString() ?? command["value"]?.ToString();
        var useVisionFallback = command["vision_fallback"]?.Value<bool>() ?? true;
        var useAgentFallback = command["agent_fallback"]?.Value<bool>() ?? false;
        var jpegQuality = command["jpeg_quality"]?.Value<int>() ?? 75;
        var clearFirst = command["clear"]?.Value<bool>() ?? false;

        if (string.IsNullOrEmpty(selectorOrText))
        {
            WriteMissingParam("selector or text");
            return;
        }

        if (string.IsNullOrEmpty(textToType))
        {
            WriteMissingParam("input or value");
            return;
        }

        var findResult = await TryFindWithVisionFallback(selectorOrText, windowSelector, useVisionFallback);

        if (!findResult.Found)
        {
            // v4.3: If agent_fallback is enabled, return screenshot instead of just error
            if (useAgentFallback)
            {
                await WriteSmartTypeFailureWithScreenshot(selectorOrText, textToType, findResult.ErrorMessage, jpegQuality);
            }
            else
            {
                WriteError("SMART_TYPE_FAILED", findResult.ErrorMessage ?? "Element not found");
            }
            return;
        }

        try
        {
            if (findResult.Method == "flaui" && findResult.Element != null)
            {
                // Use FlaUI native focus + type
                findResult.Element.Focus();
                
                if (clearFirst)
                {
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    await Task.Delay(50);
                }
            }
            else
            {
                // Click to focus via coordinates, then type
                Mouse.Click(new System.Drawing.Point(findResult.X, findResult.Y));
                await Task.Delay(100); // Wait for focus
                
                if (clearFirst)
                {
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    await Task.Delay(50);
                }
            }

            Keyboard.Type(textToType);

            WriteSuccess("smart_type", new
            {
                method_used = findResult.Method,
                typed = textToType.Length,
                x = findResult.X,
                y = findResult.Y
            });
        }
        catch (Exception ex)
        {
            WriteError("SMART_TYPE_FAILED", ex.Message);
        }
    }

    // ==================== AGENT FALLBACK HELPERS (v4.3) ====================

    /// <summary>
    /// Writes a smart_click failure response with an optimized screenshot for agent vision.
    /// The AI agent can analyze the screenshot and suggest click coordinates.
    /// </summary>
    private static async Task WriteSmartClickFailureWithScreenshot(string searchText, string? errorMessage, int jpegQuality)
    {
        try
        {
            var visionRequest = new JObject
            {
                ["use_cache"] = true,
                ["jpeg_quality"] = jpegQuality
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/screenshot_cached",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                var screenshotResult = JObject.Parse(content);

                // Build response with screenshot for agent analysis
                var result = new JObject
                {
                    ["status"] = "agent_fallback",
                    ["action"] = "smart_click",
                    ["code"] = "ELEMENT_NOT_FOUND",
                    ["message"] = errorMessage ?? $"Element '{searchText}' not found with any method",
                    ["search_text"] = searchText,
                    ["suggestion"] = "Analyze the screenshot and provide click coordinates using click_at command",
                    ["screenshot"] = screenshotResult["screenshot"]
                };

                Console.WriteLine(result.ToString(Formatting.None));
            }
            else
            {
                // Fallback: return error without screenshot
                WriteError("SMART_CLICK_FAILED", errorMessage ?? $"Element '{searchText}' not found");
            }
        }
        catch (Exception ex)
        {
            // Fallback: return error without screenshot
            Console.Error.WriteLine($"[AgentFallback] Screenshot capture failed: {ex.Message}");
            WriteError("SMART_CLICK_FAILED", errorMessage ?? $"Element '{searchText}' not found");
        }
    }

    /// <summary>
    /// Writes a smart_type failure response with an optimized screenshot for agent vision.
    /// The AI agent can analyze the screenshot and suggest input field coordinates.
    /// </summary>
    private static async Task WriteSmartTypeFailureWithScreenshot(string searchText, string textToType, string? errorMessage, int jpegQuality)
    {
        try
        {
            var visionRequest = new JObject
            {
                ["use_cache"] = true,
                ["jpeg_quality"] = jpegQuality
            };

            var response = await _httpClient.PostAsync($"{PythonBridgeUrl}/vision/screenshot_cached",
                new StringContent(visionRequest.ToString(), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                var screenshotResult = JObject.Parse(content);

                // Build response with screenshot for agent analysis
                var result = new JObject
                {
                    ["status"] = "agent_fallback",
                    ["action"] = "smart_type",
                    ["code"] = "ELEMENT_NOT_FOUND",
                    ["message"] = errorMessage ?? $"Element '{searchText}' not found with any method",
                    ["search_text"] = searchText,
                    ["pending_text"] = textToType,
                    ["suggestion"] = "Analyze the screenshot to find the input field, then use click_at to focus it and type_here to enter text",
                    ["screenshot"] = screenshotResult["screenshot"]
                };

                Console.WriteLine(result.ToString(Formatting.None));
            }
            else
            {
                // Fallback: return error without screenshot
                WriteError("SMART_TYPE_FAILED", errorMessage ?? $"Element '{searchText}' not found");
            }
        }
        catch (Exception ex)
        {
            // Fallback: return error without screenshot
            Console.Error.WriteLine($"[AgentFallback] Screenshot capture failed: {ex.Message}");
            WriteError("SMART_TYPE_FAILED", errorMessage ?? $"Element '{searchText}' not found");
        }
    }

    /// <summary>
    /// Helper: Calculate string similarity (Levenshtein-based)
    /// </summary>
    private static double CalculateSimilarity(string s1, string s2)
    {
        if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2)) return 0.0;
        if (s1 == s2) return 1.0;

        // Contains check for quick partial match
        if (s1.Contains(s2) || s2.Contains(s1))
        {
            return 0.8 + 0.2 * ((double)Math.Min(s1.Length, s2.Length) / Math.Max(s1.Length, s2.Length));
        }

        // Levenshtein distance
        var distance = LevenshteinDistance(s1, s2);
        var maxLen = Math.Max(s1.Length, s2.Length);
        return 1.0 - ((double)distance / maxLen);
    }

    private static int LevenshteinDistance(string s1, string s2)
    {
        var m = s1.Length;
        var n = s2.Length;
        var dp = new int[m + 1, n + 1];

        for (var i = 0; i <= m; i++) dp[i, 0] = i;
        for (var j = 0; j <= n; j++) dp[0, j] = j;

        for (var i = 1; i <= m; i++)
        {
            for (var j = 1; j <= n; j++)
            {
                var cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                dp[i, j] = Math.Min(Math.Min(
                    dp[i - 1, j] + 1,      // deletion
                    dp[i, j - 1] + 1),     // insertion
                    dp[i - 1, j - 1] + cost); // substitution
            }
        }

        return dp[m, n];
    }

    // ==================== SUPPORTING CLASSES ====================

    /// <summary>
    /// Metrics for tracking command performance
    /// </summary>
    public class CommandMetrics
    {
        public int CallCount { get; set; }
        public long TotalMs { get; set; }
        public long MinMs { get; set; } = long.MaxValue;
        public long MaxMs { get; set; }
        public DateTime LastCalled { get; set; }
    }

    /// <summary>
    /// Event subscription info
    /// </summary>
    public class EventSubscription
    {
        public string Id { get; set; } = "";
        public string Selector { get; set; } = "";
        public string EventType { get; set; } = "";
        public DateTime CreatedAt { get; set; }
    }

    /// <summary>
    /// Human-like behavior configuration
    /// </summary>
    public class HumanModeSettings
    {
        public bool Enabled { get; set; } = false;
        public int MinPauseMs { get; set; } = 50;
        public int MaxPauseMs { get; set; } = 300;
        public double MouseJitter { get; set; } = 2.0;
        public double TypingErrorRate { get; set; } = 0.03;
        public int TypingMinDelayMs { get; set; } = 50;
        public int TypingMaxDelayMs { get; set; } = 150;
        public bool FatigueEnabled { get; set; } = false;
        public double FatigueMultiplier { get; set; } = 0.5; // 50% slower per hour
        public bool ThinkingDelayEnabled { get; set; } = true;
        public int ThinkingMinMs { get; set; } = 200;
        public int ThinkingMaxMs { get; set; } = 800;
    }
}
}
