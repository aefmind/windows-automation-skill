"""
Windows Desktop Automation - Python Bridge
Fallback automation engine using pywinauto for legacy Win32 applications.
Also provides vision-based automation using OmniParser + Windows OCR for Flutter/Electron apps.
Runs on port 5001.
"""

import os
import sys

# Ensure the bridge_python directory is in the path for vision imports
# We use __file__ which should resolve to the full path of bridge.py
_this_file = __file__
if hasattr(_this_file, "__fspath__"):
    _this_file = _this_file.__fspath__()
_bridge_dir = os.path.dirname(os.path.abspath(_this_file))
print(f"Bridge startup: Adding {_bridge_dir} to sys.path")
if _bridge_dir not in sys.path:
    sys.path.insert(0, _bridge_dir)

from flask import Flask, request, jsonify
from flask_sock import Sock
import pywinauto
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import logging
import time
from PIL import ImageGrab
import io
import base64

# Vision imports (lazy loaded to avoid startup delay)
_vision_service = None

# Context Manager for caching and compression
from context_manager import get_context_manager


def get_vision_service():
    """Lazy load VisionService to avoid slow startup"""
    global _vision_service
    if _vision_service is None:
        # Ensure vision module is importable
        import os
        import sys

        # Get the directory containing this file
        this_file = os.path.abspath(__file__)
        bridge_dir = os.path.dirname(this_file)

        # Debug logging
        import logging

        logging.info(f"Vision import: __file__={this_file}")
        logging.info(f"Vision import: bridge_dir={bridge_dir}")
        logging.info(
            f"Vision import: sys.path[0]={sys.path[0] if sys.path else 'empty'}"
        )

        # Add to path if not present
        if bridge_dir not in sys.path:
            sys.path.insert(0, bridge_dir)
            logging.info(f"Vision import: Added {bridge_dir} to sys.path")

        # Check if vision folder exists
        vision_dir = os.path.join(bridge_dir, "vision")
        if os.path.isdir(vision_dir):
            logging.info(f"Vision import: vision folder exists at {vision_dir}")
        else:
            logging.error(f"Vision import: vision folder NOT FOUND at {vision_dir}")

        from vision.vision_service import VisionService

        _vision_service = VisionService(confidence_threshold=0.10)
    return _vision_service


app = Flask(__name__)
sock = Sock(app)
logging.basicConfig(level=logging.INFO)

# ==================== RESPONSE HELPERS ====================


def success_response(action=None, **kwargs):
    """Returns a standardized success response."""
    response = {"status": "success", "engine": "pywinauto"}
    if action:
        response["action"] = action
    response.update(kwargs)
    return jsonify(response)


def error_response(code, message, status_code=500):
    """Returns a standardized error response."""
    return jsonify(
        {"status": "error", "code": code, "message": message, "engine": "pywinauto"}
    ), status_code


# ==================== HELPER FUNCTIONS ====================


def get_target(data):
    """Gets target window from request data."""
    title = data.get("title")
    selector = data.get("selector")
    backend = data.get("backend", "win32")

    desktop = Desktop(backend=backend)
    if title:
        return desktop.window(title_re=title)
    if selector:
        return desktop.window(title_re=f".*{selector}.*")
    return desktop.active()


def safe_get_window_info(window):
    """Safely extracts window information."""
    try:
        return {
            "title": window.window_text(),
            "class_name": window.class_name(),
            "handle": window.handle,
            "rect": list(window.rectangle()) if hasattr(window, "rectangle") else None,
        }
    except Exception:
        return {"title": "Unknown", "class_name": "", "handle": 0, "rect": None}


# ==================== DISCOVERY ENDPOINTS ====================


@app.route("/explore", methods=["GET"])
def explore():
    """Lists all visible windows. Uses context manager caching."""
    try:
        ctx = get_context_manager()
        params = {}

        # Try cache first
        hit, cached_response = ctx.get_cached("explore", params)
        if hit:
            cached_response["cache_hit"] = True
            return jsonify(cached_response)

        # Execute actual query
        windows = Desktop(backend="win32").windows()
        result = [safe_get_window_info(w) for w in windows]

        # Build response dict
        response_dict = {
            "status": "success",
            "engine": "pywinauto",
            "action": "explore",
            "count": len(result),
            "data": result,
        }

        # Cache the response
        ctx.process_response("explore", params, response_dict, compress=False)

        return jsonify(response_dict)
    except Exception as e:
        return error_response("EXPLORE_FAILED", str(e))


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return success_response("health", message="Python bridge is running")


# ==================== MOUSE ENDPOINTS ====================


@app.route("/click", methods=["POST"])
def click():
    """Single click on target element."""
    try:
        target = get_target(request.json)
        target.set_focus()
        target.click_input()
        return success_response("click")
    except Exception as e:
        return error_response("CLICK_FAILED", str(e))


@app.route("/double_click", methods=["POST"])
def double_click():
    """Double click on target element."""
    try:
        target = get_target(request.json)
        target.set_focus()
        target.double_click_input()
        return success_response("double_click")
    except Exception as e:
        return error_response("DOUBLE_CLICK_FAILED", str(e))


@app.route("/right_click", methods=["POST"])
def right_click():
    """Right click on target element."""
    try:
        target = get_target(request.json)
        target.set_focus()
        target.right_click_input()
        return success_response("right_click")
    except Exception as e:
        return error_response("RIGHT_CLICK_FAILED", str(e))


@app.route("/click_at", methods=["POST"])
def click_at():
    """Click at specific coordinates."""
    try:
        data = request.json
        x = data.get("x", 0)
        y = data.get("y", 0)
        button = data.get("button", "left")

        if button == "right":
            pywinauto.mouse.right_click(coords=(x, y))
        elif button == "middle":
            pywinauto.mouse.middle_click(coords=(x, y))
        else:
            pywinauto.mouse.click(coords=(x, y))

        return success_response("click_at", x=x, y=y, button=button)
    except Exception as e:
        return error_response("CLICK_AT_FAILED", str(e))


@app.route("/mouse_move", methods=["POST"])
def mouse_move():
    """Move mouse to coordinates."""
    try:
        data = request.json
        x = data.get("x", 0)
        y = data.get("y", 0)
        pywinauto.mouse.move(coords=(x, y))
        return success_response("mouse_move", x=x, y=y)
    except Exception as e:
        return error_response("MOUSE_MOVE_FAILED", str(e))


@app.route("/scroll", methods=["POST"])
def scroll():
    """Scroll mouse wheel."""
    try:
        data = request.json
        direction = data.get("direction", "down")
        amount = data.get("amount", 3)
        wheel_dist = amount if direction == "up" else -amount
        pywinauto.mouse.scroll(coords=(0, 0), wheel_dist=wheel_dist)
        return success_response("scroll", direction=direction, amount=amount)
    except Exception as e:
        return error_response("SCROLL_FAILED", str(e))


@app.route("/drag_and_drop", methods=["POST"])
def drag_and_drop():
    """Drag from one point to another."""
    try:
        data = request.json
        from_x = data.get("from_x", 0)
        from_y = data.get("from_y", 0)
        to_x = data.get("to_x", 0)
        to_y = data.get("to_y", 0)

        pywinauto.mouse.press(coords=(from_x, from_y))
        time.sleep(0.1)
        pywinauto.mouse.move(coords=(to_x, to_y))
        time.sleep(0.1)
        pywinauto.mouse.release(coords=(to_x, to_y))

        return success_response(
            "drag_and_drop",
            from_point={"x": from_x, "y": from_y},
            to_point={"x": to_x, "y": to_y},
        )
    except Exception as e:
        return error_response("DRAG_FAILED", str(e))


# ==================== KEYBOARD ENDPOINTS ====================


@app.route("/type", methods=["POST"])
def type_text():
    """Type text into target element."""
    try:
        data = request.json
        target = get_target(data)
        text = data.get("text", "")
        target.set_focus()
        target.type_keys(text, with_spaces=True)
        return success_response("type")
    except Exception as e:
        return error_response("TYPE_FAILED", str(e))


@app.route("/hotkey", methods=["POST"])
def hotkey():
    """Send keyboard shortcut."""
    try:
        data = request.json
        keys = data.get("keys", "")

        # Convert format: "ctrl+s" -> "^s", "alt+f4" -> "%{F4}"
        key_map = {"ctrl": "^", "alt": "%", "shift": "+", "win": "#"}

        parts = keys.lower().split("+")
        pywinauto_keys = ""

        for part in parts:
            if part in key_map:
                pywinauto_keys += key_map[part]
            elif part.startswith("f") and part[1:].isdigit():
                pywinauto_keys += "{" + part.upper() + "}"
            elif len(part) == 1:
                pywinauto_keys += part
            else:
                pywinauto_keys += "{" + part.upper() + "}"

        send_keys(pywinauto_keys)
        return success_response("hotkey", keys=keys)
    except Exception as e:
        return error_response("HOTKEY_FAILED", str(e))


@app.route("/key_press", methods=["POST"])
def key_press():
    """Press a single key."""
    try:
        data = request.json
        key = data.get("key", "")

        # Special key mapping
        special_keys = {
            "enter": "{ENTER}",
            "tab": "{TAB}",
            "escape": "{ESC}",
            "backspace": "{BACKSPACE}",
            "delete": "{DELETE}",
            "space": " ",
            "up": "{UP}",
            "down": "{DOWN}",
            "left": "{LEFT}",
            "right": "{RIGHT}",
            "home": "{HOME}",
            "end": "{END}",
            "pageup": "{PGUP}",
            "pagedown": "{PGDN}",
        }

        key_to_send = special_keys.get(key.lower(), key)
        send_keys(key_to_send)
        return success_response("key_press", key=key)
    except Exception as e:
        return error_response("KEY_PRESS_FAILED", str(e))


# ==================== WINDOW ENDPOINTS ====================


@app.route("/focus_window", methods=["POST"])
def focus_window():
    """Bring window to foreground."""
    try:
        target = get_target(request.json)
        target.set_focus()
        return success_response("focus_window", window=target.window_text())
    except Exception as e:
        return error_response("FOCUS_FAILED", str(e))


@app.route("/close_window", methods=["POST"])
def close_window():
    """Close a window."""
    try:
        target = get_target(request.json)
        title = target.window_text()
        target.close()
        return success_response("close_window", window=title)
    except Exception as e:
        return error_response("CLOSE_FAILED", str(e))


# ==================== VISION ENDPOINTS (v4.0) ====================


@app.route("/vision/debug", methods=["GET"])
def vision_debug():
    """Debug endpoint to check vision import status"""
    import sys
    import os

    bridge_dir = os.path.dirname(os.path.abspath(__file__))
    vision_dir = os.path.join(bridge_dir, "vision")

    debug_info = {
        "bridge_dir": bridge_dir,
        "vision_dir_exists": os.path.isdir(vision_dir),
        "sys_path_0": sys.path[0] if sys.path else "empty",
        "bridge_in_path": bridge_dir in sys.path,
    }

    if os.path.isdir(vision_dir):
        debug_info["vision_files"] = os.listdir(vision_dir)

    try:
        from vision.vision_service import VisionService

        debug_info["import_success"] = True
        debug_info["vision_service_class"] = str(VisionService)
    except Exception as e:
        debug_info["import_success"] = False
        debug_info["import_error"] = str(e)
        import traceback

        debug_info["traceback"] = traceback.format_exc()

    return jsonify(debug_info)


@app.route("/vision/detect", methods=["POST"])
def vision_detect():
    """
    Detect all UI elements in the current screen using OmniParser.

    Request body:
        - use_tiling (bool, optional): Force tiling mode for large screens
        - confidence (float, optional): Minimum confidence threshold (0-1)

    Returns:
        - elements: List of detected UI elements with coordinates
    """
    try:
        data = request.json or {}
        use_tiling = data.get("use_tiling", False)
        confidence = data.get("confidence", 0.25)

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Detect elements
        start = time.time()
        elements = service.detect_elements(screenshot, use_tiling=use_tiling)
        elapsed = time.time() - start

        return success_response(
            "vision_detect",
            element_count=len(elements),
            elapsed_ms=int(elapsed * 1000),
            screen_size={"width": screenshot.size[0], "height": screenshot.size[1]},
            elements=[e.to_dict() for e in elements],
        )
    except Exception as e:
        logging.exception("Vision detect failed")
        return error_response("VISION_DETECT_FAILED", str(e))


@app.route("/vision/ocr", methods=["POST"])
def vision_ocr():
    """
    Perform OCR on the current screen.

    Request body:
        - full_text (bool, optional): Return full concatenated text instead of regions

    Returns:
        - text_regions: List of text regions with positions
        - full_text: All text concatenated (if requested)
    """
    try:
        data = request.json or {}
        full_text_only = data.get("full_text", False)

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Run OCR
        start = time.time()
        regions = service.recognize_text(screenshot)
        elapsed = time.time() - start

        if full_text_only:
            full_text = " ".join(r.text for r in regions)
            return success_response(
                "vision_ocr", elapsed_ms=int(elapsed * 1000), full_text=full_text
            )
        else:
            return success_response(
                "vision_ocr",
                region_count=len(regions),
                elapsed_ms=int(elapsed * 1000),
                text_regions=[r.to_dict() for r in regions],
            )
    except Exception as e:
        logging.exception("Vision OCR failed")
        return error_response("VISION_OCR_FAILED", str(e))


@app.route("/vision/find_text", methods=["POST"])
def vision_find_text():
    """
    Find a UI element containing specific text.
    Combines OmniParser detection with OCR for accurate text-based element finding.

    Request body:
        - text (str, required): Text to search for
        - case_sensitive (bool, optional): Case-sensitive matching (default: false)

    Returns:
        - found: Whether the element was found
        - element: The element details if found (x, y, width, height, confidence, text)
    """
    try:
        data = request.json or {}
        text = data.get("text")
        case_sensitive = data.get("case_sensitive", False)

        if not text:
            return error_response(
                "MISSING_PARAMETER", "text parameter is required", 400
            )

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Find element
        start = time.time()
        element = service.find_element_by_text(screenshot, text, case_sensitive)
        elapsed = time.time() - start

        if element:
            return success_response(
                "vision_find_text",
                found=True,
                elapsed_ms=int(elapsed * 1000),
                search_text=text,
                element=element.to_dict(),
            )
        else:
            return success_response(
                "vision_find_text",
                found=False,
                elapsed_ms=int(elapsed * 1000),
                search_text=text,
                element=None,
            )
    except Exception as e:
        logging.exception("Vision find text failed")
        return error_response("VISION_FIND_TEXT_FAILED", str(e))


@app.route("/vision/click_text", methods=["POST"])
def vision_click_text():
    """
    Find text on screen and click it.
    This is the key endpoint for Flutter/Electron apps where FlaUI cannot see elements.

    Request body:
        - text (str, required): Text to find and click
        - case_sensitive (bool, optional): Case-sensitive matching (default: false)
        - button (str, optional): Mouse button to use (left, right, middle)
        - double_click (bool, optional): Perform double-click instead of single click

    Returns:
        - clicked: Whether the click was performed
        - element: The element that was clicked
    """
    try:
        data = request.json or {}
        text = data.get("text")
        case_sensitive = data.get("case_sensitive", False)
        button = data.get("button", "left")
        double_click = data.get("double_click", False)

        if not text:
            return error_response(
                "MISSING_PARAMETER", "text parameter is required", 400
            )

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Find element
        start = time.time()
        element = service.find_element_by_text(screenshot, text, case_sensitive)
        find_elapsed = time.time() - start

        if not element:
            return success_response(
                "vision_click_text",
                clicked=False,
                elapsed_ms=int(find_elapsed * 1000),
                search_text=text,
                reason="Text not found on screen",
            )

        # Click on the element
        x, y = element.center

        if double_click:
            pywinauto.mouse.double_click(coords=(x, y))
        elif button == "right":
            pywinauto.mouse.right_click(coords=(x, y))
        elif button == "middle":
            pywinauto.mouse.middle_click(coords=(x, y))
        else:
            pywinauto.mouse.click(coords=(x, y))

        total_elapsed = time.time() - start

        return success_response(
            "vision_click_text",
            clicked=True,
            elapsed_ms=int(total_elapsed * 1000),
            search_text=text,
            click_coords={"x": x, "y": y},
            element=element.to_dict(),
        )
    except Exception as e:
        logging.exception("Vision click text failed")
        return error_response("VISION_CLICK_TEXT_FAILED", str(e))


@app.route("/vision/analyze", methods=["POST"])
def vision_analyze():
    """
    Perform full screen analysis - detect elements and OCR.
    Returns a comprehensive view of the current screen state.

    Returns:
        - screen_size: Width and height of the screen
        - elements: Detected UI elements
        - text_regions: OCR text regions
        - full_text: All recognized text
    """
    try:
        data = request.json or {}
        use_cache = data.get("use_cache", True)

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Full analysis
        start = time.time()
        analysis = service.analyze_screen(screenshot, use_cache=use_cache)
        elapsed = time.time() - start

        return success_response(
            "vision_analyze", elapsed_ms=int(elapsed * 1000), **analysis
        )
    except Exception as e:
        logging.exception("Vision analyze failed")
        return error_response("VISION_ANALYZE_FAILED", str(e))


# ==================== OPTIMIZED OCR ENDPOINTS (v4.1) ====================


@app.route("/vision/ocr_region", methods=["POST"])
def vision_ocr_region():
    """
    Perform OCR on a specific screen region (4x faster than full screen).

    Request body:
        - x (int, required): Left coordinate of region
        - y (int, required): Top coordinate of region
        - width (int, required): Width of region
        - height (int, required): Height of region

    Performance: ~120ms vs ~530ms for full screen

    Returns:
        - text_regions: List of text regions within the specified area
    """
    try:
        data = request.json or {}
        x = data.get("x")
        y = data.get("y")
        width = data.get("width")
        height = data.get("height")

        if x is None or y is None or width is None or height is None:
            return error_response(
                "MISSING_PARAMETER", "x, y, width, height are all required", 400
            )

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # OCR the region
        start = time.time()
        regions = service.recognize_text_region(screenshot, x, y, width, height)
        elapsed = time.time() - start

        return success_response(
            "vision_ocr_region",
            region_count=len(regions),
            elapsed_ms=int(elapsed * 1000),
            search_region={"x": x, "y": y, "width": width, "height": height},
            text_regions=[r.to_dict() for r in regions],
        )
    except Exception as e:
        logging.exception("Vision OCR region failed")
        return error_response("VISION_OCR_REGION_FAILED", str(e))


@app.route("/vision/find_text_fast", methods=["POST"])
def vision_find_text_fast():
    """
    Fast text search with optional region hints.

    If hint_regions are provided, searches those areas first (4x faster).
    Falls back to full image search if not found.

    Request body:
        - text (str, required): Text to find
        - hint_regions (list, optional): List of {x, y, width, height} to search first

    Example hint_regions for common areas:
        - Taskbar: [{"x": 0, "y": 1040, "width": 1920, "height": 40}]
        - Title bar: [{"x": 0, "y": 0, "width": 1920, "height": 50}]
        - Center dialog: [{"x": 660, "y": 340, "width": 600, "height": 400}]

    Returns:
        - found: Whether the text was found
        - region: Text region details if found
    """
    try:
        data = request.json or {}
        text = data.get("text")
        hint_regions = data.get("hint_regions", [])

        if not text:
            return error_response(
                "MISSING_PARAMETER", "text parameter is required", 400
            )

        # Convert hint_regions from dict to tuple format
        hints = None
        if hint_regions:
            hints = [(r["x"], r["y"], r["width"], r["height"]) for r in hint_regions]

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get vision service
        service = get_vision_service()

        # Fast search
        start = time.time()
        region = service.find_text_fast(screenshot, text, hints)
        elapsed = time.time() - start

        if region:
            return success_response(
                "vision_find_text_fast",
                found=True,
                elapsed_ms=int(elapsed * 1000),
                search_text=text,
                used_hints=len(hints) if hints else 0,
                region=region.to_dict(),
            )
        else:
            return success_response(
                "vision_find_text_fast",
                found=False,
                elapsed_ms=int(elapsed * 1000),
                search_text=text,
                used_hints=len(hints) if hints else 0,
                region=None,
            )
    except Exception as e:
        logging.exception("Vision find text fast failed")
        return error_response("VISION_FIND_TEXT_FAST_FAILED", str(e))


@app.route("/vision/ocr_cached", methods=["POST"])
def vision_ocr_cached():
    """
    Perform OCR with explicit cache control.

    Request body:
        - use_cache (bool, optional): Use cached results if available (default: true)
        - clear_cache (bool, optional): Clear cache before OCR (default: false)
        - full_text (bool, optional): Return full concatenated text (default: false)

    Performance:
        - First call: ~530ms
        - Cached call: <1ms

    Returns:
        - text_regions: List of text regions with positions
        - cache_stats: Current cache statistics
    """
    try:
        data = request.json or {}
        use_cache = data.get("use_cache", True)
        clear_cache = data.get("clear_cache", False)
        full_text_only = data.get("full_text", False)

        # Get vision service
        service = get_vision_service()

        # Clear cache if requested
        if clear_cache:
            service.clear_ocr_cache()

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Run OCR
        start = time.time()
        regions = service.recognize_text_cached(screenshot, use_cache=use_cache)
        elapsed = time.time() - start

        # Get cache stats
        cache_stats = service.get_ocr_cache_stats()

        if full_text_only:
            full_text = " ".join(r.text for r in regions)
            return success_response(
                "vision_ocr_cached",
                elapsed_ms=int(elapsed * 1000),
                full_text=full_text,
                cache_stats=cache_stats,
            )
        else:
            return success_response(
                "vision_ocr_cached",
                region_count=len(regions),
                elapsed_ms=int(elapsed * 1000),
                text_regions=[r.to_dict() for r in regions],
                cache_stats=cache_stats,
            )
    except Exception as e:
        logging.exception("Vision OCR cached failed")
        return error_response("VISION_OCR_CACHED_FAILED", str(e))


@app.route("/vision/cache_stats", methods=["GET"])
def vision_cache_stats():
    """
    Get OCR cache statistics.

    Returns:
        - entries: Number of cached entries
        - ttl_seconds: Cache time-to-live in seconds
    """
    try:
        service = get_vision_service()
        stats = service.get_ocr_cache_stats()
        return success_response("vision_cache_stats", **stats)
    except Exception as e:
        logging.exception("Vision cache stats failed")
        return error_response("VISION_CACHE_STATS_FAILED", str(e))


@app.route("/vision/cache_clear", methods=["POST"])
def vision_cache_clear():
    """
    Clear the OCR cache.

    Returns:
        - message: Confirmation message
    """
    try:
        service = get_vision_service()
        service.clear_ocr_cache()
        return success_response("vision_cache_clear", message="OCR cache cleared")
    except Exception as e:
        logging.exception("Vision cache clear failed")
        return error_response("VISION_CACHE_CLEAR_FAILED", str(e))


# ==================== AGENT VISION ENDPOINTS (v4.2) ====================
# These endpoints support AI agent vision mode - returning optimized screenshots
# instead of (or in addition to) local OmniParser/OCR processing


@app.route("/vision/config", methods=["GET"])
def vision_config_get():
    """
    Get current vision configuration.

    Returns:
        - mode: Current vision mode ('local', 'agent', 'auto')
        - jpeg_quality: JPEG compression quality (1-100)
        - max_width: Maximum screenshot width
        - max_height: Maximum screenshot height
        - include_thumbnail: Whether to include thumbnails
        - thumbnail_max_size: Maximum thumbnail dimension
    """
    try:
        from vision.vision_service import VisionConfig

        return success_response("vision_config", **VisionConfig.get_config())
    except Exception as e:
        logging.exception("Vision config get failed")
        return error_response("VISION_CONFIG_GET_FAILED", str(e))


@app.route("/vision/config", methods=["POST"])
def vision_config_set():
    """
    Update vision configuration.

    Request body (all optional):
        - mode (str): Vision mode ('local', 'agent', 'auto')
        - jpeg_quality (int): JPEG quality 1-100 (default 75)
        - max_width (int): Maximum screenshot width (default 1920)
        - max_height (int): Maximum screenshot height (default 1080)
        - include_thumbnail (bool): Include thumbnails in responses
        - thumbnail_max_size (int): Maximum thumbnail dimension

    Returns:
        - Updated configuration
    """
    try:
        from vision.vision_service import VisionConfig

        data = request.json or {}

        VisionConfig.update(**data)

        return success_response(
            "vision_config",
            message="Configuration updated",
            **VisionConfig.get_config(),
        )
    except ValueError as e:
        return error_response("INVALID_CONFIG", str(e), 400)
    except Exception as e:
        logging.exception("Vision config set failed")
        return error_response("VISION_CONFIG_SET_FAILED", str(e))


@app.route("/vision/screenshot", methods=["POST"])
def vision_screenshot():
    """
    Take an optimized screenshot for agent vision mode.

    This endpoint is designed for AI agents that perform their own vision
    analysis. Returns a compressed, optionally resized screenshot.

    Request body (all optional):
        - jpeg_quality (int): Override JPEG quality (1-100)
        - max_width (int): Override maximum width
        - max_height (int): Override maximum height
        - include_thumbnail (bool): Include a smaller preview

    Returns:
        - screenshot: Base64 JPEG data with metadata
            - data: Base64-encoded JPEG image
            - width, height: Dimensions after resizing
            - original_width, original_height: Original screen dimensions
            - size_bytes: Compressed size
            - compression_ratio: Compression efficiency
            - timestamp: Capture timestamp

    Performance notes:
        - 1920x1080 @ quality 75: ~150-300KB (vs ~6MB raw)
        - Token-efficient for AI vision APIs
    """
    try:
        from vision.vision_service import VisionConfig, optimize_screenshot

        data = request.json or {}

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Get parameters (use request overrides or config defaults)
        jpeg_quality = data.get("jpeg_quality", VisionConfig.jpeg_quality)
        max_width = data.get("max_width", VisionConfig.max_width)
        max_height = data.get("max_height", VisionConfig.max_height)
        include_thumbnail = data.get(
            "include_thumbnail", VisionConfig.include_thumbnail
        )
        thumbnail_max_size = data.get(
            "thumbnail_max_size", VisionConfig.thumbnail_max_size
        )

        # Optimize screenshot
        start = time.time()
        optimized = optimize_screenshot(
            screenshot,
            max_width=max_width,
            max_height=max_height,
            jpeg_quality=jpeg_quality,
            include_thumbnail=include_thumbnail,
            thumbnail_max_size=thumbnail_max_size,
        )
        elapsed = time.time() - start

        return success_response(
            "vision_screenshot",
            elapsed_ms=int(elapsed * 1000),
            screenshot=optimized.to_dict(),
        )
    except Exception as e:
        logging.exception("Vision screenshot failed")
        return error_response("VISION_SCREENSHOT_FAILED", str(e))


@app.route("/vision/screenshot_region", methods=["POST"])
def vision_screenshot_region():
    """
    Take an optimized screenshot of a specific screen region.

    More token-efficient than full screenshots when you know where to look.
    Useful for monitoring specific UI areas or following up on previous detections.

    Request body:
        - x (int, required): Left coordinate of region
        - y (int, required): Top coordinate of region
        - width (int, required): Width of region
        - height (int, required): Height of region
        - jpeg_quality (int, optional): Override JPEG quality (1-100)

    Returns:
        - screenshot: Base64 JPEG data with metadata (same as /vision/screenshot)

    Example regions:
        - Taskbar: {"x": 0, "y": 1040, "width": 1920, "height": 40}
        - Title bar: {"x": 0, "y": 0, "width": 1920, "height": 50}
        - Center dialog: {"x": 660, "y": 340, "width": 600, "height": 400}
    """
    try:
        from vision.vision_service import VisionConfig, optimize_region

        data = request.json or {}

        x = data.get("x")
        y = data.get("y")
        width = data.get("width")
        height = data.get("height")
        jpeg_quality = data.get("jpeg_quality", VisionConfig.jpeg_quality)

        if x is None or y is None or width is None or height is None:
            return error_response(
                "MISSING_PARAMETER", "x, y, width, height are all required", 400
            )

        # Take screenshot
        screenshot = ImageGrab.grab()

        # Optimize region
        start = time.time()
        optimized = optimize_region(
            screenshot,
            x=x,
            y=y,
            width=width,
            height=height,
            jpeg_quality=jpeg_quality,
        )
        elapsed = time.time() - start

        return success_response(
            "vision_screenshot_region",
            elapsed_ms=int(elapsed * 1000),
            requested_region={"x": x, "y": y, "width": width, "height": height},
            screenshot=optimized.to_dict(),
        )
    except Exception as e:
        logging.exception("Vision screenshot region failed")
        return error_response("VISION_SCREENSHOT_REGION_FAILED", str(e))


@app.route("/vision/analyze_or_screenshot", methods=["POST"])
def vision_analyze_or_screenshot():
    """
    Smart endpoint that respects vision mode configuration.

    Behavior by mode:
        - 'local': Full OmniParser + OCR analysis (returns elements, text)
        - 'agent': Returns optimized screenshot only (for AI agent vision)
        - 'auto': Tries local first; if it fails or returns nothing useful,
                  includes screenshot as fallback

    Request body:
        - use_cache (bool, optional): Use OCR cache (default True, local mode only)
        - jpeg_quality (int, optional): Override JPEG quality for screenshot
        - force_screenshot (bool, optional): Always include screenshot even in local mode

    Returns:
        Depends on mode:
        - local: elements, text_regions, full_text
        - agent: screenshot (base64 with metadata)
        - auto: elements, text_regions, full_text, and screenshot if local produced
                limited results
    """
    try:
        from vision.vision_service import VisionConfig, optimize_screenshot

        data = request.json or {}
        use_cache = data.get("use_cache", True)
        jpeg_quality = data.get("jpeg_quality", VisionConfig.jpeg_quality)
        force_screenshot = data.get("force_screenshot", False)

        mode = VisionConfig.mode
        screenshot = ImageGrab.grab()

        start = time.time()
        result = {"mode": mode}

        if mode == "agent":
            # Agent mode: just return screenshot
            optimized = optimize_screenshot(screenshot, jpeg_quality=jpeg_quality)
            result["screenshot"] = optimized.to_dict()

        elif mode == "local":
            # Local mode: full analysis
            service = get_vision_service()
            analysis = service.analyze_screen(screenshot, use_cache=use_cache)
            result.update(analysis)

            # Optionally include screenshot
            if force_screenshot:
                optimized = optimize_screenshot(screenshot, jpeg_quality=jpeg_quality)
                result["screenshot"] = optimized.to_dict()

        else:  # auto mode
            # Try local analysis first
            service = get_vision_service()
            try:
                analysis = service.analyze_screen(screenshot, use_cache=use_cache)
                result.update(analysis)

                # If local produced limited results, include screenshot as fallback
                element_count = analysis.get("element_count", 0)
                text_count = analysis.get("text_count", 0)

                if element_count < 3 and text_count < 5:
                    # Limited results - include screenshot for agent fallback
                    result["fallback_reason"] = (
                        f"Limited local results: {element_count} elements, {text_count} text regions"
                    )
                    optimized = optimize_screenshot(
                        screenshot, jpeg_quality=jpeg_quality
                    )
                    result["screenshot"] = optimized.to_dict()
                elif force_screenshot:
                    optimized = optimize_screenshot(
                        screenshot, jpeg_quality=jpeg_quality
                    )
                    result["screenshot"] = optimized.to_dict()

            except Exception as local_error:
                # Local failed - fall back to screenshot
                result["fallback_reason"] = f"Local analysis failed: {str(local_error)}"
                optimized = optimize_screenshot(screenshot, jpeg_quality=jpeg_quality)
                result["screenshot"] = optimized.to_dict()

        elapsed = time.time() - start
        result["elapsed_ms"] = int(elapsed * 1000)

        return success_response("vision_analyze_or_screenshot", **result)

    except Exception as e:
        logging.exception("Vision analyze_or_screenshot failed")
        return error_response("VISION_ANALYZE_OR_SCREENSHOT_FAILED", str(e))


# ==================== SCREENSHOT CACHE ENDPOINTS (v4.3) ====================


@app.route("/vision/screenshot_cache/stats", methods=["GET"])
def vision_screenshot_cache_stats():
    """
    Get screenshot cache statistics.

    Returns:
        - entries: Number of cached screenshots
        - max_entries: Maximum cache capacity
        - ttl_seconds: Time-to-live for cached screenshots
        - hits: Number of cache hits
        - misses: Number of cache misses
        - hit_rate: Cache hit ratio (0-1)
    """
    try:
        from vision.vision_service import get_screenshot_cache

        cache = get_screenshot_cache()
        return success_response("vision_screenshot_cache_stats", **cache.get_stats())
    except Exception as e:
        logging.exception("Vision screenshot cache stats failed")
        return error_response("VISION_SCREENSHOT_CACHE_STATS_FAILED", str(e))


@app.route("/vision/screenshot_cache/clear", methods=["POST"])
def vision_screenshot_cache_clear():
    """
    Clear the screenshot cache.

    Returns:
        - message: Confirmation message
    """
    try:
        from vision.vision_service import get_screenshot_cache

        cache = get_screenshot_cache()
        cache.clear()
        return success_response(
            "vision_screenshot_cache_clear", message="Screenshot cache cleared"
        )
    except Exception as e:
        logging.exception("Vision screenshot cache clear failed")
        return error_response("VISION_SCREENSHOT_CACHE_CLEAR_FAILED", str(e))


@app.route("/vision/screenshot_cache/config", methods=["POST"])
def vision_screenshot_cache_config():
    """
    Update screenshot cache configuration.

    Request body (all optional):
        - ttl_seconds (float): Time-to-live for cached screenshots (default 2.0)
        - max_entries (int): Maximum number of cached screenshots (default 5)

    Returns:
        - Updated cache statistics
    """
    try:
        from vision.vision_service import get_screenshot_cache

        data = request.json or {}
        cache = get_screenshot_cache()

        if "ttl_seconds" in data:
            cache.set_ttl(float(data["ttl_seconds"]))
        if "max_entries" in data:
            cache.set_max_entries(int(data["max_entries"]))

        return success_response(
            "vision_screenshot_cache_config",
            message="Cache configuration updated",
            **cache.get_stats(),
        )
    except Exception as e:
        logging.exception("Vision screenshot cache config failed")
        return error_response("VISION_SCREENSHOT_CACHE_CONFIG_FAILED", str(e))


@app.route("/vision/screenshot_cached", methods=["POST"])
def vision_screenshot_cached():
    """
    Take an optimized screenshot with caching support.

    If a cached screenshot exists within the TTL window, returns it instead
    of capturing a new one. Useful for rapid successive requests.

    Request body (all optional):
        - use_cache (bool): Use cached screenshot if available (default True)
        - jpeg_quality (int): Override JPEG quality (1-100)
        - max_width (int): Override maximum width
        - max_height (int): Override maximum height
        - include_thumbnail (bool): Include a smaller preview

    Returns:
        - screenshot: Base64 JPEG data with metadata
        - cache_hit: Whether this was served from cache
    """
    try:
        from vision.vision_service import (
            VisionConfig,
            optimize_screenshot,
            get_screenshot_cache,
        )

        data = request.json or {}
        use_cache = data.get("use_cache", True)

        # Get parameters
        jpeg_quality = data.get("jpeg_quality", VisionConfig.jpeg_quality)
        max_width = data.get("max_width", VisionConfig.max_width)
        max_height = data.get("max_height", VisionConfig.max_height)
        include_thumbnail = data.get(
            "include_thumbnail", VisionConfig.include_thumbnail
        )
        thumbnail_max_size = data.get(
            "thumbnail_max_size", VisionConfig.thumbnail_max_size
        )

        cache = get_screenshot_cache()
        cache_hit = False
        optimized = None

        # Try cache first
        if use_cache:
            optimized = cache.get(
                region=None,
                quality=jpeg_quality,
                max_width=max_width,
                max_height=max_height,
            )
            if optimized:
                cache_hit = True

        start = time.time()

        if not optimized:
            # Take new screenshot
            screenshot = ImageGrab.grab()
            optimized = optimize_screenshot(
                screenshot,
                max_width=max_width,
                max_height=max_height,
                jpeg_quality=jpeg_quality,
                include_thumbnail=include_thumbnail,
                thumbnail_max_size=thumbnail_max_size,
            )
            # Store in cache
            cache.put(
                optimized,
                region=None,
                quality=jpeg_quality,
                max_width=max_width,
                max_height=max_height,
            )

        elapsed = time.time() - start

        return success_response(
            "vision_screenshot_cached",
            elapsed_ms=int(elapsed * 1000),
            cache_hit=cache_hit,
            screenshot=optimized.to_dict(),
        )
    except Exception as e:
        logging.exception("Vision screenshot cached failed")
        return error_response("VISION_SCREENSHOT_CACHED_FAILED", str(e))


# ==================== CONTEXT MANAGER ENDPOINTS ====================


@app.route("/context/stats", methods=["GET"])
def context_stats():
    """Get context manager statistics (cache hits, misses, etc.)"""
    try:
        ctx = get_context_manager()
        stats = ctx.get_stats()
        stats["instance_id"] = id(ctx)  # Debug: show object ID
        return success_response(action="context_stats", **stats)
    except Exception as e:
        logging.exception("Context stats failed")
        return error_response("CONTEXT_STATS_FAILED", str(e))


@app.route("/context/clear", methods=["POST"])
def context_clear():
    """Clear all context manager caches."""
    try:
        ctx = get_context_manager()
        ctx.clear()
        return success_response(action="context_clear", message="Cache cleared")
    except Exception as e:
        logging.exception("Context clear failed")
        return error_response("CONTEXT_CLEAR_FAILED", str(e))


@app.route("/context/enable", methods=["POST"])
def context_enable():
    """Enable context manager caching."""
    try:
        ctx = get_context_manager()
        ctx.enable()
        return success_response(action="context_enable", enabled=True)
    except Exception as e:
        logging.exception("Context enable failed")
        return error_response("CONTEXT_ENABLE_FAILED", str(e))


@app.route("/context/disable", methods=["POST"])
def context_disable():
    """Disable context manager caching (pass-through mode)."""
    try:
        ctx = get_context_manager()
        ctx.disable()
        return success_response(action="context_disable", enabled=False)
    except Exception as e:
        logging.exception("Context disable failed")
        return error_response("CONTEXT_DISABLE_FAILED", str(e))


@app.route("/context/invalidate", methods=["POST"])
def context_invalidate():
    """Invalidate cache for a specific window."""
    try:
        data = request.get_json() or {}
        window = data.get("window")

        if not window:
            return error_response("MISSING_PARAM", "window is required")

        ctx = get_context_manager()
        ctx.cache.invalidate_window(window)
        return success_response(
            action="context_invalidate", window=window, message="Cache invalidated"
        )
    except Exception as e:
        logging.exception("Context invalidate failed")
        return error_response("CONTEXT_INVALIDATE_FAILED", str(e))


# ==================== STREAMING ENDPOINTS (v4.3) ====================


@sock.route("/vision/stream")
def vision_stream(ws):
    """
    WebSocket endpoint for real-time screenshot streaming.
    """
    from vision.streaming import get_vision_streamer

    # Get parameters from query string (Flask-Sock handles this via request.args)
    fps = int(request.args.get("fps", 5))
    quality = request.args.get("quality")
    if quality:
        quality = int(quality)
    max_width = request.args.get("max_width")
    if max_width:
        max_width = int(max_width)
    max_height = request.args.get("max_height")
    if max_height:
        max_height = int(max_height)

    streamer = get_vision_streamer()
    streamer.stream_screenshots(
        ws, fps=fps, quality=quality, max_width=max_width, max_height=max_height
    )


# ==================== MAIN ====================

if __name__ == "__main__":
    print("Starting Python Bridge (pywinauto) on port 5001...")
    app.run(port=5001, host="127.0.0.1")
