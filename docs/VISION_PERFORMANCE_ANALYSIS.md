# Vision Layer Performance Analysis

## Executive Summary

The Vision Layer (OmniParser + ONNX + WinRT OCR) delivers **sub-300ms full analysis** with a combined pipeline taking ~250ms. After WinRT native OCR optimization, we achieved **3.2x faster OCR** by bypassing PowerShell subprocess overhead.

**Key Achievement**: Full screen OCR in 166ms (was 530ms), region OCR in 60ms (was 120ms).

---

## Performance Benchmarks (v4.2 - WinRT Native)

### Timing Breakdown

| Component            | Time (ms) | % of Total |
| -------------------- | --------- | ---------- |
| Screenshot Capture   | 49        | 20%        |
| Image Preprocessing  | 17        | 7%         |
| ONNX Inference       | 66        | 26%        |
| Postprocessing (NMS) | 12        | 5%         |
| **OmniParser Total** | **~76**   | **30%**    |
| **WinRT OCR**        | **~166**  | **66%**    |
| **Full Analyze**     | **~250**  | **100%**   |

### Key Metrics

```
Screen Resolution: 1920x1080
Model Input Size: 640x640 (letterboxed)
Scale Factor: 3.0x
Padding: (0, 140) - top/bottom for 16:9 aspect
Confidence Threshold: 0.10 (optimized from 0.30)
```

---

## OmniParser ONNX Analysis

### Model Specifications

- **Model**: `omniparser-icon_detect_fp16.onnx` (6MB, FP16)
- **Architecture**: YOLO-style detector
- **Output Shape**: `[1, 5, 8400]` (8400 candidate boxes)
- **Format**: `[x_center, y_center, width, height, confidence]`

### Raw Output Statistics

```
Total Predictions: 8400
Confidence Distribution:
  Min:  0.0000
  Max:  0.4087
  Mean: 0.0044
  Std:  0.0203

Detections by Threshold:
  >= 0.01: 675
  >= 0.05: 182
  >= 0.10: 90  ← Current threshold
  >= 0.20: 21
  >= 0.30: 3
  >= 0.50: 0
  >= 0.70: 0
```

### After NMS (Non-Maximum Suppression)

| Threshold | Detections | Processing Time | Recommendation           |
| --------- | ---------- | --------------- | ------------------------ |
| 0.01      | 174        | 85ms            | Too many false positives |
| **0.10**  | **37**     | **12ms**        | **Optimal (current)**    |
| 0.30      | 3          | 7ms             | Too aggressive           |

### Detected Elements (Sample @ 0.10 threshold)

```
[0] pos=(1228, 1058) size=33x44 conf=0.409  # Taskbar icon
[1] pos=(1272, 1057) size=34x47 conf=0.354  # Taskbar icon
[2] pos=(1418, 380)  size=462x23 conf=0.319 # UI element
[3] pos=(1153, 307)  size=48x25 conf=0.286  # Button
[4] pos=(1185, 1057) size=29x37 conf=0.268  # Taskbar icon
[5] pos=(958, 977)   size=38x42 conf=0.250  # Icon
[6] pos=(1878, 1054) size=56x32 conf=0.250  # System tray
[7] pos=(1275, 52)   size=373x24 conf=0.245 # Title bar
[8] pos=(867, 1058)  size=45x39 conf=0.229  # Taskbar icon
[9] pos=(38, 210)    size=40x48 conf=0.224  # Desktop icon
```

---

## Windows OCR Analysis

### Performance (v4.2 - WinRT Native)

- **Time**: **166ms** for full 1920x1080 screen (was 530-570ms with PowerShell)
- **Region OCR**: **60ms** for small areas (taskbar, dialogs)
- **Cached**: **<1ms** for repeated queries
- **Text Regions Detected**: 340-370 (depends on screen content)
- **Confidence**: 1.0 for all detected text (Windows OCR is binary)

### OCR Optimization (v4.2 - WinRT Native)

After extensive benchmarking, we implemented several optimizations including **WinRT native OCR** which bypasses PowerShell subprocess overhead:

| Optimization | Time | Improvement |
|--------------|------|-------------|
| **WinRT Native OCR** | **166ms avg** | **3.2x faster than PowerShell** |
| **Caching** | **<1ms** | **3400x faster for cache hits** |
| **Region-based OCR** | **60ms** | **8.8x faster for small areas** |

### Benchmark Results (v4.2)

```
WinRT Full Screen OCR:   166ms (avg), 156-200ms range
PowerShell OCR:          535ms (baseline - no longer used by default)
Cached OCR:              0.16ms (instant for repeated queries)
Region OCR (taskbar):    60ms (was 120ms with PowerShell)
```

### WinRT Native OCR Implementation

We replaced PowerShell subprocess calls with direct Windows Runtime API access via `winrt-runtime`:

```python
# Old approach (PowerShell subprocess) - 535ms
subprocess.run(["powershell", "-File", "ocr_script.ps1", image_path])

# New approach (WinRT native) - 166ms
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.graphics.imaging import BitmapDecoder
engine = OcrEngine.try_create_from_language(Language("en-US"))
result = await engine.recognize_async(bitmap)
```

**Packages installed:**
- `winrt-runtime`
- `winrt-Windows.Media.Ocr`
- `winrt-Windows.Graphics.Imaging`
- `winrt-Windows.Storage`
- `winrt-Windows.Globalization`
- `winrt-Windows.Foundation`

### PowerShell Bottleneck (Solved!)

The primary bottleneck was PowerShell subprocess overhead:
- **PowerShell startup**: ~200-250ms (now bypassed)
- **Windows.Media.Ocr processing**: ~166ms (direct access)
- **Total savings**: ~370ms per OCR call

### Alternative OCR Engines Tested

| Engine | Time | Regions | Result |
|--------|------|---------|--------|
| **WinRT Native** | **166ms** | 344 | ✅ **Best - 3.2x faster** |
| Windows OCR (PowerShell) | 530ms | 354 | ✅ Fallback option |
| RapidOCR (ONNX) | 7,400ms | 135 | ❌ Too slow on CPU |
| Image scaling (50%) | 445ms | 19 | ❌ Loses 95% of text |
| pythonnet (WinRT via CLR) | 1,600ms | 0 | ❌ Doesn't work |

### Coverage Quality

The OCR detected all visible text including:

- Application titles ("Calculator", "OpenCode")
- Button labels ("CE", numbers 0-9)
- Menu items and status bar text
- System tray information
- Even small text in complex UIs

---

## Comparison: OmniParser vs OCR

| Aspect               | OmniParser (th=0.10)       | Windows OCR        |
| -------------------- | -------------------------- | ------------------ |
| **Speed**            | ~76ms                      | ~650ms             |
| **Elements Found**   | 37                         | 185-207            |
| **Best For**         | Icons, buttons, UI regions | Text labels, menus |
| **Confidence Range** | 0.10-0.41                  | 1.0 (binary)       |
| **Training Domain**  | Web/mobile UI icons        | Any text           |

---

## Threshold Optimization Results

### Before Optimization (v4.0 initial)

```
Threshold: 0.30
Detections: 2-5
Status: Missed most UI elements
```

### After Optimization (v4.0 updated)

```
Threshold: 0.10
Detections: 37
Status: Good coverage of icons and UI regions
Improvement: 7-18x more detections
```

### Files Modified

- `src/bridge_python/bridge.py`: `confidence_threshold=0.10`
- `src/bridge_python/vision/detector.py`: default `0.15`
- `src/bridge_python/vision/vision_service.py`: default `0.15`

---

## What Gets Detected

### Well Detected (confidence > 0.20)

- Taskbar icons
- Desktop shortcuts
- Application icons in title bars
- Buttons with distinct graphical elements
- Large UI regions/panels

### Partially Detected (confidence 0.10-0.20)

- Text fields with icons
- Status bar icons
- Tool buttons
- Navigation elements

### Not Detected (use OCR)

- Text-only buttons
- Windows native controls (checkboxes, sliders)
- Flat/minimal UI elements without icons

---

## Recommendations

### 1. Current Configuration (Optimal)

```python
# Production settings
confidence_threshold = 0.10  # 37 detections, good balance
```

### 2. Hybrid Strategy (Implemented)

The current implementation correctly uses **hybrid detection**:

1. **OmniParser** for icon/button detection (37 elements)
2. **Windows OCR** for text-based elements (200+ regions)
3. **Smart fallback** when FlaUI fails

### 3. Use Case Guide

| Task                 | Best Method                       | Expected Performance |
| -------------------- | --------------------------------- | -------------------- |
| Click button by text | WinRT OCR + `find_text`           | **166ms**            |
| Click icon           | OmniParser + `vision_click`       | 76ms                 |
| Read screen content  | OCR (cached)                      | **<1ms cached**      |
| Search in taskbar    | Region-based OCR                  | **60ms**             |
| Automated UI testing | FlaUI (first) → Vision (fallback) | <50ms / 250ms        |

### 4. OCR Optimization Strategy

Choose the right OCR method based on your use case:

```
┌─────────────────────────────────────────────────────────────────┐
│                     OCR DECISION TREE (v4.2)                    │
├─────────────────────────────────────────────────────────────────┤
│ Q: Is this a repeated query (same screen)?                      │
│    YES → Use /vision/ocr_cached (<1ms)                          │
│    NO  ↓                                                         │
│                                                                  │
│ Q: Do you know where the text should be?                        │
│    YES → Use /vision/ocr_region with coordinates (60ms)         │
│    NO  ↓                                                         │
│                                                                  │
│ Q: Do you have location hints?                                  │
│    YES → Use /vision/find_text_fast with hint_regions (60ms)    │
│    NO  → Use /vision/ocr (166ms full scan - WinRT)              │
└─────────────────────────────────────────────────────────────────┘
```

### 5. Future Improvements

1. ~~**pythonnet direct Windows API**: Bypass PowerShell overhead~~ ✅ **Done - WinRT native**
2. **Adaptive thresholding**: Automatically lower threshold if < 10 detections
3. **Alternative models**: Florence-2, Grounding DINO for better Windows UI detection
4. **Custom fine-tuning**: Train on Windows 11 UI specifically

---

## Conclusion

The Vision Layer performs excellently after WinRT native OCR optimization:

| Metric               | v4.0   | v4.1       | v4.2 WinRT     | Assessment         |
| -------------------- | ------ | ---------- | -------------- | ------------------ |
| Full analysis time   | ~720ms | ~600ms     | **~250ms**     | **3x faster**      |
| OmniParser inference | ~90ms  | ~76ms      | ~76ms          | Very fast          |
| Icon detection       | 2-5    | 37         | 37             | **7x improvement** |
| Text detection       | 185+   | 354+       | 344            | Comprehensive      |
| OCR (full screen)    | 530ms  | 530ms      | **166ms**      | **3.2x faster**    |
| OCR (cached)         | N/A    | <1ms       | **<1ms**       | **Instant**        |
| OCR (region)         | N/A    | 120ms      | **60ms**       | **8.8x faster**    |

**Key Achievements in v4.2:**
- ✅ **WinRT native OCR** bypasses PowerShell - 3.2x faster (166ms vs 535ms)
- ✅ **OCR caching** provides instant results for repeated queries (<1ms)
- ✅ **Region-based OCR** is 8.8x faster for targeted searches (60ms)
- ✅ New optimized API endpoints for fine-grained control
- ✅ PowerShell bottleneck **SOLVED** via winrt-runtime package

**The hybrid approach (FlaUI → Vision fallback) with optimized OmniParser and Windows OCR provides robust automation capabilities for Windows desktop applications.**

---

## Test Commands

```powershell
# Run full diagnostic (use your venv)
& ".\\.venv\\Scripts\\python.exe" `
  ".\\src\\bridge_python\\vision\\diagnostic.py"

# Test vision endpoints
curl -X POST http://127.0.0.1:5001/vision/detect -H "Content-Type: application/json" -d "{}"
curl -X POST http://127.0.0.1:5001/vision/analyze -H "Content-Type: application/json" -d "{}"

# Check health
curl http://127.0.0.1:5001/health
curl http://127.0.0.1:5001/context/stats
```

---

## OCR API Reference (v4.1)

### Standard Endpoints

| Endpoint | Method | Description | Performance |
|----------|--------|-------------|-------------|
| `/vision/ocr` | POST | Full screen OCR | ~530ms |
| `/vision/find_text` | POST | Find element by text | ~650ms |
| `/vision/click_text` | POST | Click on text | ~700ms |
| `/vision/analyze` | POST | Full screen analysis | ~720ms |

### Optimized Endpoints (New in v4.1)

| Endpoint | Method | Description | Performance |
|----------|--------|-------------|-------------|
| `/vision/ocr_cached` | POST | OCR with caching | <1ms cached |
| `/vision/ocr_region` | POST | Region-based OCR | ~120ms |
| `/vision/find_text_fast` | POST | Fast text search with hints | ~120-380ms |
| `/vision/cache_stats` | GET | Cache statistics | <1ms |
| `/vision/cache_clear` | POST | Clear OCR cache | <1ms |

### `/vision/ocr_cached`

OCR with explicit cache control.

**Request:**
```json
{
  "use_cache": true,
  "clear_cache": false,
  "full_text": false
}
```

**Response:**
```json
{
  "status": "success",
  "elapsed_ms": 0,
  "region_count": 354,
  "cache_stats": {"entries": 1, "ttl_seconds": 2.0}
}
```

### `/vision/ocr_region`

OCR only a specific screen region (4x faster).

**Request:**
```json
{
  "x": 0,
  "y": 1040,
  "width": 1920,
  "height": 40
}
```

**Response:**
```json
{
  "status": "success",
  "elapsed_ms": 122,
  "region_count": 7,
  "search_region": {"x": 0, "y": 1040, "width": 1920, "height": 40}
}
```

### `/vision/find_text_fast`

Fast text search with optional region hints.

**Request:**
```json
{
  "text": "Search",
  "hint_regions": [
    {"x": 0, "y": 1040, "width": 1920, "height": 40}
  ]
}
```

**Common hint regions:**
- **Taskbar**: `{"x": 0, "y": 1040, "width": 1920, "height": 40}`
- **Title bar**: `{"x": 0, "y": 0, "width": 1920, "height": 50}`
- **Center dialog**: `{"x": 660, "y": 340, "width": 600, "height": 400}`

---

_Analysis performed: January 11, 2026_
_Vision Layer v4.1 (OCR optimized)_
