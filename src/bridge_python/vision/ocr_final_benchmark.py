"""
Final OCR Benchmark - Compare all optimization strategies
"""

import time
from PIL import ImageGrab
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def run_final_benchmark():
    print("=" * 70)
    print("OCR Optimization Final Benchmark")
    print("=" * 70)

    # Capture screenshot
    screenshot = ImageGrab.grab()
    print(f"Screenshot: {screenshot.size[0]}x{screenshot.size[1]}")

    results = {}

    # ==================== ORIGINAL WINDOWS OCR ====================
    print("\n[1] Original Windows OCR (PNG)...")
    from vision.ocr import WindowsOCR

    original_ocr = WindowsOCR()

    # Warmup
    _ = original_ocr.recognize(screenshot)

    times = []
    for i in range(3):
        t0 = time.perf_counter()
        regions = original_ocr.recognize(screenshot)
        times.append((time.perf_counter() - t0) * 1000)

    results["original"] = {
        "avg_ms": sum(times) / len(times),
        "regions": len(regions),
        "name": "Original (PNG)",
    }
    print(f"    Avg: {results['original']['avg_ms']:.0f}ms, Regions: {len(regions)}")

    # ==================== OPTIMIZED OCR (JPEG) ====================
    print("\n[2] Optimized OCR (JPEG + Cache)...")
    from vision.ocr_optimized import OptimizedOCR

    optimized_ocr = OptimizedOCR(cache_ttl=0.1)  # Short TTL for benchmark

    # First run (no cache)
    t0 = time.perf_counter()
    regions_opt = optimized_ocr.recognize(screenshot, use_cache=False)
    t_first = (time.perf_counter() - t0) * 1000
    print(f"    First run (no cache): {t_first:.0f}ms, Regions: {len(regions_opt)}")

    # Subsequent runs (fresh, no cache)
    times = []
    for i in range(3):
        optimized_ocr.clear_cache()
        t0 = time.perf_counter()
        regions_opt = optimized_ocr.recognize(screenshot, use_cache=False)
        times.append((time.perf_counter() - t0) * 1000)

    results["optimized"] = {
        "avg_ms": sum(times) / len(times),
        "regions": len(regions_opt),
        "name": "Optimized (JPEG)",
    }
    print(f"    Avg (no cache): {results['optimized']['avg_ms']:.0f}ms")

    # With cache
    optimized_ocr.clear_cache()
    _ = optimized_ocr.recognize(screenshot, use_cache=True)  # Prime cache

    t0 = time.perf_counter()
    regions_cached = optimized_ocr.recognize(screenshot, use_cache=True)
    t_cached = (time.perf_counter() - t0) * 1000

    results["cached"] = {
        "avg_ms": t_cached,
        "regions": len(regions_cached),
        "name": "Cached",
    }
    print(f"    Cached: {t_cached:.1f}ms (cache hit)")

    # ==================== REGION-BASED OCR ====================
    print("\n[3] Region-based OCR (taskbar only)...")

    # Define taskbar region (bottom of screen)
    taskbar_region = (0, 1040, 1920, 40)  # x, y, width, height

    times = []
    for i in range(3):
        t0 = time.perf_counter()
        regions_taskbar = optimized_ocr.recognize_region(
            screenshot,
            taskbar_region[0],
            taskbar_region[1],
            taskbar_region[2],
            taskbar_region[3],
        )
        times.append((time.perf_counter() - t0) * 1000)

    results["region_taskbar"] = {
        "avg_ms": sum(times) / len(times),
        "regions": len(regions_taskbar),
        "name": "Region (taskbar)",
    }
    print(
        f"    Avg: {results['region_taskbar']['avg_ms']:.0f}ms, Regions: {len(regions_taskbar)}"
    )

    # ==================== SMALL REGION OCR ====================
    print("\n[4] Small region OCR (400x200 area)...")

    small_region = (100, 100, 400, 200)

    times = []
    for i in range(3):
        t0 = time.perf_counter()
        regions_small = optimized_ocr.recognize_region(
            screenshot,
            small_region[0],
            small_region[1],
            small_region[2],
            small_region[3],
        )
        times.append((time.perf_counter() - t0) * 1000)

    results["region_small"] = {
        "avg_ms": sum(times) / len(times),
        "regions": len(regions_small),
        "name": "Region (400x200)",
    }
    print(
        f"    Avg: {results['region_small']['avg_ms']:.0f}ms, Regions: {len(regions_small)}"
    )

    # ==================== SUMMARY ====================
    print("\n" + "=" * 70)
    print("PERFORMANCE SUMMARY")
    print("=" * 70)

    baseline = results["original"]["avg_ms"]

    print(
        f"\n  | Method              | Avg Time | Regions | Speedup | Notes           |"
    )
    print(f"  |---------------------|----------|---------|---------|-----------------|")

    for key in ["original", "optimized", "cached", "region_taskbar", "region_small"]:
        r = results[key]
        speedup = baseline / r["avg_ms"] if r["avg_ms"] > 0 else float("inf")
        notes = ""
        if key == "optimized":
            notes = f"~{baseline - r['avg_ms']:.0f}ms saved"
        elif key == "cached":
            notes = "Cache hit"
        elif "region" in key:
            notes = "Partial screen"

        print(
            f"  | {r['name']:<19} | {r['avg_ms']:>6.0f}ms | {r['regions']:>7} | {speedup:>6.1f}x | {notes:<15} |"
        )

    # ==================== RECOMMENDATIONS ====================
    print("\n" + "=" * 70)
    print("RECOMMENDATIONS")
    print("=" * 70)

    savings = baseline - results["optimized"]["avg_ms"]
    print(f"""
  1. JPEG optimization saves ~{savings:.0f}ms ({savings / baseline * 100:.0f}% improvement)
     - Apply to: ocr.py (already updated)
     
  2. Caching provides instant results for repeated queries
     - Cache hit time: {results["cached"]["avg_ms"]:.1f}ms
     - Effective for: Static UIs, repeated searches
     
  3. Region-based OCR is MUCH faster for targeted searches
     - Taskbar (1920x40): {results["region_taskbar"]["avg_ms"]:.0f}ms vs {baseline:.0f}ms
     - Small (400x200): {results["region_small"]["avg_ms"]:.0f}ms vs {baseline:.0f}ms
     - Use when: Looking for specific UI element
     
  4. BOTTLENECK: PowerShell subprocess overhead (~200-250ms minimum)
     - This is unavoidable with current Windows OCR approach
     - To go faster: Would need GPU-based OCR or native Windows API
""")

    return results


if __name__ == "__main__":
    run_final_benchmark()
