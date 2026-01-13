"""
OCR Benchmark - Compare Windows.Media.Ocr vs RapidOCR performance
"""

import time
from PIL import ImageGrab
import sys
import os

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def benchmark_ocr():
    print("=" * 70)
    print("OCR Performance Benchmark: Windows.Media.Ocr vs RapidOCR")
    print("=" * 70)

    # Capture screenshot
    print("\n[1] Capturing screenshot...")
    t0 = time.perf_counter()
    screenshot = ImageGrab.grab()
    t_capture = (time.perf_counter() - t0) * 1000
    print(
        f"    Screenshot: {screenshot.size[0]}x{screenshot.size[1]} in {t_capture:.1f}ms"
    )

    # ==================== WINDOWS OCR ====================
    print("\n" + "=" * 70)
    print("[2] Windows.Media.Ocr (PowerShell subprocess)")
    print("=" * 70)

    try:
        from vision.ocr import WindowsOCR

        windows_ocr = WindowsOCR()

        # Warmup
        print("    Warmup run...")
        _ = windows_ocr.recognize(screenshot)

        # Benchmark runs
        times_windows = []
        for i in range(3):
            t0 = time.perf_counter()
            regions_windows = windows_ocr.recognize(screenshot)
            elapsed = (time.perf_counter() - t0) * 1000
            times_windows.append(elapsed)
            print(f"    Run {i + 1}: {elapsed:.1f}ms - {len(regions_windows)} regions")

        avg_windows = sum(times_windows) / len(times_windows)
        print(f"\n    AVERAGE: {avg_windows:.1f}ms")
        print(f"    Text regions: {len(regions_windows)}")

        # Sample text
        sample_texts = [r.text for r in regions_windows[:10]]
        print(f"    Sample text: {sample_texts}")

    except Exception as e:
        print(f"    ERROR: {e}")
        avg_windows = None
        regions_windows = []

    # ==================== RAPIDOCR ====================
    print("\n" + "=" * 70)
    print("[3] RapidOCR (ONNX Runtime)")
    print("=" * 70)

    try:
        from vision.ocr_rapid import RapidOCREngine

        print("    Loading RapidOCR engine (first time includes model load)...")
        t0 = time.perf_counter()
        rapid_ocr = RapidOCREngine()

        # First run (includes model loading)
        t_load_start = time.perf_counter()
        regions_rapid = rapid_ocr.recognize(screenshot)
        t_first = (time.perf_counter() - t_load_start) * 1000
        print(
            f"    First run (with model load): {t_first:.1f}ms - {len(regions_rapid)} regions"
        )

        # Benchmark runs (models already loaded)
        times_rapid = []
        for i in range(3):
            t0 = time.perf_counter()
            regions_rapid = rapid_ocr.recognize(screenshot)
            elapsed = (time.perf_counter() - t0) * 1000
            times_rapid.append(elapsed)
            print(f"    Run {i + 1}: {elapsed:.1f}ms - {len(regions_rapid)} regions")

        avg_rapid = sum(times_rapid) / len(times_rapid)
        print(f"\n    AVERAGE: {avg_rapid:.1f}ms")
        print(f"    Text regions: {len(regions_rapid)}")

        # Sample text
        sample_texts = [r.text for r in regions_rapid[:10]]
        print(f"    Sample text: {sample_texts}")

    except Exception as e:
        import traceback

        print(f"    ERROR: {e}")
        traceback.print_exc()
        avg_rapid = None
        regions_rapid = []

    # ==================== COMPARISON ====================
    print("\n" + "=" * 70)
    print("COMPARISON SUMMARY")
    print("=" * 70)

    if avg_windows and avg_rapid:
        speedup = avg_windows / avg_rapid
        savings = avg_windows - avg_rapid

        print(f"\n  | OCR Engine       | Avg Time | Regions | Speedup |")
        print(f"  |------------------|----------|---------|---------|")
        print(
            f"  | Windows.Media.Ocr | {avg_windows:>6.1f}ms | {len(regions_windows):>7} | 1.0x    |"
        )
        print(
            f"  | RapidOCR         | {avg_rapid:>6.1f}ms | {len(regions_rapid):>7} | {speedup:.1f}x    |"
        )
        print(f"\n  Time savings: {savings:.1f}ms per OCR call ({speedup:.1f}x faster)")

        # Accuracy comparison
        print(
            f"\n  Region count difference: {abs(len(regions_windows) - len(regions_rapid))}"
        )
        if len(regions_rapid) >= len(regions_windows) * 0.8:
            print("  Coverage: RapidOCR has comparable or better coverage")
        else:
            print("  Coverage: RapidOCR may miss some text (check accuracy)")
    else:
        print("  Unable to complete comparison (one or both OCR engines failed)")

    return {
        "windows_avg_ms": avg_windows,
        "rapid_avg_ms": avg_rapid,
        "windows_regions": len(regions_windows) if regions_windows else 0,
        "rapid_regions": len(regions_rapid) if regions_rapid else 0,
    }


if __name__ == "__main__":
    benchmark_ocr()
