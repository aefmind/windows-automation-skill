"""
Diagnostic script for OmniParser performance analysis
Run this to get detailed timing and raw output statistics
"""

import time
import numpy as np
from PIL import ImageGrab, Image
import sys
import os

# Add parent to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from vision.detector import VisionDetector, Detection


def run_diagnostic():
    results = {}

    print("=" * 60)
    print("OmniParser Vision Layer Diagnostic")
    print("=" * 60)

    # 1. Screenshot capture timing
    print("\n[1] Screenshot Capture...")
    t0 = time.perf_counter()
    screenshot = ImageGrab.grab()
    t_screenshot = (time.perf_counter() - t0) * 1000
    print(f"    Time: {t_screenshot:.1f}ms")
    print(f"    Size: {screenshot.size[0]}x{screenshot.size[1]}")
    print(f"    Mode: {screenshot.mode}")
    results["screenshot_ms"] = t_screenshot
    results["screen_size"] = screenshot.size

    # 2. Model loading timing
    print("\n[2] Model Loading...")
    t0 = time.perf_counter()
    detector = VisionDetector(
        confidence_threshold=0.01
    )  # Very low threshold for diagnostic
    _ = detector.session  # Force load
    t_load = (time.perf_counter() - t0) * 1000
    print(f"    Time: {t_load:.1f}ms")
    print(f"    Model: {os.path.basename(detector.model_path)}")
    results["model_load_ms"] = t_load

    # 3. Preprocessing timing
    print("\n[3] Image Preprocessing...")
    if screenshot.mode != "RGB":
        screenshot = screenshot.convert("RGB")
    t0 = time.perf_counter()
    input_tensor, scale_x, scale_y, pad_x, pad_y = detector._preprocess_image(
        screenshot
    )
    t_preprocess = (time.perf_counter() - t0) * 1000
    print(f"    Time: {t_preprocess:.1f}ms")
    print(f"    Input shape: {input_tensor.shape}")
    print(f"    Scale: {scale_x:.4f}")
    print(f"    Padding: ({pad_x}, {pad_y})")
    print(f"    Value range: [{input_tensor.min():.3f}, {input_tensor.max():.3f}]")
    results["preprocess_ms"] = t_preprocess

    # 4. ONNX Inference timing
    print("\n[4] ONNX Inference...")
    input_name = detector.session.get_inputs()[0].name
    t0 = time.perf_counter()
    outputs = detector.session.run(None, {input_name: input_tensor})
    t_inference = (time.perf_counter() - t0) * 1000
    print(f"    Time: {t_inference:.1f}ms")
    print(f"    Output shape: {outputs[0].shape}")
    results["inference_ms"] = t_inference

    # 5. Raw output analysis
    print("\n[5] Raw Model Output Analysis...")
    raw_output = outputs[0]
    predictions = raw_output[0].T  # [N, 5]
    print(f"    Total predictions: {len(predictions)}")

    # Confidence distribution
    confidences = predictions[:, 4]
    print(f"    Confidence stats:")
    print(f"      Min: {confidences.min():.4f}")
    print(f"      Max: {confidences.max():.4f}")
    print(f"      Mean: {confidences.mean():.4f}")
    print(f"      Std: {confidences.std():.4f}")

    # Count by threshold
    thresholds = [0.01, 0.05, 0.1, 0.2, 0.3, 0.5, 0.7]
    print(f"\n    Detections by threshold:")
    for th in thresholds:
        count = np.sum(confidences >= th)
        print(f"      >= {th}: {count}")
    results["raw_predictions"] = len(predictions)
    results["max_confidence"] = float(confidences.max())

    # 6. Postprocessing with different thresholds
    print("\n[6] Postprocessing Analysis...")
    for th in [0.01, 0.1, 0.3]:
        detector.confidence_threshold = th
        t0 = time.perf_counter()
        detections = detector._postprocess_output(outputs[0], scale_x, pad_x, pad_y)
        t_post = (time.perf_counter() - t0) * 1000
        print(f"    Threshold {th}: {len(detections)} detections ({t_post:.1f}ms)")
        if detections:
            print(
                f"      Top 3 confidences: {[round(d.confidence, 3) for d in detections[:3]]}"
            )

    # 7. Full detection timing (threshold 0.1)
    print("\n[7] Full Detection Pipeline (threshold=0.1)...")
    detector.confidence_threshold = 0.1
    t0 = time.perf_counter()
    detections = detector.detect(screenshot)
    t_full = (time.perf_counter() - t0) * 1000
    print(f"    Time: {t_full:.1f}ms")
    print(f"    Detections: {len(detections)}")
    results["full_pipeline_ms"] = t_full
    results["detections_th01"] = len(detections)

    if detections:
        print("\n    Detected elements:")
        for i, d in enumerate(detections[:10]):
            print(
                f"      [{i}] pos=({d.x}, {d.y}) size={d.width}x{d.height} conf={d.confidence:.3f}"
            )

    # 8. Summary
    print("\n" + "=" * 60)
    print("PERFORMANCE SUMMARY")
    print("=" * 60)
    print(f"  Screenshot:    {results['screenshot_ms']:6.1f}ms")
    print(f"  Preprocess:    {results['preprocess_ms']:6.1f}ms")
    print(f"  Inference:     {results['inference_ms']:6.1f}ms")
    print(f"  Full pipeline: {results['full_pipeline_ms']:6.1f}ms")
    print(f"\n  Raw predictions: {results['raw_predictions']}")
    print(f"  Max confidence:  {results['max_confidence']:.4f}")
    print(f"  Final detections (th=0.1): {results['detections_th01']}")

    return results


if __name__ == "__main__":
    run_diagnostic()
