"""
VisionDetector - OmniParser ONNX wrapper for UI element detection
Detects clickable UI elements in screenshots using the OmniParser icon detection model
"""

import os
import numpy as np
import onnxruntime as ort
from PIL import Image
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
import io
import base64


@dataclass
class Detection:
    """Represents a detected UI element"""

    x: int  # Center X coordinate (original resolution)
    y: int  # Center Y coordinate (original resolution)
    width: int  # Width in pixels (original resolution)
    height: int  # Height in pixels (original resolution)
    confidence: float  # Detection confidence 0-1

    @property
    def bounds(self) -> Tuple[int, int, int, int]:
        """Returns (left, top, right, bottom)"""
        half_w = self.width // 2
        half_h = self.height // 2
        return (self.x - half_w, self.y - half_h, self.x + half_w, self.y + half_h)

    @property
    def center(self) -> Tuple[int, int]:
        """Returns (x, y) center point"""
        return (self.x, self.y)

    def to_dict(self) -> Dict:
        """Convert to dictionary for JSON serialization"""
        left, top, right, bottom = self.bounds
        return {
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "confidence": round(self.confidence, 3),
        }


class VisionDetector:
    """
    OmniParser-based UI element detector.

    Uses the OmniParser icon detection ONNX model to find UI elements
    in screenshots. Handles coordinate scaling from model input size
    back to original resolution.
    """

    # OmniParser preprocessor settings
    MODEL_INPUT_SIZE = 640  # Longest edge
    RESCALE_FACTOR = 1.0 / 255.0  # 0.00392156862745098

    def __init__(
        self, model_path: Optional[str] = None, confidence_threshold: float = 0.15
    ):
        """
        Initialize the VisionDetector.

        Args:
            model_path: Path to OmniParser ONNX model. If None, uses default location.
            confidence_threshold: Minimum confidence for detections (0-1)
        """
        if model_path is None:
            # Default path relative to this file
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            model_path = os.path.join(
                base_dir, "models", "omniparser-icon_detect_fp16.onnx"
            )

        if not os.path.exists(model_path):
            raise FileNotFoundError(f"OmniParser model not found at: {model_path}")

        self.model_path = model_path
        self.confidence_threshold = confidence_threshold
        self._session: Optional[ort.InferenceSession] = None

    @property
    def session(self) -> ort.InferenceSession:
        """Lazy load the ONNX session"""
        if self._session is None:
            # Use CPU provider for compatibility
            self._session = ort.InferenceSession(
                self.model_path, providers=["CPUExecutionProvider"]
            )
        return self._session

    def _preprocess_image(
        self, image: Image.Image
    ) -> Tuple[np.ndarray, float, float, int, int]:
        """
        Preprocess image for OmniParser model.

        The OmniParser model expects square input (640x640) with letterboxing.
        We resize to fit within 640x640 while maintaining aspect ratio,
        then pad with gray to fill the square.

        Args:
            image: PIL Image in RGB format

        Returns:
            Tuple of (preprocessed_array, scale_x, scale_y, pad_x, pad_y)
        """
        original_width, original_height = image.size

        # Calculate scale to fit within MODEL_INPUT_SIZE while maintaining aspect ratio
        scale = min(
            self.MODEL_INPUT_SIZE / original_width,
            self.MODEL_INPUT_SIZE / original_height,
        )
        new_width = int(original_width * scale)
        new_height = int(original_height * scale)

        # Resize image
        resized = image.resize((new_width, new_height), Image.Resampling.BILINEAR)

        # Create square canvas with gray background (114 is standard YOLO padding)
        canvas = Image.new(
            "RGB", (self.MODEL_INPUT_SIZE, self.MODEL_INPUT_SIZE), (114, 114, 114)
        )

        # Calculate padding to center the image
        pad_x = (self.MODEL_INPUT_SIZE - new_width) // 2
        pad_y = (self.MODEL_INPUT_SIZE - new_height) // 2

        # Paste resized image onto canvas
        canvas.paste(resized, (pad_x, pad_y))

        # Convert to numpy and normalize
        img_array = np.array(canvas, dtype=np.float32)
        img_array = img_array * self.RESCALE_FACTOR

        # Transpose from HWC to CHW format
        img_array = np.transpose(img_array, (2, 0, 1))

        # Add batch dimension
        img_array = np.expand_dims(img_array, axis=0)

        # Calculate scale factor to convert back to original resolution
        # We use the same scale for both axes since we maintain aspect ratio
        inverse_scale = 1.0 / scale

        return img_array, inverse_scale, inverse_scale, pad_x, pad_y

    def _postprocess_output(
        self,
        output: np.ndarray,
        scale: float,
        pad_x: int,
        pad_y: int,
    ) -> List[Detection]:
        """
        Convert model output to Detection objects.

        The model outputs YOLO-style format: [batch, 5, num_boxes]
        where 5 = [x_center, y_center, width, height, confidence]

        Args:
            output: Model output array
            scale: Scale factor for coordinate conversion (same for x and y)
            pad_x: X padding added during preprocessing
            pad_y: Y padding added during preprocessing

        Returns:
            List of Detection objects
        """
        # Output shape: [1, 5, N] where N is number of detection boxes
        # Transpose to [N, 5] for easier processing
        predictions = output[0].T  # [N, 5]

        detections = []
        for pred in predictions:
            x_center, y_center, width, height, confidence = pred

            if confidence < self.confidence_threshold:
                continue

            # Remove padding offset and scale to original resolution
            x = int((x_center - pad_x) * scale)
            y = int((y_center - pad_y) * scale)
            w = int(width * scale)
            h = int(height * scale)

            # Skip detections that are mostly in the padding area
            if x < 0 or y < 0:
                continue

            detections.append(
                Detection(x=x, y=y, width=w, height=h, confidence=float(confidence))
            )

        # Sort by confidence (highest first)
        detections.sort(key=lambda d: d.confidence, reverse=True)

        # Apply NMS (Non-Maximum Suppression) to remove overlapping boxes
        detections = self._apply_nms(detections, iou_threshold=0.5)

        return detections

    def _apply_nms(
        self, detections: List[Detection], iou_threshold: float = 0.5
    ) -> List[Detection]:
        """
        Apply Non-Maximum Suppression to remove overlapping detections.

        Args:
            detections: List of Detection objects (already sorted by confidence)
            iou_threshold: IoU threshold for suppression

        Returns:
            Filtered list of detections
        """
        if len(detections) <= 1:
            return detections

        kept = []
        suppressed = set()

        for i, det in enumerate(detections):
            if i in suppressed:
                continue

            kept.append(det)

            # Suppress overlapping detections
            for j in range(i + 1, len(detections)):
                if j in suppressed:
                    continue

                iou = self._calculate_iou(det, detections[j])
                if iou > iou_threshold:
                    suppressed.add(j)

        return kept

    def _calculate_iou(self, det1: Detection, det2: Detection) -> float:
        """Calculate Intersection over Union between two detections"""
        l1, t1, r1, b1 = det1.bounds
        l2, t2, r2, b2 = det2.bounds

        # Calculate intersection
        left = max(l1, l2)
        top = max(t1, t2)
        right = min(r1, r2)
        bottom = min(b1, b2)

        if left >= right or top >= bottom:
            return 0.0

        intersection = (right - left) * (bottom - top)

        # Calculate union
        area1 = (r1 - l1) * (b1 - t1)
        area2 = (r2 - l2) * (b2 - t2)
        union = area1 + area2 - intersection

        return intersection / union if union > 0 else 0.0

    def detect(self, image: Image.Image) -> List[Detection]:
        """
        Detect UI elements in an image.

        Args:
            image: PIL Image (will be converted to RGB if needed)

        Returns:
            List of Detection objects with coordinates in original resolution
        """
        # Ensure RGB format
        if image.mode != "RGB":
            image = image.convert("RGB")

        # Preprocess - now returns 5 values including padding
        input_tensor, scale_x, scale_y, pad_x, pad_y = self._preprocess_image(image)

        # Get model input/output names
        input_name = self.session.get_inputs()[0].name

        # Run inference
        outputs = self.session.run(None, {input_name: input_tensor})

        # Postprocess - use scale_x (same as scale_y) and padding info
        detections = self._postprocess_output(outputs[0], scale_x, pad_x, pad_y)

        return detections

    def detect_from_bytes(self, image_bytes: bytes) -> List[Detection]:
        """
        Detect UI elements from image bytes.

        Args:
            image_bytes: Raw image bytes (PNG, JPEG, etc.)

        Returns:
            List of Detection objects
        """
        image = Image.open(io.BytesIO(image_bytes))
        return self.detect(image)

    def detect_from_base64(self, base64_string: str) -> List[Detection]:
        """
        Detect UI elements from base64-encoded image.

        Args:
            base64_string: Base64-encoded image data

        Returns:
            List of Detection objects
        """
        image_bytes = base64.b64decode(base64_string)
        return self.detect_from_bytes(image_bytes)

    def detect_from_file(self, file_path: str) -> List[Detection]:
        """
        Detect UI elements from an image file.

        Args:
            file_path: Path to image file

        Returns:
            List of Detection objects
        """
        image = Image.open(file_path)
        return self.detect(image)


# Singleton instance for reuse
_detector_instance: Optional[VisionDetector] = None


def get_detector(confidence_threshold: float = 0.15) -> VisionDetector:
    """Get or create the singleton VisionDetector instance"""
    global _detector_instance
    if _detector_instance is None:
        _detector_instance = VisionDetector(confidence_threshold=confidence_threshold)
    return _detector_instance
