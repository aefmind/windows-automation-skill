"""
RapidOCR - Fast ONNX-based OCR for UI text recognition
Alternative to Windows.Media.Ocr with ~5x better performance
"""

import numpy as np
from PIL import Image
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
import time


@dataclass
class TextRegion:
    """Represents a detected text region"""

    text: str
    x: int  # Center X coordinate
    y: int  # Center Y coordinate
    width: int
    height: int
    confidence: float = 1.0

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
            "text": self.text,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "confidence": self.confidence,
        }


class RapidOCREngine:
    """
    RapidOCR-based text recognition using ONNX Runtime.

    Much faster than Windows.Media.Ocr:
    - Windows OCR: ~630ms (via PowerShell subprocess)
    - RapidOCR: ~100-200ms (native Python, ONNX)

    Uses lightweight ONNX models (~15MB total):
    - Text detection model
    - Text recognition model
    """

    def __init__(self):
        """Initialize RapidOCR engine (lazy loads models on first use)"""
        self._engine = None
        self._load_time = 0

    @property
    def engine(self):
        """Lazy load RapidOCR engine"""
        if self._engine is None:
            t0 = time.perf_counter()
            from rapidocr_onnxruntime import RapidOCR

            # Initialize with optimized settings for UI text
            self._engine = RapidOCR()
            self._load_time = (time.perf_counter() - t0) * 1000
        return self._engine

    def recognize(self, image: Image.Image) -> List[TextRegion]:
        """
        Recognize text in an image.

        Args:
            image: PIL Image

        Returns:
            List of TextRegion objects with detected text and positions
        """
        # Convert PIL to numpy array (RapidOCR expects numpy)
        img_array = np.array(image)

        # RapidOCR expects BGR format (OpenCV style)
        if len(img_array.shape) == 3 and img_array.shape[2] == 3:
            # RGB to BGR
            img_array = img_array[:, :, ::-1]

        # Run OCR
        result, elapse = self.engine(img_array)

        # Parse results
        regions = []
        if result is not None:
            for detection in result:
                # detection format: [box_points, text, confidence]
                # box_points: [[x1,y1], [x2,y1], [x2,y2], [x1,y2]]
                box = detection[0]
                text = detection[1]
                confidence = detection[2] if len(detection) > 2 else 1.0

                # Calculate bounding box from polygon points
                x_coords = [p[0] for p in box]
                y_coords = [p[1] for p in box]

                left = int(min(x_coords))
                top = int(min(y_coords))
                right = int(max(x_coords))
                bottom = int(max(y_coords))

                width = right - left
                height = bottom - top
                center_x = left + width // 2
                center_y = top + height // 2

                regions.append(
                    TextRegion(
                        text=text,
                        x=center_x,
                        y=center_y,
                        width=width,
                        height=height,
                        confidence=float(confidence),
                    )
                )

        return regions

    def recognize_from_file(self, file_path: str) -> List[TextRegion]:
        """
        Recognize text from an image file.

        Args:
            file_path: Path to image file

        Returns:
            List of TextRegion objects
        """
        image = Image.open(file_path)
        return self.recognize(image)

    def find_text(
        self, image: Image.Image, search_text: str, case_sensitive: bool = False
    ) -> List[TextRegion]:
        """
        Find specific text in an image.

        Args:
            image: PIL Image
            search_text: Text to find
            case_sensitive: Whether to use case-sensitive matching

        Returns:
            List of TextRegion objects containing the search text
        """
        regions = self.recognize(image)

        if case_sensitive:
            return [r for r in regions if search_text in r.text]
        else:
            search_lower = search_text.lower()
            return [r for r in regions if search_lower in r.text.lower()]

    def get_full_text(self, image: Image.Image) -> str:
        """
        Get all text from an image as a single string.

        Args:
            image: PIL Image

        Returns:
            All recognized text concatenated with spaces
        """
        regions = self.recognize(image)
        return " ".join(r.text for r in regions)


# Singleton instance
_rapidocr_instance: Optional[RapidOCREngine] = None


def get_rapidocr() -> RapidOCREngine:
    """Get or create the singleton RapidOCR instance"""
    global _rapidocr_instance
    if _rapidocr_instance is None:
        _rapidocr_instance = RapidOCREngine()
    return _rapidocr_instance
