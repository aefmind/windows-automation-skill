"""
VisionService - Combined detection + OCR service for UI automation
Provides high-level API for finding and interacting with UI elements

v4.1 - Now uses OptimizedOCR with caching and region-based search
Performance: Full OCR ~530ms, Region OCR ~120ms, Cached <1ms

v4.2 - Agent Vision Mode
- Supports 'local', 'agent', and 'auto' vision modes
- Agent mode returns optimized screenshots for AI agent vision
- Auto mode tries local first, falls back to agent mode on failure

v4.3 - Screenshot Caching
- ScreenshotCache class with TTL-based expiration
- Reduces redundant screen captures for rapid successive requests
- Thread-safe with configurable TTL and max entries
"""

from PIL import Image
from typing import List, Dict, Tuple, Optional, Union
from dataclasses import dataclass, field
import io
import base64
import os
import time
import hashlib
import threading

from .detector import VisionDetector, Detection, get_detector
from .ocr import WindowsOCR, TextRegion, get_ocr
from .ocr_optimized import OptimizedOCR, get_optimized_ocr


# =========================================================================
# SCREENSHOT CACHE (v4.3)
# =========================================================================


@dataclass
class CachedScreenshot:
    """A cached screenshot entry with TTL support."""

    data: "OptimizedScreenshot"
    timestamp: float
    cache_key: str
    hits: int = 0


class ScreenshotCache:
    """
    LRU cache for screenshots with TTL expiration.

    Reduces redundant screen captures when multiple operations happen
    within a short time window (e.g., find element + click).

    Thread-safe for concurrent access.
    """

    def __init__(
        self,
        ttl_seconds: float = 2.0,
        max_entries: int = 5,
    ):
        """
        Initialize screenshot cache.

        Args:
            ttl_seconds: Time-to-live for cached screenshots (default 2s)
            max_entries: Maximum number of cached screenshots (default 5)
        """
        self.ttl_seconds = ttl_seconds
        self.max_entries = max_entries
        self._cache: Dict[str, CachedScreenshot] = {}
        self._lock = threading.RLock()
        self._stats = {
            "hits": 0,
            "misses": 0,
            "evictions": 0,
            "stores": 0,
        }

    def _make_key(
        self,
        region: Optional[Tuple[int, int, int, int]] = None,
        quality: int = 75,
        max_width: int = 1920,
        max_height: int = 1080,
    ) -> str:
        """
        Generate cache key for a screenshot request.

        Args:
            region: Optional (x, y, width, height) tuple for region captures
            quality: JPEG quality setting
            max_width: Maximum width after resize
            max_height: Maximum height after resize

        Returns:
            String cache key
        """
        if region:
            key_data = (
                f"region:{region[0]},{region[1]},{region[2]},{region[3]}:q{quality}"
            )
        else:
            key_data = f"full:w{max_width}:h{max_height}:q{quality}"
        return hashlib.md5(key_data.encode()).hexdigest()[:16]

    def get(
        self,
        region: Optional[Tuple[int, int, int, int]] = None,
        quality: int = 75,
        max_width: int = 1920,
        max_height: int = 1080,
    ) -> Optional["OptimizedScreenshot"]:
        """
        Get a cached screenshot if available and not expired.

        Args:
            region: Optional region tuple
            quality: JPEG quality
            max_width: Max width
            max_height: Max height

        Returns:
            OptimizedScreenshot if cache hit, None otherwise
        """
        key = self._make_key(region, quality, max_width, max_height)
        now = time.time()

        with self._lock:
            if key in self._cache:
                entry = self._cache[key]
                age = now - entry.timestamp

                if age <= self.ttl_seconds:
                    # Cache hit
                    entry.hits += 1
                    self._stats["hits"] += 1
                    return entry.data
                else:
                    # Expired - remove and return miss
                    del self._cache[key]
                    self._stats["evictions"] += 1

            self._stats["misses"] += 1
            return None

    def put(
        self,
        screenshot: "OptimizedScreenshot",
        region: Optional[Tuple[int, int, int, int]] = None,
        quality: int = 75,
        max_width: int = 1920,
        max_height: int = 1080,
    ) -> str:
        """
        Store a screenshot in the cache.

        Args:
            screenshot: OptimizedScreenshot to cache
            region: Optional region tuple
            quality: JPEG quality
            max_width: Max width
            max_height: Max height

        Returns:
            Cache key used for storage
        """
        key = self._make_key(region, quality, max_width, max_height)

        with self._lock:
            # Evict oldest entries if at capacity
            while len(self._cache) >= self.max_entries:
                oldest_key = min(
                    self._cache.keys(), key=lambda k: self._cache[k].timestamp
                )
                del self._cache[oldest_key]
                self._stats["evictions"] += 1

            # Store new entry
            self._cache[key] = CachedScreenshot(
                data=screenshot,
                timestamp=time.time(),
                cache_key=key,
            )
            self._stats["stores"] += 1

        return key

    def clear(self):
        """Clear all cached screenshots."""
        with self._lock:
            count = len(self._cache)
            self._cache.clear()
            self._stats["evictions"] += count

    def get_stats(self) -> Dict:
        """Get cache statistics."""
        with self._lock:
            total_requests = self._stats["hits"] + self._stats["misses"]
            hit_rate = (
                self._stats["hits"] / total_requests if total_requests > 0 else 0.0
            )
            return {
                "entries": len(self._cache),
                "max_entries": self.max_entries,
                "ttl_seconds": self.ttl_seconds,
                "hits": self._stats["hits"],
                "misses": self._stats["misses"],
                "stores": self._stats["stores"],
                "evictions": self._stats["evictions"],
                "hit_rate": round(hit_rate, 3),
            }

    def set_ttl(self, ttl_seconds: float):
        """Update TTL setting."""
        with self._lock:
            self.ttl_seconds = max(0.1, ttl_seconds)

    def set_max_entries(self, max_entries: int):
        """Update max entries setting."""
        with self._lock:
            self.max_entries = max(1, max_entries)
            # Evict if over new limit
            while len(self._cache) > self.max_entries:
                oldest_key = min(
                    self._cache.keys(), key=lambda k: self._cache[k].timestamp
                )
                del self._cache[oldest_key]
                self._stats["evictions"] += 1


# Global screenshot cache singleton
_screenshot_cache: Optional[ScreenshotCache] = None


def get_screenshot_cache() -> ScreenshotCache:
    """Get or create the singleton ScreenshotCache instance."""
    global _screenshot_cache
    if _screenshot_cache is None:
        ttl = float(os.environ.get("SCREENSHOT_CACHE_TTL", "2.0"))
        max_entries = int(os.environ.get("SCREENSHOT_CACHE_MAX_ENTRIES", "5"))
        _screenshot_cache = ScreenshotCache(ttl_seconds=ttl, max_entries=max_entries)
    return _screenshot_cache


# =========================================================================
# VISION MODE CONFIGURATION
# =========================================================================


class VisionConfig:
    """
    Global vision mode configuration.

    Modes:
    - 'local': Use OmniParser + WinRT OCR (default, full processing)
    - 'agent': Skip local processing, return optimized screenshots for AI agent
    - 'auto': Try local first, fall back to agent mode on failure/empty results
    """

    # Vision mode: 'local', 'agent', or 'auto'
    mode: str = os.environ.get("VISION_MODE", "auto")

    # Screenshot compression settings
    jpeg_quality: int = int(os.environ.get("VISION_JPEG_QUALITY", "75"))
    max_width: int = int(os.environ.get("VISION_MAX_WIDTH", "1920"))
    max_height: int = int(os.environ.get("VISION_MAX_HEIGHT", "1080"))

    # Whether to include thumbnail in responses
    include_thumbnail: bool = (
        os.environ.get("VISION_INCLUDE_THUMBNAIL", "false").lower() == "true"
    )
    thumbnail_max_size: int = int(os.environ.get("VISION_THUMBNAIL_SIZE", "400"))

    @classmethod
    def set_mode(cls, mode: str):
        """Set vision mode ('local', 'agent', or 'auto')"""
        if mode not in ("local", "agent", "auto"):
            raise ValueError(
                f"Invalid vision mode: {mode}. Must be 'local', 'agent', or 'auto'"
            )
        cls.mode = mode

    @classmethod
    def get_config(cls) -> Dict:
        """Get current configuration as dict"""
        return {
            "mode": cls.mode,
            "jpeg_quality": cls.jpeg_quality,
            "max_width": cls.max_width,
            "max_height": cls.max_height,
            "include_thumbnail": cls.include_thumbnail,
            "thumbnail_max_size": cls.thumbnail_max_size,
        }

    @classmethod
    def update(cls, **kwargs):
        """Update configuration settings"""
        if "mode" in kwargs:
            cls.set_mode(kwargs["mode"])
        if "jpeg_quality" in kwargs:
            cls.jpeg_quality = max(1, min(100, int(kwargs["jpeg_quality"])))
        if "max_width" in kwargs:
            cls.max_width = max(100, int(kwargs["max_width"]))
        if "max_height" in kwargs:
            cls.max_height = max(100, int(kwargs["max_height"]))
        if "include_thumbnail" in kwargs:
            cls.include_thumbnail = bool(kwargs["include_thumbnail"])
        if "thumbnail_max_size" in kwargs:
            cls.thumbnail_max_size = max(50, int(kwargs["thumbnail_max_size"]))


# =========================================================================
# SCREENSHOT OPTIMIZATION HELPERS
# =========================================================================


@dataclass
class OptimizedScreenshot:
    """
    Optimized screenshot data for agent vision mode.
    Contains base64-encoded image with metadata.
    """

    data: str  # base64-encoded JPEG
    width: int
    height: int
    original_width: int
    original_height: int
    format: str  # 'jpeg'
    quality: int
    size_bytes: int
    compression_ratio: float
    timestamp: float
    thumbnail: Optional[str] = None  # Optional smaller preview

    def to_dict(self) -> Dict:
        """Convert to dictionary for JSON serialization"""
        result = {
            "data": self.data,
            "width": self.width,
            "height": self.height,
            "original_width": self.original_width,
            "original_height": self.original_height,
            "format": self.format,
            "quality": self.quality,
            "size_bytes": self.size_bytes,
            "compression_ratio": round(self.compression_ratio, 2),
            "timestamp": self.timestamp,
        }
        if self.thumbnail:
            result["thumbnail"] = self.thumbnail
        return result


def optimize_screenshot(
    image: Image.Image,
    max_width: Optional[int] = None,
    max_height: Optional[int] = None,
    jpeg_quality: Optional[int] = None,
    include_thumbnail: bool = False,
    thumbnail_max_size: int = 400,
) -> OptimizedScreenshot:
    """
    Optimize a screenshot for agent vision mode.

    Performs:
    - Resize if larger than max dimensions (maintains aspect ratio)
    - JPEG compression with configurable quality
    - Optional thumbnail generation

    Args:
        image: PIL Image to optimize
        max_width: Maximum width (default from VisionConfig)
        max_height: Maximum height (default from VisionConfig)
        jpeg_quality: JPEG quality 1-100 (default from VisionConfig)
        include_thumbnail: Generate smaller thumbnail
        thumbnail_max_size: Max dimension for thumbnail

    Returns:
        OptimizedScreenshot with base64 data and metadata
    """
    # Use defaults from config
    max_width = max_width or VisionConfig.max_width
    max_height = max_height or VisionConfig.max_height
    jpeg_quality = jpeg_quality or VisionConfig.jpeg_quality

    original_width, original_height = image.size

    # Calculate original uncompressed size (RGB)
    original_size = original_width * original_height * 3

    # Resize if needed (maintain aspect ratio)
    resized = image.copy()
    if original_width > max_width or original_height > max_height:
        resized.thumbnail((max_width, max_height), Image.LANCZOS)

    # Convert to RGB if needed (JPEG doesn't support alpha)
    if resized.mode in ("RGBA", "P"):
        resized = resized.convert("RGB")

    # Compress to JPEG
    buffer = io.BytesIO()
    resized.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True)
    compressed_data = buffer.getvalue()

    # Base64 encode
    b64_data = base64.b64encode(compressed_data).decode("ascii")

    # Generate thumbnail if requested
    thumbnail_b64 = None
    if include_thumbnail:
        thumb = image.copy()
        thumb.thumbnail((thumbnail_max_size, thumbnail_max_size), Image.LANCZOS)
        if thumb.mode in ("RGBA", "P"):
            thumb = thumb.convert("RGB")
        thumb_buffer = io.BytesIO()
        thumb.save(thumb_buffer, format="JPEG", quality=60, optimize=True)
        thumbnail_b64 = base64.b64encode(thumb_buffer.getvalue()).decode("ascii")

    return OptimizedScreenshot(
        data=b64_data,
        width=resized.size[0],
        height=resized.size[1],
        original_width=original_width,
        original_height=original_height,
        format="jpeg",
        quality=jpeg_quality,
        size_bytes=len(compressed_data),
        compression_ratio=original_size / len(compressed_data)
        if len(compressed_data) > 0
        else 0,
        timestamp=time.time(),
        thumbnail=thumbnail_b64,
    )


def optimize_region(
    image: Image.Image,
    x: int,
    y: int,
    width: int,
    height: int,
    jpeg_quality: Optional[int] = None,
) -> OptimizedScreenshot:
    """
    Extract and optimize a region of a screenshot.

    More token-efficient than full screenshots when you know where to look.

    Args:
        image: Full PIL Image
        x, y: Top-left corner of region
        width, height: Size of region
        jpeg_quality: JPEG quality 1-100

    Returns:
        OptimizedScreenshot of the cropped region
    """
    # Clamp to image bounds
    img_width, img_height = image.size
    x = max(0, min(x, img_width - 1))
    y = max(0, min(y, img_height - 1))
    right = min(x + width, img_width)
    bottom = min(y + height, img_height)

    # Crop region
    region = image.crop((x, y, right, bottom))

    # Optimize (no resize for regions - they're already targeted)
    jpeg_quality = jpeg_quality or VisionConfig.jpeg_quality

    if region.mode in ("RGBA", "P"):
        region = region.convert("RGB")

    buffer = io.BytesIO()
    region.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True)
    compressed_data = buffer.getvalue()

    b64_data = base64.b64encode(compressed_data).decode("ascii")

    region_width, region_height = region.size
    original_size = region_width * region_height * 3

    return OptimizedScreenshot(
        data=b64_data,
        width=region_width,
        height=region_height,
        original_width=img_width,
        original_height=img_height,
        format="jpeg",
        quality=jpeg_quality,
        size_bytes=len(compressed_data),
        compression_ratio=original_size / len(compressed_data)
        if len(compressed_data) > 0
        else 0,
        timestamp=time.time(),
        thumbnail=None,
    )


@dataclass
class UIElement:
    """
    Represents a detected UI element with optional text.
    Combines detection bounding box with OCR text if available.
    """

    x: int
    y: int
    width: int
    height: int
    confidence: float
    text: Optional[str] = None
    element_type: str = "unknown"  # 'button', 'text', 'icon', etc.

    @property
    def bounds(self) -> Tuple[int, int, int, int]:
        """Returns (left, top, right, bottom)"""
        half_w = self.width // 2
        half_h = self.height // 2
        return (self.x - half_w, self.y - half_h, self.x + half_w, self.y + half_h)

    @property
    def center(self) -> Tuple[int, int]:
        """Returns (x, y) center point for clicking"""
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
            "text": self.text,
            "type": self.element_type,
        }


class VisionService:
    """
    High-level service combining OmniParser detection with Windows OCR.

    Provides:
    - Element detection (visual recognition of UI elements)
    - Text recognition (OCR)
    - Combined find_element_by_text (detect + OCR correlation)
    - Tiling support for high-resolution displays
    """

    # Only tiling is used now (no global resize)
    TILE_SIZE = 640  # OmniParser optimal input size
    TILE_OVERLAP = 64  # Overlap between tiles to catch elements on boundaries

    def __init__(
        self,
        detector: Optional[VisionDetector] = None,
        ocr: Optional[WindowsOCR] = None,
        confidence_threshold: float = 0.15,
        use_optimized_ocr: bool = True,
    ):
        """
        Initialize VisionService.

        Args:
            detector: VisionDetector instance (uses singleton if None)
            ocr: WindowsOCR instance (uses singleton if None)
            confidence_threshold: Minimum confidence for detections
            use_optimized_ocr: Use OptimizedOCR with caching (default True)
        """
        self.detector = detector or get_detector(confidence_threshold)
        self.ocr = ocr or get_ocr()
        self.optimized_ocr = get_optimized_ocr() if use_optimized_ocr else None
        self.confidence_threshold = confidence_threshold
        self._use_optimized = use_optimized_ocr

    def detect_elements(
        self, image: Image.Image, use_tiling: bool = True
    ) -> List[UIElement]:
        """
        Detect all UI elements in an image using tiling only (no global resize).

        Args:
            image: PIL Image
            use_tiling: unused (kept for signature compatibility)

        Returns:
            List of UIElement objects
        """
        # Always use tiling to avoid downscaling before OmniParser
        detections = self._detect_with_tiling(image)

        # Convert Detection to UIElement
        return [
            UIElement(
                x=d.x,
                y=d.y,
                width=d.width,
                height=d.height,
                confidence=d.confidence,
                element_type="icon",
            )
            for d in detections
        ]

    def _detect_with_tiling(self, image: Image.Image) -> List[Detection]:
        """
        Detect elements using tiling for large images.

        Divides the image into overlapping tiles, processes each,
        and merges results with deduplication.
        """
        width, height = image.size
        all_detections = []

        # Calculate tile positions
        step = self.TILE_SIZE - self.TILE_OVERLAP

        for y in range(0, height, step):
            for x in range(0, width, step):
                # Calculate tile bounds
                tile_left = x
                tile_top = y
                tile_right = min(x + self.TILE_SIZE, width)
                tile_bottom = min(y + self.TILE_SIZE, height)

                # Skip tiny edge tiles
                if tile_right - tile_left < 100 or tile_bottom - tile_top < 100:
                    continue

                # Extract tile
                tile = image.crop((tile_left, tile_top, tile_right, tile_bottom))

                # Detect in tile
                tile_detections = self.detector.detect(tile)

                # Offset coordinates to original image space
                for d in tile_detections:
                    d.x += tile_left
                    d.y += tile_top
                    all_detections.append(d)

        # Deduplicate overlapping detections from different tiles
        return self._deduplicate_detections(all_detections)

    def _deduplicate_detections(
        self, detections: List[Detection], iou_threshold: float = 0.5
    ) -> List[Detection]:
        """Remove duplicate detections from overlapping tiles"""
        if len(detections) <= 1:
            return detections

        # Sort by confidence
        detections.sort(key=lambda d: d.confidence, reverse=True)

        kept = []
        suppressed = set()

        for i, det in enumerate(detections):
            if i in suppressed:
                continue

            kept.append(det)

            for j in range(i + 1, len(detections)):
                if j in suppressed:
                    continue

                iou = self.detector._calculate_iou(det, detections[j])
                if iou > iou_threshold:
                    suppressed.add(j)

        return kept

    def recognize_text(self, image: Image.Image) -> List[TextRegion]:
        """
        Recognize all text in an image.

        Args:
            image: PIL Image

        Returns:
            List of TextRegion objects with text and positions
        """
        return self.ocr.recognize(image)

    def find_element_by_text(
        self,
        image: Image.Image,
        text: str,
        case_sensitive: bool = False,
        fuzzy: bool = False,
    ) -> Optional[UIElement]:
        """
        Find a UI element containing specific text.

        This is the key method for Flutter/Electron apps:
        1. Uses OCR to find text location
        2. Correlates with detected UI elements
        3. Returns the best matching element

        Args:
            image: PIL Image
            text: Text to search for
            case_sensitive: Whether to use case-sensitive matching
            fuzzy: Allow partial/fuzzy matching

        Returns:
            UIElement if found, None otherwise
        """
        # First, find text regions
        text_regions = self.ocr.find_text(image, text, case_sensitive)

        if not text_regions:
            return None

        # For each text region, check if it overlaps with a detected element
        detections = self.detect_elements(image)

        best_match: Optional[UIElement] = None
        best_score = 0.0

        for text_region in text_regions:
            # Check direct OCR result first
            if best_match is None:
                # Use OCR bounds directly as fallback
                best_match = UIElement(
                    x=text_region.x,
                    y=text_region.y,
                    width=text_region.width,
                    height=text_region.height,
                    confidence=text_region.confidence,
                    text=text_region.text,
                    element_type="text",
                )

            # Try to find containing detection (like a button around the text)
            for det in detections:
                overlap = self._calculate_overlap(text_region, det)

                # Text is inside detection - use detection for element type/confidence
                # but KEEP the text region's center for clicking accuracy
                if overlap > 0.5:
                    score = det.confidence * overlap
                    if score > best_score:
                        best_score = score
                        # Use TEXT REGION center for accurate clicking,
                        # but detection's confidence and type info
                        best_match = UIElement(
                            x=text_region.x,  # Use OCR center, not detection center
                            y=text_region.y,
                            width=text_region.width,
                            height=text_region.height,
                            confidence=det.confidence,
                            text=text_region.text,
                            element_type="button",  # Inside a detection = button
                        )

        return best_match

    def _calculate_overlap(
        self, text_region: TextRegion, detection: Union[Detection, UIElement]
    ) -> float:
        """Calculate how much of the text region is inside the detection"""
        tl, tt, tr, tb = text_region.bounds
        dl, dt, dr, db = detection.bounds

        # Calculate intersection
        left = max(tl, dl)
        top = max(tt, dt)
        right = min(tr, dr)
        bottom = min(tb, db)

        if left >= right or top >= bottom:
            return 0.0

        intersection = (right - left) * (bottom - top)
        text_area = (tr - tl) * (tb - tt)

        return intersection / text_area if text_area > 0 else 0.0

    def get_clickable_point(
        self, image: Image.Image, text: str, case_sensitive: bool = False
    ) -> Optional[Tuple[int, int]]:
        """
        Get the click coordinates for a UI element with specific text.

        Args:
            image: PIL Image
            text: Text to search for
            case_sensitive: Whether to use case-sensitive matching

        Returns:
            (x, y) click coordinates, or None if not found
        """
        element = self.find_element_by_text(image, text, case_sensitive)
        if element:
            return element.center
        return None

    # =========================================================================
    # OPTIMIZED OCR METHODS (v4.1) - Use these for better performance
    # =========================================================================

    def recognize_text_cached(
        self, image: Image.Image, use_cache: bool = True
    ) -> List[TextRegion]:
        """
        Recognize text with caching support.

        Performance: ~530ms first call, <1ms cached

        Args:
            image: PIL Image
            use_cache: Use cache for repeated queries (default True)

        Returns:
            List of TextRegion objects
        """
        if self.optimized_ocr:
            return self.optimized_ocr.recognize(image, use_cache=use_cache)
        return self.ocr.recognize(image)

    def recognize_text_region(
        self, image: Image.Image, x: int, y: int, width: int, height: int
    ) -> List[TextRegion]:
        """
        Recognize text in a specific region only (4x faster).

        Performance: ~120ms vs ~530ms for full screen

        Args:
            image: PIL Image
            x, y: Top-left corner of region
            width, height: Size of region

        Returns:
            List of TextRegion objects with absolute coordinates
        """
        if self.optimized_ocr:
            return self.optimized_ocr.recognize_region(image, x, y, width, height)

        # Fallback: crop and OCR
        region = image.crop((x, y, x + width, y + height))
        regions = self.ocr.recognize(region)
        for r in regions:
            r.x += x
            r.y += y
        return regions

    def find_text_fast(
        self,
        image: Image.Image,
        text: str,
        hint_regions: Optional[List[Tuple[int, int, int, int]]] = None,
    ) -> Optional[TextRegion]:
        """
        Fast text search with optional region hints.

        If hint_regions are provided, searches those areas first (4x faster).
        Falls back to full image search if not found in hints.

        Args:
            image: PIL Image
            text: Text to find
            hint_regions: Optional list of (x, y, width, height) to search first

        Returns:
            TextRegion if found, None otherwise
        """
        if self.optimized_ocr:
            return self.optimized_ocr.find_text_fast(image, text, hint_regions)

        # Fallback to standard find
        regions = self.ocr.find_text(image, text)
        return regions[0] if regions else None

    def get_ocr_cache_stats(self) -> Dict:
        """Get OCR cache statistics"""
        if self.optimized_ocr:
            return self.optimized_ocr.get_cache_stats()
        return {"entries": 0, "ttl_seconds": 0, "optimized": False}

    def clear_ocr_cache(self):
        """Clear OCR cache"""
        if self.optimized_ocr:
            self.optimized_ocr.clear_cache()

    def analyze_screen(self, image: Image.Image, use_cache: bool = True) -> Dict:
        """
        Perform full analysis of a screen.

        Args:
            image: PIL Image
            use_cache: Use OCR cache (default True)

        Returns:
            Dict with elements, text, and summary
        """
        elements = self.detect_elements(image)
        text_regions = self.recognize_text_cached(image, use_cache=use_cache)

        result = {
            "width": image.size[0],
            "height": image.size[1],
            "element_count": len(elements),
            "text_count": len(text_regions),
            "elements": [e.to_dict() for e in elements],
            "text": [t.to_dict() for t in text_regions],
            "full_text": " ".join(t.text for t in text_regions),
        }

        # Add cache stats if optimized OCR
        if self.optimized_ocr:
            result["ocr_cache"] = self.get_ocr_cache_stats()

        return result


# Singleton instance
_service_instance: Optional[VisionService] = None


def get_vision_service(confidence_threshold: float = 0.15) -> VisionService:
    """Get or create the singleton VisionService instance"""
    global _service_instance
    if _service_instance is None:
        _service_instance = VisionService(confidence_threshold=confidence_threshold)
    return _service_instance
