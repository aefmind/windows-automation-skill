"""
OptimizedOCR - Faster OCR strategies for Windows Desktop Automation

Optimizations implemented:
1. WinRT native OCR (bypasses PowerShell subprocess - 3x faster)
2. Region-based OCR (only scan specific areas)
3. Caching with smart invalidation
4. Parallel OCR for multiple regions
5. JPEG format fallback for PowerShell (if WinRT unavailable)
"""

import subprocess
import json
import os
import tempfile
import hashlib
import time
from PIL import Image
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor
import threading

from .ocr import TextRegion, OCR_POWERSHELL_SCRIPT

# Try to import WinRT OCR engine (3x faster than PowerShell)
_WINRT_AVAILABLE = False
_winrt_ocr_instance = None

try:
    from .ocr_native import WinRTOCR

    _test_winrt = WinRTOCR()
    if _test_winrt.is_available():
        _WINRT_AVAILABLE = True
        _winrt_ocr_instance = _test_winrt
except ImportError:
    pass
except Exception:
    pass


class OptimizedOCR:
    """
    Optimized Windows.Media.Ocr wrapper with caching and region support.

    Performance improvements vs base WindowsOCR:
    - WinRT native: ~166ms (3.2x faster than PowerShell)
    - Region-based: Only OCR what you need
    - Caching: Skip repeated OCR calls (<1ms for cache hits)
    - Parallel: OCR multiple regions simultaneously
    """

    def __init__(self, cache_ttl: float = 2.0, prefer_winrt: bool = True):
        """
        Initialize OptimizedOCR.

        Args:
            cache_ttl: Cache time-to-live in seconds
            prefer_winrt: Use WinRT OCR if available (3x faster)
        """
        self._script_path: str = ""
        self._create_script()

        # WinRT OCR engine (preferred - 3x faster)
        self._use_winrt = prefer_winrt and _WINRT_AVAILABLE
        self._winrt_ocr = _winrt_ocr_instance if self._use_winrt else None

        # Cache for OCR results
        self._cache: Dict[str, Tuple[List[TextRegion], float]] = {}
        self._cache_ttl = cache_ttl
        self._cache_lock = threading.Lock()

        # Thread pool for parallel OCR
        self._executor = ThreadPoolExecutor(max_workers=4)

    def _create_script(self):
        """Create the PowerShell script file"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self._script_path = os.path.join(script_dir, "_ocr_script.ps1")

        with open(self._script_path, "w", encoding="utf-8") as f:
            f.write(OCR_POWERSHELL_SCRIPT)

    def _get_image_hash(self, image: Image.Image) -> str:
        """Get a quick hash of image for caching"""
        # Sample pixels for fast hash (not full image)
        width, height = image.size
        sample = image.crop((0, 0, min(100, width), min(100, height)))
        return hashlib.md5(sample.tobytes()).hexdigest()[:16]

    def _check_cache(self, cache_key: str) -> Optional[List[TextRegion]]:
        """Check if we have a valid cached result"""
        with self._cache_lock:
            if cache_key in self._cache:
                regions, timestamp = self._cache[cache_key]
                if time.time() - timestamp < self._cache_ttl:
                    return regions
                else:
                    del self._cache[cache_key]
        return None

    def _store_cache(self, cache_key: str, regions: List[TextRegion]):
        """Store result in cache"""
        with self._cache_lock:
            self._cache[cache_key] = (regions, time.time())

            # Limit cache size
            if len(self._cache) > 50:
                # Remove oldest entries
                sorted_keys = sorted(
                    self._cache.keys(), key=lambda k: self._cache[k][1]
                )
                for key in sorted_keys[:10]:
                    del self._cache[key]

    def recognize(self, image: Image.Image, use_cache: bool = True) -> List[TextRegion]:
        """
        Recognize text in an image with optimizations.

        Uses WinRT OCR (3x faster) if available, falls back to PowerShell.

        Args:
            image: PIL Image
            use_cache: Whether to use caching

        Returns:
            List of TextRegion objects
        """
        # Check cache first
        cache_key = None
        if use_cache:
            cache_key = self._get_image_hash(image)
            cached = self._check_cache(cache_key)
            if cached is not None:
                return cached

        # Use WinRT OCR if available (3x faster than PowerShell)
        if self._use_winrt and self._winrt_ocr:
            try:
                regions = self._winrt_ocr.recognize(image)
                if use_cache and cache_key:
                    self._store_cache(cache_key, regions)
                return regions
            except Exception:
                # Fall back to PowerShell on error
                pass

        # Fallback: PowerShell OCR
        return self._recognize_powershell(image, use_cache, cache_key)

    def _recognize_powershell(
        self,
        image: Image.Image,
        use_cache: bool = True,
        cache_key: Optional[str] = None,
    ) -> List[TextRegion]:
        """
        Recognize text using PowerShell subprocess (fallback method).
        """
        # Save as JPEG (faster than PNG)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
            temp_path = tmp.name
            if image.mode == "RGBA":
                image = image.convert("RGB")
            image.save(tmp, format="JPEG", quality=90)

        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    self._script_path,
                    temp_path,
                ],
                capture_output=True,
                timeout=30,
                encoding="utf-8",
                errors="replace",
            )

            if result.returncode != 0:
                return []

            output = result.stdout.strip()
            if not output or output == "[]":
                return []

            import re

            output = re.sub(r"[\x00-\x1f\x7f]", "", output)

            try:
                raw_results = json.loads(output)
            except json.JSONDecodeError:
                return []

            regions = []
            for r in raw_results:
                if isinstance(r, str):
                    r = raw_results
                    regions.append(
                        TextRegion(
                            text=r.get("text", ""),
                            x=r["left"] + r["width"] // 2,
                            y=r["top"] + r["height"] // 2,
                            width=r["width"],
                            height=r["height"],
                        )
                    )
                    break
                else:
                    regions.append(
                        TextRegion(
                            text=r.get("text", ""),
                            x=r["left"] + r["width"] // 2,
                            y=r["top"] + r["height"] // 2,
                            width=r["width"],
                            height=r["height"],
                        )
                    )

            # Store in cache
            if use_cache and cache_key:
                self._store_cache(cache_key, regions)

            return regions

        finally:
            try:
                os.unlink(temp_path)
            except:
                pass

    def recognize_region(
        self,
        image: Image.Image,
        x: int,
        y: int,
        width: int,
        height: int,
    ) -> List[TextRegion]:
        """
        Recognize text in a specific region (much faster for small areas).

        Args:
            image: Full PIL Image
            x, y: Top-left corner of region
            width, height: Size of region

        Returns:
            List of TextRegion objects with adjusted coordinates
        """
        # Crop to region
        region = image.crop((x, y, x + width, y + height))

        # OCR the region
        regions = self.recognize(region, use_cache=True)

        # Adjust coordinates back to full image space
        for r in regions:
            r.x += x
            r.y += y

        return regions

    def recognize_regions_parallel(
        self,
        image: Image.Image,
        regions: List[Tuple[int, int, int, int]],
    ) -> List[TextRegion]:
        """
        OCR multiple regions in parallel.

        Args:
            image: Full PIL Image
            regions: List of (x, y, width, height) tuples

        Returns:
            Combined list of TextRegion objects
        """

        def ocr_region(region_coords):
            x, y, w, h = region_coords
            return self.recognize_region(image, x, y, w, h)

        # Submit all regions for parallel processing
        futures = [self._executor.submit(ocr_region, r) for r in regions]

        # Collect results
        all_regions = []
        for future in futures:
            try:
                all_regions.extend(future.result(timeout=30))
            except Exception:
                pass

        return all_regions

    def find_text_fast(
        self,
        image: Image.Image,
        search_text: str,
        hint_regions: Optional[List[Tuple[int, int, int, int]]] = None,
    ) -> Optional[TextRegion]:
        """
        Fast text search with optional region hints.

        If hint_regions provided, only searches those areas first.
        Falls back to full image if not found.

        Args:
            image: PIL Image
            search_text: Text to find
            hint_regions: Optional list of (x, y, width, height) to search first

        Returns:
            TextRegion if found, None otherwise
        """
        search_lower = search_text.lower()

        # Try hint regions first
        if hint_regions:
            for x, y, w, h in hint_regions:
                regions = self.recognize_region(image, x, y, w, h)
                for r in regions:
                    if search_lower in r.text.lower():
                        return r

        # Fall back to full image
        regions = self.recognize(image)
        for r in regions:
            if search_lower in r.text.lower():
                return r

        return None

    def clear_cache(self):
        """Clear all cached results"""
        with self._cache_lock:
            self._cache.clear()

    def get_cache_stats(self) -> Dict:
        """Get cache statistics"""
        with self._cache_lock:
            return {
                "entries": len(self._cache),
                "ttl_seconds": self._cache_ttl,
                "engine": "winrt" if self._use_winrt else "powershell",
                "winrt_available": _WINRT_AVAILABLE,
            }


# Singleton instance
_optimized_ocr_instance: Optional[OptimizedOCR] = None


def get_optimized_ocr() -> OptimizedOCR:
    """Get or create the singleton OptimizedOCR instance"""
    global _optimized_ocr_instance
    if _optimized_ocr_instance is None:
        _optimized_ocr_instance = OptimizedOCR()
    return _optimized_ocr_instance
