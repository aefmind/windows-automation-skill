"""
WindowsOCR - Windows.Media.Ocr wrapper for text recognition
Uses Windows 10/11 built-in OCR capabilities via subprocess
"""

import subprocess
import json
import os
import tempfile
from PIL import Image
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
import io


@dataclass
class TextRegion:
    """Represents a detected text region"""

    text: str
    x: int  # Center X coordinate
    y: int  # Center Y coordinate
    width: int
    height: int
    confidence: float = 1.0  # Windows OCR doesn't provide confidence

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


# PowerShell script for Windows.Media.Ocr
OCR_POWERSHELL_SCRIPT = """
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.RandomAccessStream, Windows.Foundation, ContentType = WindowsRuntime]

function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | 
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and 
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

$imagePath = $args[0]
$stream = [System.IO.File]::OpenRead($imagePath)
$randomAccessStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($stream)

$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($randomAccessStream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$softwareBitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($ocrEngine -eq $null) {
    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage("en-US")
}

$ocrResult = Await ($ocrEngine.RecognizeAsync($softwareBitmap)) ([Windows.Media.Ocr.OcrResult])

$results = @()
foreach ($line in $ocrResult.Lines) {
    foreach ($word in $line.Words) {
        $rect = $word.BoundingRect
        $results += @{
            text = $word.Text
            left = [int]$rect.X
            top = [int]$rect.Y
            width = [int]$rect.Width
            height = [int]$rect.Height
        }
    }
}

$stream.Close()
ConvertTo-Json -InputObject $results -Compress
"""


class WindowsOCR:
    """
    Windows.Media.Ocr based text recognition.

    Uses the built-in Windows 10/11 OCR engine via PowerShell.
    This has zero additional dependencies and excellent accuracy.
    """

    def __init__(self):
        """Initialize WindowsOCR"""
        self._script_path: str = ""
        self._create_script()

    def _create_script(self):
        """Create the PowerShell script file"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self._script_path = os.path.join(script_dir, "_ocr_script.ps1")

        with open(self._script_path, "w", encoding="utf-8") as f:
            f.write(OCR_POWERSHELL_SCRIPT)

    def recognize(self, image: Image.Image) -> List[TextRegion]:
        """
        Recognize text in an image.

        Args:
            image: PIL Image

        Returns:
            List of TextRegion objects with detected text and positions
        """
        # Save image to temp file - use JPEG for ~60ms faster save
        # JPEG is ~5x faster to save than PNG with negligible quality loss for OCR
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
            temp_path = tmp.name
            # Convert to RGB if needed (JPEG doesn't support RGBA)
            if image.mode == "RGBA":
                image = image.convert("RGB")
            image.save(tmp, format="JPEG", quality=90)

        try:
            # Run PowerShell OCR with UTF-8 encoding
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
                errors="replace",  # Replace undecodable chars with ?
            )

            if result.returncode != 0:
                stderr = result.stderr or ""
                raise RuntimeError(f"OCR failed: {stderr}")

            # Parse JSON output
            stdout = result.stdout or ""
            output = stdout.strip()
            if not output or output == "[]":
                return []

            # Sanitize output - remove any control characters that might have slipped through
            import re

            output = re.sub(r"[\x00-\x1f\x7f]", "", output)

            try:
                raw_results = json.loads(output)
            except json.JSONDecodeError as e:
                # Log the problematic output for debugging
                import logging

                logging.warning(f"OCR JSON parse error: {e}")
                logging.debug(f"Problematic output (first 500 chars): {output[:500]}")
                return []

            # Convert to TextRegion objects
            regions = []
            for r in raw_results:
                # Handle single result (not array)
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

            return regions

        finally:
            # Clean up temp file
            try:
                os.unlink(temp_path)
            except:
                pass

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
_ocr_instance: Optional[WindowsOCR] = None


def get_ocr() -> WindowsOCR:
    """Get or create the singleton WindowsOCR instance"""
    global _ocr_instance
    if _ocr_instance is None:
        _ocr_instance = WindowsOCR()
    return _ocr_instance
