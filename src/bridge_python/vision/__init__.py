"""
Vision module for Windows Desktop Automation v4.0
Provides AI-powered UI element detection using OmniParser + Windows OCR
"""

from .detector import VisionDetector
from .ocr import WindowsOCR
from .vision_service import VisionService

__all__ = ["VisionDetector", "WindowsOCR", "VisionService"]
__version__ = "4.0.0"
