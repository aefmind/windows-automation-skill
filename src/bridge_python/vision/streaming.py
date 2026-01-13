"""
Vision Streaming - WebSocket-based real-time screenshot streaming
"""

import time
import logging
import threading
from PIL import ImageGrab
from .vision_service import VisionConfig, optimize_screenshot, get_screenshot_cache


class VisionStreamer:
    """
    Handles real-time screenshot streaming over WebSockets.
    """

    def __init__(self):
        self.active_streams = {}
        self._lock = threading.Lock()

    def stream_screenshots(
        self, ws, fps=5, quality=None, max_width=None, max_height=None
    ):
        """
        Stream screenshots over WebSocket.

        Args:
            ws: WebSocket connection
            fps: Frames per second
            quality: JPEG quality override
            max_width: Max width override
            max_height: Max height override
        """
        interval = 1.0 / max(1, min(30, fps))
        quality = quality or VisionConfig.jpeg_quality
        max_width = max_width or VisionConfig.max_width
        max_height = max_height or VisionConfig.max_height

        logging.info(f"Starting screenshot stream at {fps} FPS (quality={quality})")

        try:
            cache = get_screenshot_cache()

            while True:
                start_time = time.time()

                # Check cache first (reuse if possible)
                optimized = cache.get(
                    quality=quality, max_width=max_width, max_height=max_height
                )

                if not optimized:
                    # Capture new
                    screenshot = ImageGrab.grab()
                    optimized = optimize_screenshot(
                        screenshot,
                        max_width=max_width,
                        max_height=max_height,
                        jpeg_quality=quality,
                    )
                    # Cache it so other endpoints benefit
                    cache.put(
                        optimized,
                        quality=quality,
                        max_width=max_width,
                        max_height=max_height,
                    )

                # Send over websocket
                ws.send(optimized.data)

                # Sleep to maintain FPS
                elapsed = time.time() - start_time
                sleep_time = max(0, interval - elapsed)
                if sleep_time > 0:
                    time.sleep(sleep_time)

        except Exception as e:
            logging.info(f"Screenshot stream closed: {str(e)}")


_streamer = VisionStreamer()


def get_vision_streamer():
    return _streamer
