"""
Vision Streaming - WebSocket-based real-time screenshot streaming
"""

import time
import logging
from PIL import ImageGrab
from .vision_service import VisionConfig, optimize_screenshot, get_screenshot_cache


class VisionStreamer:
    """Handles real-time screenshot streaming over WebSockets."""

    def stream_screenshots(
        self, ws, fps=5, quality=None, max_width=None, max_height=None
    ):
        """
        Stream screenshots over WebSocket.

        Args:
            ws: WebSocket connection
            fps: Frames per second (1-30)
            quality: JPEG quality override
            max_width: Max width override
            max_height: Max height override
        """
        # Apply defaults and clamp FPS
        fps = max(1, min(30, fps))
        interval = 1.0 / fps
        quality = quality or VisionConfig.jpeg_quality
        max_width = max_width or VisionConfig.max_width
        max_height = max_height or VisionConfig.max_height
        cache = get_screenshot_cache()

        logging.info(f"Starting screenshot stream: {fps} FPS, quality={quality}")

        try:
            while True:
                frame_start = time.time()

                # Try cache first, capture if miss
                optimized = cache.get(
                    quality=quality, max_width=max_width, max_height=max_height
                )
                if not optimized:
                    screenshot = ImageGrab.grab()
                    optimized = optimize_screenshot(
                        screenshot,
                        max_width=max_width,
                        max_height=max_height,
                        jpeg_quality=quality,
                    )
                    cache.put(
                        optimized,
                        quality=quality,
                        max_width=max_width,
                        max_height=max_height,
                    )

                ws.send(optimized.data)

                # Maintain target FPS
                sleep_time = interval - (time.time() - frame_start)
                if sleep_time > 0:
                    time.sleep(sleep_time)

        except Exception as e:
            logging.info(f"Stream closed: {e}")


_streamer = VisionStreamer()


def get_vision_streamer():
    return _streamer
