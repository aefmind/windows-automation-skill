"""
Context Manager for Windows Desktop Automation v4.0

Provides intelligent caching and response compression to reduce token consumption
for AI agents. Acts as a proxy layer that:
1. Caches exploration results with window-aware invalidation
2. Compresses verbose responses to essential information
3. Implements progressive disclosure (minimal context first)
4. Tracks window state changes for cache invalidation

Author: AI Engineering Team
Version: 4.0
"""

import hashlib
import json
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from collections import OrderedDict


@dataclass
class CacheEntry:
    """A cached response with metadata for invalidation."""

    key: str
    value: Any
    window_hash: str
    created_at: float
    ttl_seconds: float
    access_count: int = 0
    last_accessed: float = field(default_factory=time.time)

    def is_expired(self) -> bool:
        return time.time() - self.created_at > self.ttl_seconds

    def touch(self):
        self.access_count += 1
        self.last_accessed = time.time()


class SemanticCache:
    """
    Cache with TTL and window-aware invalidation.

    Features:
    - TTL-based expiration
    - Window hash invalidation (cache invalidates when UI changes)
    - LRU eviction for memory management
    - Semantic key normalization
    """

    DEFAULT_TTL = 30.0  # 30 seconds default TTL
    MAX_ENTRIES = 100  # Maximum cache entries

    # TTL by command type
    TTL_BY_COMMAND = {
        "explore": 60.0,  # Windows list changes slowly
        "explore_window": 30.0,  # Window contents change moderately
        "get_window_summary": 30.0,
        "get_interactive_elements": 20.0,
        "find_element": 10.0,  # Element positions may change
        "element_exists": 5.0,  # Quick checks, low TTL
        "get_element_brief": 10.0,
    }

    def __init__(self, max_entries: int = MAX_ENTRIES):
        self.max_entries = max_entries
        self._cache: OrderedDict[str, CacheEntry] = OrderedDict()
        self._window_hashes: Dict[str, str] = {}  # window_selector -> hash
        self._stats = {"hits": 0, "misses": 0, "evictions": 0, "invalidations": 0}

    def _normalize_key(self, action: str, params: Dict[str, Any]) -> str:
        """Create a normalized cache key from action and params."""
        # Sort params for consistent hashing
        sorted_params = json.dumps(params, sort_keys=True)
        key_str = f"{action}:{sorted_params}"
        return hashlib.md5(key_str.encode()).hexdigest()

    def _get_window_hash(self, window_selector: Optional[str]) -> str:
        """Get or compute window state hash."""
        if not window_selector:
            return "desktop"
        return self._window_hashes.get(window_selector, "unknown")

    def get(self, action: str, params: Dict[str, Any]) -> Tuple[bool, Any]:
        """
        Try to get a cached response.

        Returns:
            Tuple of (hit: bool, value: Any or None)
        """
        key = self._normalize_key(action, params)

        if key not in self._cache:
            self._stats["misses"] += 1
            return (False, None)

        entry = self._cache[key]

        # Check expiration
        if entry.is_expired():
            del self._cache[key]
            self._stats["misses"] += 1
            return (False, None)

        # Check window hash validity
        window_selector = params.get("window") or params.get("selector")
        current_hash = self._get_window_hash(window_selector)
        if entry.window_hash != current_hash and entry.window_hash != "desktop":
            del self._cache[key]
            self._stats["invalidations"] += 1
            return (False, None)

        # Cache hit
        entry.touch()
        self._cache.move_to_end(key)  # LRU: move to end
        self._stats["hits"] += 1
        return (True, entry.value)

    def set(self, action: str, params: Dict[str, Any], value: Any):
        """Store a response in cache."""
        key = self._normalize_key(action, params)
        ttl = self.TTL_BY_COMMAND.get(action, self.DEFAULT_TTL)

        window_selector = params.get("window") or params.get("selector")
        window_hash = self._get_window_hash(window_selector)

        # Evict if at capacity
        while len(self._cache) >= self.max_entries:
            oldest_key = next(iter(self._cache))
            del self._cache[oldest_key]
            self._stats["evictions"] += 1

        self._cache[key] = CacheEntry(
            key=key,
            value=value,
            window_hash=window_hash,
            created_at=time.time(),
            ttl_seconds=ttl,
        )

    def update_window_hash(self, window_selector: str, new_hash: str):
        """Update window hash (invalidates related cache entries)."""
        old_hash = self._window_hashes.get(window_selector)
        if old_hash != new_hash:
            self._window_hashes[window_selector] = new_hash
            # Invalidate entries for this window
            keys_to_remove = []
            for key, entry in self._cache.items():
                if entry.window_hash == old_hash:
                    keys_to_remove.append(key)
            for key in keys_to_remove:
                del self._cache[key]
                self._stats["invalidations"] += 1

    def invalidate_window(self, window_selector: str):
        """Explicitly invalidate all cache entries for a window."""
        window_hash = self._window_hashes.get(window_selector, "unknown")
        keys_to_remove = [
            k for k, v in self._cache.items() if v.window_hash == window_hash
        ]
        for key in keys_to_remove:
            del self._cache[key]
            self._stats["invalidations"] += 1

    def clear(self):
        """Clear all cache entries."""
        self._cache.clear()
        self._window_hashes.clear()

    def get_stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        total = self._stats["hits"] + self._stats["misses"]
        hit_rate = self._stats["hits"] / total if total > 0 else 0.0
        return {
            **self._stats,
            "entries": len(self._cache),
            "hit_rate": round(hit_rate, 3),
        }


class ResponseCompressor:
    """
    Compresses verbose responses to reduce token consumption.

    Strategies:
    - Remove empty/null fields
    - Truncate long text content
    - Summarize large element lists
    - Extract only essential properties
    """

    MAX_TEXT_LENGTH = 100
    MAX_ELEMENTS_FULL = 10
    MAX_ELEMENTS_SUMMARY = 50

    # Essential fields by response type
    ESSENTIAL_FIELDS = {
        "element": ["name", "type", "id", "x", "y", "w", "h"],
        "window": ["title", "bounds", "is_enabled"],
        "summary": ["title", "element_count", "key_elements"],
    }

    @classmethod
    def compress(
        cls, response: Dict[str, Any], context: str = "default"
    ) -> Dict[str, Any]:
        """
        Compress a response based on context.

        Args:
            response: The raw response from MainAgentService
            context: Hint about what kind of response this is

        Returns:
            Compressed response
        """
        if not isinstance(response, dict):
            return response

        # Don't compress errors
        if response.get("status") == "error":
            return response

        compressed = {}

        for key, value in response.items():
            if value is None or value == "" or value == []:
                continue  # Skip empty values

            if isinstance(value, str) and len(value) > cls.MAX_TEXT_LENGTH:
                compressed[key] = value[: cls.MAX_TEXT_LENGTH] + "..."
            elif isinstance(value, list) and len(value) > cls.MAX_ELEMENTS_FULL:
                # Summarize large lists
                compressed[key] = cls._compress_element_list(value)
            elif isinstance(value, dict):
                compressed[key] = cls._compress_dict(value)
            else:
                compressed[key] = value

        return compressed

    @classmethod
    def _compress_element_list(cls, elements: List[Any]) -> Dict[str, Any]:
        """Compress a list of elements to a summary."""
        if not elements:
            return {"count": 0, "elements": []}

        # Take first N full elements
        full_elements = elements[: cls.MAX_ELEMENTS_FULL]

        # Compress each element
        compressed_elements = [
            cls._compress_element(e) if isinstance(e, dict) else e
            for e in full_elements
        ]

        return {
            "count": len(elements),
            "showing": len(compressed_elements),
            "elements": compressed_elements,
        }

    @classmethod
    def _compress_element(cls, element: Dict[str, Any]) -> Dict[str, Any]:
        """Extract essential fields from an element."""
        essential = {}
        for field in cls.ESSENTIAL_FIELDS.get("element", []):
            if field in element and element[field]:
                value = element[field]
                if isinstance(value, str) and len(value) > 50:
                    value = value[:50] + "..."
                essential[field] = value

        # Add automation_id if present and different from name
        if element.get("automation_id") and element.get("automation_id") != element.get(
            "name"
        ):
            essential["id"] = element["automation_id"]

        return essential if essential else element

    @classmethod
    def _compress_dict(cls, d: Dict[str, Any]) -> Dict[str, Any]:
        """Recursively compress a dictionary."""
        result = {}
        for key, value in d.items():
            if value is None or value == "" or value == []:
                continue
            if isinstance(value, dict):
                result[key] = cls._compress_dict(value)
            elif isinstance(value, list) and len(value) > 10:
                result[key] = cls._compress_element_list(value)
            else:
                result[key] = value
        return result


class ProgressiveDisclosure:
    """
    Implements progressive disclosure pattern for AI agents.

    Strategy:
    1. Return minimal context first (summary/counts)
    2. If agent requests more detail, provide fuller response
    3. Track what the agent has already seen to avoid repetition
    """

    DISCLOSURE_LEVELS = {
        "minimal": 0,  # Just counts and key elements
        "standard": 1,  # Interactive elements with essential fields
        "detailed": 2,  # Full element tree with all properties
    }

    def __init__(self):
        self._agent_context: Dict[str, int] = {}  # window -> disclosure level seen
        self._failure_count: Dict[str, int] = {}  # track failures to auto-escalate

    def get_disclosure_level(self, window_selector: str) -> str:
        """Get current disclosure level for a window."""
        level = self._agent_context.get(window_selector, 0)
        failures = self._failure_count.get(window_selector, 0)

        # Auto-escalate after failures
        effective_level = min(level + failures, 2)

        if effective_level == 0:
            return "minimal"
        elif effective_level == 1:
            return "standard"
        else:
            return "detailed"

    def record_request(
        self, window_selector: str, explicit_level: Optional[str] = None
    ):
        """Record that agent made a request (auto-escalates disclosure)."""
        if explicit_level and explicit_level in self.DISCLOSURE_LEVELS:
            self._agent_context[window_selector] = self.DISCLOSURE_LEVELS[
                explicit_level
            ]
        else:
            current = self._agent_context.get(window_selector, 0)
            self._agent_context[window_selector] = min(current + 1, 2)

    def record_failure(self, window_selector: str):
        """Record that an operation failed (triggers more detail on next request)."""
        current = self._failure_count.get(window_selector, 0)
        self._failure_count[window_selector] = current + 1

    def reset(self, window_selector: Optional[str] = None):
        """Reset disclosure state."""
        if window_selector:
            self._agent_context.pop(window_selector, None)
            self._failure_count.pop(window_selector, None)
        else:
            self._agent_context.clear()
            self._failure_count.clear()


class WindowStateTracker:
    """
    Tracks window state changes for cache invalidation.

    Uses a lightweight hash of:
    - Window bounds
    - Element count
    - Key element names
    """

    def __init__(self):
        self._state_hashes: Dict[str, str] = {}
        self._last_check: Dict[str, float] = {}
        self._check_interval = 5.0  # Seconds between state checks

    def compute_hash(self, window_info: Dict[str, Any]) -> str:
        """Compute a state hash from window info."""
        # Extract key state indicators
        state_parts = [
            str(window_info.get("bounds", "")),
            str(window_info.get("element_count", 0)),
            str(window_info.get("title", "")),
        ]

        # Include key element names if available
        key_elements = window_info.get("key_elements", [])
        if key_elements:
            element_names = [e.get("name", "") for e in key_elements[:5]]
            state_parts.extend(element_names)

        state_str = "|".join(state_parts)
        return hashlib.md5(state_str.encode()).hexdigest()[:12]

    def check_changed(self, window_selector: str, new_info: Dict[str, Any]) -> bool:
        """Check if window state has changed since last check."""
        new_hash = self.compute_hash(new_info)
        old_hash = self._state_hashes.get(window_selector)

        if old_hash is None:
            self._state_hashes[window_selector] = new_hash
            self._last_check[window_selector] = time.time()
            return False  # First check, not a change

        changed = new_hash != old_hash
        self._state_hashes[window_selector] = new_hash
        self._last_check[window_selector] = time.time()

        return changed

    def should_recheck(self, window_selector: str) -> bool:
        """Check if enough time has passed to recheck window state."""
        last = self._last_check.get(window_selector, 0)
        return time.time() - last > self._check_interval

    def get_hash(self, window_selector: str) -> Optional[str]:
        """Get current hash for a window."""
        return self._state_hashes.get(window_selector)


class ContextManager:
    """
    Main context manager that orchestrates caching, compression, and disclosure.

    Acts as an intelligent proxy layer between AI agents and MainAgentService.

    Usage:
        ctx = ContextManager()

        # Before calling MainAgentService:
        cached, response = ctx.get_cached("explore_window", {"selector": "Calculator"})
        if cached:
            return response

        # After getting response from MainAgentService:
        compressed = ctx.process_response("explore_window", params, raw_response)
        return compressed
    """

    # Commands that should be cached
    CACHEABLE_COMMANDS = {
        "explore",
        "explore_window",
        "get_window_summary",
        "get_interactive_elements",
        "find_element",
        "element_exists",
        "get_element_brief",
        "get_window_info",
    }

    # Commands that modify state (invalidate cache)
    STATE_MODIFYING_COMMANDS = {
        "click",
        "double_click",
        "right_click",
        "type",
        "hotkey",
        "key_press",
        "smart_click",
        "smart_type",
        "vision_click",
        "close_window",
        "minimize_window",
        "maximize_window",
    }

    def __init__(self):
        self.cache = SemanticCache()
        self.compressor = ResponseCompressor
        self.disclosure = ProgressiveDisclosure()
        self.state_tracker = WindowStateTracker()
        self._enabled = True

    def is_enabled(self) -> bool:
        return self._enabled

    def enable(self):
        self._enabled = True

    def disable(self):
        self._enabled = False

    def get_cached(self, action: str, params: Dict[str, Any]) -> Tuple[bool, Any]:
        """
        Try to get a cached response for the command.

        Returns:
            Tuple of (cache_hit: bool, response: Any or None)
        """
        if not self._enabled:
            return (False, None)

        if action not in self.CACHEABLE_COMMANDS:
            return (False, None)

        return self.cache.get(action, params)

    def process_response(
        self,
        action: str,
        params: Dict[str, Any],
        response: Dict[str, Any],
        compress: bool = True,
    ) -> Dict[str, Any]:
        """
        Process and optionally cache a response.

        This should be called after getting a response from MainAgentService.

        Args:
            action: The command that was executed
            params: The parameters used
            response: The raw response from MainAgentService
            compress: Whether to compress the response

        Returns:
            Processed (possibly compressed) response
        """
        if not self._enabled:
            return response

        # Check for state changes
        if action in self.STATE_MODIFYING_COMMANDS:
            window_selector = params.get("window") or params.get("selector")
            if window_selector:
                self.cache.invalidate_window(window_selector)

        # Cache if appropriate
        if action in self.CACHEABLE_COMMANDS and response.get("status") == "success":
            self.cache.set(action, params, response)

            # Update window state hash if we have summary info
            if action in ("get_window_summary", "explore_window"):
                window_selector = params.get("selector") or params.get("window")
                if window_selector and "data" in response:
                    new_hash = self.state_tracker.compute_hash(response.get("data", {}))
                    self.cache.update_window_hash(window_selector, new_hash)

        # Track failures for progressive disclosure
        if response.get("status") == "error":
            window_selector = params.get("window") or params.get("selector")
            if window_selector:
                self.disclosure.record_failure(window_selector)

        # Compress if requested
        if compress:
            return self.compressor.compress(response, context=action)

        return response

    def get_stats(self) -> Dict[str, Any]:
        """Get context manager statistics."""
        return {"enabled": self._enabled, "cache": self.cache.get_stats()}

    def clear(self):
        """Clear all caches and state."""
        self.cache.clear()
        self.disclosure.reset()


# Singleton instance for use across the bridge
_context_manager: Optional[ContextManager] = None


def get_context_manager() -> ContextManager:
    """Get or create the singleton ContextManager instance."""
    global _context_manager
    if _context_manager is None:
        _context_manager = ContextManager()
    return _context_manager
