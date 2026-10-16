"""How many accepts and rejects a package has seen. docx4j ``Generation``.

CR-010 section 3.2: a view is valid until the next accept or reject on its
package, as Office JS's is (a paragraph join removes a paragraph, and Office JS
throws on a stale change). A :class:`~.change.TrackedChange` listed before the
last one is stale, and its ``accept()`` / ``reject()`` raise.

The counter is kept beside the package, weakly, so that a package nobody holds
takes its counter with it. A body with no package keys by its container.
"""

from __future__ import annotations

import threading
import weakref
from typing import Any

__all__ = ["Generation"]

_counters: weakref.WeakKeyDictionary[Any, list[int]] = weakref.WeakKeyDictionary()
_fallback: dict[int, tuple[Any, list[int]]] = {}
_lock = threading.Lock()


def _counter(key: Any) -> list[int]:
    with _lock:
        try:
            found = _counters.get(key)
            if found is None:
                found = _counters[key] = [0]
            return found
        except TypeError:  # not weakly referenceable: kept with the key alive
            entry = _fallback.get(id(key))
            if entry is None or entry[0] is not key:
                entry = _fallback[id(key)] = (key, [0])
            return entry[1]


class Generation:
    """A package's accept-and-reject count at the moment a change was listed."""

    __slots__ = ("_at", "_counter")

    def __init__(self, counter: list[int]) -> None:
        """Remember where the counter stands now."""
        self._counter = counter
        self._at = counter[0]

    @classmethod
    def of(cls, key: Any) -> Generation:
        """The generation of a package (or of a detached container) now."""
        return cls(_counter(key))

    def is_current(self) -> bool:
        """True when nothing has been accepted or rejected since."""
        return self._counter[0] == self._at

    def advance(self) -> None:
        """Make every view listed before now stale."""
        with _lock:
            self._counter[0] += 1
