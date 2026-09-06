# -*- coding: utf-8 -*-
"""Shared filename-template values for Layout output."""

from random import SystemRandom


_random = SystemRandom()
_used_random9 = set()


def random_filename_values(template):
    """Allocate a nine-digit value once per filename, without session repeats."""
    raw = str(template)
    if "{random9" not in raw:
        return {}
    while True:
        value = _random.randrange(100000000, 1000000000)
        if value not in _used_random9:
            _used_random9.add(value)
            return {"random9": str(value)}


def with_random_prefix(template):
    """Replace the leading nine digits, or prepend a random placeholder."""
    raw = str(template or "{form_id}").strip() or "{form_id}"
    if raw.startswith("{random9}"):
        return raw
    for prefix in ("{seq:09}", "{seq:09d}"):
        if raw.startswith(prefix):
            return "{random9}" + raw[len(prefix):]
    if len(raw) >= 9 and all(char in "0123456789" for char in raw[:9]):
        raw = raw[9:]
    return "{random9}" + raw
