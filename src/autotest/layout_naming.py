# -*- coding: utf-8 -*-
"""Shared filename-template values for Layout output."""

from random import SystemRandom


_random = SystemRandom()
_used_random9 = set()


def random_filename_values(template):
    """Allocate a nine-digit value once per filename, without session repeats."""
    raw = str(template)
    if "{random9" not in raw and "{see" not in raw:
        return {}
    while True:
        value = _random.randrange(100000000, 1000000000)
        if value not in _used_random9:
            _used_random9.add(value)
            values = {}
            if "{random9" in raw:
                values["random9"] = str(value)
            if "{see" in raw:
                values["see"] = value
            return values


def with_random_prefix(template):
    """Replace the leading nine digits, or prepend a random placeholder."""
    raw = str(template or "{form_id}").strip() or "{form_id}"
    if raw.startswith(("{random9}", "{see:09}", "{see:09d}")):
        return raw
    if len(raw) >= 9 and all(char in "0123456789" for char in raw[:9]):
        raw = raw[9:]
    return "{random9}" + raw
