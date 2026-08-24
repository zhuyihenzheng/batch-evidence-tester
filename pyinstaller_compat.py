# -*- coding: utf-8 -*-
"""Build-time compatibility helpers for old Anaconda Python distributions."""

import inspect
import sysconfig


def patch_anaconda_sysconfig():
    """Allow PyInstaller 4.10's no-argument private sysconfig call.

    Anaconda 5.2's Python 3.6 build made ``check_exists`` a required argument
    to ``_get_sysconfigdata_name``.  Upstream CPython and PyInstaller 4.10
    expect the function to take no arguments.  Only patch the function when
    its signature proves that a positional argument is required.
    """
    getter = getattr(sysconfig, "_get_sysconfigdata_name", None)
    if getter is None or getattr(getter, "_layout_txt_compat", False):
        return False

    try:
        signature = inspect.signature(getter)
    except (TypeError, ValueError):
        return False

    try:
        signature.bind()
    except TypeError:
        original_getter = getter

        def compatible_getter(check_exists=True):
            return original_getter(check_exists)

        compatible_getter._layout_txt_compat = True
        sysconfig._get_sysconfigdata_name = compatible_getter
        return True

    return False
