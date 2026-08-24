# -*- coding: utf-8 -*-

import sysconfig
import unittest
from unittest.mock import patch

from pyinstaller_compat import patch_anaconda_sysconfig


class PyInstallerCompatTest(unittest.TestCase):
    def test_standard_sysconfig_signature_is_unchanged(self):
        def get_name():
            return "_sysconfigdata_standard"

        with patch.object(sysconfig, "_get_sysconfigdata_name", get_name):
            self.assertFalse(patch_anaconda_sysconfig())
            self.assertIs(sysconfig._get_sysconfigdata_name, get_name)
            self.assertEqual(sysconfig._get_sysconfigdata_name(), "_sysconfigdata_standard")

    def test_anaconda_required_argument_gets_compatible_default(self):
        calls = []

        def get_name(check_exists):
            calls.append(check_exists)
            return "_sysconfigdata_anaconda"

        with patch.object(sysconfig, "_get_sysconfigdata_name", get_name):
            self.assertTrue(patch_anaconda_sysconfig())
            self.assertEqual(
                sysconfig._get_sysconfigdata_name(), "_sysconfigdata_anaconda"
            )
            self.assertEqual(calls, [True])

            self.assertEqual(
                sysconfig._get_sysconfigdata_name(False), "_sysconfigdata_anaconda"
            )
            self.assertEqual(calls, [True, False])

            self.assertFalse(patch_anaconda_sysconfig())

    def test_missing_private_function_is_ignored(self):
        with patch.object(
            sysconfig, "_get_sysconfigdata_name", new=None, create=True
        ):
            self.assertFalse(patch_anaconda_sysconfig())


if __name__ == "__main__":
    unittest.main()
