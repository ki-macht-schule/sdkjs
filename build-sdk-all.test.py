#!/usr/bin/env python3
"""Guard the min/common split that the Document Server load order requires."""
import importlib.util
import os
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))


def load_mod():
    spec = importlib.util.spec_from_file_location(
        "build_sdk_all", os.path.join(HERE, "build-sdk-all.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Sources(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mod = load_mod()
        cls.sdk = cls.mod.load_slide_sdk(HERE)
        cls.common = cls.mod.sources(HERE)

    def test_common_excludes_every_min_file(self):
        leaked = set(self.sdk["min"]) & set(self.common)
        self.assertEqual(leaked, set())

    def test_api_defines_stays_in_min(self):
        self.assertIn("slide/apiDefines.js", self.sdk["min"])
        self.assertNotIn("slide/apiDefines.js", self.common)

    def test_rail_files_are_in_common(self):
        for name in (
            "slide/Editor/Format/PresentationSections.js",
            "slide/Editor/Format/PresentationThemeRemount.js",
            "slide/Editor/Format/PresentationSectionRail.js",
            "slide/Drawing/DrawingDocument.js",
        ):
            self.assertIn(name, self.common)


class BuiltBundle(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mod = load_mod()
        cls.tmp = tempfile.NamedTemporaryFile(suffix=".js", delete=False)
        cls.tmp.close()
        rc = cls.mod.build(HERE, cls.tmp.name)
        if rc != 0:
            raise RuntimeError("build-sdk-all.py failed")
        with open(cls.tmp.name, encoding="utf-8") as fh:
            cls.text = fh.read()

    @classmethod
    def tearDownClass(cls):
        os.unlink(cls.tmp.name)

    def test_wrapped_in_one_iife(self):
        self.assertIn("(function(window, undefined) {", self.text)
        self.assertTrue(self.text.rstrip().endswith("})(window);"))

    def test_min_const_does_not_reappear(self):
        self.assertNotIn("const c_oAscPresentationViewMode", self.text)

    def test_rail_took(self):
        self.assertIn("kiwiProtocol", self.text)


if __name__ == "__main__":
    unittest.main()
