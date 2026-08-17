import json
import os
import tempfile
import unittest

from settings_store import DEFAULT_SETTINGS, load_settings, save_settings


class SettingsStoreTests(unittest.TestCase):
    def test_settings_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "settings.json")
            data = dict(DEFAULT_SETTINGS)
            data.update(
                {
                    "language": "en",
                    "render_dpi": 216,
                    "ocr_language": "kor",
                    "open_output_folder": False,
                }
            )
            save_settings(data, path)
            loaded = load_settings(path)
            self.assertEqual(loaded["language"], "en")
            self.assertEqual(loaded["render_dpi"], 216)
            self.assertEqual(loaded["ocr_language"], "kor")
            self.assertFalse(loaded["open_output_folder"])

    def test_corrupt_settings_fall_back_to_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "settings.json")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("{broken json")
            self.assertEqual(load_settings(path), DEFAULT_SETTINGS)

    def test_unknown_keys_are_not_persisted(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "settings.json")
            data = dict(DEFAULT_SETTINGS)
            data["secret_password"] = "must-not-be-saved"
            save_settings(data, path)
            with open(path, "r", encoding="utf-8") as handle:
                raw = json.load(handle)
            self.assertNotIn("secret_password", raw)


if __name__ == "__main__":
    unittest.main()
