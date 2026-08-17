import os
import tempfile
import unittest

import fitz

from page_editor import PagePlanEntry, apply_page_edits


def make_three_page_pdf(path: str) -> None:
    doc = fitz.open()
    try:
        for index in range(3):
            page = doc.new_page(width=400, height=250)
            page.insert_text((40, 80), f"PAGE-{index + 1}", fontsize=18)
        doc.save(path)
    finally:
        doc.close()


class PageEditorTests(unittest.TestCase):
    def test_reorder_delete_and_rotate(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "source.pdf")
            target = os.path.join(tmp, "edited.pdf")
            make_three_page_pdf(source)

            success, message, saved_path = apply_page_edits(
                source,
                target,
                [PagePlanEntry(2, 90), PagePlanEntry(0, 0)],
            )
            self.assertTrue(success, message)
            self.assertEqual(saved_path, target)

            edited = fitz.open(target)
            try:
                self.assertEqual(len(edited), 2)
                self.assertIn("PAGE-3", edited[0].get_text("text"))
                self.assertEqual(edited[0].rotation, 90)
                self.assertIn("PAGE-1", edited[1].get_text("text"))
            finally:
                edited.close()

    def test_invalid_plan_does_not_destroy_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "source.pdf")
            target = os.path.join(tmp, "existing.pdf")
            make_three_page_pdf(source)
            original = b"existing-output"
            with open(target, "wb") as handle:
                handle.write(original)

            success, _message, _saved_path = apply_page_edits(
                source,
                target,
                [PagePlanEntry(99, 0)],
            )
            self.assertFalse(success)
            with open(target, "rb") as handle:
                self.assertEqual(handle.read(), original)


if __name__ == "__main__":
    unittest.main()
