import unittest

from main import App, OP_MERGE


class MainStateTests(unittest.TestCase):
    def test_merge_input_comes_only_from_queue(self):
        app = App.__new__(App)
        app.selected_input = "stale-selection.pdf"
        app.file_queue = []
        self.assertEqual(app._resolved_input_for_operation(OP_MERGE), ())

        app.file_queue = ["a.pdf", "b.pdf"]
        self.assertEqual(app._resolved_input_for_operation(OP_MERGE), ("a.pdf", "b.pdf"))

        app.file_queue.clear()
        self.assertEqual(app._resolved_input_for_operation(OP_MERGE), ())


if __name__ == "__main__":
    unittest.main()
