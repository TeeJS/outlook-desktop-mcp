"""
Unit tests for _split_folder_path — the folder-path splitter that honors
backslash-escaped slashes so folders with literal '/' in their name can be
addressed.

These tests do not require Outlook to be running.
"""
import os
import sys
import unittest

# Make src/ importable when running the tests from a source checkout
# without `pip install -e .`.
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from outlook_desktop_mcp.server import _split_folder_path  # noqa: E402


class SplitFolderPathTests(unittest.TestCase):

    # --- Basic, no-escape paths (unchanged behavior) ---

    def test_single_name(self):
        self.assertEqual(_split_folder_path("Inbox"), ["Inbox"])

    def test_simple_path(self):
        self.assertEqual(
            _split_folder_path("Inbox/Receipts"),
            ["Inbox", "Receipts"],
        )

    def test_deep_path(self):
        self.assertEqual(
            _split_folder_path("Inbox/Receipts/2026/Q1"),
            ["Inbox", "Receipts", "2026", "Q1"],
        )

    def test_segments_are_stripped(self):
        self.assertEqual(
            _split_folder_path(" Inbox / Receipts "),
            ["Inbox", "Receipts"],
        )

    def test_empty_string(self):
        self.assertEqual(_split_folder_path(""), [""])

    # --- Escaped slashes (the new behavior) ---

    def test_escaped_slash_single_segment(self):
        # 'Qlik\/RMM\/SAP Alerts' -> one segment with literal slashes
        self.assertEqual(
            _split_folder_path(r"Qlik\/RMM\/SAP Alerts"),
            ["Qlik/RMM/SAP Alerts"],
        )

    def test_escaped_slash_inside_path(self):
        # 'Inbox/Qlik\/RMM\/SAP Alerts' -> two segments
        self.assertEqual(
            _split_folder_path(r"Inbox/Qlik\/RMM\/SAP Alerts"),
            ["Inbox", "Qlik/RMM/SAP Alerts"],
        )

    def test_escaped_slash_mid_path(self):
        # 'Inbox/IT/SAP\/apps/QlikSense' -> four segments
        self.assertEqual(
            _split_folder_path(r"Inbox/IT/SAP\/apps/QlikSense"),
            ["Inbox", "IT", "SAP/apps", "QlikSense"],
        )

    def test_escaped_backslash(self):
        # '\\\\' -> literal '\'
        self.assertEqual(
            _split_folder_path(r"weird\\name"),
            ["weird\\name"],
        )

    def test_escaped_backslash_in_path(self):
        self.assertEqual(
            _split_folder_path(r"Inbox/weird\\name/Sub"),
            ["Inbox", "weird\\name", "Sub"],
        )

    # --- Edge cases ---

    def test_unescaped_backslash_preserved(self):
        # Lone backslash before a non-escape character is preserved literally,
        # so callers that never used the escape syntax keep working.
        self.assertEqual(
            _split_folder_path(r"Inbox/foo\bar/Sub"),
            ["Inbox", r"foo\bar", "Sub"],
        )

    def test_trailing_backslash(self):
        # Backslash at end-of-string has nothing to escape; keep it literal.
        self.assertEqual(_split_folder_path("Inbox\\"), ["Inbox\\"])

    def test_only_escaped_slashes(self):
        self.assertEqual(_split_folder_path(r"\/\/\/"), ["///"])

    def test_consecutive_separators(self):
        # 'Inbox//Sub' -> three segments (middle empty). Caller can decide
        # to reject this; the splitter just reports what it saw.
        self.assertEqual(_split_folder_path("Inbox//Sub"), ["Inbox", "", "Sub"])


if __name__ == "__main__":
    unittest.main()
