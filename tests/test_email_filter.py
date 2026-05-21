"""
Unit tests for _escape_dasl_value and _build_email_filter — the Restrict
filter construction used by list_emails and search_emails.

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

from outlook_desktop_mcp.server import (  # noqa: E402
    _build_email_filter,
    _escape_dasl_value,
)


class EscapeDaslValueTests(unittest.TestCase):

    def test_no_special_chars(self):
        self.assertEqual(_escape_dasl_value("plain text"), "plain text")

    def test_single_quote_doubled(self):
        self.assertEqual(_escape_dasl_value("O'Brien"), "O''Brien")

    def test_multiple_single_quotes(self):
        self.assertEqual(_escape_dasl_value("a'b'c"), "a''b''c")

    def test_wildcards_not_escaped(self):
        # _safe_dasl handles LIKE wildcards; _escape_dasl_value is for
        # exact-equality where % and _ are just literal characters.
        self.assertEqual(_escape_dasl_value("100% off"), "100% off")
        self.assertEqual(_escape_dasl_value("foo_bar"), "foo_bar")

    def test_empty(self):
        self.assertEqual(_escape_dasl_value(""), "")


class BuildEmailFilterTests(unittest.TestCase):

    # --- Empty / no-filter case ---

    def test_no_clauses_returns_empty(self):
        self.assertEqual(_build_email_filter(), "")

    # --- Simple-Restrict mode (no subject_like) ---

    def test_unread_only_no_dasl_prefix(self):
        # Without subject_like, no @SQL= prefix needed; uses simple Restrict.
        f = _build_email_filter(unread_only=True)
        self.assertEqual(f, "[UnRead] = True")

    def test_sender_email_uses_field_reference(self):
        # [SenderEmailAddress] reads through Outlook's COM property accessor
        # — same value the rest of the codebase sees as item.SenderEmailAddress.
        f = _build_email_filter(sender_email="alice@example.com")
        self.assertEqual(f, "[SenderEmailAddress] = 'alice@example.com'")

    def test_sender_email_strips_whitespace(self):
        f = _build_email_filter(sender_email="  alice@example.com  ")
        self.assertIn("= 'alice@example.com'", f)

    def test_sender_email_with_single_quote(self):
        # Apostrophes in addresses are rare but legal; verify escaping.
        f = _build_email_filter(sender_email="o'brien@example.com")
        self.assertIn("= 'o''brien@example.com'", f)

    def test_sender_name(self):
        f = _build_email_filter(sender_name="Qlikview, Administrator")
        self.assertEqual(f, "[SenderName] = 'Qlikview, Administrator'")

    def test_start_date_only_adds_end_now(self):
        f = _build_email_filter(start_date="2026-01-01")
        # Without subject_like, simple Restrict (no @SQL prefix). Two
        # date clauses (the >= start and the implicit <= now).
        self.assertFalse(f.startswith("@SQL="))
        self.assertEqual(f.count("[ReceivedTime]"), 2)
        self.assertIn(">= '01/01/2026 00:00'", f)
        self.assertIn("<= '", f)

    def test_end_date_only_no_implicit_start(self):
        f = _build_email_filter(end_date="2026-01-01")
        self.assertEqual(f, "[ReceivedTime] <= '01/01/2026 00:00'")

    def test_explicit_start_and_end(self):
        f = _build_email_filter(start_date="2026-01-01", end_date="2026-02-01")
        self.assertIn(">= '01/01/2026 00:00'", f)
        self.assertIn("<= '02/01/2026 00:00'", f)
        self.assertIn(" AND ", f)

    def test_sender_plus_date_simple_restrict(self):
        f = _build_email_filter(
            sender_email="bot@example.com",
            end_date="2026-05-20",
        )
        self.assertFalse(f.startswith("@SQL="))
        self.assertIn("[SenderEmailAddress] = 'bot@example.com'", f)
        self.assertIn("[ReceivedTime] <= '05/20/2026 00:00'", f)
        self.assertIn(" AND ", f)

    # --- DASL mode (subject_like forces @SQL= prefix) ---

    def test_subject_like_forces_dasl_prefix(self):
        f = _build_email_filter(subject_like="invoice")
        self.assertTrue(f.startswith("@SQL="))

    def test_subject_like_wraps_with_percent(self):
        f = _build_email_filter(subject_like="invoice")
        self.assertIn("[Subject] LIKE '%invoice%'", f)
        self.assertIn(
            "\"urn:schemas:httpmail:textdescription\" LIKE '%invoice%'",
            f,
        )

    def test_subject_like_user_wildcards_escaped(self):
        # _safe_dasl escapes % and _ so user-supplied wildcards don't expand.
        f = _build_email_filter(subject_like="100%")
        self.assertIn("[%]", f)

    def test_subject_plus_sender_combined_dasl(self):
        f = _build_email_filter(
            subject_like="urgent",
            sender_email="bot@example.com",
        )
        self.assertTrue(f.startswith("@SQL="))
        self.assertIn("[SenderEmailAddress] = 'bot@example.com'", f)
        self.assertIn("[Subject] LIKE '%urgent%'", f)

    def test_all_filters_combined(self):
        f = _build_email_filter(
            unread_only=True,
            start_date="2026-01-01",
            end_date="2026-02-01",
            sender_email="bot@example.com",
            sender_name="Robot",
            subject_like="urgent",
        )
        self.assertTrue(f.startswith("@SQL="))
        # Six top-level clauses joined by AND: unread, start, end,
        # sender_email, sender_name, subject_or_body. " AND " count = 5.
        self.assertEqual(f.count(" AND "), 5)


if __name__ == "__main__":
    unittest.main()
