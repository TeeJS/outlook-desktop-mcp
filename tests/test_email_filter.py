"""
Unit tests for _escape_dasl_value and _build_email_filter — the DASL @SQL
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

    # --- Individual clause shapes ---

    def test_unread_only(self):
        f = _build_email_filter(unread_only=True)
        self.assertEqual(f, '@SQL="urn:schemas:httpmail:read" = 0')

    def test_sender_email_uses_smtp_proptag(self):
        f = _build_email_filter(sender_email="alice@example.com")
        # PR_SENDER_SMTP_ADDRESS_W = 0x5D01001F — universal SMTP regardless of
        # whether the sender is Exchange-internal or external.
        self.assertIn(
            '"http://schemas.microsoft.com/mapi/proptag/0x5D01001F" = \'alice@example.com\'',
            f,
        )

    def test_all_filters_combined_uses_correct_proptag(self):
        # Sanity-check that the proptag in the combined filter matches the
        # one used in the single-filter case (no drift between code paths).
        f = _build_email_filter(
            unread_only=True,
            sender_email="x@y.com",
        )
        self.assertIn("0x5D01001F", f)
        self.assertNotIn("0x39FE001E", f)

    def test_sender_email_strips_whitespace(self):
        f = _build_email_filter(sender_email="  alice@example.com  ")
        self.assertIn("= 'alice@example.com'", f)

    def test_sender_email_with_single_quote(self):
        # Apostrophes in addresses are rare but legal; verify escaping.
        f = _build_email_filter(sender_email="o'brien@example.com")
        self.assertIn("= 'o''brien@example.com'", f)

    def test_sender_name(self):
        f = _build_email_filter(sender_name="Qlikview, Administrator")
        self.assertIn(
            "\"urn:schemas:httpmail:fromname\" = 'Qlikview, Administrator'",
            f,
        )

    def test_subject_like_wraps_with_percent(self):
        f = _build_email_filter(subject_like="invoice")
        self.assertIn(
            '"urn:schemas:httpmail:subject" LIKE \'%invoice%\'',
            f,
        )
        self.assertIn(
            '"urn:schemas:httpmail:textdescription" LIKE \'%invoice%\'',
            f,
        )

    def test_subject_like_user_wildcards_escaped(self):
        # _safe_dasl escapes % and _ so user-supplied wildcards don't expand.
        f = _build_email_filter(subject_like="100%")
        self.assertIn("[%]", f)

    def test_start_date_only_adds_end_now(self):
        # When only start_date is provided, end_date defaults to "now".
        # Verify both >= and <= clauses are present.
        f = _build_email_filter(start_date="2026-01-01")
        self.assertEqual(f.count("datereceived"), 2)
        self.assertIn(">= '01/01/2026 00:00'", f)
        self.assertIn("<= '", f)

    def test_end_date_only_no_implicit_start(self):
        f = _build_email_filter(end_date="2026-01-01")
        self.assertEqual(f.count("datereceived"), 1)
        self.assertIn("<= '01/01/2026 00:00'", f)
        self.assertNotIn(">= '", f)

    def test_explicit_start_and_end(self):
        f = _build_email_filter(start_date="2026-01-01", end_date="2026-02-01")
        self.assertIn(">= '01/01/2026 00:00'", f)
        self.assertIn("<= '02/01/2026 00:00'", f)

    # --- Combined filters AND-ed together ---

    def test_sender_plus_date(self):
        f = _build_email_filter(
            sender_email="bot@example.com",
            end_date="2026-05-20",
        )
        self.assertTrue(f.startswith("@SQL="))
        self.assertIn(" AND ", f)
        self.assertIn("0x5D01001F", f)
        self.assertIn("datereceived", f)

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
        # All six clauses present, joined by AND.
        # Count of " AND " separators should be (clauses - 1) where clauses
        # = unread + 2 date + sender_email + sender_name + subject = 6
        self.assertEqual(f.count(" AND "), 5)

    def test_filter_prefix(self):
        # Any non-empty filter must start with @SQL=.
        f = _build_email_filter(sender_email="x@y.com")
        self.assertTrue(f.startswith("@SQL="))


if __name__ == "__main__":
    unittest.main()
