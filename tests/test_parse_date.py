"""Unit tests for server._parse_date - no Outlook or COM required."""
from datetime import datetime, timezone

from outlook_desktop_mcp.server import _parse_date


class TestParseDate:
    def test_utc_z_suffix_converts_to_local_naive(self):
        result = _parse_date("2026-11-06T10:30:00Z")
        expected = datetime(2026, 11, 6, 10, 30, tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
        assert result == expected
        assert result.tzinfo is None

    def test_offset_suffix_converts_to_local_naive(self):
        result = _parse_date("2026-11-06T10:30:00+00:00")
        expected = datetime(2026, 11, 6, 10, 30, tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
        assert result == expected

    def test_naive_input_unchanged(self):
        result = _parse_date("2026-11-06 12:30")
        assert result == datetime(2026, 11, 6, 12, 30)
        assert result.tzinfo is None
