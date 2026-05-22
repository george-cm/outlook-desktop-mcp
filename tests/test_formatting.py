"""
Unit tests for utils/formatting - no Outlook or COM required.
COM objects are simulated with MagicMock.
"""
from datetime import datetime
from unittest.mock import MagicMock
import pytest

from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_event_summary,
    format_task_summary,
)


def _pyTime(dt: datetime) -> MagicMock:
    """Simulate a COM PyTime object that has .isoformat() and .year."""
    m = MagicMock(spec=[])
    m.isoformat = lambda: dt.isoformat()
    m.year = dt.year
    return m


def _no_date() -> MagicMock:
    """Simulate Outlook's 'no date set' sentinel (01/01/4501)."""
    m = MagicMock(spec=[])
    m.year = 4501
    m.isoformat = lambda: "4501-01-01T00:00:00"
    return m


def _mock_email(received: datetime) -> MagicMock:
    item = MagicMock()
    item.EntryID = "ENTRY001"
    item.Subject = "Test"
    item.SenderEmailAddress = "a@b.com"
    item.SenderName = "A"
    item.ReceivedTime = _pyTime(received)
    item.UnRead = False
    item.Attachments.Count = 0
    return item


def _mock_event(start: datetime, end: datetime) -> MagicMock:
    item = MagicMock()
    item.EntryID = "EVT001"
    item.Subject = "Meeting"
    item.Start = _pyTime(start)
    item.End = _pyTime(end)
    item.Duration = 60
    item.Location = ""
    item.Organizer = "George"
    item.IsRecurring = False
    item.AllDayEvent = False
    item.BusyStatus = 2
    item.MeetingStatus = 0
    item.RequiredAttendees = ""
    item.OptionalAttendees = ""
    return item


def _mock_task(due_year: int = 4501) -> MagicMock:
    item = MagicMock()
    item.EntryID = "TASK001"
    item.Subject = "Fix bug"
    item.Status = 0
    item.PercentComplete = 0
    item.DueDate = _no_date() if due_year == 4501 else _pyTime(datetime(due_year, 6, 1))
    item.StartDate = _no_date()
    item.Importance = 1
    item.Complete = False
    item.Categories = ""
    item.Owner = "George"
    return item


# ---- email ----

class TestFormatEmailSummary:
    def test_received_time_is_iso8601(self):
        dt = datetime(2026, 3, 22, 14, 0, 0)
        result = format_email_summary(_mock_email(dt))
        assert result["received_time"] == "2026-03-22T14:00:00"

    def test_received_time_parseable(self):
        dt = datetime(2026, 1, 5, 9, 30, 0)
        result = format_email_summary(_mock_email(dt))
        assert datetime.fromisoformat(result["received_time"]) == dt

    def test_attachment_count_cached(self):
        """Attachments.Count must be accessed once (cached), not twice."""
        from unittest.mock import PropertyMock
        item = _mock_email(datetime(2026, 1, 1))
        count_mock = PropertyMock(return_value=3)
        type(item.Attachments).Count = count_mock
        result = format_email_summary(item)
        assert result["attachment_count"] == 3
        assert result["has_attachments"] is True
        assert count_mock.call_count == 1, (
            f"Attachments.Count accessed {count_mock.call_count} times - expected 1 (cached)"
        )

    def test_no_attachments(self):
        item = _mock_email(datetime(2026, 1, 1))
        item.Attachments.Count = 0
        result = format_email_summary(item)
        assert result["has_attachments"] is False
        assert result["attachment_count"] == 0


# ---- event ----

class TestFormatEventSummary:
    def test_start_is_iso8601(self):
        start = datetime(2026, 4, 1, 10, 0, 0)
        end = datetime(2026, 4, 1, 11, 0, 0)
        result = format_event_summary(_mock_event(start, end))
        assert result["start"] == "2026-04-01T10:00:00"

    def test_end_is_iso8601(self):
        start = datetime(2026, 4, 1, 10, 0, 0)
        end = datetime(2026, 4, 1, 11, 30, 0)
        result = format_event_summary(_mock_event(start, end))
        assert result["end"] == "2026-04-01T11:30:00"

    def test_start_parseable(self):
        start = datetime(2026, 6, 15, 9, 0, 0)
        result = format_event_summary(_mock_event(start, start))
        assert datetime.fromisoformat(result["start"]) == start


# ---- task ----

class TestFormatTaskSummary:
    def test_no_due_date_returns_none(self):
        result = format_task_summary(_mock_task(due_year=4501))
        assert result["due_date"] is None

    def test_no_start_date_returns_none(self):
        result = format_task_summary(_mock_task())
        assert result["start_date"] is None

    def test_due_date_iso8601_when_set(self):
        result = format_task_summary(_mock_task(due_year=2026))
        assert result["due_date"] == "2026-06-01T00:00:00"
        datetime.fromisoformat(result["due_date"])

    def test_sentinel_check_uses_year_not_string(self):
        """Sentinel detection must use .year == 4501, not str() comparison.

        If str() were used, the comparison would be locale-dependent
        ('01/01/4501' vs '01.01.4501' vs '4501-01-01') and would silently
        fail on non-US Windows.
        """
        sentinel = _no_date()
        assert sentinel.year == 4501
        item = _mock_task()
        item.DueDate = sentinel
        result = format_task_summary(item)
        assert result["due_date"] is None, (
            "due_date must be None for sentinel year=4501 (locale-independent check)"
        )

    def test_due_date_accessed_once(self):
        """DueDate must be cached - no double COM round-trip."""
        from unittest.mock import PropertyMock
        item = _mock_task(due_year=2026)
        due_mock = PropertyMock(return_value=_pyTime(datetime(2026, 6, 1)))
        type(item).DueDate = due_mock
        format_task_summary(item)
        assert due_mock.call_count == 1, (
            f"DueDate accessed {due_mock.call_count} times - expected 1 (cached)"
        )
