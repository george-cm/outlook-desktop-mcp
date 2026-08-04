"""
Unit tests for draft-safe mode gating on send_email, reply_email, create_meeting.
No Outlook or COM installation required - bridge.call() is replaced with a
synchronous passthrough over MagicMock outlook/namespace objects.
"""
import asyncio
from unittest.mock import MagicMock

import pytest

from outlook_desktop_mcp import server
from outlook_desktop_mcp.utils.constants import OL_CLASS_MAIL


@pytest.fixture
def fake_com(monkeypatch):
    """Replace bridge.call with a synchronous passthrough over fake COM objects."""
    fake_outlook = MagicMock()
    fake_namespace = MagicMock()
    fake_outlook.Session.Accounts = []

    async def _fake_call(func, *args, **kwargs):
        return func(fake_outlook, fake_namespace, *args, **kwargs)

    monkeypatch.setattr(server.bridge, "call", _fake_call)
    return fake_outlook, fake_namespace


# ---- send_email ----

def test_send_email_sends_immediately_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, _ = fake_com
    asyncio.run(server.send_email(to="alice@example.com", subject="s", body="b"))
    mail = fake_outlook.CreateItem.return_value
    mail.Send.assert_called_once()
    mail.Save.assert_not_called()


def test_send_email_drafts_when_draft_mode_on(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    result = asyncio.run(server.send_email(to="alice@example.com", subject="s", body="b"))
    mail = fake_outlook.CreateItem.return_value
    mail.Save.assert_called_once()
    mail.Send.assert_not_called()
    assert "Draft saved" in result


def test_send_email_sends_when_draft_mode_on_and_confirm_send_true(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    asyncio.run(server.send_email(
        to="alice@example.com", subject="s", body="b", confirm_send=True,
    ))
    mail = fake_outlook.CreateItem.return_value
    mail.Send.assert_called_once()
    mail.Save.assert_not_called()


# ---- reply_email ----

def test_reply_email_sends_immediately_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.Reply.return_value.Body = "original quoted body"
    asyncio.run(server.reply_email(entry_id="AAA", body="my note"))
    reply = source.Reply.return_value
    reply.Send.assert_called_once()
    reply.Save.assert_not_called()
    assert reply.Body == "my note\n\noriginal quoted body"


def test_reply_email_drafts_via_reply_never_createitem(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.Reply.return_value.Body = "original quoted body"
    result = asyncio.run(server.reply_email(entry_id="AAA", body="my note"))
    reply = source.Reply.return_value
    reply.Save.assert_called_once()
    reply.Send.assert_not_called()
    fake_outlook.CreateItem.assert_not_called()
    assert "Reply draft saved" in result


def test_reply_email_reply_all_uses_replyall(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.ReplyAll.return_value.Body = "original quoted body"
    asyncio.run(server.reply_email(entry_id="AAA", body="note", reply_all=True))
    source.ReplyAll.assert_called_once()
    source.ReplyAll.return_value.Save.assert_called_once()


# ---- create_meeting ----

def test_create_meeting_sends_invites_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, _ = fake_com
    asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com",
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Send.assert_called_once()
    appt.Save.assert_not_called()


def test_create_meeting_saves_unsent_when_draft_mode_on(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    result = asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com",
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Recipients.ResolveAll.assert_called_once()
    appt.Save.assert_called_once()
    appt.Send.assert_not_called()
    assert "NOT sent" in result


def test_create_meeting_sends_when_draft_mode_on_and_confirm_send_true(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com", confirm_send=True,
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Send.assert_called_once()
    appt.Save.assert_not_called()
