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
