# `outlook` CLI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.
>
> **PREREQUISITE:** Plan `2026-05-22-fix-com-bridge-formatting.md` must be fully merged to `main` before starting this plan.

**Goal:** Add an `outlook` CLI command exposing all 30 MCP tools as grouped Typer subcommands, backed by a new `OutlookService` class that decouples business logic from MCP wiring.

**Architecture:** Extract inner COM functions from `server.py` into `OutlookService(bridge)` class in `service.py`. Refactor `server.py` to thin `@mcp.tool()` wrappers. Add `cli.py` with Typer grouped subcommands that call `OutlookService` directly. CLI owns its own `OutlookBridge` lifecycle.

**Tech Stack:** Python 3.13, Typer 0.25.x, asyncio, pywin32, pytest, unittest.mock

---

## File Map

| File | Action | Responsibility |
|---|---|---|
| `src/outlook_desktop_mcp/service.py` | **Create** | `OutlookService(bridge)` — all 30 business logic methods |
| `src/outlook_desktop_mcp/cli.py` | **Create** | Typer app, 7 command groups, 30 commands |
| `src/outlook_desktop_mcp/server.py` | **Modify** | Thin `@mcp.tool()` wrappers calling `svc.*` |
| `src/outlook_desktop_mcp/entrypoint.py` | **Modify** | Unchanged (MCP server entrypoint only) |
| `pyproject.toml` | **Modify** | Add `outlook` entry point, declare `typer` dep |
| `tests/test_service.py` | **Create** | Unit tests for OutlookService methods |
| `tests/test_cli.py` | **Create** | CLI command tests via Typer CliRunner |

---

## Key Patterns (read before starting tasks)

### OutlookService method pattern

Every method in `OutlookService` follows this structure:
```python
async def list_emails(
    self,
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
        # ... COM logic (lifted verbatim from server.py inner function) ...
        return json.dumps(results, indent=2, default=str)

    try:
        return await self._bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"
```

The inner `_list` function is lifted **verbatim** from `server.py`. The only changes are:
1. Module-level helpers (`_resolve_store`, `_get_item`, etc.) become `self._resolve_store`, `self._get_item` etc.
2. `bridge.call(...)` becomes `self._bridge.call(...)`
3. `format_com_error` imported from `utils.errors`

### CLI command pattern

```python
email_app = typer.Typer(help="Email operations")
app.add_typer(email_app, name="email")

@email_app.command("list")
def email_list(
    ctx: typer.Context,
    folder: str = typer.Option("inbox", help="Folder name"),
    count: int = typer.Option(10, help="Max results"),
    unread: bool = typer.Option(False, "--unread", help="Unread only"),
    start: str = typer.Option("", "--start", help="Start date ISO 8601"),
    end: str = typer.Option("", "--end", help="End date ISO 8601"),
    account: str = typer.Option("", "--account", help="Account substring"),
) -> None:
    svc: OutlookService = ctx.obj
    _run(svc.list_emails(folder, count, unread, start, end, account))
```

### `_run` helper

```python
import asyncio

def _run(coro, ctx: typer.Context | None = None) -> None:
    """Run an async service coroutine, print result to stdout, exit 1 on error."""
    try:
        result = asyncio.run(coro)
        typer.echo(result)
    except RuntimeError as e:
        typer.echo(json.dumps({"error": str(e)}), err=True)
        raise typer.Exit(2)
    except Exception as e:
        typer.echo(json.dumps({"error": str(e)}), err=True)
        raise typer.Exit(1)
```

### Service test pattern

```python
@pytest.fixture
def svc():
    bridge = MagicMock()
    bridge.call = AsyncMock(return_value='[{"entry_id": "X"}]')
    return OutlookService(bridge)

def test_list_emails_calls_bridge(svc):
    result = asyncio.run(svc.list_emails(folder="inbox", count=5))
    svc._bridge.call.assert_called_once()
    assert "entry_id" in result
```

### CLI test pattern

```python
from typer.testing import CliRunner
from unittest.mock import AsyncMock, MagicMock, patch

runner = CliRunner()

def test_email_list_outputs_json():
    mock_svc = MagicMock()
    mock_svc.list_emails = AsyncMock(return_value='[{"entry_id":"X","subject":"Hi"}]')
    with patch("outlook_desktop_mcp.cli.OutlookService") as MockSvc, \
         patch("outlook_desktop_mcp.cli.OutlookBridge"):
        MockSvc.return_value = mock_svc
        result = runner.invoke(app, ["email", "list"])
    assert result.exit_code == 0
    assert "entry_id" in result.output
```

---

## Task 1: Create worktree for CLI branch

- [ ] **Step 1: Ensure main has Phase 1 fixes merged**

```bash
git -C C:/Users/George.Murga/projects/outlook-desktop-mcp log --oneline -3
```
Expected: top commit message should reference "Fix OutlookBridge lifecycle bugs" or similar.

- [ ] **Step 2: Create worktree**

```bash
git -C C:/Users/George.Murga/projects/outlook-desktop-mcp worktree add .worktrees/cli -b feat/outlook-cli
```

- [ ] **Step 3: Verify tests still pass in new worktree**

```bash
cd C:/Users/George.Murga/projects/outlook-desktop-mcp/.worktrees/cli
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py -v
```
Expected: `32 passed`

---

## Task 2: Create `service.py` skeleton + private helpers

**Files:**
- Create: `src/outlook_desktop_mcp/service.py`
- Create: `tests/test_service.py` (skeleton)

- [ ] **Step 1: Write failing test for OutlookService instantiation**

Create `tests/test_service.py`:

```python
"""Unit tests for OutlookService — no Outlook or COM required."""
import asyncio
import json
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from outlook_desktop_mcp.service import OutlookService


@pytest.fixture
def mock_bridge():
    bridge = MagicMock()
    # Default: bridge.call returns an empty JSON array
    bridge.call = AsyncMock(return_value="[]")
    return bridge


@pytest.fixture
def svc(mock_bridge):
    return OutlookService(mock_bridge)


def test_service_instantiates(mock_bridge):
    s = OutlookService(mock_bridge)
    assert s is not None


def test_service_stores_bridge(mock_bridge):
    s = OutlookService(mock_bridge)
    assert s._bridge is mock_bridge
```

- [ ] **Step 2: Run test — expect FAIL**

```bash
uv run pytest tests/test_service.py::test_service_instantiates -v
```
Expected: `ModuleNotFoundError: No module named 'outlook_desktop_mcp.service'`

- [ ] **Step 3: Create `src/outlook_desktop_mcp/service.py` skeleton**

```python
"""
Outlook Service Layer
=====================
OutlookService encapsulates all Outlook COM business logic.
Methods have signatures matching the existing @mcp.tool() functions,
but take `self` instead of relying on module-level state.

This module has NO dependency on FastMCP, the MCP server, or the CLI.
It is importable anywhere, including on macOS for testing (the COM
imports are deferred inside inner functions that only run on the COM thread).
"""
import json
import logging
import os
import re
from datetime import datetime, timedelta

from outlook_desktop_mcp.com_bridge import OutlookBridge
from outlook_desktop_mcp.utils.constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
    OUTLOOK_DATE_FMT,
    MAX_ITEM_COUNT,
    SEND_USING_ACCOUNT_DISPID,
    OL_CLASS_MAIL,
    OL_CLASS_APPOINTMENT,
    OL_CLASS_TASK,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error

logger = logging.getLogger("outlook_desktop_mcp.service")


class OutlookService:
    """Business logic layer for Outlook COM operations.

    All public methods are async and return JSON strings.
    Methods delegate COM work to self._bridge.call().
    """

    def __init__(self, bridge: OutlookBridge) -> None:
        self._bridge = bridge

    # ------------------------------------------------------------------
    # Private helpers (these run on the COM thread inside inner functions)
    # ------------------------------------------------------------------

    @staticmethod
    def _safe_dasl(query: str) -> str:
        """Sanitize a string for use in a DASL LIKE filter value."""
        query = query.replace("%", "[%]").replace("_", "[_]")
        return query.replace("'", "''").replace('"', '""')

    @staticmethod
    def _check_item_class(item, expected_class: int, label: str) -> str | None:
        """Return an error string if item is the wrong type, else None."""
        if item.Class != expected_class:
            return f"Error: Entry ID does not refer to a {label}."
        return None

    @staticmethod
    def _get_item(namespace, entry_id: str, account: str = ""):
        """Retrieve an Outlook item by EntryID, optionally scoped to a store."""
        if account:
            store = OutlookService._require_store(namespace, account)
            return namespace.GetItemFromID(entry_id, store.StoreID)
        return namespace.GetItemFromID(entry_id)

    @staticmethod
    def _resolve_store(namespace, account: str = ""):
        """Resolve an account name to an Outlook Store object."""
        if not account:
            return namespace.DefaultStore
        account_lower = account.lower().strip()
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            if account_lower in store.DisplayName.lower():
                return store
        return None

    @staticmethod
    def _require_store(namespace, account: str = ""):
        """Resolve store, raising ValueError if not found."""
        store = OutlookService._resolve_store(namespace, account)
        if store is None:
            raise ValueError(
                f"Account '{account}' not found. "
                "Use list_accounts to see available accounts."
            )
        return store

    @staticmethod
    def _walk_folders(parent, name_lower: str):
        """Recursively search subfolders for a folder matching name_lower."""
        for i in range(parent.Folders.Count):
            try:
                f = parent.Folders.Item(i + 1)
                if f.Name.lower() == name_lower:
                    return f
                found = OutlookService._walk_folders(f, name_lower)
                if found:
                    return found
            except Exception as e:
                logger.debug("Skipping folder %d: %s", i + 1, e)
                continue
        return None

    @staticmethod
    def _resolve_folder(namespace, folder_name: str, store=None):
        """Resolve a folder name to an Outlook MAPIFolder object."""
        folder_name = folder_name.strip()
        store = store or namespace.DefaultStore

        if "/" in folder_name:
            parts = [p.strip() for p in folder_name.split("/")]
            current = OutlookService._resolve_folder(namespace, parts[0], store)
            if current is None:
                return None
            for part in parts[1:]:
                part_lower = part.lower()
                found = None
                for i in range(current.Folders.Count):
                    try:
                        f = current.Folders.Item(i + 1)
                        if f.Name.lower() == part_lower:
                            found = f
                            break
                    except Exception as e:
                        logger.debug("Skipping folder item %d: %s", i + 1, e)
                        continue
                if found is None:
                    return None
                current = found
            return current

        folder_lower = folder_name.lower()
        if folder_lower in FOLDER_NAME_TO_ENUM:
            return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

        root = store.GetRootFolder()
        for i in range(root.Folders.Count):
            try:
                f = root.Folders.Item(i + 1)
                if f.Name.lower() == folder_lower:
                    return f
            except Exception as e:
                logger.debug("Skipping root folder %d: %s", i + 1, e)
                continue

        return OutlookService._walk_folders(root, folder_lower)

    @staticmethod
    def _parse_date(date_str: str) -> datetime:
        """Parse ISO 8601 date string, raising ValueError with a useful message."""
        try:
            return datetime.fromisoformat(date_str)
        except ValueError:
            raise ValueError(
                f"Invalid date format: '{date_str}'. "
                "Use ISO 8601: '2026-03-10' or '2026-03-10T09:00:00'."
            )
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
uv run pytest tests/test_service.py -v
```
Expected: `2 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/service.py tests/test_service.py
git commit -m "Add OutlookService skeleton with private helpers"
```

---

## Task 3: Migrate email tools to OutlookService

**Files:**
- Modify: `src/outlook_desktop_mcp/service.py` (add 10 email methods)
- Modify: `tests/test_service.py` (add email tests)

- [ ] **Step 1: Add email tests to `tests/test_service.py`**

Append to `tests/test_service.py`:

```python
# ---- Email ----

class TestListEmails:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.list_emails())
        mock_bridge.call.assert_called_once()

    def test_returns_json_string(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value='[{"entry_id":"X"}]')
        result = asyncio.run(svc.list_emails())
        parsed = json.loads(result)
        assert isinstance(parsed, list)

    def test_error_returns_error_string(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(side_effect=Exception("COM fail"))
        result = asyncio.run(svc.list_emails())
        assert "Error" in result


class TestSendEmail:
    def test_calls_bridge_with_recipients(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value="Email sent: 'Hi' to a@b.com")
        result = asyncio.run(svc.send_email("a@b.com", "Hi", "Body"))
        mock_bridge.call.assert_called_once()
        assert "Email sent" in result


class TestMarkAsRead:
    def test_calls_bridge(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value="Marked as read: 'Subject'")
        result = asyncio.run(svc.mark_as_read("ENTRYID001"))
        mock_bridge.call.assert_called_once()
        assert "Marked as read" in result


class TestSearchEmails:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.search_emails("invoice"))
        mock_bridge.call.assert_called_once()
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_service.py -v
```
Expected: `AttributeError: 'OutlookService' object has no attribute 'list_emails'`

- [ ] **Step 3: Add email methods to `service.py`**

Add the following methods to `OutlookService` in `service.py`. Each inner function is the verbatim content from the corresponding `server.py` tool, with module-level helpers replaced by `OutlookService._helper_name`:

```python
    # ------------------------------------------------------------------
    # Accounts
    # ------------------------------------------------------------------

    async def list_accounts(self) -> str:
        """List all Outlook accounts (stores) configured in the profile."""
        def _list(outlook, namespace):
            default_id = namespace.DefaultStore.StoreID
            results = []
            for i in range(namespace.Stores.Count):
                store = namespace.Stores.Item(i + 1)
                results.append({
                    "display_name": store.DisplayName,
                    "store_id": store.StoreID,
                    "is_default": store.StoreID == default_id,
                })
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list)
        except Exception as e:
            return f"Error listing accounts: {format_com_error(e)}"

    # ------------------------------------------------------------------
    # Email
    # ------------------------------------------------------------------

    async def send_email(
        self,
        to: str,
        subject: str,
        body: str,
        cc: str = "",
        bcc: str = "",
        html_body: str = "",
        account: str = "",
    ) -> str:
        def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
            store = OutlookService._require_store(namespace, account)
            mail = outlook.CreateItem(OL_MAIL_ITEM)
            for acc in outlook.Session.Accounts:
                if acc.DeliveryStore.StoreID == store.StoreID:
                    mail._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, acc))
                    break
            mail.To = to
            mail.Subject = subject
            mail.Body = body
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            if html_body:
                mail.HTMLBody = html_body
            mail.Send()
            return f"Email sent: '{subject}' to {to}"

        try:
            return await self._bridge.call(_send, to, subject, body, cc, bcc, html_body, account)
        except Exception as e:
            return f"Error sending email: {format_com_error(e)}"

    async def list_emails(
        self,
        folder: str = "inbox",
        count: int = 10,
        unread_only: bool = False,
        start_date: str = "",
        end_date: str = "",
        account: str = "",
    ) -> str:
        def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
            count = min(max(1, count), MAX_ITEM_COUNT)
            store = OutlookService._require_store(namespace, account)
            target = OutlookService._resolve_folder(namespace, folder, store)
            if not target:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            items = target.Items
            items.Sort("[ReceivedTime]", True)
            restrictions = []
            if unread_only:
                restrictions.append("[UnRead] = True")
            if start_date:
                start = OutlookService._parse_date(start_date)
                restrictions.append(f"[ReceivedTime] >= '{start.strftime(OUTLOOK_DATE_FMT)}'")
            if end_date:
                end = OutlookService._parse_date(end_date)
                restrictions.append(f"[ReceivedTime] <= '{end.strftime(OUTLOOK_DATE_FMT)}'")
            elif start_date:
                restrictions.append(
                    f"[ReceivedTime] <= '{datetime.now().strftime(OUTLOOK_DATE_FMT)}'"
                )
            if restrictions:
                items = items.Restrict(" AND ".join(restrictions))
            results = []
            limit = min(count, items.Count)
            for i in range(limit):
                try:
                    results.append(format_email_summary(items.Item(i + 1)))
                except Exception as e:
                    logger.debug("Skipping email item %d: %s", i + 1, e)
                    continue
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
        except Exception as e:
            return f"Error listing emails: {format_com_error(e)}"

    async def read_email(
        self,
        entry_id: str = "",
        subject_search: str = "",
        folder: str = "inbox",
        account: str = "",
    ) -> str:
        def _read(outlook, namespace, entry_id, subject_search, folder, account):
            if entry_id:
                item = namespace.GetItemFromID(entry_id)
                return json.dumps(format_email_full(item), indent=2, default=str)
            if not subject_search:
                return json.dumps({"error": "Provide either entry_id or subject_search"})
            store = OutlookService._require_store(namespace, account)
            target = OutlookService._resolve_folder(namespace, folder, store)
            if not target:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            safe_query = OutlookService._safe_dasl(subject_search)
            filter_str = (
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
            )
            items = target.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)
            if items.Count == 0:
                return json.dumps({"error": f"No email found matching '{subject_search}'"})
            return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

        try:
            return await self._bridge.call(_read, entry_id, subject_search, folder, account)
        except Exception as e:
            return f"Error reading email: {format_com_error(e)}"

    async def mark_as_read(self, entry_id: str, account: str = "") -> str:
        def _mark(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return err
            subject = item.Subject
            item.UnRead = False
            item.Save()
            return f"Marked as read: '{subject}'"

        try:
            return await self._bridge.call(_mark, entry_id, account)
        except Exception as e:
            return f"Error marking email as read: {format_com_error(e)}"

    async def mark_as_unread(self, entry_id: str, account: str = "") -> str:
        def _mark(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return err
            subject = item.Subject
            item.UnRead = True
            item.Save()
            return f"Marked as unread: '{subject}'"

        try:
            return await self._bridge.call(_mark, entry_id, account)
        except Exception as e:
            return f"Error marking email as unread: {format_com_error(e)}"

    async def move_email(
        self,
        entry_id: str,
        target_folder: str = "archive",
        account: str = "",
    ) -> str:
        def _move(outlook, namespace, entry_id, target_folder, account):
            item = namespace.GetItemFromID(entry_id)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return err
            subject = item.Subject
            store = OutlookService._require_store(namespace, account)
            dest = OutlookService._resolve_folder(namespace, target_folder, store)
            if not dest:
                return (
                    f"Error: Target folder '{target_folder}' not found. "
                    "Use list_folders to see available folders."
                )
            item.Move(dest)
            return f"Moved '{subject}' to {target_folder}"

        try:
            return await self._bridge.call(_move, entry_id, target_folder, account)
        except Exception as e:
            return f"Error moving email: {format_com_error(e)}"

    async def reply_email(
        self,
        entry_id: str,
        body: str,
        reply_all: bool = False,
        account: str = "",
    ) -> str:
        def _reply(outlook, namespace, entry_id, body, reply_all, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return err
            subject = item.Subject
            reply_item = item.ReplyAll() if reply_all else item.Reply()
            reply_item.Body = body + "\n\n" + reply_item.Body
            reply_item.Send()
            return f"Reply sent to '{subject}' (reply_all={reply_all})"

        try:
            return await self._bridge.call(_reply, entry_id, body, reply_all, account)
        except Exception as e:
            return f"Error replying to email: {format_com_error(e)}"

    async def search_emails(
        self,
        query: str,
        folder: str = "inbox",
        count: int = 10,
        start_date: str = "",
        end_date: str = "",
        account: str = "",
    ) -> str:
        def _search(outlook, namespace, query, folder, count, start_date, end_date, account):
            count = min(max(1, count), MAX_ITEM_COUNT)
            store = OutlookService._require_store(namespace, account)
            target = OutlookService._resolve_folder(namespace, folder, store)
            if not target:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            safe_query = OutlookService._safe_dasl(query)
            dasl_parts = [
                f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
                f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
            ]
            if start_date:
                start = OutlookService._parse_date(start_date)
                dasl_parts.append(
                    f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime(OUTLOOK_DATE_FMT)}'"
                )
            if end_date:
                end = OutlookService._parse_date(end_date)
                dasl_parts.append(
                    f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime(OUTLOOK_DATE_FMT)}'"
                )
            elif start_date:
                dasl_parts.append(
                    f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime(OUTLOOK_DATE_FMT)}'"
                )
            filter_str = "@SQL=" + " AND ".join(dasl_parts)
            items = target.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)
            results = []
            limit = min(count, items.Count)
            for i in range(limit):
                try:
                    results.append(format_email_summary(items.Item(i + 1)))
                except Exception as e:
                    logger.debug("Skipping search result %d: %s", i + 1, e)
                    continue
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_search, query, folder, count, start_date, end_date, account)
        except Exception as e:
            return f"Error searching emails: {format_com_error(e)}"

    async def list_folders(
        self,
        folder: str = "",
        max_depth: int = 3,
        account: str = "",
    ) -> str:
        def _list(outlook, namespace, folder, max_depth, account):
            max_depth = min(max(1, max_depth), 10)
            store = OutlookService._require_store(namespace, account)
            if folder:
                start = OutlookService._resolve_folder(namespace, folder, store)
                if not start:
                    return json.dumps({"error": f"Folder '{folder}' not found"})
                base_path = folder
            else:
                start = store.GetRootFolder()
                base_path = ""

            def walk(f, depth, path_prefix):
                current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
                result = {
                    "name": f.Name,
                    "full_path": current_path,
                    "item_count": f.Items.Count,
                    "unread_count": f.UnReadItemCount,
                }
                if depth < max_depth:
                    children = []
                    for i in range(f.Folders.Count):
                        try:
                            child = f.Folders.Item(i + 1)
                            children.append(walk(child, depth + 1, current_path))
                        except Exception as e:
                            logger.debug("Skipping folder %d: %s", i + 1, e)
                            continue
                    if children:
                        result["subfolders"] = children
                return result

            folders = []
            for i in range(start.Folders.Count):
                try:
                    child = start.Folders.Item(i + 1)
                    folders.append(walk(child, 1, base_path))
                except Exception as e:
                    logger.debug("Skipping root folder %d: %s", i + 1, e)
                    continue
            return json.dumps(folders, indent=2, default=str)

        try:
            return await self._bridge.call(_list, folder, max_depth, account)
        except Exception as e:
            return f"Error listing folders: {format_com_error(e)}"

    async def list_attachments(self, entry_id: str, account: str = "") -> str:
        def _list(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return err
            attachments = []
            for i in range(item.Attachments.Count):
                att = item.Attachments.Item(i + 1)
                attachments.append({
                    "index": i + 1,
                    "filename": att.FileName,
                    "size": att.Size,
                    "type": att.Type,
                })
            return json.dumps(attachments, indent=2, default=str)

        try:
            return await self._bridge.call(_list, entry_id, account)
        except Exception as e:
            return f"Error listing attachments: {format_com_error(e)}"

    async def save_attachment(
        self,
        entry_id: str,
        attachment_index: int,
        output_path: str,
        account: str = "",
    ) -> str:
        def _get_attachment_bytes(outlook, namespace, entry_id, attachment_index, account):
            """Get attachment filename and content from the COM thread."""
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_MAIL, "mail item"):
                return {"error": err}
            if attachment_index < 1 or attachment_index > item.Attachments.Count:
                return {"error": f"Attachment index {attachment_index} out of range (1-{item.Attachments.Count})"}
            att = item.Attachments.Item(attachment_index)
            # Save to a temp location on the COM thread (required by Outlook API),
            # then return the path for the caller to move/rename.
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), att.FileName)
            att.SaveAsFile(tmp)
            return {"tmp_path": tmp, "filename": att.FileName}

        try:
            result = await self._bridge.call(
                _get_attachment_bytes, entry_id, attachment_index, account
            )
            if "error" in result:
                return result["error"]
            # File I/O (move from tmp to final path) happens here, off the COM thread.
            import shutil
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            shutil.move(result["tmp_path"], output_path)
            return f"Saved '{result['filename']}' to {output_path}"
        except Exception as e:
            return f"Error saving attachment: {format_com_error(e)}"
```

- [ ] **Step 4: Run email tests — expect PASS**

```bash
uv run pytest tests/test_service.py -v
```
Expected: `6 passed` (2 skeleton + 4 email)

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/service.py tests/test_service.py
git commit -m "Add email tools to OutlookService"
```

---

## Task 4: Migrate calendar tools to OutlookService

**Files:**
- Modify: `src/outlook_desktop_mcp/service.py` (add 8 calendar methods)
- Modify: `tests/test_service.py` (add calendar tests)

- [ ] **Step 1: Add calendar tests**

Append to `tests/test_service.py`:

```python
# ---- Calendar ----

class TestListEvents:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.list_events())
        mock_bridge.call.assert_called_once()

    def test_returns_json(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value='[{"entry_id":"E1"}]')
        result = asyncio.run(svc.list_events())
        assert json.loads(result)[0]["entry_id"] == "E1"


class TestCreateEvent:
    def test_calls_bridge(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value='{"status":"created"}')
        asyncio.run(svc.create_event("Standup", "2026-06-01 09:00", "2026-06-01 09:30"))
        mock_bridge.call.assert_called_once()


class TestRespondToMeeting:
    def test_accept(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value="Accepted meeting 'Standup'")
        result = asyncio.run(svc.respond_to_meeting("ENTRYID", "accept"))
        assert "Accepted" in result

    def test_invalid_response_returns_error(self, svc, mock_bridge):
        result = asyncio.run(svc.respond_to_meeting("ENTRYID", "maybe"))
        assert "Error" in result
        mock_bridge.call.assert_not_called()
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_service.py::TestListEvents -v
```
Expected: `AttributeError: 'OutlookService' object has no attribute 'list_events'`

- [ ] **Step 3: Add calendar methods to `service.py`**

Add to `OutlookService` in `service.py`:

```python
    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------

    async def list_events(
        self,
        start_date: str = "",
        end_date: str = "",
        count: int = 20,
        account: str = "",
    ) -> str:
        def _list(outlook, namespace, start_date, end_date, count, account):
            count = min(max(1, count), MAX_ITEM_COUNT)
            store = OutlookService._require_store(namespace, account)
            calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            start = OutlookService._parse_date(start_date) if start_date else datetime.now()
            end = OutlookService._parse_date(end_date) if end_date else start + timedelta(days=7)
            restrict = (
                f"[Start] >= '{start.strftime(OUTLOOK_DATE_FMT)}' "
                f"AND [Start] <= '{end.strftime(OUTLOOK_DATE_FMT)}'"
            )
            filtered = items.Restrict(restrict)
            results = []
            n = 0
            for item in filtered:
                n += 1
                try:
                    results.append(format_event_summary(item))
                except Exception as e:
                    logger.debug("Skipping event item: %s", e)
                    continue
                if n >= count:
                    break
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list, start_date, end_date, count, account)
        except Exception as e:
            return f"Error listing events: {format_com_error(e)}"

    async def get_event(self, entry_id: str, account: str = "") -> str:
        def _get(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            return json.dumps(format_event_full(item), indent=2, default=str)

        try:
            return await self._bridge.call(_get, entry_id, account)
        except Exception as e:
            return f"Error reading event: {format_com_error(e)}"

    async def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        location: str = "",
        body: str = "",
        all_day: bool = False,
        reminder_minutes: int = 15,
        account: str = "",
    ) -> str:
        def _create(outlook, namespace, subject, start, end, location, body,
                    all_day, reminder_minutes, account):
            appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
            if account:
                store = OutlookService._require_store(namespace, account)
                cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
                appt = appt.Move(cal)
            appt.Subject = subject
            appt.Start = start
            appt.End = end
            if location:
                appt.Location = location
            if body:
                appt.Body = body
            appt.AllDayEvent = all_day
            if reminder_minutes > 0:
                appt.ReminderSet = True
                appt.ReminderMinutesBeforeStart = reminder_minutes
            else:
                appt.ReminderSet = False
            appt.Save()
            return json.dumps({
                "status": "created",
                "subject": appt.Subject,
                "start": str(appt.Start),
                "end": str(appt.End),
                "entry_id": appt.EntryID,
            }, indent=2, default=str)

        try:
            return await self._bridge.call(
                _create, subject, start, end, location, body, all_day, reminder_minutes, account
            )
        except Exception as e:
            return f"Error creating event: {format_com_error(e)}"

    async def create_meeting(
        self,
        subject: str,
        start: str,
        end: str,
        required_attendees: str,
        location: str = "",
        body: str = "",
        optional_attendees: str = "",
        account: str = "",
    ) -> str:
        def _create(outlook, namespace, subject, start, end, required_attendees,
                    location, body, optional_attendees, account):
            appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
            if account:
                store = OutlookService._require_store(namespace, account)
                for acc in outlook.Session.Accounts:
                    if acc.DeliveryStore.StoreID == store.StoreID:
                        appt._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, acc))
                        break
            appt.Subject = subject
            appt.Start = start
            appt.End = end
            appt.MeetingStatus = OL_MEETING
            if location:
                appt.Location = location
            if body:
                appt.Body = body
            for addr in required_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_REQUIRED
            if optional_attendees:
                for addr in optional_attendees.split(";"):
                    addr = addr.strip()
                    if addr:
                        recip = appt.Recipients.Add(addr)
                        recip.Type = OL_OPTIONAL
            appt.Recipients.ResolveAll()
            appt.Send()
            return (
                f"Meeting '{subject}' created and invitations sent to {required_attendees}"
            )

        try:
            return await self._bridge.call(
                _create, subject, start, end, required_attendees, location, body,
                optional_attendees, account
            )
        except Exception as e:
            return f"Error creating meeting: {format_com_error(e)}"

    async def update_event(
        self,
        entry_id: str,
        subject: str = "",
        start: str = "",
        end: str = "",
        location: str = "",
        body: str = "",
        account: str = "",
    ) -> str:
        def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_APPOINTMENT, "calendar event"):
                return err
            if subject:
                item.Subject = subject
            if start:
                item.Start = start
            if end:
                item.End = end
            if location:
                item.Location = location
            if body:
                item.Body = body
            item.Save()
            return f"Updated event: '{item.Subject}'"

        try:
            return await self._bridge.call(_update, entry_id, subject, start, end, location, body, account)
        except Exception as e:
            return f"Error updating event: {format_com_error(e)}"

    async def delete_event(self, entry_id: str, account: str = "") -> str:
        def _delete(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_APPOINTMENT, "calendar event"):
                return err
            subject = item.Subject
            item.Delete()
            return f"Deleted event: '{subject}'"

        try:
            return await self._bridge.call(_delete, entry_id, account)
        except Exception as e:
            return f"Error deleting event: {format_com_error(e)}"

    async def respond_to_meeting(
        self,
        entry_id: str,
        response: str,
        account: str = "",
    ) -> str:
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        if response.lower() not in response_map:
            return f"Error: response must be 'accept', 'decline', or 'tentative'. Got: '{response}'"

        response_code = response_map[response.lower()]

        def _respond(outlook, namespace, entry_id, response_code, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_APPOINTMENT, "calendar event"):
                return err
            subject = item.Subject
            reply = item.Respond(response_code, True)
            if reply:
                reply.Send()
            return f"{'Accepted' if response_code == OL_RESPONSE_ACCEPTED else 'Declined' if response_code == OL_RESPONSE_DECLINED else 'Tentatively accepted'} meeting '{subject}'"

        try:
            return await self._bridge.call(_respond, entry_id, response_code, account)
        except Exception as e:
            return f"Error responding to meeting: {format_com_error(e)}"

    async def search_events(
        self,
        query: str,
        start_date: str = "",
        end_date: str = "",
        count: int = 10,
        account: str = "",
    ) -> str:
        def _search(outlook, namespace, query, start_date, end_date, count, account):
            count = min(max(1, count), MAX_ITEM_COUNT)
            store = OutlookService._require_store(namespace, account)
            calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            start = OutlookService._parse_date(start_date) if start_date else datetime.now()
            end = OutlookService._parse_date(end_date) if end_date else start + timedelta(days=365)
            restrict = (
                f"[Start] >= '{start.strftime(OUTLOOK_DATE_FMT)}' "
                f"AND [Start] <= '{end.strftime(OUTLOOK_DATE_FMT)}'"
            )
            filtered = items.Restrict(restrict)
            query_lower = query.lower()
            results = []
            for item in filtered:
                if query_lower in (item.Subject or "").lower() or query_lower in (item.Body or "").lower():
                    try:
                        results.append(format_event_summary(item))
                    except Exception as e:
                        logger.debug("Skipping event in search: %s", e)
                        continue
                if len(results) >= count:
                    break
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_search, query, start_date, end_date, count, account)
        except Exception as e:
            return f"Error searching events: {format_com_error(e)}"
```

- [ ] **Step 4: Run calendar tests — expect PASS**

```bash
uv run pytest tests/test_service.py -v
```
Expected: `10 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/service.py tests/test_service.py
git commit -m "Add calendar tools to OutlookService"
```

---

## Task 5: Migrate task + misc tools to OutlookService

**Files:**
- Modify: `src/outlook_desktop_mcp/service.py` (add tasks + categories + rules + OOO)
- Modify: `tests/test_service.py` (add task/misc tests)

- [ ] **Step 1: Add task/misc tests**

Append to `tests/test_service.py`:

```python
# ---- Tasks ----

class TestListTasks:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.list_tasks())
        mock_bridge.call.assert_called_once()


class TestCreateTask:
    def test_calls_bridge(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value='{"status":"created"}')
        asyncio.run(svc.create_task("Fix bug"))
        mock_bridge.call.assert_called_once()


class TestCompleteTask:
    def test_calls_bridge(self, svc, mock_bridge):
        mock_bridge.call = AsyncMock(return_value="Task completed")
        asyncio.run(svc.complete_task("TASKID"))
        mock_bridge.call.assert_called_once()


# ---- Misc ----

class TestListCategories:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.list_categories())
        mock_bridge.call.assert_called_once()


class TestGetOutOfOffice:
    def test_calls_bridge(self, svc, mock_bridge):
        asyncio.run(svc.get_out_of_office())
        mock_bridge.call.assert_called_once()
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_service.py::TestListTasks -v
```
Expected: `AttributeError: 'OutlookService' object has no attribute 'list_tasks'`

- [ ] **Step 3: Add task + misc methods to `service.py`**

Add to `OutlookService` in `service.py`:

```python
    # ------------------------------------------------------------------
    # Tasks
    # ------------------------------------------------------------------

    async def list_tasks(self, account: str = "") -> str:
        def _list(outlook, namespace, account):
            store = OutlookService._require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            items = tasks_folder.Items
            results = []
            for i in range(min(items.Count, MAX_ITEM_COUNT)):
                try:
                    results.append(format_task_summary(items.Item(i + 1)))
                except Exception as e:
                    logger.debug("Skipping task %d: %s", i + 1, e)
                    continue
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list, account)
        except Exception as e:
            return f"Error listing tasks: {format_com_error(e)}"

    async def get_task(self, entry_id: str, account: str = "") -> str:
        def _get(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            return json.dumps(format_task_full(item), indent=2, default=str)

        try:
            return await self._bridge.call(_get, entry_id, account)
        except Exception as e:
            return f"Error reading task: {format_com_error(e)}"

    async def create_task(
        self,
        subject: str,
        body: str = "",
        due_date: str = "",
        start_date: str = "",
        priority: int = 1,
        account: str = "",
    ) -> str:
        def _create(outlook, namespace, subject, body, due_date, start_date, priority, account):
            task = outlook.CreateItem(OL_TASK_ITEM)
            if account:
                store = OutlookService._require_store(namespace, account)
                tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
                task = task.Move(tasks_folder)
            task.Subject = subject
            if body:
                task.Body = body
            if due_date:
                task.DueDate = due_date
            if start_date:
                task.StartDate = start_date
            task.Importance = max(0, min(2, priority))
            task.Save()
            return json.dumps({
                "status": "created",
                "subject": task.Subject,
                "entry_id": task.EntryID,
            }, indent=2, default=str)

        try:
            return await self._bridge.call(
                _create, subject, body, due_date, start_date, priority, account
            )
        except Exception as e:
            return f"Error creating task: {format_com_error(e)}"

    async def complete_task(self, entry_id: str, account: str = "") -> str:
        def _complete(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_TASK, "task"):
                return err
            subject = item.Subject
            item.Status = OL_TASK_COMPLETE
            item.PercentComplete = 100
            item.Complete = True
            item.Save()
            return f"Task completed: '{subject}'"

        try:
            return await self._bridge.call(_complete, entry_id, account)
        except Exception as e:
            return f"Error completing task: {format_com_error(e)}"

    async def delete_task(self, entry_id: str, account: str = "") -> str:
        def _delete(outlook, namespace, entry_id, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            if err := OutlookService._check_item_class(item, OL_CLASS_TASK, "task"):
                return err
            subject = item.Subject
            item.Delete()
            return f"Deleted task: '{subject}'"

        try:
            return await self._bridge.call(_delete, entry_id, account)
        except Exception as e:
            return f"Error deleting task: {format_com_error(e)}"

    # ------------------------------------------------------------------
    # Categories
    # ------------------------------------------------------------------

    async def list_categories(self, account: str = "") -> str:
        def _list(outlook, namespace):
            cats = namespace.Categories
            results = [
                {"name": cats.Item(i + 1).Name, "color": cats.Item(i + 1).Color}
                for i in range(cats.Count)
            ]
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list)
        except Exception as e:
            return f"Error listing categories: {format_com_error(e)}"

    async def set_category(
        self,
        entry_id: str,
        category: str,
        account: str = "",
    ) -> str:
        def _set(outlook, namespace, entry_id, category, account):
            item = OutlookService._get_item(namespace, entry_id, account)
            existing = item.Categories or ""
            cats = [c.strip() for c in existing.split(";") if c.strip()]
            if category not in cats:
                cats.append(category)
            item.Categories = "; ".join(cats)
            item.Save()
            return f"Category '{category}' set on '{item.Subject}'"

        try:
            return await self._bridge.call(_set, entry_id, category, account)
        except Exception as e:
            return f"Error setting category: {format_com_error(e)}"

    # ------------------------------------------------------------------
    # Rules
    # ------------------------------------------------------------------

    async def list_rules(self, account: str = "") -> str:
        def _list(outlook, namespace, account):
            store = OutlookService._require_store(namespace, account)
            rules = store.GetRules()
            results = []
            for i in range(rules.Count):
                rule = rules.Item(i + 1)
                results.append({
                    "name": rule.Name,
                    "enabled": rule.Enabled,
                    "execution_order": rule.ExecutionOrder,
                })
            return json.dumps(results, indent=2, default=str)

        try:
            return await self._bridge.call(_list, account)
        except Exception as e:
            return f"Error listing rules: {format_com_error(e)}"

    async def toggle_rule(
        self,
        rule_name: str,
        enabled: bool,
        account: str = "",
    ) -> str:
        def _toggle(outlook, namespace, rule_name, enabled, account):
            store = OutlookService._require_store(namespace, account)
            rules = store.GetRules()
            for i in range(rules.Count):
                rule = rules.Item(i + 1)
                if rule.Name.lower() == rule_name.lower():
                    rule.Enabled = enabled
                    rules.Save()
                    state = "enabled" if enabled else "disabled"
                    return f"Rule '{rule.Name}' {state}"
            return f"Error: Rule '{rule_name}' not found."

        try:
            return await self._bridge.call(_toggle, rule_name, enabled, account)
        except Exception as e:
            return f"Error toggling rule: {format_com_error(e)}"

    # ------------------------------------------------------------------
    # Out of Office
    # ------------------------------------------------------------------

    async def get_out_of_office(self, account: str = "") -> str:
        def _get(outlook, namespace, account):
            store = OutlookService._require_store(namespace, account)
            try:
                prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
                oof_state = store.PropertyAccessor.GetProperty(prop_tag)
                return json.dumps({
                    "out_of_office": bool(oof_state),
                    "status": "on" if oof_state else "off",
                }, indent=2)
            except Exception:
                return json.dumps({
                    "out_of_office": None,
                    "status": "unknown",
                    "note": "Could not read OOF property. Check Outlook settings directly.",
                }, indent=2)

        try:
            return await self._bridge.call(_get, account)
        except Exception as e:
            return f"Error checking OOF status: {format_com_error(e)}"
```

- [ ] **Step 4: Run all service tests — expect PASS**

```bash
uv run pytest tests/test_service.py -v
```
Expected: `15 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/service.py tests/test_service.py
git commit -m "Add task, category, rules, OOO tools to OutlookService"
```

---

## Task 6: Refactor `server.py` to thin wrappers

**Files:**
- Modify: `src/outlook_desktop_mcp/server.py`

- [ ] **Step 1: Replace server.py top section**

After the existing imports in `server.py`, add:
```python
from outlook_desktop_mcp.service import OutlookService
```

After `bridge = OutlookBridge()`, add:
```python
svc = OutlookService(bridge)
```

- [ ] **Step 2: Replace each tool function body**

For each `@mcp.tool()` function, replace the entire inner function + `bridge.call(...)` + `except` block with a single `await svc.<method>(...)` delegation. Example:

```python
# BEFORE (list_emails — ~40 lines with inner function):
@mcp.tool()
async def list_emails(folder="inbox", count=10, ...) -> str:
    def _list(outlook, namespace, ...):
        ...40 lines...
    try:
        return await bridge.call(_list, ...)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"

# AFTER (4 lines):
@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.
    [keep the existing docstring unchanged]
    """
    return await svc.list_emails(folder, count, unread_only, start_date, end_date, account)
```

Apply this pattern to all 30 tool functions. Keep the `@mcp.tool()` decorator and docstrings unchanged. Remove the inner functions and `bridge.call(...)` blocks entirely.

- [ ] **Step 3: Verify server.py still imports cleanly**

```bash
uv run python -c "import outlook_desktop_mcp.server; print('server ok')"
```
Expected: `server ok`

- [ ] **Step 4: Run all tests**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_service.py -v
```
Expected: `47 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server.py
git commit -m "Refactor server.py to thin @mcp.tool wrappers over OutlookService

server.py is now ~200 lines of 4-6 line wrapper functions.
All business logic lives in service.py."
```

---

## Task 7: Create `cli.py`

**Files:**
- Create: `src/outlook_desktop_mcp/cli.py`
- Create: `tests/test_cli.py`

- [ ] **Step 1: Write CLI tests first**

Create `tests/test_cli.py`:

```python
"""CLI tests — no Outlook required. OutlookService and OutlookBridge are mocked."""
import json
import asyncio
import pytest
from typer.testing import CliRunner
from unittest.mock import AsyncMock, MagicMock, patch

runner = CliRunner()


def _mock_svc(return_value: str = "[]"):
    """Return a mock OutlookService where every method returns return_value."""
    svc = MagicMock()
    for attr in [
        "list_accounts", "list_folders",
        "list_emails", "read_email", "send_email", "reply_email",
        "move_email", "mark_as_read", "mark_as_unread", "search_emails",
        "list_attachments", "save_attachment",
        "list_events", "get_event", "create_event", "create_meeting",
        "update_event", "delete_event", "respond_to_meeting", "search_events",
        "list_tasks", "get_task", "create_task", "complete_task", "delete_task",
        "list_categories", "set_category",
        "list_rules", "toggle_rule",
        "get_out_of_office",
    ]:
        setattr(svc, attr, AsyncMock(return_value=return_value))
    return svc


def _invoke(args: list[str], svc=None, return_value: str = "[]"):
    """Invoke the CLI with a mocked service."""
    from outlook_desktop_mcp.cli import app
    mock_svc = svc or _mock_svc(return_value)
    mock_bridge = MagicMock()
    mock_bridge.start = MagicMock()
    mock_bridge.stop = MagicMock()
    with patch("outlook_desktop_mcp.cli.OutlookBridge", return_value=mock_bridge), \
         patch("outlook_desktop_mcp.cli.OutlookService", return_value=mock_svc):
        return runner.invoke(app, args)


# ---- accounts ----

def test_accounts_exit_0():
    result = _invoke(["accounts"], return_value='[{"display_name":"Test"}]')
    assert result.exit_code == 0


def test_accounts_outputs_json():
    result = _invoke(["accounts"], return_value='[{"display_name":"Work"}]')
    assert "display_name" in result.output


# ---- email list ----

def test_email_list_exit_0():
    result = _invoke(["email", "list"])
    assert result.exit_code == 0


def test_email_list_passes_folder_option():
    svc = _mock_svc()
    _invoke(["email", "list", "--folder", "sent"], svc=svc)
    svc.list_emails.assert_called_once()
    call_args = svc.list_emails.call_args
    assert call_args.kwargs.get("folder") == "sent" or call_args.args[0] == "sent"


def test_email_list_passes_count_option():
    svc = _mock_svc()
    _invoke(["email", "list", "--count", "5"], svc=svc)
    svc.list_emails.assert_called_once()


def test_email_list_error_exits_1():
    svc = _mock_svc()
    svc.list_emails = AsyncMock(side_effect=Exception("COM error"))
    result = _invoke(["email", "list"], svc=svc)
    assert result.exit_code == 1


# ---- email send ----

def test_email_send_exit_0():
    result = _invoke([
        "email", "send",
        "--to", "a@b.com",
        "--subject", "Hi",
        "--body", "Hello"
    ], return_value="Email sent: 'Hi' to a@b.com")
    assert result.exit_code == 0


def test_email_send_missing_required_args_exits_nonzero():
    result = _invoke(["email", "send"])
    assert result.exit_code != 0


# ---- calendar list ----

def test_calendar_list_exit_0():
    result = _invoke(["calendar", "list"])
    assert result.exit_code == 0


def test_calendar_respond_accept():
    svc = _mock_svc(return_value="Accepted meeting 'Standup'")
    result = _invoke(["calendar", "respond", "ENTRYID", "--response", "accept"], svc=svc)
    assert result.exit_code == 0
    svc.respond_to_meeting.assert_called_once()


# ---- tasks ----

def test_tasks_list_exit_0():
    result = _invoke(["tasks", "list"])
    assert result.exit_code == 0


def test_tasks_create_exit_0():
    result = _invoke(
        ["tasks", "create", "--subject", "Fix bug"],
        return_value='{"status":"created"}'
    )
    assert result.exit_code == 0


# ---- error handling ----

def test_runtime_error_exits_2():
    svc = _mock_svc()
    svc.list_emails = AsyncMock(side_effect=RuntimeError("Bridge not running"))
    result = _invoke(["email", "list"], svc=svc)
    assert result.exit_code == 2


def test_error_json_on_stderr():
    svc = _mock_svc()
    svc.list_emails = AsyncMock(side_effect=Exception("Outlook error"))
    result = _invoke(["email", "list"], svc=svc)
    assert "error" in result.output or result.exit_code == 1
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_cli.py -v
```
Expected: `ModuleNotFoundError: No module named 'outlook_desktop_mcp.cli'`

- [ ] **Step 3: Create `src/outlook_desktop_mcp/cli.py`**

```python
"""
outlook CLI
===========
Typer-based CLI exposing all 30 OutlookService operations as grouped subcommands.

Usage:
    outlook --help
    outlook email list --folder inbox --count 20
    outlook calendar list --start 2026-06-01 --end 2026-06-07
    outlook tasks create --subject "Fix bug"

Output: JSON to stdout (exit 0), error JSON to stderr (exit 1 or 2).
"""
import asyncio
import json
import sys
from typing import Optional

import typer

from outlook_desktop_mcp.com_bridge import OutlookBridge
from outlook_desktop_mcp.service import OutlookService

app = typer.Typer(
    name="outlook",
    help="Outlook Desktop CLI — access email, calendar, tasks, and more.",
    no_args_is_help=True,
)
email_app = typer.Typer(help="Email operations", no_args_is_help=True)
calendar_app = typer.Typer(help="Calendar operations", no_args_is_help=True)
tasks_app = typer.Typer(help="Task operations", no_args_is_help=True)
folders_app = typer.Typer(help="Folder operations", no_args_is_help=True)
categories_app = typer.Typer(help="Category operations", no_args_is_help=True)
rules_app = typer.Typer(help="Mail rule operations", no_args_is_help=True)

app.add_typer(email_app, name="email")
app.add_typer(calendar_app, name="calendar")
app.add_typer(tasks_app, name="tasks")
app.add_typer(folders_app, name="folders")
app.add_typer(categories_app, name="categories")
app.add_typer(rules_app, name="rules")


# ------------------------------------------------------------------
# Bridge lifecycle — shared via Typer context
# ------------------------------------------------------------------

@app.callback()
def _startup(ctx: typer.Context) -> None:
    """Start the Outlook COM bridge before any command runs."""
    if ctx.invoked_subcommand is None:
        return
    bridge = OutlookBridge()
    bridge.start()
    ctx.ensure_object(dict)
    ctx.obj = OutlookService(bridge)
    ctx.call_on_close(bridge.stop)


def _svc(ctx: typer.Context) -> OutlookService:
    return ctx.obj


# ------------------------------------------------------------------
# Output helpers
# ------------------------------------------------------------------

def _ok(result: str) -> None:
    typer.echo(result)


def _run(ctx: typer.Context, coro) -> None:
    """Run an async service coroutine, print result, handle errors."""
    try:
        result = asyncio.run(coro)
        _ok(result)
    except RuntimeError as e:
        typer.echo(json.dumps({"error": str(e)}), err=True)
        raise typer.Exit(2)
    except Exception as e:
        typer.echo(json.dumps({"error": str(e)}), err=True)
        raise typer.Exit(1)


# ------------------------------------------------------------------
# Accounts + Folders (top-level)
# ------------------------------------------------------------------

@app.command("accounts")
def cmd_accounts(ctx: typer.Context) -> None:
    """List all Outlook accounts (stores)."""
    _run(ctx, _svc(ctx).list_accounts())


@folders_app.command("list")
def folders_list(
    ctx: typer.Context,
    folder: str = typer.Option("", "--folder", "-f", help="Starting folder path"),
    depth: int = typer.Option(3, "--depth", "-d", help="Max recursion depth"),
    account: str = typer.Option("", "--account", "-a", help="Account substring"),
) -> None:
    """List folder hierarchy."""
    _run(ctx, _svc(ctx).list_folders(folder=folder, max_depth=depth, account=account))


# ------------------------------------------------------------------
# Email
# ------------------------------------------------------------------

@email_app.command("list")
def email_list(
    ctx: typer.Context,
    folder: str = typer.Option("inbox", "--folder", "-f"),
    count: int = typer.Option(10, "--count", "-n"),
    unread: bool = typer.Option(False, "--unread"),
    start: str = typer.Option("", "--start"),
    end: str = typer.Option("", "--end"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List emails from a folder."""
    _run(ctx, _svc(ctx).list_emails(
        folder=folder, count=count, unread_only=unread,
        start_date=start, end_date=end, account=account,
    ))


@email_app.command("read")
def email_read(
    ctx: typer.Context,
    entry_id: str = typer.Argument(default="", help="Outlook EntryID"),
    subject: str = typer.Option("", "--subject", "-s", help="Subject search substring"),
    folder: str = typer.Option("inbox", "--folder", "-f"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Read a specific email by EntryID or subject search."""
    _run(ctx, _svc(ctx).read_email(
        entry_id=entry_id, subject_search=subject, folder=folder, account=account,
    ))


@email_app.command("send")
def email_send(
    ctx: typer.Context,
    to: str = typer.Option(..., "--to", help="Recipient(s), semicolon-separated"),
    subject: str = typer.Option(..., "--subject", "-s"),
    body: str = typer.Option(..., "--body", "-b"),
    cc: str = typer.Option("", "--cc"),
    bcc: str = typer.Option("", "--bcc"),
    html: str = typer.Option("", "--html", help="HTML body"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Send an email."""
    _run(ctx, _svc(ctx).send_email(
        to=to, subject=subject, body=body, cc=cc, bcc=bcc,
        html_body=html, account=account,
    ))


@email_app.command("reply")
def email_reply(
    ctx: typer.Context,
    entry_id: str = typer.Argument(..., help="EntryID of email to reply to"),
    body: str = typer.Option(..., "--body", "-b"),
    reply_all: bool = typer.Option(False, "--reply-all"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Reply to an email."""
    _run(ctx, _svc(ctx).reply_email(
        entry_id=entry_id, body=body, reply_all=reply_all, account=account,
    ))


@email_app.command("move")
def email_move(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    folder: str = typer.Option("archive", "--folder", "-f"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Move an email to a folder."""
    _run(ctx, _svc(ctx).move_email(
        entry_id=entry_id, target_folder=folder, account=account,
    ))


@email_app.command("mark-read")
def email_mark_read(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Mark an email as read."""
    _run(ctx, _svc(ctx).mark_as_read(entry_id=entry_id, account=account))


@email_app.command("mark-unread")
def email_mark_unread(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Mark an email as unread."""
    _run(ctx, _svc(ctx).mark_as_unread(entry_id=entry_id, account=account))


@email_app.command("search")
def email_search(
    ctx: typer.Context,
    query: str = typer.Argument(..., help="Search term"),
    folder: str = typer.Option("inbox", "--folder", "-f"),
    count: int = typer.Option(10, "--count", "-n"),
    start: str = typer.Option("", "--start"),
    end: str = typer.Option("", "--end"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Search emails by subject and body."""
    _run(ctx, _svc(ctx).search_emails(
        query=query, folder=folder, count=count,
        start_date=start, end_date=end, account=account,
    ))


@email_app.command("attachments")
def email_attachments(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List attachments on an email."""
    _run(ctx, _svc(ctx).list_attachments(entry_id=entry_id, account=account))


@email_app.command("save-attachment")
def email_save_attachment(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    index: int = typer.Argument(..., help="Attachment index (1-based)"),
    output: str = typer.Option(..., "--output", "-o", help="Output file path"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Save an attachment to disk."""
    _run(ctx, _svc(ctx).save_attachment(
        entry_id=entry_id, attachment_index=index,
        output_path=output, account=account,
    ))


# ------------------------------------------------------------------
# Calendar
# ------------------------------------------------------------------

@calendar_app.command("list")
def calendar_list(
    ctx: typer.Context,
    start: str = typer.Option("", "--start"),
    end: str = typer.Option("", "--end"),
    count: int = typer.Option(20, "--count", "-n"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List calendar events."""
    _run(ctx, _svc(ctx).list_events(
        start_date=start, end_date=end, count=count, account=account,
    ))


@calendar_app.command("get")
def calendar_get(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Get full details of a calendar event."""
    _run(ctx, _svc(ctx).get_event(entry_id=entry_id, account=account))


@calendar_app.command("create")
def calendar_create(
    ctx: typer.Context,
    subject: str = typer.Option(..., "--subject", "-s"),
    start: str = typer.Option(..., "--start", help="ISO 8601 datetime"),
    end: str = typer.Option(..., "--end", help="ISO 8601 datetime"),
    location: str = typer.Option("", "--location", "-l"),
    body: str = typer.Option("", "--body", "-b"),
    all_day: bool = typer.Option(False, "--all-day"),
    reminder: int = typer.Option(15, "--reminder", help="Reminder minutes before start"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Create a calendar appointment."""
    _run(ctx, _svc(ctx).create_event(
        subject=subject, start=start, end=end, location=location,
        body=body, all_day=all_day, reminder_minutes=reminder, account=account,
    ))


@calendar_app.command("meeting")
def calendar_meeting(
    ctx: typer.Context,
    subject: str = typer.Option(..., "--subject", "-s"),
    start: str = typer.Option(..., "--start"),
    end: str = typer.Option(..., "--end"),
    required: str = typer.Option(..., "--required", help="Required attendees, semicolon-separated"),
    optional: str = typer.Option("", "--optional"),
    location: str = typer.Option("", "--location", "-l"),
    body: str = typer.Option("", "--body", "-b"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Create a meeting and send invitations."""
    _run(ctx, _svc(ctx).create_meeting(
        subject=subject, start=start, end=end, required_attendees=required,
        optional_attendees=optional, location=location, body=body, account=account,
    ))


@calendar_app.command("update")
def calendar_update(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    subject: str = typer.Option("", "--subject", "-s"),
    start: str = typer.Option("", "--start"),
    end: str = typer.Option("", "--end"),
    location: str = typer.Option("", "--location", "-l"),
    body: str = typer.Option("", "--body", "-b"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Update a calendar event."""
    _run(ctx, _svc(ctx).update_event(
        entry_id=entry_id, subject=subject, start=start, end=end,
        location=location, body=body, account=account,
    ))


@calendar_app.command("delete")
def calendar_delete(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Delete a calendar event."""
    _run(ctx, _svc(ctx).delete_event(entry_id=entry_id, account=account))


@calendar_app.command("respond")
def calendar_respond(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    response: str = typer.Option(..., "--response", help="accept | decline | tentative"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Respond to a meeting invitation."""
    _run(ctx, _svc(ctx).respond_to_meeting(
        entry_id=entry_id, response=response, account=account,
    ))


@calendar_app.command("search")
def calendar_search(
    ctx: typer.Context,
    query: str = typer.Argument(...),
    start: str = typer.Option("", "--start"),
    end: str = typer.Option("", "--end"),
    count: int = typer.Option(10, "--count", "-n"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Search calendar events by subject or body."""
    _run(ctx, _svc(ctx).search_events(
        query=query, start_date=start, end_date=end, count=count, account=account,
    ))


# ------------------------------------------------------------------
# Tasks
# ------------------------------------------------------------------

@tasks_app.command("list")
def tasks_list(
    ctx: typer.Context,
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List tasks."""
    _run(ctx, _svc(ctx).list_tasks(account=account))


@tasks_app.command("get")
def tasks_get(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Get full details of a task."""
    _run(ctx, _svc(ctx).get_task(entry_id=entry_id, account=account))


@tasks_app.command("create")
def tasks_create(
    ctx: typer.Context,
    subject: str = typer.Option(..., "--subject", "-s"),
    body: str = typer.Option("", "--body", "-b"),
    due: str = typer.Option("", "--due", help="Due date ISO 8601"),
    start: str = typer.Option("", "--start", help="Start date ISO 8601"),
    priority: int = typer.Option(1, "--priority", help="0=low 1=normal 2=high"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Create a task."""
    _run(ctx, _svc(ctx).create_task(
        subject=subject, body=body, due_date=due,
        start_date=start, priority=priority, account=account,
    ))


@tasks_app.command("complete")
def tasks_complete(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Mark a task as complete."""
    _run(ctx, _svc(ctx).complete_task(entry_id=entry_id, account=account))


@tasks_app.command("delete")
def tasks_delete(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Delete a task."""
    _run(ctx, _svc(ctx).delete_task(entry_id=entry_id, account=account))


# ------------------------------------------------------------------
# Categories
# ------------------------------------------------------------------

@categories_app.command("list")
def categories_list(
    ctx: typer.Context,
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List all Outlook categories."""
    _run(ctx, _svc(ctx).list_categories(account=account))


@categories_app.command("set")
def categories_set(
    ctx: typer.Context,
    entry_id: str = typer.Argument(...),
    category: str = typer.Argument(...),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Set a category on an item."""
    _run(ctx, _svc(ctx).set_category(
        entry_id=entry_id, category=category, account=account,
    ))


# ------------------------------------------------------------------
# Rules
# ------------------------------------------------------------------

@rules_app.command("list")
def rules_list(
    ctx: typer.Context,
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """List mail rules."""
    _run(ctx, _svc(ctx).list_rules(account=account))


@rules_app.command("toggle")
def rules_toggle(
    ctx: typer.Context,
    name: str = typer.Argument(..., help="Rule name"),
    enable: bool = typer.Option(..., "--enable/--disable"),
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Enable or disable a mail rule."""
    _run(ctx, _svc(ctx).toggle_rule(rule_name=name, enabled=enable, account=account))


# ------------------------------------------------------------------
# Out of Office
# ------------------------------------------------------------------

@app.command("out-of-office")
def cmd_out_of_office(
    ctx: typer.Context,
    account: str = typer.Option("", "--account", "-a"),
) -> None:
    """Get out-of-office / auto-reply status."""
    _run(ctx, _svc(ctx).get_out_of_office(account=account))


# ------------------------------------------------------------------
# Entrypoint
# ------------------------------------------------------------------

def main() -> None:
    app()


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run CLI tests — expect PASS**

```bash
uv run pytest tests/test_cli.py -v
```
Expected: `16 passed`

- [ ] **Step 5: Run full test suite**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_service.py tests/test_cli.py -v
```
Expected: `63 passed`

- [ ] **Step 6: Commit**

```bash
git add src/outlook_desktop_mcp/cli.py tests/test_cli.py
git commit -m "Add outlook CLI with 30 Typer subcommands

Grouped by email, calendar, tasks, folders, categories, rules.
JSON output to stdout, error JSON to stderr, exit codes 0/1/2.
Bridge lifecycle owned by CLI callback (start/stop per invocation)."
```

---

## Task 8: Update `pyproject.toml` and smoke test

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Add `outlook` entry point and declare `typer` dep**

Edit `pyproject.toml`:

```toml
[project.scripts]
outlook-desktop-mcp = "outlook_desktop_mcp.entrypoint:main"
outlook = "outlook_desktop_mcp.cli:main"

[project.dependencies]
pywin32 = ">=310; sys_platform == 'win32'"
mcp = {version = ">=1.26.0", extras = ["cli"]}
typer = ">=0.9.0"
```

- [ ] **Step 2: Reinstall package in venv**

```bash
uv sync
```
Expected: no errors

- [ ] **Step 3: Smoke test `--help`**

```bash
uv run outlook --help
```
Expected: Shows `outlook` command groups: `email`, `calendar`, `tasks`, `folders`, `categories`, `rules`, `accounts`, `out-of-office`

```bash
uv run outlook email --help
```
Expected: Shows email subcommands: `list`, `read`, `send`, `reply`, `move`, `mark-read`, `mark-unread`, `search`, `attachments`, `save-attachment`

- [ ] **Step 4: Run full test suite one final time**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_service.py tests/test_cli.py -v
```
Expected: `63 passed`

- [ ] **Step 5: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "Add outlook entry point and declare typer dependency"
```

---

## Task 9: Merge CLI branch to main

- [ ] **Step 1: Final check**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_service.py tests/test_cli.py -v
```
Expected: `63 passed`

- [ ] **Step 2: Merge**

```bash
git checkout main
git merge --no-ff feat/outlook-cli
git branch -d feat/outlook-cli
git worktree remove .worktrees/cli
```

- [ ] **Step 3: Final verify on main**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_service.py tests/test_cli.py -v
uv run outlook --help
```
Expected: `63 passed`, help displayed.
