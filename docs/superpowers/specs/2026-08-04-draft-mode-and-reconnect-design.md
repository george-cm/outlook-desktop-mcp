# Design: Draft-safe mode + COM reconnect resilience

**Date:** 2026-08-04
**Status:** Approved by George
**Author:** George Murga + Devin

---

## Summary

Two independent fixes to `outlook-desktop-mcp` (Windows/COM path only):

1. **Draft-safe mode.** `send_email`, `reply_email`, `create_meeting` transmit
   immediately today with no way to review first. Add an opt-in env flag that
   makes them create a reviewable Outlook artifact (Drafts item / unsent
   meeting) instead, unless the caller explicitly confirms a send.
2. **Reconnect resilience.** The COM bridge dispatches `Outlook.Application`
   once at process startup and never again. If Outlook isn't running yet, the
   whole MCP server process dies before `mcp.run()`. If Outlook restarts while
   the server is up, every subsequent call fails forever with
   `0x800706BA RPC server unavailable` — the server is unusable until the
   entire MCP session is restarted.

Both diverge from `upstream/Aanerud/outlook-desktop-mcp` — (1) is opt-in via
env var to stay upstream-mergeable; (2) is a pure bug fix, no behavior change
when Outlook is healthy, safe to upstream as-is.

---

## Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Platform scope | Windows only (`server.py`, `com_bridge.py`) | Reconnect bug is COM-specific. macOS (`server_mac.py`) is stateless AppleScript per call, no equivalent bug. Draft-mode parity for macOS is a follow-up. |
| Draft-mode activation | Env var `OUTLOOK_MCP_DRAFT_MODE` (`"1"`/`"true"`/`"yes"`, case-insensitive) | Unset/false = today's exact behavior, byte-identical, upstream-safe default. Matches how MCP clients already configure server env (`claude mcp add --env`). |
| Send override | New `confirm_send: bool = False` param on the 3 gated tools, not a new tool | Additive, backward compatible signature. No tool-list bloat. No-op when draft mode is off. |
| `respond_to_meeting` / `delete_event` cancel path | **Out of scope**, left sending immediately | No Outlook equivalent of a reviewable "draft" for either (see Out of Scope). Forcing a fake draft step would either do nothing useful or silently drop the notification — worse than today. |
| Reconnect scope | Only disconnect-class HRESULTs trigger reconnect + one retry | Avoid masking real errors (bad entry_id, permission errors) as connection problems. |
| Startup behavior | `bridge.start()` no longer requires Outlook to already be running | Connection becomes lazy, attempted on first queued call. Server process survives "Outlook not open yet." |

---

## Part 1 — Draft-safe mode

### Gated tools

| Tool | `DRAFT_MODE` off (default) | `DRAFT_MODE` on, `confirm_send=False` (default) | `DRAFT_MODE` on, `confirm_send=True` |
|---|---|---|---|
| `send_email` | `.Send()` (unchanged) | `mail.Save()` → lands in Drafts, tool returns `entry_id` | `.Send()` |
| `reply_email` | `.Send()` (unchanged) | `item.Reply()` / `item.ReplyAll()` (always threaded — Outlook chains it to the source item) → prepend `body` → `.Save()` | `.Send()` |
| `create_meeting` | `.Send()` (unchanged) | build appointment + recipients, `Recipients.ResolveAll()`, `.Save()` only → sits on the organizer's calendar unsent; open it in Outlook and hit Send to actually invite | `.Send()` |

Threading for `reply_email` is guaranteed by construction: the draft is always
created via `source_item.Reply()`/`.ReplyAll()`, never a fresh `CreateItem`, so
it's structurally the same threaded-reply object Outlook would send — just
saved instead.

### Docstring changes

Each gated tool's docstring gets a new paragraph:

```
When the server is running in draft-safe mode (OUTLOOK_MCP_DRAFT_MODE=1),
this creates a Drafts item / unsent meeting instead of sending. Pass
confirm_send=true ONLY if the user has explicitly asked, in this turn,
to send/dispatch immediately rather than draft it.
```

### Out of Scope: `respond_to_meeting`, `delete_event` (meeting cancel)

Both call `.Send()`. Neither maps onto "reviewable draft":

- `respond_to_meeting`: whether skipping `.Send()` still updates the local
  calendar RSVP status is undocumented COM behavior — not verified, not
  assumed. Also low-stakes/high-frequency; forcing review adds friction
  without a real safety win.
- `delete_event` cancel path: Outlook has no "draft cancellation" artifact.
  Skipping `.Send()` either leaves the meeting un-cancelled on your calendar,
  or (if `.Delete()` is used instead) silently removes it with attendees
  never notified — strictly worse than today's behavior.

Both are left sending immediately, as they do today. Revisit only if a
concrete, verified Outlook mechanism for reviewable cancellations/responses
is found.

---

## Part 2 — COM reconnect resilience

### Current failure modes

1. `com_bridge.start()` → `_com_thread_main()` dispatches `Outlook.Application`
   synchronously; if it throws, `_init_error` is set and `start()` re-raises
   in `server.py:main()`, which runs **before** `mcp.run()`. The MCP stdio
   transport never comes up. The client sees a crashed server, not a usable
   session that can retry.
2. Once connected, `self._outlook` / `self._namespace` are cached for the
   life of the thread. If Outlook restarts, these become stale COM proxies.
   Every `bridge.call()` afterward raises `pythoncom.com_error` with HRESULT
   `0x800706BA` (RPC server unavailable) — permanently, since nothing ever
   re-dispatches.

### Fix

Split "COM thread alive" from "connected to Outlook":

- `start()` spawns the thread and waits only for `_alive` (thread running,
  ready to pull off the queue) — **not** for a successful Outlook dispatch.
  `start()` no longer fails because Outlook isn't open yet.
- Connection becomes lazy: before running a queued `func`, if
  `self._outlook is None`, attempt to connect
  (`Dispatch("Outlook.Application")` + `GetNamespace("MAPI")`). Failure →
  that call's `result_holder["error"]` is a clear
  `RuntimeError("Outlook Desktop is not running. Open it and retry.")` — the
  thread keeps running, the next call retries the connection.
- On a queued call raising one of the disconnect-class HRESULTs:

  | HRESULT | Meaning |
  |---|---|
  | `0x800706BA` | RPC server unavailable |
  | `0x800706BE` | RPC call failed |
  | `0x80010108` | RPC_E_DISCONNECTED |
  | `0x800401FD` | CO_E_OBJNOTCONNECTED |

  → mark disconnected (`self._outlook = None`), reconnect, retry the *same*
  `func` once. If the retry also fails, surface that error normally. Any
  other COM error (bad entry_id, permission denied, etc.) passes straight
  through — no reconnect attempt, no masking.
- `server.py:main()`: `bridge.start()` call is unchanged syntactically, but
  it now essentially always succeeds (barring the COM thread itself failing
  to spawn, which doesn't depend on Outlook). Outlook-not-running errors
  surface per-tool-call instead of killing the process.

### Non-goals

- No polling/health-check loop. Reconnect is purely reactive, attempted only
  when a call needs the connection and it's missing/dead. No background
  timers, no extra threads.
- No exponential backoff. `Dispatch()` failures aren't hot-looped — calls are
  driven by actual MCP tool invocations, not a tight retry loop.

---

## Testing (TDD)

### `tests/test_com_bridge.py` (extend existing mock pattern)

- `start()` succeeds with zero Outlook dispatch attempts (fake thread main
  that never connects).
- First `call()` triggers a connect attempt.
- Simulated RPC-unavailable (`0x800706BA`) mid-loop → reconnect attempted,
  failed call retried once, succeeds on retry.
- Reconnect itself fails (Outlook still down) → clear `RuntimeError`
  surfaced, COM thread still alive, next `call()` retries cleanly.
- A non-disconnect COM error (or plain `ValueError`) → no reconnect attempt,
  original error propagates unchanged (regression guard for existing
  traceback-preservation tests).

### New `tests/test_send_gating.py`

Pure logic against `MagicMock()` outlook/namespace/mail objects, no live COM:

- `DRAFT_MODE` off → `confirm_send` is irrelevant, `.Send()` always called,
  `.Save()` never called (regression guard: today's behavior unchanged).
- `DRAFT_MODE` on, `confirm_send=False` → `.Save()` called, `.Send()` never
  called.
- `DRAFT_MODE` on, `confirm_send=True` → `.Send()` called.
- `reply_email` in draft mode: assert the mock's `.Reply()`/`.ReplyAll()` was
  called on the *source* item (never a bare `CreateItem`), and the drafted
  body is the original quoted body with the new text prepended.
- `create_meeting` in draft mode: `Recipients.ResolveAll()` called,
  `.Save()` called, `.Send()` never called.

---

## Files touched

| File | Change |
|---|---|
| `src/outlook_desktop_mcp/com_bridge.py` | Lazy connect, disconnect-HRESULT detection, reconnect + single retry |
| `src/outlook_desktop_mcp/server.py` | `DRAFT_MODE` module flag from env; `confirm_send` param + branching on `send_email`, `reply_email`, `create_meeting`; docstring updates |
| `tests/test_com_bridge.py` | New reconnect/lazy-connect test cases |
| `tests/test_send_gating.py` | New file, draft/send branching tests |
| `README.md` | New "Draft-safe mode" section documenting `OUTLOOK_MCP_DRAFT_MODE` and `confirm_send` |

---

## Out of Scope

- macOS (`server_mac.py`) draft-mode parity — same send-immediately issue
  exists there but no reconnect bug (stateless AppleScript per call). Follow-up.
- `respond_to_meeting`, `delete_event` cancel-path gating — see rationale above.
- Attachments on drafted emails/replies (not requested; can follow the
  `outlook-email` skill's `--attach` pattern later if needed).
- Any change to the existing `outlook` CLI plan in
  `docs/superpowers/plans/2026-05-22-outlook-cli.md` (Phase 2, not yet
  started) — this work lands independently and first, since both touch
  `server.py`/`com_bridge.py`.
