"""
Outlook COM constants
=====================
Two distinct constant sets - do NOT mix them up:

olItemType - used with Outlook.CreateItem(n). OL_MAIL_ITEM=0, OL_APPOINTMENT_ITEM=1, OL_TASK_ITEM=3.
olObjectClass - used to verify item.Class. OL_CLASS_MAIL=43, OL_CLASS_APPOINTMENT=26, OL_CLASS_TASK=48.

CreateItem(0) creates a mail item (olItemType=0).
item.Class == 43 checks that an existing item IS a mail item (olObjectClass=43).
"""

# --- olDefaultFolders ---
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_TASKS = 13
OL_FOLDER_JUNK = 23
OL_FOLDER_OUTBOX = 4
OL_FOLDER_ARCHIVE = 28  # olFolderArchive - was missing, causing 'archive' default to DFS-search

FOLDER_NAME_TO_ENUM = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "sentmail": OL_FOLDER_SENT_MAIL,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "trash": OL_FOLDER_DELETED_ITEMS,
    "drafts": OL_FOLDER_DRAFTS,
    "calendar": OL_FOLDER_CALENDAR,
    "tasks": OL_FOLDER_TASKS,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
    "archive": OL_FOLDER_ARCHIVE,
}

# --- olItemType - used with Outlook.CreateItem() ---
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1
OL_TASK_ITEM = 3

# --- olObjectClass - used to verify item.Class ---
OL_CLASS_MAIL = 43
OL_CLASS_APPOINTMENT = 26
OL_CLASS_TASK = 48

# --- Formatting ---
OUTLOOK_DATE_FMT = "%m/%d/%Y %H:%M"
MAX_ITEM_COUNT = 200

# --- COM dispatch IDs ---
SEND_USING_ACCOUNT_DISPID = 64209  # Outlook SendUsingAccount property DISPID

# --- olBusyStatus ---
OL_BUSY_FREE = 0
OL_BUSY_TENTATIVE = 1
OL_BUSY_BUSY = 2
OL_BUSY_OUT_OF_OFFICE = 3
OL_BUSY_WORKING_ELSEWHERE = 4

# --- olMeetingStatus ---
OL_NON_MEETING = 0
OL_MEETING = 1
OL_MEETING_RECEIVED = 3
OL_MEETING_CANCELED = 5

# --- olMeetingResponse ---
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_DECLINED = 4

# --- olRecipientType ---
OL_REQUIRED = 1
OL_OPTIONAL = 2
OL_RESOURCE = 3

# --- Display name maps ---
BUSY_STATUS_NAMES = {
    0: "free", 1: "tentative", 2: "busy",
    3: "out_of_office", 4: "working_elsewhere",
}
MEETING_STATUS_NAMES = {
    0: "appointment", 1: "meeting", 3: "received", 5: "canceled",
}
RESPONSE_NAMES = {
    0: "none", 1: "organized", 2: "tentative",
    3: "accepted", 4: "declined", 5: "not_responded",
}

# --- olTaskStatus ---
OL_TASK_NOT_STARTED = 0
OL_TASK_IN_PROGRESS = 1
OL_TASK_COMPLETE = 2
OL_TASK_WAITING = 3
OL_TASK_DEFERRED = 4

TASK_STATUS_NAMES = {
    0: "not_started", 1: "in_progress", 2: "complete",
    3: "waiting", 4: "deferred",
}

# --- olImportance ---
OL_IMPORTANCE_LOW = 0
OL_IMPORTANCE_NORMAL = 1
OL_IMPORTANCE_HIGH = 2

IMPORTANCE_NAMES = {0: "low", 1: "normal", 2: "high"}

# --- olRuleActionType ---
OL_RULE_ACTION_MOVE = 1
OL_RULE_ACTION_DELETE = 8
OL_RULE_ACTION_MARK_READ = 11
