"""Verify constants are importable from utils.constants and values are correct."""
from outlook_desktop_mcp.utils.constants import (
    FOLDER_NAME_TO_ENUM,
    OL_FOLDER_ARCHIVE,
    OL_CLASS_MAIL,
    OL_CLASS_APPOINTMENT,
    OL_CLASS_TASK,
    OUTLOOK_DATE_FMT,
    MAX_ITEM_COUNT,
    SEND_USING_ACCOUNT_DISPID,
)


def test_archive_in_folder_name_to_enum():
    assert "archive" in FOLDER_NAME_TO_ENUM
    assert FOLDER_NAME_TO_ENUM["archive"] == 28


def test_ol_folder_archive_value():
    assert OL_FOLDER_ARCHIVE == 28


def test_ol_class_constants_distinct_from_item_type():
    # olObjectClass values (for item.Class checks)
    assert OL_CLASS_MAIL == 43
    assert OL_CLASS_APPOINTMENT == 26
    assert OL_CLASS_TASK == 48


def test_formatting_constants():
    assert OUTLOOK_DATE_FMT == "%m/%d/%Y %H:%M"
    assert MAX_ITEM_COUNT == 200


def test_send_using_account_dispid():
    assert SEND_USING_ACCOUNT_DISPID == 64209
