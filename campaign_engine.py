"""
Campaign engine for PE Deal Sourcing Automator.
Template field extraction, merge, and draft creation logic.
No UI code — testable independently.
"""

import re
from datetime import datetime
from typing import Optional

from email_engine import OutlookClient
from file_manager import read_contact_list, read_template


def extract_template_fields(html: str) -> list[str]:
    """
    Parse {{field_name}} placeholders from template HTML.
    Returns sorted list of unique field names.
    """
    return sorted(set(re.findall(r"\{\{(.+?)\}\}", html)))


def merge_template(html: str, row: dict, field_mapping: dict) -> str:
    """
    Replace all {{placeholder}} in html using row data + field_mapping.
    field_mapping maps template_field_name -> contact_list_column_name.
    Missing values become empty string.
    """
    def replacer(match):
        field = match.group(1)
        col = field_mapping.get(field, field)
        return row.get(col, "")
    return re.sub(r"\{\{(.+?)\}\}", replacer, html)


def merge_subject(subject_template: str, row: dict, field_mapping: dict) -> str:
    """Same as merge_template but for the subject line (plain text)."""
    def replacer(match):
        field = match.group(1)
        col = field_mapping.get(field, field)
        return row.get(col, "")
    return re.sub(r"\{\{(.+?)\}\}", replacer, subject_template)


def guess_email_column(columns: list[str]) -> str | None:
    """Try to auto-detect which column contains email addresses."""
    for col in columns:
        if "email" in col.lower() or "e-mail" in col.lower():
            return col
    return columns[0] if columns else None


def guess_column_match(field_name: str, columns: list[str]) -> str | None:
    """Try to match a template field to a column by name similarity."""
    field_lower = field_name.lower()
    for col in columns:
        if col.lower() == field_lower:
            return col
    for col in columns:
        if field_lower in col.lower() or col.lower() in field_lower:
            return col
    return None


def create_campaign_drafts(
    outlook: OutlookClient,
    template_html: str,
    subject_template: str,
    rows: list[dict],
    field_mapping: dict,
    email_column: str,
    *,
    deferred_delivery: Optional[datetime] = None,
    signature_html: str = "",
    on_progress: Optional[callable] = None,
) -> tuple[int, list[str]]:
    """
    Create Outlook drafts for all rows.

    Args:
        outlook: OutlookClient instance.
        template_html: Raw template HTML with {{placeholders}}.
        subject_template: Subject line with optional {{placeholders}}.
        rows: List of row dicts from contact list.
        field_mapping: Template field -> column name mapping.
        email_column: Column name containing email addresses.
        deferred_delivery: Optional datetime for DeferredDeliveryTime.
        signature_html: Cached Outlook signature HTML to append.
        on_progress: Optional callback(current: int, total: int).

    Returns:
        (success_count, error_messages) tuple.
    """
    success = 0
    errors = []
    total = len(rows)

    for i, row in enumerate(rows):
        email = row.get(email_column, "").strip()
        if not email:
            errors.append(f"Row {i + 1}: no email address")
            if on_progress:
                on_progress(i + 1, total)
            continue

        merged_html = merge_template(template_html, row, field_mapping)
        merged_subject = merge_subject(subject_template, row, field_mapping)

        ok = outlook.create_draft_with_signature(
            to=email,
            subject=merged_subject,
            html_body=merged_html,
            deferred_delivery=deferred_delivery,
            signature_html=signature_html,
        )

        if ok:
            success += 1
        else:
            errors.append(f"Row {i + 1} ({email}): Outlook error")

        if on_progress:
            on_progress(i + 1, total)

    return success, errors
