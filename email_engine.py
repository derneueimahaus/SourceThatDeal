"""
Email engine for PE Deal Sourcing Automator.
Provides OutlookClient with real Outlook (Windows/win32com) or mock (Linux) backend.
"""

import logging
import platform
import re
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)

# --- OS detection and backend selection ---
_IS_WINDOWS = platform.system() == "Windows"
_win32com = None

if _IS_WINDOWS:
    try:
        import win32com.client
        import pywintypes

        _win32com = win32com.client
    except ImportError as e:
        logger.warning(
            "Windows detected but win32com unavailable (%s). Using mock backend.",
            e,
        )
        _IS_WINDOWS = False

# Use real Outlook only when we have win32com on Windows
_USE_REAL_OUTLOOK = _IS_WINDOWS and _win32com is not None


class OutlookClient:
    """
    Client for creating drafts and scanning Inbox via local Outlook (Windows)
    or a mock implementation (Linux / Windows without win32com).
    """

    def __init__(self) -> None:
        self._outlook = None
        self._namespace = None
        if _USE_REAL_OUTLOOK:
            self._ensure_outlook()

    def _ensure_outlook(self) -> bool:
        """Get or create Outlook Application instance. Returns False if unavailable."""
        if not _USE_REAL_OUTLOOK:
            return False
        try:
            if self._outlook is None:
                try:
                    self._outlook = _win32com.GetActiveObject("Outlook.Application")
                except pywintypes.com_error:
                    logger.info(
                        "Outlook.Application not active, attempting to dispatch new instance."
                    )
                    self._outlook = _win32com.Dispatch("Outlook.Application")
                except Exception as e:
                    logger.warning(
                        "Failed to get active Outlook object: %s. Attempting to dispatch.",
                        e,
                    )
                    self._outlook = _win32com.Dispatch("Outlook.Application")
                self._namespace = self._outlook.GetNamespace("MAPI")
            return True
        except pywintypes.com_error as e:
            logger.exception(
                "Outlook COM error: %s. Please ensure Outlook is running and responsive.", e
            )
            self._outlook = None
            self._namespace = None
            return False
        except Exception as e:
            logger.exception("Failed to connect to Outlook: %s", e)
            self._outlook = None
            self._namespace = None
            return False

    def create_draft(
        self,
        to: str,
        subject: str,
        body: str,
        *,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        html_body: Optional[str] = None,
    ) -> bool:
        """
        Create an email draft in the Outlook Drafts folder.

        Args:
            to: Recipient email address(es), comma-separated if multiple.
            subject: Email subject.
            body: Plain-text body (used when html_body is not provided).
            cc: Optional CC addresses (comma-separated).
            bcc: Optional BCC addresses (comma-separated).
            html_body: Optional HTML body; when set, draft uses rich HTML (and body as fallback).

        Returns:
            True if draft was created (or mocked) successfully, False otherwise.
        """
        if _USE_REAL_OUTLOOK:
            return self._create_draft_real(
                to, subject, body, cc=cc, bcc=bcc, html_body=html_body
            )
        return self._create_draft_mock(
            to, subject, body, cc=cc, bcc=bcc, html_body=html_body
        )

    def _create_draft_real(
        self,
        to: str,
        subject: str,
        body: str,
        *,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        html_body: Optional[str] = None,
    ) -> bool:
        if not self._ensure_outlook():
            return False
        try:
            mail = self._outlook.CreateItem(0)  # 0 = olMail
            mail.To = to
            mail.Subject = subject
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            # Access the inspector to ensure the mail item is fully initialized.
            # This is a robust way to prepare the item before modifying its body.
            _ = mail.GetInspector

            if html_body is not None:
                # Overwrite the entire body (including any default signature) with our HTML.
                mail.HTMLBody = html_body
            else:
                mail.Body = body

            mail.Save()
            return True
        except Exception as e:
            logger.exception("Failed to create draft: %s", e)
            return False

    def _create_draft_mock(
        self,
        to: str,
        subject: str,
        body: str,
        *,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        html_body: Optional[str] = None,
    ) -> bool:
        logger.info(
            "[MOCK] create_draft to=%r subject=%r body_len=%d html=%s cc=%r bcc=%r",
            to,
            subject,
            len(body),
            bool(html_body),
            cc,
            bcc,
        )
        return True

    def get_default_signature_html(self) -> str:
        """
        Extract the default Outlook signature HTML by creating a throwaway mail item.
        Returns empty string on mock/error.
        """
        if _USE_REAL_OUTLOOK:
            return self._get_signature_real()
        return ""

    def _get_signature_real(self) -> str:
        if not self._ensure_outlook():
            return ""
        try:
            mail = self._outlook.CreateItem(0)
            _ = mail.GetInspector  # triggers Outlook to insert default signature
            sig = mail.HTMLBody or ""
            mail.Close(1)  # 1 = olDiscard
            return sig
        except Exception as e:
            logger.exception("Could not extract signature: %s", e)
            return ""

    def create_draft_with_signature(
        self,
        to: str,
        subject: str,
        html_body: str,
        *,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        deferred_delivery: Optional[datetime] = None,
        signature_html: str = "",
    ) -> bool:
        """
        Create a draft that preserves the Outlook default signature
        and optionally sets DeferredDeliveryTime.

        Args:
            to: Recipient email address(es).
            subject: Email subject.
            html_body: HTML body (template content, already merged).
            cc: Optional CC addresses.
            bcc: Optional BCC addresses.
            deferred_delivery: Optional datetime for deferred send.
            signature_html: Pre-cached signature HTML (from get_default_signature_html).
                If empty, signature is extracted per-draft (slower for bulk).
        """
        if _USE_REAL_OUTLOOK:
            return self._create_draft_with_sig_real(
                to, subject, html_body,
                cc=cc, bcc=bcc,
                deferred_delivery=deferred_delivery,
                signature_html=signature_html,
            )
        return self._create_draft_with_sig_mock(
            to, subject, html_body,
            cc=cc, bcc=bcc,
            deferred_delivery=deferred_delivery,
        )

    def _create_draft_with_sig_real(
        self, to, subject, html_body, *, cc=None, bcc=None,
        deferred_delivery=None, signature_html="",
    ) -> bool:
        if not self._ensure_outlook():
            return False
        try:
            mail = self._outlook.CreateItem(0)

            # Get signature: use cached version or extract from this mail item
            if signature_html:
                sig = signature_html
            else:
                _ = mail.GetInspector
                sig = mail.HTMLBody or ""

            # Merge template HTML with signature
            # Insert our content right after <body...> tag in the signature HTML
            if sig:
                body_match = re.search(r"(<body[^>]*>)", sig, re.IGNORECASE)
                if body_match:
                    insert_pos = body_match.end()
                    final_html = sig[:insert_pos] + html_body + sig[insert_pos:]
                else:
                    final_html = html_body + sig
            else:
                final_html = html_body

            mail.HTMLBody = final_html
            mail.To = to
            mail.Subject = subject
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            if deferred_delivery:
                mail.DeferredDeliveryTime = deferred_delivery

            mail.Save()
            return True
        except Exception as e:
            logger.exception("Failed to create draft with signature: %s", e)
            return False

    def _create_draft_with_sig_mock(
        self, to, subject, html_body, *, cc=None, bcc=None,
        deferred_delivery=None,
    ) -> bool:
        logger.info(
            "[MOCK] create_draft_with_signature to=%r subject=%r html_len=%d "
            "cc=%r bcc=%r deferred=%s",
            to, subject, len(html_body), cc, bcc, deferred_delivery,
        )
        return True

    def scan_for_replies(
        self,
        email_list: list[str],
        *,
        folder_name: str = "Inbox",
        max_items: int = 500,
    ) -> list[dict]:
        """
        Scan the local Inbox (or named folder) for replies from the given addresses.

        Args:
            email_list: List of sender email addresses to look for (case-insensitive).
            folder_name: Outlook folder to scan (default "Inbox").
            max_items: Maximum number of items to scan (default 500).

        Returns:
            List of dicts with keys: sender_email, subject, received_time, entry_id (or similar).
            Empty list on error or when using mock and no mock data is configured.
        """
        if not email_list:
            return []
        normalized = {addr.strip().lower() for addr in email_list if addr}
        if _USE_REAL_OUTLOOK:
            return self._scan_for_replies_real(normalized, folder_name, max_items)
        return self._scan_for_replies_mock(normalized, folder_name, max_items)

    def _scan_for_replies_real(
        self,
        email_set: set[str],
        folder_name: str,
        max_items: int,
    ) -> list[dict]:
        if not self._ensure_outlook():
            return []
        try:
            inbox = self._namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if folder_name != "Inbox":
                folders = inbox.Parent.Folders
                folder = None
                for i in range(1, folders.Count + 1):
                    if folders.Item(i).Name == folder_name:
                        folder = folders.Item(i)
                        break
                if folder is None:
                    folder = inbox
            else:
                folder = inbox
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # descending
            results = []
            count = 0
            for item in items:
                if count >= max_items:
                    break
                try:
                    sender = getattr(item, "SenderEmailAddress", None) or getattr(
                        item, "Sender", None
                    )
                    if sender is None:
                        continue
                    sender_str = str(sender).strip().lower()
                    if sender_str not in email_set:
                        continue
                    received = getattr(item, "ReceivedTime", None)
                    results.append({
                        "sender_email": sender_str,
                        "subject": getattr(item, "Subject", "") or "",
                        "received_time": received,
                        "entry_id": getattr(item, "EntryID", "") or "",
                    })
                    count += 1
                except Exception as e:
                    logger.debug("Skipping item during scan: %s", e)
                    continue
            return results
        except Exception as e:
            logger.exception("Failed to scan for replies: %s", e)
            return []

    def _scan_for_replies_mock(
        self,
        email_set: set[str],
        folder_name: str,
        max_items: int,
    ) -> list[dict]:
        logger.info(
            "[MOCK] scan_for_replies email_list=%s folder=%s max_items=%s",
            list(email_set),
            folder_name,
            max_items,
        )
        return []


# Convenience instance; callers can also instantiate OutlookClient().
def get_outlook_client() -> OutlookClient:
    """Return a shared OutlookClient instance."""
    return OutlookClient()
