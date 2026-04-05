"""
Email engine for PE Deal Sourcing Automator.
Provides OutlookClient with real Outlook (Windows/win32com) or mock (Linux) backend.
"""

import logging
import platform
import re
from datetime import datetime, timedelta
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

# Reply/forward subject prefixes stripped during subject normalisation (multilingual)
_REPLY_PREFIXES = re.compile(r"^\s*(re|aw|fwd?|sv|vs|wg)\s*:\s*", re.IGNORECASE)


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

    # ------------------------------------------------------------------
    # Draft creation
    # ------------------------------------------------------------------

    def _save_draft_oom(
        self,
        to: str,
        subject: str,
        html_body: str,
        *,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        deferred_delivery: Optional[datetime] = None,
    ) -> bool:
        """Save an Outlook draft via OOM only (no MAPI, no OMG).

        Pure OOM approach: Create MailItem, set Subject/HTMLBody, set To/CC/BCC
        as display strings, set DeferredDeliveryTime if needed, then Save.

        Setting mail.To/CC/BCC to strings does NOT trigger OMG—only actual
        Sending or reading addresses from existing items does. This is safe
        for new drafts.
        """
        if not self._ensure_outlook():
            return False

        try:
            mail = self._outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = subject
            mail.HTMLBody = html_body

            # Set recipients as display strings (no address-book resolution, no OMG)
            if to:
                mail.To = to
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            if deferred_delivery is not None:
                mail.DeferredDeliveryTime = pywintypes.Time(deferred_delivery.timestamp())

            mail.Save()
            mail.Close(1)  # olDiscard (already saved)
            return True

        except Exception as e:
            logger.exception("Failed to save draft: %s", e)
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
        # Route through signature path so all drafts include signature
        return self._create_draft_with_sig_real(
            to, subject,
            html_body if html_body is not None else body,
            cc=cc, bcc=bcc,
        )

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

    # ------------------------------------------------------------------
    # Signature extraction
    # ------------------------------------------------------------------

    def get_default_signature_html(self) -> str:
        """
        Extract the default Outlook signature HTML by creating a throwaway mail item.
        Returns empty string on mock/error.
        """
        if _USE_REAL_OUTLOOK:
            return self._get_signature_real()
        return ""

    def _get_signature_real(self) -> str:
        """Read default Outlook signature HTML from disk. No COM, no OMG.

        Avoids GetInspector which triggers internal CurrentUser.Address resolution.
        Instead reads .htm file from %%APPDATA%%\\Microsoft\\Signatures\\
        """
        try:
            import os
            import winreg

            sig_dir = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures")
            if not os.path.isdir(sig_dir):
                return ""

            # Try registry first to find configured default signature name
            sig_name = None
            for ver in ["16.0", "15.0", "14.0", "12.0"]:
                try:
                    reg_path = f"Software\\Microsoft\\Office\\{ver}\\Common\\MailSettings"
                    with winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER,
                        reg_path,
                    ) as key:
                        val, _ = winreg.QueryValueEx(key, "NewSignature")
                        if val:
                            sig_name = val
                            break
                except OSError:
                    continue

            if sig_name:
                sig_file = os.path.join(sig_dir, f"{sig_name}.htm")
                if os.path.exists(sig_file):
                    with open(sig_file, "r", encoding="utf-8", errors="ignore") as f:
                        return f.read()

            # Fallback: use the only .htm file if exactly one exists
            htm_files = [f for f in os.listdir(sig_dir) if f.lower().endswith(".htm")]
            if len(htm_files) == 1:
                with open(os.path.join(sig_dir, htm_files[0]), "r",
                          encoding="utf-8", errors="ignore") as f:
                    return f.read()

            return ""
        except Exception as e:
            logger.warning("Could not read signature from filesystem: %s", e)
            return ""

    # ------------------------------------------------------------------
    # Draft with signature + optional deferred delivery
    # ------------------------------------------------------------------

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
            # Get signature: use cached version or extract fresh
            sig = signature_html if signature_html else self._get_signature_real()

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

            return self._save_draft_oom(
                to, subject, final_html,
                cc=cc, bcc=bcc,
                deferred_delivery=deferred_delivery,
            )
        except Exception as e:
            logger.exception("Failed to create draft with signature: %s", e)
            return False

    # ------------------------------------------------------------------
    # Threaded reply draft (for follow-up campaigns)
    # ------------------------------------------------------------------

    def create_reply_draft(
        self,
        original_subject: str,
        to_email: str,
        html_body: str,
        *,
        signature_html: str = "",
        sent_lookback_days: int = 180,
    ) -> bool:
        """
        Find the original sent email for this contact and create a threaded reply draft.

        Locates the sent item by subject + recipient in the Sent Items folder, calls
        .Reply() on it so Outlook sets In-Reply-To and References automatically, then
        replaces the body with the follow-up template content and saves to Drafts.

        Args:
            original_subject: The exact subject that was sent to this contact (merged).
            to_email: Recipient email address (used to identify the correct sent item).
            html_body: Follow-up template HTML (already merged with contact data).
            signature_html: Cached signature HTML to inject, same as other draft methods.
            sent_lookback_days: How far back to search Sent Items (default 180 days).

        Returns:
            True if reply draft was saved; False if sent item not found or on error.
        """
        if _USE_REAL_OUTLOOK:
            return self._create_reply_draft_real(
                original_subject, to_email, html_body,
                signature_html=signature_html,
                sent_lookback_days=sent_lookback_days,
            )
        return self._create_reply_draft_mock(original_subject, to_email, html_body)

    def _create_reply_draft_real(
        self, original_subject, to_email, html_body,
        *, signature_html="", sent_lookback_days=180,
    ) -> bool:
        if not self._ensure_outlook():
            return False
        try:
            sent_folder = self._namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail

            # Chain two Restrict() calls — more reliable than AND in a single Jet filter.
            # Use 12-hour clock (%I + %p) to avoid the %H/%p mixing bug.
            cutoff = datetime.now() - timedelta(days=sent_lookback_days)
            cutoff_str = cutoff.strftime("%m/%d/%Y %I:%M %p")
            items = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_str}'")
            # Escape single quotes in subject for Jet filter
            safe_subject = original_subject.replace("'", "''")
            items = items.Restrict(f"[Subject] = '{safe_subject}'")
            items.Sort("[SentOn]", True)  # most recent first

            # Find the sent item addressed to this recipient.
            # Use index-based Recipients access (more reliable than iteration in win32com).
            # recip.Address may be an Exchange DN for GAL contacts, so also check the
            # To string by extracting email addresses from angle-bracket notation.
            sent_item = None
            to_lower = to_email.strip().lower()
            for mail in items:
                try:
                    if mail.Class != 43:  # skip non-mail items
                        continue
                    matched = False
                    # Check Recipients collection via index (1-based)
                    try:
                        for j in range(1, mail.Recipients.Count + 1):
                            recip = mail.Recipients.Item(j)
                            addr = (recip.Address or "").strip().lower()
                            if addr == to_lower:
                                matched = True
                                break
                    except Exception:
                        pass
                    # Fallback: extract bare addresses from the To string
                    # e.g. "John Doe <john@example.com>; Jane <jane@example.com>"
                    if not matched:
                        to_str = (mail.To or "").lower()
                        addrs_in_to = re.findall(r"<([^>]+)>", to_str)
                        if not addrs_in_to:
                            # No angle brackets — treat each semicolon-separated token as an address
                            addrs_in_to = [a.strip() for a in to_str.split(";")]
                        if to_lower in addrs_in_to:
                            matched = True
                    if matched:
                        logger.debug("Found sent item for %r: %r", to_email, mail.Subject)
                        sent_item = mail
                        break
                    else:
                        logger.debug(
                            "Skipping sent item (recipient mismatch): To=%r looking for %r",
                            mail.To, to_email,
                        )
                except Exception as item_err:
                    logger.debug("Skipping sent item: %s", item_err)
                    continue

            if sent_item is None:
                # Log all candidates that were found (subject matched) to diagnose mismatch
                try:
                    debug_items = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_str}'")
                    debug_items = debug_items.Restrict(f"[Subject] = '{safe_subject}'")
                    candidates = []
                    for m in debug_items:
                        try:
                            if m.Class != 43:
                                continue
                            recip_addrs = []
                            try:
                                for j in range(1, m.Recipients.Count + 1):
                                    recip_addrs.append(m.Recipients.Item(j).Address)
                            except Exception:
                                pass
                            candidates.append(f"To={m.To!r} Recipients={recip_addrs}")
                        except Exception:
                            pass
                    logger.warning(
                        "No sent item found for reply: subject=%r to=%r — "
                        "candidates with this subject: %s",
                        original_subject, to_email,
                        candidates if candidates else "NONE (subject filter matched nothing)",
                    )
                except Exception:
                    logger.warning(
                        "No sent item found for reply: subject=%r to=%r",
                        original_subject, to_email,
                    )
                return False

            # .Reply() creates a MailItem with In-Reply-To + References set correctly.
            # On a sent item, Outlook populates Recipients with yourself (the sender);
            # remove those and add the correct contact.
            reply = sent_item.Reply()
            for i in range(reply.Recipients.Count, 0, -1):
                reply.Recipients.Remove(i)
            reply.Recipients.Add(to_email)
            if not reply.Recipients.ResolveAll():
                logger.warning("Could not resolve recipient %r for reply draft", to_email)

            # Extract just the inner body content from signature HTML (full HTML doc).
            sig = signature_html
            sig_inner = ""
            if sig:
                sig_body_start = re.search(r"<body[^>]*>", sig, re.IGNORECASE)
                sig_body_end = re.search(r"</body>", sig, re.IGNORECASE)
                if sig_body_start:
                    end = sig_body_end.start() if sig_body_end else len(sig)
                    sig_inner = sig[sig_body_start.end():end]
                else:
                    sig_inner = sig

            # Build quoted block directly from the original sent item.
            # reply.HTMLBody from .Reply() does not reliably contain the quoted original
            # via COM automation (it is only rendered when the inspector opens), so we
            # construct it ourselves from the sent item's properties.
            try:
                sent_date_str = sent_item.SentOn.strftime("%A, %B %d, %Y %I:%M %p")
            except Exception:
                sent_date_str = ""
            sent_from = sent_item.SenderName or sent_item.SenderEmailAddress or ""
            sent_subj = sent_item.Subject or ""
            orig_html = sent_item.HTMLBody or ""
            orig_inner_match = re.search(
                r"<body[^>]*>(.*?)</body>", orig_html, re.IGNORECASE | re.DOTALL
            )
            orig_inner = orig_inner_match.group(1) if orig_inner_match else orig_html

            quoted_block = (
                f'<hr style="display:inline-block;width:98%" tabindex="-1">'
                f'<p style="margin:0"><b>From:</b> {sent_from}<br>'
                f'<b>Sent:</b> {sent_date_str}<br>'
                f'<b>To:</b> {sent_item.To or ""}<br>'
                f'<b>Subject:</b> {sent_subj}</p>'
                f'<div>{orig_inner}</div>'
            )

            # Assemble final HTML: new content + signature + quoted original.
            # Use reply.HTMLBody as the outer HTML shell (preserves Outlook's head/styles).
            existing_body = reply.HTMLBody
            body_tag = re.search(r"(<body[^>]*>)", existing_body, re.IGNORECASE)
            if body_tag:
                insert_pos = body_tag.end()
                final_html = (existing_body[:insert_pos]
                              + html_body + sig_inner
                              + quoted_block
                              + existing_body[insert_pos:])
            else:
                final_html = html_body + sig_inner + quoted_block

            reply.HTMLBody = final_html
            reply.Save()
            reply.Close(1)  # olDiscard — already saved to Drafts
            return True

        except Exception as e:
            logger.exception("Failed to create reply draft: %s", e)
            return False

    def _create_reply_draft_mock(
        self, original_subject, to_email, html_body,
    ) -> bool:
        logger.info(
            "[MOCK] create_reply_draft to=%r original_subject=%r html_len=%d",
            to_email, original_subject, len(html_body),
        )
        return True

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

    # ------------------------------------------------------------------
    # Sender address resolution
    # ------------------------------------------------------------------

    def _get_smtp_address(self, mail_item) -> str:
        """Return the best SMTP sender address for a received mail item.

        mail.SenderEmailAddress is unreliable for internet senders (Gmail,
        web.de, etc.) in Outlook — it can return an Exchange DN instead of
        the plain SMTP address. Try multiple properties in order.
        """
        # 1. SenderEmailAddress — reliable for plain SMTP/IMAP accounts
        addr = (mail_item.SenderEmailAddress or "").strip()
        if addr and "@" in addr:
            return addr.lower()
        # 2. AddressEntry.Address — works when Outlook has resolved the sender
        try:
            entry = mail_item.Sender
            if entry:
                addr2 = (entry.Address or "").strip()
                if addr2 and "@" in addr2:
                    return addr2.lower()
                # Exchange user — fetch primary SMTP via GetExchangeUser()
                try:
                    ex_user = entry.GetExchangeUser()
                    if ex_user:
                        smtp = (ex_user.PrimarySmtpAddress or "").strip()
                        if smtp:
                            return smtp.lower()
                except Exception:
                    pass
        except Exception:
            pass
        # 3. Extract from SenderName "Display Name <email@domain>"
        name = mail_item.SenderName or ""
        m = re.search(r"<([^>]+)>", name)
        if m:
            return m.group(1).strip().lower()
        return ""

    # ------------------------------------------------------------------
    # Subject normalisation
    # ------------------------------------------------------------------

    def _normalize_subject(self, subject: str) -> str:
        """Strip all leading reply/forward prefixes (Re:, AW:, FW:, SV:, WG:, VS:).

        Loops until no more prefixes remain so chained replies like
        "Re: Re: Intro" are fully unwrapped.
        """
        s = subject or ""
        while True:
            stripped = _REPLY_PREFIXES.sub("", s)
            if stripped == s:
                return s.strip()
            s = stripped

    # ------------------------------------------------------------------
    # Reply scanning
    # ------------------------------------------------------------------

    def scan_for_replies(
        self,
        email_list: list[str],
        *,
        folder_name: str = "Inbox",
        max_items: int = 500,
        subject_filters: list[str] | None = None,
        lookback_days: int = 14,
    ) -> list[dict]:
        """
        Scan the local Inbox (or named folder) for replies from the given addresses.

        Args:
            email_list: List of sender email addresses to look for (case-insensitive).
            folder_name: Outlook folder to scan (default "Inbox").
            max_items: Maximum number of items to scan within the lookback window.
            subject_filters: Optional list of original (bare) subject strings. When
                provided, an inbox item must match both a sender address AND contain
                one of the normalised subjects as a substring.
            lookback_days: Only scan items received within this many days (default 14).
                Uses Items.Restrict() so filtering is done by the Outlook data store,
                not in Python — much faster for large inboxes.

        Returns:
            List of dicts with keys: sender_email, subject, received_time, entry_id.
            Empty list on error or when using mock backend.
        """
        if not email_list:
            return []
        normalized = {addr.strip().lower() for addr in email_list if addr}
        if _USE_REAL_OUTLOOK:
            return self._scan_for_replies_real(
                normalized, folder_name, max_items,
                subject_filters=subject_filters, lookback_days=lookback_days,
            )
        return self._scan_for_replies_mock(
            normalized, folder_name, max_items,
            subject_filters=subject_filters, lookback_days=lookback_days,
        )

    def _scan_for_replies_real(
        self,
        email_set: set[str],
        folder_name: str,
        max_items: int,
        *,
        subject_filters: list[str] | None = None,
        lookback_days: int = 14,
    ) -> list[dict]:
        if not self._ensure_outlook():
            return []
        try:
            inbox = self._namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

            # Restrict() delegates date filtering to the Outlook/MAPI data store so
            # only items within the lookback window are iterated — far faster than
            # loading all items and comparing dates in Python.
            # Use 12-hour clock (%I) + AM/PM (%p); mixing 24-hour %H with %p is invalid.
            cutoff = datetime.now() - timedelta(days=lookback_days)
            cutoff_str = cutoff.strftime("%m/%d/%Y %I:%M %p")
            items = inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'")
            items.Sort("[ReceivedTime]", True)  # newest first

            norm_filters = (
                [sf.lower() for sf in subject_filters if sf]
                if subject_filters else None
            )

            results = []
            count = 0
            for mail in items:
                if count >= max_items:
                    break
                count += 1
                try:
                    if mail.Class != 43:  # 43 = olMail; skip calendar/task items
                        continue
                    sender = self._get_smtp_address(mail)
                    if not sender or sender not in email_set:
                        continue
                    if norm_filters is not None:
                        norm_subj = self._normalize_subject(mail.Subject or "").lower()
                        if not any(f in norm_subj for f in norm_filters):
                            continue
                    results.append({
                        "sender_email": sender,
                        "subject": mail.Subject or "",
                        "received_time": str(mail.ReceivedTime),
                        "entry_id": mail.EntryID,
                    })
                except Exception as item_err:
                    logger.debug("Skipping inbox item: %s", item_err)
            return results
        except Exception as e:
            logger.exception("scan_for_replies error: %s", e)
            return []

    def _scan_for_replies_mock(
        self,
        email_set: set[str],
        folder_name: str,
        max_items: int,
        *,
        subject_filters: list[str] | None = None,
        lookback_days: int = 14,
    ) -> list[dict]:
        logger.info(
            "[MOCK] scan_for_replies email_list=%s folder=%s max_items=%s "
            "subject_filters=%s lookback_days=%s",
            list(email_set),
            folder_name,
            max_items,
            subject_filters,
            lookback_days,
        )
        return []


# Convenience instance; callers can also instantiate OutlookClient().
def get_outlook_client() -> OutlookClient:
    """Return a shared OutlookClient instance."""
    return OutlookClient()
