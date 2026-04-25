"""
emailer.email_sender
====================

Network layer — takes a pre-built HTML report and ships it via SMTP.

Design notes
------------
Kept deliberately thin: this module does not build the HTML, does not
aggregate the data, and does not touch the GUI. Its only job is to
turn ``(subject, html, config)`` into a delivered message or a clear
error string the GUI can show.

Error surface
-------------
The sender catches three specific exception classes and turns them
into short, actionable error messages:

* ``SMTPAuthenticationError`` — wrong user/app-password. The message
  tells the user to check ``EMAIL_PASSWORD`` (a Gmail App Password,
  not their account password).
* ``SMTPException`` — any other SMTP protocol error (connection,
  recipient refused, etc).
* ``OSError`` — network unreachable, DNS failure, firewall block.

A final ``Exception`` catch-all is NOT included — if something
unexpected breaks, we want the crash to be visible during
development so the specific cause can be added to the list above.

Return contract
---------------
``send()`` returns ``(True, '')`` on success or ``(False, reason)``
on failure. Never raises.
"""

from __future__ import annotations
import logging
import smtplib
from email.message import EmailMessage
from typing import Any, Dict, Tuple

from online_po_processor.data.models import ProcessingResult
from online_po_processor.emailer.email_builder import EmailBuilder


class EmailSender:
    """
    Send a pre-built HTML email via SMTP.

    Usage::

        from online_po_processor.config.email_config import get_email_config
        ok, err = EmailSender.send(result, get_email_config())
        if not ok:
            show_error(err)
    """

    # ── Public API ─────────────────────────────────────────────────────

    @staticmethod
    def send(
        result: ProcessingResult,
        config: Dict[str, Any],
    ) -> Tuple[bool, str]:
        """
        Build the HTML via :class:`EmailBuilder` and send it.

        Args:
            result: Populated ``ProcessingResult`` to describe in the
                    email. Must have at least one row — we guard
                    against empty results because an empty report is
                    just noise in the recipient's inbox.
            config: Effective email config dict (see
                    :func:`~online_po_processor.config.email_config.
                    get_email_config`).

        Returns:
            ``(success, error_reason)``. On success, ``error_reason``
            is an empty string. On failure, ``success`` is False and
            ``error_reason`` is a short, user-presentable message.
        """
        # ── Guard 1: nothing to report ─────────────────────────────────
        if not result.rows:
            return False, (
                "Cannot send email — no rows in the latest run. "
                "Generate a successful SO first."
            )

        # ── Guard 2: config sanity ─────────────────────────────────────
        # We only check the keys that are strictly necessary to send.
        # Bad CC or bad SMTP_PORT would surface as the SMTP library's
        # own error — handled in the try/except below.
        if not config.get('EMAIL_SENDER'):
            return False, "Email config missing EMAIL_SENDER."
        if not config.get('EMAIL_PASSWORD'):
            return False, "Email config missing EMAIL_PASSWORD."
        if not config.get('DEFAULT_RECIPIENT'):
            return False, "Email config missing DEFAULT_RECIPIENT."
        if not config.get('SMTP_SERVER'):
            return False, "Email config missing SMTP_SERVER."

        # ── Build message ──────────────────────────────────────────────
        try:
            subject = EmailBuilder.build_subject(result)
            html = EmailBuilder.build_html(result)
        except (KeyError, ValueError, AttributeError) as e:
            logging.exception("Email build failed")
            return False, f"Could not build email body: {e}"

        msg = EmailSender._assemble_message(subject, html, config)

        # ── Send via SMTP ──────────────────────────────────────────────
        return EmailSender._deliver(msg, config)

    # ── Internal helpers ───────────────────────────────────────────────

    @staticmethod
    def _assemble_message(
        subject: str,
        html: str,
        config: Dict[str, Any],
    ) -> EmailMessage:
        """
        Build the RFC 5322 :class:`EmailMessage` with plain-text fallback
        plus the HTML alternative.

        Plain-text body is a single hint line telling the recipient
        their client doesn't support HTML. Every modern client does,
        but including the fallback makes the message pass spam
        filters with better scores.
        """
        msg = EmailMessage()
        msg['From'] = config['EMAIL_SENDER']
        msg['To'] = config['DEFAULT_RECIPIENT']

        cc_list = config.get('CC_RECIPIENTS') or []
        if cc_list:
            msg['Cc'] = ', '.join(cc_list)

        msg['Subject'] = subject

        # Text fallback (invisible to HTML-capable clients).
        msg.set_content(
            "This email contains an HTML report. "
            "Please view it in an HTML-capable client."
        )
        msg.add_alternative(html, subtype='html')

        return msg

    @staticmethod
    def _deliver(
        msg: EmailMessage,
        config: Dict[str, Any],
    ) -> Tuple[bool, str]:
        """
        Push the assembled message through SMTP.

        Uses STARTTLS on Gmail's default port 587. The connection is
        explicitly closed via ``quit()`` on success and ``close()``
        in the exception paths — needed because ``smtplib.SMTP`` does
        not implement context-manager cleanup on every Python version.

        Returns:
            ``(success, error_reason)`` — see :meth:`send`.
        """
        server_host = config['SMTP_SERVER']
        server_port = int(config.get('SMTP_PORT', 587))
        sender = config['EMAIL_SENDER']
        password = config['EMAIL_PASSWORD']
        to_addr = config['DEFAULT_RECIPIENT']
        cc_list = config.get('CC_RECIPIENTS') or []

        server: smtplib.SMTP | None = None

        try:
            server = smtplib.SMTP(server_host, server_port, timeout=30)
            server.starttls()
            server.login(sender, password)

            recipients = [to_addr] + list(cc_list)
            server.send_message(msg, to_addrs=recipients)
            server.quit()

            logging.info(
                "Email delivered to %s (+ %d CC)",
                to_addr, len(cc_list),
            )
            return True, ""

        except smtplib.SMTPAuthenticationError as e:
            logging.error("SMTP auth failed: %s", e)
            EmailSender._safe_close(server)
            return False, (
                f"Authentication failed — double-check EMAIL_PASSWORD "
                f"(must be a Gmail App Password, not the account "
                f"password). ({e.smtp_code})"
            )

        except smtplib.SMTPException as e:
            logging.error("SMTP error: %s", e)
            EmailSender._safe_close(server)
            return False, f"SMTP error: {e}"

        except OSError as e:
            # Covers network unreachable, DNS failure, connection
            # refused, firewall blocks, TLS handshake issues.
            logging.error("Network error during SMTP: %s", e)
            EmailSender._safe_close(server)
            return False, (
                f"Network error — check internet connection and that "
                f"port {server_port} is reachable. ({e})"
            )

    @staticmethod
    def _safe_close(server: smtplib.SMTP | None) -> None:
        """
        Close the SMTP connection, swallowing any error.

        Used from the exception paths — if we're already handling an
        error, another error from ``close()`` would just mask the
        original cause.
        """
        if server is None:
            return
        try:
            server.close()
        except (smtplib.SMTPException, OSError):
            pass