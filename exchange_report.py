import base64
import csv
import io
import logging
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any, Collection, Iterator, Mapping
from urllib.parse import quote, urlsplit
from uuid import UUID

import requests
from azure.core.credentials import TokenCredential
from azure.core.exceptions import AzureError

GRAPH_ROOT = "https://graph.microsoft.com"
GRAPH_SCOPE = f"{GRAPH_ROOT}/.default"
TRACE_PATH = "/v1.0/admin/exchange/tracing/messageTraces"
MAX_ATTACHMENT_BYTES = 2_500_000
STATUS_VALUES = {"failed": "failed", "filteredasspam": "filteredAsSpam", "quarantined": "quarantined"}
ATOM = r"[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+"
DOMAIN_LABEL = r"[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?"
EMAIL = re.compile(rf"{ATOM}(?:\.{ATOM})*@{DOMAIN_LABEL}(?:\.{DOMAIN_LABEL})+")


class ReportError(RuntimeError):
    pass


def _emails(env: Mapping[str, str], key: str, *, max_addresses: int | None = None) -> tuple[str, ...]:
    values = tuple(dict.fromkeys(value.strip().lower() for value in env.get(key, "").split(",")))
    if any(not EMAIL.fullmatch(value) or len(value) > 254 or len(value.rsplit("@", 1)[0]) > 64 for value in values):
        raise ReportError(f"{key} must contain comma-separated email addresses.")
    if max_addresses is not None and len(values) > max_addresses:
        raise ReportError(f"{key} must have at most {max_addresses} distinct addresses.")
    return values


def _flag(env: Mapping[str, str], key: str) -> bool:
    value = env.get(key, "false").strip().lower()
    if value not in {"true", "false"}:
        raise ReportError(f"{key} must be true or false.")
    return value == "true"


def _sender(env: Mapping[str, str]) -> str:
    value = env.get("REPORT_SENDER", "").strip()
    try:
        return str(UUID(value))
    except ValueError:
        if len(value) <= 254 and EMAIL.fullmatch(value):
            return value
        raise ReportError("REPORT_SENDER must be a mailbox Entra object ID or user principal name.") from None


@dataclass(frozen=True)
class Settings:
    ap_addresses: tuple[str, ...]
    report_recipients: tuple[str, ...]
    sender: str
    statuses: tuple[str, ...]
    send_enabled: bool
    include_subject: bool

    @classmethod
    def from_environment(cls, env: Mapping[str, str] | None = None) -> "Settings":
        env = os.environ if env is None else env
        configured_statuses = env.get("REPORT_STATUSES", "failed").split(",")
        statuses = tuple(dict.fromkeys(value.strip().lower() for value in configured_statuses))
        if set(statuses) - STATUS_VALUES.keys():
            raise ReportError("REPORT_STATUSES must contain failed, filteredAsSpam, and/or quarantined.")
        return cls(
            ap_addresses=_emails(env, "AP_INBOXES"),
            report_recipients=_emails(env, "REPORT_RECIPIENTS", max_addresses=500),
            sender=_sender(env),
            statuses=tuple(STATUS_VALUES[status] for status in statuses),
            send_enabled=_flag(env, "REPORT_SEND_ENABLED"),
            include_subject=_flag(env, "REPORT_INCLUDE_SUBJECT"),
        )


def prior_week(now: datetime) -> tuple[datetime, datetime]:
    if now.tzinfo is None:
        raise ValueError("Report time must be timezone-aware.")
    now = now.astimezone(timezone.utc)
    end = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    return end - timedelta(days=7), end


def _iso_utc(value: datetime) -> str:
    return value.isoformat().replace("+00:00", "Z")


def _odata(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _valid_trace_link(url: str) -> bool:
    try:
        parsed = urlsplit(url)
    except ValueError:
        return False
    return (parsed.scheme, parsed.netloc, parsed.path, parsed.fragment) == (
        "https", "graph.microsoft.com", TRACE_PATH, ""
    )


class GraphClient:
    def __init__(self, session: requests.Session, credential: TokenCredential) -> None:
        self.session = session
        self.credential = credential

    def _request(
        self,
        method: str,
        url: str,
        operation: str,
        expected_status: int,
        **kwargs: Any,
    ) -> requests.Response:
        retryable = {429, 500, 502, 503, 504} if method == "GET" else {429}
        for attempt in range(4):
            try:
                token = self.credential.get_token(GRAPH_SCOPE).token
            except AzureError:
                raise ReportError("Could not get a Microsoft Graph token from the managed identity.") from None
            try:
                response = self.session.request(
                    method, url,
                    headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                    timeout=(5, 30), allow_redirects=False, **kwargs,
                )
            except (requests.ConnectionError, requests.Timeout):
                if method != "GET":
                    raise ReportError(
                        f"{operation} did not return a response; delivery may have been accepted."
                    ) from None
                if attempt == 3:
                    break
                time.sleep(2 ** attempt)
                continue
            except requests.RequestException:
                raise ReportError(f"{operation} could not complete.") from None
            if response.status_code == expected_status:
                return response
            retry_after = response.headers.get("Retry-After", "")
            delay = int(retry_after) if re.fullmatch(r"[0-9]{1,9}", retry_after) else 2 ** attempt
            if attempt == 3 or response.status_code not in retryable or delay > 120:
                raise ReportError(f"{operation} failed (Graph HTTP {response.status_code}).")
            time.sleep(delay)
        raise ReportError(f"{operation} failed after retries.")

    def traces(self, address: str, status: str, start: datetime, end: datetime) -> Iterator[dict[str, Any]]:
        url = f"{GRAPH_ROOT}{TRACE_PATH}"
        params: dict[str, str | int] | None = {
            "$top": 5000,
            "$filter": f"receivedDateTime ge {_iso_utc(start)} and receivedDateTime le {_iso_utc(end)}"
                       f" and recipientAddress eq {_odata(address)} and status eq {_odata(status)}",
        }
        seen_links: set[str] = set()
        while True:
            response = self._request("GET", url, "Message trace", 200, params=params)
            try:
                page = response.json()
            except ValueError:
                raise ReportError("Message trace returned invalid JSON.") from None
            if not isinstance(page, dict) or not isinstance(page.get("value"), list):
                raise ReportError("Message trace returned an invalid page.")
            for item in page["value"]:
                if not isinstance(item, dict):
                    raise ReportError("Message trace returned an invalid row.")
                yield item
            next_link = page.get("@odata.nextLink")
            if next_link is None:
                return
            if not isinstance(next_link, str) or not _valid_trace_link(next_link) or next_link in seen_links:
                raise ReportError("Message trace returned an invalid or repeated continuation link.")
            seen_links.add(next_link)
            url, params = next_link, None

    def send(
        self,
        settings: Settings,
        subject: str,
        body: str,
        filename: str,
        content: bytes,
    ) -> None:
        self._request(
            "POST",
            f"{GRAPH_ROOT}/v1.0/users/{quote(settings.sender, safe='')}/sendMail",
            "Send report",
            202,
            json={
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": [{"emailAddress": {"address": address}} for address in settings.report_recipients],
                    "attachments": [{
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": filename,
                        "contentType": "text/csv",
                        "contentBytes": base64.b64encode(content).decode("ascii"),
                    }],
                },
            },
        )


@dataclass(frozen=True)
class Trace:
    received_at: datetime
    sender: str
    recipient: str
    status: str
    message_id: str
    trace_id: str
    subject: str

    @classmethod
    def from_graph(cls, row: Mapping[str, Any]) -> "Trace":
        def field(name: str, *, required: bool = False) -> str:
            value = row.get(name)
            if value is None and not required:
                return ""
            if not isinstance(value, str) or (required and not value):
                raise ReportError("Message trace returned an invalid row.")
            return value

        try:
            received_at = datetime.fromisoformat(field("receivedDateTime", required=True).replace("Z", "+00:00"))
            if received_at.tzinfo is None:
                raise ValueError
        except ValueError:
            raise ReportError("Message trace returned an invalid received timestamp.") from None
        return cls(
            received_at.astimezone(timezone.utc),
            field("senderAddress"),
            field("recipientAddress", required=True),
            field("status", required=True),
            field("messageId"),
            field("id", required=True),
            field("subject"),
        )


def _safe_cell(value: str) -> str:
    if value.startswith(("\t", "\r", "\n")) or value.lstrip(" \t\r\n\ufeff").startswith(
        ("=", "+", "-", "@", "＝", "＋", "－", "＠")
    ):
        return "'" + value
    return value


def csv_parts(traces: Collection[Trace], include_subject: bool, max_bytes: int = MAX_ATTACHMENT_BYTES) -> list[bytes]:
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, quoting=csv.QUOTE_ALL)

    def line(cells: Collection[str]) -> bytes:
        buffer.seek(0)
        buffer.truncate(0)
        writer.writerow(cells)
        return buffer.getvalue().encode("utf-8")

    fields = ("received_at_utc", "sender", "ap_recipient", "status", "message_id", "trace_id")
    header = b"\xef\xbb\xbf" + line(fields + (("subject",) if include_subject else ()))
    result = []
    current = bytearray(header)
    for trace in traces:
        cells: tuple[str, ...] = (
            _iso_utc(trace.received_at), trace.sender, trace.recipient, trace.status, trace.message_id, trace.trace_id
        )
        if include_subject:
            cells += (trace.subject,)
        row = line(tuple(map(_safe_cell, cells)))
        if len(header) + len(row) > max_bytes:
            raise ReportError("One trace row exceeds the maximum CSV attachment size.")
        if len(current) + len(row) > max_bytes:
            result.append(bytes(current))
            current = bytearray(header)
        current.extend(row)
    result.append(bytes(current))
    return result


def run_report(client: GraphClient, settings: Settings, now: datetime) -> None:
    start, end = prior_week(now)
    unique: dict[tuple[str, str], Trace] = {}
    for address in settings.ap_addresses:
        for status in settings.statuses:
            for row in client.traces(address, status, start, end):
                trace = Trace.from_graph(row)
                if (
                    start <= trace.received_at < end
                    and trace.recipient.lower() == address
                    and trace.status.lower() == status.lower()
                ):
                    unique.setdefault((trace.trace_id.lower(), address), trace)
    traces = sorted(unique.values(), key=lambda trace: (trace.received_at, trace.recipient, trace.trace_id))
    csv_files = csv_parts(traces, settings.include_subject)
    label = f"{start:%Y-%m-%d} to {end - timedelta(days=1):%Y-%m-%d} UTC"
    logging.info(
        "Prepared AP rejection report for %s: %d rows across %d CSV files.", label, len(traces), len(csv_files)
    )
    if not settings.send_enabled:
        logging.info("Report sending disabled; no email sent.")
        return

    inboxes = "\n".join(f"- {address}" for address in settings.ap_addresses)
    body = (
        f"Exchange Online message trace for {label}.\n"
        f"Matching message-recipient pairs: {len(traces)}.\n"
        f"Statuses: {', '.join(settings.statuses)}.\n\n"
        f"AP inboxes covered by this report:\n{inboxes}\n\n"
        "This report includes only failures visible in Exchange Online message trace."
    )
    for index, content in enumerate(csv_files, start=1):
        part = f" (part {index} of {len(csv_files)})" if len(csv_files) > 1 else ""
        suffix = f"-part-{index:03d}" if part else ""
        client.send(
            settings,
            f"Weekly AP rejected email report: {label}{part}",
            body,
            f"ap-rejections-{start:%Y-%m-%d}{suffix}.csv",
            content,
        )
        logging.info("Graph accepted part %d of %d of the AP rejection report.", index, len(csv_files))
