import base64
import csv
import io
import json
from datetime import datetime, timedelta, timezone
from unittest.mock import Mock

import pytest
import requests
from azure.core.credentials import AccessToken

import exchange_report
from exchange_report import GRAPH_ROOT, TRACE_PATH, GraphClient, ReportError, Settings, Trace, csv_parts, run_report
from function_app import app

NOW = datetime(2026, 11, 16, 8, tzinfo=timezone.utc)
SENDER_ID = "dec0913a-82ca-48ec-a44d-6d87630f6fe1"
ENV = {
    "AP_INBOXES": "AP@example.com, ap-uk@example.com, AP@example.com",
    "REPORT_RECIPIENTS": "finance@example.com, alerts@example.com",
    "REPORT_SENDER": SENDER_ID,
}


def reply(code, data=None, **headers):
    response = requests.Response()
    response.status_code = code
    response._content = json.dumps(data).encode()
    response.headers.update(headers)
    return response


def setup_client(*responses):
    session = Mock(spec=requests.Session)
    session.request.side_effect = responses
    credential = Mock()
    credential.get_token.return_value = AccessToken("fake-token", 2**31)
    return GraphClient(session, credential), session.request


def trace(trace_id="id-1", **changes):
    return {
        "id": trace_id,
        "receivedDateTime": "2026-11-09T00:00:00.0000000Z",
        "senderAddress": "vendor@example.net",
        "recipientAddress": "AP@example.com",
        "status": "failed",
        "messageId": "<msg@example.net>",
        "subject": 'Invoice, "Nov" – 2026',
        **changes,
    }


def report_rows(content):
    return list(csv.DictReader(io.StringIO(content.decode("utf-8-sig"), newline="")))


def test_config_and_timer_binding():
    settings = Settings.from_environment(ENV)
    assert settings.ap_addresses == ("ap@example.com", "ap-uk@example.com")
    assert settings.statuses == ("failed",)
    assert settings.sender == SENDER_ID
    assert not settings.send_enabled
    assert not settings.include_subject
    timer, = app.get_functions()
    assert timer.get_function_name() == "WeeklyApRejectionReport"
    binding, = json.loads(timer.get_function_json())["bindings"]
    assert binding["type"] == "timerTrigger"
    assert binding["schedule"] == "%REPORT_SCHEDULE%"
    assert binding["useMonitor"] is True


@pytest.mark.parametrize(("setting", "value"), [
    ("AP_INBOXES", ""),
    ("AP_INBOXES", "ap@example.com,"),
    ("REPORT_RECIPIENTS", "Finance <finance@example.com>"),
    ("REPORT_RECIPIENTS", "finance@example.com\r\nBCC: other@example.net"),
    ("REPORT_SENDER", "ap@example.com, finance@example.com"),
    ("REPORT_STATUSES", "delivered"),
    ("REPORT_SEND_ENABLED", "y"),
])
def test_invalid_settings_fail_closed(setting, value):
    with pytest.raises(ReportError, match=setting):
        Settings.from_environment(ENV | {setting: value})


def test_week_uses_utc_monday_boundary():
    previous_monday = datetime(2026, 11, 9, tzinfo=timezone.utc)
    monday = previous_monday + timedelta(days=7)
    assert exchange_report.prior_week(NOW) == (previous_monday, monday)
    sunday_local = datetime(2026, 11, 15, 20, tzinfo=timezone(timedelta(hours=-5)))
    assert exchange_report.prior_week(sunday_local) == (previous_monday, monday)
    assert exchange_report.prior_week(monday - timedelta(microseconds=1)) == (
        previous_monday - timedelta(days=7), previous_monday
    )


def test_report_pages_every_ap_address_deduplicates_and_sends_csv():
    next_link = f"{GRAPH_ROOT}{TRACE_PATH}?%24skiptoken=page-2"
    last_trace = trace(
        "id-1",
        recipientAddress="ap-uk@example.com",
        receivedDateTime="2026-11-15T23:59:59.9999999Z",
    )
    client, request = setup_client(
        reply(200, {"value": [trace()], "@odata.nextLink": next_link}),
        reply(200, {"value": [
            trace(),
            trace("end", receivedDateTime="2026-11-16T00:00:00Z"),
            trace("early", receivedDateTime="2026-11-08T23:59:59.9999999Z"),
            trace("outside", recipientAddress="unrelated@example.com"),
            trace("status", status="delivered"),
        ]}),
        reply(200, {"value": [last_trace]}),
        reply(202),
    )
    run_report(client, Settings.from_environment(ENV | {
        "REPORT_SEND_ENABLED": "true", "REPORT_INCLUDE_SUBJECT": "true",
    }), NOW)

    first, continuation, second, mail = request.call_args_list
    assert first.args == ("GET", GRAPH_ROOT + TRACE_PATH)
    assert first.kwargs["params"]["$filter"] == (
        "receivedDateTime ge 2026-11-09T00:00:00Z and receivedDateTime le 2026-11-16T00:00:00Z"
        " and recipientAddress eq 'ap@example.com' and status eq 'failed'"
    )
    assert continuation.args == ("GET", next_link)
    assert continuation.kwargs["params"] is None
    assert "recipientAddress eq 'ap-uk@example.com'" in second.kwargs["params"]["$filter"]
    assert mail.args == ("POST", f"{GRAPH_ROOT}/v1.0/users/{SENDER_ID}/sendMail")
    assert all(call.kwargs["allow_redirects"] is False for call in request.call_args_list)
    message = mail.kwargs["json"]["message"]
    assert message["toRecipients"] == [
        {"emailAddress": {"address": "finance@example.com"}},
        {"emailAddress": {"address": "alerts@example.com"}},
    ]
    assert "2026-11-09 to 2026-11-15 UTC" in message["subject"]
    attachment, = message["attachments"]
    assert attachment["name"] == "ap-rejections-2026-11-09.csv"
    rows = report_rows(base64.b64decode(attachment["contentBytes"]))
    assert [row["ap_recipient"] for row in rows] == ["AP@example.com", "ap-uk@example.com"]
    assert [row["trace_id"] for row in rows] == ["id-1", "id-1"]
    assert rows[0]["subject"] == 'Invoice, "Nov" – 2026'


def test_multiple_statuses_opt_in_and_odata_escaping():
    client, request = setup_client(reply(200, {"value": []}), reply(200, {"value": []}))
    settings = Settings.from_environment(ENV | {
        "AP_INBOXES": "ap'uk@example.com",
        "REPORT_STATUSES": "FilteredAsSpam,quarantined",
    })
    run_report(client, settings, NOW)
    filters = [call.kwargs["params"]["$filter"] for call in request.call_args_list]
    assert filters[0].endswith("recipientAddress eq 'ap''uk@example.com' and status eq 'filteredAsSpam'")
    assert filters[1].endswith("recipientAddress eq 'ap''uk@example.com' and status eq 'quarantined'")


def test_empty_report_is_mailed_when_enabled():
    client, request = setup_client(reply(200, {"value": []}), reply(200, {"value": []}), reply(202))
    run_report(client, Settings.from_environment(ENV | {"REPORT_SEND_ENABLED": "true"}), NOW)
    message = request.call_args.kwargs["json"]["message"]
    body = message["body"]["content"]
    assert "Matching message-recipient pairs: 0." in body
    assert body.split("\n\n")[1].splitlines() == [
        "AP inboxes covered by this report:",
        "- ap@example.com",
        "- ap-uk@example.com",
    ]
    content = base64.b64decode(message["attachments"][0]["contentBytes"])
    assert content.decode("utf-8-sig") == (
        '"received_at_utc","sender","ap_recipient","status","message_id","trace_id"\r\n'
    )


def test_throttling_is_retried_without_sending_in_preview(monkeypatch):
    sleep = Mock()
    monkeypatch.setattr(exchange_report.time, "sleep", sleep)
    ok = reply(200, {"value": []})
    client, request = setup_client(reply(429, **{"Retry-After": "7"}), ok, ok)
    run_report(client, Settings.from_environment(ENV), NOW)
    sleep.assert_called_once_with(7)
    assert [call.args[0] for call in request.call_args_list] == ["GET"] * 3


@pytest.mark.parametrize("bad_link", [
    f"https://graph.microsoft.com.example.net{TRACE_PATH}",
    f"http://graph.microsoft.com{TRACE_PATH}",
    f"https://graph.microsoft.com@elsewhere.example{TRACE_PATH}",
    "https://graph.microsoft.com/v1.0/users",
])
def test_bad_continuation_aborts_without_sending_or_following(bad_link):
    client, request = setup_client(reply(200, {"value": [trace()], "@odata.nextLink": bad_link}))
    with pytest.raises(ReportError, match="continuation"):
        run_report(client, Settings.from_environment(ENV | {"REPORT_SEND_ENABLED": "true"}), NOW)
    assert request.call_count == 1


def test_repeated_continuation_aborts():
    page = reply(200, {"value": [], "@odata.nextLink": f"{GRAPH_ROOT}{TRACE_PATH}?$skiptoken=same"})
    client, request = setup_client(page, page)
    with pytest.raises(ReportError, match="continuation"):
        run_report(client, Settings.from_environment(ENV), NOW)
    assert request.call_count == 2


def test_trace_failure_does_not_send_partial_report():
    client, request = setup_client(reply(200, {"value": [trace()]}), reply(403, {"error": "private context"}))
    with pytest.raises(ReportError, match=r"^Message trace failed \(Graph HTTP 403\)\.$"):
        run_report(client, Settings.from_environment(ENV | {"REPORT_SEND_ENABLED": "true"}), NOW)
    assert [call.args[0] for call in request.call_args_list] == ["GET", "GET"]


def test_send_timeout_does_not_retry_and_avoids_leaking_request_data():
    client, request = setup_client(reply(200, {"value": []}), reply(200, {"value": []}), requests.Timeout("private"))
    with pytest.raises(ReportError, match="delivery may have been accepted") as caught:
        run_report(client, Settings.from_environment(ENV | {"REPORT_SEND_ENABLED": "true"}), NOW)
    assert "private" not in str(caught.value)
    assert [call.args[0] for call in request.call_args_list] == ["GET", "GET", "POST"]


def test_csv_formula_guard_and_attachment_sizing():
    traces = [
        Trace.from_graph(trace("a", subject=" =2+2", senderAddress="+vendor@example.com")),
        Trace.from_graph(trace("b", subject="＠SUM(A1)", senderAddress="\tvendor@example.com")),
        Trace.from_graph(trace("c", subject='Invoice,"line one"\nline two')),
    ]
    max_bytes = max(len(csv_parts([row], include_subject=True)[0]) for row in traces)
    files = csv_parts(traces, include_subject=True, max_bytes=max_bytes)
    assert len(files) == 3
    assert all(len(file) <= max_bytes for file in files)
    rows = [row for file in files for row in report_rows(file)]
    assert [row["subject"] for row in rows] == ["' =2+2", "'＠SUM(A1)", 'Invoice,"line one"\nline two']
    assert [row["sender"] for row in rows[:2]] == ["'+vendor@example.com", "'\tvendor@example.com"]
    assert "subject" not in report_rows(csv_parts(traces, include_subject=False)[0])[0]


def test_sends_numbered_parts(monkeypatch):
    client, request = setup_client(reply(200, {"value": []}), reply(200, {"value": []}), reply(202), reply(202))
    monkeypatch.setattr(exchange_report, "csv_parts", lambda *args: [b"one", b"two"])
    run_report(client, Settings.from_environment(ENV | {"REPORT_SEND_ENABLED": "true"}), NOW)
    messages = [call.kwargs["json"]["message"] for call in request.call_args_list[-2:]]
    assert [message["attachments"][0]["name"] for message in messages] == [
        "ap-rejections-2026-11-09-part-001.csv", "ap-rejections-2026-11-09-part-002.csv",
    ]
    assert messages[1]["subject"].endswith("(part 2 of 2)")
