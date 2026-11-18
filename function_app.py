import logging
from datetime import datetime, timezone

import azure.functions as func
import requests
from azure.identity import ManagedIdentityCredential

from exchange_report import GraphClient, Settings, run_report

app = func.FunctionApp()


@app.function_name(name="WeeklyApRejectionReport")
@app.timer_trigger(schedule="%REPORT_SCHEDULE%", arg_name="timer", use_monitor=True)
def weekly_ap_rejection_report(timer: func.TimerRequest) -> None:
    if timer.past_due:
        logging.warning("AP rejection report timer is past due.")

    settings = Settings.from_environment()
    with ManagedIdentityCredential() as credential, requests.Session() as session:
        run_report(GraphClient(session, credential), settings, datetime.now(timezone.utc))
