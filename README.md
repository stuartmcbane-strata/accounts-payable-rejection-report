# Weekly Accounts Payable rejection report

An Azure Functions v4 Python timer reads [Exchange Online message trace](https://learn.microsoft.com/en-us/graph/api/messagetracingroot-list-messagetraces?view=graph-rest-1.0) through Microsoft Graph and emails the results as an Excel-compatible CSV. It does not read the AP mailboxes.

By default it runs Mondays at **08:00 UTC** and reports the previous full Monday-to-Sunday UTC week, based on when Exchange Online received the message. It queries each configured AP address, including addresses for shared mailboxes, and collects all result pages before emailing. The report contains one row per message trace and AP recipient, with the UTC timestamp, sender, AP recipient, trace status, message ID, and trace ID. Subject lines are optional; original bodies and attachments are never accessed. A successful run with no matches sends a header-only CSV so recipients can distinguish zero results from a failed run. Oversize reports arrive as numbered CSVs in separate emails.

**Coverage:** The default status is `failed` (Exchange could not deliver). Spam filtering and quarantine are separate statuses and can be included explicitly. This doesn't cover on-premises-only Exchange, upstream gateways, or rejections that never created a trace record, such as [Exchange Online connection filtering](https://learn.microsoft.com/en-us/defender-office-365/connection-filter-policies-configure). The report does not look up an individual SMTP failure reason; use the trace ID and recipient in Exchange admin center for investigation.

## Settings

Set these under the Function App's **Settings → Environment variables** before publishing. The deployment does not upload local settings.

| Setting | Value |
| --- | --- |
| `REPORT_SCHEDULE` | Azure NCRONTAB in UTC; use `0 0 8 * * 1` for Monday at 08:00. Changing this does not change the reporting week. |
| `AP_INBOXES` | Comma-separated SMTP addresses to query and list in the email body, including any aliases shown separately in trace. |
| `REPORT_RECIPIENTS` | Comma-separated list of approved report recipients or a distribution list. |
| `REPORT_SENDER` | Dedicated Exchange Online sender mailbox's Entra object ID (preferred) or sign-in UPN, not an arbitrary SMTP alias. |
| `REPORT_STATUSES` | Optional; defaults to `failed`. Add `filteredAsSpam` and/or `quarantined` if wanted. |
| `REPORT_INCLUDE_SUBJECT` | Optional; defaults to `false`. Set to `true` only if the list should receive subjects. |
| `REPORT_SEND_ENABLED` | Optional; defaults to `false`. Enable only after checking trace and recipients. |

## Deploy and grant permissions

1. In Azure, create a Linux Function App on **Flex Consumption** with **Python 3.12**, host storage, and Application Insights enabled. Leave the host time zone at UTC. Under **Identity**, enable its **system-assigned managed identity**. Record the tenant ID and the identity's principal (object) ID. No application secrets are required.
2. An Entra administrator provisions the [Microsoft message trace service principal](https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/graph-api-message-trace) and assigns the identity the [Microsoft Graph app role](https://learn.microsoft.com/en-us/entra/identity/managed-identities-azure-resources/how-to-assign-app-role-managed-identity-powershell) for message trace. That role is tenant-wide; `AP_INBOXES` restricts what this code retrieves, not what the identity is allowed to read.

   In Microsoft Graph PowerShell, replace the two placeholder IDs:

   ```powershell
   $tenantId = '<Entra tenant ID>'
   $identityObjectId = '<Function App managed identity principal ID>'
   Connect-MgGraph -TenantId $tenantId -Scopes 'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All'

   $traceProviderId = '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373'
   if (-not (Get-MgServicePrincipal -Filter "appId eq '$traceProviderId'")) {
       New-MgServicePrincipal -AppId $traceProviderId | Out-Null
   }
   $mi = Get-MgServicePrincipal -ServicePrincipalId $identityObjectId
   $graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
   $traceRole = $graph.AppRoles | Where-Object {
       $_.Value -eq 'ExchangeMessageTrace.Read.All' -and $_.IsEnabled -and $_.AllowedMemberTypes -contains 'Application'
   }
   if (-not $traceRole) { throw 'Microsoft Graph message trace application role not found.' }
   $assignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mi.Id -All |
       Where-Object { $_.ResourceId -eq $graph.Id -and $_.AppRoleId -eq $traceRole.Id }
   if (-not $assignment) {
       New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mi.Id -BodyParameter @{
           principalId = $mi.Id
           resourceId = $graph.Id
           appRoleId = $traceRole.Id
       } | Out-Null
   }
   ```

3. An Exchange administrator uses [Application RBAC](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac) to grant `Application Mail.Send` **only** for the dedicated report mailbox. Do **not** also grant tenant-wide `Mail.Send` in Entra: those rights are additive and would defeat the scope. Run the following once per managed identity and report mailbox, using the mailbox's real primary address:

   ```powershell
   Connect-ExchangeOnline
   $senderAddress = 'ap-reports@example.com'
   $mailbox = Get-EXOMailbox -Identity $senderAddress -Properties ExternalDirectoryObjectId -ErrorAction Stop
   $addressFilter = "EmailAddresses -eq 'smtp:$($senderAddress.Replace("'", "''"))'"
   $scopeName = 'AP report sender only'
   $scopeMatches = @(Get-Recipient -Filter $addressFilter -ResultSize 2 -ErrorAction Stop)
   if ($scopeMatches.Count -ne 1 -or $scopeMatches[0].PrimarySmtpAddress.ToString() -ne $senderAddress) {
       throw 'Sender scope did not match exactly the dedicated mailbox.'
   }
   New-ServicePrincipal -AppId $mi.AppId -ObjectId $mi.Id -DisplayName 'Weekly AP report'
   New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $addressFilter
   New-ManagementRoleAssignment -Name 'Weekly AP report Mail.Send' -App $mi.AppId `
       -Role 'Application Mail.Send' -CustomResourceScope $scopeName
   Test-ServicePrincipalAuthorization -Identity $mi.Id -Resource $senderAddress
   $mailbox.ExternalDirectoryObjectId  # Set REPORT_SENDER to this value.
   ```

   Check the test shows `Application Mail.Send` in scope for the sender and out of scope for a different mailbox. Both Microsoft message trace provisioning and Exchange permission changes can take time to propagate.

4. With Azure Functions Core Tools v4 and Python 3.12, sign in using `az login`. From the project directory, publish to your new Function App (Python uses a remote build by default):

   ```powershell
   $functionAppName = 'your-function-app-name'
   func azure functionapp publish $functionAppName
   ```

5. With sending still disabled, select **Functions → WeeklyApRejectionReport → Code + Test → Test/Run**. Compare the logged count with Exchange admin center **Mail flow → Message trace**, for the same UTC dates, recipients and statuses. Then set `REPORT_RECIPIENTS` to one approved test recipient, enable `REPORT_SEND_ENABLED`, and manually run again to verify the CSV arrives. Restore the production recipient list after verification.

## Operations

- Monitor successful executions and alert on failures or a missing weekly execution in Application Insights. Check the sender's Sent Items or Exchange trace if Graph accepted the request but a recipient didn't receive it. Report content and addresses are not written to application logs.
- A trace request failure aborts the run before any mail is sent. Graph acceptance means queued, not delivered. If a send timed out or only some CSV parts were accepted, check Sent Items before rerunning. A manual run recomputes the prior full week and can resend a report; missed older weeks aren't automatically backfilled.
- The report includes finance-related metadata. Use approved recipients, preferably an internal distribution list, and handle the attachment according to your retention policy.

## Local checks

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
python -m ruff check .
python -m mypy exchange_report.py function_app.py
```

The automated tests use mocked Graph responses, not tenant data. To index the timer locally with Core Tools and Azurite, copy the local settings example to `local.settings.json`. The example disables the timer locally because the live function uses an Azure managed identity.
