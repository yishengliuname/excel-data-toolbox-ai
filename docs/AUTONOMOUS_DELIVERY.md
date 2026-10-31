# Bounded autonomous local delivery

A goal should produce an inspected artifact, not merely a plan or a success
message. The unified UI now submits a durable local background job through
`POST /api/ai-run`. Existing business operators and approval checks perform the
work; this is not an arbitrary-code agent or a new industry template.

## User workflow

1. Upload the source files into a task. Derived reports are unselected by default.
2. State the business goal; explicitly identify authoritative rules if provided.
3. The job binds the selected originals and their content fingerprints.
4. Existing routing compiles and validates a supported local plan.
5. Unambiguous, whitelisted local operations execute on copies. Unconfirmed
   business rules return to the existing review/confirmation workflow.
6. Export and validation use the existing application delivery path. A workbook
   delivery is not successful unless there is a real downloadable file.
7. The UI polls `/api/jobs`, displays progress and retains a download link.
   Reloading the page restores the current task's latest delivery.

Chart-only requests can return a chart rather than a workbook. An unfinished
chart follow-up or unresolved evidence is a **partial** delivery, not a complete
business conclusion. The job engine's `completed` state only means its handler
returned; check the result's `delivery_status` as well.

## Confirm once, reuse within an explicit authorization

The business-rule review panel now offers optional local automation grants.
Enter the customer scope and profile title, check the rules, and explicitly
select automatic reuse. Set the validity period (1–365 days), maximum total
records, and maximum excluded fraction per table. Defaults do **not** grant
automatic reuse; old profiles remain review-only.

On a subsequent task, choose the uploaded originals, identify the customer,
query their profiles and click **按已授权配置自动执行**. The server checks the
schema, unambiguous source binding, status dictionary, duplicate-key conflicts,
metric units/currency, dates, record/exclusion limits, request topics, rule
fingerprint and engine revision. It checks again immediately before execution.
An expired/revoked/drifted grant stops execution with a specific explanation.
Engine changes require renewed review rather than silently changing meaning.

Approved dimension joins can supply selected master-data attributes only if
keys are unique and fact-key coverage is complete. This is evidence reuse, not
permission to scan directories, invent unit conversions or join facts and
inflate amounts. Changing data values within a valid contract computes fresh
results; it does not reuse old totals.

`POST /api/ai-run` optionally accepts `customer_scope` and `profile_id`.
The initial human `POST /api/business-confirm` optionally accepts
`automation_limits` (`valid_days`, `max_rows`, `max_excluded_ratio`). No AI plan
can approve those fields. `POST /api/business-profile-revoke` requires the
customer, profile ID and `confirmed: true`. Revocation does not delete prior
deliverables, but prevents a new submission from bypassing it through the cache.
These scopes organize local operator configuration; they are **not** multi-user
authentication or tenant security.

## Task-level acceptance and bounded recovery

File acceptance reopens the XLSX and compares its contents. Task acceptance
separately checks unchanged inputs, completed planned steps, export acceptance,
pending review items and each declared analysis topic's evidence coverage.
Missing topic evidence is **needs_review**, never an implicit pass. The UI shows
the checks and evidence gaps. `completed` jobs may therefore deliver a useful
**partial** workbook. This is not certification of every possible question,
accounting audit, causal inference or an external transaction.

Only typed `TimeoutError`/`ConnectionError` failures in pre-execution planning
receive one bounded retry, with a sanitized event record. Business-rule errors,
execution errors and ambiguous state are not retried. Cancellation is checked
again before execution; neither retry nor saved authorization enables external
payment, purchase, publishing or HR actions.

## Safety and recovery

- Same task + same goal + same selected content reuses an active/completed job.
  Concurrent duplicate requests are deduplicated transactionally.
- Changed inputs invalidate a queued goal instead of applying an old plan.
- Background session scopes do not change another browser's default task.
- Cancellation is cooperative at safe stage boundaries, not an immediate rollback.
  A running calculation may have written partial output before cancellation.
- An executing checkpoint is never blindly replayed after a crash. Interrupted
  jobs require reconciliation; a repeated goal returns the interrupted record.
- A completed checkpoint can restore its file link without repeating calculations.
- Failed pre-execution jobs may be resubmitted. A missing saved file requires
  regeneration after checking the task state.
- This is local single-process orchestration, not a distributed exactly-once
  system. Queued tasks also revalidate the restored source signatures on restart.

Publishing, payment, ordering, deleting source files and external engineering
actions are outside this autonomous endpoint. Optional AI can help plan; it cannot
approve missing business rules or grant authority for these actions.

## Quality boundaries

The explicitly declared multi-fact process also checks refund/return dates,
opening and movement periods, nonnegative standard costs, cash/application timing,
document capacity and same-date inventory counts. Tiny refund allocations conserve
cents without a negative rounding tail.

`关键问题与行动` groups the same priority/problem/action for a shorter boss view;
`风险行动明细` retains every original object and evidence row. Counts are grouped,
but incompatible amounts/units are not summed into a misleading headline.

This does **not** promise to complete every boss task. Unknown businesses, absent
facts, ambiguous units and causal claims require more evidence. A safe partial
report with disclosed gaps is preferable to invented profitability or decisions.

## Verification

```powershell
$env:PYTHONUTF8 = "1"
python -m pytest -q tests/test_autopilot.py tests/test_multifact_pipeline.py tests/test_trusted_automation.py
python -m pytest -q
```

Tests cover actual HTTP submission/execution/download, artifact contents,
parallel duplicate requests, reload/checkpoint recovery, changed-input rejection,
cancellation boundaries and isolation from derived outputs. Independent benchmark
answers are read only after the actual application exports a workbook, never by
the business execution code. These tests are not native Office GUI certification.
