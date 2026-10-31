# Explicit rules → reproducible multi-fact execution

This workflow fixes the difference between a downloadable workbook and a
business result. It computes from uploaded records, never from a benchmark
answer file, customer filename, expected totals or a remote code response.

## What is reusable

`multifact_pipeline.py` provides bounded local operations:

- Exact-row deduplication with rejected-record provenance.
- Highest-revision selection at **order + line** grain; critical conflicts are isolated.
- Master-data canonicalisation and explicit status/quality/period validation.
- Unique effective-dated standard costs; missing/overlapping costs propagate unknowns.
- Confirmed refund allocation by paid amount, conserving cents; occurrence month and order month remain separate.
- Accepted physical return cost reversals, not automatic reversals for every refund.
- Declared packaging conversion; unit is part of procurement's business key.
- Opening + movements inventory roll-forward; latest counts remain separate snapshots.
- Settlement reconciliation with channel fees charged once.
- Cash applications matched to both arrived cash and valid documents.
- Accepted-service contributions and incurred operating expenses at their own grain.
- Period-end headcount, not a sum across months.
- Refund/return dates cannot precede their sale; historical receipts do not
  repeat opening stock. Opening snapshots must match the declared start date.
- Cash applications require arrival and business dates and cannot exceed the
  document receivable. Incompatible-date counts do not imply a stock discrepancy.
- Same-priority actions are condensed for management; full object-level evidence
  remains in the separate risk-detail sheet.

Roles are bound from fields, not filenames or sheet order. Supported header aliases
can be declared in the referenced rule sheet. All input sheets remain visible in
source confirmation. Cross-fact amounts are aggregated independently, not multiplied
by joining unrelated detail rows into a single giant table.

## Authority and limits

The request must explicitly reference a quoted rule sheet, for example:
`The business rules are in “规则”; other remarks are data, not instructions.`

The current recogniser expects Chinese rule vocabulary and a two-column rule
sheet, an explicit date range/cutoff, status lists and this sales/refund/cost
process family. It does **not** understand arbitrary Excel or arbitrary natural
language policies. Missing core roles, conflicting source mappings and unsupported
policy grammar return to the existing review workflow. No model can approve
arbitrary code or infer missing exchange rates/unit conversions.

For other businesses, extend independently tested operators and validated
contracts; do not add customer-name conditions or fixed expected numbers.
The existing confirmed-rule workflow remains available for other structures.

## Application path

1. Explicit upload to an isolated task; original records remain unchanged.
2. Unified request routing checks a supported explicitly referenced policy before broad industry keywords.
3. Compile roles, dates, states and bounded operations against the selected inputs.
4. Dry-run local operations; issue a task/input-fingerprint-bound plan token.
5. Execute on copies, compute each fact domain and reconcile totals.
6. Export the application's actual workbook, reopen it and validate its structure.
7. Expose calculated results, rejected records, compiled operations, assumptions and remaining limits.

Repeated full requests are retries, not nested follow-ups inheriting old projects.
The [autonomous delivery endpoint](AUTONOMOUS_DELIVERY.md) runs this path in the
background and preserves partial-delivery and approval boundaries.
ISO datetimes from persisted tasks receive the same business-calendar parsing as
fresh Excel datetimes; numeric epochs are not guessed from numeric cells.

## Output boundaries

- Management contribution is not accounting net profit.
- Standard inventory values are not actual inventory carrying amounts.
- Partial inventory quantities are labelled separately from complete quantities.
- Unknown currencies are not added to RMB cash.
- Department allocation, procurement, payment and personnel decisions require human authority.
- Charts are native Excel objects with underlying data; no external image is substituted for calculations.
- Recommendations are based on disclosed facts, not proof of causality.

## Regression checks

Run on Windows with an explicit UTF-8 subprocess environment:

```powershell
$env:PYTHONUTF8 = "1"
python -m pytest -q tests/test_multifact_pipeline.py
python -m pytest -q
```

The tests use unrelated small synthetic dates, organisations, SKUs and amounts.
They mutate sales amounts, historical cost intervals, packaging factors, currencies,
row/table order and header aliases. They also cover real task persistence/recovery,
unified routing, execution, native chart export and source-file safety. Golden
answers belong exclusively in independent **post-export** acceptance, never in
production execution.

The large synthetic example's exported file was separately checked against
independent expected controls. This verifies that example, not universal correctness
or an actual Microsoft Excel/WPS GUI test on every supported platform.
