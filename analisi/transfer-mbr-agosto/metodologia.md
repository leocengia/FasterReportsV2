# Methodology note — Transfer Data Submission, August 2026

Testo pronto da incollare nelle celle `Notes` del template e da allegare alla submission.
Va tradotto/adattato, ma i numeri e i limiti sono quelli misurati.

---

## 1. How transfers are identified

Our reporting source is the Salesforce case export that feeds the weekly Omni Report
(143 columns, one row per case). There is **no field in the export that records the queue a
case was transferred *to***. `Case Queue Name` holds only two values and is empty on 71% of
rows; `user_agent_queue_group_name` and `user_routing_profile` are constant across the whole
site. The destination is therefore reconstructed **indirectly**, from the receiving case.

**Receiving leg (the figures reported in Section 1).** A case created by a transfer is
identified by `Case Origin`:

| Template column | Rule |
|---|---|
| **Voice** | `Case Origin = "Transferred from Phone"` |
| **BOF** | `Case Origin = "Transferred from Live Agent"` |

**Destination queue type** comes from `work_function` on that same receiving case:
`Advanced` and `Basic` are taken as-is. Values that are unmapped omni levels
(`Technical Advanced`, `Product Bulk Basic`) are folded into the matching tier. Values that
are not a tier at all (`Call Assignment`, `UNMAPPED`, `Project`) are reported on a separate
**Other** line and are **not** silently added to either row. `Omni Level Indicator` is used
as a cross-check; it is finer (it also distinguishes `Intermediate`) but mixes skill and
tier, so it is not used as the primary axis.

## 2. Known limitations

These are properties of the export, not of our process, and they apply to any site using the
same source:

1. **No transfer destination field.** See above. Section 1 measures where transferred cases
   *landed*, inferred from the receiving case, not a routing decision recorded at transfer time.
2. **The two legs do not reconcile.** The ceding leg (`case_status = "Closed - Transferred"`,
   or `Case Resolution Category = "Case Transfer"`) and the receiving leg are different rows
   with no key linking them. In our reference week the ceding leg counted 64 cases and the
   receiving leg 139. Both are reported; they are not the same population and the difference
   is not an error.
3. **Voice-to-voice call transfers are a separate population.**
   `Case Resolution Category = "Call Transfer"` (76 cases in the reference week) always
   carries `work_function = Call Assignment`, so it has no Advanced/Basic destination and is
   reported in the annex only.
4. **No transfer reason.** `Transfer Type` and `Transfer Reason` exist in the agent-workitem
   export schema but are entirely unpopulated, and that export has no case number, so it
   cannot be joined to the case data.
5. **No transfer timestamp.** Case-level dates are day-granular; there is no time at which
   the transfer itself occurred.

## 3. Scrubbing methodology and coverage

We are **not** at 100% scrubbing. Coverage is a designed sample, not an ad-hoc one:

- **100% of high-risk transfers are scrubbed.** A transfer is flagged high-risk if any of:
  - `Misrouted Cases = 1`;
  - it shares a parent case with another transfer — i.e. one contact produced more than one
    transfer (*bounce*);
  - the transferred case itself spawned child cases (`Has Child Cases = Yes`);
  - its handle time is at or below the 10th percentile of all transfers that period
    (a transfer that was passed on without being worked).
- **A stratified random sample of the remaining transfers**, stratified by case type, sized
  for a ±5% margin of error at 95% confidence. For a population of ~700 monthly transfers
  that is ~250 cases; for a single week (~140 transfers) it is ~103.
- Each sampled case is marked valid / invalid by a team member against the agreed
  definition, with a free-text reason recorded.

`Invalid rate` is therefore a **sample estimate**, and is reported with its confidence
interval. It is not directly comparable to a site that scrubs 100%.

## 4. Open question

The template refers to an "agreed definition" of a valid transfer that is not included in
the template itself. Our marking follows the definition supplied by the requester; until it
is supplied, the case-level export is produced but the valid/invalid column is left blank.
