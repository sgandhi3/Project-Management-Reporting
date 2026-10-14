# Report context — facts ADO can't tell us

The weekly report task reads this file (it is passed to the narrative writer as
`reportContext`) and weaves it into the exec summary, PDM/Enrollment/Claims
bullets and footnotes. **Update it whenever a blocker or date changes; delete
lines that are no longer true.** Last updated: 2026-10-13 (from the manual edits
made to the 10/13 deck).

## Why execution is behind (overall bullet)
- Delays in test design, execution, and defect resolution across core domains.
- Active efforts are in flight to de-risk and re-baseline execution timelines.

## PDM
- 169 PDM SIT test cases (Roster) are blocked pending roster ingestion — expected by 10/30. (-> `pdmBlockedNote` footnote **** and PDM bullet)
- Revised PDM SIT target date is in discussion; keep `tbd: true` on SIT_SCHEDULE.PDM in config.js until confirmed (shows "(TBD)" and a sentence in the PDM status box).

## Enrollment
- Missing members in the QA region are required to complete the test cases. Members expected by 10/12; execution expected to wrap by 10/23.

## Claims (EDI workstream)
- Claims SIT is at risk and progressing slowly due to delays in test design and execution.
