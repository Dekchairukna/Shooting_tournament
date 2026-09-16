# PATCH: Approved score row

- Overview marks a score as APPROVED when either:
  - the existing bypass/approve path is used (`bypass_signed`), or
  - all 3 signature images are present: score recorder, athlete, score/referee judge.
- APPROVED rows are blue across the entire row and the status pill reads `APPROVED`.
- Live `/overview-data` refresh preserves the APPROVED blue state.
- Finished-but-not-approved rows remain green with status `ตีแล้ว`.
- Visible scorecard role label `กรรมการตัดสิน` was renamed to `กรรมการยกคะแนน`; database field names are unchanged for compatibility.
