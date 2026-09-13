# Auto Court IDs by event lane count

- Creating an event now pre-creates event-scoped Court IDs from `lane_count` (e.g. 8 lanes -> court01..court08).
- Auto-created Court IDs start disabled (`!UNSET!`) and cannot log in until superadmin sets a password.
- The event Athletes page lists every Court ID and lets superadmin set/change each password only; no manual ID creation step.
- Older events are backfilled automatically when superadmin opens the Athletes page.
- Editing lane_count automatically adds missing Court IDs or removes IDs above the new lane count.
- Existing event/court restrictions, backend scorecard guards, and audit-log behavior are preserved.
