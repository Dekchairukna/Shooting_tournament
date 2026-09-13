# Live + 48 Countries

- Removed the experimental Control dashboard/menu.
- Added public Live results page per event: `/events/<id>/live?round=1`.
- Live page refreshes from the existing overview-data API every 1 second.
- Added flag mapping for the 48 requested participating countries/teams.
- Added Thai canonical country names plus common English aliases.
- Athlete entry now offers the 48-country list and shows flags.
- Court queue now shows country flags.
- Score entry flow remains: court account -> event -> my court queue -> scorecard.
- No database schema change required for country flags; existing `affiliation` is reused.
