# Event-scoped Court IDs

## What changed
- Court account creation moved from the main Dashboard to each event's **Athletes** page.
- An event may use no Court IDs at all; normal admin workflow remains unchanged.
- Court accounts are now scoped to both **event + court number**.
- Staff still log in with the simple public ID `court01`, `court02`, ...
- Internally, usernames are unique per event (for example `event12_court01`) so different events do not overwrite one another.
- Court login redirects directly to that event's Round 1 overview.
- Court users cannot open another event's overview or scorecard and cannot autosave scores for another court.
- Round 1 and Round 2 filtering both use the effective court assignment for that round.
- Generic User Management no longer creates Court accounts; Court IDs are managed only inside each event.
- Court IDs can be removed from the event's Athletes page.
- When an event is deleted, its Court accounts are deleted too.

## Login ambiguity safeguard
If the same public ID (e.g. `court01`) exists in multiple events **and** those accounts use the same password, login is intentionally rejected as ambiguous. Use different passwords for Court IDs across concurrently stored events.

## Existing score audit features retained
- Backend court guard on scorecard/autosave.
- Score edit history.
- Edit-after-signature requires editor signature and records the change.
