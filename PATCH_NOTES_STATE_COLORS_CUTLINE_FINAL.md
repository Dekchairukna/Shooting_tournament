# State Colors + Cut Line cleanup

- Overview row background now represents only shooting state:
  - waiting = grey
  - active = yellow
  - finished = light green
- Removed qualification/progression classes from row class names so old R2/qualified/eliminated CSS cannot override shooting-state colours.
- Status column now prioritizes and displays รอตี / กำลังตี / ตีแล้ว.
- R2 qualification no longer colours the athlete row.
- Shoot-off is shown as a separate action while preserving the shooting-state badge.
- Cut lines are generated only after every relevant athlete in the whole event has finished the qualification round.
- Round 2 direct-qualified placeholders remain neutral.
