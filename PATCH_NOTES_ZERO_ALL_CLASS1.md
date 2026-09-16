# Zero-score ranking fix

- Keep the existing Round 1 ranking logic unchanged.
- Special case only when every row has TOTAL = 0: show Class = 1 for every athlete.
- As soon as any athlete has a non-zero total, the normal ranking rules apply again.
