# Import DEMO events from local SQLite to PostgreSQL

This package is for the current Petanque Shooting project.

It imports only the two DEMO events from the local SQLite database:

- `DEMO • FINISHED • 48 NATIONS • PETANQUE SHOOTING 2026`
- `DEMO • NOT STARTED • 48 NATIONS • PETANQUE SHOOTING 2026`

It does **not** overwrite PostgreSQL IDs. New IDs are created and all athlete/bracket references are remapped.

## Recommended way with Railway

Place `import_demo_to_postgres.py` in the project root next to `app.py`.

First test what will be imported:

```bash
python3 import_demo_to_postgres.py --sqlite instance/shooting.db --dry-run
```

Then, if your Railway project exposes `DATABASE_PUBLIC_URL` to `railway run`:

```bash
railway login
railway link
railway run python3 import_demo_to_postgres.py   --sqlite instance/shooting.db   --use-public-url
```

If the same demo event names already exist and you intentionally want to replace them:

```bash
railway run python3 import_demo_to_postgres.py   --sqlite instance/shooting.db   --use-public-url   --replace
```

## Alternative: explicit PostgreSQL public URL

Avoid pasting the password into chat. Set it only in your terminal:

```bash
export TARGET_DATABASE_URL='postgresql://USER:PASSWORD@HOST:PORT/DBNAME'
python3 import_demo_to_postgres.py --sqlite instance/shooting.db
unset TARGET_DATABASE_URL
```

## What is copied

- Event
- Athletes / countries
- Score entries
- Score signatures
- Tie-break entries
- Score edit logs
- Bracket matches
- Results-approved settings

Local user IDs are intentionally **not** copied, because user IDs on the website may be different.
