# Banner Capability Map, runbook

What it answers: **where your institution already owns Banner functionality and does
not use it.** It turns "we think Banner does this" into evidence, so you can answer a
purchase request for a third-party product Banner already covers.

Everything is read-only. Every query is a `SELECT` against your own Banner database.
The point is that a person runs the query, from your own read-only account, and the
result stays on your machine.

---

## Before you run

1. **Network / VPN to the database.** If you cannot reach the host you will see
   `ORA-12545: target host or object does not exist`.
2. **`.env`** in the folder you run from, filled with your read-only connection
   (copy `.env.example` first):
   ```
   BANNER_HOST=your-banner-db-host
   BANNER_PORT=1521
   BANNER_SERVICE=YOUR_SERVICE_NAME
   BANNER_USER=your_readonly_user
   BANNER_PASS=your_password
   ```
   Optional: `ORACLE_CLIENT_DIR=` only if the tool cannot find your Oracle Instant
   Client on its own (the same client SQL Developer uses).
3. **Packages** (once):
   ```
   pip install oracledb python-dotenv openpyxl
   ```

---

## The command

```
python bcm_report.py --out capability_map.xlsx --html capability_map.html
```

- `--out`   the Excel file (defaults to `capability_map.xlsx` if omitted).
- `--html`  the interactive one-page presentation (optional, but this is the one you show).

It runs in under a minute and prints one line per capability with its verdict.

---

## What it produces

**Excel** (`capability_map.xlsx`):
- **Capability Map**: one row per capability, colored by verdict.
- **Access by Group**: which security group or class opens each capability, how many
  screens it reaches, how many of those it can **change**, and how many people hold it.
- **Evidence**: the actual screens and tables measured.
- **Method**: how each row is measured, the verdicts, and the one caveat.
- **Since last run**: appears only once a previous run exists to compare against.

**HTML** (`capability_map.html`): an interactive map. Click a capability and the panel
shows its screens, its tables (with the canonical judge table marked "judges this"),
record and user counts, the last activity date, and the groups that open it. A
**"How we got here"** button opens the method as a plain diagram (the room / doors /
keys model). It is self-contained: no external references, opens straight from the file.

At the top, **one chart**: every capability plotted as people-who-can-reach-it against
records-held. The strip along the bottom is the argument in a picture. Click a dot to
open that capability. Each card also owns a link (`capability_map.html#fixed-assets`),
so one finding can be sent on its own.

---

## The group layer, and where it stops

For each capability the report shows **how the access is wired**: through a security
**group**, a security **class**, or pinned straight onto a person with no group
governing it. Banner names groups after the job ("Admissions Manager"), which is what
makes this readable to a director rather than to a DBA.

It reports a **headcount and never a name.** Say that out loud when you present it. The
group is the unit you can actually govern; a per-person listing is a different review
with different rules and different approvals. Presenting it as a deliberate boundary
rather than a limitation is the difference between a governance tool and a witch hunt.

**Access debt** is flagged when a capability holds no live data *and* a group still
grants the right to change it. Nothing is misconfigured. It is standing permission on a
room nobody enters, and it is usually the cheapest cleanup available, because removing
it takes nothing away from anyone who is working.

---

## History: what moved

Banner's security tables carry **no history.** A revoked grant is deleted, not
end-dated. "Who could reach this last March" cannot be reconstructed after the fact, by
any query, ever. So the tool keeps its own: every run appends a small JSON snapshot to
`history/`.

```
python bcm_report.py --runs                           # runs stored so far
python bcm_report.py --since 2026-03-02 --out ...     # today, compared to that date
python bcm_report.py --compare 2026-03-02 2026-08-05  # two stored runs, NO database
python bcm_report.py --snapshot-only                  # measure and store, write nothing
```

Once a second run exists the report opens with a **what moved since** band, each card
carries its own changes, and the workbook gains a *Since last run* sheet. `--compare`
reads two files and never connects, so it runs off-network, in a meeting, in a second.

- **The first run shows nothing, and that is correct.** There is nothing to compare
  against yet. The value arrives on the second run, which is exactly why the saving has
  to start before anyone asks for it.
- **Never prune `history/`.** It is the one thing here that cannot be regenerated.
- If you schedule it, weekly is a good default: a single missed run then costs a week,
  not a quarter. On an ephemeral worker, point the history at durable storage
  (subclass `HistoryStore`, three methods), or the feature dies silently.

---

## The three questions behind every verdict

1. **Do the screens exist?** Banner forms shipped, from `GUBOBJS` (object type FORM).
   Proves the capability is installed.
2. **Can anyone reach them?** Distinct users with access, from the security view
   `GUVUACC`. The size of the open door, not the traffic.
3. **Is there data, and recent?** A real `COUNT(*)` and the newest activity date on the
   capability's one canonical table. This is the judge.

`Forms exist + people hold the keys + zero data` = owned, wired up, never used.

---

## The verdicts

| Verdict | What it means |
|---|---|
| **Owned, not used** | Screens ship and people hold access, but the table Banner writes to for this holds no rows. A signal worth confirming with the product owner, not a conclusion. |
| **Custom-built** | The native table is empty, but a local custom table quietly does the job. |
| **Abandoned** | Data exists, but the newest activity is years old. |
| **In use** | Recent activity. A department wanting a tool for this already has a working one. |
| **Not installed** | No Banner table for this capability is present. |
| **Expected absent** | Empty on purpose (for example, Housing at a college with no residence halls). Not a gap. |
| **Unable to verify** | The table exists but could not be read (timeout or privilege). NOT a use signal in either direction. Re-run or check the grant. |

---

## Verify before you present

Run the command and read the console:

1. **Nothing lands on "Unable to verify".** If it does, that is a timeout or a
   privilege problem, not a finding. Re-run.
2. **Spot-check a last-activity date** against what you already know. The tool discovers
   the activity-date column from the catalog, so the date should match reality.
3. **The counts that grow are the "In use" ones.** The idle ones stay at zero. That is
   how you know it is live data, not a stale snapshot.

---

## How to present it honestly

- **Lead with content-based findings, not empty tables.** The strongest evidence is not
  "this table is empty", it is "the only thing in this table is the sample the vendor
  ships". Example: a fresh Banner Workflow install carries one process named
  "system verification". If that is the only process defined, the engine was never put
  to work. That is hard to argue with. An empty table alone is weaker.
- **Present a plain empty table as a candidate, never a verdict.** Say "Banner ships
  this and we see no current data; worth a conversation with the product owner before we
  buy", not "nobody uses this".
- **Watch for a sibling table with data.** A drill-down table can be busy while the
  canonical judge table is empty (items get tagged but never capitalized, events get
  rooms but no events, people get certifications but no skills records). Pin the judge
  table, and own the nuance out loud before someone points at the sibling.
- **The one caveat, always**: database evidence shows technical artifacts and rows
  visible to a read-only account. It does NOT prove license ownership, deployment,
  adoption, usability, or functional fit. "Empty" is not "unlicensed"; a recent activity
  date is not "human use" (a load job can set it). Validate entitlement and workflow with
  the Banner product owner before any procurement decision.

---

## Make it yours

The catalogue of capabilities in the script is a starting set of the ones institutions
most often license and under-use. Add your own: each entry is a form prefix (or a
description match) and a canonical table. The verdicts are computed from **your** data,
so they are yours alone. Nothing about your institution ships with this tool.
