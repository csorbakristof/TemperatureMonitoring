# Software & Skill Environment Plan

Plan derived from [`description-and-plans.md`](../description-and-plans.md) and the current Python codebase (master, Oct 2026). Planning only; nothing here is implemented yet.

## 1. Where we are

| Area | Current state | Problem |
|---|---|---|
| Ingestion | `src/data_importer.py` (IN001): `TempLogs*.zip` → one 57 MB `temperature_database.json`, MD5 dedup | Whole file loaded/rewritten on every run; JSON is slow and not diff-friendly; data is gitignored, so not version-controlled as the description wants |
| Gas meter | `loadGasmeterValuesIntoDatabase.py` merges a CSV into the same JSON | Manual, separate entry point |
| Analyses | `detect_heating.py`, `heating_statistics.py`, `generate_calendars.py`, `create_heatmap.py`, `src/temperature_statistics.py` (STAT001/002) | Root-level scripts with hardcoded device names, each re-reads the JSON, outputs chained via files (`heating_cycles.json`) |
| Alignment | 15-min tolerance matching inside STAT002 only | No shared 5-min aligned grid (requirement in §4 of the description) |
| Sensor metadata | Device names baked into code; spec has wrong meanings (`T1_BE` = "bedroom", `T3_Kek` = "kitchen") | Device renames/moves (`T5_Lampa`→`T5_KamraLent`, `T7_Pince`/`T7_Kut`/`T7_AUX`) are not modelled |
| Docs/AI context | `specification.md`, `InformationForAI.md`, `.github/copilot-instructions.md` (says timestamps are Excel format, wrong), per-script `.md` specs | Overlapping and drifting |
| UI | tkinter GUI (`temperature_gui.py`) | Desktop-only, no use in cloud/agent workflows |
| Automation | None (no CI, no scheduled runs) | Everything manual |

## 2. Target architecture

```
Google Drive (TempLogs*.zip, gas photos/CSV)
        │  rclone / Drive API
        ▼
 raw/  (immutable zips, content-hashed)         ── private data repo or Drive
        │  tempmon ingest
        ▼
 readings.parquet  (long: device, ts, temp, hum, vbat)   ── version-controlled
 gas.csv           (ts, meter_m3, source)                ── hand-editable
        │  tempmon align
        ▼
 grid_5min.parquet (wide: one column per *location*, interpolated, gap-limited)
        │  tempmon analyze <id>
        ▼
 results/<ID>.parquet|json  ──►  tempmon report  ──►  HTML report / PNG / CSV
```

### 2.1 Package layout (`src/tempmon/`, installable, single CLI)

```
src/tempmon/
  cli.py            # typer: ingest | gas | align | analyze | report | doctor
  config.py         # loads sensors.yaml + thresholds.yaml (pydantic)
  io/raw_zip.py     # ZIP→CSV parsing (reuse temperature_processor logic)
  io/store.py       # parquet read/append/dedup (DuckDB for queries)
  io/gas.py
  align.py          # 5-min grid, interpolation (limit=2 samples), gap mask
  analyses/
    base.py         # Analysis protocol: id, requires, run(ctx) -> Result
    gaps.py         # STAT001 coverage + battery health
    ventilation.py  # STAT002 supply-vent (T1_BE) vs room/outdoor
    heating.py      # HEAT001 cycle detection Z1/Z2 (from detect_heating.py)
    degree_days.py  # DD001 daily ΣΔT (inside − outside)
    gas_model.py    # GAS001 regression: gas ~ a·ΣΔT + b·cycles(Z1,Z2)
    anomalies.py    # ANOM001 sensor dead/stuck/out-of-range
  viz/
    calendars.py    # from generate_calendars.py / create_heatmap.py
    charts.py
  report/html.py    # single self-contained HTML (plotly) + summary.md
```

Key decisions:

- **Storage: Parquet + DuckDB** instead of one JSON. Append-only, ~10× smaller, queryable with SQL, and small enough (a few MB/year) to commit. Dedup key = `(device, ts)` instead of MD5 per row.
- **Sensor registry `config/sensors.yaml`** maps physical device → location → role over time intervals, so analyses ask for `role: outdoor` or `location: radiator_z2`, never `T2_Terasz`:
  ```yaml
  - device: T6_Z2
    periods: [{from: 2024-10-01, location: radiator_z2, role: heating_probe}]
  - device: T1_BE
    periods: [{from: 2024-10-01, location: kek_supply_vent, role: vent_supply}]
  ```
- **Analyses are plugins** with a feature ID (keeps the existing `(STAT001)`/`[TempCal_Z1]` naming convention). A tiny DAG (`requires=["HEAT001"]`) replaces "please run detect_heating.py first" messages.
- **Thresholds in `config/thresholds.yaml`** (`TEMP_RISE_ABOVE_MIN`, gap minutes, Z1 boost schedule) so the agent can tune them without code edits.
- **Z1 scheduled boost** (2×20 min/day) is modelled explicitly so HEAT001 can separate forced runs from demand-driven cycles.
- **GUI**: keep tkinter as-is (frozen), new viewing goes through the HTML report. Revisit only if needed.
- **Raw data location**: raw zips stay in Google Drive (source of truth) and are mirrored into a **private** `TemperatureMonitoring-data` repo (or Git LFS) together with `readings.parquet` and `gas.csv`; code repo stays data-free except samples.

### 2.2 Quality gates

- `pytest` with synthetic fixtures (generated 5-min series with known heating cycles / vent on-off) and golden-file tests per analysis ID.
- `ruff` + `mypy` (type hints already a project rule).
- GitHub Actions: lint + tests on every PR.

## 3. Automation & Claude skill environment

### 3.1 Repo-level AI context

- **`CLAUDE.md`** replaces `InformationForAI.md` + `copilot-instructions.md`: architecture map, commands, feature-ID convention, work-phase prefixes from `prompts.md` (`Q:`, `DOC:`, `TEST:`, `REFACTOR`), "never commit real data to the code repo".
- **`specification.md`** stays the human spec, one section per feature ID; per-script `.md` files fold into it.
- **SessionStart hook** (`.claude/hooks/session-start.sh`): `pip install -e .[dev]`, clone/pull the data repo, so every cloud session can run analyses immediately.

### 3.2 Claude skills (`.claude/skills/`)

| Skill | Trigger | What it does |
|---|---|---|
| `ingest-logs` | "new temperature logs", a zip uploaded/added to Drive | `tempmon ingest` → `tempmon doctor` → summary (new rows per device, gaps, low battery < 2.6 V) → commit to data repo |
| `add-gas-reading` | "gas 12345.678", or a **photo of the meter** | Reads value (Claude vision replaces the old C# OCR prototype), sanity-checks monotonicity and rate vs. degree-days, appends to `gas.csv` |
| `thermal-report` | "how was last month", "report for winter 2025" | Runs analyses for the period, builds HTML report, writes a short narrative (heating hours, gas/degree-day, ventilation off periods, anomalies), publishes as an Artifact |
| `new-analysis` | "add an analysis that …" | Spec-first workflow: assign feature ID → add spec section → synthetic-data test → implement plugin → run on real data → show chart. Encodes the methodology in `prompts.md` |
| `sensor-change` | "moved T7 to the well", "replaced battery" | Updates `sensors.yaml` periods with dates, re-runs alignment |

Each skill only calls the `tempmon` CLI; logic lives in Python, not in prompts, so it is testable and reproducible.

### 3.3 Scheduled workflows

| When | Where | Job |
|---|---|---|
| Daily | GitHub Actions cron | rclone-sync Drive folder → if new zip: ingest, commit parquet to data repo |
| Monthly (heating season: weekly) | Claude routine (`create_trigger`) | Run `thermal-report` for the period, post summary + Artifact link in this project, flag anomalies (dead sensor, battery, heating cycles without gas increase) |
| Every PR | GitHub Actions | ruff, mypy, pytest |

Download from the sensors themselves stays manual (Bluetooth, phone app); the plan automates everything after the zip lands in Drive.

## 4. Roadmap

| Phase | Deliverable | Notes |
|---|---|---|
| 0. Hygiene | `CLAUDE.md`, fix doc errors, CI (tests+ruff), pyproject deps (drop `zipfile36`, add `typer`, `pyarrow`, `duckdb`) | Small PR |
| 1. Store | `tempmon ingest` → Parquet, one-off migration from `temperature_database.json`, `sensors.yaml`, gas.csv | Verify row counts equal old DB |
| 2. Align | 5-min grid + gap mask, tests | Unblocks all cross-sensor analyses |
| 3. Port analyses | STAT001/002, HEAT001, calendars → plugins; delete root scripts | Golden tests: same cycle counts as old scripts |
| 4. Energy | DD001, GAS001 regression + report | The headline insight from §5 of the description |
| 5. Skills | `ingest-logs`, `add-gas-reading`, `thermal-report`, SessionStart hook | Needs phases 1–4 |
| 6. Automation | Drive sync Action, monthly Claude routine, `new-analysis`, `sensor-change` skills | Needs Drive credentials as repo secret |

## 5. Open decisions (defaults chosen)

1. **Data location**: private `TemperatureMonitoring-data` repo *(default)* vs. Git LFS in this repo vs. Drive only.
2. **Drive access from CI**: rclone with a Google service account shared on the logs folder *(default)* vs. manual upload to the project.
3. **GUI**: freeze tkinter *(default)* vs. replace with a local web dashboard (e.g. Streamlit/Panel).
4. **Timezone**: store UTC, render Europe/Budapest *(default)*; today `temperature_processor.py:140` converts to naive local time of whatever machine runs it (differs between laptop and cloud).
