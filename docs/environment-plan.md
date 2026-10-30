# Software & Skill Environment Plan

Plan derived from [`description-and-plans.md`](../description-and-plans.md) and the current Python codebase (master, Oct 2026). Planning only; nothing here is implemented yet.

Ground rules: everything runs **manually, on demand** (no CI, no scheduled jobs); new measurement data is **uploaded manually**; all times are **Europe/Budapest** (all sensors are near Budapest).

## 1. Where we are

| Area | Current state | Problem |
|---|---|---|
| Ingestion | `src/data_importer.py` (IN001): `TempLogs*.zip` → one 57 MB `temperature_database.json`, MD5 dedup | Whole file loaded/rewritten on every run; JSON is slow and not diff-friendly; data is gitignored, so not version-controlled as the description wants |
| Gas meter | `loadGasmeterValuesIntoDatabase.py` merges a CSV into the same JSON | Manual, separate entry point |
| Analyses | `detect_heating.py`, `heating_statistics.py`, `generate_calendars.py`, `create_heatmap.py`, `src/temperature_statistics.py` (STAT001/002) | Root-level scripts with hardcoded device names, each re-reads the JSON, outputs chained via files (`heating_cycles.json`) |
| Alignment | 15-min tolerance matching inside STAT002 only | No shared 5-min aligned grid (requirement in §4 of the description) |
| Sensor metadata | Device names baked into code. Naming convention stays as is (e.g. `T1_BE` = bedroom, `T3_Kek` = kitchen) | Device renames/moves (`T5_Lampa`→`T5_KamraLent`, `T7_Pince`/`T7_Kut`/`T7_AUX`) are not modelled |
| Docs/AI context | `specification.md`, `InformationForAI.md`, `.github/copilot-instructions.md` (says timestamps are Excel format; the CSV header even carries the Excel formula `(A2/86400)+25569`, but the stored value is a **Unix epoch in seconds, UTC**, e.g. `1609459200`), per-script `.md` specs | Overlapping and drifting |
| UI | tkinter GUI (`temperature_gui.py`) | Desktop-only, no use in cloud/agent workflows; to be removed |
| Automation | None (no CI, no scheduled runs) | Stays manual by design |

## 2. Target architecture

```
Manual upload (TempLogs*.zip, gas readings as text)
        │
        ▼
 raw/  (immutable zips, content-hashed)         ── private data repo
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
    periods:
      - {start: "2024-10-01 00:00:00", end: null, location: radiator_z2, role: heating_probe}
  - device: T5_Lampa
    periods:
      - {start: "2024-10-01 00:00:00", end: "2025-03-14 18:30:00", location: lamp, role: room}
  ```
  Each period has an explicit `start` and `end` (`null` = still active) with date + `HH:MM:SS`, Europe/Budapest local time.
- **Analyses are plugins** with a feature ID (keeps the existing `(STAT001)`/`[TempCal_Z1]` naming convention). A tiny DAG (Directed Acyclic Graph: each analysis lists the analyses it depends on, `requires=["HEAT001"]`, and they run in dependency order) replaces "please run detect_heating.py first" messages.
- **Thresholds in `config/thresholds.yaml`** (`TEMP_RISE_ABOVE_MIN`, gap minutes, boost detection parameters) so the agent can tune them without code edits.
- **Heating start and boost events are inferred from the measurements**, not from documented schedules (those are not reliably recorded). HEAT001 detects cycle starts from the radiator probes and classifies short, regular runs as boosts vs. demand-driven cycles.
- **GUI**: tkinter GUI is **removed** (`temperature_gui.py`, `GUI_README.md`). Every function it had (device/date selection, plots, STAT002 differences, PNG/CSV export) must be available in the HTML report and a local web dashboard.
- **Raw data location**: a separate **private** `TemperatureMonitoring-data` repo holds raw zips, `readings.parquet` and `gas.csv` (plain git; no Google Drive, no Git LFS). Code repo stays data-free except samples.
- **Time**: everything is stored and used as Europe/Budapest time (tz-aware, so DST is unambiguous). Ingest converts the raw UTC epoch once; no other timezones are handled.

### 2.2 Quality gates

- `pytest` with synthetic fixtures (generated 5-min series with known heating cycles / vent on-off) and golden-file tests per analysis ID. A golden-file test runs an analysis on a fixed input and compares the output with a reviewed, committed expected result file (e.g. `tests/golden/HEAT001.json`); any difference fails the test, and intended changes are accepted by regenerating the file and reviewing the diff.
- `ruff` + `mypy` (type hints already a project rule).
- Run manually (`pytest`, `ruff check`, `mypy`) before merging; no GitHub Actions.

## 3. Automation & Claude skill environment

### 3.1 Repo-level AI context

- **`CLAUDE.md`** replaces `InformationForAI.md` + `copilot-instructions.md`: architecture map, commands, feature-ID convention, "never commit real data to the code repo".
- **`specification.md`** stays the human spec, one section per feature ID; per-script `.md` files fold into it.
- **`prompts.md`** is dropped as methodology documentation (it was only preliminary chat history).
- **SessionStart hook** (`.claude/hooks/session-start.sh`): `pip install -e .[dev]`, clone/pull the data repo, so every cloud session can run analyses immediately.

### 3.2 Claude skills (`.claude/skills/`)

| Skill | Trigger | What it does |
|---|---|---|
| `ingest-logs` | "new temperature logs", a zip uploaded manually | `tempmon ingest` → `tempmon doctor` → summary (new rows per device, gaps, low battery < 2.6 V) → commit to data repo |
| `add-gas-reading` | "gas 12345.678 at 2026-10-20 07:45" | Text reading only, no photos. User gives the value with date, hour and minute (asks for the time if missing), sanity-checks monotonicity and rate vs. degree-days, appends to `gas.csv` |
| `thermal-report` | "how was last month", "report for winter 2025" | Runs analyses for the period, builds HTML report, writes a short narrative (heating hours, gas/degree-day, ventilation off periods, anomalies), publishes as an Artifact |
| `new-analysis` | "add an analysis that …" | Spec-first workflow: assign feature ID → add spec section → synthetic-data test → implement plugin → run on real data → show chart |
| `sensor-change` | "moved T7 to the well at 2026-10-20 14:00", "replaced battery" | Closes the current period (`end`) and opens a new one (`start`) in `sensors.yaml` with date + time, re-runs alignment |

Each skill only calls the `tempmon` CLI; logic lives in Python, not in prompts, so it is testable and reproducible.

### 3.3 Execution

No scheduled workflows. Sensor download (Bluetooth, phone app), upload of new zips and gas readings, ingestion, tests and reports are all run manually on demand, via the skills above or the `tempmon` CLI.

## 4. Roadmap

| Phase | Deliverable | Notes |
|---|---|---|
| 0. Hygiene | `CLAUDE.md`, fix doc errors (timestamp format), remove tkinter GUI, drop `prompts.md`, pyproject deps (drop `zipfile36`, add `typer`, `pyarrow`, `duckdb`) | Small PR |
| 1. Store | `tempmon ingest` → Parquet (Europe/Budapest), one-off migration from `temperature_database.json`, `sensors.yaml` with start/end periods, gas.csv | Verify row counts equal old DB |
| 2. Align | 5-min grid + gap mask, tests | Unblocks all cross-sensor analyses |
| 3. Port analyses | STAT001/002, HEAT001 (heating start + boost inference from data), calendars → plugins; delete root scripts | Golden tests: same cycle counts as old scripts |
| 4. Energy | DD001, GAS001 regression + report | The headline insight from §5 of the description |
| 5. UI + skills | HTML report and local web dashboard covering former GUI functions; `ingest-logs`, `add-gas-reading`, `thermal-report`, `new-analysis`, `sensor-change`, SessionStart hook | Needs phases 1–4 |

## 5. Decisions

1. **Data location**: separate private `TemperatureMonitoring-data` repo, plain git (no Google Drive, no Git LFS).
2. **Data intake**: manual upload of new measurement data; no Drive/CLI cloud access.
3. **GUI**: tkinter removed; all its functionality moves to the HTML report and a local web dashboard.
4. **Timezone**: store and use Europe/Budapest only. Fixes today's `temperature_processor.py:140`, which converts to the naive local time of whatever machine runs it (differs between laptop and cloud).
5. **Automation**: none; no CI, no scheduled workflows, no Claude routines.
