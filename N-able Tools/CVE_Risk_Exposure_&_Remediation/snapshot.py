"""
snapshot.py — structured snapshot storage for historical trend tracking.

Directory layout
----------------
<output_dir>/
    snapshots/
        index.json           master index of all runs (appended on each save)
        2026-04.json         monthly aggregate (latest run per month)
        2026-04-25T12-30.json  full per-run record

Design goals
------------
- Accumulate history without requiring the previous Excel workbook
- Allow trend reconstruction from snapshots alone
- Stay readable (JSON, not binary) and queryable without a database
- Non-fatal: snapshot failures are logged as warnings, never crash the run
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_SNAPSHOT_DIR = 'snapshots'


def _snap_dir(output_path: str) -> Path:
    return Path(output_path).parent / _SNAPSHOT_DIR


def save(output_path: str,
         customer: str,
         threshold: float,
         unique_cves: int,
         unique_devices: int,
         trend_metrics: Optional[dict] = None,
         root_cause_summary: Optional[dict] = None) -> None:
    """
    Persist a compact snapshot of this run.  Writes three files:
      - per-run timestamped JSON
      - monthly aggregate JSON (overwritten with latest run each month)
      - index.json (appended)

    All writes are wrapped in try/except — snapshot failure is non-fatal.
    """
    try:
        snap_dir = _snap_dir(output_path)
        snap_dir.mkdir(parents=True, exist_ok=True)

        now   = datetime.now()
        stamp = now.strftime('%Y-%m-%dT%H-%M')
        month = now.strftime('%Y-%m')

        record: dict = {
            'run_date':          now.isoformat(timespec='seconds'),
            'customer':          customer,
            'output_file':       Path(output_path).name,
            'threshold':         threshold,
            'unique_cves':       unique_cves,
            'unique_devices':    unique_devices,
            'trend_metrics':     trend_metrics,
            'root_cause_summary':root_cause_summary,
        }

        # 1. Per-run file
        run_path = snap_dir / f'{stamp}.json'
        _write_json(run_path, record)
        log.info("Snapshot saved: %s", run_path)

        # 2. Monthly aggregate (latest run per month — overwrites)
        monthly_path = snap_dir / f'{month}.json'
        _write_json(monthly_path, record)

        # 3. Index — append this run's summary
        index_path = snap_dir / 'index.json'
        index = _read_json(index_path) if index_path.exists() else []
        if not isinstance(index, list):
            index = []
        index.append({
            'run_date':       record['run_date'],
            'customer':       customer,
            'unique_cves':    unique_cves,
            'unique_devices': unique_devices,
            'file':           run_path.name,
        })
        _write_json(index_path, index)

    except Exception as exc:
        log.warning("Could not save snapshot: %s", exc)


def load_history(output_path: str, months: int = 12) -> list[dict]:
    """
    Load up to `months` monthly snapshots from the snapshots directory.
    Returns a list of records sorted oldest-first, suitable for trend charts.
    Returns empty list if no snapshots exist.
    """
    snap_dir = _snap_dir(output_path)
    if not snap_dir.exists():
        return []

    records: list[dict] = []
    now = datetime.now()
    for i in range(months - 1, -1, -1):
        # Walk backwards month by month
        yr  = now.year  - (now.month - 1 - i < 0)  # type: ignore[operator]
        mo  = (now.month - 1 - i) % 12 + 1
        key = f'{yr:04d}-{mo:02d}'
        p   = snap_dir / f'{key}.json'
        if p.exists():
            rec = _read_json(p)
            if isinstance(rec, dict):
                records.append(rec)

    return records


def get_root_cause_trend(output_path: str, months: int = 6) -> dict[str, list]:
    """
    Return a dict of root_cause → list of monthly counts for charting.
    Useful for answering "are mismatches increasing or decreasing?"
    """
    history = load_history(output_path, months=months)
    result: dict[str, list] = {}
    for rec in history:
        rcs = rec.get('root_cause_summary') or {}
        for cause, count in rcs.items():
            if cause not in result:
                result[cause] = []
            result[cause].append({'month': rec.get('run_date', '')[:7], 'count': count})
    return result


# ==============================================================================
# HELPERS
# ==============================================================================

def _write_json(path: Path, data) -> None:
    with open(path, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, indent=2, default=str)


def _read_json(path: Path):
    with open(path, 'r', encoding='utf-8') as fh:
        return json.load(fh)
