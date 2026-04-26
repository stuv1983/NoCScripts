"""
run_dashboard.py — command-line entrypoint.

Runs the full dashboard pipeline without launching the GUI.
Suitable for scheduled jobs, CI pipelines, and server-side automation.

Usage examples
--------------
# Minimal (skip RMM, all dates):
python run_dashboard.py \\
    --input  reports/april_cve.xlsx \\
    --output output/April_Dashboard.xlsx \\
    --skip-rmm

# With RMM and patch match:
python run_dashboard.py \\
    --input   reports/april_cve.xlsx \\
    --rmm     reports/device_inventory.xlsx \\
    --patch   reports/patch_report.csv \\
    --output  output/April_Dashboard.xlsx \\
    --threshold 9.0 \\
    --since   2026-04-01

# With trend comparison:
python run_dashboard.py \\
    --input    reports/april_cve.xlsx \\
    --rmm      reports/device_inventory.xlsx \\
    --output   output/April_Dashboard.xlsx \\
    --previous output/March_Dashboard.xlsx
"""

import argparse
import logging
import sys
from pathlib import Path

from orchestrator import DashboardRequest, run as run_dashboard

# CLI uses the same logging format as the GUI
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(name)s — %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='run_dashboard',
        description='Generate an N-able CVE Dashboard without launching the GUI.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument('--input',     required=True,  metavar='FILE',
                   help='Vulnerability / CVE report (CSV or XLSX)')
    p.add_argument('--output',    required=True,  metavar='FILE',
                   help='Output workbook path (.xlsx)')
    p.add_argument('--rmm',       default=None,   metavar='FILE',
                   help='Device Inventory / RMM report (CSV or XLSX)')
    p.add_argument('--skip-rmm',  action='store_true',
                   help='Skip RMM merge (CVE export already includes device info)')
    p.add_argument('--patch',     default=None,   metavar='FILE',
                   help='Patch report for patch-match analysis (CSV or XLSX)')
    p.add_argument('--previous',  default=None,   metavar='FILE',
                   help='Previous dashboard (.xlsx) for month-over-month trends')
    p.add_argument('--threshold', default=9.0,    type=float, metavar='SCORE',
                   help='Minimum CVE score to include (default: 9.0)')
    p.add_argument('--since',     default=None,   metavar='YYYY-MM-DD',
                   help='Only include detections on or after this date')
    p.add_argument('--all-dates', action='store_true',
                   help='Include all detection dates (overrides --since)')
    p.add_argument('--verbose',   action='store_true',
                   help='Enable DEBUG-level logging')
    return p


def main() -> int:
    parser = build_parser()
    args   = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Validate paths upfront so errors are clear before any processing starts
    if not Path(args.input).exists():
        log.error("Input file not found: %s", args.input)
        return 1
    if args.rmm and not Path(args.rmm).exists():
        log.error("RMM file not found: %s", args.rmm)
        return 1
    if args.patch and not Path(args.patch).exists():
        log.error("Patch file not found: %s", args.patch)
        return 1
    if args.previous and not Path(args.previous).exists():
        log.error("Previous report not found: %s", args.previous)
        return 1
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)

    request = DashboardRequest(
        vuln_path        = args.input,
        output_path      = args.output,
        rmm_path         = args.rmm,
        skip_rmm         = args.skip_rmm,
        patch_path       = args.patch,
        include_patch    = bool(args.patch),
        prev_report_path = args.previous,
        include_trend    = bool(args.previous),
        threshold        = args.threshold,
        cutoff_date      = None if args.all_dates else args.since,
        show_all_dates   = args.all_dates,
    )

    log.info("Starting headless dashboard generation")
    result = run_dashboard(request)

    if result.success:
        log.info("Done: %s", result.output_path)
        if result.trend_summary:
            ts = result.trend_summary
            log.info(
                "Trend: +%d new  -%d resolved  %d persisting",
                ts['new_cve_count'], ts['resolved_cve_count'], ts['persisting_cve_count'],
            )
        for w in result.warnings:
            log.warning(w)
        return 0
    else:
        log.error("Failed: %s", result.message)
        return 1


if __name__ == '__main__':
    sys.exit(main())
