"""
diagnostics.py — patch failure classification and remediation guidance.

Classifies each unresolved device-CVE pair into one of five defensible states
based only on data present in the patch report.

Internal cause codes are kept for logic; Excel output uses plain "Patch Evidence Notes".

Author : Stu Villanti <s.villanti@kenstra.com>
"""

from __future__ import annotations
import logging, re
from typing import Optional
import pandas as pd
from data_pipeline import extract_cve_id, get_base_product, normalize_device_name

log = logging.getLogger(__name__)

# ── Display mapping (internal code → plain English for Excel output) ──────────
# Keep internal codes for classification logic; these labels are what
# stakeholders and L1/L2 see. Wording implies the action, not just the state.
DISPLAY_MAP: dict[str, str] = {
    "version_below_fixed": "Patch required",
    "below_baseline":      "Below current product baseline",
    "version_compliant":   "Patched but still detected (rescan required)",
    "detection_mismatch":  "Patched but still detected (rescan required)",  # same L1 action
    "coverage_gap":        "Device missing from patch report",
    "unmanaged_app":       "Product not tracked",
    "patch_installing":    "Patch installing — in progress",
    "patch_pending":       "Patch pending — scheduled, not yet started",
    "patch_missing":       "Patch missing — not yet deployed to device",
    "no_version_data":     "Installed but version unknown",
    "no_fixed_baseline":   "No patch baseline defined",
}

# ── Health score penalties ────────────────────────────────────────────────────
_PENALTIES: dict[str, float] = {
    "version_below_fixed": 2.5,
    "coverage_gap":        2.0,
    "unmanaged_app":       1.5,
    "below_baseline":      1.0,   # CVE-patched but not on current safe release
    "version_compliant":   1.0,
    "detection_mismatch":  1.0,
    "no_fixed_baseline":   0.5,
    "patch_installing":    0.3,
    "patch_pending":       0.5,   # scheduled but not delivered
    "patch_missing":       1.5,   # patch tool knows device needs it but hasn't delivered
    "no_version_data":     0.5,
}

def compute_health_score(root_cause_df: pd.DataFrame, total_pairs: int) -> dict:
    """Patch Reliability Score 0-100. Fewer gaps = higher score."""
    if total_pairs == 0 or root_cause_df.empty:
        return {"score": 100, "grade": "A", "breakdown": {}, "interpretation": "No data"}
    counts = root_cause_df["Patch Evidence Notes"].value_counts().to_dict()
    # Map display labels back to internal codes for penalty lookup
    _rev = {v: k for k, v in DISPLAY_MAP.items()}
    breakdown, total_penalty = {}, 0.0
    for label, count in counts.items():
        cause   = _rev.get(label, label)
        w       = _PENALTIES.get(cause, 1.0)
        penalty = min(round((count / total_pairs) * 100 * w, 1), 40.0)
        breakdown[label] = {"count": count, "weight": w, "penalty": penalty}
        total_penalty   += penalty
    score = max(0, round(100 - total_penalty))
    grade = ("A" if score >= 90 else "B" if score >= 75 else
             "C" if score >= 60 else "D" if score >= 40 else "F")
    interp = {
        "A": "Excellent — patching is well-managed with minimal gaps",
        "B": "Good — minor gaps present, targeted remediation advised",
        "C": "Fair — significant patching issues, prioritise action items below",
        "D": "Poor — systemic patching failures, immediate remediation required",
        "F": "Critical — environment is largely unpatched or unmanaged",
    }[grade]
    log.info("Health score: %d (%s) — %s", score, grade, interp)
    return {"score": score, "grade": grade, "breakdown": breakdown,
            "interpretation": interp, "total_pairs": total_pairs}

# ── Classification rules (internal — not shown in Excel) ─────────────────────
_RULES = [
    # (pmr_substring, resolved_value, vcr_substring, internal_cause)
    ("Not found in patch report",                  None,        None,               "coverage_gap"),
    ("Device in patch report - product not found", None,        None,               "unmanaged_app"),
    (None,                                         "Patch confirmed - pending rescan", None,          None),
    ("Matched - installed",   "Unresolved", "Version compliant",    "version_compliant"),
    ("Matched - installed",   "Unresolved", "Below fixed version",  "version_below_fixed"),
    ("Matched - installed",   "Unresolved", "no fixed baseline",    "no_fixed_baseline"),
    ("Matched - installed",   "Unresolved", None,                   "no_version_data"),
    # Fix: these statuses have defined DISPLAY_MAP entries and _PENALTIES weights —
    # they must map to their cause codes so rows appear in root_cause_df and
    # contribute to the health score instead of being silently dropped.
    ("Matched - installing",  None,         None,                   "patch_installing"),
    ("Matched - pending",     None,         None,                   "patch_pending"),
    ("Matched - missing",     None,         None,                   "patch_missing"),
    ("Matched - failed",      None,         None,                   "patch_missing"),
]

def classify_baseline_root_cause(row) -> Optional[str]:
    """
    Return 'below_baseline' if the row is CVE-patched but below the product
    baseline, or None otherwise.

    This runs AFTER classify_root_cause — it adds a second, independent note
    for rows that are CVE-confirmed but still behind the recommended baseline.
    A device can have both a root_cause (from CVE matching) AND a baseline cause.
    """
    bl_status = str(row.get('Baseline Compliance', '')).strip()
    if bl_status == 'Below baseline':
        return 'below_baseline'
    return None

def classify_root_cause(row) -> Optional[str]:
    """Returns internal cause code or None. No shadow_it guessing — no path data available."""
    pmr = str(row.get("Patch Match Result",          "")).strip()
    res = str(row.get("Patch Evidence Status","Unresolved")).strip()
    vcr = str(row.get("Version Check Result",        "")).strip()
    for pmr_s, res_v, vcr_s, cause in _RULES:
        if ((pmr_s is None or pmr_s.lower() in pmr.lower()) and
            (res_v is None or res == res_v) and
            (vcr_s is None or vcr_s.lower() in vcr.lower())):
            return cause
    return None

# ── Recommendations (config.json remediation_rules + generic fallbacks) ──────
_GENERIC: dict[str, list[str]] = {
    "coverage_gap":        ["Verify RMM agent is active and reporting on affected devices",
                            "Confirm patch report scope includes all sites/clients"],
    "unmanaged_app":       ["Add product to config.json product_map so it is tracked",
                            "Deploy managed installer via RMM to replace untracked version"],
    "version_below_fixed": ["Force-push latest version via RMM software deployment",
                            "Check for failed or deferred update policies on affected devices"],
    "below_baseline":      ["Update product to current baseline version via RMM",
                            "CVE is addressed but device is below the current minimum safe release",
                            "Run version_sync.py to confirm _baseline is current"],
    "no_fixed_baseline":   ["Add minimum fixed version to config.json fixed_version_rules",
                            "Check NVD for published fixed version and update config"],
    "version_compliant":   ["Trigger a fresh N-able vulnerability scan on affected devices",
                            "Verify N-able detection signatures are up to date"],
    "patch_installing":    ["Patch is in progress — re-run report after next RMM sync to confirm completion",
                            "If status remains Installing for >24h, check RMM agent and patch policy for errors"],
    "patch_pending":       ["Patch is scheduled — verify patch policy is active and device is online",
                            "Re-run report after next RMM maintenance window"],
    "patch_missing":       ["Force deploy patch via RMM software deployment",
                            "Check patch policy scope includes this device",
                            "Verify RMM agent is reporting correctly on affected device"],
    "no_version_data":     ["Force RMM agent inventory sync on affected devices",
                            "Reinstall or update RMM agent if version data is consistently missing"],
    "detection_mismatch":  ["Trigger a fresh N-able vulnerability scan on affected devices",
                            "Verify N-able detection signatures are up to date",
                            "Confirm installed patch version actually addresses this CVE"],
}


def compute_recommended_actions(root_cause_df: pd.DataFrame,
                                max_actions: int = 3) -> list[dict]:
    """
    Top N prioritised actions from the evidence data — sorted by device count.
    Returns list of dicts ready for the overview sheet or a stakeholder summary.
    """
    if root_cause_df.empty or 'Patch Evidence Notes' not in root_cause_df.columns:
        return []

    _rev = {v: k for k, v in DISPLAY_MAP.items()}
    _TEMPLATE: dict[str, str] = {
        'version_below_fixed': 'Update {product} on {n} device(s) — installed version is below the required fix',
        'coverage_gap':        'Investigate {n} device(s) missing from patch report — check RMM agent status',
        'version_compliant':   'Trigger re-scan on {n} device(s) — patch installed but CVE still detected',
        'detection_mismatch':  'Trigger re-scan on {n} device(s) — patched but still showing as vulnerable',
        'unmanaged_app':       'Add {product} to patch tracking — {n} device(s) affected and untracked',
        'no_fixed_baseline':   'Define fixed version for {product} in config.json — {n} device(s) unverifiable',
        'no_version_data':     'Fix RMM telemetry on {n} device(s) — patch installed but version not recorded',
    }

    agg: dict[tuple[str, str], set[str]] = {}
    for row in root_cause_df.itertuples(index=False):
        label  = str(getattr(row, 'Patch_Evidence_Notes', ''))
        prod   = str(getattr(row, 'Product', ''))
        device = str(getattr(row, 'Device', ''))
        agg.setdefault((label, prod), set()).add(device)

    actions = []
    for (label, prod), devices in agg.items():
        cause = _rev.get(label, '')
        n     = len(devices)
        bp    = get_base_product(prod).title()
        actions.append({
            'label':   label,
            'cause':   cause,
            'product': bp,
            'count':   n,
            'action':  _TEMPLATE.get(cause, f'{label} on {{n}} device(s)').format(n=n, product=bp),
        })

    return sorted(actions, key=lambda x: -x['count'])[:max_actions]

def get_recommendations(cause: str, product: str,
                        product_rules: Optional[dict] = None) -> list[str]:
    steps: list[str] = []
    if product_rules:
        bp = get_base_product(product).lower()
        steps += product_rules.get(bp, [])
    steps += _GENERIC.get(cause, [])
    seen: set[str] = set()
    return [s for s in steps if not (s in seen or seen.add(s))]  # type: ignore

# ── Main entry point ──────────────────────────────────────────────────────────
def compute_patch_diagnostics(patch_full_df: pd.DataFrame,
                               product_rules: Optional[dict] = None,
                               resolved_pairs: Optional[set] = None) -> dict:
    """
    Classify patch evidence for each unresolved device-CVE pair.

    resolved_pairs: set of (normalised_device, cve_id) tuples already resolved
    by any method — pipeline-confirmed OR manually marked ☑ in product sheets.
    These pairs are excluded from Patch Evidence Notes so a manual resolution
    is not contradicted by a classification.

    Returns:
        patch_lag_df      resolved pairs with days-to-fix
        version_drift_df  products with multiple installed versions
        root_cause_df     per-pair classification (Patch Evidence Notes)
        health_score      environment reliability score 0-100
    """
    df = patch_full_df.copy()
    _e = pd.DataFrame()
    _no_h = {"score": None, "grade": None, "breakdown": {}, "interpretation": "No data"}
    required = {"Name","Vulnerability Name","Patch Match Result","Patch Evidence Status"}
    if not required.issubset(df.columns):
        log.warning("compute_patch_diagnostics: missing columns %s", required - set(df.columns))
        return {"patch_lag_df": _e, "version_drift_df": _e, "root_cause_df": _e, "health_score": _no_h}

    df["_cause"]          = df.apply(classify_root_cause, axis=1)
    df["_baseline_cause"] = df.apply(classify_baseline_root_cause, axis=1)

    # ── Root cause / Patch Evidence Notes table (simplified columns) ──────────
    rows = []
    for row in df[df["_cause"].notna() | df["_baseline_cause"].notna()].itertuples(index=False):
        cause          = getattr(row, "_cause", None)
        baseline_cause = getattr(row, "_baseline_cause", None)
        prod           = str(getattr(row, "Affected_Products", ""))
        device_name    = str(getattr(row, "Name", ""))
        cve_id         = extract_cve_id(str(getattr(row, "Vulnerability_Name", "")))

        # Skip pairs already resolved by any method.
        # resolved_pairs may contain 2-tuples (device, cve) or 3-tuples (device, cve, product).
        if resolved_pairs:
            nk = normalize_device_name(device_name)
            from data_pipeline import _detect_product as _dp_det
            _pk = _dp_det(str(getattr(row, "Affected_Products", "")))
            if (nk, cve_id) in resolved_pairs or (nk, cve_id, _pk) in resolved_pairs:
                continue

        # Primary CVE cause row
        if cause:
            steps = get_recommendations(cause, prod, product_rules)
            rows.append({
                "Device":               device_name,
                "Product":              prod,
                "CVE":                  cve_id,
                "Patch Match Result":   getattr(row, "Patch_Match_Result", ""),
                "Resolved":             getattr(row, "Patch_Evidence_Status", ""),
                "Patch Evidence Notes": DISPLAY_MAP.get(cause, "Unresolved"),
                "Baseline Compliance":  getattr(row, "Baseline_Compliance", ""),
                "Recommended Steps":    "\n".join(f"{i+1}. {s}" for i, s in enumerate(steps)),
                "_cause_internal":      cause,
            })
        elif baseline_cause:
            # No CVE-specific issue but below baseline — emit a standalone baseline row
            steps = get_recommendations(baseline_cause, prod, product_rules)
            rows.append({
                "Device":               device_name,
                "Product":              prod,
                "CVE":                  cve_id,
                "Patch Match Result":   getattr(row, "Patch_Match_Result", ""),
                "Resolved":             getattr(row, "Patch_Evidence_Status", ""),
                "Patch Evidence Notes": DISPLAY_MAP.get(baseline_cause, ""),
                "Baseline Compliance":  getattr(row, "Baseline_Compliance", ""),
                "Recommended Steps":    "\n".join(f"{i+1}. {s}" for i, s in enumerate(steps)),
                "_cause_internal":      baseline_cause,
            })
    root_cause_df = (pd.DataFrame(rows).sort_values("Patch Evidence Notes", ignore_index=True)
                     if rows else _e)

    health = compute_health_score(root_cause_df, total_pairs=len(df))

    # ── Patch Evidence Notes summary (for overview sheet) ─────────────────────
    if not root_cause_df.empty:
        summary = root_cause_df["Patch Evidence Notes"].value_counts().to_dict()
        log.info("Patch evidence summary: %s", summary)
    else:
        summary = {}

    # ── Patch lag (resolved pairs) ────────────────────────────────────────────
    lag_rows = []
    if "Patch Install Date" in df.columns and "First detected" in df.columns:
        for row in df[df["Patch Evidence Status"] == "Patch confirmed - pending rescan"].itertuples(index=False):
            idt = pd.to_datetime(getattr(row, "Patch_Install_Date", None), errors="coerce")
            fdt = pd.to_datetime(getattr(row, "First_detected", None),     errors="coerce")
            if pd.isna(idt) or pd.isna(fdt): continue
            lag_rows.append({
                "Device":          getattr(row, "Name", ""),
                "CVE":             extract_cve_id(str(getattr(row, "Vulnerability_Name", ""))),
                "Product":         getattr(row, "Affected_Products", ""),
                "First Detected":  fdt.date(),
                "Patch Installed": idt.date(),
                "Lag (days)":      (idt - fdt).days,
            })
    patch_lag_df = (pd.DataFrame(lag_rows).sort_values("Lag (days)", ascending=False, ignore_index=True)
                    if lag_rows else _e)

    # ── Version drift ─────────────────────────────────────────────────────────
    # Groups by base product AND architecture (x64/x86) so a fleet with mixed
    # 32-bit and 64-bit Chrome/Firefox installs shows separate drift rows for
    # each variant rather than merging them into one misleading bucket.
    _ARCH_RE = re.compile(r'\((x64|x86|32[\-\s]?bit|64[\-\s]?bit)\)', re.IGNORECASE)

    def _arch_suffix(text: str) -> str:
        """Return ' (x64)' / ' (x86)' / '' from a patch name or product string."""
        m = _ARCH_RE.search(str(text))
        if not m:
            return ''
        a = m.group(1).lower()
        return ' (x86)' if ('x86' in a or '32' in a) else ' (x64)'

    def _drift_key(row) -> str:
        """
        Stable grouping key for version drift: base product + arch suffix.

        Arch is preferred from 'Matched Patch' (e.g. 'Chrome (x64) 147.0…')
        because N-able's CVE export often omits it from 'Affected Products'
        (e.g. 'Google Chrome' with no arch tag).  Falls back to 'Affected
        Products' arch when the patch name is absent or has no arch tag.
        """
        bp   = get_base_product(str(row.get("Affected Products", "")))
        arch = _arch_suffix(str(row.get("Matched Patch", "")))
        if not arch:
            arch = _arch_suffix(str(row.get("Affected Products", "")))
        return bp + arch

    drift_rows = []
    no_version_data_products = []
    if "Matched Patch Version" in df.columns:
        df["_bp"]       = df["Affected Products"].apply(get_base_product)
        df["_drift_key"] = df.apply(_drift_key, axis=1)

        for dk, grp in df.groupby("_drift_key"):
            vers = (grp["Matched Patch Version"].dropna().astype(str).str.strip()
                    .loc[lambda s: s.str.len() > 0].unique().tolist())
            if len(set(vers)) < 2:
                if len(vers) == 0:
                    # Only log the base product name (no arch suffix) for the
                    # "no version data" note — avoids duplicating e.g.
                    # "Google Chrome (x64)" and "Google Chrome (x86)" separately
                    # when both have no data.
                    bp = str(grp["_bp"].iloc[0]) if not grp.empty else str(dk)
                    if bp not in no_version_data_products:
                        no_version_data_products.append(bp)
                continue
            drift_rows.append({
                "Product":           dk,           # e.g. "Google Chrome (x64)"
                "Distinct Versions": len(set(vers)),
                "Versions Seen":     ", ".join(sorted(set(vers))),
                "Device Count":      grp["Name"].nunique(),
            })
    version_drift_df = (pd.DataFrame(drift_rows).sort_values(
        "Distinct Versions", ascending=False, ignore_index=True) if drift_rows else _e)

    if no_version_data_products:
        log.info("Version drift: no version data for %s — not in patch tool scope",
                 ", ".join(no_version_data_products))

    return {
        "patch_lag_df":            patch_lag_df,
        "version_drift_df":        version_drift_df,
        "version_drift_no_data":   no_version_data_products,
        "root_cause_df":           root_cause_df,
        "health_score":            health,
        "evidence_summary":        summary,
        "recommended_actions":     compute_recommended_actions(root_cause_df),
    }