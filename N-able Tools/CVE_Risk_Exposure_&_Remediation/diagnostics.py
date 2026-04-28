"""
diagnostics.py — patch failure classification and remediation guidance.

Classifies each unresolved device-CVE pair into one of five defensible states
based only on data that is actually present in the patch report.  No guessing.

Internal cause codes are kept for logic; Excel output uses plain "Patch Evidence Notes".
"""

from __future__ import annotations
import logging, re
from typing import Optional
import pandas as pd
from data_pipeline import extract_cve_id, get_base_product

log = logging.getLogger(__name__)

# ── Display mapping (internal code → plain English for Excel output) ──────────
# Keep internal codes for classification logic; these labels are what
# stakeholders and L1/L2 see. Wording implies the action, not just the state.
DISPLAY_MAP: dict[str, str] = {
    "version_below_fixed": "Patch required",
    "version_compliant":   "Patched but still detected (rescan required)",
    "detection_mismatch":  "Patched but still detected (rescan required)",  # same L1 action
    "coverage_gap":        "Device missing from patch report",
    "unmanaged_app":       "Product not tracked",
    "no_version_data":     "Installed but version unknown",
    "no_fixed_baseline":   "No patch baseline defined",
}

# ── Health score penalties ────────────────────────────────────────────────────
_PENALTIES: dict[str, float] = {
    "version_below_fixed": 2.5,
    "coverage_gap":        2.0,
    "unmanaged_app":       1.5,
    "version_compliant":   1.0,
    "detection_mismatch":  1.0,   # scanner lying or incorrect fix — not healthy
    "no_fixed_baseline":   0.5,
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
    (None,                                         "Resolved",  None,               None),
    ("Matched - installed",   "Unresolved", "Version compliant",    "version_compliant"),
    ("Matched - installed",   "Unresolved", "Below fixed version",  "version_below_fixed"),
    ("Matched - installed",   "Unresolved", "no fixed baseline",    "no_fixed_baseline"),
    ("Matched - installed",   "Unresolved", None,                   "no_version_data"),
    ("Matched - installing",  None,         None,                   None),
    ("Matched - pending",     None,         None,                   None),
    ("Matched - missing",     None,         None,                   None),
    ("Matched - failed",      None,         None,                   None),
]

def classify_root_cause(row) -> Optional[str]:
    """Returns internal cause code or None. No shadow_it guessing — no path data available."""
    pmr = str(row.get("Patch Match Result",          "")).strip()
    res = str(row.get("Resolved (from Patch Report)","Unresolved")).strip()
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
    "no_fixed_baseline":   ["Add minimum fixed version to config.json fixed_version_rules",
                            "Check NVD for published fixed version and update config"],
    "version_compliant":   ["Trigger a fresh N-able vulnerability scan on affected devices",
                            "Verify N-able detection signatures are up to date"],
    "no_version_data":     ["Force RMM agent inventory sync on affected devices",
                            "Reinstall or update RMM agent if version data is consistently missing"],
    "detection_mismatch":  ["Trigger a fresh N-able vulnerability scan on affected devices",
                            "Verify N-able detection signatures are up to date",
                            "Confirm installed patch version actually addresses this CVE"],
}



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
                               product_rules: Optional[dict] = None) -> dict:
    """
    Classify patch evidence for each unresolved device-CVE pair.

    Returns:
        patch_lag_df      resolved pairs with days-to-fix
        version_drift_df  products with multiple installed versions
        root_cause_df     per-pair classification (Patch Evidence Notes)
        health_score      environment reliability score 0-100
    """
    df = patch_full_df.copy()
    _e = pd.DataFrame()
    _no_h = {"score": None, "grade": None, "breakdown": {}, "interpretation": "No data"}
    required = {"Name","Vulnerability Name","Patch Match Result","Resolved (from Patch Report)"}
    if not required.issubset(df.columns):
        log.warning("compute_patch_diagnostics: missing columns %s", required - set(df.columns))
        return {"patch_lag_df": _e, "version_drift_df": _e, "root_cause_df": _e, "health_score": _no_h}

    df["_cause"] = df.apply(classify_root_cause, axis=1)

    # ── Root cause / Patch Evidence Notes table (simplified columns) ──────────
    rows = []
    for _, row in df[df["_cause"].notna()].iterrows():
        cause = row["_cause"]
        prod  = str(row.get("Affected Products", ""))
        steps = get_recommendations(cause, prod, product_rules)
        rows.append({
            "Device":               row.get("Name", ""),
            "Product":              prod,
            "CVE":                  extract_cve_id(str(row.get("Vulnerability Name", ""))),
            "Patch Match Result":   row.get("Patch Match Result", ""),
            "Resolved":             row.get("Resolved (from Patch Report)", ""),
            "Patch Evidence Notes": DISPLAY_MAP.get(cause, "Unresolved"),
            "Recommended Steps":    "\n".join(f"{i+1}. {s}" for i, s in enumerate(steps)),
            "_cause_internal":      cause,   # kept for health score, not written to Excel
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
        for _, row in df[df["Resolved (from Patch Report)"] == "Resolved"].iterrows():
            idt = pd.to_datetime(row.get("Patch Install Date"), errors="coerce")
            fdt = pd.to_datetime(row.get("First detected"),     errors="coerce")
            if pd.isna(idt) or pd.isna(fdt): continue
            lag_rows.append({
                "Device":          row.get("Name", ""),
                "CVE":             extract_cve_id(str(row.get("Vulnerability Name", ""))),
                "Product":         row.get("Affected Products", ""),
                "First Detected":  fdt.date(),
                "Patch Installed": idt.date(),
                "Lag (days)":      (idt - fdt).days,
            })
    patch_lag_df = (pd.DataFrame(lag_rows).sort_values("Lag (days)", ascending=False, ignore_index=True)
                    if lag_rows else _e)

    # ── Version drift ─────────────────────────────────────────────────────────
    drift_rows = []
    if "Matched Patch Version" in df.columns:
        df["_bp"] = df["Affected Products"].apply(get_base_product)
        for prod, grp in df.groupby("_bp"):
            vers = (grp["Matched Patch Version"].dropna().astype(str).str.strip()
                    .loc[lambda s: s.str.len() > 0].unique().tolist())
            if len(set(vers)) < 2: continue
            drift_rows.append({
                "Product":           prod,
                "Distinct Versions": len(set(vers)),
                "Versions Seen":     ", ".join(sorted(set(vers))),
                "Device Count":      grp["Name"].nunique(),
            })
    version_drift_df = (pd.DataFrame(drift_rows).sort_values(
        "Distinct Versions", ascending=False, ignore_index=True) if drift_rows else _e)

    return {
        "patch_lag_df":        patch_lag_df,
        "version_drift_df":    version_drift_df,
        "root_cause_df":       root_cause_df,
        "health_score":        health,
        "evidence_summary":    summary,
    }
