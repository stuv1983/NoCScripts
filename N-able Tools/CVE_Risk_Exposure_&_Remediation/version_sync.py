"""
version_sync.py — automatically fetch current stable browser/app versions
                  and update _baseline in config.json.

Run manually:    python version_sync.py
Or call from code:  from version_sync import sync_baselines; sync_baselines()

On each successful fetch the _baseline value in config.json is updated in-place.
Per-CVE rules are left untouched.

Sources
-------
Chrome   versionhistory.googleapis.com  (Google's own API, no key required)
Firefox  product-details.mozilla.org    (Mozilla's public JSON feed)
Edge     edgeupdates.microsoft.com      (Microsoft's public update API)
VLC      api.github.com/repos/videolan  (GitHub releases API)
"""

from __future__ import annotations

import json
import logging
import os
import urllib.request
import urllib.error
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
_TIMEOUT = 10  # seconds


def _fetch(url: str) -> dict | list:
    req = urllib.request.Request(
        url,
        headers={'User-Agent': _UA, 'Accept': 'application/json'},
    )
    with urllib.request.urlopen(req, timeout=_TIMEOUT) as r:
        return json.loads(r.read())


# ==============================================================================
# PER-PRODUCT FETCHERS
# Each returns the current stable version string, or None on failure.
# ==============================================================================

def _get_chrome() -> Optional[str]:
    """Google Chrome stable for Windows."""
    try:
        data = _fetch(
            'https://versionhistory.googleapis.com/v1/chrome/platforms/win/'
            'channels/stable/versions?filter=channel%3Dstable&orderBy=version+desc&pageSize=1'
        )
        return data['versions'][0]['version']
    except Exception as e:
        log.warning("version_sync: Chrome fetch failed — %s", e)
        return None


def _get_firefox() -> Optional[str]:
    """Mozilla Firefox latest release."""
    try:
        data = _fetch('https://product-details.mozilla.org/1.0/firefox_versions.json')
        return data['LATEST_FIREFOX_VERSION']
    except Exception as e:
        log.warning("version_sync: Firefox fetch failed — %s", e)
        return None


def _get_edge() -> Optional[str]:
    """Microsoft Edge stable."""
    try:
        data = _fetch('https://edgeupdates.microsoft.com/api/products')
        stable = next((p for p in data if p.get('Product') == 'Stable'), None)
        if stable and stable.get('Releases'):
            return stable['Releases'][0]['ProductVersion']
    except Exception as e:
        log.warning("version_sync: Edge fetch failed — %s", e)
    return None


def _get_vlc() -> Optional[str]:
    """VLC media player latest release via GitHub."""
    try:
        data = _fetch('https://api.github.com/repos/videolan/vlc/releases/latest')
        return data['tag_name'].lstrip('v')
    except Exception as e:
        log.warning("version_sync: VLC fetch failed — %s", e)
        return None


# ==============================================================================
# MAIN SYNC FUNCTION
# ==============================================================================

_FETCHERS: dict[str, callable] = {
    'chrome':  _get_chrome,
    'firefox': _get_firefox,
    'edge':    _get_edge,
    'vlc':     _get_vlc,
}


def sync_baselines(config_path: Optional[str] = None,
                   dry_run: bool = False) -> dict[str, str]:
    """
    Fetch current stable versions for all tracked products and update
    _baseline in config.json.

    Parameters
    ----------
    config_path : path to config.json (default: same directory as this file)
    dry_run     : if True, fetch and return results but don't write to disk

    Returns
    -------
    dict of product → new version string (only products that were successfully fetched)
    """
    if config_path is None:
        config_path = Path(__file__).parent / 'config.json'
    config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"config.json not found at {config_path}")

    with open(config_path, 'r', encoding='utf-8') as fh:
        cfg = json.load(fh)

    fvr = cfg.get('fixed_version_rules', {})
    updated: dict[str, str] = {}
    skipped: list[str] = []

    for product, fetcher in _FETCHERS.items():
        if product not in fvr:
            skipped.append(product)
            continue

        log.info("version_sync: fetching %s ...", product)
        new_version = fetcher()

        if not new_version:
            log.warning("version_sync: could not fetch %s — baseline unchanged", product)
            continue

        old_version = fvr[product].get('_baseline', '(not set)')
        fvr[product]['_baseline'] = new_version
        updated[product] = new_version
        log.info("version_sync: %-10s  %s  →  %s", product, old_version, new_version)

    if updated and not dry_run:
        cfg['fixed_version_rules'] = fvr
        with open(config_path, 'w', encoding='utf-8') as fh:
            json.dump(cfg, fh, indent=2)
        log.info("version_sync: config.json updated (%d product(s))", len(updated))
    elif dry_run:
        log.info("version_sync: dry run — no changes written")

    if skipped:
        log.debug("version_sync: skipped (not in fixed_version_rules): %s", skipped)

    return updated


# ==============================================================================
# CLI ENTRY POINT
# ==============================================================================

if __name__ == '__main__':
    import sys
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s  %(levelname)-8s  %(message)s',
        datefmt='%H:%M:%S',
    )

    p = argparse.ArgumentParser(
        description='Sync current stable browser/app versions into config.json baselines.'
    )
    p.add_argument('--dry-run', action='store_true',
                   help='Fetch and display results without writing to config.json')
    p.add_argument('--config', default=None, metavar='PATH',
                   help='Path to config.json (default: same folder as this script)')
    args = p.parse_args()

    results = sync_baselines(config_path=args.config, dry_run=args.dry_run)

    if results:
        print("\nUpdated baselines:")
        for product, version in results.items():
            print(f"  {product:<12} {version}")
    else:
        print("No baselines updated (fetch failed or no products configured).")
        sys.exit(1)
