"""
formatting.py — named color palette and shared xlsxwriter format factories
for the dashboard workbook.

Note: some sheet-builder modules still have local add_format() calls using
the same hex colors under different names — not yet all migrated to COLORS
below.
"""
from __future__ import annotations

# ── Named palette ──────────────────────────────────────────────────────────────
# Names describe the color itself, not "what it means" — the same hex is
# reused for different meanings in different sheets (peach = "unmanaged app"
# in the product-sheet legend, but is also just a neutral highlight
# elsewhere). Keying by meaning would be actively misleading; keying by the
# color keeps this file honest about what it actually controls.
COLORS = {
    # Text colors
    'GRAY_TEXT':        '#595959',
    'RED_TEXT':         '#C00000',
    'GREEN_TEXT':       '#375623',
    'DARK_RED_TEXT':    '#9C0006',
    'AMBER_TEXT':       '#7F6000',
    'NAVY_TEXT':        '#1F3864',
    'BROWN_TEXT':       '#7B3F00',
    'LINK_BLUE':        '#0563C1',
    'SLATE':            '#2E4057',   # used as both bg_color and font_color at different call sites
    'MED_BLUE_TEXT':    '#2E75B6',

    # Background fills
    'AMBER_BG':         '#FFF2CC',
    'LIGHT_GREEN_BG':   '#E2EFDA',
    'PEACH_BG':         '#FCE4D6',
    'LIGHT_GRAY_BG':    '#F2F2F2',
    'LIGHT_BLUE_BG':    '#D6E4F0',
    'DARK_BLUE_BG':     '#1F4E79',
    'ORANGE_BG':        '#ED7D31',
    'PALE_BLUE_BG':     '#EBF3FB',
    'LIGHT_AMBER_BG':   '#FFF3E0',
    'GRAY_BG':          '#D9D9D9',
    'PINK_RED_BG':      '#FFEBEE',
    'STEEL_BLUE_BG':    '#BDD7EE',
    'PALE_YELLOW_BG':   '#FFFFE0',
    'LIGHT_ORANGE_BG':  '#FFE0CC',
    'LIGHT_RED_BG':     '#FFCCCC',
    'OFF_WHITE_BG':     '#F5F5F5',
    'PINK_BG':          '#F2CEEF',
    'TEAL_BG':          '#D9F0F4',
    'DARK_RED_BG':      '#7B0000',
    'GREEN_ACCENT_BG':  '#70AD47',
    'ROSE_BG':          '#FFC7CE',
    'NEAR_WHITE_BG':    '#F9F9F9',
    'WHITE':            '#FFFFFF',
}


def get_workbook_styles(wb) -> dict:
    """The common, reusable format set — sheet builders should pull from
    this dict instead of redefining their own copy of the same style."""
    C = COLORS
    return {
        'title':        wb.add_format({'bold': True, 'font_size': 14,
                                       'bg_color': C['DARK_BLUE_BG'], 'font_color': 'white', 'border': 1}),
        'header':       wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': C['GRAY_BG'], 'border': 1}),
        'sub_header':   wb.add_format({'bold': True, 'bg_color': C['LIGHT_BLUE_BG'], 'border': 1}),
        'section':      wb.add_format({'bold': True, 'bg_color': C['LIGHT_GRAY_BG'], 'border': 1}),
        'alert':        wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': C['RED_TEXT'], 'font_color': 'white'}),
        'warn':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': C['ORANGE_BG'], 'font_color': 'white'}),
        'info':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': C['GREEN_TEXT'], 'font_color': 'white'}),
        'bold':         wb.add_format({'bold': True}),
        'note':         wb.add_format({'italic': True, 'font_color': C['GRAY_TEXT']}),
        'note_sm':      wb.add_format({'italic': True, 'font_color': C['GRAY_TEXT'], 'font_size': 9}),
        'note_amber':   wb.add_format({'italic': True, 'font_color': C['AMBER_TEXT'], 'font_size': 8,
                                       'bg_color': C['PALE_YELLOW_BG'], 'border': 1, 'text_wrap': True}),
        'link':         wb.add_format({'font_color': 'blue', 'underline': True}),
        'up':           wb.add_format({'font_color': C['RED_TEXT'], 'bold': True}),
        'down':         wb.add_format({'font_color': C['GREEN_TEXT'], 'bold': True}),
        'same':         wb.add_format({'font_color': C['GRAY_TEXT']}),
        'row_red':      wb.add_format({'bg_color': C['PEACH_BG']}),
        'row_green':    wb.add_format({'bg_color': C['LIGHT_GREEN_BG']}),
        'row_amber':    wb.add_format({'bg_color': C['AMBER_BG']}),
        'row_blue':     wb.add_format({'bg_color': C['STEEL_BLUE_BG'], 'font_color': C['NAVY_TEXT']}),  # deeper blue — resolved/confirmed
        'row_pink':     wb.add_format({'bg_color': C['PINK_BG']}),
        'row_teal':     wb.add_format({'bg_color': C['TEAL_BG']}),
        'row_missing':  wb.add_format({'bg_color': C['ROSE_BG'], 'font_color': C['DARK_RED_TEXT']}),
        'score_good':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': C['GREEN_TEXT']}),
        'score_warn':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': C['AMBER_TEXT']}),
        'score_bad':    wb.add_format({'bold': True, 'font_size': 18, 'font_color': C['DARK_RED_TEXT']}),
    }


def get_band_formats(wb) -> dict:
    """
    Bold, color-coded 4-tier band formats — for compact risk/age breakdowns
    like N-Day Exposure Age (critical / high / amber / ok bands with counts).

    Distinct from get_workbook_styles()'s row_red/row_amber/row_green: those
    are meant for shading a full data row, these are bold and meant to sit
    next to a number in a small summary table. Same underlying colors,
    different weight — kept as a separate factory rather than overloading
    one dict with both variants.
    """
    C = COLORS
    return {
        'header':   wb.add_format({'bold': True, 'bg_color': C['SLATE'], 'font_color': 'white', 'border': 1}),
        'critical': wb.add_format({'bold': True, 'bg_color': C['RED_TEXT'], 'font_color': 'white'}),
        'high':     wb.add_format({'bg_color': C['PEACH_BG'], 'bold': True}),
        'amber':    wb.add_format({'bg_color': C['AMBER_BG'], 'bold': True}),
        'ok':       wb.add_format({'bg_color': C['LIGHT_GREEN_BG'], 'bold': True}),
        'label':    wb.add_format({'font_size': 9, 'italic': True, 'font_color': C['GRAY_TEXT']}),
    }


def build_legend_entries(stale_warning_days: int) -> list:
    """The row-color legend shown at the bottom of every product sheet."""
    C = COLORS
    return [
        (C['STEEL_BLUE_BG'], 'blue row',   'Patch via RMM — install confirmed after CVE first detected'),
        (C['LIGHT_ORANGE_BG'], 'orange row', 'Known active exploit — unresolved, prioritise immediately'),
        (C['LIGHT_AMBER_BG'], 'amber-orange row',
         f'Approaching stale — device offline \u2265 {stale_warning_days}d; '
         f'patch confirmation unreliable (overrides blue)'),
        (C['AMBER_BG'], 'yellow row', 'Coverage gap — device not in patch report'),
        (C['PEACH_BG'], 'peach row',  'Unmanaged app — product not tracked in patch report'),
        (C['PINK_BG'], 'pink row',   'Detection mismatch — CVE detected but no matching patch found'),
        (C['TEAL_BG'], 'teal row',   'Patch installing — patch is in progress, re-check after next RMM sync'),
        (C['LIGHT_RED_BG'], 'red row',    'Unresolved — patch not yet applied'),
    ]