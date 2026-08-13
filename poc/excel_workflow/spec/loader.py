"""Load nursery_spec.json — the single source of truth for workbook structure.

Both the openpyxl builder and the generated SpecConstants.bas read from here,
so tab names and ordering cannot drift between them.
"""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

SPEC_PATH = Path(__file__).resolve().parent / "nursery_spec.json"


@lru_cache(maxsize=1)
def load_spec() -> dict:
    with SPEC_PATH.open(encoding="utf-8") as fh:
        return json.load(fh)


def default_tabs() -> list[str]:
    """The 11 named default entries, before fan-out and conditionals."""
    return list(load_spec()["default_tabs"])


def tabs_for(nursery_types, planting_dates: int) -> list[str]:
    """Full ordered sheet list for a workbook.

    Packet Prep fans out to one tab per planting date; Date recording appears
    only for AB; extras are unioned across the selected nursery types and
    appended after the core tabs.
    """
    if planting_dates < 1:
        raise ValueError(f"planting_dates must be >= 1, got {planting_dates}")

    spec = load_spec()
    selected = set(nursery_types)

    tabs: list[str] = list(spec["always_first"])
    tabs += spec["hidden"]

    conditional = spec["conditional"]
    for name in spec["default_tabs"]:
        # A conditional tab is placed immediately before its anchor, so the
        # anchor's own position in default_tabs stays authoritative.
        for cond_name, rule in conditional.items():
            if rule["before"] == name and selected & set(rule["types"]):
                tabs.append(cond_name)
        if name in spec["fan_out"]:
            tabs += [f"{name} {i}" for i in range(1, planting_dates + 1)]
        else:
            tabs.append(name)

    tabs += [t for t, types in spec["extras"].items() if selected & set(types)]
    return tabs
