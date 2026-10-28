"""Load nursery_spec.json — the single source of truth for workbook structure.

Both the openpyxl builder and the generated SpecConstants.bas read from here,
so tab names and ordering cannot drift between them.

The desktop app has since moved ahead of the workbook: it splits Replacements
from Planting errors and anchors Date recording after Fieldbook. Rather than
fork the file, those differences live in an ``app_overrides`` block that only
``app_spec()`` merges. The workbook keeps reading the base keys and keeps
building exactly as before; when the VBA catches up, the block is deleted and
the two converge again.
"""
from __future__ import annotations

import copy
import json
from functools import lru_cache
from pathlib import Path

SPEC_PATH = Path(__file__).resolve().parent / "nursery_spec.json"


@lru_cache(maxsize=1)
def load_spec() -> dict:
    with SPEC_PATH.open(encoding="utf-8") as fh:
        return json.load(fh)


@lru_cache(maxsize=1)
def app_spec() -> dict:
    """The spec as the desktop app sees it: base plus ``app_overrides``.

    Merging is one level deep — a dict override updates the matching keys and
    leaves the rest, anything else replaces wholesale. That is enough for the
    overrides we have and keeps the merge easy to reason about.
    """
    spec = copy.deepcopy(load_spec())
    for key, value in spec.pop("app_overrides", {}).items():
        if isinstance(value, dict) and isinstance(spec.get(key), dict):
            spec[key] = {**spec[key], **value}
        else:
            spec[key] = value
    return spec


def default_tabs() -> list[str]:
    """The named default entries, before fan-out and conditionals."""
    return list(load_spec()["default_tabs"])


def _anchor(name: str, rule: dict) -> tuple[str, str]:
    """Return ``(position, anchor_tab)`` for a conditional rule.

    A rule that names both anchors, or neither, has no single defensible
    placement — so it is refused rather than resolved by precedence.
    """
    has_before = "before" in rule
    has_after = "after" in rule
    if has_before == has_after:
        raise ValueError(
            f"conditional rule {name!r} must set exactly one of "
            f"'before' or 'after'")
    return ("before", rule["before"]) if has_before else ("after", rule["after"])


def tabs_for(nursery_types, planting_dates: int,
             spec: dict | None = None) -> list[str]:
    """Full ordered sheet list for a workbook.

    Packet Prep fans out to one tab per planting date; conditional tabs appear
    only for their nursery types, anchored immediately before or after a named
    tab; extras are unioned across the selected types and appended after the
    core tabs.

    Pass ``spec=app_spec()`` for the desktop app's tab list, or leave it as the
    default for the workbook's.
    """
    if planting_dates < 1:
        raise ValueError(f"planting_dates must be >= 1, got {planting_dates}")

    spec = load_spec() if spec is None else spec
    selected = set(nursery_types)

    tabs: list[str] = list(spec["always_first"])
    tabs += spec["hidden"]

    conditional = spec["conditional"]
    anchors = {name: _anchor(name, rule) for name, rule in conditional.items()}

    def matching(position: str, name: str) -> list[str]:
        return [cond for cond, rule in conditional.items()
                if anchors[cond] == (position, name)
                and selected & set(rule["types"])]

    for name in spec["default_tabs"]:
        # A conditional is placed relative to its anchor, so the anchor's own
        # position in default_tabs stays authoritative.
        tabs += matching("before", name)
        if name in spec["fan_out"]:
            tabs += [f"{name} {i}" for i in range(1, planting_dates + 1)]
        else:
            tabs.append(name)
        tabs += matching("after", name)

    tabs += [t for t, types in spec["extras"].items() if selected & set(types)]
    return tabs
