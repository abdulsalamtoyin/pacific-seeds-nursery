// GENERATED FILE, DO NOT EDIT
//
// Produced by excel_workflow/gen_spec_constants.py from
// excel_workflow/spec/nursery_spec.json. Edit the JSON and re-run
// the generator; any manual change here is overwritten.
//
// This is the spec with app_overrides applied — what the desktop app
// sees. It matches tabs_for(..., spec=app_spec()) value for value,
// which tests/test_js_parity.py checks.

export const SPEC = {
  "always_first": [
    "Home"
  ],
  "hidden": [
    "Settings"
  ],
  "default_tabs": [
    "Nursery site",
    "Material Map",
    "Field Map",
    "Nursery data",
    "Packet Prep",
    "Nursery list",
    "Replacements",
    "Planting errors",
    "Updated nursery site",
    "Fieldbook",
    "Operations",
    "Comments"
  ],
  "fan_out": {
    "Packet Prep": "planting_dates"
  },
  "conditional": {
    "Date recording": {
      "types": [
        "AB"
      ],
      "after": "Fieldbook"
    }
  },
  "extras": {
    "BC0 labels": [
      "AB"
    ],
    "BC0 TFMSA record": [
      "AB"
    ],
    "Pulling bags": [
      "AB"
    ],
    "TFMSA Spray plots": [
      "AB"
    ],
    "Hy Heights": [
      "Hybrid"
    ]
  },
  "nursery_types": [
    "Selection",
    "AB",
    "Hybrid",
    "Other"
  ],
  "fieldbook_columns": [
    "Range",
    "Row",
    "Material ID",
    "Source ID",
    "Gen",
    "CMS",
    "Plot",
    "Comments"
  ],
  "date_recording_columns": [
    {
      "name": "Group",
      "width": 6.29,
      "hidden": true
    },
    {
      "name": "Range",
      "width": 6.86,
      "hidden": false
    },
    {
      "name": "Row",
      "width": 5.43,
      "hidden": false
    },
    {
      "name": "O/E",
      "width": 9.14,
      "hidden": true
    },
    {
      "name": "S 1",
      "width": 5.14,
      "hidden": false
    },
    {
      "name": "S1 Month",
      "width": 8.43,
      "hidden": false
    },
    {
      "name": "S 2",
      "width": 5.57,
      "hidden": false
    },
    {
      "name": "S2 Month",
      "width": 8.43,
      "hidden": false
    },
    {
      "name": "S 3",
      "width": 5.14,
      "hidden": false
    },
    {
      "name": "S3 Month",
      "width": 8.43,
      "hidden": false
    },
    {
      "name": "S 4",
      "width": 5.14,
      "hidden": false
    },
    {
      "name": "S4 Month",
      "width": 8.43,
      "hidden": false
    },
    {
      "name": "Material ID",
      "width": 21.43,
      "hidden": false
    },
    {
      "name": "Source ID",
      "width": 28.57,
      "hidden": false
    },
    {
      "name": "Gen",
      "width": 5.14,
      "hidden": false
    },
    {
      "name": "CMS",
      "width": 4.71,
      "hidden": false
    },
    {
      "name": "In. Code",
      "width": 8.43,
      "hidden": false
    }
  ],
  "nursery_data_labels": [
    "GPS coordinates",
    "Planter",
    "Seeds/side",
    "Plot length",
    "Alley way spacing",
    "1st DOP and rows",
    "2nd DOP and rows",
    "3rd DOP and rows",
    "Tag rows",
    "Tagging date",
    "1st bagging date",
    "Last bagging date",
    "1st crossing date",
    "Last crossing date",
    "Unused A-line bags removal date",
    "Injection dates",
    "Opening bags date",
    "Closing bags date",
    "Harvest date"
  ],
  "nursery_data_widths": {
    "A": 10.71,
    "B": 58.29,
    "D": 9.14,
    "F": 29.14,
    "G": 59.57,
    "H": 9.14
  },
  "fieldbook_print": {
    "orientation": "landscape",
    "paper": "A4",
    "duplex": "short-edge",
    "print_title_rows": "$1:$1",
    "center_footer": "&P/&N",
    "right_header": "filename"
  },
  "_app_overrides_note": [
    "Desktop-app-only changes from 'Toyin- Updated app changes.docx'.",
    "Merged by app_spec() and by gen_spec_constants.render_js(); ignored by",
    "load_spec() consumers, so the Excel workbook is unaffected. Delete this",
    "block once the VBA carries the same changes."
  ],
  "nursery_data_defaults": {
    "Planter": "Almaco Precision Planter",
    "Seeds/side": "34",
    "Plot length": "4.5m",
    "Alley way spacing": "0.75m"
  }
};

export function defaultTabs() {
  return [...SPEC.default_tabs];
}

// Mirrors _anchor() in excel_workflow/spec/loader.py. A rule naming
// both anchors, or neither, has no defensible placement.
function anchorOf(name, rule) {
  const hasBefore = 'before' in rule;
  const hasAfter = 'after' in rule;
  if (hasBefore === hasAfter) {
    throw new Error(
      `conditional rule '${name}' must set exactly one of ` +
      `'before' or 'after'`);
  }
  return hasBefore ? ['before', rule.before] : ['after', rule.after];
}

// Mirrors tabs_for() in excel_workflow/spec/loader.py.
export function tabsFor(nurseryTypes, plantingDates) {
  if (plantingDates < 1) {
    throw new RangeError(
      `plantingDates must be >= 1, got ${plantingDates}`);
  }
  const selected = new Set(nurseryTypes);
  const hits = (types) => types.some((t) => selected.has(t));

  const entries = Object.entries(SPEC.conditional);
  const anchors = new Map(
    entries.map(([name, rule]) => [name, anchorOf(name, rule)]));
  const matching = (position, name) => entries
    .filter(([cond, rule]) => anchors.get(cond)[0] === position
      && anchors.get(cond)[1] === name && hits(rule.types))
    .map(([cond]) => cond);

  const tabs = [...SPEC.always_first, ...SPEC.hidden];
  for (const name of SPEC.default_tabs) {
    tabs.push(...matching('before', name));
    if (name in SPEC.fan_out) {
      for (let i = 1; i <= plantingDates; i++) tabs.push(`${name} ${i}`);
    } else {
      tabs.push(name);
    }
    tabs.push(...matching('after', name));
  }
  for (const [tab, types] of Object.entries(SPEC.extras)) {
    if (hits(types)) tabs.push(tab);
  }
  return tabs;
}
