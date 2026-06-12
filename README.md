# SlideCharts — an open think-cell-style add-in for PowerPoint

SlideCharts replicates the core workflow of [think-cell](https://www.think-cell.com/)
as a free Office.js taskpane add-in. Paste data from Excel, click insert, and get
consulting-grade charts drawn as **native PowerPoint shapes** — fully editable
afterwards with normal PowerPoint tools.

## Features

### Charts (the core)
| Chart | think-cell equivalent | Notes |
| --- | --- | --- |
| **Waterfall** | Waterfall chart | `e` / `total` / `=` markers compute subtotal & total bars, dashed connectors at cumulative levels, red/green gain-loss coloring, baseline handles negative cumulatives |
| **Stacked column** | Stacked chart | Segment labels (auto-hidden when too small, auto black/white for contrast), bold totals, legend |
| **100% stacked** | 100% chart | Columns normalized to 100%, percentage labels |
| **Clustered column** | Clustered chart | Side-by-side bars with value labels |
| **Mekko** | Marimekko chart | Column widths proportional to column totals, share labels, totals row |
| **Gantt** | Gantt/timeline | Numeric time units (weeks, months…), milestone diamonds for zero-duration tasks, alternating row shading |

### The datasheet workflow (as in think-cell's "Get started" tutorial)
- **Excel-style datasheet grid** in the taskpane: type or paste straight from
  Excel; the grid grows with the used area; **year sequences in the category
  row auto-continue** as you enter data in new columns.
- **Live preview** above the datasheet, rendered from the exact same layout
  primitives that get inserted on the slide.
- **Insert once, then update in place**: after the first insert the button
  becomes *Update chart on slide* — the previous chart's shapes are replaced,
  mirroring think-cell's datasheet→slide sync ("Insert as new chart" is one
  click away).
- **Totals are computed for you** — enter only raw numbers, never totals.

### Decorations
- **Value axis with gridlines** ("nice" tick steps, left gutter)
- **CAGR arrow** across the top, computed from first→last column totals,
  labelled think-cell-style in a white ellipse bubble: "+12.7% p.a."
- **Total difference arrow** at the right edge with a "+43%" ellipse bubble
- **Mean value line**: dashed line at the average of column totals with a
  "Ø 150"-style label
- **Number format**: auto / 0 / 1 / 2 decimals on all labels
- **Color schemes**: Office palette or think-cell's classic green/gray look

### Fidelity details (matched against think-cell's manual & screenshots)
- Negative values in stacked charts stack **below the baseline**; column
  totals are net and sit above the positive stack.
- Waterfall value labels go **inside** the bar (auto black/white) when the
  bar is tall enough, outside otherwise; totals stay outside and bold.
- The datasheet supports **transpose** (swap rows/columns), like
  think-cell's datasheet toolbar.

### Agenda
Scans your deck for slide titles, lets you edit the list, and inserts a styled
table-of-contents slide (numbered chips + divider rules).

### Smart elements
Harvey balls (0/25/50/75/100%), checkboxes, pentagon/chevron process flows,
traffic lights, and DRAFT/CONFIDENTIAL/FINAL stamps.

### Layout tools
One-click align (left/center/right/top/middle/bottom), distribute
horizontally/vertically, and same width/height/size on the current shape
selection — the first-selected shape is the size reference.

## Data formats

Paste straight from Excel (tab-separated); commas and semicolons also work.

**Waterfall** — two columns; `e` means "compute the total here":

```
Label        Value
2024 Revenue 820
Volume       95
Churn        -60
2025 Revenue e
```

**Stacked / Mekko** — matrix with categories across, series down:

```
         Q1   Q2   Q3
Americas 120  135  150
EMEA     90   95   100
```

**Gantt** — task, start, end in any numeric unit:

```
Task    Start End
Design  1     3
Build   2     6
Launch  6     6     <- start == end renders a milestone
```

Each panel has a **Load example** button.

## Getting started

```bash
npm install
npm run certs        # one-time: install trusted localhost dev certificates
npm start            # serves the add-in at https://localhost:3000
```

Then sideload `manifest.xml`:

- **PowerPoint on the web** — Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸
  Upload My Add-in ▸ pick `manifest.xml`.
- **Windows** — share a folder, add it as a Trusted Add-in Catalog
  (File ▸ Options ▸ Trust Center ▸ Trusted Add-in Catalogs), put the manifest
  there, then Insert ▸ My Add-ins ▸ Shared Folder.
- **macOS** — copy `manifest.xml` to
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`,
  then Insert ▸ My Add-ins.

A **SlideCharts** button appears on the Home tab and opens the taskpane.

Requires the PowerPointApi **1.4** requirement set for shape insertion
(Microsoft 365 desktop or PowerPoint on the web); the selection-based layout
tools need **1.5**.

## Development

The chart layout engines (`src/lib/`) are pure, dependency-free ES modules:
they turn data + a frame into a flat list of rect/line/text primitives.
`src/office/render.js` is the only place that touches Office.js for charts,
so the geometry is fully unit-testable:

```bash
npm test             # node --test, no dependencies needed
npm run icons        # regenerate assets/icon-*.png
```

```
src/
  lib/        chartmath.js (layout engines) · parse.js · format.js · palette.js
  office/     render.js · agenda.js · elements.js · layout.js
  taskpane/   taskpane.html · taskpane.css · app.js
```

## Known limitations

- Inserted charts are plain shape collections (not grouped — Office.js does not
  expose grouping in broadly available requirement sets). Editing the datasheet
  and clicking *Update chart on slide* replaces the chart in place, but moving
  individual chart shapes manually won't survive an update.
- Harvey balls use Unicode glyphs (◔ ◑ ◕) because shape adjustment handles
  (pie angles) aren't exposed by the PowerPoint JS API.
- The agenda slide is appended at the end of the deck; drag it into position.
- Chart placement uses a fixed frame sized for 16:9 decks.
