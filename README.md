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
| **Mekko** | Marimekko chart | Column widths proportional to column totals, share labels, totals row |
| **Gantt** | Gantt/timeline | Numeric time units (weeks, months…), milestone diamonds for zero-duration tasks, alternating row shading |

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
  expose grouping in broadly available requirement sets), and they don't
  re-layout when you edit the data afterwards; re-insert to update.
- Harvey balls use Unicode glyphs (◔ ◑ ◕) because shape adjustment handles
  (pie angles) aren't exposed by the PowerPoint JS API.
- The agenda slide is appended at the end of the deck; drag it into position.
- Chart placement uses a fixed frame sized for 16:9 decks.
