# OrgJam

A FigJam widget for building and visualizing org charts and team rosters. Drop people chips onto the canvas, connect them to show reporting structure, and use counters to keep headcount stats up to date — all without leaving FigJam.

## Widget modes

OrgJam is a single widget that operates in four modes. When you first drop it onto the canvas you'll see the **Chooser**, which lets you pick a mode.

### People Chip

The core building block. Each chip represents one person and stores:

- **Name** and **Project Area** — editable inline on the card
- **Role** — choose from Design, Eng, PM, TPM, Ops, Brand, Writing, Research, or Data (color-coded pill)
- **Location**
- **Status** — Active (white card), Open headcount (grey card), or New headcount (yellow card)
- **IC / Manager toggle** — marks whether the person is an individual contributor or a manager
- **Avatar** — upload a custom profile photo (PNG, JPG, WEBP, or GIF; auto-cropped to a square)

Chips can be connected with FigJam connectors to show reporting lines.

### Counter

A live headcount widget that scans the canvas for people chips and displays a count. Place it inside a FigJam section to count only the chips in that section; place it outside any section to count across the entire page.

Configurable via the property menu:

- **Preset** — Total, Active, Open, New, or Locations (number of distinct offices)
- **Filter** — All, ICs only, or Managers only

Clicking the refresh icon re-scans the canvas and updates the number and role/location breakdown.

### Bulk Create

Creates multiple people chips at once from a variety of input formats. Paste text or JSON into the input field, or switch to **Selection** mode to read directly from shapes already on the canvas.

Supported text formats:

| Format | Example |
|---|---|
| Plain name list | `Alice`<br>`Bob` |
| CSV / TSV / pipe-delimited | `Alice, Search, Design, NYC` |
| CSV with header row | `name,role,location`<br>`Alice,Design,NYC` |
| Bulleted / indented hierarchy | `- Alice`<br>`-- Bob` |
| Arrow notation | `Alice -> Bob` |
| JSON flat array | `[{"name":"Alice","role":"Design"}]` |
| JSON 2-D array | `[["Alice","","Design"],["Bob","Alice","Eng"]]` |
| JSON tree | `{"name":"Alice","children":[{"name":"Bob"}]}` |

When the input encodes connections (hierarchy, arrows, or a JSON tree), OrgJam places the chips in a layered layout and adds FigJam connectors between them automatically.

**Selection import** — select any existing FigJam nodes and OrgJam will read the text inside each one to generate chips.

**OCR import** — upload a screenshot or photo of an existing org chart. OrgJam sends it to the [OCR.space](https://ocr.space) API, parses the detected names and titles, infers the hierarchy from their spatial layout, and creates chips with connectors.

### Chooser

The default state of a newly placed widget. Click a card to convert it into a People Chip, Counter, or Bulk Create widget.

---

## Development

This widget is built with TypeScript and compiled to a single JS bundle with esbuild.

**Prerequisites:** Node.js (includes npm) — https://nodejs.org/en/download/

**Install dependencies:**

```
npm install
```

**Build once:**

```
npm run build
```

**Watch mode** (rebuilds on save):

```
npm run watch
```

**Type-check without emitting:**

```
npm run tsc
```

**Lint:**

```
npm run lint
```

The compiled output is written to `dist/code.js`, which is what Figma loads. Load the widget in Figma via **Plugins → Development → Import plugin from manifest…** and point it at `manifest.json`.

For more on the Figma Widget API: https://www.figma.com/widget-docs/setup-guide/
