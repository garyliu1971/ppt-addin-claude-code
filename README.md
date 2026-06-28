# 🤖 Claude Code for PowerPoint — Office JS Add-in (POC)

An AI-style task pane add-in for PowerPoint that lets you manipulate slides via natural-language commands and quick actions. Uses **Office JS PowerPoint API 1.10+**.

## Features

| Capability | Description |
|---|---|
| **Shapes** | Add geometric shapes (rectangle, oval, triangle, arrow, heart, star, etc.), text boxes, modify fill/position/size/text, delete |
| **Charts** | Insert bar, column, pie, line, doughnut, area charts with custom data |
| **Tables** | Insert and populate tables with headers and data rows |
| **Masters & Layouts** | List all slide masters and layouts, apply layouts to slides |
| **Themes** | Read current theme, set theme |
| **Background** | Set slide/master/layout background color |
| **Events** | Respond to shape selection changes and slide navigation changes |

## Project Structure

```
ppt-addin-claude-code/
├── manifest.xml                          # Office Add-in manifest
├── package.json                          # NPM dependencies & scripts
├── tsconfig.json                         # TypeScript config
├── webpack.config.js                     # Webpack bundler config
├── README.md
├── assets/                               # Icon files (16, 32, 80 px)
└── src/
    ├── taskpane.html                     # Task pane HTML
    ├── taskpane.css                      # Task pane styles
    ├── taskpane.ts                       # Main UI + command parser
    ├── commands.ts                       # Ribbon command handlers
    ├── commands.html                     # Commands page (stub)
    └── services/
        ├── pptApi.ts                     # Core PowerPoint API wrapper
        ├── shapeService.ts               # Shape CRUD operations
        ├── chartTableService.ts          # Chart & Table operations
        ├── masterLayoutThemeService.ts   # Master/Layout/Theme ops
        └── eventService.ts               # Selection change event handlers
```

## Prerequisites

- **Node.js** ≥ 18
- **PowerPoint** (Desktop, Office 365 / Microsoft 365)
- **Edge WebView2** (bundled with Windows 10/11)

## Setup

```bash
# 1. Install dependencies
npm install

# 2. Generate self-signed dev certificates (first time only)
npm run cert

# 3. Build the project
npm run build

# 4. Sideload and start debugging in PowerPoint
npm run sideload
```

Alternatively, start the dev server alone:

```bash
npm run start
```

Then manually sideload `manifest.xml` via PowerPoint → Insert → Add-ins → Upload My Add-in.

## Usage

### Natural Language Commands

Type commands in the input area (Ctrl+Enter to send):

| Command | Example |
|---|---|
| `add rectangle shape` | Inserts a blue rectangle |
| `add oval shape in red` | Inserts a red oval |
| `add text box "Hello World"` | Inserts a text box |
| `make all shapes blue` | Sets fill of all shapes to blue |
| `fill shape "Rectangle" with green` | Fills a named shape |
| `add table` | Inserts a 3×4 sample table |
| `add bar chart` | Inserts a clustered bar chart |
| `add pie chart` | Inserts a pie chart |
| `apply layout "Title Slide"` | Applies a layout by name |
| `set background purple` | Sets slide background |
| `add slide` | Adds a new slide |
| `delete slide` | Deletes current slide |
| `list slides` | Lists all slides |
| `list shapes` | Lists shapes on current slide |
| `list layouts` | Lists all layouts |
| `list masters` | Lists all slide masters |
| `show theme` | Shows current theme |
| `resize shape 300×200` | Resizes last shape to 300×200 |
| `help` | Shows all commands |

### Quick Actions

One-click buttons in the sidebar:

- **📐 List All Shapes** — enumerate shapes on current slide
- **📑 List Slides** — enumerate all slides
- **🎨 List Masters** — list slide masters
- **📋 List Layouts** — list all layouts across masters
- **🌈 Get Theme** — display current theme
- **✏️ Modify Selected Shape** — demo: resize + recolor last shape
- **📊 Insert Table** — insert sample data table
- **📈 Insert Chart** — insert sample column chart
- **🎨 Set Background** — set dark background on current slide

## Events

The add-in listens for:

- **`DocumentSelectionChanged`** — fires when shape selection changes, logs selected shape names
- **`ActiveViewChanged`** — fires when user navigates between slides, logs the new slide ID

## PowerPoint API Requirement Sets

This add-in targets the following requirement sets:

- `PowerPointApi` — all available PowerPoint JS APIs
- `DialogApi 1.1` — for any dialog-based features

## Development Notes

- The `manifest.xml` uses `https://localhost:3000` — ensure your dev server matches
- For production, replace the `Id` GUID in `manifest.xml` with your own
- Icons in `assets/` are referenced by the manifest; generate actual PNG icons before production
- The command parser is simple regex-based; extend `executeCommand()` in `taskpane.ts` for more complex NLP

## License

POC — for demonstration purposes.
