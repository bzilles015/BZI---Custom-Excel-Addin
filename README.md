[README.md](https://github.com/user-attachments/files/24133599/README.md)
# Excel FP&A Modeling Shortcuts Add-In (VBA)

A lightweight Excel **VBA add-in** that adds fast, keyboard-driven formatting and modeling helpers for FP&A / financial modeling work.

The add-in is organized as exported `.bas` modules so the VBA source can be version-controlled in Git/GitHub (Excel files themselves are binary and painful to diff).

## What it does

### Format cycles (keyboard-first)
- Cycle common **number formats** (number/date/currency/percent/other)
- Adjust **decimal places** up/down on demand
- Scale values **up/down** (e.g., units conversion)

### Styles & layout
- Auto-color formulas vs constants
- Cycle **font / fill / text case / font color**
- Apply consistent **input** and **header** styles
- Zoom, font-size, indent, center-across-selection helpers
- Apply/clear **“zero-check”** conditional formatting

### Borders
- Cycle top/bottom/left/right borders
- Apply outline + inside borders

### Unit tags
- Uniform “unit tags” appended to cell text (e.g., `[#]`, `[%]`, `[mln $]`)
- Remove the last bracket tag from each selected cell

### Engine features
- Shared cycle state (so repeated taps move through options)
- Optional **Performance Mode** toggle for heavy operations
- Simple logging hooks and an undo buffer (where applicable)

## Installation (recommended: use Releases)

1. Download the latest `.xlam` from **GitHub Releases**.
2. In Windows: right-click the downloaded file → **Properties** → check **Unblock** (if present).
3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**
4. Select the `.xlam`, enable it, and restart Excel.

> If you don’t have a compiled `.xlam` yet, see “Developer setup” below.

## Developer setup (build your own add-in)

1. Create a new workbook in Excel.
2. Press `Alt+F11` to open the VBA editor.
3. **Import modules** from `src/vba/*.bas`:
   - File → Import File…
4. Save as an add-in:
   - File → Save As → **Excel Add-In (*.xlam)**

Tip: Treat the `.bas` files as the source of truth. Only ship `.xlam` via Releases.

## Keyboard shortcuts (from `modBindings`)

> Shortcuts are defined in `src/vba/modBindings.bas` via `Application.OnKey`.
> If your keyboard layout differs, adjust the OnKey strings.

### modFormatCycles
- **Ctrl+Shift+1** — CycleNumberFormat
- **Ctrl+Shift+3** — CycleDateFormat
- **Ctrl+Shift+4** — CycleCurrencyFormat
- **Ctrl+Shift+5** — CyclePercentFormat
- **Ctrl+Shift+8** — CycleOtherNumbers
- **Ctrl+Shift+.** — IncreaseDecimal
- **Ctrl+Shift+,** — DecreaseDecimal
- **Alt+Shift+<** — ScaleUp
- **Alt+Shift+>** — ScaleDown

### modStyles
- **Ctrl+Alt+A** — AutoColorSelection
- **Ctrl+'** — CycleFont
- **Ctrl+Shift+K** — CycleFill
- **Ctrl+Alt+Shift+I** — CycleTextCase
- **Ctrl+Shift+C** — CycleFontColor
- **Ctrl+Shift+F** — IncreaseFontSize
- **Ctrl+Shift+G** — DecreaseFontSize
- **Ctrl+Shift+N** — InsertStaticNow
- **Ctrl+Alt+E** — CenterAcrossSelection
- **Ctrl+Shift+]** — IndentIn
- **Ctrl+Shift+[** — IndentOut
- **(see modBindings for ZoomIn/ZoomOut mapping)**

Input/Header tools:
- **Ctrl+Alt+Shift+U** — CycleInputStyle
- **Ctrl+Alt+Shift+H** — CycleHeaderStyle
- **Ctrl+Alt+Shift+Y** — InsertHeadersFromPrompt
- **Ctrl+Alt+Shift+D** — InsertVarianceHeaders

Zero-check CF:
- **Ctrl+Alt+Shift+Z** — ApplyZeroCheckCF
- **Ctrl+Alt+Shift+X** — ClearZeroCheckCF

### modBorders
- **Ctrl+Alt+Shift+↑** — BorderTop
- **Ctrl+Alt+Shift+↓** — BorderBottom
- **Ctrl+Alt+Shift+←** — BorderLeft
- **Ctrl+Alt+Shift+→** — BorderRight
- **Ctrl+Alt+Shift+B** — BordersOutlineInside

### modUnitTags
- **Ctrl+Alt+Shift+T** — CycleUnitTag_Value_Uniform
- **Ctrl+Alt+Shift+O** — CycleUnitTag_Duration_Uniform
- **Ctrl+Alt+Shift+P** — CycleUnitTag_Rate_Uniform
- **Ctrl+Alt+Shift+Backspace** — RemoveUnitTag

### modCore helpers
- **Ctrl+Alt+Shift+A** — MakeRefsAbsolute
- **Ctrl+Alt+Shift+R** — MakeRefsRelative
- **Ctrl+Alt+Shift+N** — GoToNextBlank
- **Ctrl+Alt+Shift+E** — GoToNextError
- **Ctrl+Alt+Shift+L** — BreakExternalLinksInSelection
- **Ctrl+Alt+Shift+V** — PasteValuesKeepFormat
- **Ctrl+Alt+Shift+M** — TogglePerformanceMode

## Repo layout

```
.
├─ src/
│  └─ vba/
│     ├─ modCore.bas
│     ├─ modFormatCycles.bas
│     ├─ modStyles.bas
│     ├─ modBorders.bas
│     ├─ modUnitTags.bas
│     └─ modBindings.bas
├─ dist/              # optional: compiled .xlam (recommended via Releases, not git)
└─ README.md
```

## Versioning & releases

Use **Semantic Versioning** (`vMAJOR.MINOR.PATCH`) for releases.
- MAJOR: breaking changes to hotkeys/behavior
- MINOR: new features/macros
- PATCH: bug fixes / comment-only cleanup

Create a Git tag (e.g., `v1.2.0`) and attach the compiled `.xlam` file to a GitHub Release.

## Troubleshooting

- **Shortcuts don’t work**: run `BindAllKeys` manually (Macro dialog) or re-open Excel.
- **Macros blocked**: check Excel Trust Center macro settings.
- **OnKey calls the wrong workbook**: the bindings target `ThisWorkbook.Name` (the add-in). If you see “...” or truncated macro names inside `modBindings`, fix those strings before building.

## License

Choose a license before making the repo public (MIT is common for small utilities).
