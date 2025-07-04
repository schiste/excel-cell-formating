# 📘 FormatCellsByCriteriaWithProgress

An optimized Excel VBA macro to format cells based on content with a built‑in progress indicator.

---

## 🎯 Overview

- Clears all background colors.
- Applies light green `#EBF1DE` to numeric constants.
- Applies light blue `#B8CCE4` to formulas referencing other sheets **only if the result is numeric**.
- Displays a dynamic progress bar in Excel’s Status Bar during execution.

---

## ⚙️ Requirements

- Excel with VBA support (Windows or macOS).
- No external dependencies.

---

## 🚀 Installation

1. Open your `.xlsm` workbook.
2. Press `⌘⌥F11` to open the VBA editor (or `Alt + F11` on Windows).
3. Insert a new module.
4. Paste the `FormatCellsByCriteriaWithProgress` macro into the module.
5. Save the workbook as an **Excel Macro‑Enabled Workbook (`.xlsm`)**.

---

## ▶️ Usage

1. Open the workbook.
2. Press `⌘ + F8` (or `Alt + F8` on Windows).
3. Select **FormatCellsByCriteriaWithProgress** and click **Run**.
4. Watch the live progress bar in Excel’s lower-left status bar.

---

## 📘 Code Sample

```vb
Sub FormatCellsByCriteriaWithProgress()
  ' [Optimized code with performance settings, progress bar, and selective formatting]
End Sub
