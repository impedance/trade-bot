# Repository Guidelines

## Project Structure & Module Organization
- `robot.xlsb` is the primary binary workbook; open it in Excel to view sheets, formulas, and VBA macros. It implements a trading robot that connects to QUIK via DDE to read market data and place orders. Текущая цель репозитория — анализ существующего робота, без переписывания его логики. Следуем шагам из `plan.md`, отмечая прогресс чекбоксами.
- `robot-analyze/` holds an extracted view of the workbook contents (XML/bin) for inspection; key parts include `xl/vbaProject.bin` (macros) and `xl/worksheets/` (sheet data).
- Keep workbook changes paired: edit in Excel, then refresh the extracted copy with a clean export so diffs stay meaningful.

## Build, Test, and Development Commands
- Preview workbook contents without Excel: `unzip -l robot.xlsb` lists entries; `bsdtar -xf robot.xlsb -s '/^/robot-analyze\//'` refreshes the extracted tree.
- Create a safe working copy before macro edits: `cp robot.xlsb robot.backup.xlsb`.
- Use Excel’s VBA editor for code changes; export/import modules (`File > Export File...`) so logic can be reviewed and versioned.

## Coding Style & Naming Conventions
- VBA: use explicit declarations (`Option Explicit`), camelCase for procedures (`calculateSignals`), and ALL_CAPS for constants.
- Sheets/tables: prefer descriptive sheet names (e.g., `Signals`, `Orders`) and consistent named ranges for cross-sheet references.
- Keep macros modular—separate data IO, calculations, and orchestration into distinct modules.

## Testing Guidelines
- Functional: open `robot.xlsb`, enable macros, and run macros against a small test dataset; verify key sheets (orders, signals) for expected values.
- Regression: record pre/post cell snapshots for critical ranges before committing; rerun after macro changes.
- Error handling: ensure macros handle missing data gracefully and surface clear `MsgBox` errors rather than silent failures.

## Commit & Pull Request Guidelines
- Commits: prefer concise messages in the form `area: change` (e.g., `vba: fix order sizing guard`); group related workbook and extracted updates together.
- Pull requests: include a brief summary of the change, steps to reproduce/verify in Excel, and screenshots or cell references for before/after results on key sheets.
- Note any macro security settings or external data connections touched; call out if users must re-enable trusted locations.

## Security & Configuration Tips
- Remove or obfuscate credentials and connection strings from macros before sharing.
- Store macros in trusted locations and sign them if distributing broadly to reduce security prompts.

## Agent Instructions
- При общении с пользователями отвечайте на русском языке, сохраняя профессиональный и лаконичный тон.
- Работайте по `plan.md`: обновляйте чекбоксы прогресса, фиксируйте выполненные/оставшиеся задачи по анализу.
- Все материалы анализа сохраняйте в Markdown-файлы внутри `docs/`, поддерживая краткие структурированные сводки.
