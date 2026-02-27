# CLAUDE.md — AI Assistant Guide for excel-coding

## Project Overview

This repository contains **Excel Sequential Data Copier** (엑셀 순차 복사기), a Windows-specific Excel VBA macro utility. The tool automates sequential clipboard-based data copying from an Excel worksheet using keyboard shortcuts, designed to work in corporate environments with restricted VBA object access.

## Repository Structure

```
excel-coding/
├── CLAUDE.md       # This file — AI assistant guidance
├── README.md       # User-facing documentation (written in Korean)
└── code.txt        # VBA macro source code (copy into Excel VBA editor)
```

This is an intentionally minimal repository. There is no build system, package manager, test framework, or CI/CD pipeline — VBA macros execute directly inside Microsoft Excel on Windows.

## Language & Platform

| Attribute | Details |
|-----------|---------|
| **Language** | VBA (Visual Basic for Applications) |
| **Host Application** | Microsoft Excel (Windows only) |
| **Documentation Language** | Korean |
| **OS Requirement** | Windows (uses `user32.dll` and `kernel32.dll`) |
| **Excel Compatibility** | Both 32-bit (`#Else`) and 64-bit (`#If VBA7`) via conditional compilation |

## Code Architecture

### File: `code.txt`

The sole source file, structured in three sections:

#### 1. System Setup (Lines 1–31)
Windows API declarations using VBA conditional compilation to support both 64-bit (`VBA7`) and legacy 32-bit Excel:
- `user32.dll`: `GetAsyncKeyState`, `OpenClipboard`, `EmptyClipboard`, `CloseClipboard`, `SetClipboardData`
- `kernel32.dll`: `Sleep`, `GlobalAlloc`, `GlobalLock`, `GlobalUnlock`, `CopyMemory` (alias `RtlMoveMemory`)
- Constants: `GMEM_MOVEABLE = &H2`, `CF_UNICODETEXT = 13`

#### 2. Main Subroutine: `StartSequentialCopy()` (Lines 36–157)
Entry point invoked by the user via Excel's macro runner (`Alt+F8`):
1. Reads user-configurable constants (data range, decimal formatting)
2. Converts column letters to column numbers via `ws.Range(CHAR & "1").Column`
3. Loads all cell values from the configured range into a `Collection`
4. Applies decimal rounding using `Format()` with dynamic format strings
5. Enters a polling loop monitoring keyboard state with `GetAsyncKeyState`:
   - `&HA0` (Left Shift) → copy next item to clipboard, advance index
   - `&HA1` (Right Shift) → go back one item, re-copy previous value
   - `&H1B` (ESC) → exit loop and terminate

#### 3. Helper Function: `PutInClipboard()` (Lines 159–176)
Low-level clipboard writer using direct Windows API calls (avoids `DataObject` which may be blocked in restricted environments):
1. Opens clipboard with `OpenClipboard(0)`
2. Allocates global memory with `GlobalAlloc`
3. Copies string bytes with `CopyMemory` (Unicode, `LenB + 2` bytes for null terminator)
4. Writes to clipboard with `SetClipboardData(CF_UNICODETEXT, ...)`
5. Returns `Boolean` success flag

## User-Configurable Constants

All user configuration is embedded as `Const` declarations inside `StartSequentialCopy()` (Lines 54–63). There is **no separate config file**.

| Constant | Default | Description |
|----------|---------|-------------|
| `START_ROW_NUM` | `1` | First data row (set to `2` if there's a header row) |
| `START_COL_CHAR` | `"B"` | First column letter to copy from |
| `END_COL_CHAR` | `"G"` | Last column letter to copy from |
| `DIGITS_DEFAULT` | `3` | Decimal places for standard columns |
| `SPECIAL_COL_CHAR` | `"E"` | Column letter that gets special decimal treatment |
| `DIGITS_SPECIAL` | `5` | Decimal places for the special column |

## How to Run (No Build Step)

VBA macros have no compilation or build step. To use the macro:

1. Open target Excel workbook
2. Press `Alt + F11` to open the VBA editor
3. Insert a new module (`Insert` → `Module`)
4. Copy the entire contents of `code.txt` and paste into the module
5. Close the VBA editor
6. Press `Alt + F8`, select `StartSequentialCopy`, click **Run**

## Key Conventions

### Code Style
- `Option Explicit` is used — all variables must be declared before use
- Section delimiters use lines of `=` characters in comments
- Comments are written in **Korean**
- Variable names are in **English**
- The user-editable zone is clearly marked with `⚡ [사용자 설정 구역]` comments

### VBA-Specific Patterns
- Column letters are converted to numbers via `ws.Range(CHAR & "1").Column` (standard VBA idiom)
- Format strings are built dynamically: `"0." & String(DIGITS, "#")`
- `DoEvents` is called in the polling loop to prevent Excel from freezing
- `Sleep 300` debounces key press detection after each action
- `Application.StatusBar` is used to display real-time progress to the user

### Clipboard Implementation Note
The `PutInClipboard` function deliberately avoids VBA's `DataObject` (from `Microsoft Forms 2.0 Object Library`) because that reference may be unavailable in security-restricted corporate Excel environments. The raw Windows API approach is intentional and necessary.

## Known Issues / Gotchas

- **`IsNumeric()` on line 94** is a VBA built-in function — it does **not** need to be defined in this module. It is correctly used as provided by the VBA runtime.
- **Platform lock-in**: The Windows API declarations mean this macro **cannot run on macOS Excel**. Any port would require removing/replacing all `Declare` statements and the clipboard logic.
- **`LongPtr` type**: Used in `PutInClipboard` for the `hGlobal` and `lpString` variables, but the function signature itself uses `LongPtr` only under `VBA7`. Ensure the `#If VBA7` block in the API declarations matches the Excel version in use.
- **No error handling**: The main loop has no `On Error` handler. Runtime errors (e.g., clipboard locked by another process) will display a generic VBA error dialog and halt execution.

## Development Workflow

Since there is no automated build/test/lint pipeline:

1. **Edit** `code.txt` in any text editor
2. **Manually test** by pasting into Excel VBA editor and running `StartSequentialCopy`
3. **Commit** changes with a clear message describing what was changed and why
4. **Push** to the feature branch

### Git Branch Convention
- Active development happens on feature branches named `claude/<task-id>`
- The `master` branch holds stable, tested versions
- Commit messages are written in **English**

## What AI Assistants Should Know

- **Do not add a build system, test framework, or package manager** — this project intentionally has none. VBA macros are not unit-testable in the traditional sense.
- **Do not convert to Python/JavaScript/etc.** unless the user explicitly requests a rewrite.
- **Preserve the Korean comments** — they are intentional and serve the target user base.
- **Preserve `Option Explicit`** — removing it would reduce code quality.
- **Preserve the `#If VBA7` / `#Else` block** — removing it breaks compatibility with 32-bit Excel.
- **The `PutInClipboard` approach is intentional** — do not simplify it to use `DataObject`.
- When modifying user-configurable constants, only change values inside the `[사용자 설정 구역]` block (Lines 51–67).
- This code uses **Windows-only APIs**. Any suggested alternatives must also work without `DataObject` in a restricted corporate environment.
