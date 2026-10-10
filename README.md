# VbeComplete

**IntelliJ IDEA–style autocomplete for the VBA editor.**

VbeComplete brings modern, IDE-grade code completion to the VBA editor (VBE) shipped with
Microsoft Office. Start typing a name and a live-filtered list of the variables, constants,
procedures, keywords and built-in functions that are actually in scope pops up next to the caret
— press `Tab` to accept the highlighted one.

> **The gap it fills.** The VBE only ever completes `object.member`; it never suggests *your own*
> variable or procedure names. VbeComplete does exactly that, and layers a set of IDEA-style
> editing helpers on top.

*Read this in other languages: **English** · [简体中文](README.zh-CN.md)*

## Highlights

**Completion**

- **Suggestions as you type.** Type a letter, digit or `_` and the list appears; keep typing and it filters live.
- **Fuzzy and camelCase matching.** You don't have to start at the first letter — type `ds` to get
  `dataSheet` (**D**ata**S**heet). Matched characters are highlighted so you can see *why* a name matched.
- **Quality-first ranking.** Results are ordered exact › prefix › word-initial acronym › contiguous
  block › scattered, then by hit position, span and name length. The closest match is always first,
  and a plain prefix behaves exactly like a plain prefix search.
- **Scope-aware ranking** *(IDEA's "local scope first")*. A local variable or parameter outranks a
  module-level name, which outranks a `Public` name from another module.
- **Recently-used ranking** *(IDEA's "recently used")*. Names you just accepted are promoted on ties.
- **Context-aware follow-up.** Accepting a `Sub` inserts a trailing space (`DoWork `), a `Function`
  inserts `()` with the caret placed inside (`GetName()`), a `Property` inserts nothing, and a
  keyword gets its own trailing space (`public `).
- **Unicode identifiers.** Names such as `Dim 用户名 As String` are fully supported.
- **Non-intrusive.** Suggestions appear only inside a VBE *code pane* — never in the Properties
  window, the Project explorer or the form designer.

**Editing helpers**

- Auto-paired brackets and quotes (`(`, `[`, `"`, …), with the caret placed between the two halves.
- `Backspace` deletes both halves of an empty pair, and removes a whitespace-only line in one stroke.
- `Enter` at end of line indents to the right depth, aligns `Else` / `ElseIf` / `Case`, and closes a
  block you just opened (`End Sub`, `Next i`, …).
- `Shift+Enter` starts a new line below without splitting the current one — the same indentation
  logic as `Enter`, at any column.

**Coexistence**

- After `object.`, VbeComplete steps aside and lets the VBE's own, type-aware member list take over.
- Global hooks are attached **only while a VBE code pane has focus**, so the tool is essentially
  invisible (and free) the rest of the time.

## Requirements

- Windows with Microsoft Office (any host that ships the VBE).
- Python 3.8+ (a python.org build is recommended; tick *Add Python to PATH* during setup).
- **"Trust access to the VBA project object model" must be enabled** — the tool reads the code
  model over COM.
  *File → Options → Trust Center → Trust Center Settings → Macro Settings →*
  *Trust access to the VBA project object model.*

## Installation

```bat
git clone https://github.com/CoffeeBean001/VbeComplete-V587.git
cd VbeComplete-V587
run.bat
```

`run.bat` installs the dependencies (`pywin32`, `pynput`) on first launch and then runs the tool
silently in the background. To watch the output instead, run `python main.py` in a console.

Then open the VBE (`Alt+F11`) and start typing a name — the list should appear.

- `stop.bat` — stop the background process.
- `run_tests.bat` — run the regression suite.

## Usage

### Keyboard shortcuts

| Key | Action |
| --- | --- |
| letter / digit / `_` | Open the candidate list and filter it live |
| `↑` / `↓` | Move the selection (wraps); the list scrolls when it overflows |
| `Shift+↑` / `Shift+↓` | Close the list and move the caret one line up / down |
| `Tab` | **Accept** the highlighted candidate (the only accept key) |
| `Enter` *(at end of line)* | New line with smart indentation, plus an automatic block closer |
| `Shift+Enter` | Start a new line below, leaving the current line intact |
| `Backspace` | Delete a whitespace-only line, or both halves of an empty pair |
| `Esc` | Dismiss the list |
| `←` / `→` | Dismiss the list (the key still moves the caret in the editor) |
| Click | Accept a candidate; click anywhere else to dismiss |
| `Ctrl+Space` | Force the list to open |

### What gets suggested

- Variables and constants declared in the project, plus procedures (`Sub` / `Function` /
  `Property`), modules, forms and control names.
- Without `Option Explicit`, variables that are "used but never declared" are suggested too.
- VBA keywords and built-in functions / data types (≈260 names) are included by default.
- **Enums are deliberately excluded** — the `vb*`, `xl*` and `mso*` families — to keep the list quiet.
- Visibility follows VBA's own scoping rules, so only names reachable from the caret are shown.

## How it compares to IntelliJ IDEA

VbeComplete reproduces the *feel* of IDEA's completion; it does not try to reimplement an IDE's
semantic infrastructure.

**Implemented**

| IDEA | VbeComplete |
| --- | --- |
| Completion as you type, live filtering | ✅ |
| Fuzzy / camelCase matching with highlighted hits | ✅ |
| Quality-ranked results | ✅ |
| Local scope first | ✅ |
| Recently used promoted | ✅ |
| Auto-paired brackets and quotes | ✅ |
| Paired deletion on `Backspace` | ✅ |
| Smart indentation and automatic block closers | ✅ |
| Start a new line (`Shift+Enter`) | ✅ |
| Trailing space for keywords; `()` after functions | ✅ |
| Type-aware member completion | Delegated to the VBE's native member list |

**Deliberately not implemented**

Go-to-definition, Find Usages, rename refactoring, live syntax/error checking, postfix completion
and live templates all depend on a full, always-current semantic model of the project — type
inference, symbol tables, a dependency graph. The VBE exposes only the *plain text* of a
`CodeModule`, so rebuilding a compiler front-end on top of it would be costly and fragile; the
risk/benefit simply doesn't add up.

## Configuration

Behaviour is tuned through environment variables (set them before launching):

| Variable | Default | Effect |
| --- | --- | --- |
| `VBECOMPLETE_LOG` | off | `1` writes a diagnostic log to `VbeComplete.log` (`VBECOMPLETE_LOG_FILE` relocates it). |
| `VBECOMPLETE_AUTO_PAREN` | on | `0` disables inserting `()` after procedures and functions. |
| `VBECOMPLETE_AUTO_SPACE` | on | `0` disables the automatic trailing space after keywords. |
| `VBECOMPLETE_MEMBER_POPUP` | off | `1` restores VbeComplete's own list after `object.` (off by default). |
| `VBECOMPLETE_NO_SCOPE_RANK` | off | `1` disables scope-aware ranking. |
| `VBECOMPLETE_NO_RECENT_RANK` | off | `1` disables recently-used ranking. |
| `VBECOMPLETE_RECENT_MAX` | 64 | How many recently-used names are remembered. |
| `VBECOMPLETE_NO_YIELD` | off | `1` disables yielding to the VBE's native member list. |
| `VBECOMPLETE_NO_AUTOPAIR` | off | `1` disables auto-pairing. |
| `VBECOMPLETE_NO_BS_DELLINE` / `VBECOMPLETE_NO_BS_DELPAIR` | off | `1` disables the two `Backspace` helpers. |
| `VBECOMPLETE_NO_ENTER_INDENT` / `VBECOMPLETE_NO_AUTOCLOSE` | off | `1` disables `Enter`-time indentation / automatic block closers. |

## Autostart on login

Double-click `install_autostart.bat`. It installs the dependencies and drops a shortcut to
`pythonw main.py` into your Startup folder, so the tool runs in the background after every login.
Because it is focus-driven, it attaches its hooks only while a VBE code pane is active.
Double-click `uninstall_autostart.bat` to remove it.

## Running the tests

```bat
run_tests.bat
```

- Runs offline against a real 14-module sample project in `tests/testdata/` — **no Excel or COM required**.
- Three cases read the Excel type library; on a machine without Excel they fail, so
  `962 PASS / 3 FAIL` is expected there and is **not** a code fault.
- Use a Python interpreter that has `pywin32` installed.

## Project structure

| File | Role |
| --- | --- |
| `main.py` | Entry point: global keyboard hook, event loop, wiring |
| `engine.py` | Completion engine: scoring, ranking, filtering, popup lifecycle |
| `vbe_bridge.py` | VBE COM bridge: read code, read caret, write back, caret placement |
| `parser.py` | VBA text parsing (statements, scope, comments and strings) |
| `vba_builtins.py` | VBA built-in vocabulary (functions, types, keywords, enums) |
| `ui.py` | The floating popup (tkinter) and its screen-avoidance logic |
| `log.py` | Optional diagnostic log (off by default, zero overhead) |
| `tests/` | Regression suite |

## Troubleshooting

- **No popup appears** — make sure *Trust access to the VBA project object model* is enabled.
- **Is the tool running?** — look for `pythonw.exe` in Task Manager; if it is missing, run `run.bat` again.
- **A list won't go away, or accepts everything** — you likely have stale instances: run `stop.bat`,
  then `run.bat`.
- **Nothing appears after `.` inside a `With` block** — that position is intentionally left to the
  VBE's own member list, which needs the object's type to be resolvable. Compile the project
  (*Debug → Compile*) so the VBE can resolve it.
- **Dependency installation failed** — run `python main.py` in a console to see the exact error.

## Credits

Built by [CoffeeBean001](https://github.com/CoffeeBean001). Inspired by the completion experience
of IntelliJ IDEA.
