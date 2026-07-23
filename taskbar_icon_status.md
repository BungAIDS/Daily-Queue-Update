# Taskbar icon — status and debugging notes

**Status (2026-07-23): NOT WORKING on the user's machine.** The Explorer app
window's taskbar button still shows the generic blank-file icon. The title-bar
icon (top-left of the window) works and has always worked — that one comes
from the page favicon. This file records what was tried, the one run where the
taskbar fan icon actually appeared, and where to pick the problem up.

## The one confirmed working run

On 2026-07-23, the **first open after pulling commit `8cc966d`** ("Give the
Explorer app window its own taskbar icon"):

- The taskbar button showed the fan icon immediately — rendered small (it was
  the window's favicon-derived 16px icon being upscaled, not our .ico).
- After roughly **10 seconds** it flipped to the generic blank-file icon.

Interpretation at the time: the first phase was the window's own HICON (set by
Chromium from the favicon — proof the window itself was found and is
stampable). The flip happened when the shell lazily applied the stamped
`RelaunchIconResource` and failed to load it (then referenced on the network
share, and a PNG-only single-frame .ico). The taskbar button was **pinned
during/after the fan phase**, which froze the broken identity into
`%APPDATA%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\`.

Every later attempt showed the blank icon. Note: the later fixes were never
tested with the poisoned pin removed — **unpinning the button and doing one
fresh open was still outstanding** when work stopped.

## What is implemented (all still in place, all best-effort/harmless)

Published beside the page by `write_explorer` (order_explorer.py):

- `glq_taskbar.ps1` — run hidden by every open path; does all of the below.
  - Copies `GL Queue Explorer Fan.ico` and `Open GL Queue Explorer.vbs` to
    `%LOCALAPPDATA%\GL Queue Explorer\` and references ONLY those copies.
  - Compiles its C# interop once into `%LOCALAPPDATA%\GL Queue
    Explorer\GLQTaskbar.<rev>.dll` (recompiling on each open was the cause of
    "opens a lot slower"; bump `$dllRev` when the C# changes).
  - Writes a Start Menu shortcut `GL Queue Explorer.lnk` carrying
    `System.AppUserModel.ID = CBCInsider.GLQueueExplorer`, the local icon and
    the local opener — the canonical source the taskbar checks first for a
    group's icon; pins inherit it.
  - Finds the app window (exact title "GL Queue Explorer", class
    `Chrome_WidgetWin*` — a browser tab never matches) and stamps its property
    store via `SHGetPropertyStoreForWindow`: AppUserModelID +
    RelaunchCommand/DisplayName/IconResource.
- `Open GL Queue Explorer.vbs` — silent opener (Edge→Chrome→msedge fallback);
  bakes the absolute page path so its %LOCALAPPDATA% copy works; kicks the ps1.
- `explorer_icon.py` now emits a real multi-frame ICO (BMP frames at
  16/20/24/32/40/48/64 + 256px PNG). PNG-only ICOs are unloadable by the
  shell's small-size icon loaders — keep the BMP frames (test-pinned in
  `test_order_explorer.py::test_icon_assets_are_shell_loadable`).

## Next-session checklist (in order)

1. **Unpin** the blank taskbar button (right-click → Unpin). The pinned
   shortcut's icon overrides everything; nothing can work while it exists.
2. Fresh open, wait ~15 s. If the fan appears and survives: done; re-pin.
3. If still blank, inspect on the machine:
   - `dir "%LOCALAPPDATA%\GL Queue Explorer"` — expect the .ico, the .vbs and
     `GLQTaskbar.1.dll`. Missing dll ⇒ Add-Type/csc blocked (AV/policy);
     missing everything ⇒ the ps1 never ran.
   - Does `%APPDATA%\Microsoft\Windows\Start Menu\Programs\GL Queue
     Explorer.lnk` exist and show the fan icon in Explorer? If the .lnk itself
     shows the fan but the taskbar doesn't, the window stamp isn't matching
     (title/class mismatch or stamp failing).
   - Run the stamper VISIBLY to see errors: open a console in the page folder,
     `powershell -NoProfile -ExecutionPolicy Bypass -File glq_taskbar.ps1`
     with the Explorer window open (temporarily remove
     `$ErrorActionPreference = "SilentlyContinue"` for real errors).
4. Remaining hypotheses, roughly in likelihood order:
   - The poisoned pin was still present during all post-`8cc966d` tests.
   - Add-Type compile blocked on the machine → no stamp at all after the
     dll-cache change (the in-memory fallback also uses Add-Type).
   - Edge re-stamps its own AUMID on some event, reverting ours → would need
     re-stamping on an interval while the window lives, or the pywin32 route
     inside the launcher (`win32com.propsys`) with a retry loop.
   - Explorer's icon cache holding the blank result for the AUMID → delete
     `%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache*` + restart
     explorer.exe once.
5. To **revert/neutralize** the whole feature: delete the two helper files
   from the share (write_explorer will stop regenerating only if the code is
   also reverted), delete `%LOCALAPPDATA%\GL Queue Explorer\`, the Start Menu
   .lnk, and unpin. The window then behaves like any Edge app window again.

## What must not regress while iterating

- The title-bar/favicon fan icon (works).
- Open speed: never recompile C# per open; never launch scripts from the
  share on pinned/relaunch paths (both caused user-visible slowness).
- The `.ico` must keep its BMP frames (blank-page icon otherwise).
