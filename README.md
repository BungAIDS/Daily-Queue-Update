# debug/launcher branch

This branch exists **only** to collect debug reports from the desktop launcher
(`launcher.py`), kept separate from feature branches so the reports never clutter
code review.

- The launcher's **Publish Debug Report** button pushes a fresh
  `diagnostics/launcher_report.txt` here automatically (git plumbing — it does
  not touch your working checkout).
- To read the latest report:
  `git fetch origin debug/launcher && git show origin/debug/launcher:diagnostics/launcher_report.txt`

Each publish is a new commit, so history keeps every prior snapshot.
