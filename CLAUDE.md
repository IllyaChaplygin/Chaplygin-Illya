# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

This is a personal Git version-control coursework repository ("Homework_chaplygin_illya" per
README.md), not an application codebase. It contains no source code, no build system, no
dependency manifest, and no tests. Commit history (`git log --all --oneline`) shows it exists
to practice Git operations (commits, merges, file add/remove) as class assignments (e.g. "VCS
Homework№7").

Do not assume there is a project to build, lint, or test. If asked to "run the tests" or
"build the project," there is nothing to run — check with the user before inventing tooling
or scaffolding a new project structure.

## Repository layout and known quirks

- `README.md` — one-line project description.
- `.gitignore` — ignores `file2new` and `file3`.
- `file3` — tracked in git despite being listed in `.gitignore` (it was added before the
  ignore rule existed, so `git status` will not show it as untracked, but new copies would be
  ignored).
- `"VCS Homework№7"` — a literal filename containing `№` and a space; it is a homework
  artifact, not a script or config. Quote/escape it in shell commands.
- `Chaplygin-Illya/` — appears as a normal empty directory on disk, but is actually tracked as
  a **git submodule gitlink** (mode `160000`) pointing at commit `fc77b9e...` with **no
  `.gitmodules` file** present. This is a broken/dangling submodule reference (likely from an
  accidental `git add` inside a nested `.git` repo), not an intentional submodule setup. Don't
  try to "fix" this by running `git submodule update` — there's no `.gitmodules` entry to
  resolve it against. If new work needs to go in that directory, treat it as a plain directory
  and be aware that `git add` there may recreate a gitlink instead of tracking files normally
  (verify with `git status`/`git ls-files -s` after adding).

## Working in this repo

- Confirm with the user before adding new project scaffolding — this repo's purpose so far has
  been Git practice, not building software.
- When making commits, follow the existing history's style: short, plain commit messages.
