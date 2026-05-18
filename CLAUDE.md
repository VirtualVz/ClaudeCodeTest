# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a small collection of standalone web experiments and scripts. No build system, package manager, or framework — everything runs directly in the browser or Python interpreter.

## Running Files

Open HTML files in the browser (always use the latest version number):
```
start burning-joint-v1.html
start tictactoe-v1.html
```

Run Python scripts:
```
python py1-v1.py
```

## File Versioning

All project files (HTML, Python) are versioned with a `-vN` suffix (e.g., `burning-joint-v1.html`, `burning-joint-v2.html`). **CLAUDE.md is the only file updated in place.**

**When modifying any project file:**
1. Identify the current highest version number for that file (scan the directory)
2. Write the updated content to a new file with the next version number — never overwrite the existing versioned file
3. Stage and commit only the new file

Example: if `burning-joint-v2.html` exists and a change is requested, create `burning-joint-v3.html`.

## GitHub Repository

**Repo:** https://github.com/VirtualVz/ClaudeCodeTest

**Required workflow — after every file change:**
1. Stage the new versioned file by name (never `git add -A` or `git add .`)
2. Commit with a short descriptive message and the co-author trailer:
   ```
   Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
   ```
3. Push to `origin master`:
   ```
   git push origin master
   ```

This must happen automatically at the end of every session in which any file is created or modified — do not wait for the user to ask.

## Prompt Logging

Every prompt the user sends must be appended to `prompts.txt` in the project root. This file is updated in place (never versioned). Format each entry as a numbered line:

```
N. <prompt text>
```

To find the next number, read the current contents of `prompts.txt` and increment the last entry's number. Do this at the start of every turn, before any other work.

## Architecture

### HTML files
Each HTML file is fully self-contained: inline `<style>` and `<script>` in a single file, no external dependencies. This pattern should be maintained for any new web experiments.

**`burning-joint-vN.html`** — Canvas 2D rendering of a joint that animates as a slider moves from 0–100%:
- Drawing pipeline: ambient glow → ash → paper → twist → cherry → filter → specular → smoke particles (layered in this order each frame)
- `ParticleSystem` class manages smoke; `flicker` object drives ember animation via spring-damped oscillator
- Key constants at the top of the script (`CW`, `CH`, `CY`, `JOINT_R`, `TWIST_START`, `BODY_START`, `FILTER_START`, etc.) define the joint geometry
- `cherryX = BODY_START + PAPER_LEN * burnPercent / 100` is the central position driving all visual state
- Slider is a plain HTML element (not canvas); `setPercent()` is the single update function

**`tictactoe-vN.html`** — Tic-tac-toe with 2-player and AI modes (minimax algorithm), score tracking, dark theme.

### Python
`py1-vN.py` is a scratch file.
