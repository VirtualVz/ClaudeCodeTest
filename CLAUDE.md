# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a small collection of standalone web experiments and scripts. No build system, package manager, or framework — everything runs directly in the browser or Python interpreter.

## Running Files

Open HTML files in the browser:
```
start burning-joint.html
start tictactoe.html
```

Run Python scripts:
```
python py1.py
```

## Architecture

### HTML files
Each HTML file is fully self-contained: inline `<style>` and `<script>` in a single file, no external dependencies. This pattern should be maintained for any new web experiments.

**`burning-joint.html`** — Canvas 2D rendering of a joint that animates as a slider moves from 0–100%:
- Drawing pipeline: ambient glow → ash → paper → twist → cherry → filter → specular → smoke particles (layered in this order each frame)
- `ParticleSystem` class manages smoke; `flicker` object drives ember animation via spring-damped oscillator
- Key constants at the top of the script (`CW`, `CH`, `CY`, `JOINT_R`, `TWIST_START`, `BODY_START`, `FILTER_START`, etc.) define the joint geometry
- `cherryX = BODY_START + PAPER_LEN * burnPercent / 100` is the central position driving all visual state
- Slider is a plain HTML element (not canvas); `setPercent()` is the single update function

**`tictactoe.html`** — Tic-tac-toe with 2-player and AI modes (minimax algorithm), score tracking, dark theme.

### Python
`py1.py` is a scratch file.

## GitHub Repository

**Repo:** https://github.com/VirtualVz/ClaudeCodeTest

**Required workflow — after every file change:**
1. Stage the modified file(s) by name (never `git add -A` or `git add .`)
2. Commit with a short descriptive message and the co-author trailer:
   ```
   Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
   ```
3. Push to `origin master`:
   ```
   git push origin master
   ```

This must happen automatically at the end of every session in which any file is created or modified — do not wait for the user to ask.
