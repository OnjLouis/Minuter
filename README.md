# Minuter

Minuter provides optional time cues while you work: second ticks/beeps, minute chime, end-of-minute beeps, and optional spoken seconds.

## Quick Start

1. Press `NVDA+Windows+Backspace` to open Minuter settings.
2. Enable one or more cue options.
3. Close settings and confirm cues are audible.

## Keyboard Shortcut

| Shortcut | Action |
| --- | --- |
| `NVDA+Windows+Backspace` | Open Minuter settings |

## Main Options

- Tick each second
- Low beep each second
- Ding each minute
- Beep at end of minute
- Speak seconds

## Sound Customization

- Minuter uses files in its `sounds` folder (for example `tick.wav`, `minute.wav`).
- Replacing those files in the add-on package changes cue sounds.

## Full Documentation

- Full help: [`source/doc/en/readme.html`](source/doc/en/readme.html)

## Source Code

- Extracted source for this build: [`source/`](source/)
- Main plugin: [`source/globalPlugins/minuter.py`](source/globalPlugins/minuter.py)

## Install

1. Download the `.nvda-addon` file from Releases.
2. In NVDA, open Add-on Manager and choose Install.
3. Select the file and restart NVDA when prompted.
