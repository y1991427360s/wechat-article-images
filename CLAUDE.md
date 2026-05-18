# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Two standalone Python scripts that download images from a WeChat public-account (公众号) article and package them into Word (`.docx`) and PowerPoint (`.pptx`) files. There is no package, no build, and no test suite — each script is run directly with `python`.

- `weixin-gui.py` — the primary, current implementation. A Tkinter desktop app (Apple-style theming) that runs the download + export pipeline on a worker thread and streams progress back to the UI via a `queue.Queue`.
- `weixin-word-ppt.py` — the original CLI prototype. Simpler, synchronous, prompts for URL/prefix on stdin, and writes output to `~/Desktop/weixin_images/<timestamp>/`. Kept for reference; prefer editing the GUI version for new behavior.

## Running

```bash
python weixin-gui.py        # GUI
python weixin-word-ppt.py   # CLI prototype
```

Runtime dependencies (no `requirements.txt` in repo): `requests`, `beautifulsoup4`, `lxml`, `python-docx`, `python-pptx`, `Pillow`. Standard library covers Tkinter, threading, queue.

## Architecture notes

The GUI script is organized as **pure module-level functions** (`extract_image_urls`, `download_single_image`, `download_images_from_url`, `generate_word_document`, `generate_ppt_presentation`, helpers) plus a single `AppleStyleApp` class that owns the Tk root and orchestrates a background `threading.Thread` worker. Two conventions matter when editing:

1. **Status flows through `status_queue` (a `queue.Queue`), never direct UI calls.** Worker code emits three event types via helpers: `log_status` (`log` — appended to the status text area with color level), `set_stage` (`stage` — updates the right-aligned state label), and `post_event(..., "progress", ...)`. `AppleStyleApp._poll_status_queue` drains the queue every 120 ms on the Tk main loop. Do not touch Tk widgets from inside `_worker` or any of the download/export functions — add new feedback by introducing a new event type and a branch in `_poll_status_queue`.
   - Cancellation flows the other way via `self.cancel_event` (a `threading.Event`) passed into `download_images_from_url` and `download_single_image`. The retry backoff uses `cancel_event.wait(...)` so cancel is responsive between attempts; an in-flight `requests.get` is not interrupted but its result is discarded once the outer loop sees the flag.
2. **Image discovery tries multiple attributes in order** (`data-src`, `data-actualsrc`, `data-original`, `data-lazy-src`, `src`) because WeChat lazy-loads. The mobile-Safari `USER_AGENT` is also load-bearing for fetching the article HTML.

Downloads run through a `ThreadPoolExecutor` (`MAX_DOWNLOAD_WORKERS=5`) with per-image retry (`MAX_DOWNLOAD_RETRIES=3`) and linear backoff (`RETRY_BACKOFF_SECONDS * attempt`). File extensions are inferred from `Content-Type` first, falling back to Pillow's format sniff (`choose_extension`), normalizing `jpeg` → `jpg`. Each run creates a fresh timestamped folder (`微信素材_YYYYMMDD_HHMMSS`).

PPT layout uses a fixed 25.4 × 19.05 cm slide with a 1 cm margin (`fit_image_to_slide`); Word inserts pictures via `word_width_for(...)`, which caps width at `WORD_MAX_WIDTH_CM` (16 cm) but does not upscale small images — natural width is computed from pixels at `WORD_IMAGE_DPI` (96). The PPT generator also strips the default first slide from the `Presentation()` template by mutating `_sldIdLst` directly — preserve that block if rewriting, or the output will have a blank opening slide.
