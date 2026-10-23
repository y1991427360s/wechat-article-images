# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Two standalone Python scripts that download images from a WeChat public-account (公众号) article and package them into Word (`.docx`) and PowerPoint (`.pptx`) files. There is no package, no build, and no test suite — each script is run directly with `python`.

- `weixin-gui.py` — the primary, current implementation. A Tkinter desktop app styled after the Oracle-Sync desktop-pet menu (white rounded cards, `#E4E6EB` borders, Segoe MDL2 icons) that runs the download + export pipeline on a worker thread and streams progress back to the UI via a `queue.Queue`.
- `weixin-word-ppt.py` — the original CLI prototype. Simpler, synchronous, prompts for URL/prefix on stdin, and writes output to `~/Desktop/weixin_images/<timestamp>/`. Kept for reference; prefer editing the GUI version for new behavior.

## Running

```bash
python weixin-gui.py        # GUI
python weixin-word-ppt.py   # CLI prototype
```

Runtime dependencies (listed in `requirements.txt`): `requests`, `beautifulsoup4`, `lxml`, `python-docx`, `python-pptx`, `Pillow`. Standard library covers Tkinter, threading, queue.

## Architecture notes

The GUI script has three layers:

1. **Pipeline (pure module-level functions)**: `fetch_article` → `download_images` → `generate_word_document` / `generate_ppt_presentation`, orchestrated by `run_export`, the worker-thread entry point. `run_export` receives every option as a plain argument (never reads Tk variables) and **always** posts a final `done` event (`outcome` = success / cancelled / failed, `folder`, `elapsed`).
2. **Widgets**: `RoundedBox` (canvas whose background is an anti-aliased rounded rect rendered by Pillow via the cached `render_round_rect`, children go in `.inner`), `IconButton`, `InputField`, `OptionTile`, `ProgressBar`. All pixel sizes go through `px()` (scaled by `UI_SCALE`, set from DPI in `main()` after `SetProcessDpiAwareness`).
3. **`MaterialApp`**: owns the Tk root, validates inputs inline (red outline + state label, no message boxes), persists settings to `%APPDATA%\微信公众号素材整理\config.json`.

Conventions when editing:

- **Status flows through `status_queue`, never direct Tk calls from the worker** (including `root.after`). Event types: `log`, `stage`, `progress` (`value` 0–1), `done`. `MaterialApp._poll_status_queue` drains every 100 ms. Add new feedback as a new event type.
- Cancellation flows via `cancel_event` (`threading.Event`); retry backoff uses `cancel_event.wait(...)`, and on cancel the executor is shut down without waiting.
- `requests` exceptions subclass `OSError` — catch `RequestException` before `OSError`. 4xx (except 429) is not retried.
- Image discovery is scoped to `#js_content` and tries `data-src`, `data-actualsrc`, `data-original`, `data-lazy-src`, `src` (WeChat lazy-loads). If none found, `extract_script_image_urls` regex-scans the page for `mmbiz.qpic.cn` URLs (picture-message posts). The mobile-Safari `USER_AGENT` is load-bearing.
- `save_image_bytes` sniffs the format with Pillow (not Content-Type), skips tiny placeholder images (`SkippedImage`), keeps JPEG/PNG/GIF/BMP as-is and converts everything else (webp…) to PNG so Word/PPT can embed it.
- Output folder is `<name>_<timestamp>` where name = prefix or article title, created only after the article page has been fetched successfully.
- Word: A4 with 2.5 cm margins; `word_width_for` never upscales and also caps height for long images. PPT uses a fixed 25.4 × 19.05 cm slide with 1 cm margin (`fit_image_to_slide`) and strips the template's default first slide by mutating `_sldIdLst` — preserve that block.
- Segoe MDL2 folder glyphs `E8B7`/`E838` render as vertical shapes; use `F12B` (folder) and `ED25` (open folder).
