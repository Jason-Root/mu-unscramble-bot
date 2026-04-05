# MU Unscramble Bot

GitHub project:

- Repository: `https://github.com/Jason-Root/mu-unscramble-bot`

This project watches the middle of the screen for the yellow MU Online event text, extracts the current `Round`, `scrambled word`, and `Hint`, solves it, and can submit the answer back into the game chat.

## What it does

- Captures a configurable box around the center of the screen.
- Runs OCR over multiple yellow-text-friendly variants of the image.
- Parses lines like:
  - `Round 5: unscramble this word: eoarobng`
  - `Difficulty Level: 1 Hint: WHAT IS THE CAPITAL OF BOTSWANA?`
- Solves the clue with:
  - a local spreadsheet memory cache checked before any API call
  - an offline capital-city solver for country-capital hints
  - a fast local anagram dictionary for common one-word answers
  - an OpenAI-compatible solver only as a fallback for misses
- Types the answer into the active MU window using `pydirectinput` by default.
- Shows a small always-on-top overlay at the top of the screen with the current round, hint, and answer.
- Shows a live OCR feed in the overlay so you can see exactly what text is being read from the MU client.

## Quick start

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -e .
Copy-Item .env.example .env
```

Set your `OPENAI_API_KEY` in `.env` if you want random/general hint solving. Without an API key, the built-in offline solver only handles country-capital clues.

For OpenRouter free models, use:

```env
OPENAI_API_KEY=your_rotated_openrouter_key
OPENAI_MODEL=qwen/qwen3.6-plus:free
OPENAI_BASE_URL=https://openrouter.ai/api/v1
OPENAI_HTTP_REFERER=http://localhost
OPENAI_APP_TITLE=MU Unscramble Bot
```

You can also replace `qwen/qwen3.6-plus:free` with another specific free model that has a `:free` suffix.

## Main files

- `config.json`: capture area, OCR thresholds, and submit behavior
- `.env`: OpenAI credentials and optional model override

## Useful commands

Test against a screenshot:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot debug-image --image "c:\Users\Jason\Pictures\Screenshots\Screenshot 2026-04-04 092618.png"
```

Capture the live screen once and save debug images:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot debug-screen
```

List the matching client windows:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot list-windows
```

Run the live bot without typing anything:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot run --dry-run
```

Test the configured API/model by itself:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot test-api
```

Saved question/answer memory:

- Solved questions are written automatically to `data/question_memory.csv`
- That CSV is checked before any API request
- You can open it directly in Excel
- You can also seed fast local answers manually in `data/local_dictionary.txt`

Community/shared answer sheet:

- The bot now uses GitHub as the shared community answer sheet source.
- Every client still keeps its normal local CSV, but it can also pull from one GitHub CSV and push new answers back automatically.
- Reads hot-reload on a timer while the bot is running, so new answers from other PCs show up without restarting.
- To enable it, set these in `config.json`:

```json
"github_answer_sheet_enabled": true,
"github_answer_sheet_repository": "Jason-Root/mu-unscramble-bot",
"github_answer_sheet_branch": "main",
"github_answer_sheet_path": "data/question_memory.csv",
"github_answer_sheet_sync_interval_seconds": 30.0
```

- Then put a GitHub fine-grained or classic token with repo content write access in `.env`:

```env
GITHUB_TOKEN=your_github_token_here
```

- If you leave `GITHUB_TOKEN` blank, the bot can still read from a public GitHub repo but it will not be able to push newly learned answers back.

Run the live bot with auto-submit enabled:

```powershell
.\.venv\Scripts\python -m mu_unscramble_bot run
```

Double-click launcher:

- You can start the bot by opening [Start MU Unscramble Bot.cmd](d:/project/Start%20MU%20Unscramble%20Bot.cmd)
- It gives you a small menu for `dry-run`, `live`, or `test-api`
- `dry-run` and `live` now show the visible Divine MU clients and let you choose the window index at startup

## Config notes

- `center_offset_y` defaults to `220` because the yellow message usually sits a bit below true screen center.
- `unsolved_retry_seconds` keeps retrying the same puzzle quickly if OCR or solving misses once.
- `open_chat_key` and `submit_key` both default to `enter`.
- `submit_backend` defaults to `directinput` because many games handle it better than normal GUI typing.
- `focus_window_title_contains` plus `require_window_match=true` helps the bot avoid typing into the wrong window.
- `capture_source="window"` makes OCR use a specific client window instead of the whole monitor.
- `target_window_title_contains`, `target_window_exact_title`, and `target_window_index` let you choose which MU client to read.
- `show_overlay=true` opens a small topmost status window near the top-center of the screen.
- `show_live_ocr_overlay=true` shows the current OCR lines inside the overlay.
- `OPENAI_BASE_URL` lets you point the solver at OpenRouter or another OpenAI-compatible endpoint.
- `test_api_on_startup=true` sends a tiny prompt when the bot starts so you can see whether the configured model is reachable.
- `question_memory_path` points to the spreadsheet-style CSV file used for saved question/answer recall.
- `github_answer_sheet_*` settings let every bot sync that local sheet with one shared GitHub CSV.
- `local_dictionary_enabled`, `local_dictionary_max_words`, and `local_dictionary_path` control the offline anagram solver.
- `online_solver_timeout_seconds` caps the slower API fallback so the bot does not wait too long on a miss.

## Practical notes

- Start with `debug-screen` and `run --dry-run` before turning on live submission.
- If OCR misses the text, increase `capture_width` / `capture_height` or tune the yellow HSV range in `config.json`.
- If the game ignores typed keys, try running the game and the bot with the same privilege level.
