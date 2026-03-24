# copilot-cli

A command-line tool to send prompts to Microsoft 365 Copilot and stream the response back to your terminal, using browser automation via Playwright.

## How it works

Launches a persistent Chromium browser pointed at `m365.cloud.microsoft/chat`, finds the chat input, submits your prompt, and streams the response text to stdout as it arrives. The browser profile is saved locally so you stay logged in between runs.

## Requirements

- Python 3.8+
- A Microsoft 365 account with Copilot access

## Setup

```bash
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
# Prompt as argument
python copilot.py "Summarize the key points of agile development"

# Prompt from stdin
echo "What is the capital of France?" | python copilot.py

# Prompt from a file
python copilot.py -f prompt.txt

# Debug mode — pauses before submitting so you can inspect the browser
python copilot.py --debug "your question"
```

## First run

On first launch the browser will open and may prompt you to log in to your Microsoft 365 account. Once logged in, the session is saved to `./browser-profile/` and reused on subsequent runs.

## Notes

- The browser runs in headed mode (not headless) — this is required for M365 authentication to work correctly.
- Response streaming polls for text changes and stops after ~2.5 seconds of no new output.
- If the chat input can't be found after navigation, the browser stays open for 120 seconds so you can log in manually before re-running.
