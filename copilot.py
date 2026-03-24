#!/usr/bin/env python3
"""
copilot-cli: Send prompts to Microsoft 365 Copilot via browser automation.

Usage:
    python copilot.py "your question here"
    echo "question" | python copilot.py
    python copilot.py -f prompt.txt
    python copilot.py --debug "question"   # pause before submit so you can inspect
"""

import sys
import time
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

COPILOT_URL = "https://m365.cloud.microsoft/chat"
PROFILE_DIR = str(Path(__file__).parent / "browser-profile")

# How long to wait for the page/input to be ready (ms)
PAGE_READY_TIMEOUT = 60_000
# How many seconds of no text change = response is done
DONE_IDLE_SECS = 2.5
# Poll interval while streaming (seconds)
POLL_INTERVAL = 0.1


def find_input(page):
    """Find the chat input element using multiple selector strategies."""
    candidates = [
        "textarea[aria-label*='message' i]",
        "textarea[placeholder*='message' i]",
        "div[contenteditable='true'][aria-label*='message' i]",
        "div[contenteditable='true'][aria-label*='Copilot' i]",
        "div[contenteditable='true'][aria-placeholder*='message' i]",
        "textarea",
        "div[contenteditable='true']",
    ]
    for sel in candidates:
        try:
            el = page.wait_for_selector(sel, timeout=2000, state="visible")
            if el:
                return el, sel
        except PlaywrightTimeout:
            continue
    return None, None


def submit_prompt(page, input_el, prompt: str):
    """Type the prompt and submit it."""
    input_el.click()

    # Use fill for textarea, type for contenteditable
    tag = input_el.evaluate("el => el.tagName.toLowerCase()")
    if tag == "textarea":
        input_el.fill(prompt)
    else:
        # contenteditable — select all then type
        input_el.press("Control+a")
        input_el.type(prompt, delay=0)

    # Try send button first, fall back to Enter
    send_selectors = [
        "button[aria-label*='send' i]",
        "button[data-testid*='send' i]",
        "button[title*='send' i]",
        "button[aria-label*='Submit' i]",
    ]
    sent = False
    for sel in send_selectors:
        try:
            btn = page.wait_for_selector(sel, timeout=1500, state="visible")
            if btn and btn.is_enabled():
                btn.click()
                sent = True
                break
        except PlaywrightTimeout:
            continue

    if not sent:
        input_el.press("Enter")


def get_last_response_text(page) -> str:
    """Extract text from the last assistant message in the conversation."""
    # Try several selector patterns for the response container
    selectors = [
        # M365 Copilot common patterns
        "[data-testid='assistant-message']",
        "[class*='assistant'][class*='message']",
        "[class*='bot'][class*='message']",
        "[class*='copilot'][class*='message']",
        # Generic chat patterns
        "[role='listitem']",
        ".message",
    ]
    for sel in selectors:
        try:
            els = page.query_selector_all(sel)
            if els:
                # Last element is the most recent response
                text = els[-1].inner_text()
                if text.strip():
                    return text.strip()
        except Exception:
            continue
    return ""


def is_generating(page) -> bool:
    """Return True if Copilot is still generating a response."""
    stop_selectors = [
        "button[aria-label*='Stop' i]",
        "button[aria-label*='stop generating' i]",
        "button[title*='Stop' i]",
        "[data-testid*='stop' i]",
    ]
    for sel in stop_selectors:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                return True
        except Exception:
            continue
    return False


def stream_response(page) -> str:
    """Poll the last response and stream new text to stdout. Returns full text."""
    print("", file=sys.stderr)  # blank line before response

    # Wait for generation to start (stop button appears or text changes)
    print("Waiting for response...", file=sys.stderr, end="\r")
    deadline = time.time() + 30
    while time.time() < deadline:
        if is_generating(page) or get_last_response_text(page):
            break
        time.sleep(POLL_INTERVAL)
    else:
        print("\nTimed out waiting for response.", file=sys.stderr)
        return ""

    # Stream
    last_text = ""
    last_change = time.time()

    while True:
        current_text = get_last_response_text(page)

        if current_text != last_text:
            new_part = current_text[len(last_text):]
            print(new_part, end="", flush=True)
            last_text = current_text
            last_change = time.time()

        still_going = is_generating(page)
        idle_secs = time.time() - last_change

        if not still_going and idle_secs >= DONE_IDLE_SECS:
            break

        # Safety cap — give up after 3 min
        if idle_secs > 180:
            break

        time.sleep(POLL_INTERVAL)

    print()  # final newline
    return last_text


def run(prompt: str, debug: bool = False):
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            no_viewport=True,
            args=["--start-maximized"],
        )

        page = context.pages[0] if context.pages else context.new_page()

        # Navigate if not already on the right page
        if "m365.cloud.microsoft" not in page.url:
            print("Opening Copilot...", file=sys.stderr)
            page.goto(COPILOT_URL, wait_until="domcontentloaded")

        # Wait for the input box — this is where the user may need to log in
        print("Waiting for chat input (log in if prompted)...", file=sys.stderr)
        input_el, matched_sel = find_input(page)

        if not input_el:
            print(
                "\nCould not find the chat input box.\n"
                "The browser is open — log in if needed, then re-run.\n"
                "Keeping browser open for 120s.",
                file=sys.stderr,
            )
            time.sleep(120)
            context.close()
            return

        print(f"Input found ({matched_sel})", file=sys.stderr)

        if debug:
            input(
                "DEBUG: paused before submit — inspect the browser, then press Enter here... "
            )

        submit_prompt(page, input_el, prompt)
        stream_response(page)

        context.close()


def main():
    parser = argparse.ArgumentParser(
        description="Send a prompt to Microsoft 365 Copilot and stream the response."
    )
    parser.add_argument("prompt", nargs="?", help="Prompt text")
    parser.add_argument("-f", "--file", help="Read prompt from a file")
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Pause before submitting so you can inspect the browser",
    )
    args = parser.parse_args()

    if args.file:
        prompt = Path(args.file).read_text(encoding="utf-8").strip()
    elif args.prompt:
        prompt = args.prompt
    elif not sys.stdin.isatty():
        prompt = sys.stdin.read().strip()
    else:
        parser.print_help()
        sys.exit(1)

    if not prompt:
        print("Error: empty prompt", file=sys.stderr)
        sys.exit(1)

    run(prompt, debug=args.debug)


if __name__ == "__main__":
    main()
