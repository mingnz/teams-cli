"""Token management and Playwright-based login flow."""

import json
import time

from .config import (
    DATA_DIR,
    REGION_KEY_PATTERN,
    TEAMS_URL,
    TOKEN_AUDIENCES,
    TOKENS_FILE,
)

# Browser profile directory for persistent sessions
BROWSER_PROFILE_DIR = DATA_DIR / "browser-profile"


def save_tokens(tokens: dict) -> None:
    """Save tokens to disk."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    TOKENS_FILE.write_text(json.dumps(tokens, indent=2))


def load_tokens() -> dict | None:
    """Load tokens from disk, or None if not found."""
    if not TOKENS_FILE.exists():
        return None
    return json.loads(TOKENS_FILE.read_text())


def is_expired(token_entry: dict) -> bool:
    """Check if a token entry has expired."""
    expires_on = token_entry.get("expires_on", 0)
    return time.time() > float(expires_on)


def _try_refresh() -> bool:
    """Silently refresh tokens by launching a headless browser.

    Uses the persistent browser profile (session cookies) to navigate
    to Teams, letting MSAL silently acquire new tokens, then re-extracts
    them from localStorage.

    Returns True if tokens were refreshed successfully.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False

    try:
        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                str(BROWSER_PROFILE_DIR),
                headless=True,
            )
            page = context.new_page()
            page.goto(TEAMS_URL)

            # Wait for MSAL to silently acquire tokens (up to 30s)
            deadline = time.time() + 30
            while time.time() < deadline:
                try:
                    found = page.evaluate("""() => {
                        for (let i = 0; i < localStorage.length; i++) {
                            const key = localStorage.key(i);
                            if (key.includes('ic3.teams.office.com') && key.includes('accesstoken')) {
                                return true;
                            }
                        }
                        return false;
                    }""")
                    if found:
                        break
                except Exception:
                    pass
                page.wait_for_timeout(1000)
            else:
                context.close()
                return False

            page.wait_for_timeout(2000)

            # Re-extract tokens
            tokens = load_tokens() or {}
            for name, audience in TOKEN_AUDIENCES.items():
                result = page.evaluate(f"""() => {{
                    for (let i = 0; i < localStorage.length; i++) {{
                        const key = localStorage.key(i);
                        if (key.includes('{audience}') && key.includes('accesstoken')) {{
                            const data = JSON.parse(localStorage.getItem(key));
                            return {{ secret: data.secret, expires_on: data.expiresOn }};
                        }}
                    }}
                    return null;
                }}""")
                if result:
                    tokens[name] = result

            context.close()

        save_tokens(tokens)
        return True
    except Exception:
        return False


def get_token(name: str) -> str | None:
    """Get a specific token by name, refreshing automatically if expired."""
    tokens = load_tokens()
    if not tokens:
        return None
    entry = tokens.get(name)
    if not entry:
        return None
    if is_expired(entry):
        if _try_refresh():
            tokens = load_tokens()
            entry = tokens.get(name, {}) if tokens else {}
            if not is_expired(entry):
                return entry.get("secret")
        return None
    return entry["secret"]


def get_region() -> str | None:
    """Get the stored region code (e.g. 'au')."""
    tokens = load_tokens()
    if not tokens:
        return None
    return tokens.get("region")


def login() -> dict:
    """Launch a browser for Teams login and extract tokens.

    Opens a Chromium browser to Teams, waits for the user to complete
    login (including MFA), then extracts auth tokens from localStorage.

    Returns:
        dict with token entries and region.

    Raises:
        TimeoutError: If login doesn't complete within 5 minutes.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Playwright is required for login. Run:\n"
            "  uv run playwright install chromium"
        )

    tokens = {}
    BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(BROWSER_PROFILE_DIR),
            headless=False,
        )
        page = context.new_page()

        print("Opening Teams login page...")
        page.goto(TEAMS_URL)
        print("Please sign in to Microsoft Teams in the browser window.")
        print("Waiting for login to complete (timeout: 5 minutes)...")

        # Poll localStorage for the IC3 token
        deadline = time.time() + 300  # 5 minute timeout
        while time.time() < deadline:
            try:
                found = page.evaluate("""() => {
                    for (let i = 0; i < localStorage.length; i++) {
                        const key = localStorage.key(i);
                        if (key.includes('ic3.teams.office.com') && key.includes('accesstoken')) {
                            return true;
                        }
                    }
                    return false;
                }""")
                if found:
                    break
            except Exception:
                pass
            page.wait_for_timeout(2000)
        else:
            context.close()
            raise TimeoutError("Login timed out after 5 minutes.")

        # Brief pause to let all tokens populate
        page.wait_for_timeout(3000)
        print("Login detected! Extracting tokens...")

        # Extract all tokens
        for name, audience in TOKEN_AUDIENCES.items():
            result = page.evaluate(f"""() => {{
                for (let i = 0; i < localStorage.length; i++) {{
                    const key = localStorage.key(i);
                    if (key.includes('{audience}') && key.includes('accesstoken')) {{
                        const data = JSON.parse(localStorage.getItem(key));
                        return {{ secret: data.secret, expires_on: data.expiresOn }};
                    }}
                }}
                return null;
            }}""")
            if result:
                tokens[name] = result

        # Extract region
        region_data = page.evaluate(f"""() => {{
            for (let i = 0; i < localStorage.length; i++) {{
                const key = localStorage.key(i);
                if (key.includes('{REGION_KEY_PATTERN}')) {{
                    const data = JSON.parse(localStorage.getItem(key));
                    const gtm = typeof data.item === 'string' ? JSON.parse(data.item) : data.item;
                    // Extract region from AMS URL like "https://au-prod.asyncgw.teams.microsoft.com"
                    const ams = gtm.ams || '';
                    const match = ams.match(/https:\\/\\/([a-z]+)-prod/);
                    return match ? match[1] : null;
                }}
            }}
            return null;
        }}""")
        tokens["region"] = region_data or "amer"

        context.close()

    save_tokens(tokens)
    return tokens
