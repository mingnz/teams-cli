"""Constants and configuration."""

from pathlib import Path

# Where tokens and cached data are stored
DATA_DIR = Path.home() / ".teams-cli"
TOKENS_FILE = DATA_DIR / "tokens.json"
LAST_CHATS_FILE = DATA_DIR / "last_chats.json"

# Teams web client ID (first-party Microsoft app)
CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"

# localStorage key patterns for token extraction
TOKEN_AUDIENCES = {
    "ic3": "ic3.teams.office.com",
    "search": "outlook.office.com/search",
    "presence": "presence.teams.microsoft.com",
}

# Region discovery key pattern
REGION_KEY_PATTERN = "DISCOVER-REGION-GTM"

# Required headers for chatsvc API
CLIENT_INFO = (
    "os=mac; osVer=10.15.7; proc=x86; lcid=en-us; "
    "deviceType=1; country=us; clientName=skypeteams; "
    "clientVer=1415/26031223020; utcOffset=+12:00; timezone=Pacific/Auckland"
)

# Default page sizes
DEFAULT_CHAT_PAGE_SIZE = 30
DEFAULT_MESSAGE_PAGE_SIZE = 20
DEFAULT_SEARCH_SIZE = 25

# Teams login URL
TEAMS_URL = "https://teams.cloud.microsoft/"
