import { homedir } from "node:os";
import { join } from "node:path";

export const DATA_DIR = join(homedir(), ".teams-cli");
export const TOKENS_FILE = join(DATA_DIR, "tokens.json");
export const LAST_CHATS_FILE = join(DATA_DIR, "last_chats.json");
export const BROWSER_PROFILE_DIR = join(DATA_DIR, "browser-profile");

export const CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";

export const TOKEN_AUDIENCES: Record<string, string> = {
  ic3: "ic3.teams.office.com",
  search: "outlook.office.com/search",
  presence: "presence.teams.microsoft.com",
};

export const REGION_KEY_PATTERN = "DISCOVER-REGION-GTM";

export const CLIENT_INFO =
  "os=mac; osVer=10.15.7; proc=x86; lcid=en-us; " +
  "deviceType=1; country=us; clientName=skypeteams; " +
  "clientVer=1415/26031223020; utcOffset=+12:00; timezone=Pacific/Auckland";

export const DEFAULT_CHAT_PAGE_SIZE = 30;
export const DEFAULT_MESSAGE_PAGE_SIZE = 20;
export const DEFAULT_SEARCH_SIZE = 25;

export const TEAMS_URL = "https://teams.cloud.microsoft/";
