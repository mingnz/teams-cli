import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { Command } from "commander";
import chalk from "chalk";
import Table from "cli-table3";
import { getMyMri, login, logout } from "./auth.js";
import { getChatClient, getSearchClient } from "./client.js";
import { DATA_DIR, LAST_CHATS_FILE } from "./config.js";
import {
  type FormattedChat,
  type FormattedMessage,
  formatChatList,
  formatMember,
  formatMessage,
  formatPerson,
  getConversationDisplayName,
  stripHtml,
} from "./formatting.js";
import {
  createDmThread,
  getActivity,
  getMessages,
  getThreadMembers,
  listConversations,
  pollConversations,
  pollMessages,
  searchMessages,
  searchPeople,
  sendMessage,
} from "./api.js";

function saveChatIndex(chats: FormattedChat[]): void {
  mkdirSync(DATA_DIR, { recursive: true });
  writeFileSync(LAST_CHATS_FILE, JSON.stringify(chats, null, 2));
}

function loadChatIndex(): FormattedChat[] {
  if (!existsSync(LAST_CHATS_FILE)) return [];
  return JSON.parse(readFileSync(LAST_CHATS_FILE, "utf-8"));
}

function resolveConversationId(idOrShort: string): string {
  if (idOrShort.includes(":")) return idOrShort;
  const chats = loadChatIndex();
  for (const c of chats) {
    if (c.short_id === idOrShort) return c.id;
  }
  console.error(`Short ID '${idOrShort}' not found. Run \`teams chats\` first.`);
  process.exit(1);
}

function printMessages(msgs: FormattedMessage[]): void {
  for (const m of msgs) {
    console.log(`${chalk.green(m.time)} ${chalk.bold(m.sender)}`);
    console.log(`  ${m.body}`);
    console.log();
  }
}

export function createProgram(): Command {
  const program = new Command();
  program
    .name("teams")
    .description("CLI for Microsoft Teams - list chats, read and send messages, search, and more.")
    .version("0.1.0");

  // login
  program
    .command("login")
    .description("Launch a browser to sign in to Microsoft Teams and store auth tokens.")
    .action(async () => {
      try {
        const tokens = await login();
        const names = Object.keys(tokens).filter((k) => k !== "region");
        console.log(chalk.green("Logged in.") + ` Tokens saved: ${names.join(", ")}`);
      } catch (e) {
        console.error((e as Error).message);
        process.exit(1);
      }
    });

  // logout
  program
    .command("logout")
    .description("Remove stored tokens and browser session data.")
    .action(() => {
      logout();
      console.log("Logged out. Tokens and session data removed.");
    });

  // chats
  program
    .command("chats")
    .description("List recent conversations.")
    .option("-n, --limit <number>", "Number of chats to list", "30")
    .action(async (opts) => {
      const client = await getChatClient();
      const convos = await listConversations(client, Number(opts.limit));
      const formatted = formatChatList(convos);
      saveChatIndex(formatted);

      const table = new Table({
        head: ["ID", "Name", "Type", "Last Message", "Time"],
        style: { head: ["cyan"] },
      });
      for (const c of formatted) {
        table.push([c.short_id, c.name, c.type, c.preview.slice(0, 60), c.time]);
      }
      console.log(table.toString());
    });

  // messages
  program
    .command("messages <chat>")
    .description("Read messages from a conversation.")
    .option("-n, --limit <number>", "Number of messages to fetch", "20")
    .action(async (chat: string, opts) => {
      const convId = resolveConversationId(chat);
      const client = await getChatClient();
      const raw = await getMessages(client, convId, Number(opts.limit));
      const msgs = raw.map((r) => formatMessage(r)).filter((m): m is FormattedMessage => m !== null);

      if (msgs.length === 0) {
        console.log("No displayable messages.");
        return;
      }
      printMessages(msgs.reverse());
    });

  // send
  program
    .command("send <chat> <message>")
    .description("Send a message to a conversation.")
    .action(async (chat: string, message: string) => {
      const convId = resolveConversationId(chat);
      const client = await getChatClient();
      const result = await sendMessage(client, convId, message);
      console.log(chalk.green("Sent.") + ` Arrival: ${result.OriginalArrivalTime ?? "ok"}`);
    });

  // search
  program
    .command("search <query>")
    .description("Search messages across all conversations.")
    .option("-n, --limit <number>", "Max results", "25")
    .action(async (query: string, opts) => {
      const client = await getSearchClient();
      const data = await searchMessages(client, query, Number(opts.limit));

      const results: { sender: string; preview: string }[] = [];
      for (const entitySet of (data.entitySets as Record<string, unknown>[]) ?? []) {
        for (const resultSet of (entitySet.resultSets as Record<string, unknown>[]) ?? []) {
          for (const hit of (resultSet.results as Record<string, unknown>[]) ?? []) {
            const body = stripHtml((hit.preview as string) ?? "");
            const extensions = (hit.extensions as Record<string, string>) ?? {};
            const sender =
              extensions.Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String ?? "";
            results.push({ sender, preview: body });
          }
        }
      }

      if (results.length === 0) {
        console.log("No results found.");
        return;
      }
      for (let i = 0; i < results.length; i++) {
        const r = results[i];
        console.log(`${chalk.cyan(`${i + 1}.`)} ${chalk.bold(r.sender || "Unknown")}`);
        console.log(`   ${r.preview.slice(0, 120)}`);
        console.log();
      }
    });

  // activity
  program
    .command("activity")
    .description("Show the activity feed (notifications, mentions, or call logs).")
    .option("-f, --feed <feed>", "Feed: notifications, mentions, or calllogs", "notifications")
    .option("-n, --limit <number>", "Number of items", "20")
    .action(async (opts) => {
      const client = await getChatClient();
      const raw = await getActivity(client, opts.feed, Number(opts.limit));
      const msgs = raw.map((r) => formatMessage(r)).filter((m): m is FormattedMessage => m !== null);

      if (msgs.length === 0) {
        console.log("No activity items.");
        return;
      }
      printMessages(msgs.reverse());
    });

  // find
  program
    .command("find <query>")
    .description("Search for people by name or email.")
    .option("-n, --limit <number>", "Max results", "10")
    .action(async (query: string, opts) => {
      const client = await getSearchClient();
      const results = await searchPeople(client, query, Number(opts.limit));

      if (results.length === 0) {
        console.log("No people found.");
        return;
      }

      const table = new Table({
        head: ["#", "Name", "Email", "Title", "MRI"],
        style: { head: ["cyan"] },
      });
      for (let i = 0; i < results.length; i++) {
        const p = formatPerson(results[i]);
        table.push([String(i + 1), p.name, p.email, p.title, p.mri]);
      }
      console.log(table.toString());
    });

  // dm
  program
    .command("dm <user> <message>")
    .description("Send a direct message to a user by name, email, or MRI.")
    .action(async (user: string, message: string) => {
      const myMri = getMyMri();
      if (!myMri) {
        console.error("Could not determine your user ID. Run `teams login`.");
        process.exit(1);
      }

      let theirMri: string;

      if (user.startsWith("8:orgid:")) {
        theirMri = user;
      } else {
        const searchClient = await getSearchClient();
        const results = await searchPeople(searchClient, user, 5);
        if (results.length === 0) {
          console.error(`No user found matching '${user}'.`);
          process.exit(1);
        }

        if (results.length > 1) {
          console.log(chalk.yellow(`Multiple matches for '${user}':`));
          for (let i = 0; i < results.length; i++) {
            const emails = (results[i].EmailAddresses as string[]) ?? [];
            console.log(`  ${i + 1}. ${results[i].DisplayName ?? "?"} (${emails[0] ?? ""})`);
          }
          console.log(chalk.green("Using first match."));
        }

        theirMri = (results[0].MRI as string) ?? "";
        const name = (results[0].DisplayName as string) ?? user;
        console.log(`Sending DM to ${chalk.bold(name)}...`);
      }

      const chatClient = await getChatClient();
      const convId = await createDmThread(chatClient, myMri, theirMri);
      const result = await sendMessage(chatClient, convId, message);
      console.log(chalk.green("Sent.") + ` Arrival: ${result.OriginalArrivalTime ?? "ok"}`);
    });

  // members
  program
    .command("members <chat>")
    .description("List members of a conversation.")
    .action(async (chat: string) => {
      const convId = resolveConversationId(chat);
      const client = await getChatClient();
      const raw = await getThreadMembers(client, convId);
      const formatted = raw.map((m) => formatMember(m));

      const table = new Table({
        head: ["Name", "Role", "Type"],
        style: { head: ["cyan"] },
      });
      for (const m of formatted) {
        table.push([m.name, m.role, m.type]);
      }
      console.log(table.toString());
    });

  // watch
  program
    .command("watch [chat]")
    .description("Watch a chat (or all chats) for new messages in real-time.")
    .option("-i, --interval <seconds>", "Poll interval in seconds", "3")
    .action(async (chat: string | undefined, opts) => {
      const interval = Number(opts.interval) * 1000;

      if (chat) {
        const convId = resolveConversationId(chat);
        await watchChat(convId, interval);
      } else {
        await watchAll(interval);
      }
    });

  return program;
}

async function watchChat(convId: string, interval: number): Promise<void> {
  const client = await getChatClient();
  let { syncUrl } = await pollMessages(client, convId);
  console.log(chalk.dim(`Watching for new messages (poll every ${interval / 1000}s). Ctrl+C to stop.`));

  const loop = async () => {
    while (true) {
      await new Promise((r) => setTimeout(r, interval));
      const pollClient = await getChatClient();
      const result = await pollMessages(pollClient, convId, syncUrl);
      syncUrl = result.syncUrl;
      const msgs = result.messages
        .map((r) => formatMessage(r))
        .filter((m): m is FormattedMessage => m !== null);
      printMessages(msgs);
    }
  };

  try {
    await loop();
  } catch (e) {
    if ((e as Error).name !== "AbortError") throw e;
  }
  console.log(chalk.dim("\nStopped watching."));
}

async function watchAll(interval: number): Promise<void> {
  const chats = loadChatIndex();
  const chatNames = new Map(chats.map((c) => [c.id, c.name]));

  const client = await getChatClient();
  let { syncUrl } = await pollConversations(client);
  console.log(chalk.dim(`Watching all chats (poll every ${interval / 1000}s). Ctrl+C to stop.`));

  const loop = async () => {
    while (true) {
      await new Promise((r) => setTimeout(r, interval));
      const pollClient = await getChatClient();
      const result = await pollConversations(pollClient, syncUrl);
      syncUrl = result.syncUrl;

      for (const conv of result.conversations) {
        const convId = (conv.id as string) ?? "";
        if (convId.startsWith("48:")) continue;
        const lastMsg = (conv.lastMessage as Record<string, unknown>) ?? {};
        const msg = formatMessage(lastMsg);
        if (!msg) continue;
        const name = chatNames.get(convId) ?? getConversationDisplayName(conv);
        console.log(`${chalk.cyan(name)} ${chalk.green(msg.time)} ${chalk.bold(msg.sender)}`);
        console.log(`  ${msg.body}`);
        console.log();
      }
    }
  };

  try {
    await loop();
  } catch (e) {
    if ((e as Error).name !== "AbortError") throw e;
  }
  console.log(chalk.dim("\nStopped watching."));
}
