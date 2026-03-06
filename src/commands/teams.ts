import { Command } from "commander";
import { graph, validateId } from "../lib/graph.js";
import { outputData, handleCommandError } from "../lib/output.js";

export function registerTeamsCommands(program: Command): void {
  const teams = program.command("teams").description("Microsoft Teams operations");

  // m365 teams list
  teams
    .command("list")
    .description("List teams you are a member of")
    .option("--select <fields>", "Comma-separated fields to include")
    .action(async (opts) => {
      try {
        const params = new URLSearchParams();
        if (opts.select) params.set("$select", opts.select);
        const qs = params.toString() ? `?${params.toString()}` : "";
        const result = await graph.get<{ value: unknown[] }>(`/me/joinedTeams${qs}`);
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams get <teamId>
  teams
    .command("get <teamId>")
    .description("Get a team by ID")
    .action(async (teamId) => {
      try {
        validateId(teamId, "team ID");
        const result = await graph.get<unknown>(`/teams/${teamId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // --- channels subgroup ---
  const channels = teams.command("channels").description("Teams channel operations");

  // m365 teams channels list --team <id>
  channels
    .command("list")
    .description("List channels in a team")
    .requiredOption("--team <id>", "Team ID")
    .option("--select <fields>", "Comma-separated fields to include")
    .action(async (opts) => {
      try {
        validateId(opts.team, "team ID");
        const params = new URLSearchParams();
        if (opts.select) params.set("$select", opts.select);
        const qs = params.toString() ? `?${params.toString()}` : "";
        const result = await graph.get<{ value: unknown[] }>(
          `/teams/${opts.team}/channels${qs}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams channels get <channelId> --team <id>
  channels
    .command("get <channelId>")
    .description("Get a channel by ID")
    .requiredOption("--team <id>", "Team ID")
    .action(async (channelId, opts) => {
      try {
        validateId(opts.team, "team ID");
        validateId(channelId, "channel ID");
        const result = await graph.get<unknown>(
          `/teams/${opts.team}/channels/${channelId}`
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams channels messages --team <id> --channel <id>
  channels
    .command("messages")
    .description("List messages in a channel (requires admin consent for ChannelMessage.Read.All)")
    .requiredOption("--team <id>", "Team ID")
    .requiredOption("--channel <id>", "Channel ID")
    .option("--top <n>", "Max number of messages to return", "25")
    .action(async (opts) => {
      try {
        validateId(opts.team, "team ID");
        validateId(opts.channel, "channel ID");
        const params = new URLSearchParams();
        params.set("$top", opts.top);
        const result = await graph.get<{ value: unknown[] }>(
          `/teams/${opts.team}/channels/${opts.channel}/messages?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams channels send --team <id> --channel <id> --body <str>
  channels
    .command("send")
    .description("Send a message to a channel")
    .requiredOption("--team <id>", "Team ID")
    .requiredOption("--channel <id>", "Channel ID")
    .requiredOption("--body <str>", "Message body")
    .option("--html", "Treat body as HTML (default: plain text)")
    .option("--subject <str>", "Message subject (channel posts only)")
    .action(async (opts) => {
      try {
        validateId(opts.team, "team ID");
        validateId(opts.channel, "channel ID");
        const payload: Record<string, unknown> = {
          body: {
            contentType: opts.html ? "html" : "text",
            content: opts.body,
          },
        };
        if (opts.subject) payload.subject = opts.subject;
        const result = await graph.post<unknown>(
          `/teams/${opts.team}/channels/${opts.channel}/messages`,
          payload
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // --- chats subgroup ---
  const chats = teams.command("chats").description("Teams chat operations");

  // m365 teams chats list
  chats
    .command("list")
    .description("List your chats (1:1, group, meeting)")
    .option("--filter <odata>", "OData filter expression")
    .option("--select <fields>", "Comma-separated fields to include")
    .option("--top <n>", "Max number of chats to return", "25")
    .action(async (opts) => {
      try {
        const params = new URLSearchParams();
        params.set("$top", opts.top);
        if (opts.filter) params.set("$filter", opts.filter);
        if (opts.select) params.set("$select", opts.select);
        const result = await graph.get<{ value: unknown[] }>(
          `/me/chats?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams chats get <chatId>
  chats
    .command("get <chatId>")
    .description("Get a chat by ID")
    .action(async (chatId) => {
      try {
        validateId(chatId, "chat ID");
        const result = await graph.get<unknown>(`/me/chats/${chatId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams chats messages <chatId>
  chats
    .command("messages <chatId>")
    .description("List messages in a chat")
    .option("--top <n>", "Max number of messages to return", "25")
    .option("--select <fields>", "Comma-separated fields to include")
    .action(async (chatId, opts) => {
      try {
        validateId(chatId, "chat ID");
        const params = new URLSearchParams();
        params.set("$top", opts.top);
        if (opts.select) params.set("$select", opts.select);
        const result = await graph.get<{ value: unknown[] }>(
          `/me/chats/${chatId}/messages?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 teams chats send <chatId>
  chats
    .command("send <chatId>")
    .description("Send a message to a chat")
    .requiredOption("--body <str>", "Message body")
    .option("--html", "Treat body as HTML (default: plain text)")
    .action(async (chatId, opts) => {
      try {
        validateId(chatId, "chat ID");
        const result = await graph.post<unknown>(`/me/chats/${chatId}/messages`, {
          body: {
            contentType: opts.html ? "html" : "text",
            content: opts.body,
          },
        });
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });
}
