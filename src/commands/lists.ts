import { Command } from "commander";
import { graph, validateId } from "../lib/graph.js";
import { readConfig } from "../lib/config.js";
import { outputData, handleCommandError } from "../lib/output.js";

function resolveSite(opts: { site?: string }): string {
  const siteId = opts.site || readConfig().defaultSiteId;
  if (!siteId) throw new Error("Site ID required. Use --site or run `sp config set --site <id>`.");
  validateId(siteId, "site ID");
  return siteId;
}

export function registerListCommands(program: Command): void {
  const lists = program.command("lists").description("SharePoint list and list-item operations");

  // ── List CRUD ─────────────────────────────────────────────────────────────

  lists
    .command("list")
    .description("List all lists in a site (use --type to filter by generic or documentLibrary)")
    .option("--site <id>", "Site ID")
    .option("--type <type>", "Filter by template type: generic or documentLibrary")
    .action(async (opts) => {
      try {
        const siteId = resolveSite(opts);
        const params = new URLSearchParams();
        if (opts.type) {
          if (opts.type !== "generic" && opts.type !== "documentLibrary") {
            throw new Error("--type must be 'generic' or 'documentLibrary'");
          }
          params.set("$filter", `list/template eq '${opts.type}'`);
        }
        const url = params.toString()
          ? `/sites/${siteId}/lists?${params.toString()}`
          : `/sites/${siteId}/lists`;
        const result = await graph.get<{ value: unknown[] }>(url);
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  lists
    .command("get <listId>")
    .description("Get a specific list by ID")
    .option("--site <id>", "Site ID")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        const result = await graph.get<unknown>(`/sites/${siteId}/lists/${listId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  lists
    .command("create")
    .description("Create a new metadata list (generic template). To create a document library, use `m365 drives` or the SharePoint UI.")
    .requiredOption("--name <name>", "List display name")
    .option("--site <id>", "Site ID")
    .action(async (opts) => {
      try {
        const siteId = resolveSite(opts);
        const result = await graph.post<unknown>(`/sites/${siteId}/lists`, {
          displayName: opts.name,
          list: { template: "generic" },
        });
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  lists
    .command("update <listId>")
    .description("Update a list")
    .option("--site <id>", "Site ID")
    .option("--name <name>", "New display name")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        const body: Record<string, unknown> = {};
        if (opts.name) body.displayName = opts.name;
        if (Object.keys(body).length === 0) throw new Error("No updates provided.");
        const result = await graph.patch<unknown>(`/sites/${siteId}/lists/${listId}`, body);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  lists
    .command("delete <listId>")
    .description("Delete a list")
    .option("--site <id>", "Site ID")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        await graph.delete(`/sites/${siteId}/lists/${listId}`);
        outputData({ message: `List ${listId} deleted.` });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // ── Items subcommand ───────────────────────────────────────────────────────

  const items = lists.command("items").description("Manage list items");

  items
    .command("list <listId>")
    .description("List items in a list")
    .option("--site <id>", "Site ID")
    .option("--filter <odata>", "OData filter expression")
    .option("--select <fields>", "Comma-separated fields to return")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        const params = new URLSearchParams();
        params.set("expand", "fields");
        if (opts.filter) params.set("$filter", opts.filter);
        if (opts.select) params.set("$select", opts.select);
        const result = await graph.get<{ value: unknown[] }>(
          `/sites/${siteId}/lists/${listId}/items?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  items
    .command("get <listId> <itemId>")
    .description("Get a specific list item")
    .option("--site <id>", "Site ID")
    .action(async (listId, itemId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(itemId, "item ID");
        const result = await graph.get<unknown>(
          `/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  items
    .command("create <listId>")
    .description("Create a new list item")
    .requiredOption("--fields <json>", "JSON object of field values")
    .option("--site <id>", "Site ID")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        const fields = JSON.parse(opts.fields) as Record<string, unknown>;
        const result = await graph.post<unknown>(`/sites/${siteId}/lists/${listId}/items`, {
          fields,
        });
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  items
    .command("update <listId> <itemId>")
    .description("Update a list item's fields")
    .requiredOption("--fields <json>", "JSON object of field values to update")
    .option("--site <id>", "Site ID")
    .action(async (listId, itemId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(itemId, "item ID");
        const fields = JSON.parse(opts.fields) as Record<string, unknown>;
        const result = await graph.patch<unknown>(
          `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
          fields
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  items
    .command("delete <listId> <itemId>")
    .description("Delete a list item")
    .option("--site <id>", "Site ID")
    .action(async (listId, itemId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(itemId, "item ID");
        await graph.delete(`/sites/${siteId}/lists/${listId}/items/${itemId}`);
        outputData({ message: `Item ${itemId} deleted from list ${listId}.` });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // ── Columns subcommand ────────────────────────────────────────────────────

  const columns = lists.command("columns").description("Manage list columns");

  columns
    .command("list <listId>")
    .description("List columns in a list or document library")
    .option("--site <id>", "Site ID")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        const result = await graph.get<{ value: unknown[] }>(
          `/sites/${siteId}/lists/${listId}/columns`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  columns
    .command("get <listId> <columnId>")
    .description("Get a specific column definition")
    .option("--site <id>", "Site ID")
    .action(async (listId, columnId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(columnId, "column ID");
        const result = await graph.get<unknown>(
          `/sites/${siteId}/lists/${listId}/columns/${columnId}`
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  columns
    .command("add <listId>")
    .description("Add a column to a list")
    .requiredOption("--name <name>", "Internal name for the column")
    .requiredOption(
      "--type <type>",
      "Column type: text|number|boolean|dateTime|choice|person|lookup|hyperlink|currency"
    )
    .option("--required", "Make the column required")
    .option("--site <id>", "Site ID")
    .action(async (listId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");

        const columnDef: Record<string, unknown> = {
          name: opts.name,
          required: !!opts.required,
        };

        // Map type string to Graph column definition
        switch (opts.type) {
          case "text":
            columnDef.text = {};
            break;
          case "number":
            columnDef.number = {};
            break;
          case "boolean":
            columnDef.boolean = {};
            break;
          case "dateTime":
            columnDef.dateTime = { format: "dateTime" };
            break;
          case "choice":
            columnDef.choice = { allowTextEntry: false, choices: [] };
            break;
          case "person":
            columnDef.personOrGroup = { allowMultipleSelection: false };
            break;
          case "lookup":
            columnDef.lookup = {};
            break;
          case "hyperlink":
            columnDef.hyperlinkOrPicture = { isPicture: false };
            break;
          case "currency":
            columnDef.currency = { locale: "en-US" };
            break;
          default:
            throw new Error(`Unknown column type: ${opts.type}`);
        }

        const result = await graph.post<unknown>(
          `/sites/${siteId}/lists/${listId}/columns`,
          columnDef
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  columns
    .command("update <listId> <columnId>")
    .description("Update a column definition")
    .option("--site <id>", "Site ID")
    .option("--name <name>", "New display name")
    .option("--required <bool>", "Set required (true/false)")
    .action(async (listId, columnId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(columnId, "column ID");
        const body: Record<string, unknown> = {};
        if (opts.name) body.displayName = opts.name;
        if (opts.required !== undefined) body.required = opts.required === "true";
        if (Object.keys(body).length === 0) throw new Error("No updates provided.");
        const result = await graph.patch<unknown>(
          `/sites/${siteId}/lists/${listId}/columns/${columnId}`,
          body
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  columns
    .command("remove <listId> <columnId>")
    .description("Remove a column from a list")
    .option("--site <id>", "Site ID")
    .action(async (listId, columnId, opts) => {
      try {
        const siteId = resolveSite(opts);
        validateId(listId, "list ID");
        validateId(columnId, "column ID");
        await graph.delete(`/sites/${siteId}/lists/${listId}/columns/${columnId}`);
        outputData({ message: `Column ${columnId} removed from list ${listId}.` });
      } catch (err) {
        handleCommandError(err);
      }
    });
}
