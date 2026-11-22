import { integer, numberProp, requireString, schemaObject, stringProp } from "./args.ts";
import { ok, type ToolContext, tool } from "./registry.ts";

export function contactTools(ctx: ToolContext) {
  return [
    tool(
      "search_contacts",
      "Search the Exchange address book (GAL) by name or address. Uses ADODB/LDAP inside the helper and falls back to CreateRecipient (one result). Read-only.",
      schemaObject(
        {
          query: stringProp("name, first name or partial address"),
          limit: numberProp("maximum contacts (default 20, max 1000)", { default: 20, minimum: 1, maximum: 1000 }),
        },
        ["query"],
      ),
      async (args) => {
        const query = requireString(args, "query");
        const limit = integer(args, "limit", 20, 1, 1000);
        const result = await ctx.bridge.searchContacts({ query, limit });
        return ok({
          query,
          count: result.contacts.length,
          contacts: result.contacts,
          ...(result.hint ? { hint: result.hint } : {}),
        });
      },
    ),

    tool(
      "get_contact_details",
      "Get the full details of an Exchange contact from its SMTP address. Read-only.",
      schemaObject(
        {
          smtp_address: stringProp("contact email address"),
        },
        ["smtp_address"],
      ),
      async (args) => {
        const smtp = requireString(args, "smtp_address");
        return ok(await ctx.bridge.getContactDetails({ smtpAddress: smtp }));
      },
    ),
  ];
}
