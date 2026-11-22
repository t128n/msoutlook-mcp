import {
  arrayProp,
  boolean,
  booleanProp,
  optionalString,
  requireString,
  schemaObject,
  stringList,
  stringProp,
} from "./args.ts";
import { checkRecipients, confirmRequired, sendAllowed } from "./guardrails.ts";
import { fail, ok, type ToolContext, tool as defineTool } from "./registry.ts";

export function ruleTools(ctx: ToolContext) {
  return [
    defineTool(
      "list_rules",
      "List Outlook rules with decoded conditions and actions. Read-only.",
      schemaObject({}, []),
      async () => {
        const rules = await ctx.bridge.listRules();
        return ok({ count: rules.length, rules });
      },
    ),

    defineTool(
      "toggle_rule",
      "Enable or disable an existing rule by exact name. Needs confirm=true.",
      schemaObject(
        {
          name: stringProp("exact rule name from list_rules"),
          enabled: booleanProp("true to enable, false to disable"),
          confirm: booleanProp("must be true"),
        },
        ["name", "enabled"],
      ),
      async (args) => {
        const name = requireString(args, "name");
        const enabled = boolean(args, "enabled", true);
        if (args.confirm !== true) return confirmRequired("toggle_rule");
        if (enabled) {
          // Enabling a forward rule is a send path. The send switch must allow it.
          const existing = await ctx.bridge.listRules();
          const rule = existing.find((candidate) => candidate.name === name);
          if (rule?.actions.some((action) => action.startsWith("Forward ->"))) {
            const blockedSend = sendAllowed(ctx.guardrails);
            if (blockedSend) return blockedSend;
          }
        }
        await ctx.bridge.toggleRule({ name, enabled });
        ctx.logger.info("toggle_rule", { enabled });
        return ok({ ok: true, name, enabled });
      },
    ),

    defineTool(
      "create_rule",
      "Create a receive rule. Needs confirm=true, at least one condition and one action. Forward addresses are checked against the allowlist.",
      schemaObject(
        {
          name: stringProp("rule name (must be unique)"),
          from_addresses: arrayProp("sender addresses, OR-chain"),
          subject_contains: arrayProp("substrings in the subject, OR-chain"),
          body_or_subject_contains: arrayProp("substrings in subject or body, OR-chain"),
          sent_to: arrayProp("recipient addresses, OR-chain"),
          has_attachment: booleanProp("match mail with an attachment", { default: false }),
          importance_high: booleanProp("match high importance", { default: false }),
          move_to_folder: stringProp("folder to move matching mail to"),
          copy_to_folder: stringProp("folder to copy matching mail to"),
          assign_categories: arrayProp("Outlook categories to assign"),
          forward_to: arrayProp("addresses to forward to (allowlist checked)"),
          stop_processing: booleanProp("stop processing more rules", { default: true }),
          enabled: booleanProp("rule is active", { default: true }),
          apply_now: booleanProp("apply to existing Inbox mail", { default: false }),
          confirm: booleanProp("must be true"),
        },
        ["name"],
      ),
      async (args) => {
        const name = requireString(args, "name");
        const fromAddresses = stringList(args, "from_addresses");
        const subjectContains = stringList(args, "subject_contains");
        const bodySubject = stringList(args, "body_or_subject_contains");
        const sentTo = stringList(args, "sent_to");
        const hasAttachment = boolean(args, "has_attachment", false);
        const importanceHigh = boolean(args, "importance_high", false);
        const moveTo = optionalString(args, "move_to_folder");
        const copyTo = optionalString(args, "copy_to_folder");
        const categories = stringList(args, "assign_categories");
        const forwardTo = stringList(args, "forward_to");
        const stopProcessing = boolean(args, "stop_processing", true);
        const enabled = boolean(args, "enabled", true);
        const applyNow = boolean(args, "apply_now", false);

        const hasCondition =
          fromAddresses.length > 0 ||
          subjectContains.length > 0 ||
          bodySubject.length > 0 ||
          sentTo.length > 0 ||
          hasAttachment ||
          importanceHigh;
        if (!hasCondition) return fail("NO_CONDITION", "At least one condition is required.");
        const hasAction =
          Boolean(moveTo) || Boolean(copyTo) || categories.length > 0 || forwardTo.length > 0;
        if (!hasAction) return fail("NO_ACTION", "At least one action is required.");

        const blockedForward = checkRecipients(ctx.guardrails, forwardTo);
        if (blockedForward) {
          const detail = blockedForward.data as {
            message?: string;
            rejected?: string[];
            allowed_recipients?: string[];
            hint?: string;
          };
          return fail("ADDRESSES_NOT_ALLOWED", detail.message, {
            rejected: detail.rejected,
            allowed_recipients: detail.allowed_recipients,
            hint: detail.hint,
          });
        }

        if (args.confirm !== true) return confirmRequired("create_rule");
        if (forwardTo.length > 0) {
          // A forward rule sends mail unattended. It obeys the same switch.
          const blockedSend = sendAllowed(ctx.guardrails);
          if (blockedSend) return blockedSend;
        }

        const existing = await ctx.bridge.listRules();
        if (existing.some((rule) => rule.name === name)) {
          return fail("RULE_NAME_EXISTS", undefined, { name });
        }

        const created = await ctx.bridge.createRule({
          name,
          from_addresses: fromAddresses,
          subject_contains: subjectContains,
          body_or_subject_contains: bodySubject,
          sent_to: sentTo,
          has_attachment: hasAttachment,
          importance_high: importanceHigh,
          move_to_folder: moveTo ?? null,
          copy_to_folder: copyTo ?? null,
          assign_categories: categories,
          forward_to: forwardTo,
          stop_processing: stopProcessing,
          enabled,
          apply_now: applyNow,
        });
        ctx.logger.info("create_rule", { enabled, apply_now: applyNow, forward: forwardTo.length });
        return ok({
          ok: true,
          name,
          enabled,
          is_local_rule: created.rule.is_local_rule,
          execution_order: created.rule.execution_order,
          summary: created.rule,
          apply_now_executed: applyNow,
        });
      },
    ),

    defineTool(
      "delete_rule",
      "Delete a rule by exact name. Needs confirm=true.",
      schemaObject(
        {
          name: stringProp("exact rule name"),
          confirm: booleanProp("must be true"),
        },
        ["name"],
      ),
      async (args) => {
        const name = requireString(args, "name");
        if (args.confirm !== true) return confirmRequired("delete_rule");
        const existing = await ctx.bridge.listRules();
        if (!existing.some((rule) => rule.name === name)) {
          return fail("RULE_NOT_FOUND", undefined, { name });
        }
        await ctx.bridge.deleteRule({ name });
        ctx.logger.info("delete_rule", {});
        return ok({ ok: true, name, deleted: true });
      },
    ),
  ];
}
