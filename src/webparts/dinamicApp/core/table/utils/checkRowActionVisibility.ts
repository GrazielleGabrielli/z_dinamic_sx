import type { IListRowActionConfig, IListRowActionFieldRule } from '../../config/types';
import type { IDynamicContext } from '../../dynamicTokens/types';

function resolveRuleValue(raw: string, ctx: IDynamicContext): string {
  const t = raw.trim();
  if (t === '[Me.Id]') return String(ctx.currentUser?.id ?? '');
  if (t === '[Me.Login]') return (ctx.currentUser?.loginName ?? '').toLowerCase();
  if (t === '[Me.Email]') return (ctx.currentUser?.email ?? '').toLowerCase();
  return t;
}

function readItemField(item: Record<string, unknown>, field: string): unknown {
  const slash = field.indexOf('/');
  if (slash === -1) return item[field];
  const base = field.slice(0, slash);
  const sub = field.slice(slash + 1);
  const parent = item[base];
  if (parent && typeof parent === 'object') {
    return (parent as Record<string, unknown>)[sub];
  }
  // Fallback: SharePoint às vezes retorna Author/Id como item.AuthorId
  return item[`${base}${sub}`];
}

function evalFieldRule(
  rule: IListRowActionFieldRule,
  item: Record<string, unknown>,
  ctx: IDynamicContext
): boolean {
  const itemVal = String(readItemField(item, rule.field) ?? '').trim().toLowerCase();
  const ruleVal = resolveRuleValue(rule.value, ctx).toLowerCase();
  if (rule.op === 'eq') return itemVal === ruleVal;
  if (rule.op === 'ne') return itemVal !== ruleVal;
  return true;
}

export function checkRowActionVisibility(
  action: IListRowActionConfig,
  item: Record<string, unknown>,
  ctx: IDynamicContext,
  userGroupIds?: Set<number>
): boolean {
  const vis = action.visibility;
  if (!vis) return true;

  const hasGroupRestriction = (vis.allowedGroupIds?.length ?? 0) > 0;
  const hasUserRestriction = (vis.allowedUserLogins?.length ?? 0) > 0;
  const hasFieldRules = (vis.fieldRules?.length ?? 0) > 0;

  if (!hasGroupRestriction && !hasUserRestriction && !hasFieldRules) return true;

  // Group OR User check: pelo menos um critério de identidade deve passar
  const identityRestricted = hasGroupRestriction || hasUserRestriction;
  if (identityRestricted) {
    let identityPassed = false;

    if (hasGroupRestriction && userGroupIds) {
      identityPassed = (vis.allowedGroupIds ?? []).some((gid) => {
        const n = parseInt(gid, 10);
        return !isNaN(n) && userGroupIds.has(n);
      });
    }

    if (!identityPassed && hasUserRestriction) {
      const myLogin = (ctx.currentUser?.loginName ?? '').toLowerCase();
      identityPassed = (vis.allowedUserLogins ?? []).some(
        (l) => l.trim().toLowerCase() === myLogin
      );
    }

    if (!identityPassed) return false;
  }

  // Field rules: AND — todas devem ser verdadeiras
  if (hasFieldRules) {
    for (const rule of vis.fieldRules ?? []) {
      if (!evalFieldRule(rule, item, ctx)) return false;
    }
  }

  return true;
}
