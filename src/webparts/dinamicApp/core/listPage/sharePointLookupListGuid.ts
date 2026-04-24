export function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

export function listGuidFromLookupListField(raw: string | undefined): string {
  if (!raw) return '';
  const t = raw.trim();
  const inBraces = t.match(/\{([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}/i);
  if (inBraces) return normListGuid(inBraces[1]);
  const plain = t.match(
    /([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i
  );
  if (plain) return normListGuid(plain[1]);
  return normListGuid(t);
}
