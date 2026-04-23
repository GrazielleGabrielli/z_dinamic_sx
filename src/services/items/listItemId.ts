export function readListItemId(row: Record<string, unknown> | undefined): number | undefined {
  if (!row) return undefined;
  const id = row.Id ?? row.ID ?? row.id;
  if (typeof id === 'number' && Number.isFinite(id)) return id;
  if (typeof id === 'string' && id.trim() !== '') {
    const n = parseInt(id, 10);
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}
