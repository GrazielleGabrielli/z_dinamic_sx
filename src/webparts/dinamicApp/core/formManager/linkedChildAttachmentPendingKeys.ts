const SEP = '\x1e';

export function linkedChildAttPendingKey(cfgId: string, localKey: string, folderNodeId = ''): string {
  return `${cfgId}${SEP}${localKey}${SEP}${folderNodeId}`;
}

export function parseLinkedChildAttPendingKey(key: string): {
  cfgId: string;
  localKey: string;
  folderNodeId: string;
} {
  const parts = key.split(SEP);
  return {
    cfgId: parts[0] ?? '',
    localKey: parts[1] ?? '',
    folderNodeId: parts[2] ?? '',
  };
}
