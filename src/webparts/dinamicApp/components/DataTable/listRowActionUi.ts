import type { TListRowActionIconPreset } from '../../core/config/types';

export function listRowActionIconName(preset: TListRowActionIconPreset, customIconName?: string): string {
  switch (preset) {
    case 'view':
      return 'View';
    case 'edit':
      return 'Edit';
    case 'link':
      return 'Link';
    case 'custom':
    default:
      return (customIconName ?? '').trim() || 'OpenInNewWindow';
  }
}
