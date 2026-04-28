import * as React from 'react';
import { Text, Stack, ActionButton } from '@fluentui/react';
import type { IListPageRichEditorBlockConfig } from '../../core/config/types';
import { defaultRichEditorConfig } from '../../core/listPage/listPageBlockConfigUtils';
import { sanitizeRichEditorHtml } from '../../core/listPage/richEditorHtmlSanitize';

export interface IListPageRichEditorBlockProps {
  editor?: IListPageRichEditorBlockConfig;
  onConfigure?: () => void;
}

export const ListPageRichEditorBlock: React.FC<IListPageRichEditorBlockProps> = ({
  editor: raw,
  onConfigure,
}) => {
  const c = raw ?? defaultRichEditorConfig();
  const safe = sanitizeRichEditorHtml(c.html, c);
  const minH = Math.max(40, c.minHeightPx);
  const showPlaceholder = !safe.trim() && c.placeholder.trim();

  return (
    <div className="dinamicSxEditor" style={{ marginBottom: 8 }}>
      {onConfigure !== undefined ? (
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { marginBottom: 12 } }}
        >
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
            Editor de conteúdo
          </Text>
          <ActionButton
            iconProps={{ iconName: 'Settings' }}
            onClick={onConfigure}
            styles={{ root: { height: 28, color: '#0078d4' } }}
          >
            Configurar
          </ActionButton>
        </Stack>
      ) : null}
      {c.title.trim() ? (
        <Text variant="large" styles={{ root: { fontWeight: 600, display: 'block', marginBottom: 12 } }}>
          {c.title}
        </Text>
      ) : null}
      <div
        className="list-page-rich-editor-html"
        style={{
          minHeight: minH,
          padding: '12px 14px',
          border: '1px solid #edebe9',
          borderRadius: 4,
          background: '#fff',
          userSelect: c.readOnly ? 'none' : 'text',
        }}
        {...(c.readOnly ? { 'aria-readonly': true as const } : {})}
      >
        {showPlaceholder ? (
          <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
            {c.placeholder}
          </Text>
        ) : (
          <div dangerouslySetInnerHTML={{ __html: safe }} />
        )}
      </div>
      <style>{`
        .list-page-rich-editor-html p { margin: 0 0 0.5em; }
        .list-page-rich-editor-html p:last-child { margin-bottom: 0; }
        .list-page-rich-editor-html ul, .list-page-rich-editor-html ol { margin: 0.5em 0; padding-left: 1.5em; }
        .list-page-rich-editor-html table { border-collapse: collapse; width: 100%; margin: 0.5em 0; }
        .list-page-rich-editor-html th, .list-page-rich-editor-html td { border: 1px solid #edebe9; padding: 6px 8px; }
        .list-page-rich-editor-html img { max-width: 100%; height: auto; }
        .list-page-rich-editor-html iframe { max-width: 100%; border: 0; aspect-ratio: 16/9; width: 100%; min-height: 200px; }
        .list-page-rich-editor-html h1 { font-size: 1.5rem; margin: 0.5em 0; }
        .list-page-rich-editor-html h2 { font-size: 1.25rem; margin: 0.5em 0; }
        .list-page-rich-editor-html h3 { font-size: 1.1rem; margin: 0.5em 0; }
      `}</style>
    </div>
  );
};
