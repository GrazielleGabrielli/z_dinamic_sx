import * as React from 'react';
import { useMemo } from 'react';
import { Text } from '@fluentui/react';
import ReactQuill from 'react-quill';
import type { IListPageRichEditorBlockConfig } from '../../core/config/types';
import 'react-quill/dist/quill.snow.css';

export type TListPageRichQuillPermissions = Pick<
  IListPageRichEditorBlockConfig,
  'allowImages' | 'allowLinks' | 'allowLists' | 'allowHeaders' | 'allowVideoEmbed'
>;

export interface IListPageRichQuillEditorProps {
  value: string;
  onChange: (html: string) => void;
  placeholder?: string;
  permissions: TListPageRichQuillPermissions;
}

export const ListPageRichQuillEditor: React.FC<IListPageRichQuillEditorProps> = ({
  value,
  onChange,
  placeholder,
  permissions,
}) => {
  const modules = useMemo(() => {
    const toolbar: unknown[] = [];
    if (permissions.allowHeaders) {
      toolbar.push([{ header: [1, 2, 3, false] }]);
    }
    toolbar.push(['bold', 'italic', 'underline', 'strike']);
    if (permissions.allowLists) {
      toolbar.push([{ list: 'ordered' }, { list: 'bullet' }]);
      toolbar.push([{ indent: '-1' }, { indent: '+1' }]);
    }
    const embeds: string[] = [];
    if (permissions.allowLinks) embeds.push('link');
    if (permissions.allowImages) embeds.push('image');
    if (permissions.allowVideoEmbed) embeds.push('video');
    if (embeds.length > 0) {
      toolbar.push(embeds);
    }
    toolbar.push(['blockquote', 'code-block']);
    toolbar.push(['clean']);
    return { toolbar };
  }, [
    permissions.allowHeaders,
    permissions.allowLists,
    permissions.allowLinks,
    permissions.allowImages,
    permissions.allowVideoEmbed,
  ]);

  const formats = useMemo(() => {
    const f = ['bold', 'italic', 'underline', 'strike', 'blockquote', 'code-block'];
    if (permissions.allowHeaders) f.push('header');
    if (permissions.allowLists) f.push('list', 'indent');
    if (permissions.allowLinks) f.push('link');
    if (permissions.allowImages) f.push('image');
    if (permissions.allowVideoEmbed) f.push('video');
    return f;
  }, [
    permissions.allowHeaders,
    permissions.allowLists,
    permissions.allowLinks,
    permissions.allowImages,
    permissions.allowVideoEmbed,
  ]);

  const quillKey = [
    permissions.allowHeaders ? 'h' : '',
    permissions.allowLists ? 'l' : '',
    permissions.allowLinks ? 'a' : '',
    permissions.allowImages ? 'i' : '',
    permissions.allowVideoEmbed ? 'v' : '',
  ].join('');

  return (
    <div className="list-page-rich-quill-wrap" style={{ marginBottom: 8 }}>
      <Text variant="small" styles={{ root: { display: 'block', marginBottom: 6, fontWeight: 600 } }}>
        Conteúdo
      </Text>
      <ReactQuill
        key={quillKey}
        theme="snow"
        value={value || ''}
        onChange={onChange}
        modules={modules}
        formats={formats}
        placeholder={placeholder || ''}
      />
      <style>{`
        .list-page-rich-quill-wrap .ql-toolbar.ql-snow { border-radius: 4px 4px 0 0; border-color: #8a8886; flex-wrap: wrap; }
        .list-page-rich-quill-wrap .ql-container.ql-snow { border-radius: 0 0 4px 4px; border-color: #8a8886; font-size: 14px; }
        .list-page-rich-quill-wrap .ql-editor { min-height: 200px; }
        .list-page-rich-quill-wrap .ql-editor.ql-blank::before { font-style: normal; color: #a19f9d; }
      `}</style>
    </div>
  );
};
