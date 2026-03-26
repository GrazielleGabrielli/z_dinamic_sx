import * as React from 'react';
import { Text } from '@fluentui/react';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableLayoutLivePreviewProps {
  cssText: string;
}

const scopeClass = 'dinamicSxLayoutLivePreviewScope';

const StackWrap: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div style={{ padding: 10, border: '1px solid #edebe9', borderRadius: 8, background: '#fff' }}>{children}</div>
);

export const TableLayoutLivePreview: React.FC<ITableLayoutLivePreviewProps> = ({ cssText }) => {
  const cssTrim = (cssText ?? '').trim();
  const scopedCss = cssTrim
    ? cssTrim.replace(/\.dinamicSxTable/g, `.${scopeClass} .dinamicSxTable`)
    : '';

  return (
    <StackWrap>
      {scopedCss ? <style type="text/css">{scopedCss}</style> : null}
      <Text variant="small" styles={{ root: { color: '#a19f9d', marginBottom: 8, display: 'block', fontWeight: 600 } }}>
        Pré-visualização geral
      </Text>
      <div className={scopeClass}>
        <div className={DINAMIC_SX_TABLE_CLASS.viewRoot} style={{ padding: 8, background: '#faf9f8' }}>
          <div className={DINAMIC_SX_TABLE_CLASS.toolbar} style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
            <span style={{ fontSize: 11, border: '1px solid #edebe9', padding: '2px 8px', borderRadius: 2 }}>Visualização</span>
            <span style={{ fontSize: 11, color: '#0078d4' }}>Exportar PDF</span>
          </div>
          <div className={DINAMIC_SX_TABLE_CLASS.scrollWrap} style={{ overflowX: 'auto' }}>
            <table className={DINAMIC_SX_TABLE_CLASS.table} style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
              <thead className={DINAMIC_SX_TABLE_CLASS.thead}>
                <tr className={DINAMIC_SX_TABLE_CLASS.headerRow}>
                  <th className={DINAMIC_SX_TABLE_CLASS.headerCell} data-field="Title" style={{ border: '1px solid #edebe9', padding: 6, textAlign: 'left' }}>
                    <span className={DINAMIC_SX_TABLE_CLASS.headerCellInner} style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
                      Title
                      <span className={DINAMIC_SX_TABLE_CLASS.headerFilterTrigger}>⧩</span>
                    </span>
                  </th>
                  <th className={DINAMIC_SX_TABLE_CLASS.headerCell} data-field="Status" style={{ border: '1px solid #edebe9', padding: 6, textAlign: 'left' }}>
                    Status
                  </th>
                </tr>
              </thead>
              <tbody className={DINAMIC_SX_TABLE_CLASS.body}>
                <tr className={DINAMIC_SX_TABLE_CLASS.row}>
                  <td className={DINAMIC_SX_TABLE_CLASS.cell} data-field="Title" style={{ border: '1px solid #edebe9', padding: 6 }}>Teste</td>
                  <td className={DINAMIC_SX_TABLE_CLASS.cell} data-field="Status" style={{ border: '1px solid #edebe9', padding: 6 }}>Ativo</td>
                </tr>
                <tr className={DINAMIC_SX_TABLE_CLASS.row}>
                  <td className={DINAMIC_SX_TABLE_CLASS.cell} data-field="Title" style={{ border: '1px solid #edebe9', padding: 6 }}>Outro item</td>
                  <td className={DINAMIC_SX_TABLE_CLASS.cell} data-field="Status" style={{ border: '1px solid #edebe9', padding: 6 }}>Pendente</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div className={DINAMIC_SX_TABLE_CLASS.pagination} style={{ marginTop: 8, display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
            <span style={{ fontSize: 11, border: '1px solid #8a8886', padding: '2px 8px', borderRadius: 2 }}>Anterior</span>
            <span style={{ fontSize: 11, border: '1px solid #8a8886', padding: '2px 8px', borderRadius: 2 }}>Próxima</span>
          </div>
        </div>
        <div className={DINAMIC_SX_TABLE_CLASS.empty} style={{ marginTop: 8, padding: 10, border: '1px solid #edebe9' }}>Estado vazio</div>
        <div className={DINAMIC_SX_TABLE_CLASS.loading} style={{ marginTop: 6, padding: 10, border: '1px solid #edebe9' }}>Carregando...</div>
        <div className={DINAMIC_SX_TABLE_CLASS.error} style={{ marginTop: 6, padding: 10, border: '1px solid #edebe9' }}>Erro de exemplo</div>
      </div>
    </StackWrap>
  );
};
