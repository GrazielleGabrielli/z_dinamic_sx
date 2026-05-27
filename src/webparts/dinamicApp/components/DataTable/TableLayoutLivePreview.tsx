import * as React from 'react';
import { Text } from '@fluentui/react';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableLayoutLivePreviewProps {
  cssText: string;
  rulePreviewTokens?: readonly string[];
}

const scopeClass = 'dinamicSxLayoutLivePreviewScope';

const StackWrap: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div
    style={{
      padding: 14,
      border: '1px solid #e4e4e7',
      borderRadius: 12,
      background: 'linear-gradient(180deg, #fafafa 0%, #f4f4f5 100%)',
    }}
  >
    {children}
  </div>
);

export const TableLayoutLivePreview: React.FC<ITableLayoutLivePreviewProps> = ({ cssText, rulePreviewTokens }) => {
  const cssTrim = (cssText ?? '').trim();
  const scopedCss = cssTrim
    ? cssTrim.replace(/\.dinamicSxTable/g, `.${scopeClass} .dinamicSxTable`)
    : '';

  const C = DINAMIC_SX_TABLE_CLASS;

  return (
    <StackWrap>
      {scopedCss ? <style type="text/css">{scopedCss}</style> : null}
      <Text variant="small" styles={{ root: { color: '#71717a', marginBottom: 10, display: 'block', fontWeight: 600, letterSpacing: '0.04em', textTransform: 'uppercase', fontSize: 10 } }}>
        Pré-visualização
      </Text>
      <div className={scopeClass}>
        <div className={C.viewRoot}>
          <div className={C.toolbar}>
            <span style={{ fontSize: 11, fontWeight: 500, color: '#71717a' }}>Visualização</span>
          </div>
          <div className={C.scrollWrap}>
            <table className={C.table} role="presentation">
              <thead className={C.thead}>
                <tr className={C.headerRow}>
                  <th className={C.headerCell} data-field="Title">
                    <span className={C.headerCellInner}>
                      Título
                      <span className={C.headerFilterTrigger} aria-hidden>⧩</span>
                    </span>
                  </th>
                  <th className={C.headerCell} data-field="Status">
                    <span className={C.headerCellInner}>Estado</span>
                  </th>
                </tr>
              </thead>
              <tbody className={C.body}>
                <tr className={C.row}>
                  <td
                    className={C.cell}
                    data-field="Title"
                    {...(rulePreviewTokens?.[0] ? { 'data-dinamic-rules': rulePreviewTokens[0] } : {})}
                  >
                    Registo de exemplo
                  </td>
                  <td
                    className={C.cell}
                    data-field="Status"
                    {...(rulePreviewTokens?.[0] ? { 'data-dinamic-rules': rulePreviewTokens[0] } : {})}
                  >
                    Ativo
                  </td>
                </tr>
                <tr className={C.row}>
                  <td
                    className={C.cell}
                    data-field="Title"
                    {...(rulePreviewTokens?.[1] ? { 'data-dinamic-rules': rulePreviewTokens[1] } : {})}
                  >
                    Outro registo
                  </td>
                  <td
                    className={C.cell}
                    data-field="Status"
                    {...(rulePreviewTokens?.[1] ? { 'data-dinamic-rules': rulePreviewTokens[1] } : {})}
                  >
                    Pendente
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
          <div className={C.pagination}>
            <button type="button" className={C.paginationBtn}>
              Anterior
            </button>
            <button type="button" className={C.paginationBtn} style={{ fontWeight: 700, borderColor: '#0f6cbd', color: '#0f6cbd', background: '#f0f6fc' }}>
              1
            </button>
            <button type="button" className={C.paginationBtn}>
              Próxima
            </button>
          </div>
        </div>
      </div>
    </StackWrap>
  );
};
