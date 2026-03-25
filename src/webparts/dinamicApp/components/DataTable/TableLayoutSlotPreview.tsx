import * as React from 'react';
import { Text } from '@fluentui/react';
import type { TTableCssSlot } from '../../core/config/types';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableLayoutSlotPreviewProps {
  slot: TTableCssSlot;
  cssBody: string;
  /** Visual mais leve quando a prévia fica ao lado do editor. */
  variant?: 'default' | 'embedded';
}

const wrapEmbedded: React.CSSProperties = {
  marginTop: 0,
  padding: 8,
  borderRadius: 6,
  border: '1px solid #edebe9',
  background: '#ffffff',
  maxWidth: '100%',
  boxSizing: 'border-box',
};

const wrapDefault: React.CSSProperties = {
  marginTop: 8,
  padding: 10,
  borderRadius: 4,
  border: '1px dashed #c8c6c4',
  background: '#ffffff',
  maxWidth: '100%',
  boxSizing: 'border-box',
};

function previewMock(slot: TTableCssSlot): React.ReactNode {
  const C = DINAMIC_SX_TABLE_CLASS;
  const hint = (t: string): React.ReactNode => (
    <span style={{ fontSize: 10, color: '#a19f9d', pointerEvents: 'none' }}>{t}</span>
  );

  switch (slot) {
    case 'viewRoot':
      return (
        <div className={C.viewRoot}>
          {hint('área geral · toolbar + tabela + paginação')}
        </div>
      );
    case 'toolbar':
      return (
        <div className={C.toolbar} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <span style={{ fontSize: 11, border: '1px solid #edebe9', padding: '2px 8px', borderRadius: 2 }}>Visualização</span>
          <span style={{ fontSize: 11, color: '#0078d4' }}>PDF</span>
        </div>
      );
    case 'scrollWrap':
      return (
        <div className={C.scrollWrap} style={{ maxHeight: 72, overflow: 'auto' }}>
          <table style={{ width: '100%', fontSize: 11, borderCollapse: 'collapse' }}>
            <tbody>
              <tr>
                <td style={{ border: '1px solid #edebe9', padding: 4 }}>A</td>
                <td style={{ border: '1px solid #edebe9', padding: 4 }}>B</td>
                <td style={{ border: '1px solid #edebe9', padding: 4 }}>C</td>
              </tr>
            </tbody>
          </table>
        </div>
      );
    case 'table':
      return (
        <table className={C.table} style={{ width: '100%', fontSize: 11 }}>
          <tbody>
            <tr>
              <td style={{ border: '1px solid #edebe9', padding: 4 }}>célula</td>
            </tr>
          </tbody>
        </table>
      );
    case 'thead':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead className={C.thead}>
            <tr>
              <th style={{ border: '1px solid #edebe9', padding: 4, textAlign: 'left' }}>Coluna</th>
            </tr>
          </thead>
        </table>
      );
    case 'headerRow':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead>
            <tr className={C.headerRow}>
              <th style={{ border: '1px solid #edebe9', padding: 4, textAlign: 'left' }}>A</th>
              <th style={{ border: '1px solid #edebe9', padding: 4, textAlign: 'left' }}>B</th>
            </tr>
          </thead>
        </table>
      );
    case 'headerCell':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead>
            <tr>
              <th className={C.headerCell} data-field="Preview">
                Título
              </th>
            </tr>
          </thead>
        </table>
      );
    case 'headerCellInner':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead>
            <tr>
              <th style={{ border: '1px solid #edebe9', padding: 4, textAlign: 'left' }}>
                <span className={C.headerCellInner} style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                  Coluna
                  <span style={{ fontSize: 9, opacity: 0.7 }}>▲</span>
                </span>
              </th>
            </tr>
          </thead>
        </table>
      );
    case 'headerFilterTrigger':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead>
            <tr>
              <th style={{ border: '1px solid #edebe9', padding: 4, textAlign: 'left' }}>
                Nome
                <span className={C.headerFilterTrigger} style={{ marginLeft: 6, cursor: 'pointer', fontSize: 10 }}>
                  ⧩
                </span>
              </th>
            </tr>
          </thead>
        </table>
      );
    case 'body':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <tbody className={C.body}>
            <tr>
              <td style={{ border: '1px solid #edebe9', padding: 4 }}>linha</td>
            </tr>
          </tbody>
        </table>
      );
    case 'row':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <tbody>
            <tr className={C.row}>
              <td style={{ border: '1px solid #edebe9', padding: 4 }}>dado 1</td>
              <td style={{ border: '1px solid #edebe9', padding: 4 }}>dado 2</td>
            </tr>
          </tbody>
        </table>
      );
    case 'cell':
      return (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <tbody>
            <tr>
              <td className={C.cell} data-field="Preview">
                Valor de exemplo
              </td>
            </tr>
          </tbody>
        </table>
      );
    case 'empty':
      return (
        <div className={C.empty} style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: 56 }}>
          <span style={{ fontSize: 12, color: '#605e5c' }}>Nenhum item</span>
        </div>
      );
    case 'loading':
      return (
        <div className={C.loading} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, minHeight: 56 }}>
          <span style={{ fontSize: 11, color: '#a19f9d' }}>Carregando…</span>
        </div>
      );
    case 'error':
      return (
        <div className={C.error} style={{ padding: 8 }}>
          <div style={{ fontSize: 11, color: '#a4262c', background: '#fde7e9', padding: 8, borderRadius: 2 }}>
            Erro de exemplo
          </div>
        </div>
      );
    case 'pagination':
      return (
        <div className={C.pagination} style={{ display: 'flex', gap: 6, justifyContent: 'flex-end', alignItems: 'center' }}>
          <span style={{ fontSize: 10, marginRight: 4 }}>Página 1</span>
          <span style={{ fontSize: 11, border: '1px solid #8a8886', padding: '2px 8px', borderRadius: 2 }}>Anterior</span>
          <span style={{ fontSize: 11, border: '1px solid #8a8886', padding: '2px 8px', borderRadius: 2 }}>Próxima</span>
        </div>
      );
  }
}

export const TableLayoutSlotPreview: React.FC<ITableLayoutSlotPreviewProps> = ({
  slot,
  cssBody,
  variant = 'default',
}) => {
  const cls = DINAMIC_SX_TABLE_CLASS[slot];
  const scopeClass = `dinamicSxLayoutPreviewScope_${slot}`;
  const raw = (cssBody ?? '').trim();
  const styleBlock = raw ? `.${scopeClass} .${cls} {\n${raw}\n}` : '';
  const wrap = variant === 'embedded' ? wrapEmbedded : wrapDefault;

  return (
    <div className={scopeClass} style={wrap}>
      <Text
        variant="small"
        styles={{
          root: {
            color: '#a19f9d',
            marginBottom: variant === 'embedded' ? 6 : 8,
            display: 'block',
            fontWeight: 600,
            fontSize: variant === 'embedded' ? 11 : undefined,
          },
        }}
      >
        Pré-visualização
      </Text>
      {styleBlock ? <style type="text/css">{styleBlock}</style> : null}
      {previewMock(slot)}
    </div>
  );
};
