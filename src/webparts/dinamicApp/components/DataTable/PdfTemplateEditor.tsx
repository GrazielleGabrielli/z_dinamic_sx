import * as React from 'react';
import { useState, useCallback, useRef, useEffect } from 'react';
import {
  Stack,
  Text,
  TextField,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IconButton,
  ChoiceGroup,
  IChoiceGroupOption,
} from '@fluentui/react';
import type { IPdfTemplateConfig, IPdfTemplateElement, TPdfLayoutMode, TPdfElementScope } from '../../core/config/types';

const A4_W = 210;
const A4_H = 297;
const SCALE = 2;
const CANVAS_W = A4_W * SCALE;
const CANVAS_H = A4_H * SCALE;

function defaultTemplate(): IPdfTemplateConfig {
  return {
    pageFormat: 'A4',
    orientation: 'portrait',
    body: {
      elements: [
        { id: 'title', type: 'text', scope: 'dynamic', x: 20, y: 20, width: 170, content: '{{Title}}', fontSize: 14, fontWeight: 'bold' },
      ],
    },
  };
}

function ensureTemplate(t: IPdfTemplateConfig | undefined): IPdfTemplateConfig {
  if (t?.body?.elements && Array.isArray(t.body.elements)) return t;
  return defaultTemplate();
}

const PdfTemplateImagePreview: React.FC<{ url: string }> = ({ url }) => {
  const [error, setError] = useState(false);
  useEffect(() => setError(false), [url]);
  if (!url) {
    return (
      <div style={{ width: '100%', height: '100%', background: '#f3f2f1', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, color: '#a19f9d' }}>
        [URL da imagem]
      </div>
    );
  }
  if (error) {
    return (
      <div style={{ width: '100%', height: '100%', background: '#f3f2f1', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, color: '#a19f9d' }}>
        Erro ao carregar
      </div>
    );
  }
  return (
    <div style={{ width: '100%', height: '100%', background: '#f3f2f1', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <img
        src={url}
        alt=""
        style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
        onError={() => setError(true)}
      />
    </div>
  );
};

export interface IPdfTemplateEditorProps {
  value: IPdfTemplateConfig | undefined;
  onChange: (config: IPdfTemplateConfig) => void;
  fieldOptions: IDropdownOption[];
}

export const PdfTemplateEditor: React.FC<IPdfTemplateEditorProps> = ({ value, onChange, fieldOptions }) => {
  const config = ensureTemplate(value);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [dragState, setDragState] = useState<{ id: string; startPx: number; startPy: number; elX: number; elY: number } | null>(null);
  const [resizeState, setResizeState] = useState<{ id: string; handle: string; startPx: number; startPy: number; startW: number; startH: number; startElX: number; startElY: number } | null>(null);
  const canvasRef = useRef<HTMLDivElement>(null);

  const bodyElements = config.body?.elements ?? [];
  const selected = bodyElements.filter((e: IPdfTemplateElement) => e.id === selectedId)[0];

  const updateBodyElements = useCallback(
    (updater: (prev: IPdfTemplateElement[]) => IPdfTemplateElement[]) => {
      const next = updater(bodyElements.slice());
      onChange({
        ...config,
        body: { ...config.body, elements: next },
      });
    },
    [config, bodyElements, onChange]
  );

  const addElement = useCallback(
    (type: 'text' | 'image', scope: TPdfElementScope) => {
      const newEl: IPdfTemplateElement = {
        id: `el_${Date.now()}`,
        type,
        scope,
        x: 30,
        y: 30 + bodyElements.length * 15,
        width: type === 'text' ? 100 : 80,
        height: type === 'image' ? 40 : undefined,
        content: type === 'text' ? (scope === 'dynamic' ? '{{Title}}' : '') : '',
        fontSize: 11,
        fontWeight: 'normal',
      };
      updateBodyElements((prev) => [...prev, newEl]);
      setSelectedId(newEl.id);
    },
    [bodyElements.length, updateBodyElements]
  );

  const deleteSelected = useCallback(() => {
    if (!selectedId) return;
    updateBodyElements((prev) => prev.filter((e) => e.id !== selectedId));
    setSelectedId(null);
  }, [selectedId, updateBodyElements]);

  const updateElement = useCallback(
    (id: string, patch: Partial<IPdfTemplateElement>) => {
      updateBodyElements((prev) =>
        prev.map((e) => (e.id === id ? { ...e, ...patch } : e))
      );
    },
    [updateBodyElements]
  );

  const mmToPx = (mm: number): number => mm * SCALE;
  const pxToMm = (px: number): number => px / SCALE;

  const handleCanvasMouseMove = useCallback(
    (e: React.MouseEvent) => {
      const rect = canvasRef.current?.getBoundingClientRect();
      if (!rect) return;
      const px = e.clientX - rect.left;
      const py = e.clientY - rect.top;
      if (dragState) {
        const dx = pxToMm(px - dragState.startPx);
        const dy = pxToMm(py - dragState.startPy);
        updateElement(dragState.id, { x: dragState.elX + dx, y: dragState.elY + dy });
      } else if (resizeState) {
        const dx = pxToMm(px - resizeState.startPx);
        const dy = pxToMm(py - resizeState.startPy);
        const { handle, startW, startH, startElX, startElY } = resizeState;
        let w = startW;
        let h = startH;
        let x = startElX;
        let y = startElY;
        if (handle.indexOf('e') !== -1) {
          w = Math.max(10, startW + dx);
        }
        if (handle.indexOf('w') !== -1) {
          const nw = Math.max(10, startW - dx);
          x = startElX + (startW - nw);
          w = nw;
        }
        if (handle.indexOf('s') !== -1) h = Math.max(5, startH + dy);
        if (handle.indexOf('n') !== -1) {
          const nh = Math.max(5, startH - dy);
          y = startElY + (startH - nh);
          h = nh;
        }
        updateElement(resizeState.id, { x, y, width: w, height: h });
      }
    },
    [dragState, resizeState, updateElement]
  );

  const handleCanvasMouseUp = useCallback(() => {
    setDragState(null);
    setResizeState(null);
  }, []);

  const handleElementMouseDown = useCallback(
    (e: React.MouseEvent, id: string) => {
      e.stopPropagation();
      e.preventDefault();
      const rect = canvasRef.current?.getBoundingClientRect();
      if (!rect) return;
      const el = bodyElements.filter((x: IPdfTemplateElement) => x.id === id)[0];
      if (!el) return;
      setSelectedId(id);
      setDragState({
        id,
        startPx: e.clientX - rect.left,
        startPy: e.clientY - rect.top,
        elX: el.x,
        elY: el.y,
      });
    },
    [bodyElements]
  );

  const handleResizeStart = useCallback(
    (e: React.MouseEvent, id: string, handle: string) => {
      e.stopPropagation();
      e.preventDefault();
      const rect = canvasRef.current?.getBoundingClientRect();
      if (!rect) return;
      const el = bodyElements.filter((x: IPdfTemplateElement) => x.id === id)[0];
      if (!el) return;
      setResizeState({
        id,
        handle,
        startPx: e.clientX - rect.left,
        startPy: e.clientY - rect.top,
        startW: el.width ?? 50,
        startH: el.height ?? 20,
        startElX: el.x,
        startElY: el.y,
      });
    },
    [bodyElements]
  );

  const insertField = useCallback(
    (fieldKey: string) => {
      if (!selected || fieldKey === '') return;
      const current = selected.content ?? '';
      const placeholder = `{{${fieldKey}}}`;
      updateElement(selected.id, { content: current ? `${current} ${placeholder}` : placeholder });
    },
    [selected, updateElement]
  );

  const PDF_FUNCTIONS: { key: string; label: string }[] = [
    { key: '[now]', label: 'Data atual' },
    { key: '[date]', label: 'Data (pt-BR)' },
    { key: '[time]', label: 'Hora atual' },
    { key: '[nPage]', label: 'Número da página' },
    { key: '[totalPages]', label: 'Total de páginas' },
    { key: '[itemIndex]', label: 'Índice do item (1, 2, 3...)' },
  ];

  const insertPdfFunction = useCallback(
    (token: string) => {
      if (!selected || !token) return;
      const current = selected.content ?? '';
      updateElement(selected.id, { content: current ? `${current} ${token}` : token });
    },
    [selected, updateElement]
  );

  const hasFixedElements = bodyElements.some((e) => e.scope === 'fixed');

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
          Template PDF
        </Text>
        <DefaultButton text="Texto (fixo)" onClick={() => addElement('text', 'fixed')} />
        <DefaultButton text="Texto (dinâmico)" onClick={() => addElement('text', 'dynamic')} />
        <DefaultButton text="Imagem (fixa)" onClick={() => addElement('image', 'fixed')} />
        <DefaultButton text="Imagem (dinâmica)" onClick={() => addElement('image', 'dynamic')} />
        {selectedId && (
          <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" onClick={deleteSelected} />
        )}
      </Stack>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Fixo: uma vez na página. Dinâmico: um por item. Campos: {'{{Campo}}'}. Funções: [now] data, [time] hora, [nPage] página, [totalPages] total, [itemIndex] índice.
      </Text>
      <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { alignItems: 'flex-start' } }}>
        <div
          ref={canvasRef}
          style={{
            width: CANVAS_W,
            height: CANVAS_H,
            background: '#fff',
            border: '1px solid #edebe9',
            boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
            position: 'relative',
            flexShrink: 0,
            userSelect: 'none',
            WebkitUserSelect: 'none',
          }}
          onMouseMove={handleCanvasMouseMove}
          onMouseUp={handleCanvasMouseUp}
          onMouseLeave={handleCanvasMouseUp}
        >
          {hasFixedElements && (config.fixedBlockHeightMm ?? 0) > 0 && (
            <div
              style={{
                position: 'absolute',
                left: 0,
                right: 0,
                top: mmToPx(config.fixedBlockHeightMm ?? 0),
                height: 2,
                background: '#0078d4',
                opacity: 0.6,
                pointerEvents: 'none',
              }}
            />
          )}
          {bodyElements.map((el) => {
            const isSelected = el.id === selectedId;
            const scope = el.scope ?? 'dynamic';
            const left = mmToPx(el.x);
            const top = mmToPx(el.y);
            const w = el.width !== undefined && el.width !== null ? mmToPx(el.width) : 80;
            const h = el.height !== undefined && el.height !== null ? mmToPx(el.height) : 24;
            return (
              <div
                key={el.id}
                style={{
                  position: 'absolute',
                  left,
                  top,
                  width: w,
                  height: el.type === 'text' ? undefined : h,
                  minHeight: el.type === 'text' ? 18 : h,
                  border: isSelected ? '2px solid #0078d4' : scope === 'fixed' ? '1px dashed #107c10' : '1px dashed #a19f9d',
                  background: el.type === 'rect' ? (el.color ?? '#f3f2f1') : scope === 'fixed' ? 'rgba(16,124,16,0.06)' : 'transparent',
                  cursor: 'move',
                  fontSize: (el.fontSize ?? 11) * (SCALE * 0.6),
                  fontWeight: el.fontWeight ?? 'normal',
                  overflow: 'hidden',
                  padding: 2,
                  boxSizing: 'border-box',
                }}
                onMouseDown={(e) => handleElementMouseDown(e, el.id)}
              >
                {el.type === 'text' && (el.content ?? '').replace(/\{\{([^}]+)\}\}/g, '[$1]')}
                {el.type === 'image' && (
                  <PdfTemplateImagePreview url={(el.imageUrl ?? el.content ?? '').trim()} />
                )}
                {el.type === 'line' && <div style={{ width: '100%', height: 2, background: '#333', marginTop: (h - 2) / 2 }} />}
                {isSelected && (
                  <>
                    <div style={{ position: 'absolute', right: -4, top: '50%', marginTop: -6, width: 8, height: 12, background: '#0078d4', cursor: 'ew-resize' }} onMouseDown={(e) => handleResizeStart(e, el.id, 'e')} />
                    <div style={{ position: 'absolute', bottom: -4, left: '50%', marginLeft: -6, width: 12, height: 8, background: '#0078d4', cursor: 'ns-resize' }} onMouseDown={(e) => handleResizeStart(e, el.id, 's')} />
                  </>
                )}
              </div>
            );
          })}
        </div>
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { minWidth: 280 } }}>
          <ChoiceGroup
            label="Página"
            options={[
              { key: 'A4', text: 'A4' },
              { key: 'Letter', text: 'Letter' },
            ] as IChoiceGroupOption[]}
            selectedKey={config.pageFormat}
            onChange={(_, o) => o && onChange({ ...config, pageFormat: o.key as 'A4' | 'Letter' })}
          />
          <ChoiceGroup
            label="Orientação"
            options={[
              { key: 'portrait', text: 'Retrato' },
              { key: 'landscape', text: 'Paisagem' },
            ] as IChoiceGroupOption[]}
            selectedKey={config.orientation}
            onChange={(_, o) => o && onChange({ ...config, orientation: o.key as 'portrait' | 'landscape' })}
          />
          <ChoiceGroup
            label="Layout dos dados"
            options={[
              { key: 'onePerPage', text: 'Uma página por item' },
              { key: 'allOnOnePage', text: 'Todos na mesma página' },
              { key: 'breakWhenFull', text: 'Quebrar página ao atingir o limite' },
            ] as IChoiceGroupOption[]}
            selectedKey={config.layoutMode ?? 'onePerPage'}
            onChange={(_, o) => o && onChange({ ...config, layoutMode: (o.key as TPdfLayoutMode) ?? 'onePerPage' })}
          />
          {(config.layoutMode === 'allOnOnePage' || config.layoutMode === 'breakWhenFull') && (
            <TextField
              label="Altura por item (mm)"
              type="number"
              description="Espaço vertical ocupado por cada item no PDF"
              value={config.bodyBlockHeightMm !== undefined && config.bodyBlockHeightMm !== null ? String(config.bodyBlockHeightMm) : '40'}
              onChange={(_, v) => onChange({ ...config, bodyBlockHeightMm: v === '' ? undefined : Number(v) || 40 })}
              styles={{ root: { maxWidth: 100 } }}
            />
          )}
          {hasFixedElements && (
            <TextField
              label="Altura da área fixa (mm)"
              type="number"
              description="Espaço no topo reservado aos elementos fixos (logo, nome da empresa)"
              value={config.fixedBlockHeightMm !== undefined && config.fixedBlockHeightMm !== null ? String(config.fixedBlockHeightMm) : ''}
              onChange={(_, v) => onChange({ ...config, fixedBlockHeightMm: v === '' ? undefined : Number(v) || 0 })}
              styles={{ root: { maxWidth: 100 } }}
            />
          )}
          {selected && (
            <>
              <Separator />
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Propriedades</Text>
              <ChoiceGroup
                label="Área"
                options={[
                  { key: 'fixed', text: 'Fixo (uma vez na página)' },
                  { key: 'dynamic', text: 'Dinâmico (um por item)' },
                ] as IChoiceGroupOption[]}
                selectedKey={selected.scope ?? 'dynamic'}
                onChange={(_, o) => o && updateElement(selected.id, { scope: o.key as TPdfElementScope })}
              />
              {selected.type === 'image' ? (
                <TextField
                  label="URL da imagem"
                  placeholder="https://..."
                  value={(selected.imageUrl ?? selected.content ?? '')}
                  onChange={(_, v) => updateElement(selected.id, { imageUrl: v ?? '', content: v ?? '' })}
                  description="Cole o link da imagem para ver no preview"
                />
              ) : (
                <>
                  <TextField
                    label="Conteúdo"
                    value={selected.content ?? ''}
                    onChange={(_, v) => updateElement(selected.id, { content: v ?? '' })}
                    multiline={selected.type === 'text'}
                  />
                  <Dropdown
                    label="Inserir campo"
                    placeholder="Selecione um campo"
                    options={fieldOptions.filter((o) => (o.key as string) !== '')}
                    onChange={(_, o) => o && insertField(String(o.key))}
                  />
                  <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>Inserir função</Text>
                  <Stack horizontal tokens={{ childrenGap: 4 }} styles={{ root: { flexWrap: 'wrap' } }}>
                    {PDF_FUNCTIONS.map((f) => (
                      <DefaultButton
                        key={f.key}
                        text={f.key}
                        title={f.label}
                        onClick={() => insertPdfFunction(f.key)}
                        styles={{ root: { minWidth: 'auto', padding: '0 8px' } }}
                      />
                    ))}
                  </Stack>
                </>
              )}
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField
                  label="X (mm)"
                  type="number"
                  value={String(selected.x)}
                  onChange={(_, v) => updateElement(selected.id, { x: Number(v) || 0 })}
                  styles={{ root: { width: 70 } }}
                />
                <TextField
                  label="Y (mm)"
                  type="number"
                  value={String(selected.y)}
                  onChange={(_, v) => updateElement(selected.id, { y: Number(v) || 0 })}
                  styles={{ root: { width: 70 } }}
                />
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField
                  label="Largura"
                  type="number"
                  value={selected.width !== undefined && selected.width !== null ? String(selected.width) : ''}
                  onChange={(_, v) => updateElement(selected.id, { width: v === '' ? undefined : Number(v) || 0 })}
                  styles={{ root: { width: 70 } }}
                />
                <TextField
                  label="Altura"
                  type="number"
                  value={selected.height !== undefined && selected.height !== null ? String(selected.height) : ''}
                  onChange={(_, v) => updateElement(selected.id, { height: v === '' ? undefined : Number(v) || 0 })}
                  styles={{ root: { width: 70 } }}
                />
              </Stack>
              {selected.type === 'text' && (
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <TextField
                    label="Fonte"
                    type="number"
                    value={String(selected.fontSize ?? 11)}
                    onChange={(_, v) => updateElement(selected.id, { fontSize: Number(v) || 11 })}
                    styles={{ root: { width: 60 } }}
                  />
                  <Dropdown
                    label="Peso"
                    options={[
                      { key: 'normal', text: 'Normal' },
                      { key: 'bold', text: 'Negrito' },
                    ]}
                    selectedKey={selected.fontWeight ?? 'normal'}
                    onChange={(_, o) => o && updateElement(selected.id, { fontWeight: o.key as 'normal' | 'bold' })}
                    styles={{ root: { width: 100 } }}
                  />
                </Stack>
              )}
            </>
          )}
        </Stack>
      </Stack>
    </Stack>
  );
};

function Separator(): React.ReactElement {
  return <div style={{ height: 1, background: '#edebe9', margin: '4px 0' }} />;
}
