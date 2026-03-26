import { jsPDF } from 'jspdf';
import type { IPdfTemplateConfig, IPdfTemplateElement } from '../config/types';
import { getPdfPageSizeMm, toJsPdfFormat } from './pdfPageFormats';

const PT_TO_MM = 0.352778;
const TEXT_LINE_HEIGHT_FACTOR = 1.15;

function getImageUrl(el: IPdfTemplateElement): string {
  const url = (el.imageUrl ?? el.content ?? '').trim();
  return url;
}

export interface ILoadedImage {
  dataUrl: string;
  width: number;
  height: number;
}

function loadImageAsDataUrl(url: string): Promise<ILoadedImage | null> {
  if (!url) return Promise.resolve(null);
  if (url.indexOf('data:') === 0) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        resolve({ dataUrl: url, width: img.naturalWidth, height: img.naturalHeight });
      };
      img.onerror = () => resolve(null);
      img.src = url;
    });
  }
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      try {
        const canvas = typeof document !== 'undefined' ? document.createElement('canvas') : null;
        if (!canvas) { resolve(null); return; }
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        const ctx = canvas.getContext('2d');
        if (!ctx) { resolve(null); return; }
        ctx.drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL('image/png');
        resolve({ dataUrl, width: img.naturalWidth, height: img.naturalHeight });
      } catch {
        resolve(null);
      }
    };
    img.onerror = () => resolve(null);
    img.src = url;
  });
}

function getDataUrlFormat(dataUrl: string): 'PNG' | 'JPEG' {
  if (dataUrl.indexOf('data:image/png') === 0) return 'PNG';
  if (dataUrl.indexOf('data:image/jpeg') === 0 || dataUrl.indexOf('data:image/jpg') === 0) return 'JPEG';
  return 'PNG';
}

function getByPath(obj: unknown, path: string): unknown {
  if (obj === null || obj === undefined) return undefined;
  const parts = path.replace(/\//g, '.').split('.');
  let current: unknown = obj;
  for (let i = 0; i < parts.length; i++) {
    if (current === null || current === undefined || typeof current !== 'object') return undefined;
    current = (current as Record<string, unknown>)[parts[i]];
  }
  return current;
}

function valueToText(val: unknown): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && !Array.isArray(val) && val !== null) {
    const o = val as Record<string, unknown>;
    if (o.Title !== undefined) return String(o.Title);
    if (o.DisplayName !== undefined) return String(o.DisplayName);
    if (o.Email !== undefined) return String(o.Email);
  }
  if (Array.isArray(val)) {
    return val.map((v) => valueToText(v)).filter(Boolean).join(', ');
  }
  return String(val);
}

export interface IPdfRenderContext {
  pageNumber?: number;
  totalPages?: number;
  itemIndex?: number;
}

function replacePlaceholders(
  content: string | undefined,
  item: Record<string, unknown>
): string {
  if (!content) return '';
  return content.replace(/\{\{([^}]+)\}\}/g, (_match, fieldName: string) => {
    const trimmed = fieldName.trim();
    const val = getByPath(item, trimmed);
    return valueToText(val);
  });
}

function replacePdfFunctions(content: string, ctx: IPdfRenderContext): string {
  const now = new Date();
  const dateStr = now.toLocaleDateString('pt-BR');
  const timeStr = now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  return content
    .replace(/\[now\]/gi, dateStr)
    .replace(/\[date\]/gi, dateStr)
    .replace(/\[time\]/gi, timeStr)
    .replace(/\[nPage\]/g, String(ctx.pageNumber ?? 1))
    .replace(/\[totalPages\]/g, String(ctx.totalPages ?? 1))
    .replace(/\[itemIndex\]/g, String(ctx.itemIndex ?? 1));
}

function getTextLineHeightMm(fontSizePt: number): number {
  return fontSizePt * PT_TO_MM * TEXT_LINE_HEIGHT_FACTOR;
}

function splitTextToLines(
  doc: jsPDF,
  text: string,
  maxWidth: number | undefined
): string[] {
  if (!text) return [];
  if (maxWidth !== undefined && maxWidth > 0) {
    const lines = doc.splitTextToSize(text, maxWidth);
    return Array.isArray(lines) ? lines.map((line) => String(line)) : [String(lines)];
  }
  return String(text).split('\n');
}

function truncateLinesToHeight(
  lines: string[],
  lineHeightMm: number,
  maxHeightMm: number | undefined
): string[] {
  if (!lines.length) return lines;
  if (maxHeightMm === undefined || maxHeightMm <= 0) return lines;
  const maxLines = Math.max(1, Math.floor(maxHeightMm / lineHeightMm));
  if (lines.length <= maxLines) return lines;
  const next = lines.slice(0, maxLines);
  const lastIdx = next.length - 1;
  const last = next[lastIdx].replace(/\s+$/, '');
  next[lastIdx] = last.length > 0 ? `${last}...` : '...';
  return next;
}

function getDefaultTextHeightMm(fontSizePt: number): number {
  return getTextLineHeightMm(fontSizePt) + 2;
}

function getElementHeightMm(el: IPdfTemplateElement): number {
  if (el.height !== undefined && el.height !== null && el.height > 0) return el.height;
  if (el.type === 'text') return getDefaultTextHeightMm(el.fontSize ?? 11);
  if (el.type === 'line') return 1;
  if (el.type === 'image') return 30;
  return 10;
}

function getImageDrawBox(
  el: IPdfTemplateElement,
  loaded: ILoadedImage
): { x: number; y: number; width: number; height: number } {
  const hasWidth = el.width !== undefined && el.width !== null && el.width > 0;
  const hasHeight = el.height !== undefined && el.height !== null && el.height > 0;
  if (hasWidth && hasHeight) {
    return {
      x: el.x,
      y: el.y,
      width: el.width as number,
      height: el.height as number,
    };
  }
  if (hasWidth) {
    const width = el.width as number;
    const height = loaded.width > 0 ? (loaded.height / loaded.width) * width : 30;
    return { x: el.x, y: el.y, width, height };
  }
  if (hasHeight) {
    const height = el.height as number;
    const width = loaded.height > 0 ? (loaded.width / loaded.height) * height : 40;
    return { x: el.x, y: el.y, width, height };
  }
  return { x: el.x, y: el.y, width: 40, height: 30 };
}

function renderSection(
  doc: jsPDF,
  elements: IPdfTemplateElement[],
  item: Record<string, unknown>,
  offsetY: number,
  imageDataMap: Record<string, ILoadedImage | null>,
  ctx: IPdfRenderContext
): void {
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    const x = el.x;
    const y = el.y + offsetY;
    if (el.type === 'text') {
      let text = replacePlaceholders(el.content, item);
      text = replacePdfFunctions(text, ctx);
      if (text) {
        const fontSize = el.fontSize ?? 11;
        const width = el.width !== undefined && el.width !== null && el.width > 0 ? el.width : undefined;
        const lineHeightMm = getTextLineHeightMm(fontSize);
        const maxHeightMm = el.height !== undefined && el.height !== null && el.height > 0 ? el.height : undefined;
        const lines = truncateLinesToHeight(splitTextToLines(doc, text, width), lineHeightMm, maxHeightMm);
        if (lines.length === 0) continue;
        doc.setFontSize(fontSize);
        doc.setFont('helvetica', el.fontWeight === 'bold' ? 'bold' : 'normal');
        if (el.color) doc.setTextColor(el.color);
        doc.text(lines, x, y, { ...(width !== undefined ? { maxWidth: width } : {}), baseline: 'top' as never });
        if (el.color) doc.setTextColor(0, 0, 0);
      }
    } else if (el.type === 'image') {
      const url = getImageUrl(el);
      if (!url) continue;
      const loaded = imageDataMap[url] ?? null;
      if (!loaded) continue;
      const drawBox = getImageDrawBox({ ...el, y }, loaded);
      try {
        const format = getDataUrlFormat(loaded.dataUrl);
        doc.addImage(loaded.dataUrl, format, drawBox.x, drawBox.y, drawBox.width, drawBox.height);
      } catch {
        // ignore
      }
    } else if (el.type === 'rect' && el.width !== undefined && el.height !== undefined) {
      doc.setFillColor(el.color ?? '#f0f0f0');
      doc.rect(x, y, el.width, el.height, 'F');
    } else if (el.type === 'line' && el.width !== undefined) {
      if (el.color) doc.setDrawColor(el.color);
      else doc.setDrawColor(0, 0, 0);
      doc.line(x, y, x + el.width, y);
    }
  }
}

function splitBodyByScope(elements: IPdfTemplateElement[]): { fixed: IPdfTemplateElement[]; dynamic: IPdfTemplateElement[] } {
  const fixed: IPdfTemplateElement[] = [];
  const dynamic: IPdfTemplateElement[] = [];
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    if (el.scope === 'fixed') fixed.push(el);
    else dynamic.push(el);
  }
  return { fixed, dynamic };
}

function getFixedBlockHeightMm(template: IPdfTemplateConfig, fixedElements: IPdfTemplateElement[]): number {
  if (template.fixedBlockHeightMm !== undefined && template.fixedBlockHeightMm !== null && template.fixedBlockHeightMm > 0) {
    return template.fixedBlockHeightMm;
  }
  let max = 0;
  for (let i = 0; i < fixedElements.length; i++) {
    const el = fixedElements[i];
    const bottom = el.y + getElementHeightMm(el);
    if (bottom > max) max = bottom;
  }
  return max + 5;
}

function getBodyBlockHeightMm(template: IPdfTemplateConfig, dynamicElements: IPdfTemplateElement[]): number {
  if (template.bodyBlockHeightMm !== undefined && template.bodyBlockHeightMm !== null && template.bodyBlockHeightMm > 0) {
    return template.bodyBlockHeightMm;
  }
  let max = 40;
  for (let i = 0; i < dynamicElements.length; i++) {
    const el = dynamicElements[i];
    const bottom = el.y + getElementHeightMm(el);
    if (bottom > max) max = bottom;
  }
  return max + 5;
}

function collectImageUrls(elements: IPdfTemplateElement[]): string[] {
  const urls: string[] = [];
  for (let i = 0; i < elements.length; i++) {
    if (elements[i].type !== 'image') continue;
    const url = getImageUrl(elements[i]);
    if (url && urls.indexOf(url) === -1) urls.push(url);
  }
  return urls;
}

export async function generateAndDownloadPdf(
  template: IPdfTemplateConfig,
  items: Record<string, unknown>[],
  filename: string = 'export.pdf'
): Promise<void> {
  const orientation = template.orientation === 'landscape' ? 'l' : 'p';
  const format = toJsPdfFormat(template.pageFormat);
  const doc = new jsPDF({ orientation, unit: 'mm', format });
  const pageSizeMm = getPdfPageSizeMm(
    template.pageFormat,
    template.orientation === 'landscape' ? 'landscape' : 'portrait'
  );
  const pageHeightMm = pageSizeMm.heightMm;

  const bodyElements = template.body?.elements ?? [];
  const { fixed: fixedBodyElements, dynamic: dynamicBodyElements } = splitBodyByScope(bodyElements);
  if (bodyElements.length === 0) {
    doc.text('Nenhum conteúdo no template.', 20, 20);
    doc.save(filename);
    return;
  }

  const headerElements = template.header?.elements ?? [];
  const footerElements = template.footer?.elements ?? [];
  const allUrls = collectImageUrls(headerElements).concat(
    collectImageUrls(bodyElements),
    collectImageUrls(footerElements)
  );
  const uniqueUrls: string[] = [];
  for (let u = 0; u < allUrls.length; u++) {
    if (uniqueUrls.indexOf(allUrls[u]) === -1) uniqueUrls.push(allUrls[u]);
  }
  const imageDataMap: Record<string, ILoadedImage | null> = {};
  for (let i = 0; i < uniqueUrls.length; i++) {
    imageDataMap[uniqueUrls[i]] = await loadImageAsDataUrl(uniqueUrls[i]);
  }

  const layoutMode = template.layoutMode ?? 'onePerPage';
  const headerHeight = template.header?.height ?? 0;
  const footerHeight = template.footer?.height ?? 0;
  const marginBottom = footerHeight + 15;
  const fixedBlockHeight = getFixedBlockHeightMm(template, fixedBodyElements);
  const bodyBlockHeight = getBodyBlockHeightMm(template, dynamicBodyElements);
  const firstItem = items.length > 0 ? (items[0] as Record<string, unknown>) : ({} as Record<string, unknown>);

  let totalPages = 1;
  if (layoutMode === 'onePerPage') totalPages = Math.max(1, items.length);
  else if (layoutMode === 'allOnOnePage') totalPages = 1;
  else {
    let cp = headerHeight + (fixedBodyElements.length > 0 ? fixedBlockHeight : 0);
    for (let i = 0; i < items.length; i++) {
      if (cp + bodyBlockHeight > pageHeightMm - marginBottom) {
        totalPages += 1;
        cp = headerHeight + (fixedBodyElements.length > 0 ? fixedBlockHeight : 0);
      }
      cp += bodyBlockHeight;
    }
  }

  const ctx = (pageNum: number, itemIdx?: number): IPdfRenderContext => ({
    pageNumber: pageNum,
    totalPages,
    itemIndex: itemIdx !== undefined ? itemIdx + 1 : 1,
  });

  if (layoutMode === 'onePerPage') {
    for (let idx = 0; idx < items.length; idx++) {
      if (idx > 0) doc.addPage(format, orientation);
      const pageNum = idx + 1;
      const item = items[idx] as Record<string, unknown>;
      let offsetY = 0;
      if (headerElements.length > 0) {
        renderSection(doc, headerElements, item, offsetY, imageDataMap, ctx(pageNum, idx));
        offsetY += headerHeight;
      }
      if (fixedBodyElements.length > 0) {
        renderSection(doc, fixedBodyElements, firstItem, offsetY, imageDataMap, ctx(pageNum));
        offsetY += fixedBlockHeight;
      }
      renderSection(doc, dynamicBodyElements, item, offsetY, imageDataMap, ctx(pageNum, idx));
      if (footerElements.length > 0) {
        const footerY = pageHeightMm - footerHeight - 10;
        renderSection(doc, footerElements, item, footerY, imageDataMap, ctx(pageNum, idx));
      }
    }
  } else if (layoutMode === 'allOnOnePage') {
    const pageNum = 1;
    let offsetY = 0;
    if (headerElements.length > 0 && items.length > 0) {
      renderSection(doc, headerElements, firstItem, 0, imageDataMap, ctx(pageNum));
      offsetY += headerHeight;
    }
    if (fixedBodyElements.length > 0) {
      renderSection(doc, fixedBodyElements, firstItem, offsetY, imageDataMap, ctx(pageNum));
      offsetY += fixedBlockHeight;
    }
    for (let idx = 0; idx < items.length; idx++) {
      const item = items[idx] as Record<string, unknown>;
      renderSection(doc, dynamicBodyElements, item, offsetY, imageDataMap, ctx(pageNum, idx));
      offsetY += bodyBlockHeight;
    }
    if (footerElements.length > 0 && items.length > 0) {
      const footerY = pageHeightMm - footerHeight - 10;
      renderSection(doc, footerElements, firstItem, footerY, imageDataMap, ctx(pageNum));
    }
  } else {
    let currentPage = 1;
    let currentPageOffsetY = headerHeight;
    if (fixedBodyElements.length > 0) currentPageOffsetY += fixedBlockHeight;
    let pageStarted = false;
    let lastItemOnPage = items[0] as Record<string, unknown>;
    const footerY = pageHeightMm - footerHeight - 10;
    let fixedRenderedOnCurrentPage = false;
    for (let idx = 0; idx < items.length; idx++) {
      if (currentPageOffsetY + bodyBlockHeight > pageHeightMm - marginBottom) {
        if (footerElements.length > 0) renderSection(doc, footerElements, lastItemOnPage, footerY, imageDataMap, ctx(currentPage));
        doc.addPage(format, orientation);
        currentPage += 1;
        currentPageOffsetY = headerHeight;
        pageStarted = false;
        fixedRenderedOnCurrentPage = false;
      }
      const item = items[idx] as Record<string, unknown>;
      lastItemOnPage = item;
      if (headerElements.length > 0 && !pageStarted) {
        renderSection(doc, headerElements, item, 0, imageDataMap, ctx(currentPage, idx));
        pageStarted = true;
      }
      if (fixedBodyElements.length > 0 && !fixedRenderedOnCurrentPage) {
        renderSection(doc, fixedBodyElements, firstItem, headerHeight, imageDataMap, ctx(currentPage));
        fixedRenderedOnCurrentPage = true;
      }
      renderSection(doc, dynamicBodyElements, item, currentPageOffsetY, imageDataMap, ctx(currentPage, idx));
      currentPageOffsetY += bodyBlockHeight;
    }
    if (footerElements.length > 0 && items.length > 0) {
      renderSection(doc, footerElements, lastItemOnPage, footerY, imageDataMap, ctx(currentPage));
    }
  }

  doc.save(filename);
}
