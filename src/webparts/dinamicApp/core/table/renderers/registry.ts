import type { InternalFieldType, TTableRenderer } from '../types';
import { textRenderer } from './textRenderer';
import { noteRenderer } from './noteRenderer';
import { numberRenderer } from './numberRenderer';
import { currencyRenderer } from './currencyRenderer';
import { dateRenderer } from './dateRenderer';
import { booleanRenderer } from './booleanRenderer';
import { choiceRenderer } from './choiceRenderer';
import { multiChoiceRenderer } from './multiChoiceRenderer';
import { lookupRenderer } from './lookupRenderer';
import { lookupMultiRenderer } from './lookupMultiRenderer';
import { userRenderer } from './userRenderer';
import { userMultiRenderer } from './userMultiRenderer';
import { urlRenderer } from './urlRenderer';
import { fileRenderer } from './fileRenderer';
import { managedMetadataRenderer } from './managedMetadataRenderer';
import { calculatedRenderer } from './calculatedRenderer';
import { imageRenderer } from './imageRenderer';
import { unknownRenderer } from './unknownRenderer';

const REGISTRY: Partial<Record<InternalFieldType, TTableRenderer>> = {
  text: textRenderer,
  note: noteRenderer,
  number: numberRenderer,
  currency: currencyRenderer,
  date: dateRenderer,
  boolean: booleanRenderer,
  choice: choiceRenderer,
  multiChoice: multiChoiceRenderer,
  lookup: lookupRenderer,
  lookupMulti: lookupMultiRenderer,
  user: userRenderer,
  userMulti: userMultiRenderer,
  url: urlRenderer,
  file: fileRenderer,
  managedMetadata: managedMetadataRenderer,
  calculated: calculatedRenderer,
  image: imageRenderer,
  unknown: unknownRenderer,
};

export function getRenderer(fieldType: InternalFieldType | undefined): TTableRenderer {
  if (fieldType && REGISTRY[fieldType]) return REGISTRY[fieldType] as TTableRenderer;
  return unknownRenderer;
}

export function registerRenderer(fieldType: InternalFieldType, renderer: TTableRenderer): void {
  (REGISTRY as Record<InternalFieldType, TTableRenderer>)[fieldType] = renderer;
}
