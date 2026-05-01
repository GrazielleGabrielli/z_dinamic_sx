import { DayOfWeek } from '@fluentui/date-time-utilities';
import type { IDatePickerStrings } from '@fluentui/react';

const pad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

export function formatDatePtBr(date?: Date): string {
  if (!date) return '';
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
}

export function parseDateFromStringPtBr(value: string): Date | undefined {
  const t = value.trim();
  if (!t) return undefined;
  const parts = t.split(/[/\-.]/).map((p) => p.trim());
  if (parts.length !== 3) return undefined;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  let year = parseInt(parts[2], 10);
  if (Number.isNaN(day) || Number.isNaN(month) || Number.isNaN(year)) return undefined;
  if (parts[2].length <= 2) {
    year += year >= 70 ? 1900 : 2000;
  }
  const dt = new Date(year, month, day);
  if (dt.getFullYear() !== year || dt.getMonth() !== month || dt.getDate() !== day) return undefined;
  return dt;
}

export const DATE_PICKER_STRINGS_PT_BR: IDatePickerStrings = {
  months: [
    'janeiro',
    'fevereiro',
    'março',
    'abril',
    'maio',
    'junho',
    'julho',
    'agosto',
    'setembro',
    'outubro',
    'novembro',
    'dezembro',
  ],
  shortMonths: ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'],
  days: ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'],
  shortDays: ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'],
  goToToday: 'Ir para hoje',
  weekNumberFormatString: 'Semana {0}',
  prevMonthAriaLabel: 'Mês anterior',
  nextMonthAriaLabel: 'Próximo mês',
  prevYearAriaLabel: 'Ano anterior',
  nextYearAriaLabel: 'Próximo ano',
  prevYearRangeAriaLabel: 'Intervalo de anos anterior',
  nextYearRangeAriaLabel: 'Próximo intervalo de anos',
  closeButtonAriaLabel: 'Fechar',
  selectedDateFormatString: 'Data selecionada {0}',
  todayDateFormatString: 'Data de hoje {0}',
  monthPickerHeaderAriaLabel: '{0}, alterar ano',
  yearPickerHeaderAriaLabel: '{0}, alterar mês',
  dayMarkedAriaLabel: 'marcado',
  isRequiredErrorMessage: 'Campo obrigatório',
  invalidInputErrorMessage: 'Formato de data inválido (use dd/mm/aaaa)',
  isOutOfBoundsErrorMessage: 'Data fora do intervalo permitido',
  isResetStatusMessage: 'Entrada inválida "{0}", data restaurada para "{1}"',
};

/* eslint-disable @rushstack/no-new-null */
function parseDateFromStringForFluent(dateStr: string): Date | null {
  const d = parseDateFromStringPtBr(dateStr);
  return d === undefined ? null : d;
}
/* eslint-enable @rushstack/no-new-null */

export const FLUENT_DATE_PICKER_PT_BR = {
  firstDayOfWeek: DayOfWeek.Monday,
  formatDate: formatDatePtBr,
  parseDateFromString: parseDateFromStringForFluent,
  strings: DATE_PICKER_STRINGS_PT_BR,
  allowTextInput: true,
  placeholder: 'dd/mm/aaaa',
};
