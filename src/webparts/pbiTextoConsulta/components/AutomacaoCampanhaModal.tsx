import * as React from 'react';
import { useEffect, useState } from 'react';
import { createAutomacaoCampanhaItem } from './automacaoCampanhaService';
import type { AutomacaoCampanhaFormData, AutomacaoCampanhaFormErrors } from './automacaoCampanhaTypes';
import { ENVIAR_PARA_OPCOES, TIPO_CAMPANHA_OPCOES } from './automacaoCampanhaUtils';

export interface AutomacaoCampanhaModalProps {
  isOpen: boolean;
  initialValues: AutomacaoCampanhaFormData;
  onClose: () => void;
  onSaved?: (itemId: number) => void | Promise<void>;
}

const emptyErrors: AutomacaoCampanhaFormErrors = {};

const mergeUniqueOptions = (base: string[], currentValue: string): string[] => {
  const combined: string[] = base.slice();
  if (currentValue) {
    combined.push(currentValue);
  }
  const result: string[] = [];
  for (let i = 0; i < combined.length; i++) {
    const item = combined[i];
    if (!item) {
      continue;
    }
    if (result.indexOf(item) === -1) {
      result.push(item);
    }
  }
  return result;
};

const AutomacaoCampanhaModal = ({
  isOpen,
  initialValues,
  onClose,
  onSaved
}: AutomacaoCampanhaModalProps): React.ReactElement | null => {
  const [form, setForm] = useState<AutomacaoCampanhaFormData>(initialValues);
  const [errors, setErrors] = useState<AutomacaoCampanhaFormErrors>(emptyErrors);
  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string>('');

  useEffect(() => {
    if (isOpen) {
      const text =
        initialValues.TextoConsulta.trim().length > 0
          ? initialValues.TextoConsulta
          : initialValues.texto_regra;
      setForm({
        ...initialValues,
        TextoConsulta: text,
        texto_regra: text
      });
      setErrors(emptyErrors);
      setSaveError('');
    }
  }, [isOpen, initialValues]);

  useEffect(() => {
    if (!isOpen) {
      return undefined;
    }

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        onClose();
      }
    };

    window.addEventListener('keydown', onKeyDown);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const handleChange =
    (field: keyof AutomacaoCampanhaFormData) =>
    (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>): void => {
      const { value } = event.target;
      setForm((current) => ({ ...current, [field]: value }));
      setErrors((current) => ({ ...current, [field]: undefined }));
    };

  const handleTextoConsultaRegraChange = (event: React.ChangeEvent<HTMLTextAreaElement>): void => {
    const { value } = event.target;
    setForm((current) => ({
      ...current,
      TextoConsulta: value,
      texto_regra: value
    }));
    setErrors((current) => ({
      ...current,
      TextoConsulta: undefined,
      texto_regra: undefined
    }));
  };

  const validate = (): boolean => {
    const next: AutomacaoCampanhaFormErrors = {};

    if (!form.Title.trim()) {
      next.Title = 'Informe o titulo.';
    }

    if (!form.TextoConsulta.trim()) {
      next.TextoConsulta = 'Texto de consulta obrigatorio.';
    }

    setErrors(next);
    return Object.keys(next).length === 0;
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    setSaveError('');

    if (!validate()) {
      return;
    }

    try {
      setIsSaving(true);
      const synced = {
        ...form,
        TextoConsulta: form.TextoConsulta.trim(),
        texto_regra: form.TextoConsulta.trim()
      };
      const id = await createAutomacaoCampanhaItem(synced);
      await Promise.resolve(onSaved?.(id));
      onClose();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Falha ao salvar.';
      setSaveError(message);
    } finally {
      setIsSaving(false);
    }
  };

  const inputClass =
    'w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none transition focus:border-indigo-400 focus:ring-4 focus:ring-indigo-100';
  const labelClass = 'text-sm font-medium text-slate-800';

  const tipoOptions = mergeUniqueOptions(TIPO_CAMPANHA_OPCOES, form.Tipo_campanha);
  const enviarOptions = mergeUniqueOptions(ENVIAR_PARA_OPCOES, form.EnviarPara);

  return (
    <div
      className="fixed inset-0 z-[1001] flex items-start justify-center overflow-y-auto bg-slate-900/50 p-4 py-10 sm:p-6 sm:py-12"
      role="dialog"
      aria-modal="true"
      aria-labelledby="automacao-modal-title"
      onClick={onClose}
    >
      <div className="relative w-full max-w-2xl rounded-[28px] border border-slate-200 bg-white shadow-2xl" onClick={(event) => event.stopPropagation()}>
        <div className="flex items-start justify-between gap-4 border-b border-slate-100 px-6 py-5 sm:px-8">
          <div>
            <h2 id="automacao-modal-title" className="text-xl font-semibold text-slate-900">
              Automacao de campanha
            </h2>
            <p className="mt-1 text-sm text-slate-500">Revise os dados e conclua o cadastro na lista AutomacaoCampanhas.</p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="inline-flex min-h-[44px] min-w-[44px] shrink-0 items-center justify-center rounded-xl border border-slate-200 text-lg font-semibold text-slate-600 hover:bg-slate-50"
            aria-label="Fechar"
          >
            ×
          </button>
        </div>

        <form className="max-h-[min(70vh,720px)] overflow-y-auto px-6 py-6 sm:px-8" onSubmit={(e) => void handleSubmit(e)}>
          <div className="grid w-full grid-cols-12 gap-6">
            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-title">
                Titulo
              </label>
              <input id="ac-title" type="text" value={form.Title} onChange={handleChange('Title')} className={inputClass} />
              {errors.Title && <p className="text-sm text-red-600">{errors.Title}</p>}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-texto-consulta-regra">
                Texto consulta / texto_regra
              </label>
              <textarea
                id="ac-texto-consulta-regra"
                rows={8}
                value={form.TextoConsulta}
                onChange={handleTextoConsultaRegraChange}
                className={`${inputClass} resize-y font-mono text-xs leading-relaxed`}
              />
              {(errors.TextoConsulta || errors.texto_regra) && (
                <p className="text-sm text-red-600">{errors.TextoConsulta || errors.texto_regra}</p>
              )}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-desc">
                descricao_campanha
              </label>
              <textarea id="ac-desc" rows={4} value={form.descricao_campanha} onChange={handleChange('descricao_campanha')} className={`${inputClass} resize-y`} />
              {errors.descricao_campanha && <p className="text-sm text-red-600">{errors.descricao_campanha}</p>}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-inicio">
                Inicio
              </label>
              <input id="ac-inicio" type="datetime-local" value={form.Inicio} onChange={handleChange('Inicio')} className={inputClass} />
              {errors.Inicio && <p className="text-sm text-red-600">{errors.Inicio}</p>}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-fim">
                Fim
              </label>
              <input id="ac-fim" type="datetime-local" value={form.Fim} onChange={handleChange('Fim')} className={inputClass} />
              {errors.Fim && <p className="text-sm text-red-600">{errors.Fim}</p>}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-tipo">
                Tipo_campanha
              </label>
              <select id="ac-tipo" value={form.Tipo_campanha} onChange={handleChange('Tipo_campanha')} className={inputClass}>
                <option value="">Selecione</option>
                {tipoOptions.map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
              </select>
              {errors.Tipo_campanha && <p className="text-sm text-red-600">{errors.Tipo_campanha}</p>}
            </div>

            <div className="col-span-12 space-y-2">
              <label className={labelClass} htmlFor="ac-enviar">
                EnviarPara
              </label>
              <select id="ac-enviar" value={form.EnviarPara} onChange={handleChange('EnviarPara')} className={inputClass}>
                <option value="">Selecione</option>
                {enviarOptions.map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
              </select>
              {errors.EnviarPara && <p className="text-sm text-red-600">{errors.EnviarPara}</p>}
            </div>
          </div>

          {saveError && (
            <div className="mt-6 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
              {saveError}
            </div>
          )}

          <div className="mt-8 flex flex-col-reverse gap-3 border-t border-slate-100 pt-6 sm:flex-row sm:justify-end">
            <button
              type="button"
              onClick={onClose}
              disabled={isSaving}
              className="inline-flex min-h-[48px] items-center justify-center rounded-xl border border-slate-300 bg-white px-8 py-3 text-sm font-semibold text-slate-800 hover:bg-slate-50 disabled:opacity-50"
            >
              Cancelar
            </button>
            <button
              type="submit"
              disabled={isSaving}
              className="inline-flex min-h-[48px] items-center justify-center rounded-xl bg-indigo-600 px-8 py-3 text-sm font-semibold text-white shadow-sm hover:bg-indigo-700 disabled:bg-slate-300"
            >
              {isSaving ? 'Salvando...' : 'Salvar campanha'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default AutomacaoCampanhaModal;
