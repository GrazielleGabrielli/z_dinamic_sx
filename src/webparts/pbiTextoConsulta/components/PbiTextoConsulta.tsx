import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import styles from './PbiTextoConsulta.module.scss';
import type { IPbiTextoConsultaProps } from './IPbiTextoConsultaProps';
import {
  createValidationTemplateItem,
  DEFAULT_VALIDATION_TITLE,
  DEFAULT_VALIDATION_STATUS,
  getValidationTemplateItemById
} from './validationService';
import type {
  ValidationPhaseOneErrors,
  ValidationPhaseOneFormData,
  ValidationTemplateItem
} from './validationTypes';

const initialFormData: ValidationPhaseOneFormData = {
  textoConsulta: ''
};

const POLLING_INTERVAL_MS = 5000;
const INITIAL_POLLING_TIMEOUT_MS = 120000;
const TOTAL_POLLING_TIMEOUT_MS = 300000;
const POWER_AUTOMATE_FLOW_URL = 'https://make.powerautomate.com/environments/Default-57b0d405-5803-4a25-b41a-65f25019513e/flows/2b45d0a8-4f9f-40ca-9bf9-748cf987981e/details?v3=false';

const formatElapsedTime = (totalSeconds: number): string => {
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  const formattedMinutes = minutes < 10 ? `0${minutes}` : String(minutes);
  const formattedSeconds = seconds < 10 ? `0${seconds}` : String(seconds);
  return `${formattedMinutes}:${formattedSeconds}`;
};

const normalizeValidationResult = (title: string): 'OK' | 'ERRO' | null => {
  const normalizedTitle = title.trim().toUpperCase();

  if (normalizedTitle === 'OK' || normalizedTitle === 'ERRO') {
    return normalizedTitle;
  }

  return null;
};

const PbiTextoConsulta = (_props: IPbiTextoConsultaProps): React.ReactElement => {
  const [formData, setFormData] = useState<ValidationPhaseOneFormData>(initialFormData);
  const [errors, setErrors] = useState<ValidationPhaseOneErrors>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [createdItem, setCreatedItem] = useState<ValidationTemplateItem | null>(null);
  const [activeItemId, setActiveItemId] = useState<number | null>(null);
  const [submitError, setSubmitError] = useState<string>('');
  const [isPolling, setIsPolling] = useState(false);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);
  const [pollingMessage, setPollingMessage] = useState<string>('');
  const [validationResult, setValidationResult] = useState<'OK' | 'ERRO' | null>(null);

  const textoConsultaLength = useMemo(() => formData.textoConsulta.length, [formData.textoConsulta]);
  const isValidationPending = createdItem?.Title.trim().toUpperCase() === DEFAULT_VALIDATION_TITLE;
  const shouldShowPowerAutomateLink =
    createdItem !== null &&
    validationResult === null &&
    elapsedSeconds >= INITIAL_POLLING_TIMEOUT_MS / 1000;

  const handleFieldChange =
    (field: keyof ValidationPhaseOneFormData) =>
    (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
      const { value } = event.target;

      setFormData((current) => ({
        ...current,
        [field]: value
      }));

      setErrors((current) => ({
        ...current,
        [field]: undefined
      }));
    };

  const validateForm = (): ValidationPhaseOneErrors => {
    const nextErrors: ValidationPhaseOneErrors = {};

    if (!formData.textoConsulta.trim()) {
      nextErrors.textoConsulta = 'Informe o texto da consulta para validar.';
    }

    return nextErrors;
  };

  useEffect(() => {
    if (!activeItemId || !isPolling) {
      return undefined;
    }

    const startedAt = Date.now();

    const timerInterval = window.setInterval(() => {
      const elapsedMs = Date.now() - startedAt;
      const seconds = Math.min(Math.floor(elapsedMs / 1000), TOTAL_POLLING_TIMEOUT_MS / 1000);
      setElapsedSeconds(seconds);

      if (elapsedMs >= INITIAL_POLLING_TIMEOUT_MS && elapsedMs < TOTAL_POLLING_TIMEOUT_MS) {
        setPollingMessage('Validacao ainda em processamento. Tentando novamente por mais 3 minutos. Se preferir, acompanhe no Power Automate.');
      }
    }, 1000);

    const pollItem = async (): Promise<void> => {
      try {
        const refreshedItem = await getValidationTemplateItemById(activeItemId);
        setCreatedItem(refreshedItem);

        const result = normalizeValidationResult(refreshedItem.Title);

        if (result) {
          setValidationResult(result);
          setPollingMessage(`Validacao finalizada com resultado ${result}.`);
          setIsPolling(false);
          return;
        }

        const elapsedMs = Date.now() - startedAt;

        if (elapsedMs >= TOTAL_POLLING_TIMEOUT_MS) {
          setElapsedSeconds(TOTAL_POLLING_TIMEOUT_MS / 1000);
          setPollingMessage('Tempo limite de 5 minutos atingido. Verifique o fluxo diretamente no Power Automate.');
          setIsPolling(false);
          return;
        }

        if (elapsedMs >= INITIAL_POLLING_TIMEOUT_MS) {
          setPollingMessage('Validacao ainda em processamento. Tentando novamente por mais 3 minutos. Se preferir, acompanhe no Power Automate.');
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Nao foi possivel consultar a validacao.';
        setSubmitError(message);
        setIsPolling(false);
      }
    };

    setPollingMessage('Aguardando retorno da validacao...');

    const pollingInterval = window.setInterval(() => {
      void pollItem();
    }, POLLING_INTERVAL_MS);

    void pollItem();

    return () => {
      window.clearInterval(timerInterval);
      window.clearInterval(pollingInterval);
    };
  }, [activeItemId, isPolling]);

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const nextErrors = validateForm();
    setErrors(nextErrors);
    setSubmitError('');
    setPollingMessage('');
    setValidationResult(null);
    setCreatedItem(null);
    setActiveItemId(null);
    setElapsedSeconds(0);

    if (Object.keys(nextErrors).length > 0) {
      return;
    }

    try {
      setIsSubmitting(true);

      const item = await createValidationTemplateItem({
        Title: DEFAULT_VALIDATION_TITLE,
        TextoConsulta: formData.textoConsulta.trim(),
        Status: DEFAULT_VALIDATION_STATUS
      });

      if (!item.Id) {
        throw new Error('Item criado sem ID valido.');
      }

      setCreatedItem(item);
      setActiveItemId(item.Id);
      setPollingMessage('Item criado com sucesso. Iniciando acompanhamento da validacao...');
      setIsPolling(true);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Nao foi possivel enviar a consulta para validacao.';
      setSubmitError(message);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <section className={styles.pbiTextoConsulta}>
      <div className="mx-auto flex w-full max-w-3xl flex-col gap-6">
        <header className="space-y-3">
          <div className="space-y-2">
            <h1 className="text-3xl font-semibold tracking-tight text-slate-900">
              Validar consulta
            </h1>
            <p className="text-sm leading-6 text-slate-600 sm:text-base">
              Informe o texto da consulta e inicie a validacao.
            </p>
          </div>
        </header>

        <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm sm:p-8">
          <form className="space-y-6" onSubmit={handleSubmit}>
            <div className="space-y-2">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <label className="text-sm font-medium text-slate-800" htmlFor="validation-texto-consulta">
                  Texto de consulta
                </label>
                <span className="text-xs font-medium text-slate-500">
                  {textoConsultaLength} caracteres
                </span>
              </div>
              <textarea
                id="validation-texto-consulta"
                value={formData.textoConsulta}
                onChange={handleFieldChange('textoConsulta')}
                placeholder="Cole aqui o texto da consulta."
                className={`min-h-[320px] w-full rounded-2xl border bg-white px-4 py-4 text-sm leading-6 text-slate-900 outline-none transition focus:ring-4 ${
                  errors.textoConsulta
                    ? 'border-red-300 focus:border-red-400 focus:ring-red-100'
                    : 'border-slate-200 focus:border-indigo-400 focus:ring-indigo-100'
                }`}
              />
              {errors.textoConsulta && <p className="text-sm text-red-600">{errors.textoConsulta}</p>}
            </div>

            {submitError && (
              <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {submitError}
              </div>
            )}

            <div className="flex justify-end">
              <button
                type="submit"
                disabled={isSubmitting || isPolling || isValidationPending}
                className="inline-flex items-center justify-center rounded-xl bg-indigo-600 px-5 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
              >
                {isSubmitting ? 'Enviando...' : isPolling ? 'Aguardando retorno...' : 'Validar'}
              </button>
            </div>
          </form>
        </div>

        {createdItem && (
          <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm sm:p-8">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div className="space-y-1">
                <p className="text-sm font-semibold text-slate-900">Acompanhamento da validacao</p>
                <p className="text-sm text-slate-600">{pollingMessage || 'Validacao iniciada.'}</p>
              </div>
              <div className="rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-right">
                <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">Tempo</p>
                <p className="mt-1 text-2xl font-semibold text-slate-900">{formatElapsedTime(elapsedSeconds)}</p>
              </div>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">ID</p>
                <p className="mt-2 text-lg font-semibold text-slate-900">{createdItem.Id}</p>
              </div>
              <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">Status</p>
                <p className="mt-2 text-lg font-semibold text-slate-900">{createdItem.Status || DEFAULT_VALIDATION_STATUS}</p>
              </div>
              <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">Resultado</p>
                <p
                  className={`mt-2 text-lg font-semibold ${
                    validationResult === 'OK'
                      ? 'text-emerald-600'
                      : validationResult === 'ERRO'
                        ? 'text-red-600'
                        : 'text-slate-900'
                  }`}
                >
                  {createdItem.RespostaPBI || 'Aguardando retorno do campo RespostaPBI'}
                </p>
              </div>
            </div>

            {shouldShowPowerAutomateLink && (
              <div className="mt-6 rounded-xl border border-amber-200 bg-amber-50 p-4">
                <p className="text-sm font-medium text-amber-900">
                  A validacao ultrapassou 2 minutos. O sistema continuara tentando ate completar 5 minutos.
                </p>
                <a
                  className="mt-2 inline-flex text-sm font-semibold text-amber-800 underline underline-offset-2"
                  href={POWER_AUTOMATE_FLOW_URL}
                  target="_blank"
                  rel="noreferrer"
                >
                  Verificar Power Automate
                </a>
              </div>
            )}
          </section>
        )}
      </div>
    </section>
  );
};

export default PbiTextoConsulta;
