import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { Icon, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import styles from './PbiTextoConsulta.module.scss';
import type { IPbiTextoConsultaProps } from './IPbiTextoConsultaProps';
import AutomacaoCampanhaModal from './AutomacaoCampanhaModal';
import type { AutomacaoCampanhaFormData } from './automacaoCampanhaTypes';
import { EMPTY_AUTOMACAO_CAMPANHA_FORM } from './automacaoCampanhaTypes';
import { buildAutomacaoCampanhaInitialForm } from './automacaoCampanhaUtils';
import {
  createValidationTemplateItem,
  DEFAULT_VALIDATION_TITLE,
  DEFAULT_VALIDATION_STATUS,
  deleteValidationTemplateItem,
  getValidationTemplateItemById,
  getValidationTemplateItemsTitleOk,
  updateValidarTemplatesItemStatus
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

type ParsedRespostaPBI =
  | { kind: 'array'; firstJson: string; fullJson: string; count: number }
  | { kind: 'text'; content: string }
  | { kind: 'empty' };

const isStatusFinalizado = (status: string): boolean => status.trim().toLowerCase() === 'finalizado';

const parseRespostaPBIContent = (raw: string): ParsedRespostaPBI => {
  const trimmed = raw.trim();

  if (!trimmed) {
    return { kind: 'empty' };
  }

  try {
    const data = JSON.parse(trimmed) as unknown;

    if (Array.isArray(data)) {
      if (data.length === 0) {
        return { kind: 'text', content: '[]' };
      }

      return {
        kind: 'array',
        firstJson: JSON.stringify(data[0], null, 2),
        fullJson: JSON.stringify(data, null, 2),
        count: data.length
      };
    }

    if (data !== null && typeof data === 'object') {
      return { kind: 'text', content: JSON.stringify(data, null, 2) };
    }

    return { kind: 'text', content: String(data) };
  } catch {
    return { kind: 'text', content: trimmed };
  }
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
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [refreshError, setRefreshError] = useState<string>('');
  const [respostaPanelOpen, setRespostaPanelOpen] = useState(false);
  const [automacaoModalOpen, setAutomacaoModalOpen] = useState(false);
  const [automacaoModalValues, setAutomacaoModalValues] = useState<AutomacaoCampanhaFormData>(EMPTY_AUTOMACAO_CAMPANHA_FORM);
  const [isApproving, setIsApproving] = useState(false);
  const [approveError, setApproveError] = useState<string>('');
  const [pageView, setPageView] = useState<'lista' | 'novo'>('lista');
  const [okItems, setOkItems] = useState<ValidationTemplateItem[]>([]);
  const [listaLoading, setListaLoading] = useState(false);
  const [listaError, setListaError] = useState<string>('');
  const [listaRespostaItem, setListaRespostaItem] = useState<ValidationTemplateItem | null>(null);
  const [listaRespostaFullOpen, setListaRespostaFullOpen] = useState(false);
  const [listaStatusTab, setListaStatusTab] = useState<'pendente' | 'finalizado'>('pendente');
  const [listaDeletingId, setListaDeletingId] = useState<number | null>(null);
  const [listaDeleteModalItem, setListaDeleteModalItem] = useState<ValidationTemplateItem | null>(null);

  const textoConsultaLength = useMemo(() => formData.textoConsulta.length, [formData.textoConsulta]);
  const respostaParsed = useMemo(
    () => parseRespostaPBIContent(createdItem?.RespostaPBI ?? ''),
    [createdItem?.RespostaPBI]
  );
  const listaRespostaParsed = useMemo(
    () => parseRespostaPBIContent(listaRespostaItem?.RespostaPBI ?? ''),
    [listaRespostaItem?.RespostaPBI]
  );
  const okItemsPendente = useMemo(
    () => okItems.filter((item) => !isStatusFinalizado(item.Status)),
    [okItems]
  );
  const okItemsFinalizado = useMemo(
    () => okItems.filter((item) => isStatusFinalizado(item.Status)),
    [okItems]
  );
  const okItemsNaAba = listaStatusTab === 'finalizado' ? okItemsFinalizado : okItemsPendente;
  const isValidationPending = createdItem?.Title.trim().toUpperCase() === DEFAULT_VALIDATION_TITLE;
  const shouldShowPowerAutomateLink =
    createdItem !== null &&
    validationResult === null &&
    elapsedSeconds >= INITIAL_POLLING_TIMEOUT_MS / 1000;

  const showAprovarButton =
    validationResult === 'OK' &&
    createdItem !== null &&
    createdItem.Status.trim().toLowerCase() !== 'finalizado';

  const loadOkItems = useCallback(async (): Promise<void> => {
    setListaError('');
    setListaLoading(true);
    try {
      const items = await getValidationTemplateItemsTitleOk();
      setOkItems(items);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Falha ao carregar a lista.';
      setListaError(message);
    } finally {
      setListaLoading(false);
    }
  }, []);

  useEffect(() => {
    if (pageView === 'lista') {
      void loadOkItems();
    }
  }, [pageView, loadOkItems]);

  const truncateListaPreview = (text: string, maxLen: number): string => {
    const t = text.trim();
    if (t.length <= maxLen) {
      return t;
    }
    return `${t.slice(0, maxLen)}...`;
  };

  const handleCriarCampanhaDesdeLista = (item: ValidationTemplateItem): void => {
    setAutomacaoModalValues(buildAutomacaoCampanhaInitialForm(item.TextoConsulta, item.RespostaPBI));
    setAutomacaoModalOpen(true);
  };

  const closeListaRespostaModal = (): void => {
    setListaRespostaFullOpen(false);
    setListaRespostaItem(null);
  };

  const handleExcluirItemValidarTemplates = async (): Promise<void> => {
    if (!listaDeleteModalItem || isStatusFinalizado(listaDeleteModalItem.Status)) {
      return;
    }
    setListaError('');
    setListaDeletingId(listaDeleteModalItem.Id);
    try {
      await deleteValidationTemplateItem(listaDeleteModalItem.Id);
      if (listaRespostaItem?.Id === listaDeleteModalItem.Id) {
        closeListaRespostaModal();
      }
      setListaDeleteModalItem(null);
      await loadOkItems();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Falha ao excluir o item.';
      setListaError(message);
    } finally {
      setListaDeletingId(null);
    }
  };

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

  useEffect(() => {
    setRespostaPanelOpen(false);
  }, [createdItem?.Id]);

  useEffect(() => {
    if (respostaParsed.kind !== 'array') {
      setRespostaPanelOpen(false);
    }
  }, [respostaParsed.kind]);

  useEffect(() => {
    if (!respostaPanelOpen) {
      return undefined;
    }

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';

    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setRespostaPanelOpen(false);
      }
    };

    window.addEventListener('keydown', onKeyDown);

    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [respostaPanelOpen]);

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const nextErrors = validateForm();
    setErrors(nextErrors);
    setSubmitError('');
    setRefreshError('');
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

  const handleAprovar = async (): Promise<void> => {
    const itemId = createdItem?.Id ?? activeItemId;
    if (!itemId || validationResult !== 'OK') {
      return;
    }

    setApproveError('');

    try {
      setIsApproving(true);
      await updateValidarTemplatesItemStatus(itemId, 'Finalizado');
      const refreshedItem = await getValidationTemplateItemById(itemId);
      setCreatedItem(refreshedItem);

      const initialAutomacao = buildAutomacaoCampanhaInitialForm(
        refreshedItem.TextoConsulta,
        refreshedItem.RespostaPBI
      );
      setAutomacaoModalValues(initialAutomacao);
      setAutomacaoModalOpen(true);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Nao foi possivel aprovar.';
      setApproveError(message);
    } finally {
      setIsApproving(false);
    }
  };

  const handleRefreshItem = async (): Promise<void> => {
    const itemId = createdItem?.Id ?? activeItemId;
    if (!itemId) {
      return;
    }

    setRefreshError('');

    try {
      setIsRefreshing(true);
      const refreshedItem = await getValidationTemplateItemById(itemId);
      setCreatedItem(refreshedItem);

      const result = normalizeValidationResult(refreshedItem.Title);

      if (result) {
        setValidationResult(result);
        setPollingMessage(`Validacao finalizada com resultado ${result}.`);
        setIsPolling(false);
      } else {
        setPollingMessage('Dados atualizados. Aguardando retorno da validacao...');
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Nao foi possivel atualizar o item.';
      setRefreshError(message);
    } finally {
      setIsRefreshing(false);
    }
  };

  return (
    <section className={styles.pbiTextoConsulta}>
      <div className="mx-auto flex w-full max-w-5xl flex-col gap-8 px-4 sm:px-6 lg:px-8">
        {pageView === 'lista' ? (
          <>
            <header className="overflow-hidden rounded-[28px] border border-white/70 bg-white/90 px-6 py-7 shadow-[0_18px_45px_rgba(15,23,42,0.08)] backdrop-blur sm:px-8 sm:py-9">
              <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
                <div className="space-y-3">
                  <span className="inline-flex w-fit rounded-full border border-emerald-200 bg-emerald-50 px-3 py-1 text-xs font-semibold uppercase tracking-[0.14em] text-emerald-700">
                    Title OK
                  </span>
                  <h1 className="text-3xl font-semibold tracking-tight text-slate-900 sm:text-4xl">
                    Consultas validadas
                  </h1>
                  <p className="max-w-2xl text-sm leading-7 text-slate-600 sm:text-base">
                    Itens da lista ValidarTemplates com titulo OK. Use Criar campanha para abrir o cadastro de automacao ou Novo para enviar outra consulta.
                  </p>
                </div>
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
                  <button
                    type="button"
                    onClick={() => void loadOkItems()}
                    disabled={listaLoading}
                    className="inline-flex min-h-[48px] items-center justify-center rounded-2xl border border-slate-300 bg-white px-8 py-3.5 text-sm font-semibold text-slate-800 shadow-sm transition hover:border-slate-400 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    {listaLoading ? 'Atualizando...' : 'Atualizar lista'}
                  </button>
                  <button
                    type="button"
                    onClick={() => setPageView('novo')}
                    className="inline-flex min-h-[48px] items-center justify-center rounded-2xl bg-indigo-600 px-8 py-3.5 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-700"
                  >
                    Novo
                  </button>
                </div>
              </div>
            </header>

            {listaError && (
              <div className="rounded-2xl border border-red-200 bg-red-50 px-4 py-4 text-sm text-red-700">
                {listaError}
              </div>
            )}

            <div className="rounded-[28px] border border-slate-200/80 bg-white/95 p-6 shadow-[0_22px_55px_rgba(15,23,42,0.08)] backdrop-blur sm:p-8">
              <div className="mb-6 flex flex-wrap gap-2 border-b border-slate-200 pb-4" role="tablist" aria-label="Filtro por status">
                <button
                  type="button"
                  role="tab"
                  aria-selected={listaStatusTab === 'pendente'}
                  onClick={() => setListaStatusTab('pendente')}
                  className={`inline-flex min-h-[44px] items-center justify-center rounded-xl px-5 py-2.5 text-sm font-semibold transition ${
                    listaStatusTab === 'pendente'
                      ? 'bg-indigo-600 text-white shadow-sm'
                      : 'border border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  Pendente
                  <span
                    className={`ml-2 rounded-full px-2 py-0.5 text-xs font-bold ${
                      listaStatusTab === 'pendente' ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {okItemsPendente.length}
                  </span>
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={listaStatusTab === 'finalizado'}
                  onClick={() => setListaStatusTab('finalizado')}
                  className={`inline-flex min-h-[44px] items-center justify-center rounded-xl px-5 py-2.5 text-sm font-semibold transition ${
                    listaStatusTab === 'finalizado'
                      ? 'bg-indigo-600 text-white shadow-sm'
                      : 'border border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  Finalizado
                  <span
                    className={`ml-2 rounded-full px-2 py-0.5 text-xs font-bold ${
                      listaStatusTab === 'finalizado' ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {okItemsFinalizado.length}
                  </span>
                </button>
              </div>
              {listaLoading && okItems.length === 0 ? (
                <p className="text-center text-sm text-slate-600">Carregando itens...</p>
              ) : null}
              {!listaLoading && okItems.length === 0 && !listaError ? (
                <p className="text-center text-sm text-slate-600">Nenhum item com Title OK encontrado.</p>
              ) : null}
              {!listaLoading && okItems.length > 0 && okItemsNaAba.length === 0 && !listaError ? (
                <p className="text-center text-sm text-slate-600">
                  {listaStatusTab === 'finalizado'
                    ? 'Nenhum item com Status Finalizado.'
                    : 'Nenhum item pendente (fora de Finalizado).'}
                </p>
              ) : null}
              <ul className="space-y-4">
                {okItemsNaAba.map((item) => (
                  <li
                    key={item.Id}
                    className="flex flex-col gap-4 rounded-2xl border border-slate-200 bg-slate-50/80 p-5 sm:flex-row sm:items-center sm:justify-between"
                  >
                    <div className="min-w-0 flex-1 space-y-2">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-500">ID {item.Id}</span>
                        <span className="rounded-full bg-white px-2.5 py-0.5 text-xs font-medium text-slate-700">
                          {item.Status || DEFAULT_VALIDATION_STATUS}
                        </span>
                      </div>
                      <p className="text-sm leading-6 text-slate-700">{truncateListaPreview(item.TextoConsulta, 200)}</p>
                    </div>
                    <div
                      className="flex shrink-0 self-start divide-x divide-slate-200 overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm sm:self-center"
                      role="group"
                      aria-label={`Acoes do item ${item.Id}`}
                    >
                      <TooltipHost
                        content="Ver RespostaPBI (itens retornados pela consulta)"
                        calloutProps={{ gapSpace: 4 }}
                      >
                        <button
                          type="button"
                          aria-label="Ver itens"
                          onClick={() => {
                            setListaRespostaFullOpen(false);
                            setListaRespostaItem(item);
                          }}
                          className="flex h-9 w-9 items-center justify-center text-slate-600 transition hover:bg-slate-50"
                        >
                          <Icon iconName="List" styles={{ root: { fontSize: 16 } }} />
                        </button>
                      </TooltipHost>
                      <TooltipHost
                        content="Criar campanha na lista AutomacaoCampanhas"
                        calloutProps={{ gapSpace: 4 }}
                      >
                        <button
                          type="button"
                          aria-label="Criar campanha"
                          onClick={() => handleCriarCampanhaDesdeLista(item)}
                          className="flex h-9 w-9 items-center justify-center text-indigo-600 transition hover:bg-indigo-50"
                        >
                          <Icon iconName="Add" styles={{ root: { fontSize: 16 } }} />
                        </button>
                      </TooltipHost>
                      {!isStatusFinalizado(item.Status) ? (
                        <TooltipHost
                          content="Excluir este item pendente da lista ValidarTemplates"
                          calloutProps={{ gapSpace: 4 }}
                        >
                          <span className="inline-flex">
                            <button
                              type="button"
                              aria-label="Excluir"
                              onClick={() => setListaDeleteModalItem(item)}
                              disabled={listaDeletingId === item.Id}
                              className="flex h-9 w-9 items-center justify-center text-red-600 transition hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-50"
                            >
                              {listaDeletingId === item.Id ? (
                                <Spinner size={SpinnerSize.xSmall} />
                              ) : (
                                <Icon iconName="Delete" styles={{ root: { fontSize: 16 } }} />
                              )}
                            </button>
                          </span>
                        </TooltipHost>
                      ) : null}
                    </div>
                  </li>
                ))}
              </ul>
            </div>

            {listaRespostaItem && (
              <div
                className="fixed inset-0 z-[1000] flex items-start justify-center overflow-y-auto bg-slate-900/50 p-4 py-10 sm:p-6"
                role="dialog"
                aria-modal="true"
                aria-labelledby="lista-resposta-title"
              >
                <button
                  type="button"
                  className="fixed inset-0 cursor-default bg-transparent"
                  onClick={closeListaRespostaModal}
                  aria-label="Fechar"
                />
                <div className="relative z-[1] w-full max-w-3xl rounded-[28px] border border-slate-200 bg-white shadow-2xl">
                  <div className="flex items-start justify-between gap-4 border-b border-slate-200 px-5 py-4 sm:px-6">
                    <div>
                      <p id="lista-resposta-title" className="text-base font-semibold text-slate-900">
                        RespostaPBI
                      </p>
                      <p className="mt-0.5 text-sm text-slate-500">ID {listaRespostaItem.Id}</p>
                    </div>
                    <button
                      type="button"
                      onClick={closeListaRespostaModal}
                      className="inline-flex min-h-[44px] min-w-[44px] shrink-0 items-center justify-center rounded-xl border border-slate-200 bg-white text-lg font-semibold text-slate-600 transition hover:bg-slate-50"
                      aria-label="Fechar"
                    >
                      ×
                    </button>
                  </div>
                  <div className="max-h-[min(70vh,640px)] overflow-auto p-5 sm:p-6">
                    {listaRespostaParsed.kind === 'empty' && (
                      <p className="text-sm font-medium text-slate-600">Sem conteudo em RespostaPBI.</p>
                    )}
                    {listaRespostaParsed.kind === 'text' && (
                      <pre className="whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-slate-50 p-4 font-mono text-xs leading-relaxed text-slate-800">
                        {listaRespostaParsed.content}
                      </pre>
                    )}
                    {listaRespostaParsed.kind === 'array' && (
                      <div className="space-y-3">
                        <p className="text-xs font-medium text-slate-500">
                          Primeiro registro de {listaRespostaParsed.count}{' '}
                          {listaRespostaParsed.count === 1 ? 'item' : 'itens'}
                        </p>
                        <pre className="max-h-64 overflow-auto whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-slate-50 p-4 font-mono text-xs leading-relaxed text-slate-800">
                          {listaRespostaParsed.firstJson}
                        </pre>
                        {listaRespostaParsed.count > 1 && (
                          <button
                            type="button"
                            onClick={() => setListaRespostaFullOpen(true)}
                            className="inline-flex min-h-[44px] items-center justify-center rounded-xl border border-indigo-200 bg-indigo-50 px-6 py-2.5 text-sm font-semibold text-indigo-800 transition hover:bg-indigo-100"
                          >
                            Mostrar tudo
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {listaRespostaFullOpen && listaRespostaParsed.kind === 'array' && listaRespostaItem && (
              <div className="fixed inset-0 z-[1001] flex justify-end" role="dialog" aria-modal="true" aria-labelledby="lista-resposta-full-title">
                <button
                  type="button"
                  className="absolute inset-0 bg-slate-900/50 transition-opacity"
                  onClick={() => setListaRespostaFullOpen(false)}
                  aria-label="Fechar painel"
                />
                <aside className="relative flex h-full w-full max-w-2xl flex-col border-l border-slate-200 bg-white shadow-2xl">
                  <div className="flex items-center justify-between gap-4 border-b border-slate-200 px-5 py-4 sm:px-6">
                    <div>
                      <p id="lista-resposta-full-title" className="text-base font-semibold text-slate-900">
                        Resultado completo
                      </p>
                      <p className="mt-0.5 text-sm text-slate-500">{listaRespostaParsed.count} itens</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => setListaRespostaFullOpen(false)}
                      className="inline-flex min-h-[44px] min-w-[44px] items-center justify-center rounded-xl border border-slate-200 bg-white text-lg font-semibold text-slate-600 transition hover:bg-slate-50"
                      aria-label="Fechar"
                    >
                      ×
                    </button>
                  </div>
                  <div className="min-h-0 flex-1 overflow-auto p-5 sm:p-6">
                    <pre className="whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-slate-50 p-4 font-mono text-xs leading-relaxed text-slate-800">
                      {listaRespostaParsed.fullJson}
                    </pre>
                  </div>
                </aside>
              </div>
            )}

            {listaDeleteModalItem && (
              <div
                className="fixed inset-0 z-[1002] flex items-center justify-center bg-slate-900/50 p-4"
                role="dialog"
                aria-modal="true"
                aria-labelledby="lista-delete-title"
              >
                <button
                  type="button"
                  className="fixed inset-0 cursor-default bg-transparent"
                  onClick={() => setListaDeleteModalItem(null)}
                  aria-label="Fechar"
                />
                <div className="relative z-[1] w-full max-w-md rounded-2xl border border-slate-200 bg-white p-6 shadow-2xl">
                  <h3 id="lista-delete-title" className="text-lg font-semibold text-slate-900">
                    Excluir item
                  </h3>
                  <p className="mt-2 text-sm leading-6 text-slate-600">
                    Confirma a exclusao do item ID {listaDeleteModalItem.Id} da lista ValidarTemplates?
                  </p>
                  <div className="mt-6 flex items-center justify-end gap-3">
                    <button
                      type="button"
                      onClick={() => setListaDeleteModalItem(null)}
                      className="inline-flex min-h-[40px] items-center justify-center rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-700 transition hover:bg-slate-50"
                    >
                      Cancelar
                    </button>
                    <button
                      type="button"
                      onClick={() => void handleExcluirItemValidarTemplates()}
                      disabled={listaDeletingId === listaDeleteModalItem.Id}
                      className="inline-flex min-h-[40px] items-center justify-center rounded-xl bg-red-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-red-700 disabled:cursor-not-allowed disabled:bg-slate-300"
                    >
                      {listaDeletingId === listaDeleteModalItem.Id ? 'Excluindo...' : 'Excluir'}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </>
        ) : (
          <>
            <header className="overflow-hidden rounded-[28px] border border-white/70 bg-white/90 px-6 py-7 shadow-[0_18px_45px_rgba(15,23,42,0.08)] backdrop-blur sm:px-8 sm:py-9">
              <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
                <div className="space-y-3">
                  <div className="flex flex-wrap items-center gap-3">
                    <button
                      type="button"
                      onClick={() => setPageView('lista')}
                      className="inline-flex min-h-[40px] items-center justify-center rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-semibold text-slate-800 hover:bg-slate-50"
                    >
                      Voltar
                    </button>
                    <span className="inline-flex w-fit rounded-full border border-indigo-200 bg-indigo-50 px-3 py-1 text-xs font-semibold uppercase tracking-[0.14em] text-indigo-700">
                      Nova validacao
                    </span>
                  </div>
                  <h1 className="text-3xl font-semibold tracking-tight text-slate-900 sm:text-4xl">
                    Validar consulta
                  </h1>
                  <p className="max-w-2xl text-sm leading-7 text-slate-600 sm:text-base">
                    Envie o texto da consulta para a lista ValidarTemplates e acompanhe o retorno automatico da validacao.
                  </p>
                </div>
                <div className="grid grid-cols-2 gap-3 sm:w-fit">
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
                    <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Status inicial</p>
                    <p className="mt-1 text-sm font-semibold text-slate-900">{DEFAULT_VALIDATION_STATUS}</p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3">
                    <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Acao</p>
                    <p className="mt-1 text-sm font-semibold text-slate-900">{DEFAULT_VALIDATION_TITLE}</p>
                  </div>
                </div>
              </div>
            </header>

            <div className="rounded-[28px] border border-slate-200/80 bg-white/95 p-6 shadow-[0_22px_55px_rgba(15,23,42,0.08)] backdrop-blur sm:p-8 lg:p-10">
          <div className="mb-8 flex flex-col gap-3 border-b border-slate-100 pb-6">
            <p className="text-sm font-semibold text-slate-900">Texto de consulta</p>
            <p className="text-sm leading-6 text-slate-500">
              Cole abaixo a consulta que sera enviada para validacao automatica.
            </p>
          </div>

          <form className="space-y-6" onSubmit={handleSubmit}>
            <div className="space-y-3">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <label className="text-sm font-medium text-slate-800" htmlFor="validation-texto-consulta">
                  Texto de consulta
                </label>
                <span className="inline-flex w-fit rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600">
                  {textoConsultaLength} caracteres
                </span>
              </div>
              <textarea
                id="validation-texto-consulta"
                rows={7}
                value={formData.textoConsulta}
                onChange={handleFieldChange('textoConsulta')}
                placeholder="Cole aqui o texto da consulta."
                className={`w-full rounded-[24px] border bg-slate-50/70 px-5 py-4 text-sm leading-7 text-slate-900 outline-none transition duration-200 placeholder:text-slate-400 focus:bg-white focus:ring-4 ${
                  errors.textoConsulta
                    ? 'border-red-300 focus:border-red-400 focus:ring-red-100'
                    : 'border-slate-200 focus:border-indigo-400 focus:ring-indigo-100'
                }`}
              />
              {errors.textoConsulta && <p className="text-sm font-medium text-red-600">{errors.textoConsulta}</p>}
            </div>

            {submitError && (
              <div className="rounded-2xl border border-red-200 bg-red-50 px-4 py-4 text-sm text-red-700">
                {submitError}
              </div>
            )}

            <div className="flex flex-col gap-5 border-t border-slate-100 pt-8 sm:flex-row sm:items-center sm:justify-between sm:gap-6">
              <p className="text-sm leading-6 text-slate-500 sm:pr-4">
                O envio fica bloqueado enquanto houver uma validacao pendente.
              </p>
              <button
                type="submit"
                disabled={isSubmitting || isPolling || isValidationPending}
                className="inline-flex min-h-[52px] min-w-[200px] shrink-0 items-center justify-center rounded-2xl bg-indigo-600 px-10 py-4 text-sm font-semibold text-white shadow-[0_12px_30px_rgba(79,70,229,0.28)] transition duration-200 hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none sm:px-12 sm:py-[1.125rem]"
              >
                {isSubmitting ? 'Enviando...' : isPolling ? 'Aguardando retorno...' : 'Validar'}
              </button>
            </div>
          </form>
        </div>

        {createdItem && (
          <section className="rounded-[28px] border border-slate-200/80 bg-white/95 p-6 shadow-[0_22px_55px_rgba(15,23,42,0.08)] backdrop-blur sm:p-8 lg:p-10">
            <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
              <div className="space-y-2">
                <p className="text-sm font-semibold text-slate-900">Acompanhamento da validacao</p>
                <p className="max-w-2xl text-sm leading-6 text-slate-600">{pollingMessage || 'Validacao iniciada.'}</p>
              </div>
              <div className="flex flex-col gap-3 sm:flex-row sm:items-stretch lg:flex-col xl:flex-row xl:items-start">
                <button
                  type="button"
                  onClick={() => void handleRefreshItem()}
                  disabled={isRefreshing}
                  className="inline-flex min-h-[48px] min-w-[160px] shrink-0 items-center justify-center rounded-2xl border border-slate-300 bg-white px-8 py-3.5 text-sm font-semibold text-slate-800 shadow-sm transition hover:border-slate-400 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-60"
                >
                  {isRefreshing ? 'Atualizando...' : 'Atualizar'}
                </button>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 px-5 py-4 text-left sm:min-w-[180px] lg:text-right">
                  <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Tempo</p>
                  <p className="mt-1 text-3xl font-semibold tracking-tight text-slate-900">{formatElapsedTime(elapsedSeconds)}</p>
                </div>
              </div>
            </div>

            {refreshError && (
              <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {refreshError}
              </div>
            )}

            {approveError && (
              <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                {approveError}
              </div>
            )}

            {showAprovarButton && (
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => void handleAprovar()}
                  disabled={isApproving}
                  className="inline-flex min-h-[52px] min-w-[200px] items-center justify-center rounded-2xl bg-emerald-600 px-10 py-4 text-sm font-semibold text-white shadow-[0_12px_30px_rgba(5,150,105,0.28)] transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
                >
                  {isApproving ? 'Aprovando...' : 'Aprovar'}
                </button>
              </div>
            )}

            <div className="mt-8 grid gap-4 md:grid-cols-3">
              <div className="rounded-2xl border border-slate-200 bg-slate-50 p-5">
                <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">ID</p>
                <p className="mt-2 text-2xl font-semibold tracking-tight text-slate-900">{createdItem.Id}</p>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-slate-50 p-5">
                <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Status</p>
                <p className="mt-2 text-lg font-semibold text-slate-900">{createdItem.Status || DEFAULT_VALIDATION_STATUS}</p>
              </div>
              <div
                className={`rounded-2xl border bg-slate-50 p-5 md:col-span-3 ${
                  validationResult === 'OK'
                    ? 'border-emerald-200'
                    : validationResult === 'ERRO'
                      ? 'border-red-200'
                      : 'border-slate-200'
                }`}
              >
                <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                  <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-500">Resultado</p>
                  {validationResult && (
                    <span
                      className={`inline-flex w-fit rounded-full px-3 py-1 text-xs font-semibold ${
                        validationResult === 'OK'
                          ? 'bg-emerald-100 text-emerald-800'
                          : 'bg-red-100 text-red-800'
                      }`}
                    >
                      {validationResult}
                    </span>
                  )}
                </div>
                {respostaParsed.kind === 'empty' && (
                  <p className="mt-3 text-sm font-medium text-slate-600">Aguardando retorno do campo RespostaPBI</p>
                )}
                {respostaParsed.kind === 'text' && (
                  <pre className="mt-3 max-h-64 overflow-auto whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-white p-4 font-mono text-xs leading-relaxed text-slate-800">
                    {respostaParsed.content}
                  </pre>
                )}
                {respostaParsed.kind === 'array' && (
                  <div className="mt-3 space-y-3">
                    <p className="text-xs font-medium text-slate-500">
                      Primeiro registro de {respostaParsed.count} {respostaParsed.count === 1 ? 'item' : 'itens'}
                    </p>
                    <pre className="max-h-64 overflow-auto whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-white p-4 font-mono text-xs leading-relaxed text-slate-800">
                      {respostaParsed.firstJson}
                    </pre>
                    {respostaParsed.count > 1 && (
                      <button
                        type="button"
                        onClick={() => setRespostaPanelOpen(true)}
                        className="inline-flex min-h-[44px] items-center justify-center rounded-xl border border-indigo-200 bg-indigo-50 px-6 py-2.5 text-sm font-semibold text-indigo-800 transition hover:bg-indigo-100"
                      >
                        Mostrar tudo
                      </button>
                    )}
                  </div>
                )}
              </div>
            </div>

            {respostaPanelOpen && respostaParsed.kind === 'array' && (
              <div className="fixed inset-0 z-[1000] flex justify-end" role="dialog" aria-modal="true" aria-labelledby="resposta-panel-title">
                <button
                  type="button"
                  className="absolute inset-0 bg-slate-900/50 transition-opacity"
                  onClick={() => setRespostaPanelOpen(false)}
                  aria-label="Fechar painel"
                />
                <aside className="relative flex h-full w-full max-w-2xl flex-col border-l border-slate-200 bg-white shadow-2xl">
                  <div className="flex items-center justify-between gap-4 border-b border-slate-200 px-5 py-4 sm:px-6">
                    <div>
                      <p id="resposta-panel-title" className="text-base font-semibold text-slate-900">
                        Resultado completo
                      </p>
                      <p className="mt-0.5 text-sm text-slate-500">{respostaParsed.count} itens</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => setRespostaPanelOpen(false)}
                      className="inline-flex min-h-[44px] min-w-[44px] items-center justify-center rounded-xl border border-slate-200 bg-white text-lg font-semibold text-slate-600 transition hover:bg-slate-50"
                      aria-label="Fechar"
                    >
                      ×
                    </button>
                  </div>
                  <div className="min-h-0 flex-1 overflow-auto p-5 sm:p-6">
                    <pre className="whitespace-pre-wrap break-words rounded-xl border border-slate-200 bg-slate-50 p-4 font-mono text-xs leading-relaxed text-slate-800">
                      {respostaParsed.fullJson}
                    </pre>
                  </div>
                </aside>
              </div>
            )}

            {shouldShowPowerAutomateLink && (
              <div className="mt-6 rounded-2xl border border-amber-200 bg-amber-50 px-5 py-4">
                <p className="text-sm font-medium leading-6 text-amber-900">
                  A validacao ultrapassou 2 minutos. O sistema continuara tentando ate completar 5 minutos.
                </p>
                <a
                  className="mt-4 inline-flex min-h-[48px] items-center justify-center rounded-xl border border-amber-300/80 bg-white px-6 py-3 text-sm font-semibold text-amber-900 shadow-sm transition hover:bg-amber-100/80"
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

          </>
        )}

        <AutomacaoCampanhaModal
          isOpen={automacaoModalOpen}
          initialValues={automacaoModalValues}
          onClose={() => setAutomacaoModalOpen(false)}
          onSaved={async () => {
            await loadOkItems();
          }}
        />
      </div>
    </section>
  );
};

export default PbiTextoConsulta;
