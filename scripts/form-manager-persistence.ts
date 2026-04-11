import type { IDynamicViewConfig } from '../src/webparts/dinamicApp/core/config/types';
import { parseConfig } from '../src/webparts/dinamicApp/core/config/validators';
import { getDefaultConfig, getDefaultFormManagerConfig } from '../src/webparts/dinamicApp/core/config/utils';

function fail(msg: string): never {
  throw new Error(`[form-manager-persistence] ${msg}`);
}

function sampleFormManagerView(): IDynamicViewConfig {
  const d = getDefaultConfig();
  return {
    ...d,
    dataSource: { kind: 'list', title: 'ListaTeste' },
    mode: 'formManager',
    formManager: {
      ...getDefaultFormManagerConfig(),
      attachmentStorageKind: 'itemAttachments',
      stepLayout: 'rail',
      customButtons: [
        {
          id: 'b1',
          label: 'T',
          behavior: 'actionsOnly',
          actions: [],
          finishAfterRun: { kind: 'clearForm' },
        },
      ],
    },
  };
}

function main(): void {
  const ok = sampleFormManagerView();
  const round = parseConfig(JSON.stringify(ok));
  if (!round?.formManager) fail('round-trip: formManager ausente');
  if (round.formManager.attachmentStorageKind !== 'itemAttachments') {
    fail('round-trip: attachmentStorageKind itemAttachments');
  }
  if (round.formManager.customButtons?.[0]?.finishAfterRun?.kind !== 'clearForm') {
    fail('round-trip: finishAfterRun clearForm');
  }
  if (round.formManager.stepLayout !== 'rail') fail('round-trip: stepLayout rail');

  const again = parseConfig(JSON.stringify(round));
  if (JSON.stringify(round.formManager) !== JSON.stringify(again?.formManager)) {
    fail('idempotência: segundo parse alterou formManager');
  }

  const broken: IDynamicViewConfig = {
    ...ok,
    dashboard: {
      enabled: true,
      dashboardType: 'cards',
      cardsCount: 1,
      cards: [{ invalid: true } as unknown as (typeof ok.dashboard.cards)[0]],
      chartType: 'bar',
    },
  };
  const repaired = parseConfig(JSON.stringify(broken));
  if (!repaired?.formManager?.customButtons?.length) {
    fail('reparo formManager: customButtons perdidos com dashboard inválido');
  }
  if (!repaired.dashboard?.cards || !Array.isArray(repaired.dashboard.cards)) {
    fail('reparo: dashboard sem cards válidos');
  }

  // eslint-disable-next-line no-console
  console.log('form-manager-persistence: ok');
}

main();
