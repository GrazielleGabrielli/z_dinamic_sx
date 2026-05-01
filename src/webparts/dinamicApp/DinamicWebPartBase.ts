import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import '../../assets/dist/tailwind.css';
import DinamicApp from './components/DinamicApp';
import { IDinamicAppProps } from './components/IDinamicAppProps';
import { IDynamicViewConfig, IDynamicViewWebPartProps, TViewMode } from './core/config/types';
import { parseConfig } from './core/config/validators';
import { getDefaultConfig } from './core/config/utils';
import { getSP, getGraph } from './pnpConfig';
import {
  registerNativeEditSaveToolbarBridge,
  runNativePagePersistAfterPropertyWrite,
} from './core/sharePoint/sharePointPageToolbarDom';
import { TPersistStatus } from './core/persist/types';

const LOG = '[DinamicSX Persist]';

// Tempo assumido como proxy de conclusão da persistência nativa do SharePoint.
// runNativePagePersistAfterPropertyWrite é fire-and-forget (retorna void);
// não há callback ou promise real de confirmação do SharePoint.
const PERSIST_ASSUMED_DURATION_MS = 2500;
const SAVED_DISPLAY_DURATION_MS = 3000;

export abstract class DinamicWebPartBase extends BaseClientSideWebPart<IDynamicViewWebPartProps> {
  private _persistStatus: TPersistStatus = 'idle';
  private _nativePersistTimer: number | undefined;
  private _savedTimer: number | undefined;
  private _idleTimer: number | undefined;
  private _beforeUnloadHandler: ((e: BeforeUnloadEvent) => void) | undefined;
  private _nativeEditSaveBridgeTeardown: (() => void) | undefined;

  protected abstract getForcedMode(): TViewMode | undefined;

  protected onInit(): Promise<void> {
    getSP(this.context);
    getGraph(this.context);

    this._nativeEditSaveBridgeTeardown = registerNativeEditSaveToolbarBridge(this.domElement);

    this._beforeUnloadHandler = (e: BeforeUnloadEvent): void => {
      if (this._persistStatus === 'saving' || this._persistStatus === 'persisting') {
        e.preventDefault();
        // Alguns browsers modernos ignoram returnValue mas ainda exigem a atribuição
        e.returnValue = 'Há um salvamento em andamento. Deseja realmente sair?';
      }
    };
    window.addEventListener('beforeunload', this._beforeUnloadHandler);

    return super.onInit();
  }

  public render(): void {
    const forcedMode = this.getForcedMode();
    const element: React.ReactElement<IDinamicAppProps> = React.createElement(DinamicApp, {
      configJson: this.properties.configJson ?? '',
      siteUrl: this.context.pageContext.web.serverRelativeUrl,
      instanceScopeId: this.instanceId,
      onSaveConfig: (config: IDynamicViewConfig) => this.saveConfig(config),
      persistStatus: this._persistStatus,
      ...(forcedMode !== undefined ? { forcedMode } : {}),
    });

    ReactDom.render(element, this.domElement);
  }

  private loadConfig(): IDynamicViewConfig {
    const raw = this.properties.configJson;
    console.log(`${LOG} load — raw length: ${raw?.length ?? 0}`);
    const result = parseConfig(raw) ?? getDefaultConfig();
    console.log(`${LOG} load — parse resultado:`, result.mode, result.dataSource?.title);
    return result;
  }

  private clearAllPersistTimers(): void {
    if (this._nativePersistTimer !== undefined) {
      window.clearTimeout(this._nativePersistTimer);
      this._nativePersistTimer = undefined;
    }
    if (this._savedTimer !== undefined) {
      window.clearTimeout(this._savedTimer);
      this._savedTimer = undefined;
    }
    if (this._idleTimer !== undefined) {
      window.clearTimeout(this._idleTimer);
      this._idleTimer = undefined;
    }
  }

  private setStatus(status: TPersistStatus): void {
    this._persistStatus = status;
    this.render();
  }

  private saveConfig(config: IDynamicViewConfig): void {
    if (this._persistStatus === 'saving' || this._persistStatus === 'persisting') {
      console.warn(`${LOG} save ignorado — persistência já em andamento (status: ${this._persistStatus})`);
      return;
    }

    this.clearAllPersistTimers();

    const serialized = JSON.stringify(config);
    console.log(`${LOG} save iniciado — displayMode: ${this.displayMode === DisplayMode.Edit ? 'Edit' : 'Read'} — ${serialized.length} chars`);
    console.log(`${LOG} JSON:`, serialized);

    this.properties.configJson = serialized;

    if (this.displayMode === DisplayMode.Edit) {
      // Em Edit Mode, this.properties é serializado nativamente pelo SharePoint quando o
      // usuário salvar a página. Não é necessário o DOM hack; o banner 'pending' orienta o
      // usuário a salvar. O SharePoint recarrega a página após o save, resetando o estado.
      console.log(`${LOG} Edit Mode — config escrita em this.properties; aguardando save manual da página`);
      this.setStatus('pending');
      return;
    }

    // Read Mode: único mecanismo disponível é simular o clique nos botões nativos da toolbar.
    // runNativePagePersistAfterPropertyWrite é fire-and-forget (retorna void);
    // a transição para 'saved' é baseada em timer controlado como proxy de conclusão.
    this.setStatus('saving');
    this._nativePersistTimer = window.setTimeout(() => {
      this._nativePersistTimer = undefined;
      console.log(`${LOG} Read Mode — persistência nativa iniciada — runNativePagePersistAfterPropertyWrite`);
      try {
        runNativePagePersistAfterPropertyWrite(
          this.domElement,
          true,
          800
        );

        this._savedTimer = window.setTimeout(() => {
          this._savedTimer = undefined;
          console.log(`${LOG} persistência assumida como concluída (após ${PERSIST_ASSUMED_DURATION_MS}ms)`);
          this.setStatus('saved');

          this._idleTimer = window.setTimeout(() => {
            this._idleTimer = undefined;
            this.setStatus('idle');
          }, SAVED_DISPLAY_DURATION_MS);
        }, PERSIST_ASSUMED_DURATION_MS);
      } catch (err) {
        console.error(`${LOG} erro na persistência nativa:`, err);
        this.setStatus('error');
      }
    }, 500);
  }

  private updateConfig(partial: Partial<IDynamicViewConfig>): void {
    const current = this.loadConfig();
    this.saveConfig({ ...current, ...partial });
  }

  public applyConfigPatch(partial: Partial<IDynamicViewConfig>): void {
    this.updateConfig(partial);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    this.clearAllPersistTimers();
    if (this._nativeEditSaveBridgeTeardown !== undefined) {
      this._nativeEditSaveBridgeTeardown();
      this._nativeEditSaveBridgeTeardown = undefined;
    }
    if (this._beforeUnloadHandler !== undefined) {
      window.removeEventListener('beforeunload', this._beforeUnloadHandler);
      this._beforeUnloadHandler = undefined;
    }
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
