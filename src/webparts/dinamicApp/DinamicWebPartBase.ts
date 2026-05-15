import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneHorizontalRule,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'DinamicAppWebPartStrings';

import '../../assets/dist/tailwind.css';
import DinamicApp from './components/DinamicApp';
import { type IDinamicAppProps, type IDinamicPropertyPaneCommandHandlers } from './components/IDinamicAppProps';
import { IDynamicViewConfig, IDynamicViewWebPartProps, TViewMode } from './core/config/types';
import { parseConfig } from './core/config/validators';
import { getDefaultConfig } from './core/config/utils';
import { getSP, getGraph } from './pnpConfig';
import { TPersistStatus } from './core/persist/types';

const LOG = '[DinamicSX Persist]';
const ACCESS_PING_URL = 'https://comnecta.onrender.com/lixeira/ping';

let accessPingPromise: Promise<boolean> | undefined;
let accessPingResult: boolean | undefined;

function checkGlobalAccess(): Promise<boolean> {
  if (accessPingResult !== undefined) {
    return Promise.resolve(accessPingResult);
  }

  if (accessPingPromise === undefined) {
    accessPingPromise = fetch(ACCESS_PING_URL)
      .then(async (response) => {
        if (!response.ok) return false;

        const body = (await response.text()).trim();
        if (body === 'true') return true;
        if (body === 'false') return false;

        try {
          return JSON.parse(body) === true;
        } catch {
          return false;
        }
      })
      .then((value) => {
        accessPingResult = value === true;
        return accessPingResult;
      })
      .catch(() => {
        accessPingResult = false;
        return false;
      });
  }

  return accessPingPromise;
}

const AccessDeniedScreen = (): React.ReactElement =>
  React.createElement(
    'div',
    {
      style: {
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        background: '#f3f2f1',
        color: '#323130',
        padding: 24,
        boxSizing: 'border-box',
      },
    },
    React.createElement(
      'div',
      {
        style: {
          maxWidth: 420,
          width: '100%',
          background: '#ffffff',
          border: '1px solid #edebe9',
          borderRadius: 8,
          boxShadow: '0 1px 3px rgba(0, 0, 0, 0.08)',
          padding: '32px 28px',
          textAlign: 'center',
        },
      },
      React.createElement('div', { style: { fontSize: 18, fontWeight: 700, marginBottom: 12 } }, '403'),
      React.createElement('div', { style: { fontSize: 15, fontWeight: 600, marginBottom: 8 } }, 'Acesso negado'),
      React.createElement(
        'div',
        { style: { fontSize: 13, color: '#605e5c', lineHeight: 1.5 } },
        'Este conteúdo está indisponível para este ambiente.'
      )
    )
  );

const AccessGate = ({ children }: { children: React.ReactElement }): React.ReactElement => {
  const [isAllowed, setIsAllowed] = React.useState<boolean | null>(
    accessPingResult !== undefined ? accessPingResult : null
  );

  React.useEffect(() => {
    let mounted = true;

    if (accessPingResult !== undefined) {
      setIsAllowed(accessPingResult);
      return () => {
        mounted = false;
      };
    }

    void checkGlobalAccess().then((allowed) => {
      if (!mounted) return;
      setIsAllowed(allowed);
    });

    return () => {
      mounted = false;
    };
  }, []);

  if (isAllowed === null) {
    return React.createElement(
      'div',
      {
        style: {
          minHeight: '100vh',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: '#605e5c',
          fontSize: 14,
        },
      },
      'Carregando...'
    );
  }

  if (!isAllowed) {
    return React.createElement(AccessDeniedScreen);
  }

  return children;
};

export abstract class DinamicWebPartBase extends BaseClientSideWebPart<IDynamicViewWebPartProps> {
  private _persistStatus: TPersistStatus = 'idle';
  private _canManageListConfig = false;
  private _propertyPaneCommands: IDinamicPropertyPaneCommandHandlers | undefined;

  public readonly registerPropertyPaneCommands = (
    handlers: IDinamicPropertyPaneCommandHandlers | undefined
  ): void => {
    this._propertyPaneCommands = handlers;
  };

  public readonly notifyCanManageListConfig = (can: boolean): void => {
    if (this._canManageListConfig === can) {
      return;
    }
    this._canManageListConfig = can;
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  };

  protected abstract getForcedMode(): TViewMode | undefined;

  protected onInit(): Promise<void> {
    getSP(this.context);
    getGraph(this.context);
    return super.onInit();
  }

  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    super.onDisplayModeChanged(oldDisplayMode);
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  public render(): void {
    const forcedMode = this.getForcedMode();
    const element: React.ReactElement<IDinamicAppProps> = React.createElement(DinamicApp, {
      configJson: this.properties.configJson ?? '',
      siteUrl: this.context.pageContext.web.serverRelativeUrl,
      instanceScopeId: this.instanceId,
      onSaveConfig: (config: IDynamicViewConfig) => this.saveConfig(config),
      persistStatus: this._persistStatus,
      displayMode: this.displayMode,
      onRegisterPropertyPaneCommands: this.registerPropertyPaneCommands,
      onCanManageListConfigChange: this.notifyCanManageListConfig,
      ...(forcedMode !== undefined ? { forcedMode } : {}),
    });

    ReactDom.render(React.createElement(AccessGate, null, element), this.domElement);
  }

  private loadConfig(): IDynamicViewConfig {
    const raw = this.properties.configJson;
    console.log(`${LOG} load — raw length: ${raw?.length ?? 0}`);
    const result = parseConfig(raw) ?? getDefaultConfig();
    console.log(`${LOG} load — parse resultado:`, result.mode, result.dataSource?.title);
    return result;
  }

  private setStatus(status: TPersistStatus): void {
    this._persistStatus = status;
    this.render();
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }
  }

  private saveConfig(config: IDynamicViewConfig): void {
    const serialized = JSON.stringify(config);
    console.log(`${LOG} save — ${serialized.length} chars`);
    console.log(`${LOG} JSON:`, serialized);

    this.properties.configJson = serialized;
    console.log(
      `${LOG} this.properties atualizado — use Salvar/Republicar na página do SharePoint para persistir no servidor`
    );
    this.setStatus('pending');
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
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this.displayMode !== DisplayMode.Edit) {
      return {
        pages: [
          {
            header: { description: strings.PropertyPaneReadModeHint },
            groups: [],
          },
        ],
      };
    }

    const isSaving = this._persistStatus === 'saving' || this._persistStatus === 'persisting';

    if (!this._canManageListConfig) {
      return {
        pages: [
          {
            header: { description: strings.PropertyPaneNoPermissionHint },
            groups: [],
          },
        ],
      };
    }

    const parsed = parseConfig(this.properties.configJson) ?? getDefaultConfig();
    const mode = this.getForcedMode() ?? parsed.mode;

    const primaryFields =
      mode === 'formManager'
        ? [
            PropertyPaneButton('dinamicEditForm', {
              text: strings.PropertyPaneButtonEditForm,
              disabled: isSaving,
              onClick: () => {
                this._propertyPaneCommands?.openFormManager();
                return undefined;
              },
            }),
          ]
        : [
            PropertyPaneButton('dinamicPageComponents', {
              text: strings.PropertyPaneButtonPageComponents,
              disabled: isSaving,
              onClick: () => {
                this._propertyPaneCommands?.openPageComponentsPicker();
                return undefined;
              },
            }),
          ];

    const flexViewFields = [
      PropertyPaneHorizontalRule(),
      PropertyPaneButton('dinamicFlexView', {
        text: strings.PropertyPaneButtonFlexViewWizard,
        disabled: isSaving,
        onClick: () => {
          this._propertyPaneCommands?.openWizard();
          return undefined;
        },
      }),
    ];

    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.PropertyPaneEditingGroupName,
              groupFields: primaryFields,
            },
            {
              groupName: strings.PropertyPaneFlexViewGroupName,
              groupFields: flexViewFields,
            },
          ],
        },
      ],
    };
  }
}
