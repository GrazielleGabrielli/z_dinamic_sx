import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import '../../assets/dist/tailwind.css';
import DinamicApp from './components/DinamicApp';
import { IDinamicAppProps } from './components/IDinamicAppProps';
import { IDynamicViewConfig, IDynamicViewWebPartProps, TViewMode } from './core/config/types';
import { parseConfig } from './core/config/validators';
import { getDefaultConfig } from './core/config/utils';
import { getSP, getGraph } from './pnpConfig';
import { runNativePagePersistAfterPropertyWrite } from './core/sharePoint/sharePointPageToolbarDom';

export abstract class DinamicWebPartBase extends BaseClientSideWebPart<IDynamicViewWebPartProps> {
  private _nativePersistTimer: number | undefined;

  protected abstract getForcedMode(): TViewMode | undefined;

  protected onInit(): Promise<void> {
    getSP(this.context);
    getGraph(this.context);
    return super.onInit();
  }

  public render(): void {
    const forcedMode = this.getForcedMode();
    const element: React.ReactElement<IDinamicAppProps> = React.createElement(DinamicApp, {
      configJson: this.properties.configJson ?? '',
      siteUrl: this.context.pageContext.web.serverRelativeUrl,
      instanceScopeId: this.instanceId,
      onSaveConfig: (config: IDynamicViewConfig) => this.saveConfig(config),
      openAiApiKey: this.properties.openAiApiKey ?? '',
      ...(forcedMode !== undefined ? { forcedMode } : {}),
    });

    ReactDom.render(element, this.domElement);
  }

  private loadConfig(): IDynamicViewConfig {
    return parseConfig(this.properties.configJson) ?? getDefaultConfig();
  }

  private saveConfig(config: IDynamicViewConfig): void {
    this.properties.configJson = JSON.stringify(config);
    this.render();
    if (this.displayMode === DisplayMode.Read) return;
    if (this._nativePersistTimer !== undefined) {
      window.clearTimeout(this._nativePersistTimer);
    }
    this._nativePersistTimer = window.setTimeout(() => {
      this._nativePersistTimer = undefined;
      try {
        runNativePagePersistAfterPropertyWrite(
          this.domElement,
          false,
          800
        );
      } catch {}
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
    if (this._nativePersistTimer !== undefined) {
      window.clearTimeout(this._nativePersistTimer);
      this._nativePersistTimer = undefined;
    }
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configurações da web part' },
          groups: [
            {
              groupName: 'Integração com IA',
              groupFields: [
                PropertyPaneTextField('openAiApiKey', {
                  label: 'Chave de API OpenAI',
                  description: 'Necessária para o assistente de configuração com IA (Crie com IA).',
                  placeholder: 'sk-...',
                }),
                PropertyPaneTextField('configJson', {
                  label: 'Configuração atual (JSON)',
                  description: 'Leitura apenas — backup da configuração gerada pelas telas e pela IA.',
                  multiline: true,
                  rows: 12,
                  disabled: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
