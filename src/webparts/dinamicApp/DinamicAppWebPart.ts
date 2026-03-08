import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import '../../assets/dist/tailwind.css';
import DinamicApp from './components/DinamicApp';
import { IDinamicAppProps } from './components/IDinamicAppProps';
import { IDynamicViewConfig, IDynamicViewWebPartProps } from './core/config/types';
import { parseConfig } from './core/config/validators';
import { getDefaultConfig } from './core/config/utils';
import { getSP, getGraph } from './pnpConfig';

export default class DinamicAppWebPart extends BaseClientSideWebPart<IDynamicViewWebPartProps> {

  protected onInit(): Promise<void> {
    getSP(this.context);
    getGraph(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDinamicAppProps> = React.createElement(DinamicApp, {
      configJson: this.properties.configJson ?? '',
      siteUrl: this.context.pageContext.web.serverRelativeUrl,
      onSaveConfig: (config: IDynamicViewConfig) => this.saveConfig(config),
    });

    ReactDom.render(element, this.domElement);
  }

  /**
   * Retorna a config persistida ou a config default se não houver JSON válido.
   */
  private loadConfig(): IDynamicViewConfig {
    return parseConfig(this.properties.configJson) ?? getDefaultConfig();
  }

  /**
   * Serializa e persiste a config em configJson, depois re-renderiza.
   */
  private saveConfig(config: IDynamicViewConfig): void {
    this.properties.configJson = JSON.stringify(config);
    this.render();
  }

  /**
   * Mescla parcialmente a config atual e salva o resultado.
   */
  private updateConfig(partial: Partial<IDynamicViewConfig>): void {
    const current = this.loadConfig();
    this.saveConfig({ ...current, ...partial });
  }

  // Expõe updateConfig para uso em testes ou extensões futuras
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
    return { pages: [] };
  }
}
