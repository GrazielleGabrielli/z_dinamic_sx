import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';

import '../../assets/dist/tailwind.css';
import { getSP } from '../../services/core/sp';
import PbiTextoConsulta from './components/PbiTextoConsulta';
import { IPbiTextoConsultaProps } from './components/IPbiTextoConsultaProps';

export interface IPbiTextoConsultaWebPartProps {}

export default class PbiTextoConsultaWebPart extends BaseClientSideWebPart<IPbiTextoConsultaWebPartProps> {
  protected onInit(): Promise<void> {
    getSP(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IPbiTextoConsultaProps> = React.createElement(PbiTextoConsulta, {
      siteUrl: this.context.pageContext.web.absoluteUrl
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
