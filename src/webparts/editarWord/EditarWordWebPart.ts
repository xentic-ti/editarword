import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import EditarWord from './components/EditarWord';
import { IEditarWordProps } from './components/IEditarWordProps';

export interface IEditarWordWebPartProps {}

export default class EditarWordWebPart extends BaseClientSideWebPart<IEditarWordWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IEditarWordProps> = React.createElement(EditarWord, {
      context: this.context,
      folderServerRelativeUrl: "/sites/GrupoDePrueba2/Documentos compartidos"
    });

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
