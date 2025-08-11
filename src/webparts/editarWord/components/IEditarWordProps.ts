import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEditarWordProps {
  context: WebPartContext;
  /** Server Relative URL de la carpeta. Ej: /sites/GrupoDePrueba2/Documentos%20compartidos */
  folderServerRelativeUrl: string;
}