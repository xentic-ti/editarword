import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import { Dropdown, IDropdownOption, PrimaryButton, Label, Stack, MessageBar, MessageBarType } from '@fluentui/react';
import { IEditarWordProps } from './IEditarWordProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import PizZip from 'pizzip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const EditarWord: React.FC<IEditarWordProps> = ({ context, folderServerRelativeUrl }) => {
  const [files, setFiles] = useState<IDropdownOption[]>([]);
  const [selectedFile, setSelectedFile] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<{ type: MessageBarType, text: string } | null>(null);

  // 1) Cargar lista de .docx en la carpeta
  const loadFiles = useCallback(async () => {
    try {
      setMsg(null);
      // REST: /_api/web/GetFolderByServerRelativeUrl('...')/Files?$select=Name,ServerRelativeUrl&$filter=substringof('.docx',Name)
      const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl(@p)/Files?$select=Name,ServerRelativeUrl&$filter=substringof('.docx',Name)&@p='${encodeURIComponent(folderServerRelativeUrl)}'`;

      const r = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const j = await r.json();
      const opts: IDropdownOption[] = (j.value || []).map((f: any) => ({
        key: f.ServerRelativeUrl,
        text: f.Name
      }));
      setFiles(opts);
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: `Error al listar archivos: ${e.message || e}` });
    }
  }, [context, folderServerRelativeUrl]);

  useEffect(() => { loadFiles(); }, [loadFiles]);

  // 2) Descargar archivo como ArrayBuffer
  const downloadArrayBuffer = async (serverRelativeUrl: string): Promise<ArrayBuffer> => {
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(serverRelativeUrl)}')/$value`;
    const res: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    return await res.arrayBuffer();
  };

  // 3) Editar document.xml -> agregar fila a la tabla (ID/Autor/Fecha)
  const appendRowToDocXml = (docXmlString: string, fila: { id: string; autor: string; fecha: string }): string => {
    const parser = new DOMParser();
    const xml = parser.parseFromString(docXmlString, 'application/xml');

    // Buscar la tabla cuyo primer row tiene celdas "ID", "Autor", "Fecha"
    const tables = Array.from(xml.getElementsByTagNameNS(W_NS, 'tbl'));
    let targetTbl: Element | null = null;

    for (const tbl of tables) {
      const firstRow = tbl.getElementsByTagNameNS(W_NS, 'tr')[0];
      if (!firstRow) continue;
      const texts = Array.from(firstRow.getElementsByTagNameNS(W_NS, 't')).map(t => (t.textContent || '').trim());
      const hasHeaders = texts.includes('ID') && texts.includes('Autor') && texts.includes('Fecha');
      if (hasHeaders) { targetTbl = tbl; break; }
    }

    if (!targetTbl) {
      throw new Error('No se encontró una tabla con encabezados ID/Autor/Fecha en document.xml');
    }

    const rowXml = `
      <w:tr xmlns:w="${W_NS}">
        <w:tc><w:p><w:r><w:t>${fila.id}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${fila.autor}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${fila.fecha}</w:t></w:r></w:p></w:tc>
      </w:tr>
    `.trim();

    const rowDoc = parser.parseFromString(rowXml, 'application/xml');
    const newRow = rowDoc.getElementsByTagNameNS(W_NS, 'tr')[0];

    // Importar al document.xml y anexar
    const imported = xml.importNode(newRow, true);
    targetTbl.appendChild(imported);

    const ser = new XMLSerializer();
    return ser.serializeToString(xml);
  };

  // 4) Guardar archivo en SharePoint (nuevo nombre)
// Sube un nuevo .docx a la carpeta indicada. Devuelve la URL server-relative del nuevo archivo.
const uploadNewFile = async (
  folderSrvRel: string,     // p.ej. "/sites/GrupoDePrueba2/Documentos compartidos"  (SIN %20)
  originalName: string,
  blob: Blob
): Promise<string> => {

  const webUrl = context.pageContext.web.absoluteUrl;           // web actual
  const origin = new URL(webUrl).origin;

  // Normaliza la carpeta (sin %20)
  const folderDecoded = decodeURIComponent(folderSrvRel);

  // Deriva el sitio de destino a partir de la carpeta: "/sites/GrupoDePrueba2"
  const parts = folderDecoded.split('/').filter(Boolean);
  const sitesIdx = parts.indexOf('sites');
  const siteRel = sitesIdx >= 0 ? `/${parts.slice(0, sitesIdx + 2).join('/')}` : `/${parts[0]}`;
  const targetWebAbs = `${origin}${siteRel}`;

  const sameSite = webUrl.toLowerCase().startsWith(targetWebAbs.toLowerCase());

  const newName = originalName.replace(/\.docx$/i, '') + '-editado.docx';

  // Armamos el endpoint (mismo sitio vs otro sitio)
  const apiBase = sameSite
    ? `${webUrl}/_api/web`
    : `${webUrl}/_api/SP.AppContextSite(@target)/web`;

  const tail = `GetFolderByServerRelativePath(DecodedUrl='${folderDecoded}')/Files/AddUsingPath(DecodedUrl='${folderDecoded}/${newName}',Overwrite=true)`;

  const url = sameSite
    ? `${apiBase}/${tail}`
    : `${apiBase}/${tail}?@target='${encodeURIComponent(targetWebAbs)}'`;

  const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    },
    body: blob
  });

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Error al subir (${res.status}): ${text || res.statusText}`);
  }

  return `${folderDecoded}/${newName}`;
};



  // 5) Flujo principal al hacer clic
  const onEditAndSave = useCallback(async () => {
  if (!selectedFile) {
    setMsg({ type: MessageBarType.warning, text: 'Selecciona un archivo primero.' });
    return;
  }

  setBusy(true);
  setMsg(null);

  try {
    // a) Descargar DOCX
    const ab = await downloadArrayBuffer(selectedFile);

    // b) Abrir ZIP
    const zip = new PizZip(ab);

    // c) Leer word/document.xml
    const docXml = zip.file('word/document.xml')?.asText();
    if (!docXml) throw new Error('No se encontró word/document.xml');

    // d) Agregar la fila auto
    const updatedXml = appendRowToDocXml(docXml, {
      id: '1233',
      autor: 'Pedro',
      fecha: '12/08/2025'
    });

    // e) Reemplazar XML
    zip.file('word/document.xml', updatedXml);

    // f) Generar .docx (PizZip es síncrono)
    const blob: Blob = zip.generate({ type: 'blob' });

    // g) Subir nuevo archivo a la misma carpeta
    const fileName = selectedFile.split('/').pop()!;
    const newFileSrvRelUrl = await uploadNewFile(
      folderServerRelativeUrl, // OJO: sin %20, ruta decodificada
      fileName,
      blob
    );

    setMsg({
      type: MessageBarType.success,
      text: `Archivo guardado: ${newFileSrvRelUrl}`
    });

  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    setMsg({ type: MessageBarType.error, text: `Error: ${message}` });
  } finally {
    setBusy(false);
  }
}, [selectedFile, folderServerRelativeUrl, context]);

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Label>Selecciona un archivo Word (.docx) en: <b>{folderServerRelativeUrl}</b></Label>
      <Dropdown
        placeholder="Elegir archivo…"
        options={files}
        selectedKey={selectedFile || undefined}
        onChange={(_, opt) => setSelectedFile((opt?.key as string) || null)}
      />
      <PrimaryButton text={busy ? 'Procesando…' : 'Editar y Guardar'} onClick={onEditAndSave} disabled={!selectedFile || busy} />
      <PrimaryButton text="Recargar lista" onClick={loadFiles} disabled={busy} styles={{ root: { width: 140 } }} />
      {msg && <MessageBar messageBarType={msg.type}>{msg.text}</MessageBar>}
    </Stack>
  );
};

export default EditarWord;
