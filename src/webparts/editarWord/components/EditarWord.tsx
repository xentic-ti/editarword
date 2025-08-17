import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import { Dropdown, IDropdownOption, PrimaryButton, Label, Stack, MessageBar, MessageBarType } from '@fluentui/react';
import { IEditarWordProps } from './IEditarWordProps';
import { SPHttpClient } from '@microsoft/sp-http';
import PizZip from 'pizzip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const DEBUG_DOCX = true;

const EditarWord: React.FC<IEditarWordProps> = ({ context, folderServerRelativeUrl }) => {
  const [files, setFiles] = useState<IDropdownOption[]>([]);
  const [selectedFile, setSelectedFile] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<{ type: MessageBarType, text: string } | null>(null);

  // 1) Cargar lista de .docx en la carpeta
  const loadFiles = useCallback(async () => {
    try {
      setMsg(null);
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

  // 2) Descargar archivo como ArrayBuffer (ruta decodificada para evitar doble encoding)
  const downloadArrayBuffer = async (serverRelativeUrl: string): Promise<ArrayBuffer> => {
    const decoded = decodeURIComponent(serverRelativeUrl);
    const web = context.pageContext.web.absoluteUrl;
    const url = `${web}/_api/web/GetFileByServerRelativePath(DecodedUrl='${decoded}')/$value`;
    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    return res.arrayBuffer();
  };

  // 3) Agregar fila a la tabla (ID/Autor/Fecha) en document.xml
  const appendRowToDocXml = (docXmlString: string, fila: { id: string; autor: string; fecha: string }): string => {
    const parser = new DOMParser();
    const xml = parser.parseFromString(docXmlString, 'application/xml');

    const tables = Array.from(xml.getElementsByTagNameNS(W_NS, 'tbl'));
    let targetTbl: Element | null = null;

    for (const tbl of tables) {
      const firstRow = tbl.getElementsByTagNameNS(W_NS, 'tr')[0];
      if (!firstRow) continue;
      const texts = Array.from(firstRow.getElementsByTagNameNS(W_NS, 't')).map(t => (t.textContent || '').trim());
      const hasHeaders = texts.includes('ID') && texts.includes('Autor') && texts.includes('Fecha');
      if (hasHeaders) { targetTbl = tbl; break; }
    }

    if (!targetTbl) throw new Error('No se encontró una tabla con encabezados ID/Autor/Fecha en document.xml');

    const rowXml = `
      <w:tr xmlns:w="${W_NS}">
        <w:tc><w:p><w:r><w:t>${fila.id}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${fila.autor}</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>${fila.fecha}</w:t></w:r></w:p></w:tc>
      </w:tr>
    `.trim();

    const rowDoc = parser.parseFromString(rowXml, 'application/xml');
    const newRow = rowDoc.getElementsByTagNameNS(W_NS, 'tr')[0];
    const imported = xml.importNode(newRow, true);
    targetTbl.appendChild(imported);

    return new XMLSerializer().serializeToString(xml);
  };

  // ===== Helpers de diagnóstico y edición de título =====

  function getWVal(el?: Element | null): string | null {
    if (!el) return null;
    return el.getAttributeNS(W_NS, 'val') || el.getAttribute('w:val') || el.getAttribute('val');
  }

  function listSdtTags(xmlString: string, label: string) {
    const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
    const sdts = Array.from(doc.getElementsByTagNameNS(W_NS, 'sdt'));
    const info = sdts.map((sdt, i) => {
      const pr = sdt.getElementsByTagNameNS(W_NS, 'sdtPr')[0];
      const tagEl = pr?.getElementsByTagNameNS(W_NS, 'tag')[0];
      const aliasEl = pr?.getElementsByTagNameNS(W_NS, 'alias')[0];
      const tag = getWVal(tagEl);
      const alias = getWVal(aliasEl);
      const t = sdt.getElementsByTagNameNS(W_NS, 't')[0]?.textContent ?? '';
      return { idx: i, tag, alias, sampleText: t.slice(0, 60) };
    });
    console.groupCollapsed(`[SDT] ${label} (${sdts.length})`);
    console.table(info);
    console.groupEnd();
  }

  function trySetSdtText(xmlString: string, tagOrAlias: string, newText: string): { updated: string; changed: boolean } {
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlString, 'application/xml');

    const sdts = Array.from(xml.getElementsByTagNameNS(W_NS, 'sdt'));
    for (const sdt of sdts) {
      const pr = sdt.getElementsByTagNameNS(W_NS, 'sdtPr')[0];
      if (!pr) continue;

      const tagEl = pr.getElementsByTagNameNS(W_NS, 'tag')[0];
      const aliasEl = pr.getElementsByTagNameNS(W_NS, 'alias')[0];
      const tagVal = getWVal(tagEl);
      const aliasVal = getWVal(aliasEl);

      if (tagVal === tagOrAlias || aliasVal === tagOrAlias) {
        const content = sdt.getElementsByTagNameNS(W_NS, 'sdtContent')[0] || sdt;
        let t = content.getElementsByTagNameNS(W_NS, 't')[0];
        if (!t) {
          const p = xml.createElementNS(W_NS, 'w:p');
          const r = xml.createElementNS(W_NS, 'w:r');
          t = xml.createElementNS(W_NS, 'w:t');
          r.appendChild(t); p.appendChild(r); content.appendChild(p);
        }
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        t.textContent = newText;

        console.log(`[SDT] Reemplazado tag/alias="${tagOrAlias}" → "${newText}"`);
        return { updated: new XMLSerializer().serializeToString(xml), changed: true };
      }
    }
    console.warn(`[SDT] No se encontró tag/alias="${tagOrAlias}" en este XML`);
    return { updated: xmlString, changed: false };
  }

  // Fallback por estilo (por si el título no es SDT)
  function trySetParagraphByStyle(xmlString: string, styleIds: string[], newText: string): { updated: string; changed: boolean } {
    const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, 'p'));
    for (const p of paragraphs) {
      const pPr = p.getElementsByTagNameNS(W_NS, 'pPr')[0];
      const pStyle = pPr?.getElementsByTagNameNS(W_NS, 'pStyle')[0];
      const val = pStyle ? (pStyle.getAttributeNS(W_NS, 'val') || pStyle.getAttribute('w:val') || pStyle.getAttribute('val')) : null;
      if (val && styleIds.includes(val)) {
        Array.from(p.childNodes).filter(n => (n as Element).localName === 'r').forEach(n => p.removeChild(n));
        const r = doc.createElementNS(W_NS, 'w:r');
        const t = doc.createElementNS(W_NS, 'w:t');
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        t.textContent = newText;
        r.appendChild(t);
        p.appendChild(r);
        console.log(`[STYLE] Reemplazado pStyle="${val}" → "${newText}"`);
        return { updated: new XMLSerializer().serializeToString(doc), changed: true };
      }
    }
    console.warn('[STYLE] No se encontró párrafo con estilos:', styleIds);
    return { updated: xmlString, changed: false };
  }

  // Actualiza docProps/core.xml -> <dc:title>
  function setCoreTitle(coreXml: string, newTitle: string): string {
    const parser = new DOMParser();
    const xml = parser.parseFromString(coreXml, 'application/xml');
    let el = xml.getElementsByTagNameNS(DC_NS, 'title')[0];
    if (!el) {
      el = xml.createElementNS(DC_NS, 'dc:title');
      xml.documentElement.appendChild(el);
    }
    el.textContent = newTitle;
    return new XMLSerializer().serializeToString(xml);
  }

  // (Opcional) Subir dump de XML para inspección
  async function uploadTextDebug(folderSrvRel: string, name: string, text: string) {
    const folderDecoded = decodeURIComponent(folderSrvRel);
    const web = context.pageContext.web.absoluteUrl;
    const url = `${web}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${folderDecoded}')/Files/AddUsingPath(DecodedUrl='${name}',Overwrite=true)`;
    await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: new Blob([text], { type: 'text/plain' })
    });
  }

  // 4) Guardar archivo en SharePoint (nuevo nombre)
  const uploadNewFile = async (folderSrvRel: string, originalName: string, blob: Blob): Promise<string> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const origin = new URL(webUrl).origin;

    const folderDecoded = decodeURIComponent(folderSrvRel);

    const parts = folderDecoded.split('/').filter(Boolean);
    const sitesIdx = parts.indexOf('sites');
    const siteRel = sitesIdx >= 0 ? `/${parts.slice(0, sitesIdx + 2).join('/')}` : `/${parts[0]}`;
    const targetWebAbs = `${origin}${siteRel}`;

    const sameSite = webUrl.toLowerCase().startsWith(targetWebAbs.toLowerCase());
    const newName = originalName.replace(/\.docx$/i, '') + '-editado.docx';

    const apiBase = sameSite ? `${webUrl}/_api/web` : `${webUrl}/_api/SP.AppContextSite(@target)/web`;
    // Forma canónica: filename solo en AddUsingPath del folder
    const tail = `GetFolderByServerRelativePath(DecodedUrl='${folderDecoded}')/Files/AddUsingPath(DecodedUrl='${newName}',Overwrite=true)`;
    const url = sameSite ? `${apiBase}/${tail}` : `${apiBase}/${tail}?@target='${encodeURIComponent(targetWebAbs)}'`;

    const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: blob
    });

    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`Error al subir (${res.status}): ${text || res.statusText}`);
    }
    return `${folderDecoded}/${newName}`;
  };

  // Separar documento en 2: primera página y resto
  const splitDocByFirstPage = async () => {
    if (!selectedFile) {
      setMsg({ type: MessageBarType.warning, text: 'Selecciona un archivo primero.' });
      return;
    }

    setBusy(true);
    setMsg(null);

    try {
      // Descargar DOCX original
      const ab = await downloadArrayBuffer(selectedFile);
      const zip = new PizZip(ab);

      const docXml = zip.file("word/document.xml")?.asText();
      if (!docXml) throw new Error("No se encontró word/document.xml");

      // Buscar primer salto de página
      const parts = docXml.split(/<w:br[^>]*w:type="page"[^>]*\/>/i);

      if (parts.length < 2) {
        throw new Error("No se encontró salto de página en el documento");
      }

      const [firstPageXml, restXml] = parts;

      // ---- Primera página ----
      const zip1 = new PizZip(ab); // copia
      zip1.file("word/document.xml", firstPageXml + "</w:body></w:document>");
      const blob1: Blob = zip1.generate({ type: "blob" });
      const name1 = selectedFile.split("/").pop()!.replace(/\.docx$/i, "-pag1.docx");
      await uploadNewFile(folderServerRelativeUrl, name1, blob1);

      // ---- Resto del documento ----
      const zip2 = new PizZip(ab); // copia
      zip2.file("word/document.xml", restXml + "</w:body></w:document>");
      const blob2: Blob = zip2.generate({ type: "blob" });
      const name2 = selectedFile.split("/").pop()!.replace(/\.docx$/i, "-resto.docx");
      await uploadNewFile(folderServerRelativeUrl, name2, blob2);

      setMsg({ type: MessageBarType.success, text: `Se guardaron: ${name1} y ${name2}` });
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: `Error al dividir: ${e.message || e}` });
    } finally {
      setBusy(false);
    }
  };


  // 5) Flujo principal
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

      // c) document.xml
      const docXml = zip.file('word/document.xml')?.asText();
      if (!docXml) throw new Error('No se encontró word/document.xml');

      if (DEBUG_DOCX) listSdtTags(docXml, 'word/document.xml');

      // d1) Agregar fila a la tabla
      let updatedXml = appendRowToDocXml(docXml, { id: '1233', autor: 'Pedro', fecha: '12/08/2025' });

      // d2) Título nuevo
      const nuevoTitulo = 'Título actualizado desde SPFx';

      // 2a) Intento en document.xml por SDT
      let r = trySetSdtText(updatedXml, 'TituloDocumento', nuevoTitulo);
      updatedXml = r.updated;

      // Fallback por estilo si no es SDT
      if (!r.changed) {
        const rr = trySetParagraphByStyle(updatedXml, ['TituloDocumento', 'Title'], nuevoTitulo);
        updatedXml = rr.updated;
        if (rr.changed) console.log('[DOC] Título reemplazado por estilo.');
      }

      // Aplicar cambios a document.xml
      zip.file('word/document.xml', updatedXml);

      let cambios = { documentSdt: r.changed, headers: 0, footers: 0, core: false };

      // 2b) Headers
      const headerFiles = zip.file(/word\/header\d+\.xml/);
      for (const f of headerFiles) {
        const x = f.asText();
        if (DEBUG_DOCX) listSdtTags(x, f.name);
        const rr = trySetSdtText(x, 'TituloDocumento', nuevoTitulo);
        if (rr.changed) {
          zip.file(f.name, rr.updated);
          cambios.headers++;
        }
      }

      // 2c) Footers
      const footerFiles = zip.file(/word\/footer\d+\.xml/);
      for (const f of footerFiles) {
        const x = f.asText();
        if (DEBUG_DOCX) listSdtTags(x, f.name);
        const rr = trySetSdtText(x, 'TituloDocumento', nuevoTitulo);
        if (rr.changed) {
          zip.file(f.name, rr.updated);
          cambios.footers++;
        }
      }

      // 2d) Propiedad de documento Title
      const core = zip.file('docProps/core.xml')?.asText();
      if (core) {
        zip.file('docProps/core.xml', setCoreTitle(core, nuevoTitulo));
        cambios.core = true;
      }

      console.info('[DOCX] Título actualizado:', cambios);
      setMsg({ type: MessageBarType.info, text: `Cambios → docSDT=${cambios.documentSdt} | headers=${cambios.headers} | footers=${cambios.footers} | core=${cambios.core}` });

      // (opcional) Dump del document.xml ya editado
      if (DEBUG_DOCX) {
        await uploadTextDebug(folderServerRelativeUrl, 'document-xml-dump.txt', updatedXml);
      }

      // f) Generar .docx (PizZip es síncrono)
      const blob: Blob = zip.generate({ type: 'blob' });

      // g) Subir nuevo archivo a la misma carpeta
      const fileName = selectedFile.split('/').pop()!;
      const newFileSrvRelUrl = await uploadNewFile(folderServerRelativeUrl, fileName, blob);

      setMsg({ type: MessageBarType.success, text: `Archivo guardado: ${newFileSrvRelUrl}` });

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
      <PrimaryButton
        text="Dividir en 1ra página y resto"
        onClick={splitDocByFirstPage}
        disabled={!selectedFile || busy}
      />

    </Stack>
  );
};

export default EditarWord;
