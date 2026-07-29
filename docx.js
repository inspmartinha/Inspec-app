// Geração do relatório em .docx (OOXML nativo, via JSZip local — sem depender de internet).
// Sem logotipo ou identificação da empresa no documento gerado, por política de proteção de dados.

function xmlEscape(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function textParagraphs(text, opts) {
  opts = opts || {};
  const lines = String(text == null ? '' : text).split(/\r?\n/);
  const rprOpen = '<w:rPr>' +
    (opts.bold ? '<w:b/>' : '') +
    (opts.italic ? '<w:i/>' : '') +
    (opts.color ? `<w:color w:val="${opts.color}"/>` : '') +
    (opts.size ? `<w:sz w:val="${opts.size}"/><w:szCs w:val="${opts.size}"/>` : '') +
    '</w:rPr>';
  const pprAlign = opts.align ? `<w:jc w:val="${opts.align}"/>` : '';
  return lines.map(line => {
    const t = xmlEscape(line);
    const space = /^\s|\s$/.test(line) ? ' xml:space="preserve"' : '';
    return `<w:p><w:pPr>${pprAlign}</w:pPr><w:r>${rprOpen}<w:t${space}>${t}</w:t></w:r></w:p>`;
  }).join('');
}

function emptyParagraphXml() {
  return '<w:p/>';
}

let __imgSeq = 100;
function imageParagraphXml(rId, cx, cy, altText) {
  const id = __imgSeq++;
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Picture ${id}" descr="${xmlEscape(altText || '')}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function cellXml(contentXml, opts) {
  opts = opts || {};
  let tcPr = '<w:tcPr>';
  if (opts.width != null) tcPr += `<w:tcW w:w="${opts.width}" w:type="dxa"/>`;
  if (opts.gridSpan) tcPr += `<w:gridSpan w:val="${opts.gridSpan}"/>`;
  tcPr += '<w:tcBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>' +
    '</w:tcBorders>';
  if (opts.shade) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${opts.shade}"/>`;
  tcPr += `<w:vAlign w:val="${opts.vAlign || 'center'}"/>`;
  tcPr += '</w:tcPr>';
  const content = contentXml && contentXml.length ? contentXml : emptyParagraphXml();
  return `<w:tc>${tcPr}${content}</w:tc>`;
}

function rowXml(cellsArr) {
  return `<w:tr>${cellsArr.join('')}</w:tr>`;
}

function tableXml(rowsArr, colWidths) {
  const gridCols = colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>` +
    `</w:tblBorders><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${rowsArr.join('')}</w:tbl>`;
}

function contentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Default Extension="jpg" ContentType="image/jpeg"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`;
}

function relsRootXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" w:cs="Verdana"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    `</w:styles>`;
}

function documentRelsXml(rels) {
  let items = `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  rels.forEach(r => {
    items += `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${r.target}"/>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

function buildDocumentXmlWrapper(bodyXml) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<w:body>${bodyXml}` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1021" w:right="1021" w:bottom="1021" w:left="1021" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>` +
    `</w:body></w:document>`;
}

function dataUrlInfo(dataUrl) {
  const m = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/.exec(dataUrl || '');
  if (!m) return null;
  const mime = m[1];
  let ext = mime.split('/')[1];
  if (ext === 'jpg') ext = 'jpeg';
  if (ext !== 'png' && ext !== 'jpeg') ext = 'png';
  return { mime, ext, base64: m[2] };
}

function getImageSizeEmu(dataUrl, maxWidthIn) {
  maxWidthIn = maxWidthIn || 3.2;
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const ratio = (img.naturalHeight && img.naturalWidth) ? (img.naturalHeight / img.naturalWidth) : 0.75;
      const wIn = maxWidthIn;
      const hIn = wIn * ratio;
      resolve({ cx: Math.round(wIn * 914400), cy: Math.round(hIn * 914400) });
    };
    img.onerror = () => resolve({ cx: Math.round(maxWidthIn * 914400), cy: Math.round(maxWidthIn * 0.75 * 914400) });
    img.src = dataUrl;
  });
}

// Carrega o JSZip já vendorizado localmente em vendor/jszip.min.js (carregado via <script> no index.html).
function ensureJSZip() {
  if (!window.JSZip) throw new Error('JSZip não foi carregado (vendor/jszip.min.js ausente).');
  return Promise.resolve();
}

// sections: [{ title, blocks:[{ heading?, rows:[{label,value}], photos:[{src,caption}] }] }]
async function buildDocxBlob({ typeLabel, tag, unidade, local, sections }) {
  await ensureJSZip();

  const rels = [];
  let relCounter = 1;
  function nextRelId() { return 'rIdImg' + (relCounter++); }

  async function addImageParagraph(dataUrl, altText, maxWidthIn) {
    const info = dataUrlInfo(dataUrl);
    if (!info) return '';
    const relId = nextRelId();
    const filename = 'media/img' + relId + '.' + info.ext;
    rels.push({ id: relId, target: filename, base64: info.base64 });
    const size = await getImageSizeEmu(dataUrl, maxWidthIn);
    return imageParagraphXml(relId, size.cx, size.cy, altText);
  }

  const dateStr = new Date().toLocaleDateString('pt-BR');

  // Cabeçalho: título (sem logotipo)
  let bodyXml = tableXml([
    rowXml([
      cellXml(
        textParagraphs('RELATÓRIO DE ' + typeLabel.toUpperCase(), { bold: true, size: 26 }) +
        textParagraphs('Gerado em ' + dateStr, { size: 16, color: '666666' }),
        { gridSpan: 2, width: 9350, vAlign: 'center' }
      )
    ])
  ], [3000, 6350]);
  bodyXml += emptyParagraphXml();

  // Placa TAG / Unidade / Local
  bodyXml += tableXml([
    rowXml([
      cellXml(
        textParagraphs('TAG DO EQUIPAMENTO', { size: 14, color: 'E8A33D', bold: true }) +
        textParagraphs(tag || '—', { bold: true, size: 32, color: 'FFFFFF' }) +
        textParagraphs('UNIDADE: ' + (unidade || '—') + '    |    LOCAL: ' + (local || '—'), { size: 16, color: 'DCE2E7' }),
        { gridSpan: 2, width: 9350, shade: '1C2530' }
      )
    ])
  ], [3000, 6350]);
  bodyXml += emptyParagraphXml();

  // Seções
  for (let idx = 0; idx < sections.length; idx++) {
    const sec = sections[idx];
    const rows = [];
    rows.push(rowXml([cellXml(textParagraphs((idx + 1) + '. ' + sec.title, { bold: true, size: 23 }), { shade: 'C6D9F1', gridSpan: 2, width: 9350 })]));

    for (const block of sec.blocks) {
      if (block.heading) {
        rows.push(rowXml([cellXml(textParagraphs(block.heading, { bold: true }), { shade: 'EAF0FA', gridSpan: 2, width: 9350 })]));
      }
      for (const r of block.rows) {
        if (!r.label) continue;
        rows.push(rowXml([
          cellXml(textParagraphs(r.label, { bold: true }), { width: 3000, shade: 'F5F5F5' }),
          cellXml(textParagraphs(r.value || ''), { width: 6350 })
        ]));
      }
      if (block.photos.length) {
        let photosXml = '';
        for (const p of block.photos) {
          photosXml += await addImageParagraph(p.src, p.caption, 3.3);
          photosXml += textParagraphs(p.caption || '', { align: 'center', italic: true, size: 18, color: '666666' });
        }
        rows.push(rowXml([cellXml(photosXml, { gridSpan: 2, width: 9350 })]));
      }
    }
    bodyXml += tableXml(rows, [3000, 6350]) + emptyParagraphXml();
  }

  const documentXml = buildDocumentXmlWrapper(bodyXml);

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypesXml());
  zip.folder('_rels').file('.rels', relsRootXml());
  const wordFolder = zip.folder('word');
  wordFolder.file('document.xml', documentXml);
  wordFolder.file('styles.xml', stylesXml());
  wordFolder.folder('_rels').file('document.xml.rels', documentRelsXml(rels));
  const mediaFolder = wordFolder.folder('media');
  rels.forEach(r => {
    const fname = r.target.replace('media/', '');
    mediaFolder.file(fname, r.base64, { base64: true });
  });

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
}
