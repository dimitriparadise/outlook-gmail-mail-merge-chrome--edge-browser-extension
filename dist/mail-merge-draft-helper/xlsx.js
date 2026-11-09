// Minimal .xlsx reader for recipient lists.
// Unzips the workbook with the browser's DecompressionStream and reads cell text from the sheet XML.
// Formulas are not evaluated (the value Excel last saved is used) and dates stay as Excel serial numbers.

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function unzipEntries(input, wanted) {
  const bytes = ArrayBuffer.isView(input) ? new Uint8Array(input.buffer, input.byteOffset, input.byteLength) : new Uint8Array(input);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  // The end-of-central-directory record sits at the end, after an optional comment of up to 64 KB.
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 0xffff); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('Not a zip file.');

  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const decoder = new TextDecoder();
  const files = new Map();

  for (let n = 0; n < count; n++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Damaged zip directory.');
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    offset += 46 + nameLength + extraLength + commentLength;
    if (!wanted(name)) continue;

    const start = localOffset + 30 + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true);
    const data = bytes.subarray(start, start + compressedSize);
    if (method === 0) files.set(name, decoder.decode(data));
    else if (method === 8) files.set(name, decoder.decode(await inflateRaw(data)));
    else throw new Error(`Unsupported zip compression method ${method}.`);
  }
  return files;
}

function decodeXmlEntities(text) {
  // XML normalizes literal line breaks to \n; an escaped &#13; still yields \r.
  return text.replace(/\r\n?/g, '\n').replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (match, code) => {
    const lower = code.toLowerCase();
    if (lower === 'amp') return '&';
    if (lower === 'lt') return '<';
    if (lower === 'gt') return '>';
    if (lower === 'quot') return '"';
    if (lower === 'apos') return "'";
    return String.fromCodePoint(lower[1] === 'x' ? parseInt(lower.slice(2), 16) : parseInt(lower.slice(1), 10));
  });
}

function xmlAttributes(source) {
  const attrs = {};
  for (const [, key, value] of String(source).matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) {
    attrs[key] = decodeXmlEntities(value);
  }
  return attrs;
}

// Joins the <t> runs of a string item, skipping phonetic hints (<rPh>). Tags may carry a namespace prefix.
function xmlText(fragment) {
  return [...String(fragment)
    .replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, '')
    .matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)]
    .map(match => decodeXmlEntities(match[1]))
    .join('');
}

function columnIndex(cellRef) {
  const letters = /^[A-Z]+/i.exec(cellRef || '')?.[0];
  if (!letters) return -1;
  return [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function cellValue(attrs, inner, sharedStrings) {
  const raw = /<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/.exec(inner)?.[1];
  switch (attrs.t) {
    case 's':
      return sharedStrings[Number(raw)] ?? '';
    case 'inlineStr':
      return xmlText(inner);
    case 'b':
      return raw === '1' ? 'TRUE' : 'FALSE';
    case 'str':
    case 'e':
      return decodeXmlEntities(raw ?? '');
    default: {
      if (raw === undefined) return '';
      const number = Number(raw);
      // Print stored floats like 3.2250000000000001 the way Excel shows them.
      return raw.trim() !== '' && Number.isFinite(number) ? String(number) : decodeXmlEntities(raw);
    }
  }
}

function parseSheetRows(xml, sharedStrings) {
  const rows = [];
  let rowNumber = 0;
  for (const rowMatch of xml.matchAll(/<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g)) {
    const rowAttrs = xmlAttributes(rowMatch[1]);
    rowNumber = rowAttrs.r ? Number(rowAttrs.r) : rowNumber + 1;
    const cells = [];
    let col = -1;
    for (const cellMatch of (rowMatch[2] || '').matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)) {
      const attrs = xmlAttributes(cellMatch[1]);
      const index = columnIndex(attrs.r);
      col = index >= 0 ? index : col + 1;
      cells[col] = cellValue(attrs, cellMatch[2] || '', sharedStrings);
    }
    rows[rowNumber - 1] = Array.from(cells, value => value ?? '');
  }
  return Array.from(rows, row => row ?? []);
}

function resolveSheetPath(target) {
  const path = String(target || '').replace(/^\//, '');
  return path.startsWith('xl/') ? path : `xl/${path}`;
}

// Returns [{ name, rows }] for every worksheet, where rows is an array of string arrays.
async function readXlsx(input) {
  const files = await unzipEntries(input, name =>
    name === 'xl/workbook.xml' || name === 'xl/_rels/workbook.xml.rels' || name === 'xl/sharedStrings.xml' || name.startsWith('xl/worksheets/'));

  const workbook = files.get('xl/workbook.xml');
  if (!workbook) throw new Error('Not an Excel workbook.');

  const sharedStrings = [...(files.get('xl/sharedStrings.xml') || '')
    .matchAll(/<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g)]
    .map(match => xmlText(match[1] || ''));

  const targets = {};
  for (const [, attrs] of (files.get('xl/_rels/workbook.xml.rels') || '').matchAll(/<(?:\w+:)?Relationship\b([^>]*?)\/?>/g)) {
    const { Id, Target } = xmlAttributes(attrs);
    if (Id) targets[Id] = Target;
  }

  const sheets = [];
  for (const [, attrSource] of workbook.matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)) {
    const attrs = xmlAttributes(attrSource);
    const relId = Object.entries(attrs).find(([key]) => /(^|:)id$/.test(key))?.[1];
    const xml = files.get(resolveSheetPath(targets[relId]));
    if (xml) sheets.push({ name: attrs.name || `Sheet${sheets.length + 1}`, rows: parseSheetRows(xml, sharedStrings) });
  }
  return sheets;
}
