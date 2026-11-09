import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';
import { makeXlsx, zip } from './xlsx-fixture.mjs';

function loadXlsx() {
  const context = { Blob, DecompressionStream, Response, TextDecoder };
  vm.createContext(context);
  vm.runInContext(`${fs.readFileSync(new URL('../xlsx.js', import.meta.url), 'utf8')}
globalThis.readXlsx = readXlsx;`, context);
  return context.readXlsx;
}

// Arrays built inside the vm have another realm's prototype; compare plain copies.
const plain = value => JSON.parse(JSON.stringify(value));

const workbookXml = sheets => `<workbook xmlns:r="r"><sheets>${sheets}</sheets></workbook>`;

test('readXlsx reads shared strings, numbers, and sheet names', async () => {
  const readXlsx = loadXlsx();
  const file = makeXlsx([
    { name: 'Roster', rows: [['First Name', 'Email', 'GPA'], ['Ada', 'ada@example.com', 3.2250000000000001], ['Zoë & Co', 'zoe@example.com', 4]] },
    { name: 'Notes', rows: [['Hello']] }
  ]);

  const sheets = await readXlsx(file);

  assert.deepEqual(plain(sheets.map(sheet => sheet.name)), ['Roster', 'Notes']);
  assert.deepEqual(plain(sheets[0].rows), [
    ['First Name', 'Email', 'GPA'],
    ['Ada', 'ada@example.com', '3.225'],
    ['Zoë & Co', 'zoe@example.com', '4']
  ]);
  assert.deepEqual(plain(sheets[1].rows), [['Hello']]);
});

test('readXlsx handles inline strings, booleans, rich text, phonetic hints, gaps, and line breaks', async () => {
  const readXlsx = loadXlsx();
  const sheet = `<worksheet><sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Inline &amp; &#x4E2D;</t></is></c><c r="C1" t="b"><v>1</v></c></row>
    <row r="3"><c r="B3" t="s"><v>0</v></c><c r="D3" t="str"><v>formula text</v></c></row>
  </sheetData></worksheet>`;
  const file = zip({
    'xl/workbook.xml': workbookXml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><r><t>Line one</t></r><r><t xml:space="preserve">\r\nline two</t></r><rPh><t>ignored</t></rPh></si></sst>',
    'xl/worksheets/sheet1.xml': sheet
  });

  const [result] = await readXlsx(file);

  assert.deepEqual(plain(result.rows), [
    ['Inline & 中', '', 'TRUE'],
    [],
    ['', 'Line one\nline two', '', 'formula text']
  ]);
});

test('readXlsx reads uncompressed (stored) entries', async () => {
  const readXlsx = loadXlsx();
  const file = makeXlsx([{ name: 'Sheet1', rows: [['Email'], ['a@example.com']] }], { store: true });

  const [sheet] = await readXlsx(file);

  assert.deepEqual(plain(sheet.rows), [['Email'], ['a@example.com']]);
});

test('readXlsx rejects files that are not Excel workbooks', async () => {
  const readXlsx = loadXlsx();

  await assert.rejects(() => readXlsx(new TextEncoder().encode('Name,Email\nAda,ada@example.com')), /Not a zip file/);
  await assert.rejects(() => readXlsx(zip({ 'word/document.xml': '<doc/>' })), /Not an Excel workbook/);
});

test('readXlsx skips empty <t/> runs and decodes _xHHHH_ escapes once', async () => {
  const readXlsx = loadXlsx();
  const file = zip({
    'xl/workbook.xml': workbookXml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><r><t/></r><r><t>Ada</t></r></si><si><t>Line_x000D_break</t></si><si><t>_x005F_x0041_</t></si></sst>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="inlineStr"><is><t/></is></c><c r="E1"><v/></c><c r="F1"><v>7</v></c></row></sheetData></worksheet>'
  });

  const [sheet] = await readXlsx(file);

  assert.deepEqual(plain(sheet.rows), [['Ada', 'Line\rbreak', '_x0041_', '', '', '7']]);
});
