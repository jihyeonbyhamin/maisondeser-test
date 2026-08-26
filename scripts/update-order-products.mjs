#!/usr/bin/env node
// Syncs the direct-order page's product list/prices from the "메종드세르" tab
// of the internal pricing sheet (see scripts/products.config.json for the id)
// into data/order-products.json, which order.html reads client-side.
//
// This is a working margin-calculation sheet (not a clean catalog export):
// each category ("키워드", column A) starts a block of a few rows, one row
// per weight/quantity option. Column I ("현재 판매가") is the customer price
// per the site owner's instruction. Column P ("매핑") usually holds a retailer
// product title we clean up into a customer-facing label; when it's missing
// we fall back to "<category> <옵션>". A category can reappear later in the
// sheet as a duplicate/promo block (e.g. "수박 단독특가") — only the FIRST
// block seen per category (scripts/products.config.json's "categories" list)
// is kept, matching what the site owner asked to expose.
//
// The sheet must be shared as "Anyone with the link: Viewer" so this script
// can fetch it anonymously via Google's XLSX export endpoint (no API key
// needed).
import fs from 'node:fs/promises';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONFIG_PATH = path.join(__dirname, 'products.config.json');
const OUTPUT_PATH = path.join(__dirname, '..', 'data', 'order-products.json');

const MAX_DATA_ROWS = 150;
const PRICE_COL = 'I';
const OPTION_COL = 'F';
const DESC_COL = 'P';
const CATEGORY_COL = 'A';
const ITEM_COL = 'B';

function readZipEntries(buf, names) {
  const eocdSig = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.subarray(i, i + 4).equals(eocdSig)) { eocd = i; break; }
  }
  if (eocd === -1) throw new Error('Not a valid xlsx (zip EOCD not found)');

  const entryCount = buf.readUInt16LE(eocd + 10);
  let cdOffset = buf.readUInt32LE(eocd + 16);
  const wanted = new Set(names);
  const found = {};

  for (let i = 0; i < entryCount; i++) {
    const sig = buf.readUInt32LE(cdOffset);
    if (sig !== 0x02014b50) throw new Error(`Bad central directory signature at ${cdOffset}`);
    const method = buf.readUInt16LE(cdOffset + 10);
    const compSize = buf.readUInt32LE(cdOffset + 20);
    const nameLen = buf.readUInt16LE(cdOffset + 28);
    const extraLen = buf.readUInt16LE(cdOffset + 30);
    const commentLen = buf.readUInt16LE(cdOffset + 32);
    const localOffset = buf.readUInt32LE(cdOffset + 42);
    const name = buf.toString('utf8', cdOffset + 46, cdOffset + 46 + nameLen);

    if (wanted.has(name)) {
      const localNameLen = buf.readUInt16LE(localOffset + 26);
      const localExtraLen = buf.readUInt16LE(localOffset + 28);
      const dataStart = localOffset + 30 + localNameLen + localExtraLen;
      const raw = buf.subarray(dataStart, dataStart + compSize);
      found[name] = method === 0 ? Buffer.from(raw) : zlib.inflateRawSync(raw);
    }

    cdOffset += 46 + nameLen + extraLen + commentLen;
  }
  return found;
}

function decodeXmlEntities(str) {
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, '&');
}

function parseSharedStrings(xml) {
  const strings = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = siRe.exec(xml))) {
    const text = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join('');
    strings.push(decodeXmlEntities(text));
  }
  return strings;
}

function parseSheetCells(xml, sharedStrings) {
  const cells = {};
  const rowRe = /<row [^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRe.exec(xml))) {
    const cellRe = /<c\s+([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cellMatch;
    while ((cellMatch = cellRe.exec(rowMatch[2]))) {
      const attrs = cellMatch[1];
      const refMatch = attrs.match(/r="([A-Z]+\d+)"/);
      if (!refMatch) continue;
      const ref = refMatch[1];
      const typeMatch = attrs.match(/t="(\w+)"/);
      const type = typeMatch ? typeMatch[1] : 'n';
      const inner = cellMatch[2] || '';
      const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
      if (!vMatch) continue;
      const raw = vMatch[1];
      cells[ref] = type === 's' ? sharedStrings[Number(raw)] : Number(raw);
    }
  }
  return cells;
}

function parseRels(xml) {
  const rels = {};
  const relRe = /<Relationship\s+[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"/g;
  let m;
  while ((m = relRe.exec(xml))) {
    rels[m[1]] = decodeXmlEntities(m[2]);
  }
  return rels;
}

function findSheetTarget(workbookXml, relsXml, sheetName) {
  const sheetRe = /<sheet\s+[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"[^>]*\/>/g;
  let m;
  let rId = null;
  while ((m = sheetRe.exec(workbookXml))) {
    if (decodeXmlEntities(m[1]) === sheetName) {
      rId = m[2];
      break;
    }
  }
  if (!rId) throw new Error(`Sheet named "${sheetName}" not found in workbook.`);
  const rels = parseRels(relsXml);
  const target = rels[rId];
  if (!target) throw new Error(`No relationship target for ${rId}.`);
  return `xl/${target.replace(/^\/?xl\//, '')}`;
}

const BRACKET_RE = /\[[^\]]*\]\s*/g;
const LEADING_PAREN_RE = /^\(([^()]{0,30})\)\s*/;
const STRAY_CLOSE_PAREN_RE = /^[^()]{1,20}\)\s*/; // source has typos like "신흥유통)참외..." (missing open paren)
const WRAPPED_TAG_RE = /^[★☆●○]+[^★☆●○]{0,12}[★☆●○]+\s*/; // e.g. "★특가★ ", "●가성비●"
const LEADING_BULLET_RE = /^[★☆●○■□▲▶]+\s*/;
const TRAILING_PRICE_RE = /\s*\|\s*[\d,]+\s*원\s*$/;
const ERROR_SENTINELS = new Set(['#VALUE!', '#REF!', '#DIV/0!', '#N/A']);

function cleanDescription(desc) {
  if (typeof desc !== 'string') return null;
  let str = desc.trim();
  if (!str || ERROR_SENTINELS.has(str)) return null;

  let prev;
  do {
    prev = str;
    str = str.replace(LEADING_PAREN_RE, '');
  } while (str !== prev);
  const stray = str.match(STRAY_CLOSE_PAREN_RE);
  if (stray && !stray[0].includes('(')) str = str.slice(stray[0].length);

  str = str
    .replace(BRACKET_RE, '')
    .replace(WRAPPED_TAG_RE, '')
    .replace(LEADING_BULLET_RE, '')
    .replace(TRAILING_PRICE_RE, '')
    .trim();
  return str || null;
}

function formatOption(optionText) {
  if (optionText === undefined || optionText === null || optionText === '') return '';
  let opt = String(optionText).trim().replace(/^\*+/, '');
  if (/^\d+$/.test(opt)) {
    const n = Number(opt);
    opt = n >= 100 ? `${n}g` : `${n}kg`;
  }
  return opt;
}

async function main() {
  const config = JSON.parse(await fs.readFile(CONFIG_PATH, 'utf8'));
  const { spreadsheetId, sheetName, categories } = config.orderProductsSheet;
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;

  const res = await fetch(exportUrl);
  if (!res.ok) {
    throw new Error(
      `Failed to fetch order products sheet (${res.status}). Make sure it's shared as "Anyone with the link: Viewer".`
    );
  }
  const buf = Buffer.from(await res.arrayBuffer());

  const prelim = readZipEntries(buf, ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']);
  const sheetPath = findSheetTarget(
    prelim['xl/workbook.xml'].toString('utf8'),
    prelim['xl/_rels/workbook.xml.rels'].toString('utf8'),
    sheetName
  );

  const entries = readZipEntries(buf, [sheetPath, 'xl/sharedStrings.xml']);
  const sharedStrings = entries['xl/sharedStrings.xml']
    ? parseSharedStrings(entries['xl/sharedStrings.xml'].toString('utf8'))
    : [];
  const cells = parseSheetCells(entries[sheetPath].toString('utf8'), sharedStrings);

  const allowed = new Set(categories);
  const used = new Set();
  let currentCategory = null;
  let itemName = null;
  let collecting = false;
  const products = [];

  for (let row = 3; row < 3 + MAX_DATA_ROWS; row++) {
    const rawCategory = cells[`${CATEGORY_COL}${row}`];
    if (typeof rawCategory === 'string' && rawCategory.trim()) {
      const category = rawCategory.trim();
      if (used.has(category)) {
        collecting = false; // duplicate/promo block reappearing later — skip it
      } else {
        used.add(category);
        currentCategory = category;
        const rawItem = cells[`${ITEM_COL}${row}`];
        itemName = typeof rawItem === 'string' && rawItem.trim() ? rawItem.trim() : category;
        collecting = allowed.has(category);
      }
    }
    if (!collecting || !currentCategory) continue;

    const price = cells[`${PRICE_COL}${row}`];
    if (typeof price !== 'number' || !Number.isFinite(price) || price <= 0) continue;

    const desc = cleanDescription(cells[`${DESC_COL}${row}`]);
    const optionText = formatOption(cells[`${OPTION_COL}${row}`]);
    const label = desc || [itemName, optionText].filter(Boolean).join(' ');
    if (!label) continue;

    products.push({ category: currentCategory, label, price: Math.round(price) });
  }

  if (!products.length) {
    throw new Error('Parsed 0 products from the order sheet — check the sheet layout or categories list.');
  }

  const output = { updatedAt: new Date().toISOString(), products };
  await fs.writeFile(OUTPUT_PATH, JSON.stringify(output, null, 2) + '\n');
  console.log(`Wrote ${products.length} order product option(s) to ${OUTPUT_PATH}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
