#!/usr/bin/env node
// Syncs "신규제철"/"best" product data (name, price, Coupang link) from the
// shared Google Sheet into data/prices.json's seasonNew/seasonBest groups.
// Only owns those two groups — the price ticker is synced separately by
// scripts/update-prices.mjs from the Coupang Partners API.
//
// The sheet must be shared as "Anyone with the link: Viewer" so this script
// can fetch it anonymously via Google's XLSX export endpoint (no API key
// needed). Layout expected (see scripts/products.config.json for the id):
//   row 1: merged group headers ("신규제철" / "best")
//   row 2: column headers ("상품","가격","링크" x2)
//   row 3+: data — B/C/D = 신규제철 상품/가격/링크, E/F/G = best 상품/가격/링크
// The "링크" cell's displayed text is a Coupang product title; the actual
// product URL is read from the cell's hyperlink target, not its text.
import fs from 'node:fs/promises';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONFIG_PATH = path.join(__dirname, 'products.config.json');
const OUTPUT_PATH = path.join(__dirname, '..', 'data', 'prices.json');

const FALLBACK_URL = 'https://shop.coupang.com/A01623310';
const MAX_DATA_ROWS = 200;
const MAX_EMPTY_STREAK = 3;

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

function colLetterOf(ref) {
  return ref.match(/^[A-Z]+/)[0];
}

function parseSheet(xml, sharedStrings) {
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

  const hyperlinks = {};
  const linkRe = /<hyperlink\s+[^>]*r:id="(rId\d+)"[^>]*ref="([A-Z]+\d+)"[^>]*\/>/g;
  let linkMatch;
  while ((linkMatch = linkRe.exec(xml))) {
    hyperlinks[linkMatch[2]] = linkMatch[1];
  }

  return { cells, hyperlinks };
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

function extractGroup(cells, hyperlinks, rels, nameCol, priceCol, linkCol) {
  const items = [];
  let emptyStreak = 0;
  for (let row = 3; row < 3 + MAX_DATA_ROWS && emptyStreak < MAX_EMPTY_STREAK; row++) {
    const name = cells[`${nameCol}${row}`];
    const price = cells[`${priceCol}${row}`];
    if (name === undefined || price === undefined) {
      emptyStreak++;
      continue;
    }
    emptyStreak = 0;
    const linkRef = `${linkCol}${row}`;
    const rId = hyperlinks[linkRef];
    const url = (rId && rels[rId]) || FALLBACK_URL;
    items.push({ category: String(name).trim(), price: Math.round(Number(price)), url });
  }
  return items;
}

async function main() {
  const config = JSON.parse(await fs.readFile(CONFIG_PATH, 'utf8'));
  const { spreadsheetId, gid } = config.seasonSheet;
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${gid || '0'}`;

  const res = await fetch(exportUrl);
  if (!res.ok) {
    throw new Error(
      `Failed to fetch season sheet (${res.status}). Make sure it's shared as "Anyone with the link: Viewer".`
    );
  }
  const buf = Buffer.from(await res.arrayBuffer());

  const entries = readZipEntries(buf, [
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/_rels/sheet1.xml.rels',
    'xl/sharedStrings.xml',
  ]);

  const sharedStrings = entries['xl/sharedStrings.xml']
    ? parseSharedStrings(entries['xl/sharedStrings.xml'].toString('utf8'))
    : [];
  const { cells, hyperlinks } = parseSheet(entries['xl/worksheets/sheet1.xml'].toString('utf8'), sharedStrings);
  const rels = entries['xl/worksheets/_rels/sheet1.xml.rels']
    ? parseRels(entries['xl/worksheets/_rels/sheet1.xml.rels'].toString('utf8'))
    : {};

  const seasonNew = extractGroup(cells, hyperlinks, rels, 'B', 'C', 'D');
  const seasonBest = extractGroup(cells, hyperlinks, rels, 'E', 'F', 'G');

  if (!seasonNew.length && !seasonBest.length) {
    throw new Error('Parsed 0 products from the season sheet — check the sheet layout.');
  }

  let previous = {};
  try {
    previous = JSON.parse(await fs.readFile(OUTPUT_PATH, 'utf8'));
  } catch {
    // no previous data yet, that's fine
  }

  const output = { ...previous, seasonNew, seasonBest, seasonUpdatedAt: new Date().toISOString() };
  await fs.writeFile(OUTPUT_PATH, JSON.stringify(output, null, 2) + '\n');
  console.log(`Wrote ${seasonNew.length} seasonNew + ${seasonBest.length} seasonBest item(s) to ${OUTPUT_PATH}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
