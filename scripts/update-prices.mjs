#!/usr/bin/env node
// Fetches current prices from the Coupang Partners product-search API and
// writes them to data/prices.json for the site's price ticker to read.
// Requires COUPANG_ACCESS_KEY / COUPANG_SECRET_KEY env vars (see .github/workflows/update-prices.yml).
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONFIG_PATH = path.join(__dirname, 'products.config.json');
const OUTPUT_PATH = path.join(__dirname, '..', 'data', 'prices.json');

const ACCESS_KEY = process.env.COUPANG_ACCESS_KEY;
const SECRET_KEY = process.env.COUPANG_SECRET_KEY;

if (!ACCESS_KEY || !SECRET_KEY) {
  console.error('Missing COUPANG_ACCESS_KEY / COUPANG_SECRET_KEY environment variables.');
  process.exit(1);
}

const HOST = 'https://api-gateway.coupang.com';
const API_PATH = '/v2/providers/affiliate_open_api/apis/openapi/products/search';

function signedDate() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const yy = pad(now.getUTCFullYear() % 100);
  const MM = pad(now.getUTCMonth() + 1);
  const dd = pad(now.getUTCDate());
  const HH = pad(now.getUTCHours());
  const mm = pad(now.getUTCMinutes());
  const ss = pad(now.getUTCSeconds());
  return `${yy}${MM}${dd}T${HH}${mm}${ss}Z`;
}

function sign(method, apiPath, query) {
  const datetime = signedDate();
  const message = `${datetime}${method}${apiPath}${query}`;
  const signature = crypto.createHmac('sha256', SECRET_KEY).update(message).digest('hex');
  return { datetime, signature };
}

async function searchProduct(keyword) {
  const query = `keyword=${encodeURIComponent(keyword)}&limit=1`;
  const { datetime, signature } = sign('GET', API_PATH, query);
  const authorization = `CEA algorithm=HmacSHA256, access-key=${ACCESS_KEY}, signed-date=${datetime}, signature=${signature}`;

  const res = await fetch(`${HOST}${API_PATH}?${query}`, {
    headers: {
      Authorization: authorization,
      'Content-Type': 'application/json;charset=UTF-8',
    },
  });

  if (!res.ok) {
    throw new Error(`Coupang API error ${res.status}: ${await res.text()}`);
  }

  const json = await res.json();
  const product = json?.data?.productData?.[0];
  if (!product) return null;

  return {
    price: product.productPrice,
    url: product.productUrl,
  };
}

async function main() {
  const config = JSON.parse(await fs.readFile(CONFIG_PATH, 'utf8'));

  let previous = { items: [] };
  try {
    previous = JSON.parse(await fs.readFile(OUTPUT_PATH, 'utf8'));
  } catch {
    // no previous data yet, that's fine
  }
  const prevByCategory = Object.fromEntries((previous.items || []).map((i) => [i.category, i]));

  const items = [];
  for (const entry of config) {
    const prev = prevByCategory[entry.category];
    try {
      const result = await searchProduct(entry.keyword);
      if (!result) {
        console.warn(`No result for "${entry.keyword}" — keeping previous value if any.`);
        if (prev) items.push(prev);
        continue;
      }
      let trend = 'same';
      if (prev) {
        if (result.price > prev.price) trend = 'up';
        else if (result.price < prev.price) trend = 'down';
      }
      items.push({
        category: entry.category,
        price: result.price,
        url: entry.url || result.url,
        trend,
      });
    } catch (err) {
      console.error(`Failed to fetch "${entry.keyword}":`, err.message);
      if (prev) items.push(prev);
    }
    // Coupang Partners API allows up to 10 calls/hour — space calls out.
    await new Promise((r) => setTimeout(r, 1500));
  }

  const output = {
    updatedAt: new Date().toISOString(),
    items,
  };

  await fs.writeFile(OUTPUT_PATH, JSON.stringify(output, null, 2) + '\n');
  console.log(`Wrote ${items.length} item(s) to ${OUTPUT_PATH}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
