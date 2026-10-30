#!/usr/bin/env node
/**
 * 09_concurrency: open and read 32 documents from 8 workers at once, with logging on.
 *
 * Run: node examples/javascript/09_concurrency.mjs
 */
import { fileURLToPath } from 'node:url';
import { isMainThread, parentPort, Worker, workerData } from 'node:worker_threads';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document, createFromMarkdown } from '../../js/lib/index.js';

const WORKERS = 8;
const DOCUMENTS = 32;

const MARKDOWN = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`;

// createFromMarkdown writes to a path, so the file lives in a private temp
// directory just long enough to read its bytes back.
function markdownToBytes(markdown, format) {
  const dir = mkdtempSync(join(tmpdir(), 'oo_example_'));
  try {
    const path = join(dir, `report.${format}`);
    createFromMarkdown(markdown, format, path);
    return readFileSync(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function buildDocxBytes(title = 'Quarterly Report') {
  return markdownToBytes(MARKDOWN.replaceAll('Quarterly Report', title), 'docx');
}

function check(condition, message) {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    process.exit(1);
  }
}

// Each worker thread loads the binding itself and parses its share of the
// documents. Real threads, not Promise.all: a native call blocks its thread,
// so only separate threads exercise the parser's helper threads in parallel.
function parseInWorker() {
  const texts = workerData.items.map(([i, data]) => {
    const doc = Document.fromBytes(Buffer.from(data), 'docx');
    try {
      return [i, doc.plainText()];
    } finally {
      doc.close();
    }
  });
  parentPort.postMessage(texts);
}

async function main() {
  console.log('step 1: build 32 DOCX files in memory');
  const inputs = Array.from({ length: DOCUMENTS }, (_, i) => buildDocxBytes(`Quarterly Report ${i}`));
  console.log(`  documents: ${DOCUMENTS}`);

  console.log('step 2: skipped, the JS binding has no host logger bridge');

  console.log('step 3: parse them on 8 workers');
  // A watchdog turns a deadlock into a failure instead of a hang.
  const watchdog = setTimeout(() => check(false, 'a parse did not finish: possible deadlock'), 60_000);
  const jobs = Array.from({ length: WORKERS }, (_, w) => {
    const items = inputs.flatMap((data, i) => (i % WORKERS === w ? [[i, data]] : []));
    return new Promise((resolve, reject) => {
      const worker = new Worker(fileURLToPath(import.meta.url), { workerData: { items } });
      worker.once('message', resolve);
      worker.once('error', reject);
    });
  });
  const texts = [];
  for (const [i, text] of (await Promise.all(jobs)).flat()) texts[i] = text;
  clearTimeout(watchdog);
  console.log(`  workers: ${WORKERS}`);

  console.log('step 4: check every result');
  check(texts.length === DOCUMENTS && texts.every((t) => t), 'not every document was read');
  texts.forEach((text, i) => check(text.includes(`Quarterly Report ${i}`), `document ${i} came back with the wrong text`));

  console.log('OK 09_concurrency');
}

if (isMainThread) await main();
else parseInWorker();
