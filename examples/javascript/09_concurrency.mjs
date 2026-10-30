#!/usr/bin/env node
/**
 * 09_concurrency — open and read 32 documents from 8 worker threads at once.
 *
 * Run: OFFICE_OXIDE_LIB=.../liboffice_oxide.so node examples/javascript/09_concurrency.mjs
 */
import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'os';
import { join } from 'path';
import { readFileSync, unlinkSync } from 'fs';
import { Document, createFromMarkdown } from '../../js/lib/index.js';

function check(cond, message) {
  if (!cond) { console.error(`FAIL: ${message}`); process.exit(1); }
}

const MARKDOWN = (title) => `# ${title}

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`;

// The JS binding has no DOCX writer, so the DOCX is generated into a temp
// file that is read and deleted at once; everything after that is bytes.
function buildDocxBytes(title = 'Quarterly Report') {
  const path = join(tmpdir(), `oo_${process.pid}_${Math.random().toString(36).slice(2)}.docx`);
  try {
    createFromMarkdown(MARKDOWN(title), 'docx', path);
    return readFileSync(path);
  } finally {
    try { unlinkSync(path); } catch (_) {}
  }
}


const WORKERS = 8;
const DOCUMENTS = 32;

if (isMainThread) {
  // Real worker_threads (koffi loads in workers), not Promise.all: a native
  // call blocks its thread, so only separate threads exercise the parser's
  // helper threads in parallel. The JS binding has no host logger bridge, so
  // there is no logging to switch on; the aim is the same, to show parsing
  // from many threads cannot deadlock the host. A watchdog turns a hang into
  // a failure.
  console.log(`step 1: start ${WORKERS} workers for ${DOCUMENTS} documents`);
  const watchdog = setTimeout(() => {
    console.error('FAIL: workers finish in time (deadlock?)');
    process.exit(1);
  }, 60_000);
  const url = fileURLToPath(import.meta.url);
  const jobs = Array.from({ length: WORKERS }, (_, w) => {
    const ids = [];
    for (let i = w; i < DOCUMENTS; i += WORKERS) ids.push(i);
    return new Promise((resolve, reject) => {
      const worker = new Worker(url, { workerData: { ids } });
      worker.once('message', resolve);
      worker.once('error', reject);
      worker.once('exit', (c) => c !== 0 && reject(new Error(`worker exited ${c}`)));
    });
  });

  console.log('step 2: open and read every document');
  const results = (await Promise.all(jobs)).flat();
  clearTimeout(watchdog);

  console.log('step 3: check every result');
  console.log(`  workers: ${WORKERS}`);
  console.log(`  documents: ${results.length}`);
  check(results.length === DOCUMENTS, `${DOCUMENTS} results`);
  for (const r of results) check(r.ok, `text contains "Quarterly Report ${r.i}"`);
  console.log('OK 09_concurrency');
} else {
  const out = [];
  for (const i of workerData.ids) {
    const doc = Document.fromBytes(buildDocxBytes(`Quarterly Report ${i}`), 'docx');
    try {
      out.push({ i, ok: doc.plainText().includes(`Quarterly Report ${i}`) });
    } finally { doc.close(); }
  }
  parentPort.postMessage(out);
}
