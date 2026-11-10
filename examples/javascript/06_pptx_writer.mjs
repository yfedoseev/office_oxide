#!/usr/bin/env node
/**
 * 06_pptx_writer: write a presentation with a size, slides, titles, text and an image, then read it back.
 *
 * Run: node examples/javascript/06_pptx_writer.mjs
 */
import { Document, PptxWriter } from '../../js/lib/index.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

function buildPptxBytes() {
  const w = new PptxWriter();
  try {
    w.setPresentationSize(12192000, 6858000); // 16:9, in EMU (914400 per inch)
    const one = w.addSlide();
    w.setSlideTitle(one, 'Welcome');
    w.addSlideText(one, 'Quarterly Report');
    const two = w.addSlide();
    w.setSlideTitle(two, 'Results');
    w.addSlideText(two, 'Revenue grew in every region.');
    w.addSlideImage(two, PNG, 'png', 914400, 2743200, 1828800, 1828800);
    return Buffer.from(w.toBytes());
  } finally {
    w.close();
  }
}

function check(condition, message) {
  if (!condition) {
    console.error(`FAIL: ${message}`);
    process.exit(1);
  }
}

function has(text, expected) {
  check(text.includes(expected), `"${expected}" is missing from the text`);
}

function main() {
  console.log('step 1: build a PPTX in memory');
  const data = buildPptxBytes();

  console.log('step 2: open it from bytes');
  const doc = Document.fromBytes(data, 'pptx');
  let text, slides, withImages;
  try {
    text = doc.plainText();
    slides = doc.toIr().sections.length;
    // Plain Markdown drops images; the WithImages variant keeps them in place.
    withImages = doc.toMarkdownWithImages();
    console.log(`  format: ${doc.format}`);
    console.log(`  chars: ${text.length}`);
    console.log(`  slides: ${slides}`);
  } finally {
    doc.close();
  }

  console.log('step 3: check the slides');
  check(slides === 2, 'expected two slides');
  for (const expected of ['Welcome', 'Quarterly Report', 'Results', 'Revenue grew in every region.']) {
    has(text, expected);
  }
  check(text.indexOf('Welcome') < text.indexOf('Results'), 'the slides are out of order');
  check(withImages.includes('[image-base64:'), 'the slide image is missing');

  console.log('OK 06_pptx_writer');
}

main();
