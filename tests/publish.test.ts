import { readFileSync } from 'node:fs';
import { PDFDocument } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { blockText, buildEpub, DEFAULT_BOOK_STYLE, readDocx, readZip, typesetBook, writeZip, type BookStyle } from '../src/core';

const F = (n: string) => readFileSync(`src/app/fonts/book/eb-garamond-${n}.ttf`);
const fonts = { regular: F('400-normal'), italic: F('400-italic'), bold: F('700-normal'), boldItalic: F('700-italic') };

const P = (text: string, o: { style?: string; i?: boolean; center?: boolean; pageBreak?: boolean } = {}) =>
  `<w:p><w:pPr>${o.style ? `<w:pStyle w:val="${o.style}"/>` : ''}${o.center ? '<w:jc w:val="center"/>' : ''}</w:pPr>${o.pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''}<w:r>${o.i ? '<w:rPr><w:i/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const LOREM = 'The rain had not stopped for three days, and the lantern in the window burned lower every night while the town waited for news that never came. ';

function docx(body: string): Uint8Array {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const styles = `<?xml version="1.0"?><w:styles xmlns:w="x"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style></w:styles>`;
  return writeZip([
    { name: '[Content_Types].xml', data: '<Types/>' },
    { name: 'word/document.xml', data: doc },
    { name: 'word/styles.xml', data: styles },
  ]);
}

const novel = () =>
  docx(
    [1, 2, 3]
      .map(
        (n) =>
          P(`Chapter ${n}`, { style: 'Heading1' }) +
          Array.from({ length: 30 }, (_, k) => (k === 12 ? P('* * *', { center: true }) : `<w:p><w:r><w:t xml:space="preserve">${LOREM.repeat(3)}</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>She knew.</w:t></w:r></w:p>`)).join(''),
      )
      .join(''),
  );

describe('docx', () => {
  it('reads chapters from Heading 1, keeps italics and scene breaks', async () => {
    const ms = await readDocx(novel());
    expect(ms.chapterRule).toBe('headings');
    expect(ms.chapters.map((c) => c.title)).toEqual(['Chapter 1', 'Chapter 2', 'Chapter 3']);
    const b = ms.chapters[0]!.blocks;
    expect(b.filter((x) => x.kind === 'scene')).toHaveLength(1);
    const p = b[0]!;
    expect(p.kind === 'para' && p.runs.some((r) => r.i && r.text === 'She knew.')).toBe(true);
    expect(ms.stats.words).toBeGreaterThan(5000);
  });

  it('finds chapters by their words when no heading styles are used, and keeps front matter', async () => {
    const ms = await readDocx(docx(P('My Title', { center: true }) + P('Prologue') + P(LOREM) + P('Chapter One', { pageBreak: true }) + P(LOREM) + P('Chapter Two') + P(LOREM)));
    expect(ms.chapterRule).toBe('chapter-words');
    expect(ms.chapters.map((c) => c.title)).toEqual(['Prologue', 'Chapter One', 'Chapter Two']);
    expect(blockText(ms.front)).toBe('My Title');
  });

  it('decodes entities and rejects non-docx files', async () => {
    const ms = await readDocx(docx(P('Tom &amp; Jerry &#8212; &quot;hi&quot;')));
    expect(blockText(ms.chapters[0]!.blocks)).toBe('Tom & Jerry — "hi"');
    await expect(readDocx(new Uint8Array([1, 2, 3]))).rejects.toThrow(/Word document/);
  });
});

describe('typesetting', () => {
  const style: BookStyle = { ...DEFAULT_BOOK_STYLE, title: 'The Last Lantern', author: 'June Hart', copyright: 'Copyright 2026 June Hart.\nAll rights reserved.', chapterRecto: true };

  it('makes a trim-size PDF with chapters on right-hand pages and a gutter for its page count', async () => {
    const ms = await readDocx(novel());
    const r = await typesetBook(ms, style, { w: 5.5, h: 8.5 }, fonts);
    const doc = await PDFDocument.load(r.pdf);
    expect(doc.getPageCount()).toBe(r.pageCount);
    expect(doc.getPage(0).getSize()).toEqual({ width: 5.5 * 72, height: 8.5 * 72 });
    expect(r.pageCount).toBeGreaterThan(20);
    for (const p of r.chapterPages) expect(p % 2).toBe(1);
    expect(r.margins.inside).toBeGreaterThanOrEqual(0.375 + 0.25);
    const raw = Buffer.from(await doc.save({ useObjectStreams: false })).toString('latin1');
    expect(raw).toContain('/FontFile2');
    expect(r.warnings).toEqual([]);
  });

  it('larger type means more pages', async () => {
    const ms = await readDocx(novel());
    const a = await typesetBook(ms, style, { w: 6, h: 9 }, fonts);
    const b = await typesetBook(ms, { ...style, size: 14 }, { w: 6, h: 9 }, fonts);
    expect(b.pageCount).toBeGreaterThan(a.pageCount);
  });
});

describe('epub', () => {
  it('stores mimetype first and lists every chapter', async () => {
    const ms = await readDocx(novel());
    const bytes = buildEpub(ms, { style: { ...DEFAULT_BOOK_STYLE, title: 'T', author: 'A', copyright: 'c' }, cover: new Uint8Array([0xff, 0xd8, 0xff]), modified: '2026-11-16T00:00:00Z', id: 'urn:isbn:9781736340127' });
    expect(new TextDecoder().decode(bytes.subarray(30, 38))).toBe('mimetype');
    const files = await readZip(bytes);
    expect(new TextDecoder().decode(files.get('mimetype'))).toBe('application/epub+zip');
    const opf = new TextDecoder().decode(files.get('OEBPS/content.opf'));
    expect(opf).toContain('properties="cover-image"');
    expect(opf).toContain('chapter-3.xhtml');
    expect(new TextDecoder().decode(files.get('OEBPS/chapter-1.xhtml'))).toContain('<em>She knew.</em>');
  });
});

describe('embedded fonts', () => {
  it('embeds real TrueType programs, not web fonts', async () => {
    const { inflateSync } = await import('node:zlib');
    const { PDFRawStream } = await import('pdf-lib');
    const ms = await readDocx(novel());
    const r = await typesetBook(ms, { ...DEFAULT_BOOK_STYLE, title: 'T' }, { w: 6, h: 9 }, fonts);
    const doc = await PDFDocument.load(r.pdf);
    const magics: string[] = [];
    for (const [, o] of doc.context.enumerateIndirectObjects()) {
      if (!(o instanceof PDFRawStream)) continue;
      let c: Uint8Array = o.contents;
      try {
        c = inflateSync(c);
      } catch {
        /* not compressed */
      }
      const m = Buffer.from(c.subarray(0, 4)).toString('hex');
      if (['774f4632', '774f4646', '00010000', '4f54544f'].includes(m)) magics.push(m);
    }
    expect(magics.length).toBe(4);
    expect(magics.every((m) => m === '00010000')).toBe(true);
  });
});
