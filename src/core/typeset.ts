/**
 * Typesetting a manuscript into a print-ready book interior: justified text, first-line indents, chapter
 * openings, scene breaks, running heads, page numbers inside the margins, widow and orphan control, and a
 * gutter that follows KDP's table for the final page count.
 */

import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, rgb, type PDFFont } from 'pdf-lib';
import { gutterMargin } from './interior';
import type { Block, Manuscript, Run } from './manuscript';
import { ABSOLUTE_MAX_PAGES, MIN_PAGES } from './specs';

export type BookFont = 'garamond' | 'crimson' | 'baskerville';

export const BOOK_FONTS: readonly { id: BookFont; label: string }[] = [
  { id: 'garamond', label: 'EB Garamond (classic)' },
  { id: 'crimson', label: 'Crimson Pro (modern book)' },
  { id: 'baskerville', label: 'Libre Baskerville (large, readable)' },
];

export interface BookFontFiles {
  regular: Uint8Array;
  italic: Uint8Array;
  bold: Uint8Array;
  boldItalic: Uint8Array;
}

export interface BookStyle {
  font: BookFont;
  /** Body size in points. */
  size: number;
  /** Line height as a multiple of the size. */
  leading: number;
  /** First-line indent, inches. */
  indent: number;
  justify: boolean;
  /** Start every chapter on a right-hand page. */
  chapterRecto: boolean;
  /** Running heads: author on the left page, title (or chapter) on the right. */
  headers: 'title' | 'chapter' | 'none';
  pageNumbers: boolean;
  /** Add a title page and copyright page in front of the text. */
  frontMatter: boolean;
  title: string;
  subtitle: string;
  author: string;
  copyright: string;
  dedication: string;
}

export const DEFAULT_BOOK_STYLE: BookStyle = {
  font: 'garamond',
  size: 11.5,
  leading: 1.4,
  indent: 0.3,
  justify: true,
  chapterRecto: false,
  headers: 'title',
  pageNumbers: true,
  frontMatter: true,
  title: '',
  subtitle: '',
  author: '',
  copyright: '',
  dedication: '',
};

type F = 'r' | 'i' | 'b' | 'bi';
interface Seg {
  text: string;
  f: F;
}
interface Item extends Seg {
  /** Points from the text block's left edge. */
  x: number;
  /** Baseline, points from the page top. */
  y: number;
  size: number;
}
interface Page {
  kind: 'front' | 'body' | 'blank';
  items: Item[];
  chapterStart: boolean;
  chapter: string;
}

export interface TypesetResult {
  pdf: Uint8Array;
  pageCount: number;
  /** Physical page number where each chapter starts. */
  chapterPages: number[];
  margins: { inside: number; outside: number; top: number; bottom: number };
  warnings: string[];
}

const PT = 72;
const clampPages = (n: number) => Math.min(ABSOLUTE_MAX_PAGES, Math.max(MIN_PAGES, n));
/** Inside margin: KDP's gutter for the page count plus a comfortable reading margin. */
export const insideMargin = (pages: number) => gutterMargin(clampPages(pages)) + 0.25;
export const NOVEL_MARGINS = { outside: 0.5, top: 0.6, bottom: 0.6 } as const;

const fontOf = (r: Run): F => (r.b && r.i ? 'bi' : r.b ? 'b' : r.i ? 'i' : 'r');

/** Typeset a manuscript into a PDF at the trim size (no bleed). */
export async function typesetBook(ms: Manuscript, style: BookStyle, trim: { w: number; h: number }, files: BookFontFiles): Promise<TypesetResult> {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  doc.setTitle(style.title || 'Book interior');
  if (style.author) doc.setAuthor(style.author);
  doc.setProducer('SelfPublisher Studio');
  doc.setCreator('SelfPublisher Studio');
  const fonts: Record<F, PDFFont> = {
    r: await doc.embedFont(files.regular),
    i: await doc.embedFont(files.italic),
    b: await doc.embedFont(files.bold),
    bi: await doc.embedFont(files.boldItalic),
  };
  const charset = new Set(fonts.r.getCharacterSet());
  const missing = new Set<string>();
  const clean = (s: string) =>
    [...s.replace(/[ -  ]/g, ' ').replace(/[​­﻿]/g, '')]
      .map((ch) => {
        if (charset.has(ch.codePointAt(0)!) || ch === '\n') return ch;
        missing.add(ch);
        return '?';
      })
      .join('');
  // Measuring through fontkit is slow; a novel repeats the same words, so cache widths at 1 pt.
  const widthCache = new Map<string, number>();
  const width = (s: Seg, size: number) => {
    const key = s.f + s.text;
    let w = widthCache.get(key);
    if (w === undefined) {
      w = fonts[s.f].widthOfTextAtSize(s.text, 1);
      widthCache.set(key, w);
    }
    return w * size;
  };

  const W = trim.w * PT;
  const H = trim.h * PT;
  const size = style.size;
  const lineH = size * style.leading;

  function layout(inside: number): { pages: Page[]; chapterPages: number[] } {
    const textW = W - (inside + NOVEL_MARGINS.outside) * PT;
    const headRow = style.headers !== 'none' ? lineH * 1.6 : 0;
    const footRow = style.pageNumbers ? lineH * 1.6 : 0;
    const top = NOVEL_MARGINS.top * PT + headRow;
    const bottom = H - NOVEL_MARGINS.bottom * PT - footRow;
    const pages: Page[] = [];
    const chapterPages: number[] = [];
    let page: Page = { kind: "body", items: [], chapterStart: false, chapter: "" };
    let y = 0; // baseline of the last line placed
    let chapter = '';

    const newPage = (kind: Page['kind'] = 'body', chapterStart = false) => {
      page = { kind, items: [], chapterStart, chapter };
      pages.push(page);
      y = top - lineH + size; // so the first line's baseline sits at top + size
    };
    const fits = () => y + lineH <= bottom + 0.01;

    /** Break runs into lines of segments. */
    const breakLines = (runs: Run[], sz: number, firstIndent: number, maxW: number, forceFont?: F): { segs: Seg[]; w: number; gaps: number; last: boolean }[][] => {
      // Words are lists of segments (a word can change style mid-way). `\n` is a forced line break.
      const words: (Seg[] | 'br')[] = [];
      let cur: Seg[] = [];
      const flush = () => {
        if (cur.length) words.push(cur);
        cur = [];
      };
      for (const r of runs) {
        const f = forceFont ?? fontOf(r);
        const parts = clean(r.text).split(/(\s)/);
        for (const part of parts) {
          if (part === '\n') {
            flush();
            words.push('br');
          } else if (/^\s$/.test(part)) flush();
          else if (part) cur.push({ text: part, f });
        }
      }
      flush();
      const space = width({ text: ' ', f: 'r' }, sz);
      const lines: { segs: Seg[]; w: number; gaps: number; last: boolean }[] = [];
      let line: Seg[][] = [];
      let lw = 0;
      const ww = (w: Seg[]) => w.reduce((a, s) => a + width(s, sz), 0);
      const avail = () => maxW - (lines.length === 0 ? firstIndent : 0);
      const emit = (last: boolean) => {
        const segs: Seg[] = [];
        line.forEach((w, k) => {
          if (k) segs.push({ text: ' ', f: 'r' });
          segs.push(...w);
        });
        lines.push({ segs, w: lw, gaps: Math.max(0, line.length - 1), last });
        line = [];
        lw = 0;
      };
      for (const w of words) {
        if (w === 'br') {
          emit(true);
          continue;
        }
        const add = ww(w) + (line.length ? space : 0);
        if (line.length && lw + add > avail()) emit(false);
        lw += ww(w) + (line.length ? space : 0);
        line.push(w);
      }
      if (line.length || !lines.length) emit(true);
      else lines[lines.length - 1]!.last = true;
      return [lines];
    };

    /** Place prepared lines with widow/orphan control. */
    const placeLines = (lines: { segs: Seg[]; w: number; gaps: number; last: boolean }[], sz: number, align: 'justify' | 'left' | 'center' | 'right', firstIndent: number, maxW: number, x0 = 0) => {
      let k = 0;
      while (k < lines.length) {
        let fit = 0;
        for (let yy = y; yy + lineH <= bottom + 0.01 && fit < lines.length - k; yy += lineH) fit++;
        const left = lines.length - k;
        if (fit === 0 || (fit === 1 && left > 1 && k === 0 && page.items.length)) {
          newPage();
          continue;
        }
        let take = Math.min(fit, left);
        if (left - take === 1 && take >= 3) take--; // keep the paragraph's last line company
        for (let j = 0; j < take; j++, k++) {
          const ln = lines[k]!;
          y += lineH;
          const indent = k === 0 ? firstIndent : 0;
          const room = maxW - indent;
          let x = x0 + indent;
          let extra = 0;
          if (align === 'center') x = x0 + (maxW - ln.w) / 2;
          else if (align === 'right') x = x0 + maxW - ln.w;
          else if (align === 'justify' && !ln.last && ln.gaps) extra = (room - ln.w) / ln.gaps;
          for (const s of ln.segs) {
            if (s.text === ' ') {
              x += width({ text: ' ', f: 'r' }, sz) + extra;
              continue;
            }
            page.items.push({ ...s, x, y, size: sz });
            x += width(s, sz);
          }
        }
        if (k < lines.length) newPage();
      }
    };

    const para = (runs: Run[], o: { indent: boolean; align?: 'center' | 'right'; sz?: number; font?: F }) => {
      const sz = o.sz ?? size;
      const align = o.align ?? (style.justify ? 'justify' : 'left');
      const ind = o.indent && !o.align ? style.indent * PT : 0;
      const [lines] = breakLines(runs, sz, ind, textW, o.font);
      placeLines(lines!, sz, align, ind, textW);
    };

    const centered = (text: string, sz: number, f: F, at?: number) => {
      if (at !== undefined) y = at - lineH;
      const [lines] = breakLines([{ text, ...(f.includes('i') ? { i: true } : {}), ...(f.includes('b') ? { b: true } : {}) }], sz, 0, textW);
      for (const ln of lines!) {
        y += sz * 1.25;
        let x = (textW - ln.w) / 2;
        for (const s of ln.segs) {
          if (s.text !== ' ') page.items.push({ ...s, x, y, size: sz });
          x += s.text === ' ' ? width({ text: ' ', f: 'r' }, sz) : width(s, sz);
        }
      }
    };

    const blocks = (bs: Block[], isFront: boolean) => {
      let first = true;
      for (const b of bs) {
        if (b.kind === 'pagebreak') {
          newPage(isFront ? 'front' : 'body');
          first = true;
          continue;
        }
        if (b.kind === 'scene') {
          if (!fits()) newPage();
          y += lineH * 0.5;
          if (!fits()) newPage();
          centered('*   *   *', size, 'r');
          y += lineH * 0.5;
          first = true;
          continue;
        }
        if (b.heading) {
          if (page.items.length) y += lineH * 0.6;
          if (y + lineH * 3 > bottom) newPage();
          para(b.runs, { indent: false, align: 'center', sz: size * (b.heading === 2 ? 1.15 : 1.05), font: 'b' });
          y += lineH * 0.3;
          first = true;
          continue;
        }
        para(b.runs, { indent: !first && !isFront, align: b.align });
        first = false;
      }
    };

    // --- front matter ---
    const textH = bottom - top;
    if (style.frontMatter) {
      newPage('front');
      if (style.title) centered(style.title, size * 2.2, 'r', top + textH * 0.3);
      if (style.subtitle) {
        y += size;
        centered(style.subtitle, size * 1.2, 'i');
      }
      if (style.author) centered(style.author, size * 1.3, 'r', top + textH * 0.62);
      newPage('front');
      if (style.copyright.trim()) {
        const lines = style.copyright.split('\n');
        const est = lines.length * size * 0.85 * 1.6 + 2 * lineH;
        y = bottom - est;
        for (const l of lines) para([{ text: l }], { indent: false, align: 'center', sz: size * 0.82 });
      }
      if (style.dedication.trim()) {
        newPage('front');
        y = top + textH * 0.25;
        for (const l of style.dedication.split('\n')) para([{ text: l, i: true }], { indent: false, align: 'center' });
        newPage('blank');
      }
    }
    if (ms.front.length) {
      newPage('front');
      blocks(ms.front, true);
    }

    // --- chapters ---
    for (const ch of ms.chapters) {
      chapter = ch.title;
      if (!pages.length) newPage('body', true);
      else {
        newPage('body', true);
        if (style.chapterRecto && pages.length % 2 === 0) {
          // Landed on a left-hand page: leave it blank and start on the next.
          page.kind = 'blank';
          page.chapterStart = false;
          newPage('body', true);
        }
      }
      chapterPages.push(pages.length);
      y = top + textH * 0.22 - lineH;
      if (ch.title) {
        centered(ch.title, size * 1.7, 'r');
        y += lineH * 2;
      } else y += lineH;
      blocks(ch.blocks, false);
    }
    // Drop a trailing empty page.
    while (pages.length > 1 && !pages[pages.length - 1]!.items.length) pages.pop();
    return { pages, chapterPages };
  }

  // Lay out, then again if the page count moved the gutter into another band.
  let inside = insideMargin(300);
  let result = layout(inside);
  for (let n = 0; n < 3; n++) {
    const want = insideMargin(result.pages.length);
    if (Math.abs(want - inside) < 1e-6) break;
    inside = want;
    result = layout(inside);
  }

  // --- draw ---
  const ink = rgb(0.08, 0.08, 0.08);
  const grey = rgb(0.35, 0.35, 0.35);
  const headSize = size * 0.78;
  result.pages.forEach((p, idx) => {
    const n = idx + 1;
    const recto = n % 2 === 1;
    const left = (recto ? inside : NOVEL_MARGINS.outside) * PT;
    const textW = W - (inside + NOVEL_MARGINS.outside) * PT;
    const pg = doc.addPage([W, H]);
    for (const it of p.items) pg.drawText(it.text, { x: left + it.x, y: H - it.y, size: it.size, font: fonts[it.f], color: ink });
    if (p.kind !== 'body') return;
    if (style.headers !== 'none' && !p.chapterStart) {
      const text = clean((recto ? (style.headers === 'chapter' && p.chapter ? p.chapter : style.title) : style.author || style.title).toUpperCase());
      if (text) {
        const w = fonts.r.widthOfTextAtSize(text, headSize);
        pg.drawText(text, { x: left + (textW - w) / 2, y: H - (NOVEL_MARGINS.top * PT + headSize), size: headSize, font: fonts.r, color: grey });
      }
    }
    if (style.pageNumbers) {
      const t = String(n);
      const w = fonts.r.widthOfTextAtSize(t, headSize * 1.05);
      pg.drawText(t, { x: left + (textW - w) / 2, y: NOVEL_MARGINS.bottom * PT + headSize * 0.3, size: headSize * 1.05, font: fonts.r, color: grey });
    }
  });

  const warnings: string[] = [];
  if (missing.size) warnings.push(`Some characters are not in this font and print as "?": ${[...missing].slice(0, 12).join(' ')}. Try another font or replace them in the manuscript.`);
  if (ms.stats.images) warnings.push(`${ms.stats.images} image${ms.stats.images > 1 ? 's were' : ' was'} left out. Pictures in the text are not supported yet.`);
  if (ms.stats.tables) warnings.push(`${ms.stats.tables} table${ms.stats.tables > 1 ? 's were' : ' was'} included as plain paragraphs.`);
  if (ms.stats.footnotes) warnings.push(`${ms.stats.footnotes} footnote${ms.stats.footnotes > 1 ? 's were' : ' was'} left out.`);
  return { pdf: await doc.save(), pageCount: result.pages.length, chapterPages: result.chapterPages, margins: { inside, ...NOVEL_MARGINS }, warnings };
}
