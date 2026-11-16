/**
 * A manuscript: the book's text as chapters of paragraphs with italic and bold, read from a Word .docx.
 * Word's own fonts, sizes and spacing are deliberately dropped; the typesetter applies the book's style.
 */

import { readZip } from './zip';

export interface Run {
  text: string;
  i?: boolean;
  b?: boolean;
}

export type Block =
  | { kind: 'para'; runs: Run[]; align?: 'center' | 'right'; heading?: 2 | 3 }
  | { kind: 'scene' }
  | { kind: 'pagebreak' };

export interface Chapter {
  /** Empty for text before the first chapter heading (front matter typed into the manuscript). */
  title: string;
  blocks: Block[];
}

export interface Manuscript {
  /** Content before the first chapter heading, if any. */
  front: Block[];
  chapters: Chapter[];
  stats: { words: number; paragraphs: number; images: number; tables: number; footnotes: number };
  /** How chapters were found, for the UI. */
  chapterRule: 'headings' | 'chapter-words' | 'page-breaks' | 'none';
}

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const decode = (s: string) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e: string) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : (ENT[e] ?? m),
  );

const attr = (tag: string, name: string) => {
  const m = new RegExp(`${name}="([^"]*)"`).exec(tag);
  return m ? decode(m[1]!) : null;
};
const on = (tag: string) => {
  const v = attr(tag, 'w:val');
  return v === null || !['0', 'false', 'none'].includes(v);
};

/** styleId -> lower-case style name, from styles.xml. */
function styleNames(xml: string | undefined): Map<string, string> {
  const m = new Map<string, string>();
  if (!xml) return m;
  for (const s of xml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    const name = /<w:name w:val="([^"]+)"/.exec(s[2]!)?.[1];
    if (name) m.set(s[1]!, name.toLowerCase());
  }
  return m;
}

export interface RawPara {
  runs: Run[];
  style: string;
  align?: 'center' | 'right';
  breakBefore: boolean;
  breakAfter: boolean;
}

/** Paragraphs from document.xml, in order. */
function paragraphs(xml: string, styles: Map<string, string>, stats: Manuscript['stats']): RawPara[] {
  const out: RawPara[] = [];
  let p: RawPara | null = null;
  let i = false;
  let b = false;
  let inRun = false;
  let inRunProps = false;
  let skip = 0; // inside instrText / delText / footnote text
  let inT = false;
  let pendingBreak = false;
  const push = (text: string) => {
    if (!p || !text) return;
    const last = p.runs[p.runs.length - 1];
    if (last && !!last.i === i && !!last.b === b) last.text += text;
    else p.runs.push({ text, ...(i ? { i: true } : {}), ...(b ? { b: true } : {}) });
  };
  for (const m of xml.matchAll(/<[^>]+>|[^<]+/g)) {
    const t = m[0];
    if (t[0] !== '<') {
      if (inT && !skip) push(decode(t));
      continue;
    }
    const name = /^<\/?([\w:]+)/.exec(t)?.[1] ?? '';
    const close = t[1] === '/';
    const selfClose = t.endsWith('/>');
    switch (name) {
      case 'w:p':
        if (close) {
          if (p) out.push(p);
          p = null;
        } else {
          p = { runs: [], style: '', breakBefore: pendingBreak, breakAfter: false };
          pendingBreak = false;
          if (selfClose) {
            out.push(p);
            p = null;
          }
        }
        break;
      case 'w:pStyle':
        if (p) p.style = styles.get(attr(t, 'w:val') ?? '') ?? (attr(t, 'w:val') ?? '').toLowerCase();
        break;
      case 'w:jc':
        if (p && !inRun) {
          const v = attr(t, 'w:val');
          if (v === 'center') p.align = 'center';
          else if (v === 'right' || v === 'end') p.align = 'right';
        }
        break;
      case 'w:pageBreakBefore':
        if (p && on(t)) p.breakBefore = true;
        break;
      case 'w:r':
        inRun = !close && !selfClose;
        if (!close) {
          i = false;
          b = false;
        }
        break;
      case 'w:rPr':
        inRunProps = inRun && !close && !selfClose;
        break;
      case 'w:i':
        if (inRunProps) i = on(t);
        break;
      case 'w:b':
        if (inRunProps) b = on(t);
        break;
      case 'w:rStyle':
        if (inRunProps) {
          const n = styles.get(attr(t, 'w:val') ?? '') ?? '';
          if (/emphasis|italic/.test(n)) i = true;
          if (/strong|bold/.test(n)) b = true;
        }
        break;
      case 'w:t':
        inT = !close && !selfClose;
        break;
      case 'w:tab':
        if (inRun) push(' ');
        break;
      case 'w:br':
      case 'w:cr':
        if (attr(t, 'w:type') === 'page') {
          if (p && p.runs.some((r) => r.text.trim())) p.breakAfter = true;
          else if (p) p.breakBefore = true;
          else pendingBreak = true;
        } else if (inRun) push('\n');
        break;
      case 'w:instrText':
      case 'w:delText':
      case 'w:footnoteReference':
      case 'w:endnoteReference':
        if (name.endsWith('Reference')) stats.footnotes++;
        else if (!selfClose) skip += close ? -1 : 1;
        break;
      case 'w:drawing':
      case 'w:pict':
        if (!close) stats.images++;
        break;
      case 'w:tbl':
        if (!close) stats.tables++;
        break;
    }
  }
  return out;
}

const SCENE = /^\s*([*#~•·⁂❧✻\-_=]\s*){1,5}\s*$/;
const CHAPTER_WORD = /^\s*(chapter|prologue|epilogue|part|interlude|afterword|foreword|introduction|preface|acknowledg(e)?ments|about the author)\b/i;
const headingLevel = (style: string) => {
  if (style === 'title') return 1;
  const m = /^heading\s*(\d)/.exec(style);
  return m ? Number(m[1]) : 0;
};
const textOf = (p: { runs: Run[] }) => p.runs.map((r) => r.text).join('');

/** Turn raw paragraphs into chapters. Exported for tests. */
export function buildManuscript(raw: RawPara[], stats: Manuscript['stats']): Manuscript {
  const hasHeadings = raw.some((p) => headingLevel(p.style) === 1 && textOf(p).trim());
  const hasWords = !hasHeadings && raw.filter((p) => textOf(p).length < 60 && CHAPTER_WORD.test(textOf(p))).length >= 2;
  const isChapter = (p: RawPara) => {
    const t = textOf(p).trim();
    if (!t) return false;
    if (hasHeadings) return headingLevel(p.style) === 1;
    if (hasWords) return t.length < 60 && CHAPTER_WORD.test(t);
    return false;
  };
  const chapterRule: Manuscript['chapterRule'] = hasHeadings ? 'headings' : hasWords ? 'chapter-words' : 'none';

  const front: Block[] = [];
  const chapters: Chapter[] = [];
  let cur: Block[] = front;
  let lastWasHeadingAdjacent = false;
  for (const p of raw) {
    const t = textOf(p);
    if (isChapter(p)) {
      // A chapter subtitle on the next line (e.g. "Chapter 1" then "The Storm") is joined to the title.
      chapters.push({ title: t.trim().replace(/\s+/g, ' '), blocks: [] });
      cur = chapters[chapters.length - 1]!.blocks;
      lastWasHeadingAdjacent = true;
      continue;
    }
    if (p.breakBefore && cur.length && cur[cur.length - 1]!.kind !== 'pagebreak') cur.push({ kind: 'pagebreak' });
    if (!t.trim()) {
      lastWasHeadingAdjacent = false;
      continue; // empty paragraphs are Word spacing, not content
    }
    if (SCENE.test(t)) cur.push({ kind: 'scene' });
    else {
      const lvl = headingLevel(p.style);
      const runs = p.runs.map((r) => ({ ...r, text: r.text.replace(/[  ]{2,}/g, ' ') }));
      // Trim the paragraph's outer whitespace.
      runs[0]!.text = runs[0]!.text.replace(/^\s+/, '');
      runs[runs.length - 1]!.text = runs[runs.length - 1]!.text.replace(/\s+$/, '');
      const ch = chapters[chapters.length - 1];
      if (lastWasHeadingAdjacent && ch && cur === ch.blocks && !cur.length && lvl === 2) {
        ch.title += `: ${t.trim()}`;
        lastWasHeadingAdjacent = false;
        continue;
      }
      cur.push({ kind: 'para', runs: runs.filter((r) => r.text), ...(p.align ? { align: p.align } : {}), ...(lvl === 2 || lvl === 3 ? { heading: lvl } : {}) });
      stats.paragraphs++;
      stats.words += t.split(/\s+/).filter(Boolean).length;
    }
    if (p.breakAfter) cur.push({ kind: 'pagebreak' });
    lastWasHeadingAdjacent = false;
  }
  // No chapter markers at all: page breaks split the book, or it is one long chapter.
  if (!chapters.length) {
    const parts: Block[][] = [[]];
    for (const b of front) {
      if (b.kind === 'pagebreak') parts.push([]);
      else parts[parts.length - 1]!.push(b);
    }
    const filled = parts.filter((x) => x.length);
    return {
      front: [],
      chapters: filled.map((blocks) => ({ title: '', blocks })),
      stats,
      chapterRule: filled.length > 1 ? 'page-breaks' : 'none',
    };
  }
  return { front, chapters, stats, chapterRule };
}

/** Read a .docx file. Throws a readable Error for anything else. */
export async function readDocx(bytes: Uint8Array): Promise<Manuscript> {
  let files: Map<string, Uint8Array>;
  try {
    files = await readZip(bytes);
  } catch {
    throw new Error('That file is not a Word document (.docx). In Word, use File > Save As > Word Document.');
  }
  const doc = files.get('word/document.xml');
  if (!doc) throw new Error('That file is not a Word document (.docx).');
  const td = new TextDecoder();
  const stats = { words: 0, paragraphs: 0, images: 0, tables: 0, footnotes: 0 };
  const raw = paragraphs(td.decode(doc), styleNames(files.has('word/styles.xml') ? td.decode(files.get('word/styles.xml')!) : undefined), stats);
  return buildManuscript(raw, stats);
}

/** Plain text of a block list (for word counts and tests). */
export const blockText = (bs: Block[]) => bs.map((b) => (b.kind === 'para' ? b.runs.map((r) => r.text).join('') : '')).join('\n');
