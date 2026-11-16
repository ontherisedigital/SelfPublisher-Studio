/** Images and PDF fonts shared by the Interior, Page designer and Puzzles screens. */

import type { BookFont, BookFontFiles, PdfFonts, PdfImage } from '../core';

export interface PageImage extends PdfImage {
  id: number;
  name: string;
  widthPx: number;
  heightPx: number;
  url: string;
}

let nextId = 1;

/** Read an image file into PNG/JPEG bytes the PDF can embed (other formats are converted to PNG). */
export async function readImage(file: File): Promise<PageImage> {
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  const isPng = head[0] === 0x89 && head[1] === 0x50;
  const isJpg = head[0] === 0xff && head[1] === 0xd8;
  const bmp = await createImageBitmap(file);
  const { width, height } = bmp;
  let bytes: Uint8Array;
  let type: 'png' | 'jpg';
  let blob: Blob = file;
  if (isPng || isJpg) {
    bytes = new Uint8Array(await file.arrayBuffer());
    type = isPng ? 'png' : 'jpg';
    bmp.close();
  } else {
    const c = document.createElement('canvas');
    c.width = width;
    c.height = height;
    c.getContext('2d')!.drawImage(bmp, 0, 0);
    bmp.close();
    blob = await new Promise<Blob>((res, rej) => c.toBlob((b) => (b ? res(b) : rej(new Error('Could not convert the image'))), 'image/png'));
    bytes = new Uint8Array(await blob.arrayBuffer());
    type = 'png';
  }
  return { id: nextId++, name: file.name, widthPx: width, heightPx: height, bytes, type, url: URL.createObjectURL(blob) };
}

let fontCache: Promise<PdfFonts> | null = null;
/** The PDF fonts are bundled inline (KDP needs embedded fonts, and `fetch` cannot read `file://` in the desktop app). */
export function loadFonts(): Promise<PdfFonts> {
  const decode = (dataUrl: string) => Uint8Array.from(atob(dataUrl.slice(dataUrl.indexOf(',') + 1)), (c) => c.charCodeAt(0));
  fontCache ??= Promise.all([
    import('./fonts/fredoka-latin.ttf?inline'),
    import('./fonts/dosis-latin.ttf?inline'),
    import('./fonts/arima-bold-latin.ttf?inline'),
  ]).then(([d, b, a]) => ({ display: decode(d.default), body: decode(b.default), accent: decode(a.default) }));
  return fontCache;
}


// --- Text measuring for SVG previews (matches the PDF's fonts closely enough for line breaks) ---

const FAMILY = { display: 'Fredoka', body: 'Dosis', accent: 'Arima' } as const;
let measureCtx: CanvasRenderingContext2D | null = null;
/** Width in points of `text` at `size` pt in one of the app fonts. */
export function measureText(text: string, size: number, font: keyof typeof FAMILY): number {
  measureCtx ??= document.createElement('canvas').getContext('2d')!;
  measureCtx.font = `${font === 'accent' ? 700 : 500} ${size}px ${FAMILY[font]}, sans-serif`;
  return measureCtx.measureText(text).width;
}

// --- Book fonts (Publish tab): SIL Open Font License, see fonts/book/OFL.txt ---

const BOOK_FONT_FILES = {
  garamond: () => Promise.all([import('./fonts/book/eb-garamond-400-normal.ttf?inline'), import('./fonts/book/eb-garamond-400-italic.ttf?inline'), import('./fonts/book/eb-garamond-700-normal.ttf?inline'), import('./fonts/book/eb-garamond-700-italic.ttf?inline')]),
  crimson: () => Promise.all([import('./fonts/book/crimson-pro-400-normal.ttf?inline'), import('./fonts/book/crimson-pro-400-italic.ttf?inline'), import('./fonts/book/crimson-pro-700-normal.ttf?inline'), import('./fonts/book/crimson-pro-700-italic.ttf?inline')]),
  baskerville: () => Promise.all([import('./fonts/book/libre-baskerville-400-normal.ttf?inline'), import('./fonts/book/libre-baskerville-400-italic.ttf?inline'), import('./fonts/book/libre-baskerville-700-normal.ttf?inline'), import('./fonts/book/libre-baskerville-700-italic.ttf?inline')]),
} as const;
const bookFontCache = new Map<string, Promise<BookFontFiles>>();

export function loadBookFonts(id: BookFont): Promise<BookFontFiles> {
  const decode = (dataUrl: string) => Uint8Array.from(atob(dataUrl.slice(dataUrl.indexOf(',') + 1)), (c) => c.charCodeAt(0));
  let p = bookFontCache.get(id);
  if (!p) {
    p = BOOK_FONT_FILES[id]().then(([r, i, b, bi]) => ({ regular: decode(r.default), italic: decode(i.default), bold: decode(b.default), boldItalic: decode(bi.default) }));
    bookFontCache.set(id, p);
  }
  return p;
}
