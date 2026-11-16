/**
 * Publish: a finished book in, KDP's files out. A Word manuscript is typeset into a print interior (or a
 * finished interior PDF is checked as-is); the page count sets the cover size; the cover is checked or built
 * from art; and the interior PDF, print cover PDF, eBook cover JPG and EPUB are exported.
 */

import { PDFDocument } from 'pdf-lib';
import {
  BOOK_FONTS,
  buildCoverArtPdf,
  buildCoverScene,
  buildEpub,
  drawScene,
  formatSize,
  getPaper,
  getTrimSize,
  pageLimits,
  readDocx,
  readInteriorPdf,
  runInteriorPreflight,
  runPreflight,
  sceneToPdf,
  typesetBook,
  wrapRegions,
  type Check,
  type InteriorResult,
  type Manuscript,
  type TypesetResult,
} from '../../core';
import { cropRegion, getPdfjs, loadArtwork, newCanvas, type Artwork } from '../artwork';
import type { Ctx, Screen } from '../context';
import { $, bindFileInput, esc, segment, times, toast } from '../dom';
import { saveFile } from '../files';
import { scanInterior } from '../interiorScan';
import { loadBookFonts, loadFonts, readImage, type PageImage } from '../media';

const ICON: Record<Check['status'], string> = { pass: '✓', warn: '!', fail: '✕', info: 'i' };
const LABEL: Record<Check['status'], string> = { pass: 'Pass', warn: 'Warning', fail: 'Fail', info: 'Note' };
const checkList = (cs: Check[]) =>
  cs
    .map(
      (c) =>
        `<li class="check-row ${c.status}"><span class="ck" role="img" aria-label="${LABEL[c.status]}">${ICON[c.status]}</span><div><strong>${esc(c.title)}</strong><span>${esc(times(c.detail))}</span></div></li>`,
    )
    .join('');

const slug = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'book';

export function initPublishScreen(ctx: Ctx & { setPages(n: number): void }): Screen {
  const pub = ctx.state.publish;
  const st = pub.style;
  let manuscript: { name: string; ms: Manuscript } | null = null;
  let finishedPdf: { name: string; bytes: Uint8Array } | null = null;
  let interior: { pdf: Uint8Array; pageCount: number; warnings: string[]; result: TypesetResult | null; check: InteriorResult | null } | null = null;
  let builtSig = '';
  let building = false;
  let spread = 0;
  let coverSource: 'wrap' | 'art' = 'wrap';
  let wrap: { art: Artwork; bytes: Uint8Array } | null = null;
  let frontArt: PageImage | null = null;
  let backArt: PageImage | null = null;
  let visible = false;

  const el = {
    primary: $('pub-primary'),
    alt: $('pub-alt'),
    source: $('pub-source'),
    sourceHint: $('pub-source-hint'),
    drop: $('pub-drop'),
    file: $<HTMLInputElement>('pub-file'),
    dropTitle: $('pub-drop-title'),
    dropSub: $('pub-drop-sub'),
    status: $('pub-file-status'),
    warnings: $('pub-warnings'),
    styleCard: $('pub-style-card'),
    title: $<HTMLInputElement>('pub-title'),
    subtitle: $<HTMLInputElement>('pub-subtitle'),
    author: $<HTMLInputElement>('pub-author'),
    font: $<HTMLSelectElement>('pub-font'),
    size: $<HTMLSelectElement>('pub-size'),
    leading: $<HTMLSelectElement>('pub-leading'),
    headers: $<HTMLSelectElement>('pub-headers'),
    recto: $<HTMLInputElement>('pub-recto'),
    justify: $<HTMLInputElement>('pub-justify'),
    numbers: $<HTMLInputElement>('pub-numbers'),
    front: $<HTMLInputElement>('pub-front'),
    frontFields: $('pub-front-fields'),
    copyright: $<HTMLTextAreaElement>('pub-copyright'),
    dedication: $<HTMLTextAreaElement>('pub-dedication'),
    previewCard: $('pub-preview-card'),
    summary: $('pub-summary'),
    spread: $('pub-spread'),
    prev: $<HTMLButtonElement>('pub-prev'),
    next: $<HTMLButtonElement>('pub-next'),
    pageno: $('pub-pageno'),
    checks: $('pub-checks'),
    coverCard: $('pub-cover-card'),
    coverSize: $('pub-cover-size'),
    template: $<HTMLButtonElement>('pub-template'),
    coverSourceSeg: $('pub-cover-source'),
    wrapIn: $('pub-cover-wrap-in'),
    artIn: $('pub-cover-art-in'),
    coverDrop: $('pub-cover-drop'),
    coverFile: $<HTMLInputElement>('pub-cover-file'),
    frontArt: $<HTMLInputElement>('pub-front-art'),
    backArt: $<HTMLInputElement>('pub-back-art'),
    spineColor: $<HTMLInputElement>('pub-spine-color'),
    spineTextColor: $<HTMLInputElement>('pub-spine-text-color'),
    artStatus: $('pub-art-status'),
    coverVerdict: $('pub-cover-verdict'),
    coverChecks: $('pub-cover-checks'),
    coverPreview: $('pub-cover-preview'),
    exportCard: $('pub-export-card'),
    dlInterior: $<HTMLButtonElement>('pub-dl-interior'),
    dlCover: $<HTMLButtonElement>('pub-dl-cover'),
    dlEbookCover: $<HTMLButtonElement>('pub-dl-ebook-cover'),
    dlEpub: $<HTMLButtonElement>('pub-dl-epub'),
    checklist: $('pub-checklist'),
  };

  // ---- controls ----------------------------------------------------------------------------

  const sourceSeg = segment<'docx' | 'pdf'>(
    el.source,
    [
      { value: 'docx', label: 'Word manuscript (.docx)' },
      { value: 'pdf', label: 'Finished interior (PDF)' },
    ],
    () => pub.source,
    (v) => {
      pub.source = v;
      ctx.save();
      interior = null;
      builtSig = '';
      refresh();
    },
  );
  const coverSeg = segment<'wrap' | 'art'>(
    el.coverSourceSeg,
    [
      { value: 'wrap', label: 'I have a finished cover' },
      { value: 'art', label: 'Build it from front and back art' },
    ],
    () => coverSource,
    (v) => {
      coverSource = v;
      refreshCover();
    },
  );

  el.font.innerHTML = BOOK_FONTS.map((f) => `<option value="${f.id}">${esc(f.label)}</option>`).join('');
  const writeStyle = () => {
    el.title.value = st.title;
    el.subtitle.value = st.subtitle;
    el.author.value = st.author;
    el.font.value = st.font;
    el.size.value = String(st.size);
    el.leading.value = String(st.leading);
    el.headers.value = st.headers;
    el.recto.checked = st.chapterRecto;
    el.justify.checked = st.justify;
    el.numbers.checked = st.pageNumbers;
    el.front.checked = st.frontMatter;
    el.copyright.value = st.copyright;
    el.dedication.value = st.dedication;
    el.frontFields.hidden = !st.frontMatter;
  };
  writeStyle();

  let debounce = 0;
  const changed = () => {
    st.title = el.title.value;
    st.subtitle = el.subtitle.value;
    st.author = el.author.value;
    st.font = el.font.value as typeof st.font;
    st.size = Number(el.size.value);
    st.leading = Number(el.leading.value);
    st.headers = el.headers.value as typeof st.headers;
    st.chapterRecto = el.recto.checked;
    st.justify = el.justify.checked;
    st.pageNumbers = el.numbers.checked;
    st.frontMatter = el.front.checked;
    st.copyright = el.copyright.value;
    st.dedication = el.dedication.value;
    el.frontFields.hidden = !st.frontMatter;
    ctx.save();
    clearTimeout(debounce);
    debounce = window.setTimeout(() => void rebuild(), 450);
  };
  for (const i of [el.title, el.subtitle, el.author, el.copyright, el.dedication]) i.addEventListener('input', changed);
  for (const i of [el.font, el.size, el.leading, el.headers, el.recto, el.justify, el.numbers, el.front]) i.addEventListener('change', changed);

  // ---- manuscript in -------------------------------------------------------------------------

  bindFileInput(el.file, el.drop, async (file) => {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const isPdf = bytes[0] === 0x25 && bytes[1] === 0x50; // %P
    if (isPdf !== (pub.source === 'pdf')) {
      pub.source = isPdf ? 'pdf' : 'docx';
      ctx.save();
    }
    el.drop.classList.add('busy');
    try {
      if (isPdf) {
        finishedPdf = { name: file.name, bytes };
        manuscript = null;
      } else {
        const ms = await readDocx(bytes);
        manuscript = { name: file.name, ms };
        finishedPdf = null;
        // A first guess at the title from the file name.
        if (!st.title) {
          const small = new Set(['a', 'an', 'and', 'the', 'of', 'in', 'on', 'for', 'to', 'at', 'by', 'or']);
          st.title = file.name
            .replace(/\.docx$/i, '')
            .replace(/[_-]+/g, ' ')
            .trim()
            .split(/\s+/)
            .map((w, k) => (k && small.has(w.toLowerCase()) ? w.toLowerCase() : w[0]!.toUpperCase() + w.slice(1)))
            .join(' ');
          // Front matter already in the manuscript means it has its own title page.
          if (ms.front.length) st.frontMatter = false;
          writeStyle();
          ctx.save();
        }
      }
      builtSig = '';
      spread = 0;
      await rebuild();
    } catch (err) {
      toast(err instanceof Error ? err.message : String(err));
    } finally {
      el.drop.classList.remove('busy');
    }
  });

  function trim(): { w: number; h: number } | null {
    const l = ctx.layout();
    if (l) return { w: l.input.trimWidth, h: l.input.trimHeight };
    return null;
  }

  async function rebuild(): Promise<void> {
    const t = trim();
    if (!t || building) return;
    const sig = JSON.stringify([pub.source, manuscript?.name, finishedPdf?.name, finishedPdf?.bytes.length, st, t, ctx.state.binding, ctx.state.paper]);
    if (sig === builtSig) return;
    building = true;
    el.status.textContent = pub.source === 'docx' ? 'Typesetting…' : 'Checking the PDF…';
    try {
      if (pub.source === 'docx' && manuscript) {
        const r = await typesetBook(manuscript.ms, st, t, await loadBookFonts(st.font));
        interior = { pdf: r.pdf, pageCount: r.pageCount, warnings: r.warnings, result: r, check: null };
      } else if (pub.source === 'pdf' && finishedPdf) {
        const facts = await readInteriorPdf(finishedPdf.bytes);
        facts.pages = await scanInterior(finishedPdf.bytes, facts.pages, (d, n) => (el.status.textContent = `Checking page ${d} of ${n}…`));
        interior = { pdf: finishedPdf.bytes, pageCount: facts.pages.length, warnings: [], result: null, check: null };
        const st2 = ctx.state;
        interior.check = runInteriorPreflight(
          { trimWidth: t.w, trimHeight: t.h, binding: st2.binding, limits: pageLimits(st2.binding, getTrimSize(st2.trimId, st2.binding), getPaper(st2.paper)), coverPages: null },
          facts,
        );
      } else interior = null;
      builtSig = JSON.stringify([pub.source, manuscript?.name, finishedPdf?.name, finishedPdf?.bytes.length, st, t, ctx.state.binding, ctx.state.paper]);
      if (interior && ctx.state.pages !== interior.pageCount) ctx.setPages(interior.pageCount); // sizes the cover (refreshes this screen)
    } catch (err) {
      interior = null;
      toast(`Could not build the interior: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      building = false;
    }
    refresh();
    void renderSpread();
  }

  // ---- preview -------------------------------------------------------------------------------

  let renderToken = 0;
  async function renderSpread(): Promise<void> {
    if (!interior || !visible) return;
    const token = ++renderToken;
    const n = interior.pageCount;
    // Spread 0 is page 1 alone (a right-hand page); spread k shows pages 2k and 2k+1.
    const pages = spread === 0 ? [1] : [2 * spread, 2 * spread + 1].filter((p) => p <= n);
    el.pageno.textContent = pages.length > 1 ? `Pages ${pages[0]}–${pages[1]} of ${n}` : `Page ${pages[0]} of ${n}`;
    el.prev.disabled = spread === 0;
    el.next.disabled = 2 * spread + 1 >= n;
    const pdfjs = await getPdfjs();
    const task = pdfjs.getDocument({ data: interior.pdf.slice() });
    try {
      const doc = await task.promise;
      const canvases: HTMLCanvasElement[] = [];
      const h = Math.min(560, Math.max(320, el.spread.clientWidth * 0.62));
      for (const p of pages) {
        const page = await doc.getPage(p);
        const base = page.getViewport({ scale: 1 });
        const vp = page.getViewport({ scale: ((h / base.height) * (window.devicePixelRatio || 1)) });
        const c = newCanvas(vp.width, vp.height);
        c.className = `spread-page ${p % 2 ? 'recto' : 'verso'}`;
        c.style.height = `${h}px`;
        const g = c.getContext('2d')!;
        g.fillStyle = '#fff';
        g.fillRect(0, 0, c.width, c.height);
        await page.render({ canvas: c, viewport: vp, intent: 'print' }).promise;
        canvases.push(c);
      }
      if (token !== renderToken) return;
      if (spread === 0) {
        const blank = document.createElement('div');
        blank.className = 'spread-page spread-empty';
        blank.style.height = `${h}px`;
        blank.style.aspectRatio = String(canvases[0]!.width / canvases[0]!.height);
        el.spread.replaceChildren(blank, ...canvases);
      } else el.spread.replaceChildren(...canvases);
    } finally {
      await task.destroy();
    }
  }
  el.prev.addEventListener('click', () => {
    spread = Math.max(0, spread - 1);
    void renderSpread();
  });
  el.next.addEventListener('click', () => {
    spread++;
    void renderSpread();
  });
  el.spread.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowLeft') el.prev.click();
    if (e.key === 'ArrowRight') el.next.click();
  });

  // ---- cover ---------------------------------------------------------------------------------

  bindFileInput(el.coverFile, el.coverDrop, async (file) => {
    try {
      const art = await loadArtwork(file);
      wrap = { art, bytes: new Uint8Array(await file.arrayBuffer()) };
    } catch (err) {
      toast(err instanceof Error ? err.message : String(err));
    }
    refreshCover();
  });
  bindFileInput(el.frontArt, null, async (f) => {
    frontArt = await readImage(f);
    refreshCover();
  });
  bindFileInput(el.backArt, null, async (f) => {
    backArt = await readImage(f);
    refreshCover();
  });
  for (const i of [el.spineColor, el.spineTextColor]) i.addEventListener('input', () => refreshCover());

  el.template.addEventListener('click', async () => {
    const l = ctx.layout();
    if (!l) return;
    const bytes = await sceneToPdf(buildCoverScene(l, { mode: 'template', dpi: 300 }), { title: `Cover template for ${st.title || 'your book'}`, fonts: await loadFonts() });
    const r = await saveFile(`${slug(st.title)}_cover-template_${l.input.pageCount}pp.pdf`, bytes, 'application/pdf');
    if (r.saved) toast(r.path ? `Saved ${r.path}` : 'Template downloaded');
  });

  function coverReady(): boolean {
    return coverSource === 'wrap' ? !!wrap : !!frontArt;
  }

  function refreshCover(): void {
    const l = ctx.layout();
    coverSeg.sync();
    el.wrapIn.hidden = coverSource !== 'wrap';
    el.artIn.hidden = coverSource !== 'art';
    el.dlCover.disabled = !l?.valid || !coverReady();
    el.dlEbookCover.disabled = !coverReady();
    el.coverVerdict.hidden = true;
    el.coverChecks.innerHTML = '';
    el.coverPreview.replaceChildren();
    el.coverPreview.hidden = true;
    if (!l) return;
    el.coverSize.textContent = times(
      `${l.binding === 'hardcover' ? 'Hardcover' : 'Paperback'} cover for ${l.input.pageCount} pages: ${formatSize(l.canvas.w, l.canvas.h, 'in')} with a ${l.spineWidth.toFixed(3)}" spine.`,
    );
    if (coverSource === 'art') {
      el.artStatus.textContent = frontArt ? `Front: ${frontArt.name}${backArt ? `, back: ${backArt.name}` : ''}.` : 'Choose at least a front cover image.';
    }
    if (coverSource === 'wrap' && wrap) {
      const r = runPreflight(l, wrap.art.facts);
      el.coverVerdict.hidden = false;
      el.coverVerdict.className = `verdict ${r.verdict}`;
      el.coverVerdict.textContent = r.verdict === 'pass' ? `${wrap.art.name}: fits this book.` : r.verdict === 'warn' ? `${wrap.art.name}: check the warnings.` : `${wrap.art.name}: needs fixing for this book.`;
      el.coverChecks.innerHTML = checkList(r.checks);
      // Preview with the guides on top.
      const w = Math.min(1400, Math.round((el.coverPreview.clientWidth || 800) * (window.devicePixelRatio || 1)));
      const c = newCanvas(w, (w * l.canvas.h) / l.canvas.w);
      c.style.width = '100%';
      c.style.height = 'auto';
      const g = c.getContext('2d')!;
      g.drawImage(wrap.art.canvas, 0, 0, c.width, c.height);
      drawScene(g, buildCoverScene(l, { mode: 'overlay', dpi: 300 }), c.width / l.canvas.w);
      el.coverPreview.replaceChildren(c);
      el.coverPreview.hidden = false;
    }
  }

  async function printCoverPdf(): Promise<Uint8Array> {
    const l = ctx.layout()!;
    if (coverSource === 'art') {
      if (!frontArt) throw new Error('Choose a front cover image.');
      return buildCoverArtPdf(
        l,
        { front: frontArt, back: backArt, spineColor: el.spineColor.value, backColor: el.spineColor.value, spineText: l.spineTextAllowed ? [st.title, st.author].filter(Boolean).join('   ') : '', spineTextColor: el.spineTextColor.value, title: st.title },
        await loadFonts(),
      );
    }
    if (!wrap) throw new Error('Choose your cover file.');
    if (wrap.art.kind === 'pdf') return wrap.bytes;
    // An image cover: place it on a page of the exact cover size.
    const doc = await PDFDocument.create();
    const img = wrap.art.kind === 'png' ? await doc.embedPng(wrap.bytes) : await doc.embedJpg(wrap.bytes);
    const page = doc.addPage([l.canvas.w * 72, l.canvas.h * 72]);
    page.drawImage(img, { x: 0, y: 0, width: l.canvas.w * 72, height: l.canvas.h * 72 });
    doc.setTitle(`${st.title || 'Book'} cover`);
    return doc.save();
  }

  /** Front cover at 1600 px wide, in the trim's proportions, as a JPG (KDP eBook covers are JPG or TIFF). */
  async function ebookCoverJpg(): Promise<Uint8Array> {
    const l = ctx.layout()!;
    const W = 1600;
    const H = Math.round((W * l.input.trimHeight) / l.input.trimWidth);
    const out = newCanvas(W, H);
    const g = out.getContext('2d')!;
    g.imageSmoothingQuality = 'high';
    let src: CanvasImageSource;
    let sw: number;
    let sh: number;
    if (coverSource === 'art' && frontArt) {
      const bmp = await createImageBitmap(new Blob([frontArt.bytes as BlobPart]));
      src = bmp;
      sw = bmp.width;
      sh = bmp.height;
    } else if (wrap) {
      // The front panel of the wrap, without its bleed.
      const r = wrapRegions(l).front;
      const bleedX = 0.125 / l.canvas.w;
      const bleedY = 0.125 / l.canvas.h;
      const spineLeft = l.input.readingDirection === 'ltr';
      const crop = cropRegion(wrap.art.canvas, { x: spineLeft ? r.x : r.x + bleedX, y: r.y + bleedY, w: r.w - bleedX, h: r.h - 2 * bleedY }, 4096);
      src = crop;
      sw = crop.width;
      sh = crop.height;
    } else throw new Error('Add your cover first.');
    // Cover-fit into the frame.
    const k = Math.max(W / sw, H / sh);
    g.drawImage(src, (W - sw * k) / 2, (H - sh * k) / 2, sw * k, sh * k);
    const blob = await new Promise<Blob>((res, rej) => out.toBlob((b) => (b ? res(b) : rej(new Error('Encoding failed'))), 'image/jpeg', 0.92));
    return new Uint8Array(await blob.arrayBuffer());
  }

  // ---- export --------------------------------------------------------------------------------

  const save = async (name: string, make: () => Promise<Uint8Array>, mime: string) => {
    try {
      const r = await saveFile(name, await make(), mime);
      if (r.saved) toast(r.path ? `Saved ${r.path}` : `${name} downloaded`);
    } catch (err) {
      toast(err instanceof Error ? err.message : String(err));
    }
  };
  el.dlInterior.addEventListener('click', () => {
    if (interior) void save(`${slug(st.title)}_interior_${interior.pageCount}pp.pdf`, async () => interior!.pdf, 'application/pdf');
  });
  el.dlCover.addEventListener('click', () => void save(`${slug(st.title)}_print-cover_${ctx.state.pages}pp.pdf`, printCoverPdf, 'application/pdf'));
  el.dlEbookCover.addEventListener('click', () => void save(`${slug(st.title)}_ebook-cover.jpg`, ebookCoverJpg, 'image/jpeg'));
  el.dlEpub.addEventListener('click', () => {
    if (!manuscript) return;
    void save(
      `${slug(st.title)}.epub`,
      async () => buildEpub(manuscript!.ms, { style: st, cover: coverReady() ? await ebookCoverJpg().catch(() => undefined) : undefined }),
      'application/epub+zip',
    );
  });

  // ---- refresh -------------------------------------------------------------------------------

  function refresh(): void {
    sourceSeg.sync();
    const docx = pub.source === 'docx';
    el.sourceHint.textContent = docx
      ? 'Your finished, edited manuscript. Chapters are found from Heading 1 styles or lines like "Chapter 1"; italics, bold and scene breaks (* * *) are kept; Word’s fonts and spacing are replaced by the book style below.'
      : 'An interior PDF you made elsewhere (any kind of book). It is checked against KDP’s rules and used as-is.';
    el.dropTitle.textContent = docx ? 'Choose your Word manuscript (.docx)' : 'Choose your interior PDF';
    el.styleCard.hidden = !docx;
    el.dlEpub.hidden = !docx;

    const t = trim();
    const l = ctx.layout();
    const ms = manuscript?.ms;
    if (docx && ms) {
      const rule = { headings: 'from Heading 1', 'chapter-words': 'from lines like "Chapter 1"', 'page-breaks': 'from page breaks', none: 'none found, so the text runs as one section' }[ms.chapterRule];
      el.status.textContent = building ? 'Typesetting…' : `${manuscript!.name}: ${ms.stats.words.toLocaleString()} words, ${ms.chapters.length} chapter${ms.chapters.length === 1 ? '' : 's'} (${rule}).`;
    } else if (!docx && finishedPdf) el.status.textContent = building ? 'Checking…' : `${finishedPdf.name}`;
    else el.status.textContent = t ? '' : 'Choose a trim size on the left first.';

    const warns = interior?.warnings ?? [];
    el.warnings.hidden = !warns.length;
    el.warnings.innerHTML = warns.map((w) => `<li>${esc(w)}</li>`).join('');

    const have = !!interior;
    el.previewCard.hidden = !have;
    el.coverCard.hidden = !have;
    el.exportCard.hidden = !have;
    if (!interior || !t) {
      el.primary.textContent = 'Drop in your manuscript';
      return;
    }
    el.primary.textContent = times(`${interior.pageCount} pages · ${formatSize(t.w, t.h, 'in')}`);
    el.alt.textContent = l ? times(`${l.binding === 'hardcover' ? 'Hardcover' : 'Paperback'} · ${getPaper(l.input.paper).label.toLowerCase()} · cover ${formatSize(l.canvas.w, l.canvas.h, 'in')}`) : '';
    const m = interior.result?.margins;
    el.summary.textContent = m
      ? `No bleed. Margins: ${m.inside.toFixed(3)}" inside (gutter for ${interior.pageCount} pages), ${m.outside}" outside, ${m.top}" top and bottom. Page numbers and running heads sit inside the margins.`
      : 'Your PDF, checked against KDP’s rules:';
    el.checks.innerHTML = interior.check ? checkList(interior.check.checks.filter((c) => c.id !== 'cover-pages')) : '';
    refreshCover();

    const paper = getPaper(ctx.state.paper).label.toLowerCase();
    const bleed = interior.check?.bleed ? 'Bleed' : 'No bleed';
    el.checklist.innerHTML = [
      `<b>Paperback / hardcover content:</b> choose <b>${esc(paper)}</b>, trim <b>${esc(times(formatSize(t.w, t.h, 'in')))}</b>, <b>${bleed}</b>, and the cover finish you like.`,
      `<b>Manuscript:</b> upload the <b>Interior PDF</b>.`,
      `<b>Book cover:</b> choose <i>Upload a cover you already have</i> and upload the <b>Print cover PDF</b> (made for exactly ${interior.pageCount} pages; rebuild it if the page count changes).`,
      `<b>Launch and review:</b> open KDP’s Print Previewer and check every page before you publish.`,
      ...(docx ? [`<b>Kindle eBook:</b> upload the <b>EPUB</b> as the manuscript and the <b>eBook cover JPG</b> as the cover.`] : []),
    ]
      .map((x) => `<li>${x}</li>`)
      .join('');
  }

  return {
    id: 'publish',
    refresh: () => {
      refresh();
      void rebuild();
    },
    activate: () => {
      visible = true;
      void renderSpread();
    },
    deactivate: () => {
      visible = false;
    },
  };
}
