/**
 * EPUB 3 from a manuscript: one XHTML file per chapter, a navigation document, simple book CSS and an
 * optional cover image. Reflowable, as Kindle and other eBook stores expect.
 */

import type { Block, Manuscript, Run } from './manuscript';
import type { BookStyle } from './typeset';
import { writeZip } from './zip';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const runs = (rs: Run[]) =>
  rs
    .map((r) => {
      let t = esc(r.text).replace(/\n/g, '<br/>');
      if (r.i) t = `<em>${t}</em>`;
      if (r.b) t = `<strong>${t}</strong>`;
      return t;
    })
    .join('');

function body(bs: Block[]): string {
  let first = true;
  const out: string[] = [];
  for (const b of bs) {
    if (b.kind === 'pagebreak') continue;
    if (b.kind === 'scene') {
      out.push('<p class="scene">*&#160;&#160;*&#160;&#160;*</p>');
      first = true;
      continue;
    }
    if (b.heading) {
      out.push(`<h${b.heading} class="sub">${runs(b.runs)}</h${b.heading}>`);
      first = true;
      continue;
    }
    const cls = b.align === 'center' ? 'center' : b.align === 'right' ? 'right' : first ? 'first' : '';
    out.push(`<p${cls ? ` class="${cls}"` : ''}>${runs(b.runs)}</p>`);
    first = false;
  }
  return out.join('\n');
}

const page = (title: string, content: string, lang: string, bodyType = '') => `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="${lang}" xml:lang="${lang}">
<head><meta charset="UTF-8"/><title>${esc(title)}</title><link rel="stylesheet" type="text/css" href="style.css"/></head>
<body${bodyType ? ` epub:type="${bodyType}"` : ''}>
${content}
</body>
</html>`;

const CSS = `body { margin: 0; font-family: serif; line-height: 1.45; }
p { margin: 0; text-indent: 1.5em; text-align: justify; }
p.first, p.center, p.right, p.scene { text-indent: 0; }
p.center, p.scene { text-align: center; }
p.right { text-align: right; }
p.scene { margin: 1em 0; }
h1 { text-align: center; font-weight: normal; margin: 3em 0 2em; page-break-before: always; }
h2.sub, h3.sub { text-align: center; margin: 1.5em 0 0.8em; }
.title { text-align: center; margin-top: 30%; }
.title h1 { margin: 0 0 0.4em; font-size: 2em; page-break-before: auto; }
.title .subtitle { font-style: italic; text-indent: 0; text-align: center; }
.title .author { margin-top: 3em; text-indent: 0; text-align: center; font-size: 1.2em; }
.copyright p { text-indent: 0; text-align: center; font-size: 0.85em; margin-bottom: 0.5em; }
.dedication p { text-indent: 0; text-align: center; font-style: italic; margin-top: 30%; }
.cover { margin: 0; padding: 0; text-align: center; }
.cover img { max-width: 100%; max-height: 100%; }
`;

export interface EpubOptions {
  style: BookStyle;
  /** JPEG front cover, if any. */
  cover?: Uint8Array;
  lang?: string;
  /** Identifier: an ISBN or any unique id. */
  id?: string;
  /** ISO time for dcterms:modified (tests pass a fixed value). */
  modified?: string;
}

export function buildEpub(ms: Manuscript, o: EpubOptions): Uint8Array {
  const s = o.style;
  const lang = o.lang ?? 'en';
  const title = s.title || 'Untitled';
  const id = o.id ?? `urn:uuid:${crypto.randomUUID()}`;
  const modified = (o.modified ?? new Date().toISOString()).replace(/\.\d+Z$/, 'Z');
  const docs: { id: string; file: string; title: string; xhtml: string; nav: boolean }[] = [];

  if (o.cover) docs.push({ id: 'cover-page', file: 'cover.xhtml', title: 'Cover', xhtml: page('Cover', '<div class="cover"><img src="cover.jpg" alt="Cover"/></div>', lang, 'cover'), nav: false });
  docs.push({
    id: 'title-page',
    file: 'title.xhtml',
    title: 'Title Page',
    xhtml: page(
      title,
      `<div class="title"><h1>${esc(title)}</h1>${s.subtitle ? `<p class="subtitle">${esc(s.subtitle)}</p>` : ''}${s.author ? `<p class="author">${esc(s.author)}</p>` : ''}</div>`,
      lang,
      'titlepage',
    ),
    nav: false,
  });
  if (s.copyright.trim())
    docs.push({ id: 'copyright', file: 'copyright.xhtml', title: 'Copyright', xhtml: page('Copyright', `<div class="copyright">${s.copyright.split('\n').map((l) => `<p>${esc(l)}</p>`).join('')}</div>`, lang, 'copyright-page'), nav: false });
  if (s.dedication.trim())
    docs.push({ id: 'dedication', file: 'dedication.xhtml', title: 'Dedication', xhtml: page('Dedication', `<div class="dedication">${s.dedication.split('\n').map((l) => `<p>${esc(l)}</p>`).join('')}</div>`, lang, 'dedication'), nav: false });
  if (ms.front.length) docs.push({ id: 'front', file: 'front.xhtml', title: 'Front matter', xhtml: page('Front matter', body(ms.front), lang), nav: false });
  ms.chapters.forEach((ch, i) => {
    const t = ch.title || (ms.chapters.length > 1 ? `Part ${i + 1}` : title);
    docs.push({ id: `ch${i + 1}`, file: `chapter-${i + 1}.xhtml`, title: t, xhtml: page(t, `${ch.title ? `<h1>${esc(ch.title)}</h1>` : ''}\n${body(ch.blocks)}`, lang, 'chapter'), nav: true });
  });

  const navList = docs.filter((d) => d.nav || d.id === 'title-page').map((d) => `<li><a href="${d.file}">${esc(d.title)}</a></li>`).join('\n');
  const nav = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="${lang}" xml:lang="${lang}">
<head><meta charset="UTF-8"/><title>Contents</title><link rel="stylesheet" type="text/css" href="style.css"/></head>
<body>
<nav epub:type="toc" id="toc"><h1>Contents</h1><ol>
${navList}
</ol></nav>
<nav epub:type="landmarks" hidden=""><ol>
${o.cover ? '<li><a epub:type="cover" href="cover.xhtml">Cover</a></li>' : ''}
<li><a epub:type="bodymatter" href="${docs.find((d) => d.nav)?.file ?? 'title.xhtml'}">Start</a></li>
</ol></nav>
</body>
</html>`;

  const opf = `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid" xml:lang="${lang}">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:identifier id="bookid">${esc(id)}</dc:identifier>
<dc:title>${esc(title)}</dc:title>
${s.author ? `<dc:creator>${esc(s.author)}</dc:creator>` : ''}
<dc:language>${lang}</dc:language>
<meta property="dcterms:modified">${modified}</meta>
${o.cover ? '<meta name="cover" content="cover-image"/>' : ''}
</metadata>
<manifest>
<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
<item id="css" href="style.css" media-type="text/css"/>
${o.cover ? '<item id="cover-image" href="cover.jpg" media-type="image/jpeg" properties="cover-image"/>' : ''}
${docs.map((d) => `<item id="${d.id}" href="${d.file}" media-type="application/xhtml+xml"/>`).join('\n')}
</manifest>
<spine>
${docs.map((d) => `<itemref idref="${d.id}"${d.id === 'cover-page' ? ' linear="no"' : ''}/>`).join('\n')}
</spine>
</package>`;

  const container = `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>`;

  return writeZip([
    { name: 'mimetype', data: 'application/epub+zip' },
    { name: 'META-INF/container.xml', data: container },
    { name: 'OEBPS/content.opf', data: opf },
    { name: 'OEBPS/nav.xhtml', data: nav },
    { name: 'OEBPS/style.css', data: CSS },
    ...(o.cover ? [{ name: 'OEBPS/cover.jpg', data: o.cover }] : []),
    ...docs.map((d) => ({ name: `OEBPS/${d.file}`, data: d.xhtml })),
  ]);
}
