const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const FOOTER_REL = R + '/footer';
const FOOTER_CONTENT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const PROFILE = Object.freeze({
  id: 'unand', name: 'Standar Universitas Andalas', shortName: 'Unand',
  font: 'Times New Roman', size: 24, line: 360,
  margins: { left: 2268, top: 2268, right: 1701, bottom: 1701 },
  page: { width: 11906, height: 16838 },
  paragraph: { firstLine: 709, before: 0, after: 0, align: 'both' }
});
export const profile = PROFILE;

const direct = (node, name) => node && [...node.children].find(child => child.namespaceURI === W && child.localName === name) || null;
const all = (node, name) => node ? [...node.getElementsByTagNameNS(W, name)] : [];
const attr = (node, name) => node?.getAttributeNS(W, name) ?? node?.getAttribute('w:' + name) ?? null;
const number = (node, name) => { const raw = attr(node, name); return raw === null ? null : Number(raw); };
const first = (node, name) => all(node, name)[0] || null;
const xmlString = doc => new XMLSerializer().serializeToString(doc);
function parseXml(source, path) {
  const xml = new DOMParser().parseFromString(source, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error('Struktur XML DOCX tidak valid: ' + path);
  return xml;
}
function child(parent, name, before = null) {
  let element = direct(parent, name);
  if (!element) {
    element = parent.ownerDocument.createElementNS(W, 'w:' + name);
    let anchor = before;
    if (!anchor && parent.localName === 'style' && name === 'pPr') anchor = direct(parent, 'rPr');
    if (!anchor && parent.localName === 'docDefaults' && name === 'rPrDefault') anchor = direct(parent, 'pPrDefault');
    if (!anchor && parent.localName === 'rPr') {
      const order = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
      const rank = order.indexOf(name);
      anchor = [...parent.children].find(item => order.indexOf(item.localName) > rank) || null;
    }
    if (!anchor && parent.localName === 'pPr' && name === 'spacing') {
      anchor = [...parent.children].find(item => ['ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl'].includes(item.localName)) || null;
    }
    if (!anchor && parent.localName === 'pPr' && name === 'ind') {
      anchor = [...parent.children].find(item => ['contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl'].includes(item.localName)) || null;
    }
    if (!anchor && parent.localName === 'pPr' && name === 'jc') {
      anchor = [...parent.children].find(item => ['textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr'].includes(item.localName)) || null;
    }
    if (!anchor && parent.localName === 'sectPr' && name === 'pgSz') anchor = direct(parent, 'pgMar');
    if (!anchor && (parent.localName === 'p' || parent.localName === 'r' || parent.localName === 'rPrDefault' || parent.localName === 'pPrDefault')) anchor = parent.firstChild;
    parent.insertBefore(element, anchor);
  }
  return element;
}
function set(node, name, value) { node.setAttributeNS(W, 'w:' + name, String(value)); }
function runFormat(rPr) {
  const fonts = direct(rPr, 'rFonts');
  return {
    font: attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || (attr(fonts, 'asciiTheme') ? 'font tema' : null),
    size: number(direct(rPr, 'sz'), 'val')
  };
}
function spacing(pPr) {
  const item = direct(pPr, 'spacing');
  return { line: number(item, 'line'), rule: attr(item, 'lineRule') };
}
function styleMap(stylesXml) {
  const map = new Map();
  if (!stylesXml) return map;
  for (const style of all(stylesXml, 'style')) {
    const id = attr(style, 'styleId');
    if (id) map.set(id, style);
  }
  return map;
}
function resolveStyle(styles, id) {
  const chain = [];
  const visited = new Set();
  let current = id || 'Normal';
  while (current && styles.has(current) && !visited.has(current)) {
    visited.add(current);
    const style = styles.get(current);
    chain.unshift(style);
    current = attr(direct(style, 'basedOn'), 'val');
  }
  return chain;
}
function effectiveFormat(paragraph, styles, defaults) {
  const pPr = direct(paragraph, 'pPr');
  const styleId = attr(direct(pPr, 'pStyle'), 'val') || 'Normal';
  const chain = resolveStyle(styles, styleId);
  let font = defaults.font, size = defaults.size;
  let line = defaults.line, rule = defaults.rule;
  for (const style of chain) {
    const r = runFormat(direct(style, 'rPr'));
    if (r.font) font = r.font;
    if (r.size !== null) size = r.size;
    const s = spacing(direct(style, 'pPr'));
    if (s.line !== null) { line = s.line; rule = s.rule; }
  }
  const own = spacing(pPr);
  if (own.line !== null) { line = own.line; rule = own.rule; }
  const fonts = new Set(), sizes = new Set();
  const runs = all(paragraph, 'r');
  if (!runs.length) { if (font) fonts.add(font); if (size !== null) sizes.add(size); }
  for (const run of runs) {
    const runStyleId = attr(direct(direct(run, 'rPr'), 'rStyle'), 'val');
    let styledFont = font, styledSize = size;
    for (const style of runStyleId ? resolveStyle(styles, runStyleId) : []) {
      const sf = runFormat(direct(style, 'rPr'));
      if (sf.font) styledFont = sf.font;
      if (sf.size !== null) styledSize = sf.size;
    }
    const ownRun = runFormat(direct(run, 'rPr'));
    if (ownRun.font || styledFont) fonts.add(ownRun.font || styledFont);
    if (ownRun.size !== null || styledSize !== null) sizes.add(ownRun.size ?? styledSize);
  }
  return { fonts: [...fonts], sizes: [...sizes], line, rule, styleId };
}
function defaultsFrom(stylesXml) {
  const defaults = first(stylesXml, 'docDefaults');
  const rPr = first(direct(defaults, 'rPrDefault'), 'rPr');
  const pPr = first(direct(defaults, 'pPrDefault'), 'pPr');
  return { ...runFormat(rPr), ...spacing(pPr) };
}
function paragraphText(p) {
  return all(p, 't').map(t => t.textContent).join('').trim();
}
function headingKind(paragraph, text, styleId) {
  if (paragraph.parentElement?.localName !== 'body' || direct(direct(paragraph, 'pPr'), 'numPr')) return null;
  if (/^BAB\s+(?:\d+|[IVXLCDM]+)(?=\s|[.:\-]|$)/i.test(text)) return 'chapter';
  if (/^\d+\.\d+\.?\s+\S/.test(text) && text.length <= 180 && !/list/i.test(styleId)) return 'subchapter';
  return null;
}
function isBodyParagraph(paragraph, text, styleId, kind) {
  return paragraph.parentElement?.localName === 'body' && !kind &&
    !/heading|title|subtitle|caption|list|toc/i.test(styleId) &&
    !direct(direct(paragraph, 'pPr'), 'numPr') && !/^\s*(?:gambar|tabel)\s+\d+[.:]/i.test(text);
}
function effectiveParagraph(paragraph, styles) {
  const pPr = direct(paragraph, 'pPr');
  const styleId = attr(direct(pPr, 'pStyle'), 'val') || 'Normal';
  let align = null, firstLine = null, before = null, after = null;
  for (const part of [...resolveStyle(styles, styleId).map(style => direct(style, 'pPr')), pPr]) {
    const value = attr(direct(part, 'jc'), 'val');
    if (value !== null) align = value;
    const indent = number(direct(part, 'ind'), 'firstLine');
    if (indent !== null) firstLine = indent;
    const gap = direct(part, 'spacing');
    if (number(gap, 'before') !== null) before = number(gap, 'before');
    if (number(gap, 'after') !== null) after = number(gap, 'after');
  }
  return { align, firstLine, before, after };
}
function headingPrefix(text, kind) {
  return kind === 'chapter'
    ? text.match(/^(\s*BAB\s+)(\d+|[IVXLCDM]+)(?=\s|[.:\-]|$)/i)
    : text.match(/^(\s*)(\d+)\.(\d+)(\.?)(?=\s)/);
}
function setHeadingPrefix(paragraph, length, replacement) {
  const nodes = all(paragraph, 't');
  let remaining = length;
  for (let i = 0; i < nodes.length; i++) {
    const value = nodes[i].textContent;
    if (i === 0) nodes[i].textContent = replacement + value.slice(Math.min(remaining, value.length));
    else if (remaining > 0) nodes[i].textContent = value.slice(Math.min(remaining, value.length));
    remaining -= Math.min(remaining, value.length);
    if (remaining <= 0) break;
  }
}
function checkHeadings(entries, findings, repair = false) {
  let chapter = 0, subsection = 0;
  for (const entry of entries) {
    const kind = headingKind(entry.paragraph, entry.text, entry.styleId);
    if (!kind) continue;
    if (kind === 'chapter') { chapter++; subsection = 0; }
    else if (chapter) subsection++;
    else continue;
    const match = headingPrefix(entry.text, kind);
    if (!match) continue;
    const expected = kind === 'chapter' ? `BAB ${chapter}` : `${chapter}.${subsection}`;
    const current = kind === 'chapter' ? `BAB ${match[2]}` : `${match[2]}.${match[3]}`;
    if (current.toUpperCase() === expected.toUpperCase()) continue;
    if (repair) setHeadingPrefix(entry.paragraph, match[0].length, expected);
    else findings.push(makeFinding('structure', 'Urutan judul belum sesuai', `Tertulis ${current}; seharusnya ${expected}.`, entry.page, entry.index, entry.text));
  }
}
function readSections(documentXml) {
  return all(documentXml, 'sectPr').map((sect, index) => {
    const margins = direct(sect, 'pgMar'), size = direct(sect, 'pgSz');
    return {
      index: index + 1,
      margins: { left: number(margins, 'left'), top: number(margins, 'top'), right: number(margins, 'right'), bottom: number(margins, 'bottom') },
      page: { width: number(size, 'w'), height: number(size, 'h') }
    };
  });
}
function makeFinding(category, title, detail, page, paragraphIndex = null, snippet = '') {
  return { category, title, detail, page, paragraphIndex, snippet };
}
function analyzeXml(documentXml, stylesXml) {
  const styles = styleMap(stylesXml), defaults = defaultsFrom(stylesXml);
  const body = first(documentXml, 'body');
  if (!body) throw new Error('Isi dokumen Word tidak ditemukan.');
  const paragraphs = [], findings = [];
  const sections = readSections(documentXml);
  if (!sections.length) findings.push(makeFinding('page', 'Pengaturan halaman tidak ditemukan', 'Margin dan ukuran A4 akan ditetapkan saat perbaikan.', 1));
  for (const section of sections) {
    const bad = Object.entries(PROFILE.margins).filter(([key, target]) => section.margins[key] !== target);
    if (bad.length) findings.push(makeFinding('page', 'Margin belum sesuai', 'Bagian ' + section.index + ': kiri/atas harus 4 cm, kanan/bawah 3 cm.', 1));
    const size = section.page;
    const correct = Math.abs(size.width - PROFILE.page.width) <= 2 && Math.abs(size.height - PROFILE.page.height) <= 2;
    if (!correct) findings.push(makeFinding('page', 'Ukuran kertas bukan A4', 'Bagian ' + section.index + ': kertas harus A4 (21 × 29,7 cm).', 1));
  }
  let page = 1;
  const headingEntries = [];
  for (const paragraph of all(body, 'p')) {
    const pPr = direct(paragraph, 'pPr');
    if (direct(pPr, 'pageBreakBefore')) page += 1;
    const text = paragraphText(paragraph);
    if (!text) {
      page += all(paragraph, 'lastRenderedPageBreak').length + all(paragraph, 'br').filter(br => attr(br, 'type') === 'page').length;
      continue;
    }
    const format = effectiveFormat(paragraph, styles, defaults);
    const index = paragraphs.length;
    const item = { text, page, index, styleId: format.styleId };
    paragraphs.push(item);
    headingEntries.push({ paragraph, text, page, index, styleId: format.styleId });
    const wrongFonts = format.fonts.filter(font => font.toLowerCase() !== PROFILE.font.toLowerCase());
    if (wrongFonts.length) findings.push(makeFinding('font', 'Jenis font berbeda', 'Ditemukan ' + [...new Set(wrongFonts)].join(', ') + '; seharusnya Times New Roman.', page, index, text));
    const wrongSizes = format.sizes.filter(size => size !== PROFILE.size);
    if (wrongSizes.length) findings.push(makeFinding('font', 'Ukuran font berbeda', 'Ditemukan ' + [...new Set(wrongSizes)].map(size => size / 2 + ' pt').join(', ') + '; seharusnya 12 pt.', page, index, text));
    if (format.line !== PROFILE.line || (format.rule && format.rule !== 'auto')) {
      const description = format.line === null ? 'Spasi tidak ditetapkan' : 'Spasi ' + (format.rule === 'auto' || !format.rule ? (format.line / 240).toFixed(1) + ' baris' : format.line + ' (' + format.rule + ')');
      findings.push(makeFinding('spacing', 'Spasi baris berbeda', description + '; seharusnya 1,5 baris.', page, index, text));
    }
    if (isBodyParagraph(paragraph, text, format.styleId, headingKind(paragraph, text, format.styleId))) {
      const layout = effectiveParagraph(paragraph, styles);
      if (layout.align !== PROFILE.paragraph.align) findings.push(makeFinding('paragraph', 'Perataan paragraf berbeda', 'Paragraf isi seharusnya rata kiri-kanan.', page, index, text));
      if (layout.firstLine !== PROFILE.paragraph.firstLine) findings.push(makeFinding('paragraph', 'Awal paragraf belum sesuai', 'Baris pertama paragraf isi seharusnya menjorok 1,25 cm.', page, index, text));
      if ((layout.before ?? 0) !== 0 || (layout.after ?? 0) !== 0) findings.push(makeFinding('paragraph', 'Jarak antarapagraf berbeda', 'Jarak sebelum dan sesudah paragraf isi seharusnya 0 pt.', page, index, text));
    }
    page += all(paragraph, 'lastRenderedPageBreak').length + all(paragraph, 'br').filter(br => attr(br, 'type') === 'page').length;
  }
  checkHeadings(headingEntries, findings);
  return { paragraphs, findings, sections, pageCount: Math.max(1, page) };
}
function fixRunProperties(rPr) {
  const fonts = child(rPr, 'rFonts');
  for (const theme of ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']) {
    fonts.removeAttributeNS(W, theme);
    fonts.removeAttribute('w:' + theme);
  }
  for (const key of ['ascii', 'hAnsi', 'eastAsia', 'cs']) set(fonts, key, PROFILE.font);
  set(child(rPr, 'sz'), 'val', PROFILE.size);
  set(child(rPr, 'szCs'), 'val', PROFILE.size);
}
function fixParagraphProperties(pPr) {
  const item = child(pPr, 'spacing');
  set(item, 'line', PROFILE.line);
  set(item, 'lineRule', 'auto');
}
function fixContent(xml, includeSections = false) {
  const body = includeSections ? first(xml, 'body') : null;
  const headingEntries = [];
  for (const paragraph of all(xml, 'p')) {
    fixParagraphProperties(child(paragraph, 'pPr'));
    for (const run of all(paragraph, 'r')) fixRunProperties(child(run, 'rPr'));
    if (body && body.contains(paragraph)) {
      const text = paragraphText(paragraph);
      const styleId = attr(direct(direct(paragraph, 'pPr'), 'pStyle'), 'val') || 'Normal';
      const kind = headingKind(paragraph, text, styleId);
      if (isBodyParagraph(paragraph, text, styleId, kind)) {
        const pPr = child(paragraph, 'pPr');
        set(child(pPr, 'jc'), 'val', PROFILE.paragraph.align);
        const ind = child(pPr, 'ind');
        set(ind, 'firstLine', PROFILE.paragraph.firstLine);
        for (const key of ['firstLineChars', 'hanging', 'hangingChars']) { ind.removeAttributeNS(W, key); ind.removeAttribute('w:' + key); }
        const space = child(pPr, 'spacing');
        set(space, 'before', 0); set(space, 'after', 0);
        for (const key of ['beforeAutospacing', 'afterAutospacing']) { space.removeAttributeNS(W, key); space.removeAttribute('w:' + key); }
      }
      if (kind) headingEntries.push({ paragraph, text, styleId });
    }
  }
  if (includeSections) checkHeadings(headingEntries, [], true);
  if (includeSections) {
    const sections = all(xml, 'sectPr');
    if (!sections.length) {
      const section = xml.createElementNS(W, 'w:sectPr');
      first(xml, 'body').appendChild(section);
      sections.push(section);
    }
    for (const sect of sections) {
      const margins = child(sect, 'pgMar');
      for (const [key, value] of Object.entries(PROFILE.margins)) set(margins, key, value);
      const size = child(sect, 'pgSz');
      set(size, 'w', PROFILE.page.width); set(size, 'h', PROFILE.page.height);
      size.removeAttributeNS(W, 'orient'); size.removeAttribute('w:orient');
    }
  }
}
function fixStyles(stylesXml) {
  if (!stylesXml) return;
  for (const style of all(stylesXml, 'style')) {
    const type = attr(style, 'type');
    if (type === 'paragraph') fixParagraphProperties(child(style, 'pPr'));
    if (type === 'paragraph' || type === 'character') fixRunProperties(child(style, 'rPr'));
  }
  const defaults = first(stylesXml, 'docDefaults');
  if (defaults) {
    fixRunProperties(child(child(defaults, 'rPrDefault'), 'rPr'));
    fixParagraphProperties(child(child(defaults, 'pPrDefault'), 'pPr'));
  }
}
function packagePartPath(target) {
  const parts = (target.startsWith('/') ? target.slice(1) : 'word/' + target).split('/');
  const result = [];
  for (const part of parts) {
    if (part === '..') result.pop();
    else if (part && part !== '.') result.push(part);
  }
  return result.join('/');
}
function nextId(relsXml) {
  const ids = [...relsXml.getElementsByTagNameNS(REL, 'Relationship')].map(item => item.getAttribute('Id'));
  let number = 1;
  while (ids.includes('rId' + number)) number++;
  return 'rId' + number;
}
function footerRelationship(relsXml, id) {
  return [...relsXml.getElementsByTagNameNS(REL, 'Relationship')].find(item => item.getAttribute('Id') === id && item.getAttribute('Type') === FOOTER_REL) || null;
}
function footerReference(section, type) {
  return [...section.children].find(item => item.namespaceURI === W && item.localName === 'footerReference' && attr(item, 'type') === type) || null;
}
function attachFooter(section, type, id) {
  let reference = footerReference(section, type);
  if (!reference) {
    reference = section.ownerDocument.createElementNS(W, 'w:footerReference');
    set(reference, 'type', type);
    const anchor = [...section.children].find(item => !['headerReference', 'footerReference'].includes(item.localName)) || null;
    section.insertBefore(reference, anchor);
  }
  reference.setAttributeNS(R, 'r:id', id);
}
function placePageField(footerXml, position) {
  const root = footerXml.documentElement;
  const paragraphs = all(root, 'p');
  let paragraph = paragraphs.find(item =>
    all(item, 'instrText').some(text => /\bPAGE\b/i.test(text.textContent)) ||
    all(item, 'fldSimple').some(field => /\bPAGE\b/i.test(attr(field, 'instr') || ''))
  );
  if (!paragraph) {
    paragraph = footerXml.createElementNS(W, 'w:p');
    root.appendChild(paragraph);
    const field = footerXml.createElementNS(W, 'w:fldSimple');
    set(field, 'instr', 'PAGE');
    const run = footerXml.createElementNS(W, 'w:r');
    const text = footerXml.createElementNS(W, 'w:t');
    text.textContent = '1';
    run.appendChild(text);
    field.appendChild(run);
    paragraph.appendChild(field);
    fixRunProperties(child(run, 'rPr'));
    fixParagraphProperties(child(paragraph, 'pPr'));
  }
  set(child(child(paragraph, 'pPr'), 'jc'), 'val', position === 'bottom-center' ? 'center' : 'right');
}
async function applyPageNumberPosition(zip, documentXml, position) {
  if (position === 'keep') return;
  if (!['bottom-center', 'bottom-right'].includes(position)) throw new Error('Posisi nomor halaman tidak dikenal.');
  const relsPath = 'word/_rels/document.xml.rels';
  const relsFile = zip.file(relsPath);
  const relsXml = relsFile
    ? parseXml(await relsFile.async('string'), relsPath)
    : parseXml(`<Relationships xmlns="${REL}"/>`, relsPath);
  const contentPath = '[Content_Types].xml';
  const contentFile = zip.file(contentPath);
  if (!contentFile) throw new Error('Daftar bagian DOCX tidak ditemukan.');
  const contentXml = parseXml(await contentFile.async('string'), contentPath);
  const settingsFile = zip.file('word/settings.xml');
  const settingsXml = settingsFile ? parseXml(await settingsFile.async('string'), 'word/settings.xml') : null;
  const evenPages = !!first(settingsXml, 'evenAndOddHeaders');
  const edited = new Map();
  const sections = all(documentXml, 'sectPr');
  if (!sections.length) {
    const body = first(documentXml, 'body');
    const section = documentXml.createElementNS(W, 'w:sectPr');
    body.appendChild(section);
    sections.push(section);
  }
  const newFooter = async () => {
    let index = 1, path;
    do { path = 'word/footer' + index++ + '.xml'; } while (zip.file(path));
    const id = nextId(relsXml);
    const relation = relsXml.createElementNS(REL, 'Relationship');
    relation.setAttribute('Id', id);
    relation.setAttribute('Type', FOOTER_REL);
    relation.setAttribute('Target', path.slice(5));
    relsXml.documentElement.appendChild(relation);
    const override = contentXml.createElementNS(CT, 'Override');
    override.setAttribute('PartName', '/' + path);
    override.setAttribute('ContentType', FOOTER_CONTENT);
    contentXml.documentElement.appendChild(override);
    const footer = parseXml(`<w:ftr xmlns:w="${W}"/>`, path);
    placePageField(footer, position);
    zip.file(path, xmlString(footer));
    edited.set(path, id);
    return id;
  };
  for (const section of sections) {
    const types = ['default'];
    if (direct(section, 'titlePg') || footerReference(section, 'first')) types.push('first');
    if (evenPages || footerReference(section, 'even')) types.push('even');
    for (const type of types) {
      const reference = footerReference(section, type);
      const currentId = reference?.getAttributeNS(R, 'id') || reference?.getAttribute('r:id');
      const relation = currentId && footerRelationship(relsXml, currentId);
      const path = relation ? packagePartPath(relation.getAttribute('Target')) : null;
      if (path && zip.file(path)) {
        if (!edited.has(path)) {
          const footer = parseXml(await zip.file(path).async('string'), path);
          placePageField(footer, position);
          zip.file(path, xmlString(footer));
          edited.set(path, currentId);
        }
        attachFooter(section, type, currentId);
      } else {
        const id = await newFooter();
        attachFooter(section, type, id);
      }
    }
  }
  zip.file(relsPath, xmlString(relsXml));
  zip.file(contentPath, xmlString(contentXml));
}
export async function loadDocx(file) {
  if (!file?.name?.toLowerCase().endsWith('.docx')) throw new Error('Pilih file dengan ekstensi .docx.');
  if (file.size > 25 * 1024 * 1024) throw new Error('Ukuran file maksimal 25 MB.');
  if (file.size < 100) throw new Error('File DOCX terlalu kecil atau rusak.');
  if (!globalThis.JSZip) throw new Error('Mesin pembaca DOCX belum tersedia. Muat ulang halaman.');
  let zip;
  try { zip = await globalThis.JSZip.loadAsync(file); }
  catch { throw new Error('File tidak dapat dibaca sebagai DOCX yang valid.'); }
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) throw new Error('Isi dokumen Word tidak ditemukan. Pastikan file .docx valid.');
  const doc = parseXml(await documentFile.async('string'), 'word/document.xml');
  const stylesFile = zip.file('word/styles.xml');
  const styles = stylesFile ? parseXml(await stylesFile.async('string'), 'word/styles.xml') : null;
  return { zip, doc, styles, file, report: analyzeXml(doc, styles) };
}
export async function repairDocx(state, options = {}) {
  const { zip, doc, styles } = state;
  fixContent(doc, true);
  zip.file('word/document.xml', xmlString(doc));
  if (styles) { fixStyles(styles); zip.file('word/styles.xml', xmlString(styles)); }
  for (const path of Object.keys(zip.files)) {
    if (!/^word\/(header|footer|footnotes|endnotes)\d*\.xml$/.test(path)) continue;
    const content = parseXml(await zip.file(path).async('string'), path);
    fixContent(content);
    zip.file(path, xmlString(content));
  }
  await applyPageNumberPosition(zip, doc, options.pageNumberPosition || 'keep');
  zip.file('word/document.xml', xmlString(doc));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
export async function demoDocx() {
  const zip = new globalThis.JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
  zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.folder('word').folder('_rels').file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/></w:rPr><w:t>PROPOSAL PENELITIAN</w:t></w:r></w:p><w:p><w:r><w:t>Analisis Pengaruh Tata Letak Dokumen terhadap Produktivitas Akademik</w:t></w:r></w:p><w:p><w:r><w:t>Dokumen contoh ini sengaja memakai margin, font, ukuran huruf, dan spasi yang belum sesuai dengan profil format.</w:t></w:r></w:p><w:p><w:r><w:t>Dengan Format-Guard, Anda dapat melihat temuan dan mengunduh versi dokumen yang telah dirapikan.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  return new File([blob], 'Contoh_Proposal.docx', { type: blob.type });
}
