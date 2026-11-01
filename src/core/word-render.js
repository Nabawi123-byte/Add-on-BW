// Mengubah daftar blok menjadi HTML yang disisipkan ke Word (insertHtml).
// Blok yang sama dipakai oleh template offline maupun AI.

import { escapeHtml } from './util.js';

const FONT = "font-family:'Times New Roman',serif;font-size:12pt;";

/** Teks sederhana: **tebal**, *miring*, __garis bawah__, baris baru -> <br>. */
export function inline(text) {
  return escapeHtml(text)
    .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
    .replace(/__(.+?)__/g, '<u>$1</u>')
    .replace(/(^|[^*])\*(?!\s)(.+?)\*(?!\*)/g, '$1<i>$2</i>')
    .replace(/\r?\n/g, '<br>');
}

const align = (a) => (a ? `text-align:${a === 'justify' ? 'justify' : a};` : '');

function paragraph(b) {
  let style = FONT + align(b.align) + 'margin:0 0 8pt 0;line-height:1.3;';
  if (b.size) style += `font-size:${Number(b.size)}pt;`;
  if (b.bold) style += 'font-weight:bold;';
  if (b.italic) style += 'font-style:italic;';
  if (b.color) style += `color:${escapeHtml(b.color)};`;
  if (b.indent) style += 'text-indent:36pt;';
  if (b.border) style += 'border:1px solid #000;padding:4pt;';
  return `<p style="${style}">${inline(b.text)}</p>`;
}

function heading(b) {
  const level = Math.min(Math.max(Number(b.level) || 1, 1), 3);
  const size = { 1: 16, 2: 14, 3: 12 }[level];
  let style = FONT + `font-size:${size}pt;font-weight:bold;margin:6pt 0 6pt 0;` + align(b.align);
  if (b.underline) style += 'text-decoration:underline;';
  return `<h${level} style="${style}">${inline(b.text)}</h${level}>`;
}

function list(b) {
  const tag = b.ordered ? 'ol' : 'ul';
  const items = (b.items || []).map((it) => `<li style="${FONT}margin:0 0 4pt 0;">${inline(it)}</li>`).join('');
  return items ? `<${tag}>${items}</${tag}>` : '';
}

function table(b) {
  const rows = b.rows || [];
  if (!rows.length) return '';
  const cell = 'border:1px solid #000;padding:3pt 5pt;vertical-align:top;';
  const html = rows
    .map((r, i) => {
      const isHead = b.header && i === 0;
      const tag = isHead ? 'th' : 'td';
      const extra = isHead ? 'background:#D9E1F2;font-weight:bold;text-align:center;' : '';
      return `<tr>${r.map((c) => `<${tag} style="${FONT}${cell}${extra}">${inline(c)}</${tag}>`).join('')}</tr>`;
    })
    .join('');
  return `<table style="border-collapse:collapse;width:100%;margin:0 0 8pt 0;">${html}</table>`;
}

function fields(b) {
  const rows = (b.rows || [])
    .map(
      ([k, v]) =>
        `<tr><td style="${FONT}padding:0 6pt 2pt 0;width:150pt;vertical-align:top;">${inline(k)}</td><td style="${FONT}padding:0 6pt 2pt 0;width:10pt;vertical-align:top;">:</td><td style="${FONT}padding:0 0 2pt 0;vertical-align:top;">${inline(v === undefined || v === null || v === '' ? '..............................' : v)}</td></tr>`,
    )
    .join('');
  return `<table style="border-collapse:collapse;margin:0 0 8pt ${b.indent === false ? 0 : 24}pt;">${rows}</table>`;
}

function signatureCell(s, alignText = 'center') {
  const lines = [];
  if (s.place) lines.push(inline(s.place));
  lines.push(inline(s.title || ''));
  if (s.note) lines.push(`<span style="font-size:9pt;color:#7F7F7F;">${inline(s.note)}</span>`);
  else lines.push('<br>');
  lines.push('<br><br>');
  lines.push(`<b><u>${inline(s.name || '(..............................)')}</u></b>`);
  if (s.position) lines.push(inline(s.position));
  return `<td style="${FONT}text-align:${alignText};vertical-align:top;width:50%;padding-top:12pt;">${lines.join('<br>')}</td>`;
}

function signature(b) {
  const right = b.align !== 'left';
  const empty = `<td style="width:50%;"></td>`;
  const cellHtml = signatureCell(b);
  return `<table style="width:100%;border-collapse:collapse;margin-top:12pt;"><tr>${right ? empty + cellHtml : cellHtml + empty}</tr></table>`;
}

function signatures(b) {
  const pre = b.place ? `<p style="${FONT}text-align:right;margin:12pt 0 0 0;">${inline(b.place)}</p>` : '';
  const cells = (b.items || []).map((s) => signatureCell(s)).join('');
  return `${pre}<table style="width:100%;border-collapse:collapse;margin-top:12pt;"><tr>${cells}</tr></table>`;
}

export function renderBlock(b) {
  switch (b.type) {
    case 'heading':
      return heading(b);
    case 'paragraph':
      return paragraph(b);
    case 'list':
      return list(b);
    case 'table':
      return table(b);
    case 'fields':
      return fields(b);
    case 'signature':
      return signature(b);
    case 'signatures':
      return signatures(b);
    case 'line':
      return '<hr style="border:none;border-top:2px solid #000;margin:2pt 0 10pt 0;">';
    case 'spacer':
      return `<p style="${FONT}margin:0;">&nbsp;</p>`;
    case 'html':
      return String(b.html || '');
    default:
      throw new Error(`Jenis blok tidak dikenal: ${b.type}`);
  }
}

/** Memecah blok menjadi segmen HTML; pageBreak memisahkan segmen. */
export function renderSegments(blocks) {
  const segments = [[]];
  for (const b of blocks || []) {
    if (b.type === 'pageBreak') segments.push([]);
    else segments[segments.length - 1].push(renderBlock(b));
  }
  return segments.map((s) => s.join(''));
}

/** Teks polos dari blok (untuk pratinjau dan tes). */
export function blocksToText(blocks) {
  return (blocks || [])
    .map((b) => {
      if (b.text) return b.text;
      if (b.items) return b.items.join('\n');
      if (b.rows) return b.rows.map((r) => r.join(' | ')).join('\n');
      if (b.type === 'signature') return `${b.title} ${b.name}`;
      return '';
    })
    .join('\n');
}
