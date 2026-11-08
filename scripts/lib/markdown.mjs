// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for tplvane.
// A small Markdown to HTML converter for the documentation of tplvane. It supports what those files
// use: ATX headings with GitHub-style anchors, paragraphs, bullet and numbered lists (nested by
// indentation), pipe tables, fenced code blocks, thematic breaks, inline code, **bold**, links and
// backslash escapes. Everything is escaped: raw HTML in the source is shown as text, and links
// other than http(s), mailto, relative paths and #anchors are dropped.

/** Escapes text for HTML element content and attribute values. */
export function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

/**
 * The anchor GitHub gives a heading: lower case, characters other than letters, digits, spaces,
 * hyphens and underscores removed, spaces turned into hyphens; repeated anchors get -1, -2, ….
 */
export function slugger() {
  const seen = new Map();
  return (text) => {
    const base = text
      .toLowerCase()
      .replace(/[^\p{L}\p{N}\p{M} _-]/gu, '')
      .replace(/ /g, '-');
    const n = seen.get(base) ?? 0;
    seen.set(base, n + 1);
    return n === 0 ? base : `${base}-${String(n)}`;
  };
}

/** Plain text of inline Markdown, for anchors and the table of contents. */
export function plainText(md) {
  return md
    .replace(/\\([\\`*_[\]()#+\-.!|<>])/g, '$1')
    .replace(/`([^`]*)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1');
}

const SAFE_URL = /^(?:https?:\/\/|mailto:|#|\.{0,2}\/|[\w.-]+(?:[/#?]|$))/i;

/**
 * Inline Markdown to HTML. `link(url)` maps a link target (for example a repository path to a page
 * of the site) and returns the URL to use.
 */
export function inline(md, link = (u) => u) {
  let out = '';
  let i = 0;
  while (i < md.length) {
    const c = md[i];
    if (c === '\\' && i + 1 < md.length && /[\\`*_[\]()#+\-.!|<>{}]/.test(md[i + 1])) {
      out += escapeHtml(md[i + 1]);
      i += 2;
    } else if (c === '`') {
      let ticks = 1;
      while (md[i + ticks] === '`') ticks++;
      const fence = '`'.repeat(ticks);
      const end = md.indexOf(fence, i + ticks);
      if (end === -1) {
        out += escapeHtml(fence);
        i += ticks;
      } else {
        let code = md.slice(i + ticks, end);
        if (code.startsWith(' ') && code.endsWith(' ') && code.trim() !== '') code = code.slice(1, -1);
        out += `<code>${escapeHtml(code)}</code>`;
        i = end + ticks;
      }
    } else if (c === '*' && md[i + 1] === '*') {
      const end = md.indexOf('**', i + 2);
      if (end > i + 2) {
        out += `<strong>${inline(md.slice(i + 2, end), link)}</strong>`;
        i = end + 2;
      } else {
        out += '**';
        i += 2;
      }
    } else if (c === '[') {
      const m = /^\[((?:\\.|`[^`]*`|[^\]\\])*)\]\(([^)\s]*)\)/.exec(md.slice(i));
      if (m) {
        const target = link(m[2]);
        const text = inline(m[1], link);
        out += target !== null && SAFE_URL.test(target) ? `<a href="${escapeHtml(target)}">${text}</a>` : text;
        i += m[0].length;
      } else {
        out += '[';
        i++;
      }
    } else {
      out += escapeHtml(c);
      i++;
    }
  }
  return out;
}

function splitRow(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
  const cells = [];
  let cell = '';
  let code = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\\' && s[i + 1] === '|') {
      cell += '\\|';
      i++;
    } else if (ch === '`') {
      let n = 1;
      while (s[i + n] === '`') n++;
      code = code === n ? 0 : code === 0 ? n : code;
      cell += '`'.repeat(n);
      i += n - 1;
    } else if (ch === '|' && code === 0) {
      cells.push(cell.trim());
      cell = '';
    } else cell += ch;
  }
  cells.push(cell.trim());
  // In a table cell "\|" stands for "|", also inside code spans.
  return cells.map((c) => c.replace(/`[^`]*`/g, (span) => span.replace(/\\\|/g, '|')));
}

const LIST_ITEM = /^( *)([-*]|\d+\.) +(.*)$/;

/**
 * Converts a Markdown document. Returns the HTML, the headings (level, text, id) for a table of
 * contents, and the title (the first level-1 heading).
 */
export function markdownToHtml(source, { link = (u) => u } = {}) {
  let text = source.replace(/\r\n?/g, '\n');
  // YAML front matter (SKILL.md) is not part of the document.
  if (text.startsWith('---\n')) {
    const end = text.indexOf('\n---\n', 4);
    if (end !== -1) text = text.slice(end + 5);
  }
  const lines = text.split('\n');
  const slug = slugger();
  const headings = [];
  let title = null;
  const out = [];
  let i = 0;

  const isBlockStart = (line) => /^(#{1,6} |`{3,}|\s*\|)/.test(line) || LIST_ITEM.test(line) || /^(---|\*\*\*)\s*$/.test(line);

  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === '') {
      i++;
      continue;
    }
    const fence = /^(`{3,})([^`\s]*)\s*$/.exec(line);
    if (fence) {
      const body = [];
      const close = new RegExp(`^\`{${String(fence[1].length)},}\\s*$`);
      i++;
      while (i < lines.length && !close.test(lines[i])) body.push(lines[i++]);
      i++;
      const lang = fence[2] ? ` class="language-${escapeHtml(fence[2])}"` : '';
      out.push(`<pre><code${lang}>${escapeHtml(body.join('\n'))}</code></pre>`);
      continue;
    }
    const heading = /^(#{1,6}) +(.*?)(?:\s+#+)?\s*$/.exec(line);
    if (heading) {
      const level = heading[1].length;
      const id = slug(plainText(heading[2]));
      if (level === 1 && title === null) title = plainText(heading[2]);
      headings.push({ level, text: plainText(heading[2]), id });
      out.push(`<h${level} id="${escapeHtml(id)}"><a class="anchor" href="#${escapeHtml(id)}" aria-hidden="true">#</a>${inline(heading[2], link)}</h${level}>`);
      i++;
      continue;
    }
    if (/^(---|\*\*\*)\s*$/.test(line)) {
      out.push('<hr>');
      i++;
      continue;
    }
    if (/^\s*\|/.test(line) && i + 1 < lines.length && /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$/.test(lines[i + 1])) {
      const head = splitRow(line);
      const align = splitRow(lines[i + 1]).map((c) => (c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : null));
      i += 2;
      const rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) rows.push(splitRow(lines[i++]));
      // A class, not a style attribute: the pages' CSP allows no inline styles.
      const cell = (tag, c, k) => `<${tag}${align[k] ? ` class="align-${align[k]}"` : ''}>${inline(c, link)}</${tag}>`;
      out.push(
        `<div class="table"><table><thead><tr>${head.map((c, k) => cell('th', c, k)).join('')}</tr></thead><tbody>${rows
          .map((r) => `<tr>${head.map((_, k) => cell('td', r[k] ?? '', k)).join('')}</tr>`)
          .join('')}</tbody></table></div>`,
      );
      continue;
    }
    if (LIST_ITEM.test(line)) {
      const items = [];
      while (i < lines.length) {
        const m = LIST_ITEM.exec(lines[i]);
        // A list of the other kind (bullets after numbers, or the reverse) at the same level starts a new list.
        if (m && items.length > 0 && m[1].length <= items[0].indent && /\d/.test(m[2]) !== items[0].ordered) break;
        if (m) {
          items.push({ indent: m[1].length, ordered: /\d/.test(m[2]), text: m[3] });
          i++;
        } else if (lines[i].trim() !== '' && /^ +\S/.test(lines[i]) && items.length > 0) {
          items[items.length - 1].text += ` ${lines[i].trim()}`;
          i++;
        } else if (lines[i].trim() === '' && i + 1 < lines.length && LIST_ITEM.test(lines[i + 1]) && !(LIST_ITEM.exec(lines[i + 1])[1].length <= items[0].indent && /\d/.test(LIST_ITEM.exec(lines[i + 1])[2]) !== items[0].ordered)) i++;
        else break;
      }
      out.push(renderList(items, 0, items.length, link));
      continue;
    }
    const para = [];
    while (i < lines.length && lines[i].trim() !== '' && (para.length === 0 || !isBlockStart(lines[i]))) para.push(lines[i++].trim());
    out.push(`<p>${inline(para.join(' '), link)}</p>`);
  }
  return { html: out.join('\n'), headings, title };
}

function renderList(items, from, to, link) {
  const base = items[from].indent;
  const tag = items[from].ordered ? 'ol' : 'ul';
  let html = `<${tag}>`;
  let k = from;
  while (k < to) {
    const item = items[k];
    let end = k + 1;
    while (end < to && items[end].indent > base) end++;
    html += `<li>${inline(item.text, link)}${end > k + 1 ? renderList(items, k + 1, end, link) : ''}</li>`;
    k = end;
  }
  return `${html}</${tag}>`;
}
