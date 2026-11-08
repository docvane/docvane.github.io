// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for docvane.
// The page frame shared by every page of the site: head (with the Content-Security-Policy), header
// with navigation, optional table of contents, and footer.
import { escapeHtml } from './markdown.mjs';

/** The address the site is published at (GitHub Pages of the organization docvane). */
export const SITE = 'https://docvane.github.io/';

/** The repository of the docvane packages. */
export const REPO = 'https://github.com/DevXord/docvane';

/**
 * The canonical address of a page from its path in docs/: "index.html" is left out, so a folder has
 * one address ("guide/index.html" is "https://docvane.github.io/guide/").
 */
export function canonicalUrl(path) {
  if (path.startsWith('/') || path.includes('..')) throw new Error(`not a path inside the site: ${path}`);
  return SITE + path.replace(/(^|\/)index\.html$/, '$1');
}

/** Content-Security-Policy of every page: nothing is loaded from other origins. */
export const CSP = "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'";

/**
 * Pages that moved: the old path still gets a short page that links to the new one, with the new
 * page's canonical address, and is left out of the sitemap. None yet.
 */
export const MOVED = new Map();

const NAV = [
  ['guide/index.html', 'Guide', 'guide'],
  ['guide/syntax.html', 'Syntax', 'syntax'],
  ['demo/index.html', 'Demo', 'demo'],
  ['changelog.html', 'Changelog', 'changelog'],
];

/**
 * A complete page. `path` is its path in docs/ (for the canonical address, unless `canonicalPath`
 * names another page, as for a page that moved); `root` is the relative path to the site root ("" or "../"); `active` marks the
 * navigation entry; `toc` is a list of { level, text, id } shown beside the content. `source`
 * says where the pages come from: { npm: true } for the published packages, or { commit } for
 * packages built from a commit before they were published.
 */
export function page({ title, description, path, root, active = '', body, toc = [], version, source = { npm: true }, scripts = [], wide = false, canonicalPath = path }) {
  const url = canonicalUrl(canonicalPath);
  const nav = NAV.map(([href, label, key]) => `<a href="${root}${href}"${key === active ? ' aria-current="page"' : ''}>${label}</a>`).join('');
  const contents =
    toc.length > 0
      ? `<nav class="toc" aria-label="On this page"><p class="toc-title">On this page</p><ul>${toc
          .map((h) => `<li class="toc-${String(h.level)}"><a href="#${escapeHtml(h.id)}">${escapeHtml(h.text)}</a></li>`)
          .join('')}</ul></nav>`
      : '';
  const origin = source.npm ? 'Pages built from the packages published on npm.' : `Pages built from the packages of commit ${escapeHtml(source.commit.slice(0, 7))}, before their release on npm.`;
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="${CSP}">
<title>${escapeHtml(title)}</title>
<meta name="description" content="${escapeHtml(description)}">
<link rel="canonical" href="${escapeHtml(url)}">
<meta property="og:type" content="website">
<meta property="og:site_name" content="docvane">
<meta property="og:title" content="${escapeHtml(title)}">
<meta property="og:description" content="${escapeHtml(description)}">
<meta property="og:url" content="${escapeHtml(url)}">
<link rel="icon" href="${root}favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="${root}style.css">
${scripts.map((s) => `<script type="module" src="${s}"></script>`).join('\n')}
</head>
<body>
<a class="skip" href="#main">Skip to content</a>
<header class="site-header">
  <div class="bar">
    <a class="brand" href="${root}index.html"><span class="logo" aria-hidden="true"></span>docvane</a>
    <span class="version" title="${source.npm ? 'Version on npm' : 'Version, not yet on npm'}">${escapeHtml(version)}</span>
    <nav class="main-nav" aria-label="Main">${nav}<a href="${REPO}">GitHub</a><a href="https://www.npmjs.com/package/docvane">npm</a></nav>
  </div>
</header>
<div class="layout${toc.length > 0 ? ' with-toc' : ''}${wide ? ' wide' : ''}">
${contents}
<main id="main">
${body}
</main>
</div>
<footer class="site-footer">
  <p>docvane ${escapeHtml(version)}, MIT License. ${origin} <a href="${REPO}">Source on GitHub</a> · <a href="${REPO}/blob/main/SECURITY.md">Reporting a vulnerability</a></p>
</footer>
</body>
</html>
`;
}
