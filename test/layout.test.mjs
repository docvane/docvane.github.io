// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for docvane.
// node --test: the page frame (canonical address, Open Graph tags, where the pages come from).
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { canonicalUrl, page } from '../scripts/lib/layout.mjs';

test('canonical addresses leave out index.html', () => {
  assert.equal(canonicalUrl('index.html'), 'https://docvane.github.io/');
  assert.equal(canonicalUrl('guide/index.html'), 'https://docvane.github.io/guide/');
  assert.equal(canonicalUrl('guide/mcp.html'), 'https://docvane.github.io/guide/mcp.html');
  assert.equal(canonicalUrl('sitemap.xml'), 'https://docvane.github.io/sitemap.xml');
  assert.throws(() => canonicalUrl('../x.html'));
  assert.throws(() => canonicalUrl('/x.html'));
});

test('every page has its canonical link and Open Graph tags, escaped', () => {
  const html = page({ title: 'A "b" & c', description: 'd <e>', path: 'guide/mcp.html', root: '../', version: '1.0.0', body: '' });
  assert.match(html, /<link rel="canonical" href="https:\/\/docvane\.github\.io\/guide\/mcp\.html">/);
  assert.match(html, /<meta property="og:url" content="https:\/\/docvane\.github\.io\/guide\/mcp\.html">/);
  assert.match(html, /<meta property="og:title" content="A &quot;b&quot; &amp; c">/);
  assert.match(html, /<meta property="og:description" content="d &lt;e&gt;">/);
  assert.match(html, /<meta property="og:site_name" content="docvane">/);
});

test('pages built before the release on npm say so', () => {
  const npm = page({ title: 't', description: 'd', path: 'index.html', root: '', version: '1.0.0', body: '' });
  assert.match(npm, /Pages built from the packages published on npm\./);
  assert.match(npm, /title="Version on npm"/);
  const early = page({ title: 't', description: 'd', path: 'index.html', root: '', version: '1.0.0', body: '', source: { commit: 'b0d17da6825852e51e27bcadb29e42a4fe8e78bd' } });
  assert.match(early, /Pages built from the packages of commit b0d17da, before their release on npm\./);
  assert.doesNotMatch(early, /published on npm/);
});
