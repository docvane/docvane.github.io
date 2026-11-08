// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for docvane.
// node --test: the Markdown converter used for the documentation pages.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { escapeHtml, inline, markdownToHtml, slugger } from '../scripts/lib/markdown.mjs';

test('escapes raw HTML and attribute characters', () => {
  assert.equal(inline('<script>alert(1)</script> & "q"'), '&lt;script&gt;alert(1)&lt;/script&gt; &amp; &quot;q&quot;');
  assert.equal(escapeHtml(`'`), '&#39;');
  assert.equal(markdownToHtml('<img src=x onerror=alert(1)>').html, '<p>&lt;img src=x onerror=alert(1)&gt;</p>');
});

test('drops links with unsafe schemes and keeps their text', () => {
  assert.equal(inline('[x](javascript:void)'), 'x');
  assert.equal(inline('[x](data:text/html,1)'), 'x');
  assert.equal(inline('[x](https://example.com/a?b=1&c=2)'), '<a href="https://example.com/a?b=1&amp;c=2">x</a>');
  assert.equal(inline('[x](docs/errors.md#a)', () => 'errors.html#a'), '<a href="errors.html#a">x</a>');
  assert.equal(inline('[x](../outside.md)', () => null), 'x');
});

test('code spans keep their content as text', () => {
  assert.equal(inline('`a **b** <c>`'), '<code>a **b** &lt;c&gt;</code>');
  assert.equal(inline('``a ` b``'), '<code>a ` b</code>');
  assert.equal(inline('_x001B_ and `DV_E_*`'), '_x001B_ and <code>DV_E_*</code>');
});

test('headings get GitHub anchors, unique per document', () => {
  const slug = slugger();
  assert.equal(slug('Command line in 5 minutes'), 'command-line-in-5-minutes');
  assert.equal(slug('DV_E_FILE_EXISTS'), 'dv_e_file_exists');
  assert.equal(slug('name, id, state'), 'name-id-state');
  assert.equal(slug('Large files: streaming'), 'large-files-streaming');
  assert.equal(slug('Large files: streaming'), 'large-files-streaming-1');
  const { html, headings, title } = markdownToHtml('# Title\n\n## Tasks in one call\n\ntext');
  assert.equal(title, 'Title');
  assert.deepEqual(headings.map((h) => h.id), ['title', 'tasks-in-one-call']);
  assert.match(html, /<h2 id="tasks-in-one-call">/);
});

test('tables, with escaped pipes and pipes inside code', () => {
  const { html } = markdownToHtml('| A | B |\n|---|---:|\n| `a \\| b` | 2 |\n| x \\| y | 3 |');
  assert.match(html, /<th>A<\/th><th class="align-right">B<\/th>/);
  assert.match(html, /<td><code>a \| b<\/code><\/td><td class="align-right">2<\/td>/);
  assert.match(html, /<td>x \| y<\/td>/);
});

test('a heading may end in # (C#), and fences may be longer than three backticks', () => {
  assert.match(markdownToHtml('## C#').html, /C#<\/h2>/);
  assert.match(markdownToHtml('## Title ##').html, /Title<\/h2>/);
  const { html } = markdownToHtml('````md\n```js\nx\n```\n````\nafter');
  assert.match(html, /<pre><code class="language-md">```js\nx\n```<\/code><\/pre>/);
  assert.match(html, /<p>after<\/p>/);
});

test('nested and numbered lists, fenced code and front matter', () => {
  const { html } = markdownToHtml('---\nname: x\n---\n# T\n\n- a\n  - b\n- c\n\n1. one\n2. two\n\n```ts\nconst a = 1 < 2;\n```');
  assert.match(html, /<ul><li>a<ul><li>b<\/li><\/ul><\/li><li>c<\/li><\/ul>/);
  assert.match(html, /<ol><li>one<\/li><li>two<\/li><\/ol>/);
  assert.match(html, /<pre><code class="language-ts">const a = 1 &lt; 2;<\/code><\/pre>/);
  assert.doesNotMatch(html, /name: x/);
});
