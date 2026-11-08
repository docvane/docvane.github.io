// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for docvane.
// Builds the site into docs/ (published by GitHub Pages from the main branch, /docs folder) from
// the docvane packages and the offer example of the repository:
//   node scripts/build.mjs --version 1.0.0                      # packages published on npm, example of tag v1.0.0
//   node scripts/build.mjs --version 1.0.0 --tarballs DIR --commit SHA
// The second form is for a version not yet on npm: DIR holds docvane-X.Y.Z.tgz and
// docvane-mcp-X.Y.Z.tgz packed from commit SHA, and the pages say that they come before the release.
// The documentation pages are the Markdown files of the docvane and @docvane/mcp packages, so the
// site shows exactly the documentation of that version. The demo uses the package's own browser
// code (dist/) with xlvane's and fflate's browser code, copied into docs/vendor/.
import { createHash } from 'node:crypto';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { MOVED, REPO, canonicalUrl, page } from './lib/layout.mjs';
import { escapeHtml, markdownToHtml } from './lib/markdown.mjs';
import { fetchPackage, readTarball } from './lib/npm.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const OUT = join(ROOT, 'docs');

const arg = (name) => {
  const i = process.argv.indexOf(`--${name}`);
  return i === -1 ? undefined : process.argv[i + 1];
};
const version = arg('version');
const tarballs = arg('tarballs');
const commit = arg('commit');
const accept = process.argv.includes('--accept');
/** `--accept-dependency name@version` (repeatable): a dependency version accepted by name, never by --accept. */
const acceptedDependencies = new Set(process.argv.flatMap((a, i) => (process.argv[i - 1] === '--accept-dependency' ? [a] : [])));
if (!version || !/^\d+\.\d+\.\d+$/.test(version) || (tarballs === undefined) !== (commit === undefined) || (commit !== undefined && !/^[0-9a-f]{40}$/.test(commit))) {
  console.error('usage: node scripts/build.mjs --version X.Y.Z [--tarballs DIR --commit <40-character commit id>] [--accept] [--accept-dependency name@X.Y.Z]');
  process.exit(2);
}
/** The git ref the pages link to and the example is taken from: the release tag, or the commit the tarballs were packed from. */
const ref = commit ?? `v${version}`;
const source = commit === undefined ? { npm: true } : { commit };

const LOCK = join(ROOT, 'vendor.lock.json');
const lock = existsSync(LOCK) ? JSON.parse(readFileSync(LOCK, 'utf8')) : {};
/**
 * Every package version, file of the example and its sha512 must be in vendor.lock.json, so that a
 * new or changed file never reaches the site without a deliberate acceptance after checking it.
 * `--accept` covers only the docvane packages (check their npm provenance) and the files of the
 * example; a dependency such as xlvane or fflate is accepted by name and version with
 * `--accept-dependency name@X.Y.Z`, so a new release within docvane's range is never taken in
 * along with a new docvane version. Packages packed before the release are locked by version and
 * commit, apart from the published ones.
 */
function locked(name, ver, integrity, own) {
  const known = lock[name]?.[ver];
  if (known === integrity) return;
  if (known !== undefined) throw new Error(`${name}@${ver}: the source now gives ${integrity}, vendor.lock.json has ${known}. Do not accept this without finding out why.`);
  const allowed = own ? accept : acceptedDependencies.has(`${name}@${ver}`);
  if (!allowed) {
    const how = own ? '--accept' : `--accept-dependency ${name}@${ver}`;
    throw new Error(`${name}@${ver} (${integrity}) is not in vendor.lock.json. Check it, then run again with ${how} to add it.`);
  }
  lock[name] = { ...(lock[name] ?? {}), [ver]: integrity };
  console.log(`added to vendor.lock.json: ${name}@${ver} ${integrity}`);
}
const sha512 = (bytes) => `sha512-${createHash('sha512').update(bytes).digest('base64')}`;

/** Repository paths of the documents, the package file each comes from, and the page each becomes. */
const PAGES = [
  { repo: 'packages/docvane/README.md', pkg: 'docvane', file: 'README.md', out: 'guide/index.html', title: 'Guide', active: 'guide' },
  { repo: 'docs/syntax.md', pkg: 'docvane', file: 'docs/syntax.md', out: 'guide/syntax.html', title: 'Template syntax', active: 'syntax' },
  { repo: 'docs/errors.md', pkg: 'docvane', file: 'docs/errors.md', out: 'guide/errors.html', title: 'Errors and warnings', active: 'guide' },
  { repo: 'docs/migrating-from-docxtemplater.md', pkg: 'docvane', file: 'docs/migrating-from-docxtemplater.md', out: 'guide/migrating-from-docxtemplater.html', title: 'Migrating from docxtemplater', active: 'guide' },
  { repo: 'packages/docvane/docs/cli-json.md', pkg: 'docvane', file: 'docs/cli-json.md', out: 'guide/cli-json.html', title: 'Command line JSON', active: 'guide' },
  { repo: 'packages/mcp/README.md', pkg: '@docvane/mcp', file: 'README.md', out: 'guide/mcp.html', title: '@docvane/mcp (MCP server)', active: 'guide' },
  { repo: 'packages/docvane/skills/docvane/SKILL.md', pkg: 'docvane', file: 'skills/docvane/SKILL.md', out: 'guide/agent-skill.html', title: 'Agent Skill', active: 'guide' },
  { repo: 'packages/docvane/CHANGELOG.md', pkg: 'docvane', file: 'CHANGELOG.md', out: 'changelog.html', title: 'Changelog', active: 'changelog' },
];
const PAGE_BY_REPO = new Map(PAGES.map((p) => [p.repo, p.out]));
PAGE_BY_REPO.set('README.md', 'index.html');
/** The JSON Schemas of the command line output: in the repository and in the package, and on the site. */
const SCHEMA = /^packages\/docvane\/docs\/cli-json\/[\w.-]+\.schema\.json$/;
const schemaPage = (repoPath) => `guide/cli-json/${posix.basename(repoPath)}`;

// The docvane packages: published on npm, or packed from a commit before the release.
async function own(name, file) {
  if (tarballs === undefined) {
    const pkg = await fetchPackage(name, version);
    locked(name, version, pkg.integrity, true);
    return pkg;
  }
  const pkg = readTarball(readFileSync(join(tarballs, file)));
  locked(`${name} (packed)`, `${version}+${commit}`, pkg.integrity, true);
  const manifest = JSON.parse(new TextDecoder().decode(pkg.files.get('package.json')));
  if (manifest.name !== name || manifest.version !== version) throw new Error(`${file} holds ${String(manifest.name)}@${String(manifest.version)}, not ${name}@${version}.`);
  return pkg;
}
const core = await own('docvane', `docvane-${version}.tgz`);
const mcp = await own('@docvane/mcp', `docvane-mcp-${version}.tgz`);
const manifestOf = (pkg) => JSON.parse(new TextDecoder().decode(pkg.files.get('package.json')));
// The newest xlvane and fflate that the dependency ranges allow, as npm would install them.
const xlvaneVersion = await newestInRange('xlvane', manifestOf(core).dependencies.xlvane);
const xlvanePkg = await fetchPackage('xlvane', xlvaneVersion);
locked('xlvane', xlvaneVersion, xlvanePkg.integrity, false);
const fflateRange = manifestOf(xlvanePkg).dependencies.fflate;
if (fflateRange !== manifestOf(core).dependencies.fflate) throw new Error(`docvane wants fflate ${String(manifestOf(core).dependencies.fflate)}, xlvane ${String(fflateRange)}: check which one the browser copy must satisfy.`);
const fflateVersion = await newestInRange('fflate', fflateRange);
const fflatePkg = await fetchPackage('fflate', fflateVersion);
locked('fflate', fflateVersion, fflatePkg.integrity, false);
// The offer example of the repository, for the demo. The tag could be moved, so its files are locked like a package.
const EXAMPLE = ['offer.docx', 'data.json', 'logo.png'];
const example = new Map();
for (const name of EXAMPLE) {
  const response = await fetch(`https://raw.githubusercontent.com/DevXord/docvane/${ref}/examples/offer/${name}`);
  if (!response.ok) throw new Error(`examples/offer/${name} of ${ref}: HTTP ${String(response.status)}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  locked(`examples/offer/${name}`, ref, sha512(bytes), true);
  example.set(name, bytes);
}
writeFileSync(LOCK, `${JSON.stringify(lock, null, 2)}\n`);

const text = (files, path) => {
  const data = files.get(path);
  if (!data) throw new Error(`${path} is missing from the package.`);
  return new TextDecoder().decode(data);
};
const FILES = { docvane: core.files, '@docvane/mcp': mcp.files };
const sourceOf = (p) => text(FILES[p.pkg], p.file);

/** A link of a document, mapped to a page of the site (relative to `from`) or to GitHub at the ref. */
function linker(repoPath, from) {
  const rel = (target) => posix.relative(posix.dirname(from), target) || posix.basename(target);
  return (url) => {
    if (url.startsWith('#') || /^(https?:|mailto:)/i.test(url)) {
      const gh = /^https:\/\/github\.com\/DevXord\/docvane\/blob\/main\/([^#?]+)(#.*)?$/.exec(url);
      if (gh && PAGE_BY_REPO.has(gh[1])) return rel(PAGE_BY_REPO.get(gh[1])) + (gh[2] ?? '');
      if (gh && SCHEMA.test(gh[1])) return rel(schemaPage(gh[1]));
      return url;
    }
    const [path, hash = ''] = url.split('#');
    const target = posix.normalize(posix.join(posix.dirname(repoPath), path));
    if (PAGE_BY_REPO.has(target)) return rel(PAGE_BY_REPO.get(target)) + (hash ? `#${hash}` : '');
    if (SCHEMA.test(target)) return rel(schemaPage(target));
    if (target.startsWith('..')) return null;
    return `${REPO}/blob/${ref}/${target}${hash ? `#${hash}` : ''}`;
  };
}

const write = (path, content) => {
  const full = join(OUT, path);
  mkdirSync(dirname(full), { recursive: true });
  writeFileSync(full, content);
};

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
write('.nojekyll', '');
cpSync(join(ROOT, 'src/style.css'), join(OUT, 'style.css'));
cpSync(join(ROOT, 'src/favicon.svg'), join(OUT, 'favicon.svg'));
cpSync(join(ROOT, 'src/social.png'), join(OUT, 'social.png'));

// Documentation pages.
const guideNav = PAGES.filter((p) => p.out.startsWith('guide/'));
for (const p of PAGES) {
  const root = p.out.includes('/') ? '../' : '';
  const { html, headings } = markdownToHtml(sourceOf(p), { link: linker(p.repo, p.out) });
  const sub = p.out.startsWith('guide/')
    ? `<nav class="guide-nav" aria-label="Guide">${guideNav.map((g) => `<a href="${posix.relative('guide', g.out)}"${g === p ? ' aria-current="page"' : ''}>${escapeHtml(g.title)}</a>`).join('')}</nav>`
    : '';
  const sourceLink = `<p class="source">Source: <a href="${REPO}/blob/${ref}/${p.repo}">${escapeHtml(p.repo)}</a> (version ${escapeHtml(version)})</p>`;
  write(
    p.out,
    page({
      title: `${p.title} - docvane`,
      path: p.out,
      description: `${p.title} of docvane ${version}, a template engine for Word, PowerPoint and Excel files.`,
      root,
      active: p.active,
      version,
      source,
      toc: headings.filter((h) => h.level === 2 || h.level === 3),
      body: `${sub}<article class="doc">${html}${sourceLink}</article>`,
    }),
  );
}
for (const [path, data] of core.files) if (/^docs\/cli-json\/[\w.-]+\.schema\.json$/.test(path)) write(`guide/cli-json/${posix.basename(path)}`, data);

// Landing page, with the code of the README.
const readme = text(core.files, 'README.md');
const intro = readme.slice(0, readme.indexOf('\n## '));
const codeBlocks = (md, lang) => [...md.matchAll(/```(\w*)\n([\s\S]*?)\n```/g)].filter((m) => lang === undefined || m[1] === lang).map((m) => m[2]);
const cliStart = readme.indexOf('\n## The command line in 5 minutes\n');
if (cliStart === -1) throw new Error('README section "The command line in 5 minutes" not found.');
const cliSection = readme.slice(cliStart, readme.indexOf('\n## ', cliStart + 4));
const [inspectCode, renderCode] = codeBlocks(intro, 'ts');
if (inspectCode === undefined || renderCode === undefined) throw new Error('the README has fewer than two TypeScript examples before its first section.');
const fill = (template, values) =>
  template.replace(/\{\{(\w+)\}\}/g, (_, k) => {
    if (!(k in values)) throw new Error(`unknown placeholder ${k}`);
    return values[k];
  });
const landing = fill(readFileSync(join(ROOT, 'src/index.html'), 'utf8'), {
  version: escapeHtml(version),
  render: escapeHtml(renderCode),
  inspect: escapeHtml(inspectCode),
  cli: escapeHtml(codeBlocks(cliSection, 'bash').join('\n')),
});
write('index.html', page({ title: 'docvane - fill Word, PowerPoint and Excel templates', path: 'index.html', description: 'docvane fills Word, PowerPoint and Excel templates with one syntax and treats templates and data as untrusted. MIT license, two runtime dependencies.', root: '', version, source, body: landing }));

// Demo: the page, its script, the sample, and the browser code of docvane, xlvane and fflate.
write('demo/index.html', page({ title: 'Demo - docvane', path: 'demo/index.html', description: 'Fill a Word, PowerPoint or Excel template in the browser with docvane; the files do not leave the computer.', root: '../', active: 'demo', version, source, wide: true, scripts: ['demo.js'], body: readFileSync(join(ROOT, 'src/demo/index.html'), 'utf8') }));
cpSync(join(ROOT, 'src/demo/demo.js'), join(OUT, 'demo/demo.js'));
for (const [name, bytes] of example) write(`demo/sample/${name}`, bytes);
/** Copies an entry point and the files it imports, with the bare specifiers rewritten to the copies beside it (no import map, so the CSP needs no inline script). */
function vendor(pkg, name, rewrite) {
  const needed = new Set();
  const queue = ['dist/index.js'];
  while (queue.length > 0) {
    const path = queue.shift();
    if (needed.has(path)) continue;
    needed.add(path);
    for (const m of text(pkg.files, path).matchAll(/from "(\.\/[\w.-]+\.js)"/g)) queue.push(posix.join('dist', m[1]));
  }
  for (const path of needed) {
    let js = text(pkg.files, path);
    for (const [specifier, target] of rewrite) js = js.replaceAll(`from "${specifier}"`, `from "${target}"`);
    // Static imports and re-exports at the start of a line, and dynamic import() of a string (TSDoc examples in comments are not code).
    if (/^\s*(?:import|export)\b[^;'"]*?\bfrom\s*["'](?!\.\.?\/)|^\s*import\s*["'](?!\.\.?\/)|\bimport\s*\(\s*["'](?!\.\.?\/)/m.test(js)) throw new Error(`${name}/${path} imports a module the browser cannot load.`);
    write(`vendor/${name}/${posix.basename(path)}`, js);
  }
  write(`vendor/${name}/LICENSE`, pkg.files.get('LICENSE'));
}
vendor(core, 'docvane', [
  ['xlvane', '../xlvane/index.js'],
  ['fflate', '../fflate/browser.js'],
]);
vendor(xlvanePkg, 'xlvane', [['fflate', '../fflate/browser.js']]);
const fflateEntry = JSON.parse(text(fflatePkg.files, 'package.json')).exports['.'].import.default;
write('vendor/fflate/browser.js', fflatePkg.files.get(fflateEntry.replace(/^\.\//, '')));
write('vendor/fflate/LICENSE', fflatePkg.files.get('LICENSE'));
const origin = commit === undefined ? 'Copied from the packages published on npm' : `docvane packed from commit ${commit} before its release; xlvane and fflate from npm. Copied`;
write(
  'vendor/README.txt',
  `docvane ${version} (${core.integrity}): dist/*.js used by index.js, MIT license.\nxlvane ${xlvaneVersion} (${xlvanePkg.integrity}): dist/*.js used by index.js, MIT license.\nfflate ${fflateVersion} (${fflatePkg.integrity}): ${fflateEntry}, MIT license.\n${origin} by scripts/build.mjs.\n`,
);

// The page GitHub Pages shows for an address that does not exist. It can be served at any depth, so its links start at the root.
write('404.html', page({ title: 'Page not found - docvane', path: '404.html', description: 'This page does not exist on the docvane website.', root: '/', version, source, body: '<article class="doc"><h1>Page not found</h1><p>There is no page at this address. Try the <a href="/index.html">home page</a>, the <a href="/guide/index.html">guide</a>, the <a href="/guide/syntax.html">template syntax</a> or the <a href="/demo/index.html">demo</a>.</p></article>' }));

if (!existsSync(join(OUT, 'index.html'))) throw new Error('index.html was not written');

// For search engines: every page in a sitemap (without lastmod, so the same version gives the same
// files), and a robots.txt that allows everything and names the sitemap.
const sitePages = ['index.html', ...PAGES.map((p) => p.out), 'demo/index.html'];
write('sitemap.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${sitePages.map((p) => `  <url><loc>${escapeHtml(canonicalUrl(p))}</loc></url>\n`).join('')}</urlset>\n`);
write('robots.txt', `User-agent: *\nAllow: /\n\nSitemap: ${canonicalUrl('sitemap.xml')}\n`);
if (MOVED.size > 0) throw new Error('pages that moved are not handled by this build yet.');
console.log(`docs/: docvane ${version}${commit === undefined ? '' : ` (packed from ${commit.slice(0, 7)})`}, @docvane/mcp ${version}, xlvane ${xlvaneVersion}, fflate ${fflateVersion}; ${String(sitePages.length)} pages.`);

/** The newest published version that a caret range (^X.Y.Z) allows; other ranges are refused. */
async function newestInRange(name, range) {
  const m = /^\^(\d+)\.(\d+)\.(\d+)$/.exec(range ?? '');
  if (!m) throw new Error(`${name}: only caret ranges such as ^0.8.3 are handled, not "${String(range)}".`);
  const [major, minor, patch] = m.slice(1).map(Number);
  const meta = await (await fetch(`https://registry.npmjs.org/${name}`)).json();
  const ok = Object.keys(meta.versions)
    .map((v) => /^(\d+)\.(\d+)\.(\d+)$/.exec(v))
    .filter((v) => v !== null)
    .map((v) => v.slice(1).map(Number))
    .filter(([a, b, c]) => a === major && (major > 0 || b === minor) && (a > major || b > minor || (b === minor && c >= patch)))
    .sort((x, y) => x[0] - y[0] || x[1] - y[1] || x[2] - y[2]);
  if (ok.length === 0) throw new Error(`${name}: no version matches ${range}.`);
  return ok[ok.length - 1].join('.');
}
