# docvane.github.io

The website of [docvane](https://github.com/DevXord/docvane): a landing page, the documentation of the published version, and a demo that fills a template in the browser.

GitHub Pages publishes the `docs/` folder of the `main` branch. `docs/` is generated; edit `src/` and `scripts/` and build again.

```sh
node scripts/build.mjs --version 1.0.0   # downloads docvane, @docvane/mcp, xlvane and fflate from npm (integrity checked) and the offer example of tag v1.0.0
node scripts/check.mjs                   # every internal link and anchor exists, nothing is loaded from other origins
node --test test/*.test.mjs              # tests of the Markdown converter and the page frame
```

- No dependencies: the generator, the Markdown converter and the tar reader are in `scripts/`. They are copied from the site of xlvane ([xlvane/xlvane.github.io](https://github.com/xlvane/xlvane.github.io), MIT, commit 9bb6269) and adapted; each copied file says so.
- The documentation pages are the Markdown files of the packages, so they match that version.
- For search engines, the build writes `robots.txt` and `sitemap.xml`, and each page carries its canonical address and Open Graph tags; `check.mjs` checks that they agree.
- The demo uses the browser code of the package (`docs/vendor/docvane/`), of xlvane (`docs/vendor/xlvane/`) and fflate's browser build (`docs/vendor/fflate/`), and the offer template of the repository (`docs/demo/sample/`). Nothing is loaded from other sites, and every page sets a Content-Security-Policy that allows only this site.

## Before a version is on npm

A version can be built from packed packages before its release:

```sh
node scripts/build.mjs --version 1.0.0 --tarballs tarballs --commit <40-character commit id>
```

`tarballs/` (not committed) holds `docvane-1.0.0.tgz` and `docvane-mcp-1.0.0.tgz` made with `npm pack` from that commit of DevXord/docvane. The pages then say that they come before the release on npm, and link to that commit instead of the tag. After the release, build again from npm.

## vendor.lock.json

It holds the sha512 of every package version the site was built from and of each file of the example. A version that is not in it stops the build until it is checked and accepted: a new release of the docvane packages (and the example of its tag) with `--accept`, a newer dependency within docvane's range, such as xlvane or fflate, only by name with `--accept-dependency name@X.Y.Z`. Each accepted entry is printed. A changed sha512 for a known version always stops the build. Packages packed before the release are locked under their own names (`docvane (packed)`), by version and commit.

After each release of docvane: check its provenance on npm, build with `--version X.Y.Z --accept`, check, commit and push.

License: MIT.
