// Copied from xlvane/xlvane.github.io (MIT), commit 9bb6269956ca7f7c051da8d77344e65aac44accd, and adapted for tplvane.
// Downloads a published npm package (or reads a packed tarball) and returns its files, checked against the integrity the
// registry lists. Tarballs are gzip-compressed POSIX tar archives (USTAR); only regular files are
// returned, with the leading "package/" removed.
import { createHash } from 'node:crypto';
import { gunzipSync } from 'node:zlib';

export async function fetchPackage(name, version) {
  const meta = await (await fetch(`https://registry.npmjs.org/${name.replace('/', '%2f')}/${version}`)).json();
  if (!meta.dist?.tarball) throw new Error(`${name}@${version} was not found on the npm registry.`);
  const bytes = new Uint8Array(await (await fetch(meta.dist.tarball)).arrayBuffer());
  const [algo, expected] = meta.dist.integrity.split('-');
  const actual = createHash(algo).update(bytes).digest('base64');
  if (actual !== expected) throw new Error(`${name}@${version}: the tarball does not match the registry's integrity ${meta.dist.integrity}.`);
  return { files: untar(gunzipSync(bytes)), integrity: meta.dist.integrity };
}

/** Regular files of a tar archive, by path. */
export function untar(data) {
  const files = new Map();
  const text = (from, len) => new TextDecoder().decode(data.subarray(from, from + len)).replace(/\0.*$/s, '');
  let pos = 0;
  let longName = null;
  while (pos + 512 <= data.length) {
    if (data.subarray(pos, pos + 512).every((b) => b === 0)) break;
    let name = text(pos, 100);
    const size = parseInt(text(pos + 124, 12).trim() || '0', 8);
    const type = String.fromCharCode(data[pos + 156] || 48);
    const prefix = text(pos + 345, 155);
    if (prefix) name = `${prefix}/${name}`;
    const body = data.subarray(pos + 512, pos + 512 + size);
    if (type === 'L') longName = new TextDecoder().decode(body).replace(/\0.*$/s, '');
    else {
      if (longName !== null) name = longName;
      longName = null;
      if (type === '0' || type === '\0') files.set(name.replace(/^package\//, ''), new Uint8Array(body));
    }
    pos += 512 + Math.ceil(size / 512) * 512;
  }
  return files;
}

/**
 * A package tarball on disk (`npm pack` of a commit before the release), with the sha512 of the
 * file as its integrity.
 */
export function readTarball(bytes) {
  return { files: untar(gunzipSync(bytes)), integrity: `sha512-${createHash('sha512').update(bytes).digest('base64')}` };
}
