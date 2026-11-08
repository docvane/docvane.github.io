import { $t as metricsOf, A as checkSource$1, B as createWorksheetParser, Bt as resolveLimits, Cn as CORE_PROPERTIES_PART, Ct as relationshipsPartName, D as HEAD_BYTES, Dt as checkUniqueName, E as CORRUPT_HINT, Et as checkDefinedName, F as readCentralDirectory, Ht as printNames, I as unsupported, Jt as DRAWING_CONTENT_TYPE, K as SheetRels, L as validateLocalHeaders, Lt as DEFAULT_READ_LIMITS, M as limitExceeded, N as locateCentralDirectory, Nn as clip, O as NotLoaded, Ot as nameScopeKey, Pt as normalizeStyle, Q as createSharedStrings, R as WarningCollector, Rt as checkKeyText, Sn as APP_PROPERTIES_PART, St as openPackage, T as Bytes, Tn as corePropertiesXml, Tt as normalizeTable, Ut as printXml, Wt as sheetProtectionXml, X as cellXml, Xt as MediaParts, Y as buildStyles, Yt as IMAGE_CONTENT_TYPES, Z as contentTypesXml, Zt as drawingParts, _t as FormulaGroups, a as DATE_FORMAT, b as readNotes, bn as partKey, cn as chartPartXml, ct as sheetHead, d as convertCellInput, dt as sheetTail, en as resolveAnchor, f as createWorkbook, g as detectHeader, gn as formatCellRef, gt as workbookXml, hn as columnLetters, ht as workbookRelsXml, i as readOptions, in as classifyRelationship, it as noteParts, j as corrupt, k as TAIL_BYTES, kn as createXmlByteStreamParser, ln as chartSheetNames, lt as sheetRelsPartName, mn as MAX_ROWS, n as oneOfOption, nt as hyperlinksXml, o as DATE_TIME_FORMAT, on as CHART_CONTENT_TYPE, ot as rootRelsXml, p as sortedNotes, pn as MAX_COLUMNS, pt as tableParts, rt as nameTitles, sn as chartPartName, st as rowStart, tn as resolveImage, u as checkBookViews, un as resolveChartSheets, ut as sheetRulesXml, v as isoDate, w as workbookPartsToLoad, wn as appPropertiesXml, wt as TABLE_CONTENT_TYPE, xn as zipNameToPartName, y as analyzeWorkbook, yt as crc32, z as sheetNotFound, zn as XlvError, zt as checkOutputPositions } from "./options-BD873QNd.js";
import { Deflate, Inflate } from "../fflate/browser.js";
// src/zip/stream.ts
/**
* Compressed bytes handed to the inflater at a time. Limits are checked on every output chunk
* fflate emits, and the input is cut small so that a highly compressed piece cannot produce much
* output before the next check.
*/
const PIECE = 4096;
function memorySource(data) {
	return {
		size: data.length,
		read: (start, end) => Promise.resolve(data.subarray(start, end)),
		pieces: async function* (start, end) {
			for (let at = start; at < end; at += PIECE) yield data.subarray(at, Math.min(end, at + PIECE));
		}
	};
}
function blobSource(blob) {
	return {
		size: blob.size,
		read: async (start, end) => new Uint8Array(await blob.slice(start, end).arrayBuffer()),
		pieces: async function* (start, end) {
			const reader = blob.slice(start, end).stream().getReader();
			try {
				for (;;) {
					const { done, value } = await reader.read();
					if (done) return;
					for (let at = 0; at < value.length; at += PIECE) yield value.subarray(at, at + PIECE);
				}
			} finally {
				reader.releaseLock();
			}
		}
	};
}
function toSource(source) {
	if (source instanceof Uint8Array) return memorySource(source);
	if (source instanceof ArrayBuffer) return memorySource(new Uint8Array(source));
	if (ArrayBuffer.isView(source)) return memorySource(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
	if (typeof Blob !== "undefined" && source instanceof Blob) return blobSource(source);
	throw new XlvError("XLV_E_INVALID_INPUT", `Expected the file as a Uint8Array, ArrayBuffer or Blob, got ${source === null ? "null" : typeof source}.`, { hint: "Pass the bytes of the file, or a Blob (in Node.js: await fs.openAsBlob(path))." });
}
/** Runs `fn`, loading missing byte ranges and retrying; the parsing functions are synchronous. */
async function withLoading(b, source, fn) {
	for (let attempt = 0;; attempt++) try {
		return fn();
	} catch (error) {
		if (!(error instanceof NotLoaded) || attempt > 16) throw error;
		const start = Math.max(0, Math.min(error.start, source.size));
		const end = Math.min(source.size, Math.max(error.end, start + 1024));
		if (end <= start) throw error;
		b.add(start, await source.read(start, end));
	}
}
/**
* Opens a ZIP archive from memory or a Blob and validates its structure with the rules of
* `openZip`. Throws (rejects with) `XLV_E_NOT_ZIP`, `XLV_E_UNSUPPORTED_CONTAINER`,
* `XLV_E_ZIP_CORRUPT`, `XLV_E_ZIP_UNSUPPORTED` or `XLV_E_LIMIT_EXCEEDED`.
*/
async function openZipSource(source, limits) {
	const src = toSource(source);
	const size = src.size;
	const head = await src.read(0, Math.min(size, HEAD_BYTES));
	checkSource$1(size, head, limits);
	let b;
	if (typeof Blob !== "undefined" && source instanceof Blob) {
		b = new Bytes(size, [{
			start: 0,
			data: head
		}]);
		const tailStart = Math.max(0, size - TAIL_BYTES);
		if (size > 0) b.add(tailStart, await src.read(tailStart, size));
	} else b = new Bytes(size, [{
		start: 0,
		data: await src.read(0, size)
	}]);
	const dir = await withLoading(b, src, () => locateCentralDirectory(b, limits));
	if (!b.has(dir.centralOffset, dir.centralLimit) && dir.centralLimit > dir.centralOffset) b.add(dir.centralOffset, await src.read(dir.centralOffset, dir.centralLimit));
	const entries = await withLoading(b, src, () => readCentralDirectory(b, dir.centralOffset, dir.centralLimit, dir.entryCount));
	for (const entry of [...entries].sort((x, y) => x.headerOffset - y.headerOffset)) {
		const at = entry.headerOffset;
		if (at + 30 > dir.centralOffset || b.has(at, at + 30)) continue;
		const fixed = await src.read(at, at + 30);
		b.add(at, fixed);
		const nameLength = fixed[26] | fixed[27] << 8;
		const extraLength = fixed[28] | fixed[29] << 8;
		const end = Math.min(dir.centralOffset, at + 30 + nameLength + extraLength);
		if (end > at + 30) b.add(at + 30, await src.read(at + 30, end));
	}
	await withLoading(b, src, () => validateLocalHeaders(b, entries, dir.centralOffset));
	return new SourceZip(src, entries, limits);
}
var SourceZip = class {
	source;
	limits;
	entries;
	byName;
	/** Decompressed bytes produced so far by all entries (maxTotalBytes). */
	produced = 0;
	constructor(source, entries, limits) {
		this.source = source;
		this.limits = limits;
		this.entries = entries.map(({ nameBytes: _nameBytes, ...entry }) => entry);
		this.byName = new Map(this.entries.map((e) => [e.name.toLowerCase(), e]));
	}
	find(name) {
		return this.byName.get(name.toLowerCase());
	}
	async read(entry, maxBytes) {
		const chunks = [];
		let length = 0;
		const reader = this.stream(entry, maxBytes).getReader();
		for (;;) {
			const { done, value } = await reader.read();
			if (done) break;
			chunks.push(value);
			length += value.length;
		}
		if (chunks.length === 1) return chunks[0];
		const out = new Uint8Array(length);
		let at = 0;
		for (const chunk of chunks) {
			out.set(chunk, at);
			at += chunk.length;
		}
		return out;
	}
	/** Checks that apply before any byte of `entry` is read, in the order of `ZipReader.read`. */
	precheck(entryOrName, maxEntryBytes) {
		const entry = typeof entryOrName === "string" ? this.find(entryOrName) : entryOrName;
		if (entry === void 0) throw new XlvError("XLV_E_PART_MISSING", `The archive has no entry "${String(entryOrName)}".`, {
			location: { part: `/${String(entryOrName)}` },
			hint: "The workbook is incomplete. Re-save it in a spreadsheet application."
		});
		const { limits } = this;
		if (entry.uncompressedSize > maxEntryBytes) throw limitExceeded("maxEntryBytes", maxEntryBytes, `ZIP entry "${entry.name}" declares ${entry.uncompressedSize} bytes, more than the limit.`, entry.name);
		if (this.produced + entry.uncompressedSize > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
		if (entry.method === 0) {
			if (entry.compressedSize !== entry.uncompressedSize) throw corrupt(`Stored ZIP entry "${entry.name}" has inconsistent sizes.`, "size-mismatch", entry.headerOffset, { name: entry.name });
		} else if (entry.method !== 8) throw unsupported(`ZIP entry "${entry.name}" uses compression method ${entry.method}.`, "compression-method", "Only stored and Deflate-compressed entries are supported. Re-save the workbook in a spreadsheet application.", entry.name);
		return entry;
	}
	stream(entryOrName, maxBytes) {
		let entry;
		let pieces = null;
		let inflater = null;
		const out = [];
		let produced = 0;
		let crc = 0;
		let ended = false;
		let finished = false;
		const zip = this;
		const { limits } = this;
		const maxEntryBytes = Math.min(maxBytes ?? limits.maxEntryBytes, limits.maxEntryBytes);
		const take = (chunk) => {
			produced += chunk.length;
			zip.produced += chunk.length;
			if (produced > maxEntryBytes) throw limitExceeded("maxEntryBytes", maxEntryBytes, `ZIP entry "${entry.name}" decompresses to more than the limit.`, entry.name);
			if (zip.produced > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
			if (entry.method === 8 && produced > limits.compressionRatioMinBytes && produced / Math.max(1, entry.compressedSize) > limits.maxCompressionRatio) throw limitExceeded("maxCompressionRatio", limits.maxCompressionRatio, `ZIP entry "${entry.name}" expands more than ${String(limits.maxCompressionRatio)} times, which is typical of a zip bomb.`, entry.name);
			if (chunk.length > 0) {
				crc = crc32(chunk, crc);
				out.push(chunk);
			}
		};
		const finish = () => {
			if (entry.method === 8 && !ended) throw corrupt(`ZIP entry "${entry.name}" ends early.`, "deflate", entry.headerOffset, { name: entry.name });
			if (produced !== entry.uncompressedSize) throw corrupt(`ZIP entry "${entry.name}" is ${produced} bytes, but its header says ${entry.uncompressedSize}.`, "size-mismatch", entry.headerOffset, { name: entry.name });
			if (crc !== entry.crc32) throw corrupt(`ZIP entry "${entry.name}" failed its CRC-32 check.`, "crc-mismatch", entry.headerOffset, { name: entry.name });
			finished = true;
		};
		const inflateError = (error) => error instanceof XlvError ? error : new XlvError("XLV_E_ZIP_CORRUPT", `ZIP entry "${entry.name}" is not a valid Deflate stream.`, {
			location: {
				part: `/${entry.name}`,
				offset: entry.headerOffset
			},
			hint: CORRUPT_HINT,
			details: {
				reason: "deflate",
				name: entry.name
			},
			cause: error
		});
		return new ReadableStream({
			start: () => {
				entry = zip.precheck(entryOrName, maxEntryBytes);
				pieces = zip.source.pieces(entry.dataOffset, entry.dataOffset + entry.compressedSize);
				if (entry.method === 8) {
					inflater = new Inflate();
					inflater.ondata = (chunk, final) => {
						take(chunk);
						if (final) ended = true;
					};
				}
			},
			pull: async (controller) => {
				try {
					while (out.length === 0 && !finished) {
						const next = await pieces.next();
						if (next.done === true) {
							if (inflater && !ended) inflater.push(/* @__PURE__ */ new Uint8Array(0), true);
							finish();
						} else if (inflater) {
							if (!ended) inflater.push(next.value, false);
						} else take(next.value);
					}
					while (out.length > 0) controller.enqueue(out.shift());
					if (finished) controller.close();
				} catch (error) {
					await pieces?.return?.();
					throw entry === void 0 ? error : inflateError(error);
				}
			},
			cancel: async () => {
				await pieces?.return?.();
			}
		});
	}
};
/** Entries a ZIP without ZIP64 can hold (D-202). */
const MAX_ENTRIES = 65534;
const MAX_32 = 4294967294;
const SIG_LOCAL = 67324752;
const SIG_CENTRAL = 33639248;
const SIG_EOCD = 101010256;
const SIG_DESCRIPTOR = 134695760;
const VERSION = 20;
const FLAG_DESCRIPTOR = 8;
const DOS_DATE = 33;
const DOS_TIME = 0;
const encoder$1 = new TextEncoder();
function tooLarge(what) {
	return new XlvError("XLV_E_LIMIT_EXCEEDED", `The workbook is too large to write: ${what}.`, {
		hint: "Split the data into several workbooks. Writing ZIP64 archives is not supported yet.",
		details: { limit: "zip32" }
	});
}
function checkName(name, seen) {
	if (typeof name !== "string" || !/^[\x21-\x7e][\x20-\x7e]*$/.test(name) || name.startsWith("/") || name.length > 65535) {
		const shown = String(name).slice(0, 60);
		throw new XlvError("XLV_E_INVALID_INPUT", `"${shown}" is not a valid ZIP item name.`, {
			hint: "Item names must be printable ASCII without a leading slash. This is an internal error; please report it.",
			details: { name: shown }
		});
	}
	const key = name.toLowerCase();
	if (seen.has(key)) throw new XlvError("XLV_E_INVALID_INPUT", `The ZIP item name "${name.slice(0, 60)}" is used twice.`, {
		hint: "Item names must be unique ignoring case. This is an internal error; please report it.",
		details: { name: name.slice(0, 60) }
	});
	seen.add(key);
	return name;
}
function iteratorOf(value) {
	const asyncFactory = value[Symbol.asyncIterator];
	if (typeof asyncFactory === "function") return asyncFactory.call(value);
	return value[Symbol.iterator]();
}
/**
* Writes a ZIP archive as a stream. Entries are requested from `entries` one at a time, only
* after the previous entry's data is fully consumed, and data chunks only when the consumer
* reads; memory therefore stays bounded by the chunk sizes. Problems (invalid or duplicate names,
* archives that would need ZIP64) are reported as an `XlvError` through the stream.
*/
function createZipStream(entries, level = 6) {
	let iterator = null;
	let current = null;
	const central = [];
	const seen = /* @__PURE__ */ new Set();
	const pending = [];
	let offset = 0;
	let count = 0;
	let closed = false;
	const emit = (chunk) => {
		if (chunk.length === 0) return;
		offset += chunk.length;
		if (offset > MAX_32) throw tooLarge("more than 4 GiB");
		pending.push(chunk);
	};
	const startEntry = (entry) => {
		count++;
		if (count > MAX_ENTRIES) throw tooLarge(`more than ${String(MAX_ENTRIES)} parts`);
		const name = checkName(entry.name, seen);
		const nameBytes = encoder$1.encode(name);
		const local = new Uint8Array(30 + nameBytes.length);
		const v = new DataView(local.buffer);
		v.setUint32(0, SIG_LOCAL, true);
		v.setUint16(4, VERSION, true);
		v.setUint16(6, FLAG_DESCRIPTOR, true);
		v.setUint16(8, 8, true);
		v.setUint16(10, DOS_TIME, true);
		v.setUint16(12, DOS_DATE, true);
		v.setUint16(26, nameBytes.length, true);
		local.set(nameBytes, 30);
		const headerOffset = offset;
		emit(local);
		const data = entry.data;
		const dataIterator = data instanceof Uint8Array ? [data][Symbol.iterator]() : iteratorOf(data);
		const deflater = new Deflate({ level });
		const open = {
			nameBytes,
			offset: headerOffset,
			data: dataIterator,
			deflater,
			crc: 0,
			size: 0,
			compressed: 0
		};
		deflater.ondata = (chunk) => {
			open.compressed += chunk.length;
			emit(chunk);
		};
		current = open;
	};
	const finishEntry = (open) => {
		open.deflater.push(/* @__PURE__ */ new Uint8Array(0), true);
		if (open.size > MAX_32 || open.compressed > MAX_32) throw tooLarge("a part larger than 4 GiB");
		const descriptor = /* @__PURE__ */ new Uint8Array(16);
		const d = new DataView(descriptor.buffer);
		d.setUint32(0, SIG_DESCRIPTOR, true);
		d.setUint32(4, open.crc, true);
		d.setUint32(8, open.compressed, true);
		d.setUint32(12, open.size, true);
		emit(descriptor);
		const cd = new Uint8Array(46 + open.nameBytes.length);
		const c = new DataView(cd.buffer);
		c.setUint32(0, SIG_CENTRAL, true);
		c.setUint16(4, VERSION, true);
		c.setUint16(6, VERSION, true);
		c.setUint16(8, FLAG_DESCRIPTOR, true);
		c.setUint16(10, 8, true);
		c.setUint16(12, DOS_TIME, true);
		c.setUint16(14, DOS_DATE, true);
		c.setUint32(16, open.crc, true);
		c.setUint32(20, open.compressed, true);
		c.setUint32(24, open.size, true);
		c.setUint16(28, open.nameBytes.length, true);
		c.setUint32(42, open.offset, true);
		cd.set(open.nameBytes, 46);
		central.push(cd);
		current = null;
	};
	const finishArchive = () => {
		const cdOffset = offset;
		for (const cd of central) emit(cd);
		const cdSize = offset - cdOffset;
		const eocd = /* @__PURE__ */ new Uint8Array(22);
		const e = new DataView(eocd.buffer);
		e.setUint32(0, SIG_EOCD, true);
		e.setUint16(8, central.length, true);
		e.setUint16(10, central.length, true);
		e.setUint32(12, cdSize, true);
		e.setUint32(16, cdOffset, true);
		emit(eocd);
		closed = true;
	};
	return new ReadableStream({
		pull: async (controller) => {
			try {
				while (pending.length === 0 && !closed) {
					const open = current;
					if (open === null) {
						iterator ??= iteratorOf(entries);
						const next = await iterator.next();
						if (next.done === true) finishArchive();
						else startEntry(next.value);
						continue;
					}
					const next = await open.data.next();
					if (next.done === true) {
						finishEntry(open);
						continue;
					}
					const chunk = next.value;
					if (!(chunk instanceof Uint8Array)) throw new XlvError("XLV_E_INVALID_INPUT", "ZIP entry data must be Uint8Array chunks.", { hint: "This is an internal error; please report it." });
					open.crc = crc32(chunk, open.crc);
					open.size += chunk.length;
					if (open.size > MAX_32) throw tooLarge("a part larger than 4 GiB");
					open.deflater.push(chunk, false);
				}
				while (pending.length > 0) controller.enqueue(pending.shift());
				if (closed) controller.close();
			} catch (error) {
				await current?.data.return?.();
				await iterator?.return?.();
				throw error;
			}
		},
		cancel: async () => {
			await current?.data.return?.();
			await iterator?.return?.();
		}
	});
}
// src/stream-read.ts
const GiB = 1073741824;
/**
* Limits that differ in streaming mode (D-401): `maxCells` does not apply (but one `<row>` element
* may hold at most 16 384 cells, D-J01), and single parts may
* decompress to 4 GiB (8 GiB in total). All other `DEFAULT_READ_LIMITS` stay.
*
* @example
* ```ts
* import { STREAM_LIMIT_DEFAULTS } from 'xlvane';
*
* console.log(STREAM_LIMIT_DEFAULTS.maxEntryBytes); // 4294967296
* ```
*/
const STREAM_LIMIT_DEFAULTS = Object.freeze({
	maxCells: Infinity,
	maxEntryBytes: 4 * GiB,
	maxTotalBytes: 8 * GiB
});
/**
* A ZipReader over entries loaded in advance, so the OPC layer can stay synchronous. These parts
* (relationships, workbook, shared strings, styles, comments, VML) are held in memory whole, so
* each is bounded by the batch reader's `maxEntryBytes` (256 MiB by default), not by the larger
* streaming bound meant for worksheets that are inflated piece by piece (D-R03).
*/
var PreloadedZip = class {
	zip;
	maxBufferedBytes;
	entries;
	bytesRead = 0;
	data = /* @__PURE__ */ new Map();
	/**
	* ZIP entries by the part name they map to, as openPackage maps them: percent-encoded
	* non-ASCII characters of an item name are decoded (Part 2 §7.3.5; F-OPC-12).
	*/
	byPart = /* @__PURE__ */ new Map();
	constructor(zip, maxBufferedBytes) {
		this.zip = zip;
		this.maxBufferedBytes = maxBufferedBytes;
		this.entries = zip.entries;
		for (const entry of zip.entries) {
			if (entry.name.endsWith("/")) continue;
			const key = partKey(zipNameToPartName(entry.name));
			if (!this.byPart.has(key)) this.byPart.set(key, entry);
		}
	}
	/** The ZIP entry that stores a part, or undefined. */
	entryOf(partName) {
		return this.byPart.get(partKey(partName));
	}
	async load(entries) {
		for (const entry of entries) {
			if (entry === void 0 || this.data.has(entry.name)) continue;
			const bytes = await this.zip.read(entry, this.maxBufferedBytes);
			this.data.set(entry.name, bytes);
			this.bytesRead += bytes.length;
		}
	}
	/** Loads the parts with these names that exist. */
	async loadParts(partNames) {
		await this.load([...partNames].map((name) => this.entryOf(name)));
	}
	/** Loads the Relationships parts of these sources (null: the package) that exist. */
	async loadRelationships(sources) {
		await this.loadParts([...sources].map(relationshipsPartName));
	}
	find(name) {
		return this.zip.find(name);
	}
	read(entry) {
		const name = typeof entry === "string" ? this.zip.find(entry)?.name ?? entry : entry.name;
		const bytes = this.data.get(name);
		if (bytes === void 0) throw new XlvError("XLV_E_INTERNAL", `The ZIP entry "${clip(name)}" was read before it was loaded.`, {
			location: { part: `/${clip(name)}` },
			hint: "This is a bug in xlvane. Please report it, with the file if you can share it; readWorkbook reads the file meanwhile."
		});
		return bytes;
	}
};
/**
* The internal targets of the workbook part's relationships: the sheets among them, whose
* Relationships parts analyzeWorkbook reads. Only these, and not every .rels entry of the ZIP,
* are loaded (D-H01).
*/
function workbookTargets(pkg, workbookPart) {
	const seen = /* @__PURE__ */ new Set();
	const out = [];
	for (const rel of pkg.relationships(workbookPart)) {
		const key = partKey(rel.target);
		if (rel.external || seen.has(key) || !pkg.hasPart(rel.target)) continue;
		seen.add(key);
		out.push(rel.target);
	}
	return out;
}
function checkSource(source) {
	if (source instanceof Uint8Array || source instanceof ArrayBuffer) return source;
	if (ArrayBuffer.isView(source)) return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
	if (typeof Blob !== "undefined" && source instanceof Blob) return source;
	throw new XlvError("XLV_E_INVALID_INPUT", `Expected the file as a Uint8Array, ArrayBuffer or Blob, got ${source === null ? "null" : typeof source}.`, { hint: typeof source === "string" ? "Pass the file, not its path. In Node.js use openWorkbookStreamFile() from \"xlvane/node\", or fs.openAsBlob(path)." : "Pass the bytes, or a Blob/File (for example from an <input type=\"file\"> or fs.openAsBlob())." });
}
/**
* Opens a workbook for reading row by row. Reads the workbook structure, shared strings and
* styles; worksheet cells are decompressed and parsed only while their rows are iterated, one
* row at a time. With a `Blob` (for example `fs.openAsBlob()` in Node.js, or a `File` in the
* browser) the file itself is never held in memory either. Untrusted-input limits apply as in
* `readWorkbook`, except `maxCells` (see {@link OpenWorkbookStreamOptions}).
*
* @example
* ```ts
* import { openWorkbookStream } from 'xlvane';
*
* declare const file: Blob;
* const book = await openWorkbookStream(file);
* let total = 0;
* for await (const row of book.rows('Orders')) {
*   const amount = row.cells.find((cell) => cell.col === 3)?.value;
*   if (typeof amount === 'number') total += amount;
* }
* console.log(total);
* ```
*/
async function openWorkbookStream(source, options) {
	const input = checkSource(source);
	const userLimits = readOptions(options, ["limits"], "openWorkbookStream options").get("limits");
	const limits = resolveLimits({
		...STREAM_LIMIT_DEFAULTS,
		...userLimits
	});
	const zip = await openZipSource(input, limits);
	const preloaded = new PreloadedZip(zip, userLimits?.maxEntryBytes ?? DEFAULT_READ_LIMITS.maxEntryBytes);
	await preloaded.load(zip.entries.filter((e) => e.name.toLowerCase() === "[content_types].xml"));
	await preloaded.loadRelationships([null]);
	const pkg = openPackage(preloaded, limits);
	const workbookPart = pkg.relationships(null).find((r) => !r.external && classifyRelationship(r.type)?.name === "officeDocument")?.target;
	if (workbookPart !== void 0) {
		await preloaded.loadRelationships([workbookPart]);
		await preloaded.loadParts(workbookPartsToLoad(pkg));
		await preloaded.loadRelationships(workbookTargets(pkg, workbookPart));
	}
	const warnings = new WarningCollector();
	const info = analyzeWorkbook(pkg, limits, warnings);
	const formatCache = /* @__PURE__ */ new Map();
	const sheetEntry = (name) => {
		const entry = info.sheets.find((s) => s.name === name);
		if (entry !== void 0) return entry;
		throw sheetNotFound(name, info.sheets.map((s) => s.name));
	};
	const noteCache = /* @__PURE__ */ new Map();
	async function notesOf(name) {
		const entry = sheetEntry(name);
		if (entry.kind !== "worksheet" || entry.part === null) return [];
		const cached = noteCache.get(entry.part);
		if (cached !== void 0) return cached;
		const parts = pkg.relationships(entry.part).filter((r) => !r.external && ["comments", "vmlDrawing"].includes(classifyRelationship(r.type)?.name ?? "") && pkg.hasPart(r.target) && !entry.foreignParts.has(partKey(r.target)));
		if (parts.length > 0) await preloaded.loadParts(parts.map((r) => r.target));
		const notes = Object.freeze(readNotes(pkg, entry.part, entry.name, limits, warnings, entry.foreignParts));
		noteCache.set(entry.part, notes);
		return notes;
	}
	async function* rows(name) {
		const entry = sheetEntry(name);
		if (entry.kind !== "worksheet" || entry.part === null) return;
		const part = entry.part;
		const ready = [];
		const notes = /* @__PURE__ */ new Map();
		for (const note of await notesOf(name)) if (!notes.has(note.ref)) notes.set(note.ref, note);
		const withNotes = (cells) => {
			if (notes.size === 0) return cells;
			return cells.map((cell) => {
				const note = notes.get(cell.ref);
				return note === void 0 ? cell : {
					...cell,
					note
				};
			});
		};
		const { handler, finish } = createWorksheetParser({
			part,
			sheet: entry.name,
			limits,
			warnings,
			sharedStrings: info.sharedStrings,
			richStrings: info.richStrings,
			styles: info.styles,
			dateSystem: info.dateSystem,
			conformance: info.conformance,
			budget: {
				remaining: Infinity,
				sharedText: limits.maxSharedStringChars,
				formulaText: limits.maxSharedStringChars
			},
			formatCache
		}, (number, cells, rowInfo, styledCells) => {
			const row = {
				number,
				cells: withNotes(cells),
				styledCells,
				height: rowInfo?.height ?? null,
				hidden: rowInfo?.hidden ?? false
			};
			ready.push(rowInfo?.style === void 0 ? row : {
				...row,
				style: rowInfo.style
			});
		});
		const parser = createXmlByteStreamParser(handler, {
			part,
			limits
		});
		const stored = preloaded.entryOf(part);
		if (stored === void 0) throw new XlvError("XLV_E_PART_MISSING", `The part ${clip(part)} of sheet "${entry.name}" is not in the archive.`, {
			location: {
				part: clip(part),
				sheet: entry.name
			},
			hint: "The workbook is incomplete. Re-save it in a spreadsheet application."
		});
		const reader = zip.stream(stored).getReader();
		try {
			for (;;) {
				const { done, value } = await reader.read();
				if (done) break;
				parser.write(value);
				while (ready.length > 0) yield ready.shift();
			}
			parser.end();
			finish();
			while (ready.length > 0) yield ready.shift();
		} finally {
			await reader.cancel().catch(() => void 0);
		}
	}
	return {
		sheets: info.sheets.map((s) => ({
			name: s.name,
			index: s.index,
			sheetId: s.sheetId,
			kind: s.kind,
			state: s.state
		})),
		dateSystem: info.dateSystem,
		conformance: info.conformance,
		format: info.format,
		hasMacros: info.hasMacros,
		externalLinkCount: info.externalLinkCount,
		properties: info.properties,
		definedNames: info.definedNames,
		get warnings() {
			return warnings.list();
		},
		notes: notesOf,
		rows(name) {
			sheetEntry(name);
			return rows(name);
		},
		rowStream(name) {
			sheetEntry(name);
			const iterator = rows(name);
			return new ReadableStream({
				async pull(controller) {
					const next = await iterator.next();
					if (next.done) controller.close();
					else controller.enqueue(next.value);
				},
				async cancel() {
					await iterator.return(void 0);
				}
			});
		}
	};
}
function setOwn(target, key, value) {
	Object.defineProperty(target, key, {
		value,
		enumerable: true,
		writable: true,
		configurable: true
	});
}
async function* readTableStream(source, options) {
	const o = readOptions(options, [
		"sheet",
		"header",
		"schema",
		"dates",
		"limits"
	], "readTableStream options");
	const dates = oneOfOption("dates", o.get("dates"), ["date", "iso"], "date");
	const header = o.get("header") ?? "auto";
	if (header !== "auto" && header !== false && !(typeof header === "number" && Number.isInteger(header) && header >= 1)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"header\" must be \"auto\", a row number or false.", {
		hint: "Use header: \"auto\", header: 1 or header: false.",
		details: { option: "header" }
	});
	const schema = o.get("schema");
	if (schema !== void 0) {
		const props = schema?.["~standard"];
		if (props?.version !== 1 || typeof props.validate !== "function") throw new XlvError("XLV_E_SCHEMA_INVALID", "The option \"schema\" is not a Standard Schema (version 1).", {
			hint: "Pass a Zod, Valibot or ArkType schema.",
			details: { option: "schema" }
		});
	}
	const limits = o.get("limits");
	const book = await openWorkbookStream(source, limits === void 0 ? void 0 : { limits });
	const outputCells = limits?.maxCells ?? DEFAULT_READ_LIMITS.maxCells;
	const keyTextLimit = limits?.maxSharedStringChars ?? DEFAULT_READ_LIMITS.maxSharedStringChars;
	const sheetName = o.get("sheet") ?? (book.sheets.find((s) => s.kind === "worksheet" && s.state === "visible") ?? book.sheets[0])?.name;
	if (sheetName === void 0) return;
	const buffered = [];
	let columns = null;
	let headerRow = null;
	const value = (cell) => {
		const v = cell?.value ?? null;
		return v instanceof Date && dates === "iso" ? isoDate(v) : v;
	};
	const decide = () => {
		const nonEmpty = buffered.filter((r) => r.cells.length > 0);
		if (header === "auto") headerRow = detectHeader(nonEmpty.map((r) => r.cells), widthOf(nonEmpty), 20)?.row ?? null;
		else if (typeof header === "number") {
			if (!nonEmpty.some((r) => r.number === header)) throw new XlvError("XLV_E_HEADER_NOT_FOUND", `Row ${String(header)} of sheet "${sheetName}" has no cells to use as a header.`, {
				location: { sheet: sheetName },
				hint: "Pass the row that holds the column names, header: \"auto\", or header: false.",
				details: { row: header }
			});
			headerRow = header;
		}
		const source = headerRow === null ? nonEmpty.flatMap((r) => r.cells) : nonEmpty.find((r) => r.number === headerRow)?.cells ?? [];
		let minCol = Infinity;
		let maxCol = -Infinity;
		for (const c of source) {
			if (c.col < minCol) minCol = c.col;
			if (c.col > maxCol) maxCol = c.col;
		}
		const headerCells = new Map(headerRow === null ? [] : source.map((c) => [c.col, c]));
		const out = [];
		const used = /* @__PURE__ */ new Set();
		const nextSuffix = /* @__PURE__ */ new Map();
		for (let col = minCol; col <= maxCol; col++) {
			const letter = columnLetters(col);
			const cell = headerCells.get(col);
			const text = cell === void 0 || cell.value === null ? null : cell.value instanceof Date ? isoDate(cell.value) : String(cell.value);
			const base = text === null || text.trim() === "" ? letter : text.trim();
			let key = base;
			let n = nextSuffix.get(base) ?? 2;
			while (used.has(key)) key = `${base}_${String(n++)}`;
			nextSuffix.set(base, n);
			used.add(key);
			out.push({
				key,
				header: text === null || text.trim() === "" ? null : text,
				col,
				letter
			});
		}
		return out;
	};
	let positions = 0;
	let cellsRead = 0;
	let byKey = null;
	let byKeyOf = null;
	const columnOf = (cols, key) => {
		if (byKey === null || byKeyOf !== cols) {
			byKey = /* @__PURE__ */ new Map();
			for (const c of cols) if (!byKey.has(c.key)) byKey.set(c.key, c);
			byKeyOf = cols;
		}
		return byKey.get(key);
	};
	const emit = async (r, cols) => {
		if (headerRow !== null && r.number <= headerRow) return null;
		cellsRead += r.cells.length;
		positions += cols.length;
		checkOutputPositions(positions, cellsRead, outputCells, "The table", sheetName);
		const byCol = new Map(r.cells.map((c) => [c.col, c]));
		if (cols.every((c) => byCol.get(c.col)?.value == null)) return null;
		checkKeyText(cols, 1, keyTextLimit, sheetName);
		const row = {};
		for (const c of cols) setOwn(row, c.key, value(byCol.get(c.col)));
		if (schema === void 0) return {
			row: r.number,
			value: row,
			issues: []
		};
		const result = await schema["~standard"].validate(row);
		if (!result.issues) return {
			row: r.number,
			value: result.value,
			issues: []
		};
		const issues = result.issues.map((issue) => {
			const path = (issue.path ?? []).map((seg) => {
				const key = typeof seg === "object" && seg !== null ? seg.key : seg;
				return typeof key === "number" ? key : String(key);
			});
			const column = typeof path[0] === "string" ? columnOf(cols, path[0]) : void 0;
			return {
				row: r.number,
				cell: column ? formatCellRef(r.number, column.col) : null,
				column: column?.key ?? null,
				message: String(issue.message),
				path
			};
		});
		return {
			row: r.number,
			value: void 0,
			issues
		};
	};
	for await (const r of book.rows(sheetName)) {
		if (r.cells.length === 0) continue;
		if (columns === null) {
			buffered.push(r);
			if (buffered.filter((b) => b.cells.length > 0).length < 20) continue;
			columns = decide();
			for (const b of buffered.splice(0)) {
				const item = await emit(b, columns);
				if (item) yield item;
			}
			continue;
		}
		const item = await emit(r, columns);
		if (item) yield item;
	}
	if (columns === null && buffered.some((b) => b.cells.length > 0)) {
		columns = decide();
		for (const b of buffered.splice(0)) {
			const item = await emit(b, columns);
			if (item) yield item;
		}
	}
}
function widthOf(rows) {
	let min = Infinity;
	let max = 0;
	for (const r of rows) for (const c of r.cells) {
		if (c.col < min) min = c.col;
		if (c.col > max) max = c.col;
	}
	return max === 0 ? 0 : max - min + 1;
}
// src/stream-write.ts
const CHUNK_CHARS = 65536;
const encoder = new TextEncoder();
function isIterable(value) {
	return typeof value === "object" && value !== null && (typeof value[Symbol.iterator] === "function" || typeof value[Symbol.asyncIterator] === "function");
}
async function* iterate(source) {
	if (typeof source[Symbol.asyncIterator] === "function") yield* source;
	else yield* source;
}
/**
* Writes a workbook as a stream of XLSX bytes. Rows are pulled from each sheet's iterable only
* as fast as the stream is read, so files with millions of rows need constant memory (with the
* default inline strings). Values and styles follow the rules of `NewSheet.set`; an invalid
* value errors the stream with an `XlvError` naming the sheet and cell. The layout (columns,
* frozen panes, merges, filter) is given up front.
*
* @example
* ```ts
* import { writeWorkbookStream } from 'xlvane';
*
* function* rows() {
*   yield { values: ['id', 'value'], style: { font: { bold: true } } };
*   for (let i = 1; i <= 1_000_000; i++) yield [i, Math.sqrt(i)];
* }
* const stream = writeWorkbookStream({ sheets: [{ name: 'Data', freeze: { rows: 1 }, rows: rows() }] });
* const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
* ```
*/
function writeWorkbookStream(options) {
	const o = readOptions(options, [
		"sheets",
		"dateSystem",
		"sharedStrings",
		"properties",
		"definedNames",
		"protection",
		"images",
		"views",
		"calculation",
		"compressionLevel",
		"defaultFont"
	], "writeWorkbookStream options");
	const dateSystem = o.get("dateSystem");
	const book = createWorkbook({
		...dateSystem === void 0 ? {} : { dateSystem },
		...o.has("properties") ? { properties: o.get("properties") } : {},
		...o.has("defaultFont") ? { defaultFont: o.get("defaultFont") } : {}
	});
	if (o.has("protection")) book.protect(o.get("protection"));
	if (o.has("views")) book.views(o.get("views"));
	if (o.has("calculation")) book.calculation(o.get("calculation"));
	const level = o.get("compressionLevel") ?? 6;
	if (typeof level !== "number" || !Number.isInteger(level) || level < 0 || level > 9) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"compressionLevel\" must be a whole number from 0 to 9.", {
		hint: "Use 0 for no compression or 9 for the smallest file; the default is 6.",
		details: { option: "compressionLevel" }
	});
	const imageInput = o.get("images") ?? [];
	if (!Array.isArray(imageInput)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"images\" must be an array of image files.", {
		hint: "For example images: [pngBytes], then images: [{ image: 0, anchor: \"B2:D6\" }] on a sheet.",
		details: { option: "images" }
	});
	for (const bytes of imageInput) book.addImage(bytes);
	const media = new MediaParts(book.imageList());
	const core = corePropertiesXml(book.properties);
	const rawNames = o.get("definedNames") ?? [];
	if (!Array.isArray(rawNames)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"definedNames\" must be an array.", {
		hint: "For example definedNames: [{ name: \"Items\", refersTo: \"Data!$A$1:$A$10\" }].",
		details: { option: "definedNames" }
	});
	const names = [];
	const nameKeys = /* @__PURE__ */ new Map();
	rawNames.forEach((raw, i) => {
		const r = readOptions(raw, [
			"name",
			"refersTo",
			"sheet",
			"hidden",
			"comment"
		], `definedNames[${String(i)}]`);
		const { name, refersTo, ...rest } = Object.fromEntries(r);
		const n = checkDefinedName(name, refersTo, rest, null);
		checkUniqueName(n, nameKeys);
		names.push(n);
		nameKeys.set(nameScopeKey(n), n);
	});
	const addNames = () => {
		for (const n of names) book.defineName(n.name, n.refersTo, {
			...n.sheet === null ? {} : { sheet: n.sheet },
			hidden: n.hidden,
			...n.comment === null ? {} : { comment: n.comment }
		});
	};
	const sharedOption = o.get("sharedStrings") ?? false;
	if (typeof sharedOption !== "boolean") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"sharedStrings\" must be true or false.", {
		hint: "Use sharedStrings: true or leave it out.",
		details: { option: "sharedStrings" }
	});
	const shared = sharedOption;
	const specs = o.get("sheets");
	const dynamic = !Array.isArray(specs);
	if (dynamic ? !isIterable(specs) : specs.length === 0) throw new XlvError("XLV_E_INVALID_INPUT", "The workbook has no sheets.", { hint: "Pass sheets: [{ name: \"Sheet1\", rows: [...] }]." });
	const prepare = (raw, i) => {
		const s = readOptions(raw, [
			"name",
			"state",
			"sheetId",
			"columns",
			"freeze",
			"tabColor",
			"view",
			"properties",
			"autoFilter",
			"merges",
			"hyperlinks",
			"protect",
			"pageSetup",
			"headerFooter",
			"pageBreaks",
			"printArea",
			"printTitles",
			"dataValidations",
			"conditionalFormats",
			"tables",
			"images",
			"background",
			"charts",
			"rows"
		], `sheets[${String(i)}]`);
		const state = s.get("state");
		const sheetId = s.get("sheetId");
		const sheet = book.addSheet(s.get("name"), {
			...state === void 0 ? {} : { state },
			...sheetId === void 0 ? {} : { sheetId }
		});
		const columns = s.get("columns");
		if (columns !== void 0) {
			if (!Array.isArray(columns)) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].columns must be an array.`, {
				hint: "Pass column options by position, e.g. [{ width: 20 }, undefined, { hidden: true }].",
				details: { option: "columns" }
			});
			columns.forEach((c, k) => {
				if (c !== void 0 && c !== null) sheet.column(k + 1, c);
			});
		}
		const freeze = s.get("freeze");
		if (freeze !== void 0) sheet.freeze(freeze);
		const tabColor = s.get("tabColor");
		if (tabColor !== void 0) sheet.tabColor(tabColor);
		const view = s.get("view");
		if (view !== void 0) sheet.view(view);
		if (s.has("properties")) sheet.properties(s.get("properties"));
		if (s.has("protect")) sheet.protect(s.get("protect"));
		if (s.has("pageSetup")) sheet.pageSetup(s.get("pageSetup"));
		if (s.has("headerFooter")) sheet.headerFooter(s.get("headerFooter"));
		if (s.has("pageBreaks")) sheet.pageBreaks(s.get("pageBreaks"));
		if (s.has("printArea")) sheet.printArea(s.get("printArea"));
		if (s.has("printTitles")) sheet.printTitles(s.get("printTitles"));
		const listOf = (key, example) => {
			const v = s.get(key) ?? [];
			if (!Array.isArray(v)) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].${key} must be an array.`, {
				hint: `For example ${key}: [${example}].`,
				details: { option: key }
			});
			return v.map((item) => {
				if (typeof item !== "object" || item === null) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].${key} must hold objects.`, {
					hint: `For example ${example}.`,
					details: { option: key }
				});
				return item;
			});
		};
		for (const v of listOf("dataValidations", "{ sqref: \"A2:A9\", type: \"list\", formula1: \"\\\"a,b\\\"\" }")) {
			const { sqref, ...rule } = v;
			sheet.validate(sqref, rule);
		}
		for (const v of listOf("conditionalFormats", "{ sqref: \"A2:A9\", rules: [{ type: \"cellIs\", operator: \"greaterThan\", formulas: [0], style: { font: { bold: true } } }] }")) {
			const c = readOptions(v, ["sqref", "rules"], `sheets[${String(i)}].conditionalFormats`);
			sheet.conditionalFormat(c.get("sqref"), c.get("rules"));
		}
		for (const v of listOf("tables", "{ name: \"Sales\", ref: \"A1:B10\", columns: [{ name: \"Region\" }, { name: \"Revenue\" }] }")) {
			const t = normalizeTable(v, sheet.name, true);
			book.checkTable(t, sheet);
			sheet.pushTable(t);
		}
		const merges = s.get("merges");
		if (merges !== void 0) {
			if (!Array.isArray(merges)) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].merges must be an array of ranges.`, {
				hint: "For example merges: [\"A1:C1\"].",
				details: { option: "merges" }
			});
			for (const m of merges) sheet.merge(m);
		}
		const filter = s.get("autoFilter") ?? false;
		let autoFilter = false;
		if (typeof filter === "string") {
			sheet.autoFilter(filter);
			autoFilter = sheet.filter ?? false;
		} else if (typeof filter === "boolean") autoFilter = filter;
		else throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].autoFilter must be true, false or a range.`, {
			hint: "Use autoFilter: true for a filter over all rows.",
			details: { option: "autoFilter" }
		});
		if (autoFilter === true && sheet.tableList.length > 0) throw new XlvError("XLV_E_INVALID_TABLE", `sheets[${String(i)}] has autoFilter: true and tables; the filter would cover the table "${sheet.tableList[0].name}".`, {
			location: { sheet: sheet.name },
			hint: "Tables have their own filters. Remove autoFilter, or give it a range outside the tables.",
			details: { name: sheet.tableList[0].name }
		});
		const rows = s.get("rows");
		if (!isIterable(rows)) throw new XlvError("XLV_E_INVALID_ROWS", `sheets[${String(i)}].rows is not an iterable.`, {
			hint: "Pass an array, a generator or an async iterable of rows.",
			details: { row: 0 }
		});
		const links = s.get("hyperlinks");
		if (links !== void 0) {
			if (!Array.isArray(links)) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].hyperlinks must be an array.`, {
				hint: "For example hyperlinks: [{ ref: \"A1\", url: \"https://example.com\" }].",
				details: { option: "hyperlinks" }
			});
			for (const l of links) {
				if (typeof l !== "object" || l === null) throw new XlvError("XLV_E_INVALID_OPTION", `sheets[${String(i)}].hyperlinks must hold objects.`, {
					hint: "For example { ref: \"A1\", url: \"https://example.com\" }.",
					details: { option: "hyperlinks" }
				});
				const { ref, ...target } = l;
				sheet.link(ref, target);
			}
		}
		for (const v of listOf("images", "{ image: 0, anchor: \"B2:D6\" }")) {
			const { image, anchor, ...options } = v;
			sheet.addImage(image, anchor, options);
		}
		if (s.has("background")) sheet.background(s.get("background"));
		for (const v of listOf("charts", "{ type: \"column\", anchor: \"E2:L18\", series: [{ values: \"Data!B2:B13\" }] }")) sheet.addChart(v);
		return {
			spec: raw,
			sheet,
			autoFilter,
			rows,
			filterRange: null,
			rels: new SheetRels(),
			notes: null,
			tables: [],
			drawing: null
		};
	};
	const noVisibleSheet = () => new XlvError("XLV_E_INVALID_INPUT", "All sheets of the workbook are hidden.", { hint: "A workbook needs at least one visible sheet." });
	const prepared = dynamic ? [] : specs.map(prepare);
	if (!dynamic && !prepared.some((p) => p.sheet.state === "visible")) throw noVisibleSheet();
	if (!dynamic) addNames();
	if (!dynamic) {
		const lookup = chartSheetNames(prepared.map((q) => q.sheet.name));
		for (const p of prepared) for (const c of p.sheet.chartList) resolveChartSheets(c.chart, lookup, p.sheet.name);
	}
	/** Charts in workbook order (CH-01); their parts are written once every sheet name is known. */
	const chartQueue = [];
	const strings = shared ? createSharedStrings() : null;
	const styleCache = /* @__PURE__ */ new WeakMap();
	const dateStyle = /* @__PURE__ */ new Map();
	const normalized = (style) => {
		let n = styleCache.get(style);
		if (n === void 0) {
			n = normalizeStyle(style);
			styleCache.set(style, n);
		}
		return n;
	};
	/** First 1024-id block of VML shape ids for the next sheet with notes. */
	let shapeBlock = 1;
	async function* sheetData(p, selected, n) {
		const name = p.sheet.name;
		let buffer = sheetHead({
			freeze: p.sheet.frozen,
			columns: p.sheet.columnMap,
			tabColor: p.sheet.tab,
			view: p.sheet.viewSettings,
			print: p.sheet.print,
			properties: p.sheet.sheetProps
		}, selected, null);
		let r = 0;
		let maxCol = 0;
		const groups = new FormulaGroups(name);
		for await (const raw of iterate(p.rows)) {
			r++;
			if (r > 1048576) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `Sheet "${name}" has more than ${String(MAX_ROWS)} rows.`, {
				location: { sheet: name },
				hint: "Split the rows over several sheets.",
				details: {
					limit: "rows",
					value: MAX_ROWS
				}
			});
			let values;
			let rowStyle;
			let styles;
			let info;
			if (Array.isArray(raw)) values = raw;
			else {
				const ro = readOptions(raw, [
					"values",
					"style",
					"styles",
					"rowStyle",
					"height",
					"hidden",
					"outlineLevel",
					"collapsed",
					"hyperlinks",
					"notes"
				], `row ${String(r)} of sheet "${name}"`);
				const v = ro.get("values");
				if (!Array.isArray(v)) throw new XlvError("XLV_E_INVALID_ROWS", `Row ${String(r)} of sheet "${name}" is not an array of values.`, {
					location: { sheet: name },
					hint: "Yield arrays of values, or { values: [...] } objects.",
					details: { row: r }
				});
				values = v;
				rowStyle = ro.get("style");
				styles = ro.get("styles");
				const format = {};
				for (const key of [
					"height",
					"hidden",
					"outlineLevel",
					"collapsed"
				]) if (ro.get(key) !== void 0) format[key] = ro.get(key);
				if (ro.get("rowStyle") !== void 0) format.style = ro.get("rowStyle");
				if (Object.keys(format).length > 0) {
					p.sheet.row(r, format);
					info = p.sheet.rowMap.get(r);
					p.sheet.rowMap.delete(r);
				}
				for (const key of ["hyperlinks", "notes"]) {
					const list = ro.get(key);
					if (list === void 0) continue;
					if (!Array.isArray(list)) throw new XlvError("XLV_E_INVALID_OPTION", `${key} of row ${String(r)} of sheet "${name}" must be an array, by column position.`, {
						location: { sheet: name },
						hint: key === "notes" ? "For example notes: [undefined, \"Checked\"] for a note on column B." : "For example hyperlinks: [{ url: \"https://example.com\" }] for a link on column A.",
						details: { option: key }
					});
					list.forEach((item, c) => {
						if (item === void 0 || item === null) return;
						const ref = formatCellRef(r, c + 1);
						if (key === "notes") p.sheet.note(ref, item);
						else p.sheet.link(ref, item);
					});
				}
			}
			if (values.length > 16384) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `Row ${String(r)} of sheet "${name}" has more than ${String(MAX_COLUMNS)} values.`, {
				location: { sheet: name },
				hint: "A sheet has 16 384 columns.",
				details: {
					limit: "columns",
					value: MAX_COLUMNS
				}
			});
			let cells = "";
			for (let c = 0; c < values.length; c++) {
				const ref = formatCellRef(r, c + 1);
				const converted = convertCellInput(values[c], {
					sheet: name,
					dateSystem: book.dateSystem
				}, ref);
				const style = styles?.[c] ?? rowStyle;
				let base = style === void 0 ? {} : normalized(style);
				let id = style === void 0 ? 0 : book.intern(base);
				if (converted.isDate && base.numFmt === void 0) {
					const key = `${String(id)}|${converted.hasTime ? "t" : "d"}`;
					let dated = dateStyle.get(key);
					if (dated === void 0) {
						base = {
							...base,
							numFmt: converted.hasTime ? DATE_TIME_FORMAT : DATE_FORMAT
						};
						dated = book.intern(base);
						dateStyle.set(key, dated);
					}
					id = dated;
				}
				if (converted.stored === null && id === 0) continue;
				cells += cellXml(ref, {
					value: converted.stored,
					style: id
				}, null, strings, groups.attributes(r, c + 1, ref, converted.stored));
				if (c + 1 > maxCol) maxCol = c + 1;
			}
			const start = rowStart(r, info);
			if (cells !== "") buffer += `${start}>${cells}</row>`;
			else if (info !== void 0) buffer += `${start}/>`;
			if (buffer.length >= CHUNK_CHARS) {
				yield encoder.encode(buffer);
				buffer = "";
			}
		}
		if (p.autoFilter === true) p.filterRange = r > 0 && maxCol > 0 ? {
			start: {
				row: 1,
				col: 1
			},
			end: {
				row: r,
				col: maxCol
			}
		} : null;
		else if (p.autoFilter !== false) p.filterRange = p.autoFilter;
		p.notes = noteParts(sortedNotes(p.sheet.noteMap), n, p.rels, shapeBlock);
		p.sheet.noteMap.clear();
		if (p.notes !== null) shapeBlock += p.notes.blocks;
		const rules = sheetRulesXml({
			conditionalFormats: p.sheet.conditionalList,
			dataValidations: p.sheet.validationList
		}, n);
		const frames = p.sheet.chartList.map((c) => {
			chartQueue.push({
				chart: c.chart,
				sheet: p.sheet.name
			});
			return {
				...resolveAnchor(c.anchor, metricsOf(p.sheet.columnMap, p.sheet.rowMap)),
				name: c.chart.name,
				description: c.chart.description,
				target: `../charts/chart${String(chartQueue.length)}.xml`
			};
		});
		const drawing = drawingParts(p.sheet.imageList.map((i) => resolveImage(i, metricsOf(p.sheet.columnMap, p.sheet.rowMap))), p.sheet.backgroundImage, media, ++drawingCount, p.rels, frames);
		if (drawing.part === null) drawingCount--;
		p.drawing = drawing.part;
		const sheetTables = tableParts(p.sheet.tableList, tableCount + 1, p.rels);
		tableCount += sheetTables.parts.length;
		p.tables = sheetTables.parts;
		buffer += sheetTail(p.filterRange, p.sheet.mergeList, rules.main + hyperlinksXml([...p.sheet.linkMap.values()], p.rels) + printXml(p.sheet.print) + drawing.drawing + (p.notes?.legacyDrawing ?? "") + drawing.picture + sheetTables.element + rules.ext, sheetProtectionXml(p.sheet.print.protection));
		yield encoder.encode(buffer);
	}
	/** The Comments and VML parts of a sheet with notes, after the sheet has been written. */
	function* noteEntries(p) {
		const parts = p.notes;
		if (parts === null) return;
		yield {
			name: parts.comments.name,
			data: encoder.encode(parts.comments.xml)
		};
		yield {
			name: parts.vml.name,
			data: encoder.encode(parts.vml.xml)
		};
	}
	/** Tables are numbered across the workbook (TD-02). */
	let tableCount = 0;
	/** Drawings parts are numbered across the workbook (TD-35). */
	let drawingCount = 0;
	const tableTypes = () => prepared.flatMap((p) => p.tables.map((t) => [t.name, TABLE_CONTENT_TYPE]));
	const noteSheets = () => prepared.flatMap((p, i) => p.notes === null ? [] : [i + 1]);
	/**
	* The sheet's relationships part, if it has relationships. The ZIP stream reads an entry's data
	* completely before it asks for the next entry, so the sheet has been written by now.
	*/
	function* sheetRelsEntry(p, n) {
		const xml = p.rels.xml();
		if (xml !== null) yield {
			name: sheetRelsPartName(n),
			data: encoder.encode(xml)
		};
	}
	/** The parts written after the sheets: filter ranges and styles are known only then. */
	function* closingParts() {
		checkBookViews(book.bookViews, prepared.map((p) => p.sheet.name));
		yield {
			name: "xl/workbook.xml",
			data: encoder.encode(workbookXml(prepared.map((p) => ({
				name: p.sheet.name,
				state: p.sheet.state,
				sheetId: p.sheet.sheetId,
				autoFilter: p.filterRange
			})), book.dateSystem, allNames(), book.workbookProtection, {
				bookViews: book.bookViews,
				calculation: book.calcProps
			}))
		};
		yield {
			name: "xl/_rels/workbook.xml.rels",
			data: encoder.encode(workbookRelsXml(prepared.length, shared))
		};
		yield {
			name: "xl/styles.xml",
			data: encoder.encode(buildStyles(book.styles(), true, book.dxfStyles(), book.fontDefault).xml)
		};
		if (core !== null) yield {
			name: CORE_PROPERTIES_PART,
			data: encoder.encode(core)
		};
		const app = appXml();
		if (app !== null) yield {
			name: APP_PROPERTIES_PART,
			data: encoder.encode(app)
		};
		if (strings) yield {
			name: "xl/sharedStrings.xml",
			data: encoder.encode(strings.xml())
		};
	}
	/** `docProps/app.xml` (W2-52); known once the sheets and names are. */
	const appXml = () => appPropertiesXml(book.properties, prepared.map((p) => p.sheet.name), nameTitles(allNames()));
	/** The user's defined names and the print area and titles of every sheet (SR-30). */
	const allNames = () => [...book.definedNames, ...prepared.flatMap((p) => printNames(p.sheet.name, p.sheet.print))];
	async function* entries() {
		let selected = false;
		const sheets = dynamic ? iterate(specs) : prepared;
		let n = 0;
		for await (const raw of sheets) {
			let p;
			if (dynamic) {
				p = prepare(raw, prepared.length);
				prepared.push(p);
			} else p = raw;
			n++;
			const select = !selected && p.sheet.state === "visible";
			if (select) selected = true;
			yield {
				name: `xl/worksheets/sheet${String(n)}.xml`,
				data: sheetData(p, select, n)
			};
			yield* noteEntries(p);
			for (const t of p.tables) yield {
				name: t.name,
				data: encoder.encode(t.xml)
			};
			if (p.drawing !== null) {
				yield {
					name: p.drawing.name,
					data: encoder.encode(p.drawing.xml)
				};
				yield {
					name: p.drawing.relsName,
					data: encoder.encode(p.drawing.rels)
				};
			}
			yield* sheetRelsEntry(p, n);
		}
		if (prepared.length === 0) throw new XlvError("XLV_E_INVALID_INPUT", "The workbook has no sheets.", { hint: "Add at least one sheet before closing the stream." });
		if (!selected) throw noVisibleSheet();
		if (dynamic) addNames();
		const names = chartSheetNames(prepared.map((p) => p.sheet.name));
		const chartParts = chartQueue.map((c, k) => ({
			name: chartPartName(k + 1),
			xml: chartPartXml(resolveChartSheets(c.chart, names, c.sheet), null)
		}));
		const app = appXml() !== null;
		const overrides = [
			...tableTypes(),
			...prepared.flatMap((p) => p.drawing === null ? [] : [[p.drawing.name, DRAWING_CONTENT_TYPE]]),
			...chartParts.map((c) => [c.name, CHART_CONTENT_TYPE]),
			...media.entries().map((m) => [m.name, IMAGE_CONTENT_TYPES[m.name.slice(m.name.lastIndexOf(".") + 1)]])
		];
		yield {
			name: "[Content_Types].xml",
			data: encoder.encode(contentTypesXml(prepared.length, shared, core !== null, app, noteSheets(), overrides))
		};
		yield {
			name: "_rels/.rels",
			data: encoder.encode(rootRelsXml(core !== null, app))
		};
		yield* media.entries();
		for (const c of chartParts) yield {
			name: c.name,
			data: encoder.encode(c.xml)
		};
		yield* closingParts();
	}
	return createZipStream(entries(), level);
}
const REPORT_HEADER = {
	font: { bold: true },
	fill: { color: "D9E1F2" },
	border: { bottom: { style: "thin" } }
};
/**
* Streams row objects into a formatted XLSX table, like `writeTable` but with constant memory.
* Because rows are not read in advance, the columns must be given, and widths are not fitted to
* the content (set `width` per column).
*
* @example
* ```ts
* import { writeTableStream } from 'xlvane';
*
* async function* orders() {
*   for (let id = 1; id <= 500_000; id++) yield { id, customer: `C${id % 100}`, total: id * 1.5 };
* }
* const stream = writeTableStream(orders(), { columns: ['id', 'customer', { key: 'total', numFmt: '#,##0.00' }] });
* ```
*/
function writeTableStream(rows, options) {
	const o = readOptions(options, [
		"columns",
		"sheet",
		"preset",
		"dateSystem",
		"sharedStrings"
	], "writeTableStream options");
	const spec = o.get("columns");
	if (!Array.isArray(spec) || spec.length === 0) throw new XlvError("XLV_E_INVALID_OPTION", "writeTableStream needs the \"columns\" option.", {
		hint: "Rows are not read in advance, so list the columns, for example columns: [\"id\", \"name\", { key: \"total\", numFmt: \"#,##0.00\" }].",
		details: { option: "columns" }
	});
	const columns = spec.map((entry, i) => {
		if (typeof entry === "string") return {
			key: entry,
			header: entry,
			numFmt: void 0,
			width: void 0
		};
		const c = readOptions(entry, [
			"key",
			"header",
			"numFmt",
			"width"
		], `columns[${String(i)}]`);
		const key = c.get("key");
		if (typeof key !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `columns[${String(i)}] has no "key".`, {
			hint: "Give each column the property it shows, e.g. { key: \"total\" }.",
			details: { option: "columns" }
		});
		return {
			key,
			header: c.get("header") ?? key,
			numFmt: c.get("numFmt"),
			width: c.get("width")
		};
	});
	const preset = o.get("preset") ?? "report";
	if (preset !== "report" && preset !== "plain") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"preset\" must be \"report\" or \"plain\".", {
		hint: "Use preset: \"report\" or \"plain\".",
		details: { option: "preset" }
	});
	if (!isIterable(rows)) throw new XlvError("XLV_E_INVALID_ROWS", "writeTableStream expects an iterable of row objects.", {
		hint: "Pass an array, a generator or an async iterable.",
		details: { row: 0 }
	});
	const cellStyles = columns.map((c) => c.numFmt === void 0 ? void 0 : { numFmt: c.numFmt });
	async function* table() {
		yield {
			values: columns.map((c) => c.header),
			...preset === "report" ? { style: REPORT_HEADER } : {}
		};
		let n = 0;
		for await (const row of iterate(rows)) {
			n++;
			if (typeof row !== "object" || row === null || Array.isArray(row)) throw new XlvError("XLV_E_INVALID_ROWS", `Row ${String(n)} is not an object.`, {
				hint: "Yield objects keyed by column, such as { id: 1, name: \"Alice\" }.",
				details: { row: n + 1 }
			});
			yield {
				values: columns.map((c) => {
					const v = Object.prototype.hasOwnProperty.call(row, c.key) ? row[c.key] : void 0;
					if (v === void 0 || v === null) return null;
					if (typeof v === "number" || typeof v === "string" || typeof v === "boolean" || v instanceof Date) return v;
					throw new XlvError("XLV_E_INVALID_ROWS", `Row ${String(n)}, column "${c.key}" holds ${typeof v}, which is not a cell value.`, {
						hint: "Cell values are numbers, strings, booleans, Dates or null.",
						details: {
							row: n + 1,
							key: c.key
						}
					});
				}),
				styles: cellStyles
			};
		}
	}
	const shared = o.get("sharedStrings");
	const dateSystem = o.get("dateSystem");
	return writeWorkbookStream({
		...dateSystem === void 0 ? {} : { dateSystem },
		...shared === void 0 ? {} : { sharedStrings: shared },
		sheets: [{
			name: o.get("sheet") ?? "Sheet1",
			columns: columns.map((c) => c.width === void 0 ? void 0 : { width: c.width }),
			...preset === "report" ? {
				freeze: { rows: 1 },
				autoFilter: true
			} : {},
			rows: table()
		}]
	});
}
export { readTableStream as a, openWorkbookStream as i, writeWorkbookStream as n, STREAM_LIMIT_DEFAULTS as r, writeTableStream as t };
