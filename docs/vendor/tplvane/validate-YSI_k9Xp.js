import { Inflate, deflateSync } from "../fflate/browser.js";
import { deleteRows, editCells, insertRows, isXlvError, readWorkbook, shiftFormula } from "../xlvane/index.js";
// src/version.ts
/**
* The version of the `tplvane` package.
*
* @example
* ```ts
* import { VERSION } from 'tplvane';
* console.log(VERSION);
* ```
*/
const VERSION$1 = "1.0.0";
// src/errors.ts
/** Base URL of the error catalogue. Every code has a heading (and so an anchor) there. */
const DOCS_URL = "https://github.com/DevXord/tplvane/blob/main/docs/errors.md";
/**
* Every error code tplvane can throw. Codes are stable: a code is never reused for a different
* condition. Each one is described in docs/errors.md.
*
* @example
* ```ts
* import { ERROR_CODES } from 'tplvane';
* console.log(ERROR_CODES.includes('TV_E_TAG_UNCLOSED')); // true
* ```
*/
const ERROR_CODES = [
	"TV_E_INVALID_INPUT",
	"TV_E_INVALID_OPTION",
	"TV_E_LIMIT_EXCEEDED",
	"TV_E_NOT_ZIP",
	"TV_E_UNSUPPORTED_CONTAINER",
	"TV_E_ZIP_CORRUPT",
	"TV_E_ZIP_UNSUPPORTED",
	"TV_E_XML_MALFORMED",
	"TV_E_XML_DTD_FORBIDDEN",
	"TV_E_XML_ENTITY_FORBIDDEN",
	"TV_E_XML_ENCODING",
	"TV_E_PART_MISSING",
	"TV_E_PART_INVALID",
	"TV_E_NOT_SUPPORTED",
	"TV_E_TEMPLATE_INVALID",
	"TV_E_TAG_UNCLOSED",
	"TV_E_TAG_UNOPENED",
	"TV_E_TAG_EMPTY",
	"TV_E_TAG_SYNTAX",
	"TV_E_TAG_FORBIDDEN_NAME",
	"TV_E_SECTION_UNCLOSED",
	"TV_E_SECTION_UNOPENED",
	"TV_E_SECTION_MISMATCH",
	"TV_E_SECTION_PLACEMENT",
	"TV_E_TAG_PLACEMENT",
	"TV_E_FORMATTER_UNKNOWN",
	"TV_E_FORMATTER_ARGS",
	"TV_E_DATA_INVALID",
	"TV_E_MISSING_FIELD",
	"TV_E_INVALID_VALUE",
	"TV_E_IMAGE_FETCH",
	"TV_E_USAGE",
	"TV_E_FILE_READ",
	"TV_E_FILE_EXISTS",
	"TV_E_FILE_WRITE",
	"TV_E_INTERNAL"
];
/**
* Every warning code tplvane can report. Warnings never stop the work; they describe something
* in the template that the caller may want to know about. Each one is described in docs/errors.md.
*
* @example
* ```ts
* import { WARNING_CODES } from 'tplvane';
* console.log(WARNING_CODES.includes('TV_W_SCHEMA_CONFLICT')); // true
* ```
*/
const WARNING_CODES = [
	"TV_W_SCHEMA_CONFLICT",
	"TV_W_PART_MISSING",
	"TV_W_STYLE_MISSING",
	"TV_W_DATA_SCHEMA"
];
/**
* The only error type thrown by tplvane for problems with templates, data or options. It says
* what happened (`message`), where (`location`), what to do next (`hint`) and links to the
* description of the code (`docsUrl`). `toJSON()` returns the same information as data. When a
* template has several syntax problems, one `TV_E_TEMPLATE_INVALID` lists them all in `errors`.
*
* @example
* ```ts
* import { inspect, isTvError } from 'tplvane';
*
* try {
*   inspect(new Uint8Array([1, 2, 3]));
* } catch (error) {
*   if (isTvError(error)) console.log(error.code, error.hint); // "TV_E_NOT_ZIP", "..."
* }
* ```
*/
var TvError = class extends Error {
	name = "TvError";
	/** Stable error code, see docs/errors.md. */
	code;
	/** Where the problem is. Empty when the problem is not tied to a place in the file. */
	location;
	/** What to do next. */
	hint;
	/** Link to the description of the error code. */
	docsUrl;
	/** Machine-readable details. */
	details;
	/** The single problems behind `TV_E_TEMPLATE_INVALID` and `TV_E_DATA_INVALID`; empty for other codes. */
	errors;
	constructor(code, message, init) {
		super(message, init.cause === void 0 ? void 0 : { cause: init.cause });
		this.code = code;
		this.location = { ...init.location };
		this.hint = init.hint;
		this.docsUrl = `${DOCS_URL}#${code.toLowerCase()}`;
		this.details = { ...init.details };
		this.errors = [...init.errors ?? []];
	}
	/** Returns the error as plain JSON-safe data. */
	toJSON() {
		return {
			name: this.name,
			code: this.code,
			message: this.message,
			location: { ...this.location },
			hint: this.hint,
			docsUrl: this.docsUrl,
			details: { ...this.details },
			errors: this.errors.map((error) => error.toJSON())
		};
	}
};
/**
* Returns `true` if `value` is a {@link TvError}.
*
* @example
* ```ts
* import { isTvError } from 'tplvane';
* console.log(isTvError(new Error('x'))); // false
* ```
*/
function isTvError(value) {
	return value instanceof TvError;
}
/** One error of `code` that lists `problems` in `errors` (render, inspect). @internal */
function aggregate(code, problems, truncated) {
	const first = problems[0];
	const count = problems.length;
	const what = code === "TV_E_TEMPLATE_INVALID" ? "template" : "data";
	return new TvError(code, count === 1 ? `The ${what} has a problem: ${first.message}` : `The ${what} has ${String(count)}${truncated ? " or more" : ""} problems. The first: ${first.message}`, {
		location: first.location,
		hint: `Fix the problems listed in "errors"; each one has its place in the template${code === "TV_E_DATA_INVALID" ? " and names the field" : ""}. The first: ${first.hint}`,
		details: {
			count,
			truncated
		},
		errors: problems
	});
}
// src/util/suggest.ts
/** Names longer than this are never compared, which keeps the cost bounded on hostile input. */
const MAX_COMPARED_LENGTH = 256;
/**
* Returns the candidates that are close to `input` ("Did you mean ...?"), best match first.
* Comparison ignores case and uses the optimal string alignment distance (Levenshtein plus
* transpositions); a candidate matches when the distance is at most a third of the longer name.
*/
function suggest(input, candidates, max = 3) {
	const needle = input.toLowerCase();
	const scored = [];
	candidates.forEach((name, order) => {
		const hay = name.toLowerCase();
		if (needle === hay) {
			scored.push({
				name,
				distance: 0,
				order
			});
			return;
		}
		if (needle.length > MAX_COMPARED_LENGTH || hay.length > MAX_COMPARED_LENGTH) return;
		const threshold = Math.max(1, Math.floor(Math.max(needle.length, hay.length) / 3));
		if (Math.abs(needle.length - hay.length) > threshold) return;
		const distance = osaDistance(needle, hay);
		if (distance <= threshold) scored.push({
			name,
			distance,
			order
		});
	});
	scored.sort((a, b) => a.distance - b.distance || a.order - b.order);
	return scored.slice(0, max).map((s) => s.name);
}
function osaDistance(a, b) {
	let prev2 = [];
	let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
	for (let i = 1; i <= a.length; i++) {
		const row = [i];
		for (let j = 1; j <= b.length; j++) {
			const cost = a[i - 1] === b[j - 1] ? 0 : 1;
			let value = Math.min((prev[j] ?? 0) + 1, (row[j - 1] ?? 0) + 1, (prev[j - 1] ?? 0) + cost);
			if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) value = Math.min(value, (prev2[j - 2] ?? 0) + 1);
			row.push(value);
		}
		prev2 = prev;
		prev = row;
	}
	return prev[b.length] ?? 0;
}
// src/limits.ts
const MiB = 1048576;
/**
* The default {@link Limits}. Frozen; pass `{ limits: { ... } }` to change individual values.
*
* @example
* ```ts
* import { DEFAULT_LIMITS } from 'tplvane';
* console.log(DEFAULT_LIMITS.maxEntries, DEFAULT_LIMITS.maxTagLength); // 10000 1024
* ```
*/
const DEFAULT_LIMITS = Object.freeze({
	maxInputBytes: 256 * MiB,
	maxEntries: 1e4,
	maxEntryBytes: 256 * MiB,
	maxTotalBytes: 512 * MiB,
	maxCompressionRatio: 200,
	compressionRatioMinBytes: 1 * MiB,
	maxXmlDepth: 128,
	maxXmlAttributes: 256,
	maxXmlNameLength: 1024,
	maxXmlTextLength: 8 * MiB,
	maxXmlTreeElements: 1e6,
	maxTagLength: 1024,
	maxTags: 1e5,
	maxSectionDepth: 32,
	maxErrors: 100,
	maxExpressionDepth: 32,
	maxIterations: 1e5,
	maxOutputBytes: 256 * MiB,
	maxImageBytes: 20 * MiB,
	maxImages: 1e3,
	maxImagePixels: 1e8,
	maxMarkdownLength: 1 * MiB
});
const LIMIT_NAMES = Object.keys(DEFAULT_LIMITS);
/**
* Validates user-supplied limits and merges them over the defaults. Only own, known keys are
* read, so keys such as `__proto__` cannot change anything; unknown keys are rejected.
*/
function resolveLimits(input) {
	const limits = { ...DEFAULT_LIMITS };
	if (input === void 0) return limits;
	if (input === null || typeof input !== "object" || Array.isArray(input)) throw new TvError("TV_E_INVALID_OPTION", `Option "limits" must be an object, got ${describe$2(input)}.`, { hint: "Pass an object such as { limits: { maxEntries: 10000 } }, or leave \"limits\" out to use the defaults." });
	for (const key of Object.keys(input)) if (!LIMIT_NAMES.includes(key)) {
		const close = suggest(key, LIMIT_NAMES);
		throw new TvError("TV_E_INVALID_OPTION", `Unknown limit "${key}".`, {
			hint: close.length > 0 ? `Did you mean ${close.map((c) => `"${c}"`).join(" or ")}? Known limits: ${LIMIT_NAMES.join(", ")}.` : `Known limits: ${LIMIT_NAMES.join(", ")}.`,
			details: {
				option: key,
				known: LIMIT_NAMES
			}
		});
	}
	for (const name of LIMIT_NAMES) {
		if (!Object.hasOwn(input, name)) continue;
		const value = input[name];
		limits[name] = validate$1(name, value);
	}
	return limits;
}
function validate$1(name, value) {
	const isRatio = name === "maxCompressionRatio";
	if (!(typeof value === "number" && (value === Number.POSITIVE_INFINITY || (isRatio ? Number.isFinite(value) && value >= 1 : Number.isSafeInteger(value) && value > 0)))) throw new TvError("TV_E_INVALID_OPTION", isRatio ? `Limit "${name}" must be a number of at least 1 or Infinity, got ${describe$2(value)}.` : `Limit "${name}" must be a whole number greater than 0 or Infinity, got ${describe$2(value)}.`, {
		hint: `Pass for example { limits: { ${name}: ${String(DEFAULT_LIMITS[name])} } }, or leave it out to use the default.`,
		details: { option: name }
	});
	return value;
}
function describe$2(value) {
	if (value === null) return "null";
	if (Array.isArray(value)) return "an array";
	if (typeof value === "string") return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
	if (typeof value === "object" || typeof value === "function") return `a value of type ${typeof value}`;
	return String(value);
}
/** A `TV_E_LIMIT_EXCEEDED` error for `limit`, with the hint used by every template limit check. */
function limitError(limit, value, message, location = {}) {
	return new TvError("TV_E_LIMIT_EXCEEDED", message, {
		location,
		hint: `If you trust this template, raise the limit, for example { limits: { ${limit}: ${String(Number.isFinite(value) ? value * 4 : value)} } }.`,
		details: {
			limit,
			value
		}
	});
}
// src/schema.ts
const VALUE_TYPES = [
	"string",
	"number",
	"boolean"
];
var SchemaBuilder = class {
	root = {
		t: "object",
		props: /* @__PURE__ */ new Map()
	};
	warnings = [];
	/** Merges `incoming` into `existing`; records a conflict warning at `where` for uses that do not fit. */
	merge(existing, incoming, path, where) {
		if (existing.t === "boolean") return incoming;
		if (incoming.t === "boolean") return existing;
		if (existing.t === "object" && incoming.t === "object") {
			for (const [name, prop] of incoming.props) this.mergeProp(existing, name, prop, `${path}.${name}`);
			return existing;
		}
		if (existing.t === "array" && incoming.t === "array") {
			if (incoming.items !== void 0) existing.items = existing.items === void 0 ? incoming.items : this.merge(existing.items, incoming.items, `${path}[]`, where);
			return existing;
		}
		if (existing.t === incoming.t && existing.t !== "conflict") return existing;
		if (existing.t === "conflict") {
			const same = existing.options.findIndex((o) => o.t === incoming.t);
			if (same === -1) existing.options.push(incoming);
			else existing.options[same] = this.merge(existing.options[same], incoming, path, where);
			return existing;
		}
		this.warnings.push({
			code: "TV_W_SCHEMA_CONFLICT",
			message: `"${path}" is used in ways that do not fit together: as ${describe$1(existing)} and as ${describe$1(incoming)}.`,
			location: where,
			hint: "Use different names for different data, or use the name the same way everywhere. The schema allows each of the uses (anyOf)."
		});
		return {
			t: "conflict",
			options: [existing, incoming]
		};
	}
	mergeProp(obj, name, prop, path) {
		const existing = obj.props.get(name);
		if (existing === void 0) {
			obj.props.set(name, prop);
			return;
		}
		if (existing.node.t === "boolean" && prop.node.t !== "boolean") {
			existing.first = prop.first;
			existing.role = prop.role;
		}
		existing.node = this.merge(existing.node, prop.node, path, prop.first);
		existing.required ||= prop.required;
		existing.uses += prop.uses;
	}
	/** The object in which names of `scope` are resolved, creating or merging the loop item. */
	objectOf(scope, where) {
		if (scope.kind === "root") return scope.obj;
		const item = scope.array.items;
		if (item === void 0) {
			const obj = {
				t: "object",
				props: /* @__PURE__ */ new Map()
			};
			scope.array.items = obj;
			return obj;
		}
		if (item.t === "object") return item;
		if (item.t === "conflict") {
			const existing = item.options.find((o) => o.t === "object");
			if (existing !== void 0) return existing;
		}
		const obj = {
			t: "object",
			props: /* @__PURE__ */ new Map()
		};
		const merged = this.merge(item, obj, `${scope.path}[]`, where);
		scope.array.items = merged;
		return merged.t === "object" ? merged : merged.t === "conflict" ? merged.options.find((o) => o === obj) : obj;
	}
	/** Adds a use of `tag` with schema `leaf` in `scope`; returns the node that now holds it. */
	add(scope, tag, segments, leaf, required, role) {
		if (segments.length === 0) {
			if (scope.kind === "loop") scope.array.items = scope.array.items === void 0 ? leaf : this.merge(scope.array.items, leaf, `${scope.path}[]`, tag.location);
			return leaf;
		}
		let obj = this.objectOf(scope, tag.location);
		let path = scope.kind === "loop" ? `${scope.path}[]` : "";
		for (let i = 0; i < segments.length; i++) {
			const name = segments[i];
			path = path === "" ? name : `${path}.${name}`;
			const last = i === segments.length - 1;
			const node = last ? leaf : {
				t: "object",
				props: /* @__PURE__ */ new Map()
			};
			this.mergeProp(obj, name, {
				node,
				required,
				first: tag.location,
				role: last ? role : "",
				uses: 1
			}, path);
			const held = obj.props.get(name).node;
			if (last) return held;
			if (held.t === "object") obj = held;
			else if (held.t === "conflict") obj = held.options.find((o) => o.t === "object") ?? obj;
			else return held;
		}
		return leaf;
	}
	/**
	* Adds the paths an expression reads: `required` values (a value tag, an operand of a
	* comparison), `optional` values (before a `default` formatter) and `truth` tests (operands of
	* !, && and ||, and sections over an expression), which are optional conditions.
	*/
	addExpr(scope, tag, e, mode) {
		switch (e.t) {
			case "path":
				if (e.segments.length > 1 && e.segments[e.segments.length - 1] === "length") {
					this.add(scope, tag, e.segments.slice(0, -1), {
						t: "array",
						items: void 0
					}, mode === "required", "List (its length is read)");
					return;
				}
				if (mode === "truth") this.add(scope, tag, e.segments, { t: "boolean" }, false, "Condition");
				else this.add(scope, tag, e.segments, { t: "value" }, mode === "required", "Value");
				return;
			case "not":
				this.addExpr(scope, tag, e.e, "truth");
				return;
			case "and":
			case "or":
				if (mode === "truth") {
					this.addExpr(scope, tag, e.l, "truth");
					this.addExpr(scope, tag, e.r, "truth");
				} else {
					this.addExpr(scope, tag, e.l, mode);
					this.addExpr(scope, tag, e.r, "optional");
				}
				return;
			case "compare":
				this.addExpr(scope, tag, e.l, "required");
				this.addExpr(scope, tag, e.r, "required");
				return;
			case "pipe":
				if (e.formatters.some((f) => f.name === "link") && e.e.t === "path") {
					this.add(scope, tag, e.e.segments, { t: "url" }, !e.formatters.some((f) => f.name === "default"), "Link address (http, https or mailto)");
					return;
				}
				this.addExpr(scope, tag, e.e, e.formatters.some((f) => f.name === "default") ? "optional" : mode === "truth" ? "required" : mode);
				return;
		}
	}
	/** Adds the tags and sections of `children` in `scope`. */
	addChildren(children, scope) {
		for (const child of children) if ("tag" in child) {
			const { info, parsed } = child.tag;
			if (info.kind === "image" || info.kind === "markdown") this.add(scope, info, parsed.segments ?? [], info.kind === "image" ? { t: "image" } : { t: "markdown" }, true, info.kind === "image" ? "Image" : "Markdown text");
			else this.addExpr(scope, info, fullExpr$1(child.tag), "required");
		} else this.addSection(child.section, scope);
	}
	addSection(section, scope) {
		const { info, parsed } = section.open;
		const plain = parsed.segments !== null && parsed.formatters.length === 0;
		if (section.kind !== "loop") {
			if (plain) this.add(scope, info, parsed.segments ?? [], { t: "boolean" }, false, section.kind === "inverted" ? "Shown when false or empty" : "Condition");
			else this.addExpr(scope, info, fullExpr$1(section.open), "truth");
			this.addChildren(section.children, scope);
			return;
		}
		const array = {
			t: "array",
			items: void 0
		};
		const held = this.add(scope, info, parsed.segments ?? [], array, true, "Loop");
		const target = held.t === "array" ? held : held.t === "conflict" ? held.options.find((o) => o.t === "array") ?? array : array;
		const path = (scope.kind === "loop" ? `${scope.path}[].` : "") + info.path;
		this.addChildren(section.children, {
			kind: "loop",
			array: target,
			path
		});
	}
};
/** The whole expression of a tag, its formatters included. */
function fullExpr$1(tag) {
	return tag.parsed.formatters.length === 0 ? tag.parsed.expr : {
		t: "pipe",
		e: tag.parsed.expr,
		formatters: tag.parsed.formatters
	};
}
function describe$1(node) {
	switch (node.t) {
		case "value": return "a value";
		case "object": return "an object with fields";
		case "array": return "a loop";
		case "boolean": return "a condition";
		case "image": return "an image";
		case "markdown": return "Markdown text";
		case "url": return "a link address";
		case "conflict": return "several kinds of data";
	}
}
function defineAll(entries) {
	const out = {};
	for (const [name, value] of entries) Object.defineProperty(out, name, {
		value,
		enumerable: true,
		writable: true,
		configurable: true
	});
	return out;
}
function usedIn(prop) {
	const where = prop.first.description ?? prop.first.part ?? "the template";
	const more = prop.uses > 1 ? ` and ${String(prop.uses - 1)} more ${prop.uses === 2 ? "place" : "places"}` : "";
	return `${prop.role === "" ? "Used" : prop.role} in ${where}${more}.`;
}
function toSchema(node, locations, where, description) {
	const schema = buildSchema(node, locations, where, description);
	if (where !== void 0) locations.set(schema, where);
	return schema;
}
function buildSchema(node, locations, where, description) {
	const d = description === void 0 ? {} : { description };
	switch (node.t) {
		case "value": return {
			type: VALUE_TYPES,
			...d
		};
		case "boolean": return {
			type: "boolean",
			...d
		};
		case "image": return {
			...d,
			anyOf: [{
				type: "string",
				description: "A data URL such as data:image/png;base64,…, or an http(s) address for renderAsync with fetchImage."
			}, {
				type: "object",
				properties: {
					data: {
						type: "string",
						description: "A data URL or an address, as above (in JavaScript also the bytes)."
					},
					width: {
						type: "number",
						description: "Width in centimetres."
					},
					height: {
						type: "number",
						description: "Height in centimetres."
					},
					alt: {
						type: "string",
						description: "Alternative text."
					}
				},
				required: ["data"]
			}]
		};
		case "markdown":
		case "url": return {
			type: "string",
			...d
		};
		case "array": return {
			type: "array",
			...d,
			items: node.items === void 0 ? {} : toSchema(node.items, locations, where)
		};
		case "object": {
			const props = [...node.props];
			const required = props.filter(([, p]) => p.required).map(([name]) => name);
			return {
				type: "object",
				...d,
				properties: defineAll(props.map(([name, p]) => [name, toSchema(p.node, locations, p.first, p.role === "" ? void 0 : usedIn(p))])),
				...required.length === 0 ? {} : { required }
			};
		}
		case "conflict": return {
			...d,
			anyOf: node.options.map((o) => toSchema(o, locations, where))
		};
	}
}
/** The JSON Schema of the data for the given tree (one list of children per part, in order). */
function inferSchema(parts) {
	const builder = new SchemaBuilder();
	for (const children of parts) builder.addChildren(children, {
		kind: "root",
		obj: builder.root
	});
	const locations = /* @__PURE__ */ new WeakMap();
	return {
		schema: {
			$schema: "https://json-schema.org/draft/2020-12/schema",
			...toSchema(builder.root, locations, void 0)
		},
		warnings: builder.warnings,
		locations
	};
}
// src/opc/namespaces.ts
/** OPC namespaces (ECMA-376 Part 2 schemas opc-contentTypes.xsd and opc-relationships.xsd). */
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
/**
* The first `length` UTF-16 code units of `text`, one fewer when the cut would split a surrogate
* pair, so that no lone high surrogate is left at the end.
*/
function cutAt(text, length) {
	const code = text.charCodeAt(length - 1);
	const pairStart = code >= 55296 && code <= 56319 && text.length > length;
	return text.slice(0, pairStart ? length - 1 : length);
}
/** Shortens a value from the file for messages and details, on a code point boundary. */
function clip(value) {
	return value.length > 200 ? `${cutAt(value, 200)}…` : value;
}
// src/xml/chars.ts
/** The Char production of XML 1.0 (§2.2) for a code point. */
function isXmlChar(c) {
	return c === 9 || c === 10 || c === 13 || c >= 32 && c <= 55295 || c >= 57344 && c <= 65533 || c >= 65536 && c <= 1114111;
}
// src/xml/parser.ts
/** XML 1.0 §2.11: CR LF and lone CR become LF. The parser applies this before parsing. */
function normalizeLineEnds(xml) {
	return xml.indexOf("\r") === -1 ? xml : xml.replace(/\r\n?/g, "\n");
}
/** Namespace bound to the `xml` prefix by definition (Namespaces in XML 1.0, §3). */
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
const TAB = 9;
const LF = 10;
const CR = 13;
const SPACE = 32;
const QUOTE = 34;
const HASH = 35;
const APOS = 39;
const MINUS = 45;
const DOT$1 = 46;
const SLASH$1 = 47;
const COLON$1 = 58;
const EQ = 61;
const GT = 62;
const QUESTION = 63;
const UNDERSCORE = 95;
function isWhitespace(c) {
	return c === SPACE || c === LF || c === TAB || c === CR;
}
function isNameStartChar(c) {
	if (c < 128) return c >= 97 && c <= 122 || c >= 65 && c <= 90 || c === UNDERSCORE || c === COLON$1;
	return c >= 192 && c <= 214 || c >= 216 && c <= 246 || c >= 248 && c <= 767 || c >= 880 && c <= 893 || c >= 895 && c <= 8191 || c >= 8204 && c <= 8205 || c >= 8304 && c <= 8591 || c >= 11264 && c <= 12271 || c >= 12289 && c <= 55295 || c >= 63744 && c <= 64975 || c >= 65008 && c <= 65533 || c >= 65536 && c <= 983039;
}
function isNameChar(c) {
	if (c < 128) return c >= 97 && c <= 122 || c >= 65 && c <= 90 || c >= 48 && c <= 57 || c === UNDERSCORE || c === COLON$1 || c === MINUS || c === DOT$1;
	return isNameStartChar(c) || c === 183 || c >= 768 && c <= 879 || c >= 8255 && c <= 8256;
}
const PREDEFINED_ENTITIES = /* @__PURE__ */ new Map([
	["lt", "<"],
	["gt", ">"],
	["amp", "&"],
	["apos", "'"],
	["quot", "\""]
]);
/**
* Parses an XML document held in a string and reports events to `handler`.
* Throws {@link TvError} with `TV_E_XML_MALFORMED`, `TV_E_XML_DTD_FORBIDDEN`,
* `TV_E_XML_ENTITY_FORBIDDEN` or `TV_E_LIMIT_EXCEEDED`.
*/
function parseXml(xml, handler, options) {
	const parser = new Parser(handler, options, true);
	parser.write(xml);
	parser.end();
}
const COMMENT_TOO_LONG = "Comment is longer than the limit.";
const PI_TOO_LONG = "Processing instruction is longer than the limit.";
/** Consumed input is dropped once at least this many characters can go (amortises copying). */
const COMPACT_THRESHOLD = 65536;
var Parser = class {
	handler;
	options;
	whole;
	/** Unconsumed input after end-of-line normalisation; `base` is the offset of s[0] in the whole text. */
	s = "";
	base = 0;
	pos = 0;
	stack = [];
	/** Current prefix bindings; '' is the default namespace. */
	bindings = /* @__PURE__ */ new Map([["xml", XML_NAMESPACE]]);
	textBuffer = "";
	rootSeen = false;
	ended = false;
	declarationChecked = false;
	/** A CR at the end of a piece: it may be the first half of CR LF. */
	pendingCR = false;
	/** A high surrogate at the end of a piece: its low half may follow. */
	pendingHigh = "";
	/** Newlines in dropped input and the offset where the current line starts (for locations). */
	droppedLines = 0;
	lineStartAbs = 0;
	/** Character data run that was processed in parts: its start offset and that offset's location. */
	segmentStart = -1;
	segmentLocation = null;
	/** Resumable search for the end of the token at `tokenAt`. */
	tokenAt = -1;
	tokenFrom = 0;
	tokenQuote = 0;
	/** Offset of the quote that opened the attribute value being searched. */
	tokenQuoteAt = 0;
	/**
	* Pieces held back while the unconsumed input is long (a long token or text run), joined to `s`
	* once they are at least as long as it. Appending every piece to `s` and searching it again
	* made a long token cost time quadratic in its length; joining at doubling sizes keeps the
	* total work linear (D-H01).
	*/
	held = [];
	heldLength = 0;
	constructor(handler, options, whole) {
		this.handler = handler;
		this.options = options;
		this.whole = whole;
	}
	get buffered() {
		return this.s.length - this.pos + this.heldLength + this.textBuffer.length;
	}
	write(chunk) {
		if (this.ended) throw new Error("XML stream parser: write() after end().");
		let text = this.pendingCR ? `\r${chunk}` : chunk;
		this.pendingCR = false;
		if (!this.whole && text.endsWith("\r")) {
			this.pendingCR = true;
			text = text.slice(0, -1);
		}
		text = this.pendingHigh + normalizeLineEnds(text);
		this.pendingHigh = "";
		if (!this.whole && text.length > 0) {
			const last = text.charCodeAt(text.length - 1);
			if (last >= 55296 && last <= 56319) {
				this.pendingHigh = text.slice(-1);
				text = text.slice(0, -1);
			}
		}
		if (text.length === 0) return;
		if (this.whole) {
			this.append(text);
			return;
		}
		this.held.push(text);
		this.heldLength += text.length;
		if (this.heldLength < this.s.length - this.pos) return;
		this.appendHeld();
		this.pump(false);
		this.compact();
	}
	append(text) {
		const from = this.s.length;
		this.s += text;
		this.checkCharacters(from);
	}
	appendHeld() {
		if (this.held.length === 0) return;
		const text = this.held.length === 1 ? this.held[0] ?? "" : this.held.join("");
		this.held = [];
		this.heldLength = 0;
		this.append(text);
	}
	end() {
		if (this.ended) return;
		const tail = this.pendingHigh + (this.pendingCR ? "\n" : "");
		this.pendingCR = false;
		this.pendingHigh = "";
		this.appendHeld();
		if (tail.length > 0) this.append(tail);
		this.ended = true;
		this.pump(true);
	}
	/** Drops consumed input, keeping line counts for error locations. */
	compact() {
		const drop = this.pos;
		if (drop < COMPACT_THRESHOLD && drop * 2 < this.s.length) return;
		if (drop === 0) return;
		const s = this.s;
		let last = -1;
		for (let i = s.indexOf("\n"); i !== -1 && i < drop; i = s.indexOf("\n", i + 1)) {
			this.droppedLines++;
			last = i;
		}
		if (last !== -1) this.lineStartAbs = this.base + last + 1;
		this.s = s.slice(drop);
		this.base += drop;
		this.pos = 0;
		if (this.tokenAt !== -1) {
			this.tokenAt -= drop;
			this.tokenFrom -= drop;
			this.tokenQuoteAt -= drop;
		}
	}
	/**
	* Parses as much as the buffer allows. With `final` false, a token that may continue in the
	* next piece is left for later; with `final` true, the rest of the document is parsed and the
	* end-of-document checks run, exactly as in whole-text parsing.
	*/
	pump(final) {
		const s = this.s;
		if (!this.declarationChecked) {
			if (s.length - this.pos < 6 && !final) return;
			if (s.startsWith("<?xml", this.pos) && (s.length === 5 || isWhitespace(s.charCodeAt(5)) || s.charCodeAt(5) === QUESTION)) {
				if (!final && s.indexOf("?>", this.pos + 2) === -1) {
					if (s.length - this.pos - 2 > this.options.limits.maxXmlTextLength + 1) this.limit("maxXmlTextLength", PI_TOO_LONG, this.pos);
					return;
				}
				this.declarationChecked = true;
				this.skipProcessingInstruction(true);
			}
			this.declarationChecked = true;
		}
		while (this.pos < s.length) {
			const lt = s.indexOf("<", this.pos);
			if (this.stack.length === 0) {
				const end = lt === -1 ? s.length : lt;
				for (let i = this.pos; i < end; i++) if (!isWhitespace(s.charCodeAt(i))) this.fail(this.rootSeen ? "Text after the root element." : "Text before the root element.", i);
				this.pos = end;
				if (lt === -1) break;
				if (!final && !this.tokenComplete(lt)) return;
				this.markupOutsideRoot();
				continue;
			}
			if (lt === -1) {
				if (final) this.fail(`Element <${this.stack[this.stack.length - 1]?.qname ?? ""}> is not closed.`, s.length);
				this.partialCharacterData();
				return;
			}
			if (lt > this.pos) this.characterData(this.pos, lt);
			this.pos = lt;
			if (!final && !this.tokenComplete(lt)) return;
			this.markupInContent();
		}
		if (!final) return;
		if (this.stack.length > 0) this.fail(`Element <${this.stack[this.stack.length - 1]?.qname ?? ""}> is not closed.`, s.length);
		if (!this.rootSeen) this.fail("The document has no root element.", s.length);
	}
	/**
	* Whether the markup starting at `at` ends inside the buffer, so that parsing it cannot run
	* past the end of the input received so far. The search resumes where it stopped.
	*/
	tokenComplete(at) {
		const s = this.s;
		const length = s.length;
		if (this.tokenAt !== at) {
			this.tokenAt = at;
			this.tokenFrom = at + 1;
			this.tokenQuote = 0;
		}
		if (length - at < 2) return false;
		const next = s.charCodeAt(at + 1);
		let done;
		if (next === 33) {
			if (length - at >= 4 && s.startsWith("<!--", at)) {
				const dashes = s.indexOf("--", Math.max(at + 4, this.tokenFrom));
				done = dashes !== -1 && dashes + 2 < length;
				if (!done) {
					this.tokenFrom = dashes !== -1 ? dashes : Math.max(at + 4, length - 1);
					if (length - (at + 4) > this.options.limits.maxXmlTextLength + 2) this.limit("maxXmlTextLength", COMMENT_TOO_LONG, at);
				}
			} else if (length - at < 9) done = false;
			else if (s.startsWith("<![CDATA[", at)) {
				done = s.indexOf("]]>", Math.max(at + 9, this.tokenFrom)) !== -1;
				if (!done) {
					this.tokenFrom = Math.max(at + 9, length - 2);
					if (this.textBuffer.length + length - (at + 9) > this.options.limits.maxXmlTextLength + 2) this.limit("maxXmlTextLength", "Text node is longer than the limit.", at);
				}
			} else done = true;
		} else if (next === QUESTION) {
			done = s.indexOf("?>", Math.max(at + 2, this.tokenFrom)) !== -1;
			if (!done) {
				this.tokenFrom = Math.max(at + 2, length - 1);
				if (length - (at + 2) > this.options.limits.maxXmlTextLength + 1) this.limit("maxXmlTextLength", PI_TOO_LONG, at);
			}
		} else if (next === SLASH$1) {
			done = s.indexOf(">", Math.max(at + 2, this.tokenFrom)) !== -1;
			if (!done) this.tokenFrom = length;
		} else {
			let i = Math.max(at + 1, this.tokenFrom);
			let quote = this.tokenQuote;
			done = false;
			for (; i < length; i++) {
				const c = s.charCodeAt(i);
				if (quote !== 0) {
					if (c === quote) quote = 0;
				} else if (c === QUOTE || c === APOS) {
					quote = c;
					this.tokenQuoteAt = i;
				} else if (c === GT) {
					done = true;
					break;
				}
			}
			this.tokenFrom = i;
			this.tokenQuote = quote;
			if (!done && quote !== 0 && i - (this.tokenQuoteAt + 1) > this.options.limits.maxXmlTextLength) this.limit("maxXmlTextLength", `Value of attribute "${this.attributeNameBefore(this.tokenQuoteAt)}" is longer than the limit.`, this.tokenQuoteAt + 1);
		}
		if (done) this.tokenAt = -1;
		return done;
	}
	/**
	* Character data with no "<" after it yet. A run longer than the text limit is processed in
	* parts, up to what cannot change when more input arrives (everything before an unfinished
	* reference or a trailing "]" that may start "]]>"), so memory stays bounded.
	*/
	partialCharacterData() {
		const s = this.s;
		if (s.length - this.pos <= this.options.limits.maxXmlTextLength) return;
		let cut = s.length;
		const amp = s.indexOf("&", Math.max(this.pos, s.lastIndexOf(";") + 1));
		if (amp !== -1) {
			cut = amp;
			if (s.length - amp > this.options.limits.maxXmlNameLength + 12) {
				if (s.charCodeAt(amp + 1) !== HASH) this.limit("maxXmlNameLength", "Entity name is longer than the limit.", amp);
				this.fail("\"&\" must start an entity or character reference ending with \";\".", amp);
			}
		}
		while (cut > this.pos && s.charCodeAt(cut - 1) === 93 && cut > s.length - 2) cut--;
		if (cut <= this.pos) return;
		if (this.segmentStart === -1) {
			this.segmentStart = this.base + this.pos;
			this.segmentLocation = this.location(this.pos);
		}
		this.characterData(this.pos, cut);
		this.pos = cut;
	}
	/** Every character from `from` on must match the Char production (XML 1.0 §2.2). */
	checkCharacters(from) {
		const s = this.s;
		for (let i = from; i < s.length; i++) {
			const c = s.charCodeAt(i);
			if (c >= 32 && c < 55296) continue;
			if (c === LF || c === TAB || c === CR) continue;
			if (c >= 55296 && c <= 56319) {
				const next = s.charCodeAt(i + 1);
				if (next >= 56320 && next <= 57343) {
					i++;
					continue;
				}
			} else if (c >= 57344 && c <= 65533) continue;
			this.fail(`Character U+${c.toString(16).toUpperCase().padStart(4, "0")} is not allowed in XML.`, i);
		}
	}
	markupOutsideRoot() {
		const s = this.s;
		if (s.startsWith("<!--", this.pos)) return this.skipComment();
		if (s.startsWith("<?", this.pos)) return this.skipProcessingInstruction(false);
		if (s.startsWith("<!DOCTYPE", this.pos)) return this.failDoctype();
		if (s.startsWith("<!", this.pos)) this.fail("Unexpected markup declaration outside the root element.", this.pos);
		if (s.charCodeAt(this.pos + 1) === SLASH$1) this.fail("End tag without a matching start tag.", this.pos);
		if (this.rootSeen) this.fail("A second root element is not allowed.", this.pos);
		this.rootSeen = true;
		this.startTag();
	}
	markupInContent() {
		this.segmentStart = -1;
		this.segmentLocation = null;
		const s = this.s;
		const next = s.charCodeAt(this.pos + 1);
		if (next === SLASH$1) return this.endTag();
		if (next === 33) {
			if (s.startsWith("<!--", this.pos)) return this.skipComment();
			if (s.startsWith("<![CDATA[", this.pos)) return this.cdata();
			if (s.startsWith("<!DOCTYPE", this.pos)) return this.failDoctype();
			this.fail("Unexpected markup declaration in element content.", this.pos);
		}
		if (next === QUESTION) return this.skipProcessingInstruction(false);
		this.startTag();
	}
	characterData(from, to) {
		const raw = this.s.slice(from, to);
		const close = raw.indexOf("]]>");
		if (close !== -1) this.fail("The sequence \"]]>\" is not allowed in character data.", from + close);
		const at = this.segmentStart === -1 ? this.base + from : this.segmentStart;
		this.appendText(raw.indexOf("&") === -1 ? raw : this.expandReferences(raw, from), at);
	}
	cdata() {
		const start = this.pos + 9;
		const end = this.s.indexOf("]]>", start);
		if (end === -1) this.fail("CDATA section is not terminated.", this.pos);
		this.appendText(this.s.slice(start, end), this.base + this.pos);
		this.pos = end + 3;
	}
	/** `at` is an absolute offset. */
	appendText(text, at) {
		if (text.length === 0) return;
		if (this.textBuffer.length + text.length > this.options.limits.maxXmlTextLength) {
			const message = "Text node is longer than the limit.";
			if (at === this.segmentStart && this.segmentLocation !== null && at < this.base) {
				const details = {
					limit: "maxXmlTextLength",
					value: this.options.limits.maxXmlTextLength
				};
				throw new TvError("TV_E_LIMIT_EXCEEDED", message, {
					location: this.segmentLocation,
					hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTextLength: ${String(this.options.limits.maxXmlTextLength * 4)} } }.`,
					details
				});
			}
			this.limit("maxXmlTextLength", message, at - this.base);
		}
		this.textBuffer += text;
	}
	flushText() {
		if (this.textBuffer.length > 0) {
			const text = this.textBuffer;
			this.textBuffer = "";
			this.handler.text(text);
		}
	}
	/**
	* Expands entity and character references in `raw`, which starts at offset `base`.
	*/
	expandReferences(raw, base) {
		let out = "";
		let last = 0;
		let amp = raw.indexOf("&");
		while (amp !== -1) {
			out += raw.slice(last, amp);
			const semi = raw.indexOf(";", amp + 1);
			if (semi === -1) this.fail("\"&\" must start an entity or character reference ending with \";\".", base + amp);
			const body = raw.slice(amp + 1, semi);
			if (body.charCodeAt(0) === HASH) out += this.characterReference(body, base + amp);
			else {
				const value = PREDEFINED_ENTITIES.get(body);
				if (value === void 0) {
					if (body.length > this.options.limits.maxXmlNameLength) this.limit("maxXmlNameLength", "Entity name is longer than the limit.", base + amp);
					if (!this.isName(body)) this.fail("\"&\" must start an entity or character reference.", base + amp);
					throw new TvError("TV_E_XML_ENTITY_FORBIDDEN", `Reference to entity "&${body};", which is not one of the five predefined XML entities.`, {
						location: this.location(base + amp),
						hint: "Only &lt; &gt; &amp; &apos; &quot; and character references such as &#160; are allowed, because tplvane does not process DTDs. Re-save the file in the application that made it.",
						details: { entity: body }
					});
				}
				out += value;
			}
			last = semi + 1;
			amp = raw.indexOf("&", last);
		}
		out += raw.slice(last);
		return out;
	}
	/**
	* `&#` [0-9]+ `;` or `&#x` [0-9a-fA-F]+ `;` (XML 1.0 [66], X-02): any number of digits, so
	* leading zeros are allowed; only a lowercase `x` introduces the hexadecimal form.
	*/
	characterReference(body, at) {
		const shown = body.length > 20 ? `${body.slice(0, 20)}…` : body;
		const hex = body.charCodeAt(1) === 120;
		const start = hex ? 2 : 1;
		if (body.length === start) this.fail(`Invalid character reference "&${shown};".`, at);
		let code = 0;
		for (let i = start; i < body.length; i++) {
			const c = body.charCodeAt(i);
			let d;
			if (c >= 48 && c <= 57) d = c - 48;
			else if (hex && c >= 97 && c <= 102) d = c - 87;
			else if (hex && c >= 65 && c <= 70) d = c - 55;
			else return this.fail(`Invalid character reference "&${shown};".`, at);
			if (code <= 1114111) code = code * (hex ? 16 : 10) + d;
		}
		if (!isXmlChar(code)) this.fail(`Character reference "&${shown};" refers to a character not allowed in XML.`, at);
		return String.fromCodePoint(code);
	}
	startTag() {
		const s = this.s;
		const tagStart = this.pos;
		this.flushText();
		this.pos++;
		const qname = this.readName();
		const raw = [];
		let selfClosing = false;
		for (;;) {
			const hadSpace = this.skipWhitespace();
			const c = s.charCodeAt(this.pos);
			if (c === GT) {
				this.pos++;
				break;
			}
			if (c === SLASH$1) {
				if (s.charCodeAt(this.pos + 1) !== GT) this.fail("Expected \">\" after \"/\".", this.pos);
				this.pos += 2;
				selfClosing = true;
				break;
			}
			if (Number.isNaN(c)) this.fail(`Start tag <${qname}> is not terminated.`, tagStart);
			if (!hadSpace) this.fail("Attributes must be separated by whitespace.", this.pos);
			const at = this.pos;
			const name = this.readName();
			this.skipWhitespace();
			if (s.charCodeAt(this.pos) !== EQ) this.fail(`Attribute "${name}" has no value.`, this.pos);
			this.pos++;
			this.skipWhitespace();
			const quote = s.charCodeAt(this.pos);
			if (quote !== QUOTE && quote !== APOS) this.fail(`Value of attribute "${name}" must be quoted.`, this.pos);
			const valueStart = this.pos + 1;
			const valueEnd = s.indexOf(quote === QUOTE ? "\"" : "'", valueStart);
			if (valueEnd === -1) this.fail(`Value of attribute "${name}" is not terminated.`, this.pos);
			if (valueEnd - valueStart > this.options.limits.maxXmlTextLength) this.limit("maxXmlTextLength", `Value of attribute "${name}" is longer than the limit.`, valueStart);
			let value = s.slice(valueStart, valueEnd);
			const lt = value.indexOf("<");
			if (lt !== -1) this.fail("\"<\" is not allowed in attribute values.", valueStart + lt);
			if (value.indexOf("\n") !== -1 || value.indexOf("	") !== -1) value = value.replace(/[\n\t]/g, " ");
			if (value.indexOf("&") !== -1) value = this.expandReferences(value, valueStart);
			for (const [existing] of raw) if (existing === name) this.fail(`Attribute "${name}" appears twice.`, at);
			raw.push([
				name,
				value,
				at
			]);
			if (raw.length > this.options.limits.maxXmlAttributes) this.limit("maxXmlAttributes", `Element <${qname}> has more attributes than the limit.`, at);
			this.pos = valueEnd + 1;
		}
		if (this.stack.length + 1 > this.options.limits.maxXmlDepth) this.limit("maxXmlDepth", "Elements are nested deeper than the limit.", tagStart);
		let undo = null;
		for (const [name, value, at] of raw) {
			let prefix;
			if (name === "xmlns") prefix = "";
			else if (name.startsWith("xmlns:")) {
				prefix = name.slice(6);
				this.checkNcName(prefix, at);
			} else continue;
			this.checkDeclaration(prefix, value, at);
			(undo ??= []).push([prefix, this.bindings.get(prefix)]);
			this.bindings.set(prefix, value);
		}
		const [ns, local] = this.resolve(qname, true, tagStart);
		const attrs = /* @__PURE__ */ new Map();
		for (const [name, value, at] of raw) {
			if (name === "xmlns" || name.startsWith("xmlns:")) continue;
			const [attrNs, attrLocal] = this.resolve(name, false, at);
			const key = attrNs === "" ? attrLocal : `{${attrNs}}${attrLocal}`;
			if (attrs.has(key)) this.fail(`Attribute "${name}" appears twice (same namespace and local name).`, at);
			attrs.set(key, value);
		}
		const span = {
			start: this.base + tagStart,
			end: this.base + this.pos,
			qname,
			empty: selfClosing
		};
		this.handler.start(ns, local, attrs, span);
		if (selfClosing) {
			this.handler.end(ns, local, {
				start: this.base + this.pos,
				end: this.base + this.pos,
				qname,
				empty: true
			});
			this.restore(undo);
		} else this.stack.push({
			qname,
			ns,
			local,
			undo
		});
	}
	endTag() {
		const tagStart = this.pos;
		this.pos += 2;
		const qname = this.readName();
		this.skipWhitespace();
		if (this.s.charCodeAt(this.pos) !== GT) this.fail(`End tag </${qname}> is not terminated.`, tagStart);
		this.pos++;
		const open = this.stack.pop();
		if (open === void 0 || open.qname !== qname) this.fail(open === void 0 ? `End tag </${qname}> has no start tag.` : `End tag </${qname}> does not match <${open.qname}>.`, tagStart);
		this.flushText();
		this.handler.end(open.ns, open.local, {
			start: this.base + tagStart,
			end: this.base + this.pos,
			qname,
			empty: false
		});
		this.restore(open.undo);
	}
	restore(undo) {
		if (undo === null) return;
		for (let i = undo.length - 1; i >= 0; i--) {
			const entry = undo[i];
			if (entry === void 0) continue;
			const [prefix, previous] = entry;
			if (previous === void 0) this.bindings.delete(prefix);
			else this.bindings.set(prefix, previous);
		}
	}
	/** Namespace constraints on declarations (Namespaces in XML 1.0, §3). */
	checkDeclaration(prefix, uri, at) {
		if (uri.length > this.options.limits.maxXmlNameLength) this.limit("maxXmlNameLength", "Namespace name is longer than the limit.", at);
		if (prefix === "xmlns") this.fail("The prefix \"xmlns\" must not be declared.", at);
		if (prefix === "xml" && uri !== "http://www.w3.org/XML/1998/namespace") this.fail("The prefix \"xml\" must not be bound to another namespace.", at);
		if (prefix !== "xml" && uri === "http://www.w3.org/XML/1998/namespace") this.fail("The XML namespace must not be bound to another prefix.", at);
		if (uri === XMLNS_NAMESPACE) this.fail("The xmlns namespace must not be declared.", at);
		if (prefix !== "" && uri === "") this.fail(`Prefix "${prefix}" cannot be undeclared in XML 1.0.`, at);
		if (prefix !== "") this.checkNcName(prefix, at);
	}
	resolve(qname, element, at) {
		const colon = qname.indexOf(":");
		if (colon === -1) return [element ? this.bindings.get("") ?? "" : "", qname];
		const prefix = qname.slice(0, colon);
		const local = qname.slice(colon + 1);
		this.checkNcName(prefix, at);
		this.checkNcName(local, at);
		const ns = this.bindings.get(prefix);
		if (ns === void 0) this.fail(`Namespace prefix "${prefix}" is not declared.`, at);
		return [ns, local];
	}
	checkNcName(name, at) {
		if (name.length === 0 || name.indexOf(":") !== -1) this.fail(`"${name}" is not a valid name.`, at);
	}
	readName() {
		const s = this.s;
		const start = this.pos;
		const first = s.codePointAt(start) ?? -1;
		if (!isNameStartChar(first)) this.fail("Expected a name.", start);
		let i = start + (first > 65535 ? 2 : 1);
		for (;;) {
			const c = s.charCodeAt(i);
			if (c < 128) {
				if (!isNameChar(c)) break;
				i++;
			} else {
				const cp = s.codePointAt(i) ?? -1;
				if (!isNameChar(cp)) break;
				i += cp > 65535 ? 2 : 1;
			}
			if (i - start > this.options.limits.maxXmlNameLength) this.limit("maxXmlNameLength", "Name is longer than the limit.", start);
		}
		this.pos = i;
		return s.slice(start, i);
	}
	isName(text) {
		if (text.length === 0) return false;
		const first = text.codePointAt(0) ?? -1;
		if (!isNameStartChar(first)) return false;
		for (let i = first > 65535 ? 2 : 1; i < text.length;) {
			const cp = text.codePointAt(i) ?? -1;
			if (!isNameChar(cp)) return false;
			i += cp > 65535 ? 2 : 1;
		}
		return true;
	}
	skipWhitespace() {
		const start = this.pos;
		while (isWhitespace(this.s.charCodeAt(this.pos))) this.pos++;
		return this.pos > start;
	}
	skipComment() {
		const start = this.pos;
		const end = this.s.indexOf("--", start + 4);
		if ((end === -1 ? this.s.length : end) - (start + 4) > this.options.limits.maxXmlTextLength) this.limit("maxXmlTextLength", COMMENT_TOO_LONG, start);
		if (end === -1) this.fail("Comment is not terminated.", start);
		if (this.s.charCodeAt(end + 2) !== GT) this.fail("\"--\" is not allowed inside a comment.", end);
		this.pos = end + 3;
	}
	skipProcessingInstruction(declaration) {
		const start = this.pos;
		this.pos += 2;
		const target = this.readName();
		if (!declaration && target.toLowerCase() === "xml") this.fail("The XML declaration is allowed only at the very start of the document.", start);
		const end = this.s.indexOf("?>", this.pos);
		if ((end === -1 ? this.s.length : end) - (start + 2) > this.options.limits.maxXmlTextLength) this.limit("maxXmlTextLength", PI_TOO_LONG, start);
		if (end === -1) this.fail("Processing instruction is not terminated.", start);
		const c = this.s.charCodeAt(this.pos);
		if (this.pos !== end && !isWhitespace(c)) this.fail("Invalid processing instruction target.", start);
		if (declaration) {
			let i = this.pos;
			while (isWhitespace(this.s.charCodeAt(i))) i++;
			if (i === this.pos || !this.s.startsWith("version", i)) this.fail("The XML declaration must start with the version, e.g. <?xml version=\"1.0\"?>.", start);
		}
		this.pos = end + 2;
	}
	failDoctype() {
		throw new TvError("TV_E_XML_DTD_FORBIDDEN", "The part contains a document type declaration (DTD).", {
			location: this.location(this.pos),
			hint: "tplvane rejects DTDs to protect against external entities (XXE) and entity expansion attacks. Office Open XML parts do not use DTDs; if the file was generated by a program, report the problem to its authors."
		});
	}
	/** Location of an offset into the current buffer. */
	location(offset) {
		let line = 1 + this.droppedLines;
		let lineStart = this.lineStartAbs;
		for (let i = this.s.indexOf("\n"); i !== -1 && i < offset; i = this.s.indexOf("\n", i + 1)) {
			line++;
			lineStart = this.base + i + 1;
		}
		return {
			part: this.options.part,
			line,
			column: this.base + offset - lineStart + 1
		};
	}
	fail(message, offset) {
		throw new TvError("TV_E_XML_MALFORMED", message, {
			location: this.location(offset),
			hint: "The part is not well-formed XML. Re-save the file in the application that made it; if it was generated by a program, report the problem to its authors."
		});
	}
	/** The attribute name before `name = "` whose quote is at `quote` (for a message). */
	attributeNameBefore(quote) {
		const s = this.s;
		let i = quote - 1;
		while (i >= 0 && isWhitespace(s.charCodeAt(i))) i--;
		if (s.charCodeAt(i) === EQ) i--;
		while (i >= 0 && isWhitespace(s.charCodeAt(i))) i--;
		const end = i + 1;
		while (i >= 0 && end - i <= this.options.limits.maxXmlNameLength && !isWhitespace(s.charCodeAt(i)) && s.charCodeAt(i) !== 60) i--;
		return s.slice(i + 1, end);
	}
	limit(limit, message, offset) {
		const details = {
			limit,
			value: this.options.limits[limit]
		};
		throw new TvError("TV_E_LIMIT_EXCEEDED", message, {
			location: this.location(offset),
			hint: `If you trust this file, raise the limit, for example { limits: { ${limit}: ${String(this.options.limits[limit] * 4)} } }.`,
			details
		});
	}
};
/**
* Decodes the bytes of an XML part. Office Open XML packaging parts are UTF-8 or UTF-16
* (xlvane E-55); a byte order mark selects UTF-16, otherwise UTF-8 is used. Invalid byte sequences
* and declarations of other encodings are rejected.
*/
function decodeXml(bytes, part) {
	let encoding = "utf-8";
	if (bytes[0] === 255 && bytes[1] === 254) encoding = "utf-16le";
	else if (bytes[0] === 254 && bytes[1] === 255) encoding = "utf-16be";
	let text;
	try {
		text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
	} catch (cause) {
		throw new TvError("TV_E_XML_ENCODING", `The part is not valid ${encoding.toUpperCase()}.`, {
			location: { part },
			hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in the application that made it.",
			cause
		});
	}
	const declared = declaredEncoding(text);
	if (declared !== null) {
		const normalized = declared.toLowerCase();
		if (!(encoding === "utf-8" ? normalized === "utf-8" : normalized === "utf-16" || normalized === encoding)) throw new TvError("TV_E_XML_ENCODING", `The part declares encoding "${clip(declared)}", which is not supported.`, {
			location: { part },
			hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in the application that made it.",
			details: { encoding: clip(declared) }
		});
	}
	return text;
}
/** Returns the `encoding` pseudo-attribute of the XML declaration, if present. */
function declaredEncoding(text) {
	if (!text.startsWith("<?xml")) return null;
	const end = text.indexOf("?>");
	if (end === -1) return null;
	const decl = text.slice(5, end);
	const at = decl.indexOf("encoding");
	if (at === -1) return null;
	let i = at + 8;
	while (isWhitespace(decl.charCodeAt(i))) i++;
	if (decl.charCodeAt(i) !== EQ) return null;
	i++;
	while (isWhitespace(decl.charCodeAt(i))) i++;
	const quote = decl[i];
	if (quote !== "\"" && quote !== "'") return null;
	const close = decl.indexOf(quote, i + 1);
	return close === -1 ? null : decl.slice(i + 1, close);
}
// src/xml/tree.ts
/**
* Parses a whole part into a tree. Use only for small parts (workbook, relationships, content
* types, styles), or with `skip` for parts whose bulk is not needed; worksheets and shared
* strings are processed with {@link parseXml} directly.
*/
function parseXmlTree(xml, options) {
	const stack = [];
	let root;
	let count = 0;
	let skipped = 0;
	const skip = options.skip;
	parseXml(xml, {
		start(ns, local, attrs) {
			if (skipped > 0 || skip !== void 0 && skip(ns, local)) {
				skipped++;
				return;
			}
			count += 1 + attrs.size;
			if (count > options.limits.maxXmlTreeElements) throw new TvError("TV_E_LIMIT_EXCEEDED", "The part has more elements and attributes than the limit.", {
				location: { part: options.part },
				hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTreeElements: ${String(options.limits.maxXmlTreeElements * 4)} } }.`,
				details: {
					limit: "maxXmlTreeElements",
					value: options.limits.maxXmlTreeElements
				}
			});
			const element = {
				ns,
				local,
				attrs,
				children: [],
				text: ""
			};
			const parent = stack[stack.length - 1];
			if (parent === void 0) root = element;
			else parent.children.push(element);
			stack.push(element);
		},
		end() {
			if (skipped > 0) skipped--;
			else stack.pop();
		},
		text(value) {
			if (skipped > 0) return;
			const current = stack[stack.length - 1];
			if (current !== void 0) current.text += value;
		}
	}, options);
	if (root === void 0) return {
		ns: "",
		local: "",
		attrs: /* @__PURE__ */ new Map(),
		children: [],
		text: ""
	};
	return root;
}
// src/opc/part-name.ts
const SLASH = 47;
const BACKSLASH = 92;
const PERCENT = 37;
const DOT = 46;
const COLON = 58;
/** RFC 3986 §2.3 `unreserved`, ASCII part. */
function isUnreservedAscii(c) {
	return c >= 97 && c <= 122 || c >= 65 && c <= 90 || c >= 48 && c <= 57 || c === 45 || c === DOT || c === 95 || c === 126;
}
/** RFC 3986 §2.2 `sub-delims`. */
function isSubDelim(c) {
	return c === 33 || c === 36 || c >= 38 && c <= 44 || c === 59 || c === 61;
}
/** RFC 3987 §2.2 `ucschar` (F-OPC-01). */
function isUcschar(c) {
	if (c >= 160 && c <= 55295) return true;
	if (c >= 63744 && c <= 64975) return true;
	if (c >= 65008 && c <= 65519) return true;
	if (c >= 65536 && c <= 917503) return (c & 65535) <= 65533;
	return c >= 921600 && c <= 983037;
}
/** RFC 3987 §2.2 `iunreserved`. */
function isIunreserved(c) {
	return c < 128 ? isUnreservedAscii(c) : isUcschar(c);
}
function hexValue(c) {
	if (c >= 48 && c <= 57) return c - 48;
	if (c >= 65 && c <= 70) return c - 55;
	if (c >= 97 && c <= 102) return c - 87;
	return -1;
}
/** The octet encoded by `%HH` at `at`, or -1 if `at` does not start a valid percent-encoding. */
function octetAt(s, at) {
	if (s.charCodeAt(at) !== PERCENT) return -1;
	const hi = hexValue(s.charCodeAt(at + 1));
	const lo = hexValue(s.charCodeAt(at + 2));
	return hi < 0 || lo < 0 ? -1 : hi * 16 + lo;
}
const utf8 = new TextDecoder("utf-8", { ignoreBOM: true });
/**
* Reads a percent-encoded character at `at`: one octet for ASCII, or the UTF-8 sequence of a
* non-ASCII character (RFC 3987 §3.1 converts characters to UTF-8 octets before percent-encoding
* them; F-OPC-15). Returns
* `null` when `at` is a "%" that is not followed by two hexadecimal digits.
*/
function percentAt(s, at) {
	const lead = octetAt(s, at);
	if (lead < 0) return null;
	if (lead < 128) return {
		cp: lead,
		end: at + 3
	};
	const length = lead >= 240 ? 4 : lead >= 224 ? 3 : lead >= 192 ? 2 : 0;
	if (length === 0) return {
		cp: -1,
		end: at + 3
	};
	const bytes = new Uint8Array(length);
	bytes[0] = lead;
	for (let i = 1; i < length; i++) {
		const octet = octetAt(s, at + 3 * i);
		if (octet < 0) return {
			cp: -1,
			end: at + 3
		};
		bytes[i] = octet;
	}
	const text = utf8.decode(bytes);
	const cp = text.codePointAt(0) ?? -1;
	const single = text.length === (cp > 65535 ? 2 : 1);
	const replacement = cp === 65533 && !(bytes[0] === 239 && bytes[1] === 191 && bytes[2] === 189);
	return single && !replacement && cp >= 128 ? {
		cp,
		end: at + 3 * length
	} : {
		cp: -1,
		end: at + 3
	};
}
/** Replaces the percent-encoded characters for which `decode` returns true by the characters. */
function decodePercent(s, decode) {
	if (!s.includes("%")) return s;
	let out = "";
	let from = 0;
	for (let i = 0; i < s.length;) {
		if (s.charCodeAt(i) !== PERCENT) {
			i++;
			continue;
		}
		const p = percentAt(s, i);
		if (p === null) {
			i++;
			continue;
		}
		if (p.cp >= 0 && decode(p.cp)) {
			out += s.slice(from, i) + String.fromCodePoint(p.cp);
			from = p.end;
		}
		i = p.end;
	}
	return out + s.slice(from);
}
/**
* Folds A–Z to a–z and nothing else: the ASCII case-insensitive matching that Part 2 uses for
* part names, logical item names, extensions and Override part names (§6.2.2.3, §7.2.3.5;
* F-OPC-03). Two names are equivalent exactly when their keys are equal.
*/
function partKey(name) {
	return name.replace(/[A-Z]+/g, (run) => run.toLowerCase());
}
/**
* Returns `true` if `name` is a valid part name: `1*( "/" isegment-nz )`, where no segment holds
* a percent-encoded "/", "\" or `iunreserved` character, and no segment ends with "."
* (Part 2 §6.2.2.2; F-OPC-01).
*/
function isPartName(name) {
	if (name.length < 2 || name.charCodeAt(0) !== SLASH) return false;
	let segmentLength = 0;
	let lastWasDot = false;
	for (let i = 1; i < name.length;) {
		const c = name.codePointAt(i);
		if (c === SLASH) {
			if (segmentLength === 0 || lastWasDot) return false;
			segmentLength = 0;
			i++;
			continue;
		}
		segmentLength++;
		lastWasDot = c === DOT;
		if (c === PERCENT) {
			const p = percentAt(name, i);
			if (p === null || p.cp === SLASH || p.cp === BACKSLASH || p.cp >= 0 && isIunreserved(p.cp)) return false;
			i = p.end;
			continue;
		}
		if (!(c < 128 ? isUnreservedAscii(c) || isSubDelim(c) || c === COLON || c === 64 : isUcschar(c))) return false;
		i += c > 65535 ? 2 : 1;
	}
	return segmentLength > 0 && !lastWasDot;
}
/**
* Maps a ZIP item name to a logical item name: un-percent-encode every non-ASCII character, then
* add a leading "/" (Part 2 §7.3.5; F-OPC-12). The result is a part only if {@link isPartName}
* accepts it. Names stored with raw non-ASCII characters (UTF-8 flag) give the same result as
* their percent-encoded spelling (D-OPC-01).
*/
function zipNameToPartName(zipName) {
	return `/${decodePercent(zipName, (cp) => cp >= 128)}`;
}
/**
* Maps a part name to the ZIP item name to write: drop the leading "/" and percent-encode the
* UTF-8 bytes of every non-ASCII character, in upper-case hex. The inverse of
* {@link zipNameToPartName}: `zipNameToPartName(partNameToZipName(p)) === p` for every part name
* `p`, and the result is printable ASCII, which `writeZip` requires (xlvane D-008 gap 3, option A).
* Returns null for a name with an unpaired surrogate, which no part name has.
*/
function partNameToZipName(partName) {
	const name = partName.charCodeAt(0) === SLASH ? partName.slice(1) : partName;
	let out = "";
	for (let i = 0; i < name.length; i++) {
		const c = name.charCodeAt(i);
		if (c < 128) {
			out += name[i];
			continue;
		}
		let cp = c;
		if (c >= 55296 && c <= 56319) {
			const next = name.charCodeAt(i + 1);
			if (!(next >= 56320 && next <= 57343)) return null;
			cp = (c - 55296 << 10) + (next - 56320) + 65536;
			i++;
		} else if (c >= 56320 && c <= 57343) return null;
		for (const byte of utf8Bytes(cp)) out += `%${byte.toString(16).toUpperCase().padStart(2, "0")}`;
	}
	return out;
}
/** The UTF-8 bytes of a code point (RFC 3629 §3). */
function utf8Bytes(cp) {
	if (cp < 128) return [cp];
	if (cp < 2048) return [192 | cp >> 6, 128 | cp & 63];
	if (cp < 65536) return [
		224 | cp >> 12,
		128 | cp >> 6 & 63,
		128 | cp & 63
	];
	return [
		240 | cp >> 18,
		128 | cp >> 12 & 63,
		128 | cp >> 6 & 63,
		128 | cp & 63
	];
}
/**
* Decodes percent-encoded `iunreserved` characters, which a valid part name cannot contain
* (F-OPC-01 b). Annex A (informative) suggests the same step; see D-OPC-02.
*/
function canonicalReference(reference) {
	return decodePercent(reference, isIunreserved);
}
/**
* Returns `true` for Relationships part names: `/_rels/.rels`, or a name whose second-to-last
* segment is equivalent to `_rels` and whose last segment ends with `.rels` (Part 2 §6.2.2.2;
* F-OPC-02).
*/
function isRelationshipsPart(name) {
	const key = partKey(name);
	const last = key.lastIndexOf("/");
	if (last < 1 || !key.endsWith(".rels")) return false;
	const previous = key.lastIndexOf("/", last - 1);
	return previous >= 0 && key.slice(previous + 1, last) === "_rels";
}
/**
* RFC 3986 §5.2.4 remove_dot_segments for a path that starts with "/". Works on a segment stack,
* so it is linear in the length of the path; "/../" at the root collapses to "/" (F-OPC-06).
*/
function removeDotSegments(path) {
	const segments = path.split("/");
	const out = [];
	const lastIndex = segments.length - 1;
	for (let i = 1; i <= lastIndex; i++) {
		const segment = segments[i];
		if (segment === "." || segment === "..") {
			if (segment === "..") out.pop();
			if (i === lastIndex) out.push("");
			continue;
		}
		out.push(segment);
	}
	return `/${out.join("/")}`;
}
/** Returns `true` if the first segment of a path that does not start with "/" contains ":". */
function hasColonInFirstSegment(path) {
	if (path.charCodeAt(0) === SLASH) return false;
	const slash = path.indexOf("/");
	const colon = path.indexOf(":");
	return colon !== -1 && (slash === -1 || colon < slash);
}
// src/opc/package.ts
/** ZIP item name of the Media Types stream (Part 2 §7.3.7, F-OPC-09). */
const CONTENT_TYPES_ITEM = "[content_types].xml";
/** Used in error locations; the stream is not a part, but this is where users find it. */
const CONTENT_TYPES_LOCATION = "/[Content_Types].xml";
/** Media type of Relationships parts (Part 2 §6.5.2.1, F-OPC-07). */
const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const PACKAGE_RELATIONSHIPS = "/_rels/.rels";
const INVALID_HINT = "The package structure is damaged or was written by a program that does not follow ECMA-376 Part 2. Re-save the file in the application that made it; if a program generated it, report the problem to its authors.";
const EMPTY = Object.freeze([]);
function invalid(part, message, reason, extra = {}) {
	return new TvError("TV_E_PART_INVALID", message, {
		location: { part },
		hint: INVALID_HINT,
		details: {
			reason,
			...extra
		}
	});
}
function requireAttribute(element, name, part) {
	const value = element.attrs.get(name);
	if (value === void 0) throw invalid(part, `A ${element.local} element has no ${name} attribute.`, "missing-attribute", {
		element: element.local,
		attribute: name
	});
	return value;
}
/**
* Returns the name of the Relationships part for `source`: "/_rels/.rels" for the package (null),
* otherwise "_rels/" inserted before the last segment and ".rels" appended
* (Part 2 §6.5.2.2 and §6.5.2.3, F-OPC-07). "/word/document.xml" gives "/word/_rels/document.xml.rels".
*/
function relationshipsPartName(source) {
	if (source === null) return PACKAGE_RELATIONSHIPS;
	const slash = source.lastIndexOf("/");
	return `${source.slice(0, slash + 1)}_rels/${source.slice(slash + 1)}.rels`;
}
/**
* Resolves the Target of an internal relationship to a part name. The base is the source part, or
* the package root when `source` is null (Part 2 §6.4, §6.5.2.2, §6.5.2.3; F-OPC-06, F-OPC-07).
*
* Query and fragment are dropped (§6.3.2, F-OPC-05); percent-encoded `iunreserved` characters are
* decoded (D-OPC-02); the path is merged and its dot segments removed as in RFC 3986 §5.2, so
* "../" above the root stays at the root (D-OPC-03). Returns null when `target` is not a relative
* reference (it has a scheme or an authority) or the result is not a valid part name.
*/
function resolveTarget(source, target) {
	let end = target.indexOf("#");
	if (end === -1) end = target.length;
	const query = target.indexOf("?");
	if (query !== -1 && query < end) end = query;
	const raw = target.slice(0, end);
	if (hasColonInFirstSegment(raw) || raw.startsWith("//")) return null;
	const path = canonicalReference(raw);
	const base = source ?? "/";
	let merged;
	if (path === "") merged = base;
	else if (path.charCodeAt(0) === 47) merged = path;
	else merged = base.slice(0, base.lastIndexOf("/") + 1) + path;
	const resolved = removeDotSegments(merged);
	return isPartName(resolved) ? resolved : null;
}
/**
* Opens the OPC package stored in a ZIP archive. Maps ZIP items to parts (Part 2 §7.3.5,
* F-OPC-12, F-OPC-13) and reads the Media Types stream, which a ZIP package must contain
* (§7.2.3.1, F-OPC-09). Relationships parts are read on demand.
*
* Throws `TV_E_PART_MISSING` without `[Content_Types].xml`, `TV_E_PART_INVALID` for two items
* with equivalent part names or a malformed Media Types stream, and whatever the ZIP and XML
* layers throw.
*/
function openPackage(zip, limits) {
	const parts = /* @__PURE__ */ new Map();
	const partNames = [];
	let contentTypes;
	let interleaved = false;
	for (const entry of zip.entries) {
		const folded = partKey(entry.name);
		if (folded === CONTENT_TYPES_ITEM) {
			contentTypes = entry;
			continue;
		}
		if (folded.startsWith(`${CONTENT_TYPES_ITEM}/`)) {
			interleaved = true;
			continue;
		}
		if (entry.name.endsWith("/")) continue;
		const name = zipNameToPartName(entry.name);
		if (!isPartName(name)) continue;
		const key = partKey(name);
		const previous = parts.get(key);
		if (previous !== void 0) throw invalid(previous.name, `ZIP items "${clip(previous.entry.name)}" and "${clip(entry.name)}" both map to part ${clip(previous.name)}.`, "equivalent-part-names", { items: [clip(previous.entry.name), clip(entry.name)] });
		parts.set(key, {
			name,
			entry
		});
		partNames.push(name);
	}
	if (contentTypes === void 0) throw new TvError("TV_E_PART_MISSING", interleaved ? "The package stores [Content_Types].xml in interleaved pieces, which Office Open XML documents must not use." : "The package has no [Content_Types].xml, so it is not an Office Open XML package.", {
		location: { part: CONTENT_TYPES_LOCATION },
		hint: "Check that the file is a .docx, .pptx or .xlsx document and not another ZIP-based format (such as .odt or .zip). If it is, re-save it in the application that made it.",
		details: { reason: interleaved ? "interleaved" : "missing" }
	});
	const mediaTypes = readMediaTypes(zip.read(contentTypes), limits);
	return new Package(zip, limits, parts, Object.freeze(partNames), mediaTypes);
}
/** Parses the Media Types stream (Part 2 §7.2.3.2, F-OPC-10). */
function readMediaTypes(bytes, limits) {
	const part = CONTENT_TYPES_LOCATION;
	const root = parseXmlTree(decodeXml(bytes, part), {
		part,
		limits
	});
	if (root.ns !== "http://schemas.openxmlformats.org/package/2006/content-types" || root.local !== "Types") throw invalid(part, `The root element of [Content_Types].xml is not Types in namespace ${CONTENT_TYPES_NS}.`, "root-element", {
		element: clip(root.local),
		namespace: clip(root.ns)
	});
	const defaults = /* @__PURE__ */ new Map();
	const overrides = /* @__PURE__ */ new Map();
	for (const child of root.children) {
		if (child.ns !== "http://schemas.openxmlformats.org/package/2006/content-types") continue;
		if (child.local === "Default") {
			const extension = requireAttribute(child, "Extension", part);
			const type = requireAttribute(child, "ContentType", part);
			const key = partKey(extension);
			if (defaults.has(key)) throw invalid(part, `[Content_Types].xml has more than one Default for extension "${clip(extension)}".`, "duplicate-default", { extension: clip(extension) });
			defaults.set(key, type);
		} else if (child.local === "Override") {
			const partName = requireAttribute(child, "PartName", part);
			const type = requireAttribute(child, "ContentType", part);
			const key = partKey(canonicalReference(partName));
			if (overrides.has(key)) throw invalid(part, `[Content_Types].xml has more than one Override for part "${clip(partName)}".`, "duplicate-override", { partName: clip(partName) });
			overrides.set(key, type);
		}
	}
	return {
		defaults,
		overrides
	};
}
/** The extension of a part name: what follows the last "." of the last segment (§7.2.3.5 c). */
function extensionOf(partName) {
	const dot = partName.lastIndexOf(".");
	return dot > partName.lastIndexOf("/") ? partName.slice(dot + 1) : null;
}
var Package = class {
	zip;
	limits;
	parts;
	partNames;
	mediaTypes;
	cache = /* @__PURE__ */ new Map();
	constructor(zip, limits, parts, partNames, mediaTypes) {
		this.zip = zip;
		this.limits = limits;
		this.parts = parts;
		this.partNames = partNames;
		this.mediaTypes = mediaTypes;
	}
	hasPart(partName) {
		return this.parts.has(partKey(partName));
	}
	get mediaTypeDefaults() {
		return this.mediaTypes.defaults;
	}
	overrideType(partName) {
		return this.mediaTypes.overrides.get(partKey(partName));
	}
	contentType(partName) {
		const override = this.mediaTypes.overrides.get(partKey(partName));
		if (override !== void 0) return override;
		const extension = extensionOf(partName);
		if (extension !== null) {
			const type = this.mediaTypes.defaults.get(partKey(extension));
			if (type !== void 0) return type;
		}
		return isRelationshipsPart(partName) ? RELATIONSHIPS_CONTENT_TYPE : void 0;
	}
	relationships(source) {
		const key = partKey(relationshipsPartName(source));
		const cached = this.cache.get(key);
		if (cached !== void 0) return cached;
		const stored = this.parts.get(key);
		const list = stored === void 0 ? EMPTY : this.parseRelationships(stored, source);
		this.cache.set(key, list);
		return list;
	}
	readXml(partName) {
		const stored = this.parts.get(partKey(partName));
		return stored === void 0 ? void 0 : decodeXml(this.zip.read(stored.entry), stored.name);
	}
	readBytes(partName) {
		const stored = this.parts.get(partKey(partName));
		return stored === void 0 ? void 0 : this.zip.read(stored.entry);
	}
	/** Parses a Relationships part (Part 2 §6.5.3, F-OPC-08). */
	parseRelationships(stored, source) {
		const part = stored.name;
		const root = parseXmlTree(decodeXml(this.zip.read(stored.entry), part), {
			part,
			limits: this.limits
		});
		if (root.ns !== "http://schemas.openxmlformats.org/package/2006/relationships" || root.local !== "Relationships") throw invalid(part, `The root element of ${clip(part)} is not Relationships in namespace ${PACKAGE_REL_NS}.`, "root-element", {
			element: clip(root.local),
			namespace: clip(root.ns)
		});
		const ids = /* @__PURE__ */ new Set();
		const list = [];
		for (const child of root.children) {
			if (child.ns !== "http://schemas.openxmlformats.org/package/2006/relationships" || child.local !== "Relationship") continue;
			const id = requireAttribute(child, "Id", part);
			const type = requireAttribute(child, "Type", part);
			const target = requireAttribute(child, "Target", part);
			const mode = child.attrs.get("TargetMode");
			if (ids.has(id)) throw invalid(part, `Relationship Id "${clip(id)}" is used more than once.`, "duplicate-id", { id: clip(id) });
			ids.add(id);
			if (mode !== void 0 && mode !== "Internal" && mode !== "External") throw invalid(part, `Relationship "${clip(id)}" has TargetMode "${clip(mode)}"; expected Internal or External.`, "target-mode", {
				id: clip(id),
				targetMode: clip(mode)
			});
			if (mode === "External") {
				list.push(Object.freeze({
					id,
					type,
					target,
					external: true
				}));
				continue;
			}
			const resolved = resolveTarget(source, target);
			if (resolved === null) throw invalid(part, `Relationship "${clip(id)}" has the internal target "${clip(target)}", which is not a relative reference to a part.`, "internal-target", {
				id: clip(id),
				target: clip(target)
			});
			list.push(Object.freeze({
				id,
				type,
				target: resolved,
				external: false
			}));
		}
		return Object.freeze(list);
	}
};
// src/zip/cp437.ts
const HIGH = "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\xA0";
/** Decodes bytes in IBM Code Page 437. */
function decodeCp437(bytes) {
	let out = "";
	for (let i = 0; i < bytes.length; i++) {
		const b = bytes[i];
		out += b < 128 ? String.fromCharCode(b) : HIGH[b - 128];
	}
	return out;
}
// src/zip/crc32.ts
const TABLE = /* @__PURE__ */ (() => {
	const table = /* @__PURE__ */ new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 3988292384 ^ c >>> 1 : c >>> 1;
		table[n] = c >>> 0;
	}
	return table;
})();
/** Computes the CRC-32 of `data`, continuing from `crc` (the result of a previous call). */
function crc32(data, crc = 0) {
	let c = (crc ^ 4294967295) >>> 0;
	for (let i = 0; i < data.length; i++) c = TABLE[(c ^ data[i]) & 255] ^ c >>> 8;
	return (c ^ 4294967295) >>> 0;
}
// src/zip/reader.ts
const SIG_LOCAL$1 = 67324752;
const SIG_CENTRAL$1 = 33639248;
const SIG_EOCD$1 = 101010256;
const SIG_ZIP64_EOCD = 101075792;
const SIG_ZIP64_LOCATOR = 117853008;
const EOCD_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
const CENTRAL_SIZE = 46;
const MAX_COMMENT = 65535;
const FLAG_UTF8 = 2048;
/** Compressed bytes handed to the decompressor at a time; bounds the overshoot past a limit. */
const INFLATE_CHUNK = 8192;
const CFB_SIGNATURE = [
	208,
	207,
	17,
	224,
	161,
	177,
	26,
	225
];
const CORRUPT_HINT = "The file is damaged or is not a valid Office document. Re-save it in the application that made it or ask the sender for a new copy.";
function corrupt(message, reason, offset, extra = {}) {
	return new TvError("TV_E_ZIP_CORRUPT", message, {
		location: offset === void 0 ? {} : { offset },
		hint: CORRUPT_HINT,
		details: {
			reason,
			...extra
		}
	});
}
function unsupported$1(message, reason, hint, name) {
	return new TvError("TV_E_ZIP_UNSUPPORTED", message, {
		location: name === void 0 ? {} : { part: `/${name}` },
		hint,
		details: { reason }
	});
}
function limitExceeded(limit, value, message, name) {
	return new TvError("TV_E_LIMIT_EXCEEDED", message, {
		location: name === void 0 ? {} : { part: `/${name}` },
		hint: `If you trust this file, raise the limit, for example { limits: { ${limit}: ${String(value * 4)} } }.`,
		details: {
			limit,
			value
		}
	});
}
/** Thrown by {@link Bytes} when a byte outside the loaded segments is read (streaming sources only). */
var NotLoaded = class extends Error {
	start;
	end;
	constructor(start, end) {
		super(`Bytes ${String(start)}…${String(end)} are not loaded.`);
		this.start = start;
		this.end = end;
	}
};
const segmentEnd = (segment) => segment.start + segment.data.length;
/**
* Little-endian access to a file of `size` bytes, of which only some segments may be in memory
* (all of it for `openZip`; the tail, the central directory and local headers for streaming
* sources). Offsets are absolute.
*/
var Bytes = class Bytes {
	size;
	/** Loaded segments in order of `start` (equal starts in order of addition). */
	segments = [];
	/**
	* For each index i, the segment with the largest end among segments[0..i]. A range is loaded
	* exactly when that segment, for the last i whose start is at or before the range's start,
	* reaches its end; so a lookup is one binary search (D-H01).
	*/
	reach = [];
	last;
	constructor(size, segments) {
		this.size = size;
		for (const segment of segments) this.add(segment.start, segment.data);
	}
	static of(data) {
		return new Bytes(data.length, [{
			start: 0,
			data
		}]);
	}
	add(start, data) {
		const segment = {
			start,
			data
		};
		const at = this.upperBound(start);
		this.segments.splice(at, 0, segment);
		this.reach.splice(at, 0, segment);
		for (let i = at; i < this.segments.length; i++) {
			const own = this.segments[i];
			const before = i === 0 ? void 0 : this.reach[i - 1];
			this.reach[i] = before !== void 0 && segmentEnd(before) >= segmentEnd(own) ? before : own;
		}
		this.last = segment;
	}
	/** Whether `[start, end)` is fully inside one loaded segment. */
	has(start, end) {
		return this.find(start, end) !== void 0;
	}
	/** The number of segments whose start is at or before `start`. */
	upperBound(start) {
		let lo = 0;
		let hi = this.segments.length;
		while (lo < hi) {
			const mid = lo + hi >> 1;
			if (this.segments[mid].start <= start) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	}
	find(start, stop) {
		const last = this.last;
		if (last && start >= last.start && stop <= segmentEnd(last)) return last;
		const i = this.upperBound(start) - 1;
		if (i < 0) return void 0;
		const seg = this.reach[i];
		if (stop > segmentEnd(seg)) return void 0;
		this.last = seg;
		return seg;
	}
	/** The bytes `[start, end)`; throws {@link NotLoaded} if they are not in memory. */
	slice(start, end) {
		if (end <= start) return /* @__PURE__ */ new Uint8Array(0);
		const seg = this.find(start, end);
		if (seg === void 0) throw new NotLoaded(start, end);
		return seg.data.subarray(start - seg.start, end - seg.start);
	}
	byte(at) {
		if (at < 0 || at >= this.size) return 0;
		const seg = this.find(at, at + 1);
		if (seg === void 0) throw new NotLoaded(at, at + 1);
		return seg.data[at - seg.start];
	}
	u16(at) {
		return this.byte(at) | this.byte(at + 1) << 8;
	}
	u32(at) {
		return (this.u16(at) | this.u16(at + 2) << 16) >>> 0;
	}
	/** Reads a 64-bit value; rejects values above Number.MAX_SAFE_INTEGER. */
	u64(at) {
		const value = this.u32(at) + this.u32(at + 4) * 4294967296;
		if (!Number.isSafeInteger(value)) throw corrupt("A ZIP64 field holds an impossibly large value.", "zip64-value", at);
		return value;
	}
};
/**
* Opens a ZIP archive held in memory and validates its structure. Throws
* `TV_E_NOT_ZIP`, `TV_E_UNSUPPORTED_CONTAINER`, `TV_E_ZIP_CORRUPT`,
* `TV_E_ZIP_UNSUPPORTED` or `TV_E_LIMIT_EXCEEDED`.
*/
function openZip(data, limits) {
	checkSource(data.length, data.subarray(0, CFB_SIGNATURE.length), limits);
	const b = Bytes.of(data);
	const dir = locateCentralDirectory(b, limits);
	const entries = readCentralDirectory(b, dir.centralOffset, dir.centralLimit, dir.entryCount);
	validateLocalHeaders(b, entries, dir.centralOffset);
	return new Reader(b, entries, limits);
}
CFB_SIGNATURE.length;
/** Input size limit and CFB detection, from the size and the first bytes of the input. */
function checkSource(size, head, limits) {
	if (size > limits.maxInputBytes) throw limitExceeded("maxInputBytes", limits.maxInputBytes, `The input is ${size} bytes, more than the limit.`);
	if (CFB_SIGNATURE.every((b, i) => head[i] === b)) throw new TvError("TV_E_UNSUPPORTED_CONTAINER", "The file is a Compound File Binary container, not a ZIP-based Office document.", { hint: "This is either a legacy binary file (.doc, .ppt, .xls) or a password-protected (encrypted) document. Open it in its application and save it without a password in the current format (.docx, .pptx, .xlsx)." });
}
/**
* Finds and checks the end of central directory record (and the ZIP64 record and locator).
* Needs the tail of the file loaded (and the ZIP64 record, which may lie before the tail).
*/
function locateCentralDirectory(b, limits) {
	const eocd = findEndOfCentralDirectory(b);
	let thisDisk = b.u16(eocd + 4);
	let centralDisk = b.u16(eocd + 6);
	let entriesOnDisk = b.u16(eocd + 8);
	let entryCount = b.u16(eocd + 10);
	let centralSize = b.u32(eocd + 12);
	let centralOffset = b.u32(eocd + 16);
	/** Where the structures that follow the central directory begin. */
	let centralLimit = eocd;
	if (thisDisk === 65535 || centralDisk === 65535 || entriesOnDisk === 65535 || entryCount === 65535 || centralSize === 4294967295 || centralOffset === 4294967295) {
		const locator = eocd - ZIP64_LOCATOR_SIZE;
		if (locator < 0 || b.u32(locator) !== SIG_ZIP64_LOCATOR) throw corrupt("The ZIP64 end of central directory locator is missing.", "zip64-locator", eocd);
		if (b.u32(locator + 4) !== 0 || b.u32(locator + 16) !== 1) throw unsupported$1("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the document as a single file.");
		const record = b.u64(locator + 8);
		if (record + 56 > locator || b.u32(record) !== SIG_ZIP64_EOCD) throw corrupt("The ZIP64 end of central directory record is missing.", "zip64-record", record);
		thisDisk = b.u32(record + 16);
		centralDisk = b.u32(record + 20);
		entriesOnDisk = b.u64(record + 24);
		entryCount = b.u64(record + 32);
		centralSize = b.u64(record + 40);
		centralOffset = b.u64(record + 48);
		centralLimit = record;
	}
	if (thisDisk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount) throw unsupported$1("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the document as a single file.");
	if (entryCount > limits.maxEntries) throw limitExceeded("maxEntries", limits.maxEntries, `The archive has ${entryCount} entries, more than the limit.`);
	if (centralOffset + centralSize !== centralLimit) throw corrupt("The central directory is not where the end of central directory record says.", "central-directory-bounds", centralOffset);
	return {
		centralOffset,
		centralSize,
		centralLimit,
		entryCount
	};
}
/**
* Finds the end of central directory record by scanning backwards. The record must end the
* file exactly (its comment runs to the last byte); data after it is rejected. APPNOTE does not
* define a search algorithm (Z-03), so this reader picks the strictest one.
*/
function findEndOfCentralDirectory(b) {
	const size = b.size;
	const lowest = Math.max(0, size - EOCD_SIZE - MAX_COMMENT);
	const data = b.slice(lowest, size);
	for (let at = size - EOCD_SIZE; at >= lowest; at--) {
		const i = at - lowest;
		if (data[i] === 80 && data[i + 1] === 75 && b.u32(at) === SIG_EOCD$1) {
			if (at + EOCD_SIZE + b.u16(at + 20) === size) return at;
		}
	}
	if (size >= 4 && b.u32(0) === SIG_LOCAL$1) throw corrupt("The end of the ZIP archive is missing; the file is probably truncated.", "eocd-missing");
	if (size >= EOCD_SIZE && data.indexOf(80) !== -1 && containsSignature(data, 0)) throw corrupt("The end of central directory record is followed by unexpected data.", "trailing-data");
	throw new TvError("TV_E_NOT_ZIP", "The input is not a ZIP archive, so it cannot be an Office document.", { hint: "DOCX, PPTX and XLSX files are ZIP archives. Check that the file is a .docx, .pptx or .xlsx document and not a legacy, PDF or text file." });
}
function containsSignature(data, from) {
	for (let at = data.length - 4; at >= from; at--) if (data[at] === 80 && data[at + 1] === 75 && data[at + 2] === 5 && data[at + 3] === 6) return true;
	return false;
}
function readCentralDirectory(b, start, end, count) {
	const entries = [];
	const seen = /* @__PURE__ */ new Map();
	let p = start;
	for (let i = 0; i < count; i++) {
		if (p + CENTRAL_SIZE > end || b.u32(p) !== SIG_CENTRAL$1) throw corrupt("A central directory entry is missing or damaged.", "central-directory-entry", p);
		const flags = b.u16(p + 8);
		const method = b.u16(p + 10);
		const crc = b.u32(p + 16);
		let compressedSize = b.u32(p + 20);
		let uncompressedSize = b.u32(p + 24);
		const nameLength = b.u16(p + 28);
		const extraLength = b.u16(p + 30);
		const commentLength = b.u16(p + 32);
		let diskStart = b.u16(p + 34);
		let headerOffset = b.u32(p + 42);
		const next = p + CENTRAL_SIZE + nameLength + extraLength + commentLength;
		if (next > end) throw corrupt("A central directory entry runs past the end of the directory.", "central-directory-entry", p);
		const nameBytes = b.slice(p + CENTRAL_SIZE, p + CENTRAL_SIZE + nameLength);
		const name = decodeName(nameBytes, flags, p);
		if (flags & 8257 || method === 99) throw unsupported$1(`ZIP entry "${name}" is encrypted.`, "encrypted", "Encrypted ZIP entries are not supported. Save the document again without ZIP encryption.", name);
		if (uncompressedSize === 4294967295 || compressedSize === 4294967295 || headerOffset === 4294967295 || diskStart === 65535) {
			const zip64 = findExtraField(b, p + CENTRAL_SIZE + nameLength, extraLength, 1);
			if (zip64 === null) throw corrupt(`ZIP entry "${name}" lacks its ZIP64 extra field.`, "zip64-extra", p);
			let q = zip64.start;
			const take = (size) => {
				if (q + size > zip64.end) throw corrupt(`The ZIP64 extra field of "${name}" is too short.`, "zip64-extra", p);
				const value = size === 8 ? b.u64(q) : b.u32(q);
				q += size;
				return value;
			};
			if (uncompressedSize === 4294967295) uncompressedSize = take(8);
			if (compressedSize === 4294967295) compressedSize = take(8);
			if (headerOffset === 4294967295) headerOffset = take(8);
			if (diskStart === 65535) diskStart = take(4);
		} else findExtraField(b, p + CENTRAL_SIZE + nameLength, extraLength, -1);
		if (diskStart !== 0) throw unsupported$1("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the document as a single file.");
		const key = name.toLowerCase();
		const previous = seen.get(key);
		if (previous !== void 0) throw corrupt(`The archive contains two entries named "${previous}" and "${name}".`, "duplicate-name", p, { name });
		seen.set(key, name);
		entries.push({
			name,
			nameBytes,
			method,
			flags,
			crc32: crc,
			compressedSize,
			uncompressedSize,
			headerOffset,
			dataOffset: 0
		});
		p = next;
	}
	if (p !== end) throw corrupt("The central directory size does not match its entries.", "central-directory-size", p);
	return entries;
}
function decodeName(bytes, flags, at) {
	if (flags & FLAG_UTF8) try {
		return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
	} catch {
		throw corrupt("A ZIP entry name is not valid UTF-8.", "name-encoding", at);
	}
	return decodeCp437(bytes);
}
/** Walks the extra field blocks (APPNOTE 4.5.1) and returns the data range of block `id`. */
function findExtraField(b, start, length, id) {
	const end = start + length;
	let found = null;
	for (let q = start; q < end;) {
		if (q + 4 > end) throw corrupt("A ZIP extra field is damaged.", "extra-field", q);
		const blockId = b.u16(q);
		const size = b.u16(q + 2);
		if (q + 4 + size > end) throw corrupt("A ZIP extra field is damaged.", "extra-field", q);
		if (blockId === id && found === null) found = {
			start: q + 4,
			end: q + 4 + size
		};
		q += 4 + size;
	}
	return found;
}
/**
* Checks each local file header against its central directory entry and rejects archives in
* which entries overlap each other or the central directory (the "better zip bomb" technique,
* W-06, W-07). Fills in the data offsets.
*/
function validateLocalHeaders(b, entries, centralOffset) {
	const spans = [];
	entries.forEach((entry, index) => {
		const at = entry.headerOffset;
		if (at + 30 > centralOffset || b.u32(at) !== SIG_LOCAL$1) throw corrupt(`The local header of "${entry.name}" is missing or damaged.`, "local-header", at, { name: entry.name });
		const end = at + 30 + b.u16(at + 26) + b.u16(at + 28) + entry.compressedSize;
		if (end > centralOffset) throw corrupt(`The data of "${entry.name}" runs into the central directory.`, "entry-bounds", at, { name: entry.name });
		spans.push({
			start: at,
			end,
			entry,
			index
		});
	});
	spans.sort((x, y) => x.start - y.start);
	const first = spans[0];
	if (first !== void 0 && first.start !== 0) throw corrupt("The archive does not start with a ZIP entry.", "leading-data", 0);
	for (let i = 1; i < spans.length; i++) {
		const prev = spans[i - 1];
		const cur = spans[i];
		if (prev !== void 0 && cur !== void 0 && cur.start < prev.end) throw corrupt(`ZIP entries "${prev.entry.name}" and "${cur.entry.name}" overlap.`, "overlapping-entries", cur.start, { name: cur.entry.name });
	}
	for (const span of spans) {
		const { entry } = span;
		const at = entry.headerOffset;
		const nameLength = b.u16(at + 26);
		if (!sameBytes(b.slice(at + 30, at + 30 + nameLength), entry.nameBytes) || b.u16(at + 8) !== entry.method) throw corrupt(`The local header of "${entry.name}" does not match the central directory.`, "local-header-mismatch", at, { name: entry.name });
		entries[span.index] = {
			...entry,
			dataOffset: at + 30 + nameLength + b.u16(at + 28)
		};
	}
}
function sameBytes(a, b) {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
	return true;
}
var Reader = class {
	b;
	limits;
	entries;
	byName;
	bytesRead = 0;
	constructor(b, entries, limits) {
		this.b = b;
		this.limits = limits;
		this.entries = entries.map(({ nameBytes: _nameBytes, ...entry }) => entry);
		this.byName = new Map(this.entries.map((e) => [e.name.toLowerCase(), e]));
	}
	find(name) {
		return this.byName.get(name.toLowerCase());
	}
	raw(entry) {
		if (entry.method === 0 && entry.compressedSize === entry.uncompressedSize) return {
			method: 0,
			crc32: entry.crc32,
			uncompressedSize: entry.uncompressedSize,
			body: this.b.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize)
		};
		if (entry.method === 8) return {
			method: 8,
			crc32: entry.crc32,
			uncompressedSize: entry.uncompressedSize,
			body: this.b.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize)
		};
		return null;
	}
	read(entryOrName) {
		const entry = typeof entryOrName === "string" ? this.find(entryOrName) : entryOrName;
		if (entry === void 0) throw new TvError("TV_E_PART_MISSING", `The archive has no entry "${String(entryOrName)}".`, {
			location: { part: `/${String(entryOrName)}` },
			hint: "The document is incomplete. Re-save it in the application that made it."
		});
		const { limits } = this;
		const counted = this.counted.has(entry);
		const base = counted ? this.bytesRead - entry.uncompressedSize : this.bytesRead;
		if (entry.uncompressedSize > limits.maxEntryBytes) throw limitExceeded("maxEntryBytes", limits.maxEntryBytes, `ZIP entry "${entry.name}" declares ${entry.uncompressedSize} bytes, more than the limit.`, entry.name);
		if (base + entry.uncompressedSize > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
		const compressed = this.b.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize);
		let out;
		if (entry.method === 0) {
			if (entry.compressedSize !== entry.uncompressedSize) throw corrupt(`Stored ZIP entry "${entry.name}" has inconsistent sizes.`, "size-mismatch", entry.headerOffset, { name: entry.name });
			out = compressed;
		} else if (entry.method === 8) out = this.inflate(entry, compressed, base);
		else throw unsupported$1(`ZIP entry "${entry.name}" uses compression method ${entry.method}.`, "compression-method", "Only stored and Deflate-compressed entries are supported. Re-save the document in the application that made it.", entry.name);
		if (out.length !== entry.uncompressedSize) throw corrupt(`ZIP entry "${entry.name}" is ${out.length} bytes, but its header says ${entry.uncompressedSize}.`, "size-mismatch", entry.headerOffset, { name: entry.name });
		if (crc32(out) !== entry.crc32) throw corrupt(`ZIP entry "${entry.name}" failed its CRC-32 check.`, "crc-mismatch", entry.headerOffset, { name: entry.name });
		if (!counted) {
			this.bytesRead += out.length;
			this.counted.add(entry);
		}
		return out;
	}
	counted = /* @__PURE__ */ new Set();
	inflate(entry, compressed, base) {
		const { limits } = this;
		const chunks = [];
		let produced = 0;
		let ended = false;
		const inflater = new Inflate();
		inflater.ondata = (chunk, final) => {
			produced += chunk.length;
			if (produced > limits.maxEntryBytes) throw limitExceeded("maxEntryBytes", limits.maxEntryBytes, `ZIP entry "${entry.name}" decompresses to more than the limit.`, entry.name);
			if (base + produced > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
			if (produced > limits.compressionRatioMinBytes && produced / Math.max(1, entry.compressedSize) > limits.maxCompressionRatio) throw limitExceeded("maxCompressionRatio", limits.maxCompressionRatio, `ZIP entry "${entry.name}" expands more than ${String(limits.maxCompressionRatio)} times, which is typical of a zip bomb.`, entry.name);
			if (chunk.length > 0) chunks.push(chunk);
			if (final) ended = true;
		};
		try {
			if (compressed.length === 0) inflater.push(compressed, true);
			for (let at = 0; at < compressed.length; at += INFLATE_CHUNK) inflater.push(compressed.subarray(at, at + INFLATE_CHUNK), at + INFLATE_CHUNK >= compressed.length);
		} catch (error) {
			if (error instanceof TvError) throw error;
			throw new TvError("TV_E_ZIP_CORRUPT", `ZIP entry "${entry.name}" is not a valid Deflate stream.`, {
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
		}
		if (!ended) throw corrupt(`ZIP entry "${entry.name}" ends early.`, "deflate", entry.headerOffset, { name: entry.name });
		if (chunks.length === 1) return chunks[0];
		const out = new Uint8Array(produced);
		let at = 0;
		for (const chunk of chunks) {
			out.set(chunk, at);
			at += chunk.length;
		}
		return out;
	}
};
// src/xml/raw-tree.ts
/** Parses `xml` into a byte-preserving tree. Throws the errors of `parseXml`. */
function parseRawTree(xml, options) {
	const maxNodes = options.limits.maxXmlTreeElements ?? Number.POSITIVE_INFINITY;
	let nodes = 0;
	const count = () => {
		if (++nodes > maxNodes) throw new TvError("TV_E_LIMIT_EXCEEDED", "The part has more XML elements than the limit.", {
			location: { part: options.part },
			hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTreeElements: ${String(maxNodes * 4)} } }.`,
			details: {
				limit: "maxXmlTreeElements",
				value: maxNodes
			}
		});
	};
	const source = normalizeLineEnds(xml);
	const stack = [];
	let root;
	let prolog = "";
	let lastEnd = 0;
	let pending = "";
	const flush = (until) => {
		if (until > lastEnd) {
			const raw = source.slice(lastEnd, until);
			const parent = stack[stack.length - 1];
			if (parent === void 0) {
				if (root === void 0) prolog += raw;
			} else {
				count();
				parent.children.push({
					kind: "text",
					raw,
					text: pending,
					parent,
					start: lastEnd
				});
			}
		}
		pending = "";
	};
	parseXml(source, {
		start(ns, local, attrs, span) {
			flush(span.start);
			count();
			const parent = stack[stack.length - 1] ?? null;
			const element = {
				kind: "element",
				ns,
				local,
				qname: span.qname,
				attrs,
				open: source.slice(span.start, span.end),
				close: "",
				children: [],
				parent,
				start: span.start
			};
			if (parent === null) root = element;
			else parent.children.push(element);
			stack.push(element);
			lastEnd = span.end;
		},
		end(_ns, _local, span) {
			flush(span.start);
			const element = stack.pop();
			if (!span.empty) element.close = source.slice(span.start, span.end);
			lastEnd = span.end;
		},
		text(value) {
			pending += value;
		}
	}, options);
	return {
		prolog,
		root,
		epilog: source.slice(lastEnd)
	};
}
/** Appends the serialisation of `node` to `out`. Iterative, so deep trees do not grow the stack. */
function serializeInto(node, out) {
	const work = [node];
	while (work.length > 0) {
		const item = work.pop();
		if (typeof item === "string") out.push(item);
		else if (item.kind === "text") out.push(item.raw);
		else {
			out.push(item.open);
			if (item.close !== "") work.push(item.close);
			for (let i = item.children.length - 1; i >= 0; i--) work.push(item.children[i]);
		}
	}
}
/** The text of one node and everything inside it. */
function rawXmlText(node) {
	const out = [];
	serializeInto(node, out);
	return out.join("");
}
/** The text of a whole document. */
function serialize$2(doc) {
	const out = [doc.prolog];
	serializeInto(doc.root, out);
	out.push(doc.epilog);
	return out.join("");
}
/** The prefix of a qualified name, with the colon (`w:` for `w:p`), or '' when there is none. */
function prefixOf(qname) {
	const colon = qname.indexOf(":");
	return colon === -1 ? "" : qname.slice(0, colon + 1);
}
/** A new element without source position. `open` and `close` are written by the caller. */
function makeElement(model, open, close, attrs = /* @__PURE__ */ new Map()) {
	return {
		kind: "element",
		ns: model.ns,
		local: model.local,
		qname: model.qname,
		attrs,
		open,
		close,
		children: [],
		parent: null,
		start: -1
	};
}
// src/docx/paragraphs.ts
/** WordprocessingML namespaces: Transitional and Strict (ECMA-376 Part 1 and Part 4). */
const WML_NAMESPACES = /* @__PURE__ */ new Set(["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"]);
const NOT_TEXT_CONTAINERS = /* @__PURE__ */ new Set(["del", "moveFrom"]);
const RUN_CHARACTERS = /* @__PURE__ */ new Map([
	["tab", "	"],
	["br", "\n"],
	["cr", "\n"],
	["noBreakHyphen", "‑"],
	["softHyphen", "­"],
	["sym", "￼"],
	["drawing", "￼"],
	["pict", "￼"],
	["object", "￼"],
	["footnoteReference", "￼"],
	["endnoteReference", "￼"]
]);
const SKIPPED_NOTE_TYPES = /* @__PURE__ */ new Set([
	"separator",
	"continuationSeparator",
	"continuationNotice"
]);
const isWml$2 = (e, local) => WML_NAMESPACES.has(e.ns) && (local === void 0 || e.local === local);
const isRun = (e) => e !== null && isWml$2(e, "r");
/** A container in words, as in location descriptions. */
function describeContainer(c) {
	switch (c.kind) {
		case "table": return `table ${String(c.table)} row ${String(c.row)} cell ${String(c.cell)}`;
		case "textBox": return `text box ${String(c.textBox)}`;
		case "shape": return c.name === "" ? `shape ${String(c.shape)}` : `shape "${c.name}"`;
	}
}
/** The branch of an mc:AlternateContent that is read: the first mc:Choice, else mc:Fallback. */
function selectedBranch(alternate) {
	let fallback;
	for (const child of alternate.children) {
		if (child.kind !== "element" || child.ns !== "http://schemas.openxmlformats.org/markup-compatibility/2006") continue;
		if (child.local === "Choice") return child;
		if (child.local === "Fallback") fallback ??= child;
	}
	return fallback;
}
/** Walks the paragraphs of one story part, in document order (a paragraph before those inside it). */
function walkParagraphs(doc, options) {
	const { part, story } = options;
	const out = [];
	const storyPrefix = story === "header" || story === "footer" ? `${story} ${part}` : story;
	const scopes = [];
	const tables = [];
	const paragraphs = [];
	const alternates = [];
	let notes = 0;
	const scope = () => scopes[scopes.length - 1];
	const pushScope = (within, prefix, note) => {
		scopes.push({
			within,
			prefix,
			note,
			paragraphs: 0,
			tables: 0,
			textBoxes: 0
		});
	};
	const visit = (e) => {
		const para = paragraphs[paragraphs.length - 1];
		if (e.ns === "http://schemas.openxmlformats.org/markup-compatibility/2006") {
			if (e.local === "AlternateContent") {
				if (para !== void 0 && isRun(e.parent)) {
					para.pieces.push({
						kind: "char",
						element: e,
						textStart: para.text.length,
						textEnd: para.text.length + 1
					});
					para.text += "￼";
				}
				const branch = selectedBranch(e);
				if (branch !== void 0) {
					alternates.push(e);
					visitChildren(branch);
					alternates.pop();
				}
				return;
			}
			visitChildren(e);
			return;
		}
		if (!isWml$2(e)) {
			visitChildren(e);
			return;
		}
		if (NOT_TEXT_CONTAINERS.has(e.local)) return;
		switch (e.local) {
			case "body":
			case "hdr":
			case "ftr":
				if (scopes.length === 0) pushScope([], storyPrefix, void 0);
				visitChildren(e);
				return;
			case "footnote":
			case "endnote": {
				const type = e.attrs.get(`{${e.ns}}type`);
				if (type !== void 0 && SKIPPED_NOTE_TYPES.has(type)) return;
				notes++;
				pushScope([], `${e.local} ${String(notes)}`, notes);
				visitChildren(e);
				scopes.pop();
				return;
			}
			case "tbl": {
				const s = scope();
				if (s === void 0) break;
				tables.push({
					scope: s,
					index: ++s.tables,
					rows: 0,
					cells: 0
				});
				visitChildren(e);
				tables.pop();
				return;
			}
			case "tr": {
				const t = tables[tables.length - 1];
				if (t !== void 0) {
					t.rows++;
					t.cells = 0;
				}
				break;
			}
			case "tc": {
				const t = tables[tables.length - 1];
				if (t === void 0) break;
				const container = {
					kind: "table",
					table: t.index,
					row: t.rows,
					cell: ++t.cells
				};
				pushScope([...t.scope.within, container], `${t.scope.prefix}, ${describeContainer(container)}`, t.scope.note);
				visitChildren(e);
				scopes.pop();
				return;
			}
			case "txbxContent": {
				const s = scope();
				if (s === void 0) break;
				const container = {
					kind: "textBox",
					textBox: ++s.textBoxes
				};
				pushScope([...s.within, container], `${s.prefix}, ${describeContainer(container)}`, s.note);
				visitChildren(e);
				scopes.pop();
				return;
			}
			case "p": {
				const s = scope();
				if (s === void 0) break;
				const paragraph = ++s.paragraphs;
				const builder = {
					element: e,
					location: {
						part,
						story,
						...s.note === void 0 ? {} : { note: s.note },
						...s.within.length === 0 ? {} : { within: s.within },
						paragraph,
						description: `${s.prefix}, paragraph ${String(paragraph)}`
					},
					alternates: [...alternates],
					text: "",
					pieces: []
				};
				out.push(builder);
				paragraphs.push(builder);
				visitChildren(e);
				paragraphs.pop();
				return;
			}
			case "t":
				if (para !== void 0 && isRun(e.parent)) for (const child of e.children) {
					if (child.kind !== "text" || child.text === "") continue;
					para.pieces.push({
						kind: "text",
						node: child,
						t: e,
						run: e.parent,
						textStart: para.text.length,
						textEnd: para.text.length + child.text.length
					});
					para.text += child.text;
				}
				return;
			default: {
				const char = RUN_CHARACTERS.get(e.local);
				if (char !== void 0 && para !== void 0 && isRun(e.parent)) {
					para.pieces.push({
						kind: "char",
						element: e,
						textStart: para.text.length,
						textEnd: para.text.length + 1
					});
					para.text += char;
					if (e.local === "pict" || e.local === "drawing" || e.local === "object") visitChildren(e);
					return;
				}
			}
		}
		visitChildren(e);
	};
	const visitChildren = (e) => {
		for (const child of e.children) if (child.kind === "element") visit(child);
	};
	visit(doc.root);
	return out.map((b) => ({
		element: b.element,
		text: b.text,
		pieces: b.pieces,
		location: b.location,
		alternates: b.alternates
	}));
}
/** The source offset of a character of a paragraph, for ordering tags by their place in the part. */
function sourceOffset(record, textOffset) {
	const pieces = record.pieces;
	let lo = 0;
	let hi = pieces.length - 1;
	while (lo <= hi) {
		const mid = lo + hi >> 1;
		const piece = pieces[mid];
		if (textOffset < piece.textStart) hi = mid - 1;
		else if (textOffset >= piece.textEnd) lo = mid + 1;
		else return piece.kind === "text" ? piece.node.start + (textOffset - piece.textStart) : piece.element.start;
	}
	return record.element.start;
}
// src/pptx/paragraphs.ts
/** DrawingML: Transitional and Strict. */
const DML_NAMESPACES = /* @__PURE__ */ new Set(["http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main"]);
/** PresentationML: Transitional and Strict. */
const PML_NAMESPACES = /* @__PURE__ */ new Set(["http://schemas.openxmlformats.org/presentationml/2006/main", "http://purl.oclc.org/ooxml/presentationml/main"]);
const SHAPES = /* @__PURE__ */ new Set([
	"sp",
	"graphicFrame",
	"pic",
	"cxnSp",
	"grpSp",
	"contentPart"
]);
const isDml = (e, local) => DML_NAMESPACES.has(e.ns) && (local === void 0 || e.local === local);
const isPml = (e, local) => PML_NAMESPACES.has(e.ns) && (local === void 0 || e.local === local);
/** The name of a shape: p:cNvPr/@name inside its non-visual properties. */
function shapeName(shape) {
	for (const nv of shape.children) {
		if (nv.kind !== "element" || !isPml(nv) || !nv.local.startsWith("nv")) continue;
		for (const c of nv.children) if (c.kind === "element" && isPml(c, "cNvPr")) return c.attrs.get("name") ?? "";
	}
	return "";
}
/** Walks the paragraphs of one slide or notes part, in document order. */
function walkSlideParagraphs(doc, options) {
	const { part, story, slide } = options;
	const out = [];
	const scopes = [{
		within: [],
		prefix: story === "notes" ? `slide ${String(slide)} notes` : `slide ${String(slide)}`,
		paragraphs: 0,
		shapes: 0,
		tables: 0
	}];
	const tables = [];
	const alternates = [];
	const scope = () => scopes[scopes.length - 1];
	const push = (container) => {
		const s = scope();
		scopes.push({
			within: [...s.within, container],
			prefix: `${s.prefix}, ${describeContainer(container)}`,
			paragraphs: 0,
			shapes: 0,
			tables: 0
		});
	};
	const paragraph = (p) => {
		const s = scope();
		const n = ++s.paragraphs;
		const location = {
			part,
			story,
			slide,
			...s.within.length === 0 ? {} : { within: s.within },
			paragraph: n,
			description: `${s.prefix}, paragraph ${String(n)}`
		};
		const pieces = [];
		let text = "";
		const visitRun = (e) => {
			for (const c of e.children) {
				if (c.kind !== "element") continue;
				if (isDml(c, "r")) for (const t of c.children) {
					if (t.kind !== "element" || !isDml(t, "t")) continue;
					for (const node of t.children) {
						if (node.kind !== "text" || node.text === "") continue;
						pieces.push({
							kind: "text",
							node,
							t,
							run: c,
							textStart: text.length,
							textEnd: text.length + node.text.length
						});
						text += node.text;
					}
				}
				else if (isDml(c, "br") || isDml(c, "fld")) {
					pieces.push({
						kind: "char",
						element: c,
						textStart: text.length,
						textEnd: text.length + 1
					});
					text += isDml(c, "br") ? "\n" : "￼";
				}
			}
		};
		visitRun(p);
		out.push({
			element: p,
			text,
			pieces,
			location,
			alternates: [...alternates]
		});
	};
	const visit = (e) => {
		if (e.ns === "http://schemas.openxmlformats.org/markup-compatibility/2006" && e.local === "AlternateContent") {
			const branch = selectedBranch(e);
			if (branch !== void 0) {
				alternates.push(e);
				visitChildren(branch);
				alternates.pop();
			}
			return;
		}
		if (isPml(e) && SHAPES.has(e.local)) {
			const s = scope();
			push({
				kind: "shape",
				shape: ++s.shapes,
				name: shapeName(e)
			});
			visitChildren(e);
			scopes.pop();
			return;
		}
		if (isDml(e, "tbl")) {
			const s = scope();
			tables.push({
				scope: s,
				index: ++s.tables,
				rows: 0,
				cells: 0
			});
			visitChildren(e);
			tables.pop();
			return;
		}
		if (isDml(e, "tr")) {
			const t = tables[tables.length - 1];
			if (t !== void 0) {
				t.rows++;
				t.cells = 0;
			}
		} else if (isDml(e, "tc")) {
			const t = tables[tables.length - 1];
			if (t !== void 0) {
				push({
					kind: "table",
					table: t.index,
					row: t.rows,
					cell: ++t.cells
				});
				visitChildren(e);
				scopes.pop();
				return;
			}
		} else if (isDml(e, "p")) {
			paragraph(e);
			return;
		}
		visitChildren(e);
	};
	const visitChildren = (e) => {
		for (const c of e.children) if (c.kind === "element") visit(c);
	};
	visit(doc.root);
	return out;
}
/** The shape (p:sp) that holds a paragraph, if any. */
function shapeOf(paragraph) {
	for (let e = paragraph.parent; e !== null; e = e.parent) if (isPml(e, "sp")) return e;
}
/** The text of all paragraphs of a shape joined by line ends, for "alone in its shape" checks. */
function shapeText(shape, records) {
	return records.filter((r) => shapeOf(r.element) === shape).map((r) => r.text).join("\n");
}
// src/pptx/package.ts
const STRICT_REL$2 = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_REL$2 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const R_NAMESPACES = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"];
/** Whether a relationship has the given name in either conformance class. */
function isRelationship(rel, name) {
	return !rel.external && (rel.type === TRANSITIONAL_REL$2 + name || rel.type === STRICT_REL$2 + name);
}
/** The r:id of an element, in either conformance class. */
function relId(e) {
	for (const ns of R_NAMESPACES) {
		const id = e.attrs.get(`{${ns}}id`);
		if (id !== void 0) return id;
	}
}
/** The slides of a presentation in order. */
function presentationSlides(pkg, main, limits) {
	const list = parseXmlTree(pkg.readXml(main) ?? "", {
		part: main,
		limits
	}).children.find((c) => PML_NAMESPACES.has(c.ns) && c.local === "sldIdLst");
	const rels = pkg.relationships(main);
	const slides = [];
	const warnings = [];
	for (const entry of list?.children ?? []) {
		if (!PML_NAMESPACES.has(entry.ns) || entry.local !== "sldId") continue;
		const rId = relId(entry);
		const rel = rels.find((r) => r.id === rId);
		if (rId === void 0 || rel === void 0 || !isRelationship(rel, "slide") || !pkg.hasPart(rel.target)) {
			warnings.push({
				code: "TV_W_PART_MISSING",
				message: "The presentation lists a slide whose part does not exist; it was skipped.",
				location: { part: main },
				hint: "Re-save the presentation in PowerPoint or LibreOffice to repair it."
			});
			continue;
		}
		const notesRel = pkg.relationships(rel.target).find((r) => isRelationship(r, "notesSlide"));
		const notes = notesRel !== void 0 && pkg.hasPart(notesRel.target) ? notesRel.target : void 0;
		slides.push({
			part: rel.target,
			rId,
			id: entry.attrs.get("id") ?? "",
			notes
		});
	}
	return {
		slides,
		warnings
	};
}
/** The stories of a presentation: each slide, then its notes. */
function presentationStories(pkg, main, limits) {
	const { slides, warnings } = presentationSlides(pkg, main, limits);
	const stories = [];
	slides.forEach((s, i) => {
		stories.push({
			part: s.part,
			story: "slide",
			slide: i + 1
		});
		if (s.notes !== void 0) stories.push({
			part: s.notes,
			story: "notes",
			slide: i + 1
		});
	});
	return {
		stories,
		warnings
	};
}
// src/docx/package.ts
const STRICT_REL$1 = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_REL$1 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
/** The relationship name of a type URI of either conformance class, or null. */
function relationshipName(type) {
	if (type.startsWith(STRICT_REL$1)) return type.slice(56);
	if (type.startsWith(TRANSITIONAL_REL$1)) return type.slice(68);
	return null;
}
/** SpreadsheetML namespaces: Transitional and Strict (ECMA-376 Part 1 and Part 4), as in src/sml/namespaces.ts of xlvane. */
const SML_NAMESPACES = /* @__PURE__ */ new Set(["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"]);
/** The relationships of the main part that lead to stories, in reading order. */
const STORY_RELATIONSHIPS = [
	["header", "header"],
	["footer", "footer"],
	["footnotes", "footnote"],
	["endnotes", "endnote"]
];
var RootFound = class extends Error {
	ns;
	local;
	constructor(ns, local) {
		super("root");
		this.ns = ns;
		this.local = local;
	}
};
/** The namespace and local name of the root element; parses only up to the first start tag. */
function rootElement(xml, part, limits) {
	try {
		parseXml(xml, {
			start(ns, local) {
				throw new RootFound(ns, local);
			},
			end() {},
			text() {}
		}, {
			part,
			limits
		});
	} catch (error) {
		if (error instanceof RootFound) return {
			ns: error.ns,
			local: error.local
		};
		throw error;
	}
	throw new TvError("TV_E_INTERNAL", "The main part has no root element.", {
		location: { part },
		hint: "This is an internal error; please report it with the file."
	});
}
function find(rels, name) {
	return rels.find((r) => !r.external && relationshipName(r.type) === name);
}
/**
* Opens `bytes` as a template package: a document, a presentation or a workbook. Throws
* `TV_E_NOT_SUPPORTED` for other packages, and the ZIP, XML and package errors of the layers below.
*/
function openDocx(bytes, limits) {
	const zip = openZip(bytes, limits);
	const pkg = openPackage(zip, limits);
	const officeDocument = find(pkg.relationships(null), "officeDocument");
	if (officeDocument === void 0 || !pkg.hasPart(officeDocument.target)) throw new TvError("TV_E_PART_MISSING", "The package has no main document part.", {
		location: officeDocument === void 0 ? { part: "/_rels/.rels" } : { part: officeDocument.target },
		hint: "This is not a complete Office document: the package does not point to its main part. Check that the file is a .docx, .pptx or .xlsx file and re-save it in the application that made it."
	});
	const main = officeDocument.target;
	const root = rootElement(pkg.readXml(main) ?? "", main, limits);
	if (PML_NAMESPACES.has(root.ns) && root.local === "presentation") {
		const { stories, warnings } = presentationStories(pkg, main, limits);
		return {
			format: "pptx",
			zip,
			pkg,
			main,
			stories,
			warnings
		};
	}
	if (SML_NAMESPACES.has(root.ns) && root.local === "workbook") return {
		format: "xlsx",
		zip,
		pkg,
		main,
		stories: [],
		warnings: []
	};
	if (!(WML_NAMESPACES.has(root.ns) && root.local === "document")) throw new TvError("TV_E_NOT_SUPPORTED", `The file is not a Word document, a PowerPoint presentation or an Excel workbook: its main part is a "${clip(root.local)}" element.`, {
		location: { part: main },
		hint: "Pass a Word document (.docx), a PowerPoint presentation (.pptx) or an Excel workbook (.xlsx).",
		details: {
			root: clip(root.local),
			namespace: clip(root.ns)
		}
	});
	const warnings = [];
	const stories = [{
		part: main,
		story: "body"
	}];
	const seen = /* @__PURE__ */ new Set([partKey(main)]);
	const rels = pkg.relationships(main);
	for (const [name, story] of STORY_RELATIONSHIPS) for (const rel of rels) {
		if (rel.external || relationshipName(rel.type) !== name || seen.has(partKey(rel.target))) continue;
		seen.add(partKey(rel.target));
		if (!pkg.hasPart(rel.target)) {
			warnings.push({
				code: "TV_W_PART_MISSING",
				message: `The document refers to a ${story} part that does not exist; it was skipped.`,
				location: { part: rel.target },
				hint: "Re-save the document in Word or LibreOffice to repair it."
			});
			continue;
		}
		stories.push({
			part: rel.target,
			story
		});
	}
	return {
		format: "docx",
		zip,
		pkg,
		main,
		stories,
		warnings
	};
}
// src/expr/ast.ts
const PRECEDENCE = {
	pipe: 0,
	or: 1,
	and: 2,
	compare: 3,
	not: 4,
	literal: 5,
	path: 5
};
function literalText(value) {
	return typeof value === "string" ? JSON.stringify(value) : String(value);
}
/** The text of a formatter call, with arguments in canonical form. */
function formatterText(f) {
	return f.args.length === 0 ? f.name : `${f.name}(${f.args.map(literalText).join(", ")})`;
}
/**
* The canonical text of an expression: single spaces around operators, straight double quotes,
* parentheses only where precedence needs them. Two tags with the same canonical text mean the
* same thing, which is how a closing tag finds its section.
*/
function canonical(e) {
	const wrap = (child, min) => PRECEDENCE[child.t] < min ? `(${canonical(child)})` : canonical(child);
	switch (e.t) {
		case "literal": return literalText(e.value);
		case "path": return e.segments.length === 0 ? "." : e.segments.join(".");
		case "not": return `!${wrap(e.e, PRECEDENCE.not)}`;
		case "and":
		case "or": return `${wrap(e.l, PRECEDENCE[e.t])} ${e.t === "and" ? "&&" : "||"} ${wrap(e.r, PRECEDENCE[e.t] + 1)}`;
		case "compare": return `${wrap(e.l, PRECEDENCE.not)} ${e.op} ${wrap(e.r, PRECEDENCE.not)}`;
		case "pipe": return `${wrap(e.e, PRECEDENCE.or)} | ${e.formatters.map(formatterText).join(" | ")}`;
	}
}
/** Every data path the expression reads, as text, in order of first use. */
function pathsOf$1(e) {
	const out = [];
	const walk = (x) => {
		switch (x.t) {
			case "path": {
				const text = x.segments.length === 0 ? "." : x.segments.join(".");
				if (!out.includes(text)) out.push(text);
				return;
			}
			case "not":
			case "pipe":
				walk(x.e);
				return;
			case "and":
			case "or":
			case "compare":
				walk(x.l);
				walk(x.r);
				return;
		}
	};
	walk(e);
	return out;
}
/** Every formatter call in the expression, nested ones included. */
function formattersOf(e) {
	const out = [];
	const walk = (x) => {
		switch (x.t) {
			case "pipe":
				walk(x.e);
				out.push(...x.formatters);
				return;
			case "not":
				walk(x.e);
				return;
			case "and":
			case "or":
			case "compare":
				walk(x.l);
				walk(x.r);
				return;
		}
	};
	walk(e);
	return out;
}
// src/docx/urls.ts
const MAX_URL_LENGTH = 2048;
/** null when `url` may be a link target, otherwise a message that says why not. */
function checkLinkUrl(url) {
	if (url.length > MAX_URL_LENGTH) return `The address is longer than ${String(MAX_URL_LENGTH)} characters.`;
	for (let i = 0; i < url.length; i++) {
		const c = url.charCodeAt(i);
		if (c <= 32 || c === 127 || c === 34 || c === 60 || c === 62 || c === 92 || c === 96 || c >= 128 && c <= 159) return "The address contains a space, a quote, an angle bracket, a backslash or a control character.";
		if (c >= 55296 && c <= 56319 && !(url.charCodeAt(i + 1) >= 56320 && url.charCodeAt(i + 1) <= 57343) || c >= 56320 && c <= 57343 && !(url.charCodeAt(i - 1) >= 55296 && url.charCodeAt(i - 1) <= 56319) || c === 65534 || c === 65535 || c >= 64976 && c <= 65007 || c >= 8203 && c <= 8207 || c >= 8234 && c <= 8238 || c >= 8294 && c <= 8297 || c === 1564 || c === 65279) return "The address contains a character that cannot be in a document or that changes how text is shown (such as a bidirectional control).";
	}
	const lower = url.toLowerCase();
	if (lower.startsWith("http://") || lower.startsWith("https://")) {
		const host = url.slice(lower.startsWith("https://") ? 8 : 7).split(/[/?#]/, 1)[0] ?? "";
		return host === "" || host.startsWith("@") ? "The address has no host." : null;
	}
	if (lower.startsWith("mailto:")) return url.length > 7 ? null : "The mailto address is empty.";
	return "Only http, https and mailto addresses can be links.";
}
// src/expr/formatters.ts
/** Thrown by a formatter for a value it cannot format; becomes TV_E_INVALID_VALUE with the tag's place. */
var ValueError = class extends Error {
	hint;
	constructor(message, hint) {
		super(message);
		this.hint = hint;
	}
};
/** The result of the `link` formatter: the renderer writes it as a hyperlink. */
var LinkValue = class {
	url;
	text;
	constructor(url, text) {
		this.url = url;
		this.text = text;
	}
};
const DATE_STYLES = /* @__PURE__ */ new Set([
	"short",
	"medium",
	"long",
	"full"
]);
const MAX_DECIMALS = 20;
/** Intl objects are costly to create, so a render keeps them. */
var Formatting = class {
	ctx;
	cache = /* @__PURE__ */ new Map();
	constructor(ctx) {
		this.ctx = ctx;
	}
	number(options) {
		const key = `n${JSON.stringify(options)}`;
		let f = this.cache.get(key);
		if (f === void 0) {
			f = new Intl.NumberFormat(this.ctx.locale, {
				signDisplay: "negative",
				...options
			});
			this.cache.set(key, f);
		}
		return f;
	}
	date(locale, options) {
		const key = `d${locale}${JSON.stringify(options)}`;
		let f = this.cache.get(key);
		if (f === void 0) {
			f = new Intl.DateTimeFormat(locale, options);
			this.cache.set(key, f);
		}
		return f;
	}
};
function describe(value) {
	if (value === null) return "null";
	if (Array.isArray(value)) return "a list";
	if (value instanceof Date) return "a date";
	if (typeof value === "string") return `the text ${JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value)}`;
	if (typeof value === "object") return "an object";
	return `${typeof value} ${String(value)}`;
}
/** Text formatters accept strings, numbers and booleans; missing and null values pass through. */
function text(op, name) {
	return (value, _args, ctx) => {
		if (value === void 0 || value === null) return value;
		if (typeof value === "string") return op(value, ctx);
		if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") return op(String(value), ctx);
		throw new ValueError(`${name} cannot format ${describe(value)}.`, `Pass text, a number or true/false to ${name}.`);
	};
}
function finiteNumber(value, name) {
	if (typeof value === "number" && Number.isFinite(value)) return value;
	throw new ValueError(`${name} needs a number, but the value is ${describe(value)}.`, "Pass a number in the data, for example 1234.5, not text such as \"1234,50\".");
}
function decimals(args, at) {
	const d = args[at];
	return typeof d === "number" ? {
		minimumFractionDigits: d,
		maximumFractionDigits: d
	} : {};
}
const checkDecimals = (at) => (args) => {
	const d = args[at];
	return typeof d === "number" && (d < 0 || d > MAX_DECIMALS) ? `The number of decimals must be from 0 to ${String(MAX_DECIMALS)}.` : null;
};
const ISO = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:\d{2})?)?$/;
/** The instant of a date value, and the time zone to show it in. Times without an offset are shown as written. */
function toInstant(value, ctx) {
	if (value instanceof Date) {
		const time = value.getTime();
		if (Number.isNaN(time)) throw new ValueError("date cannot format an invalid Date.", "Pass a valid Date or an ISO 8601 text such as \"2026-11-06\".");
		return {
			time,
			timeZone: ctx.ctx.timeZone
		};
	}
	if (typeof value === "string") {
		const m = ISO.exec(value);
		if (m !== null) {
			const [, y, mo, d, h = "0", mi = "0", s = "0", frac = "", offset] = m;
			const ms = Number(frac.padEnd(3, "0").slice(0, 3));
			const day = new Date(Date.UTC(2e3, 0, 1, Number(h), Number(mi), Number(s), ms));
			day.setUTCFullYear(Number(y), Number(mo) - 1, Number(d));
			const local = day.getTime();
			const check = new Date(local);
			if (check.getUTCFullYear() === Number(y) && check.getUTCMonth() === Number(mo) - 1 && check.getUTCDate() === Number(d) && check.getUTCHours() === Number(h) && check.getUTCMinutes() === Number(mi) && check.getUTCSeconds() === Number(s)) {
				if (offset === void 0) return {
					time: local,
					timeZone: "UTC"
				};
				if (offset === "Z") return {
					time: local,
					timeZone: ctx.ctx.timeZone
				};
				if (Number(offset.slice(1, 3)) > 23 || Number(offset.slice(4, 6)) > 59) throw new ValueError(`date cannot format ${describe(value)}: the time zone offset is not valid.`, "Use an offset such as +01:00 or -05:30.");
				return {
					time: local - (offset.startsWith("-") ? -1 : 1) * (Number(offset.slice(1, 3)) * 60 + Number(offset.slice(4, 6))) * 6e4,
					timeZone: ctx.ctx.timeZone
				};
			}
		}
	}
	throw new ValueError(`date cannot format ${describe(value)}.`, "Pass a Date or an ISO 8601 text such as \"2026-11-06\" or \"2026-11-06T14:30:00Z\".");
}
const PATTERN_FIELDS = /* @__PURE__ */ new Set([
	"yyyy",
	"yy",
	"MMMM",
	"MMM",
	"MM",
	"M",
	"dd",
	"d",
	"HH",
	"H",
	"mm",
	"ss"
]);
/** Splits a date pattern into tokens, or returns a message for a pattern that is not valid. */
function parseDatePattern(pattern) {
	const out = [];
	let i = 0;
	while (i < pattern.length) {
		const c = pattern[i];
		if (c === "'") {
			let literal = "";
			i++;
			for (;;) {
				if (i >= pattern.length) return "A quote in the date pattern is not closed.";
				if (pattern[i] === "'") {
					if (pattern[i + 1] === "'") {
						literal += "'";
						i += 2;
						continue;
					}
					i++;
					break;
				}
				literal += pattern[i];
				i++;
			}
			out.push({ literal: literal === "" ? "'" : literal });
			continue;
		}
		if (c >= "a" && c <= "z" || c >= "A" && c <= "Z") {
			let j = i;
			while (pattern[j] === c) j++;
			const run = pattern.slice(i, j);
			if (!PATTERN_FIELDS.has(run)) return `"${run}" is not a field of a date pattern.`;
			out.push({ field: run });
			i = j;
			continue;
		}
		out.push({ literal: c });
		i++;
	}
	return out;
}
function formatPattern(tokens, time, timeZone, ctx) {
	const parts = /* @__PURE__ */ new Map();
	for (const p of ctx.date("en-US", {
		timeZone,
		year: "numeric",
		month: "numeric",
		day: "numeric",
		hour: "numeric",
		minute: "numeric",
		second: "numeric",
		hourCycle: "h23"
	}).formatToParts(time)) parts.set(p.type, p.value);
	const num = (type) => Number(parts.get(type) ?? "0");
	const pad = (n, width) => String(n).padStart(width, "0");
	const monthName = (width) => {
		return ctx.date(ctx.ctx.locale, {
			timeZone,
			day: "numeric",
			month: width
		}).formatToParts(time).find((p) => p.type === "month")?.value ?? "";
	};
	return tokens.map((token) => {
		if ("literal" in token) return token.literal;
		switch (token.field) {
			case "yyyy": return pad(num("year"), 4);
			case "yy": return pad(num("year") % 100, 2);
			case "MMMM": return monthName("long");
			case "MMM": return monthName("short");
			case "MM": return pad(num("month"), 2);
			case "M": return String(num("month"));
			case "dd": return pad(num("day"), 2);
			case "d": return String(num("day"));
			case "HH": return pad(num("hour"), 2);
			case "H": return String(num("hour"));
			case "mm": return pad(num("minute"), 2);
			default: return pad(num("second"), 2);
		}
	}).join("");
}
const FORMATTERS = /* @__PURE__ */ new Map([
	["link", {
		params: [{
			name: "text",
			type: "string",
			optional: true
		}],
		usage: "link or link(\"Our website\")",
		apply: (value, args) => {
			if (value === void 0 || value === null) return value;
			if (typeof value !== "string") throw new ValueError(`link needs an address as text, but the value is ${describe(value)}.`, "Pass an address such as \"https://example.com\" or \"mailto:office@example.com\".");
			const problem = checkLinkUrl(value);
			if (problem !== null) throw new ValueError(`"${value.length > 60 ? `${value.slice(0, 60)}…` : value}" cannot be a link: ${problem}`, "Use an absolute http, https or mailto address, such as \"https://example.com\".");
			return new LinkValue(value, typeof args[0] === "string" ? args[0] : value);
		}
	}],
	["upper", {
		params: [],
		usage: "upper",
		apply: text((s, ctx) => s.toLocaleUpperCase(ctx.ctx.locale), "upper")
	}],
	["lower", {
		params: [],
		usage: "lower",
		apply: text((s, ctx) => s.toLocaleLowerCase(ctx.ctx.locale), "lower")
	}],
	["capitalize", {
		params: [],
		usage: "capitalize",
		apply: text((s, ctx) => {
			const first = s.codePointAt(0);
			if (first === void 0) return s;
			const head = String.fromCodePoint(first);
			return head.toLocaleUpperCase(ctx.ctx.locale) + s.slice(head.length);
		}, "capitalize")
	}],
	["trim", {
		params: [],
		usage: "trim",
		apply: text((s) => s.trim(), "trim")
	}],
	["default", {
		params: [{
			name: "value",
			type: "any"
		}],
		usage: "default(\"-\")",
		apply: (value, args) => value === void 0 || value === null || value === "" ? args[0] : value
	}],
	["number", {
		params: [{
			name: "decimals",
			type: "integer",
			optional: true
		}],
		usage: "number or number(2)",
		check: checkDecimals(0),
		apply: (value, args, ctx) => value === void 0 || value === null ? value : ctx.number(decimals(args, 0)).format(finiteNumber(value, "number"))
	}],
	["percent", {
		params: [{
			name: "decimals",
			type: "integer",
			optional: true
		}],
		usage: "percent or percent(1)",
		check: checkDecimals(0),
		apply: (value, args, ctx) => value === void 0 || value === null ? value : ctx.number({
			style: "percent",
			...decimals(args, 0)
		}).format(finiteNumber(value, "percent"))
	}],
	["currency", {
		params: [{
			name: "code",
			type: "string"
		}, {
			name: "decimals",
			type: "integer",
			optional: true
		}],
		usage: "currency(\"PLN\") or currency(\"EUR\", 0)",
		check: (args) => /^[A-Za-z]{3}$/.test(String(args[0])) ? checkDecimals(1)(args) : "The currency must be a three-letter ISO 4217 code, such as \"PLN\", \"EUR\" or \"USD\".",
		apply: (value, args, ctx) => value === void 0 || value === null ? value : ctx.number({
			style: "currency",
			currency: String(args[0]).toUpperCase(),
			...decimals(args, 1)
		}).format(finiteNumber(value, "currency"))
	}],
	["date", {
		params: [{
			name: "format",
			type: "string",
			optional: true
		}],
		usage: "date, date(\"long\") or date(\"dd.MM.yyyy\")",
		check: (args) => {
			const format = args[0];
			if (typeof format !== "string" || DATE_STYLES.has(format)) return null;
			const tokens = parseDatePattern(format);
			return typeof tokens === "string" ? `${tokens} Use short, medium, long or full, or a pattern of yyyy, yy, MMMM, MMM, MM, M, dd, d, HH, H, mm and ss, with text in 'quotes'.` : null;
		},
		apply: (value, args, ctx) => {
			if (value === void 0 || value === null) return value;
			const { time, timeZone } = toInstant(value, ctx);
			const format = typeof args[0] === "string" ? args[0] : "medium";
			if (DATE_STYLES.has(format)) return ctx.date(ctx.ctx.locale, {
				dateStyle: format,
				timeZone
			}).format(time);
			return formatPattern(parseDatePattern(format), time, timeZone, ctx);
		}
	}]
]);
/** The names of the formatters, for messages. */
const FORMATTER_NAMES = [...FORMATTERS.keys()];
function typeOk(type, value) {
	if (type === "any") return true;
	if (type === "string") return typeof value === "string";
	return typeof value === "number" && Number.isInteger(value);
}
/** Checks a formatter call of a template: the name is allowed and the arguments fit. */
function checkFormatter(call) {
	const spec = FORMATTERS.get(call.name);
	if (spec === void 0) {
		const close = suggest(call.name, FORMATTER_NAMES);
		return {
			code: "TV_E_FORMATTER_UNKNOWN",
			message: `There is no formatter "${call.name}".`,
			hint: `${close.length > 0 ? `Did you mean ${close.map((c) => `"${c}"`).join(" or ")}? ` : ""}The formatters are ${FORMATTER_NAMES.join(", ")}.`
		};
	}
	const required = spec.params.filter((p) => p.optional !== true).length;
	const bad = (message) => ({
		code: "TV_E_FORMATTER_ARGS",
		message,
		hint: `Write it as ${spec.usage}.`
	});
	if (call.args.length < required || call.args.length > spec.params.length) return bad(`${call.name} takes ${required === spec.params.length ? String(required) : `${String(required)} to ${String(spec.params.length)}`} argument${spec.params.length === 1 ? "" : "s"}, not ${String(call.args.length)}.`);
	for (let i = 0; i < call.args.length; i++) {
		const param = spec.params[i];
		if (!typeOk(param.type, call.args[i])) return bad(`The ${param.name} of ${call.name} must be ${param.type === "string" ? "a string in quotes" : "a whole number"}.`);
	}
	const message = spec.check?.(call.args) ?? null;
	return message === null ? null : bad(message);
}
/** Applies a formatter that {@link checkFormatter} accepted. Throws {@link ValueError} for a value it cannot format. */
function applyFormatter(call, value, ctx) {
	return FORMATTERS.get(call.name).apply(value, call.args, ctx);
}
// src/template/lexer.ts
/** `{` and `}`. */
const DEFAULT_DELIMITERS = {
	start: "{",
	end: "}"
};
/** Why two strings cannot be delimiters, or null when they can. */
function delimiterProblem(d) {
	if (d.start === "" || d.end === "") return "a delimiter is empty";
	if (/\s/.test(d.start + d.end)) return "a delimiter contains a space or a line break";
	if (d.start.length > 10 || d.end.length > 10) return "a delimiter is longer than 10 characters";
	if (d.start.includes(d.end) || d.end.includes(d.start)) return "one delimiter contains the other";
	return null;
}
/** The delimiters a tag such as `=[[ ]]=` sets, a problem, or null when it is not such a tag (D-056). */
function delimiterChange(inner) {
	if (inner.length < 3 || !inner.startsWith("=") || !inner.endsWith("=")) return null;
	const parts = inner.slice(1, -1).trim().split(/\s+/);
	if (parts.length !== 2) return "write the new start and end delimiters separated by a space, for example {=[[ ]]=}";
	const next = {
		start: parts[0],
		end: parts[1]
	};
	return delimiterProblem(next) ?? next;
}
/**
* Finds the tags of `text` between `delimiters`. A tag longer than `maxTagLength` (delimiters
* included) is reported as a `TV_E_LIMIT_EXCEEDED` problem and not parsed. A tag `=start end=`
* changes the delimiters for the rest of the text; the delimiters in force at the end are
* returned, for the next paragraph. Stops collecting problems after `maxProblems`.
*/
function lexText(text, maxTagLength, maxProblems = Number.POSITIVE_INFINITY, delimiters = DEFAULT_DELIMITERS) {
	const tags = [];
	const problems = [];
	const problem = (p) => {
		if (problems.length < maxProblems) problems.push(p);
	};
	let { start: openText, end: closeText } = delimiters;
	let cachedOpen = -2;
	let cachedClose = -2;
	const next = (which, from) => {
		const cached = which === "open" ? cachedOpen : cachedClose;
		if (cached !== -2 && (cached === -1 || cached >= from)) return cached;
		const found = text.indexOf(which === "open" ? openText : closeText, from);
		if (which === "open") cachedOpen = found;
		else cachedClose = found;
		return found;
	};
	let at = 0;
	while (at < text.length) {
		const open = next("open", at);
		const close = next("close", at);
		if (close !== -1 && (open === -1 || close < open)) {
			problem({
				code: "TV_E_TAG_UNOPENED",
				start: close,
				raw: closeText
			});
			at = close + closeText.length;
			continue;
		}
		if (open === -1) break;
		const inside = open + openText.length;
		const closeAfter = next("close", inside);
		const nextOpen = next("open", inside);
		if (closeAfter === -1 || nextOpen !== -1 && nextOpen < closeAfter) {
			const end = nextOpen === -1 ? text.length : nextOpen;
			problem({
				code: "TV_E_TAG_UNCLOSED",
				start: open,
				raw: clip(text.slice(open, end))
			});
			if (nextOpen === -1) break;
			at = nextOpen;
			continue;
		}
		const end = closeAfter + closeText.length;
		const raw = text.slice(open, end);
		const inner = text.slice(inside, closeAfter);
		at = end;
		if (raw.length > maxTagLength) {
			problem({
				code: "TV_E_LIMIT_EXCEEDED",
				start: open,
				raw: clip(raw)
			});
			continue;
		}
		const change = delimiterChange(inner);
		if (typeof change === "string") {
			problem({
				code: "TV_E_TAG_SYNTAX",
				start: open,
				raw: clip(raw)
			});
			continue;
		}
		if (change !== null) {
			tags.push({
				start: open,
				end,
				raw,
				inner,
				changes: change
			});
			openText = change.start;
			closeText = change.end;
			cachedOpen = -2;
			cachedClose = -2;
			continue;
		}
		tags.push({
			start: open,
			end,
			raw,
			inner
		});
	}
	return {
		tags,
		problems,
		delimiters: {
			start: openText,
			end: closeText
		}
	};
}
// src/template/structure.ts
const sectionText = (tag) => `{${tag.info.kind === "inverted" ? "^" : "#"}${tag.info.path}}`;
/**
* Builds the tree of one part. Problems (unclosed, unopened and crossed sections, and `{.}` outside
* a section) are returned, not thrown; nesting deeper than `maxSectionDepth` throws
* `TV_E_LIMIT_EXCEEDED`.
*/
function buildTree(tags, limits) {
	const root = [];
	const stack = [];
	const problems = [];
	const children = () => stack[stack.length - 1]?.children ?? root;
	for (const record of tags) {
		const tag = record.info;
		switch (tag.kind) {
			case "section":
			case "inverted":
				if (stack.length >= limits.maxSectionDepth) throw limitError("maxSectionDepth", limits.maxSectionDepth, `Sections are nested more than ${String(limits.maxSectionDepth)} levels deep.`, tag.location);
				stack.push({
					open: record,
					children: []
				});
				break;
			case "close": {
				let at = stack.length - 1;
				if (tag.path !== "") while (at >= 0 && stack[at].open.info.path !== tag.path) at--;
				if (at < 0) {
					problems.push(new TvError("TV_E_SECTION_UNOPENED", `"${tag.raw}" closes a section that is not open.`, {
						location: tag.location,
						hint: `${tag.path === "" ? "Add an opening tag such as {#items}" : `Add {#${tag.path}}`} before it, or delete it. A section must open and close in the same part of the document (the body, a header, a footer or the notes).`,
						details: { tag: tag.raw }
					}));
					break;
				}
				if (at !== stack.length - 1) {
					const inner = stack[stack.length - 1];
					problems.push(new TvError("TV_E_SECTION_MISMATCH", `"{/${tag.path}}" comes before the end of "${sectionText(inner.open)}", which is inside it.`, {
						location: tag.location,
						hint: `Close the inner section first: put {/${inner.open.info.path}} before {/${tag.path}}.`,
						details: {
							tag: tag.raw,
							open: inner.open.info.raw
						}
					}));
				}
				const frame = stack[at];
				stack.length = at;
				const hasContent = frame.children.length > 0;
				const plain = frame.open.parsed.segments !== null && frame.open.parsed.formatters.length === 0;
				const kind = frame.open.info.kind === "inverted" ? "inverted" : hasContent && plain ? "loop" : "condition";
				children().push({ section: {
					kind,
					open: frame.open,
					close: record,
					children: frame.children
				} });
				break;
			}
			default:
				if (record.parsed.paths.includes(".") && stack.length === 0) problems.push(new TvError("TV_E_TAG_SYNTAX", `"${tag.raw}" stands for the current item of a loop, but it is not inside a section.`, {
					location: tag.location,
					hint: "Use {.} only between {#list} and {/list}, or write the name of the field, for example {name}.",
					details: { tag: tag.raw }
				}));
				children().push({ tag: record });
		}
	}
	for (const frame of stack) problems.push(new TvError("TV_E_SECTION_UNCLOSED", `The section "${sectionText(frame.open)}" is not closed.`, {
		location: frame.open.info.location,
		hint: `Add {/${frame.open.info.path}} after the content of the section, in the same part of the document (the body, a header, a footer or the notes).`,
		details: { tag: frame.open.info.raw }
	}));
	return {
		children: root,
		problems
	};
}
/** The public form of a section. */
function toSectionNode(tree) {
	const tags = [];
	const sections = [];
	for (const child of tree.children) if ("tag" in child) tags.push(child.tag.info);
	else sections.push(toSectionNode(child.section));
	return {
		kind: tree.kind,
		path: tree.open.info.path,
		formatters: tree.open.info.formatters,
		open: tree.open.info,
		close: tree.close.info,
		tags,
		sections
	};
}
// src/template/tag.ts
/** Names that would reach object prototypes in JavaScript; never valid as a path segment or formatter. */
const FORBIDDEN_NAMES = /* @__PURE__ */ new Set([
	"__proto__",
	"constructor",
	"prototype"
]);
const SIGILS = /* @__PURE__ */ new Map([
	["#", "section"],
	["^", "inverted"],
	["/", "close"],
	["%", "image"],
	["~", "markdown"]
]);
/** Characters of operators that the grammar does not have (arithmetic, conditional, indexing, calls). */
const UNSUPPORTED_CHARS = "+-*/%?:[]&=";
/** Quote characters by family: a string opened by one closes at the next of the same family. */
const DOUBLE_QUOTES = "\"„“”";
const SINGLE_QUOTES = "'‘’‚";
const ID_START = /^[\p{ID_Start}_]$/u;
const ID_CONTINUE = /^[\p{ID_Continue}\u200C\u200D]$/u;
var Problem = class extends Error {
	problem;
	constructor(problem) {
		super(problem.message);
		this.problem = problem;
	}
};
const SYNTAX_HINT = "Write a value as {name} or {customer.name}, a formatter as {total | currency(\"PLN\")}; see the syntax reference.";
const OPERATOR_HINT = "Tags support the operators == != < <= > >= && || ! and parentheses; arithmetic, function calls, [ ] and ?: are not supported. Names may contain letters, digits and _, but not -.";
var Scanner = class {
	s;
	pos = 0;
	constructor(s) {
		this.s = s;
	}
	get done() {
		return this.pos >= this.s.length;
	}
	/** The code point at the position, as a string, or '' at the end. */
	peek() {
		const cp = this.s.codePointAt(this.pos);
		return cp === void 0 ? "" : String.fromCodePoint(cp);
	}
	next() {
		const c = this.peek();
		this.pos += c.length;
		return c;
	}
	skipWs() {
		while (this.s[this.pos] === " " || this.s[this.pos] === "\xA0") this.pos++;
	}
	fail(message, hint = SYNTAX_HINT, at = this.pos) {
		throw new Problem({
			code: "TV_E_TAG_SYNTAX",
			message,
			hint,
			position: at
		});
	}
	/** Fails at an unexpected character, with the operator hint when it looks like an operator. */
	unexpected(expected) {
		const c = this.peek();
		if (c === "") this.fail(`The tag ends where ${expected} was expected.`);
		if (c === "	" || c === "\n" || c === "\r") this.fail("The tag contains a tab or a line break.", "Keep the whole tag on one line, without tabs; type it again if the document split it.");
		const isOperator = UNSUPPORTED_CHARS.includes(c) || c === "(" || c === "<" || c === ">" || c === "!" || c === "|";
		this.fail(`Unexpected "${c}" where ${expected} was expected.`, isOperator ? OPERATOR_HINT : SYNTAX_HINT);
	}
};
function isIdStart(c) {
	return c !== "" && ID_START.test(c);
}
function isIdContinue(c) {
	return c !== "" && ID_CONTINUE.test(c);
}
function readName(sc, what) {
	const start = sc.pos;
	if (!isIdStart(sc.peek())) sc.unexpected(what);
	sc.next();
	while (isIdContinue(sc.peek())) sc.next();
	const name = sc.s.slice(start, sc.pos);
	if (FORBIDDEN_NAMES.has(name)) throw new Problem({
		code: "TV_E_TAG_FORBIDDEN_NAME",
		message: `"${name}" cannot be used as a name in a tag.`,
		hint: "The names __proto__, constructor and prototype are reserved because they reach the internals of JavaScript objects. Rename the field in the template and in the data.",
		position: start
	});
	return name;
}
function readPath(sc, allowCurrent) {
	if (sc.peek() === ".") {
		if (!allowCurrent) sc.fail("\".\" (the current item) cannot be used here.", "Close a section with its name, for example {/items}.");
		sc.next();
		return [];
	}
	const segments = [readName(sc, "a name")];
	for (;;) {
		const save = sc.pos;
		sc.skipWs();
		if (sc.peek() !== ".") {
			sc.pos = save;
			return segments;
		}
		sc.next();
		sc.skipWs();
		segments.push(readName(sc, "a name after \".\""));
	}
}
function readString(sc) {
	const start = sc.pos;
	const open = sc.next();
	const family = DOUBLE_QUOTES.includes(open) ? DOUBLE_QUOTES : SINGLE_QUOTES;
	let out = "";
	for (;;) {
		if (sc.done) sc.fail("A string in the tag is not closed.", "End the string with the same kind of quote that starts it, for example \"PLN\".", start);
		const c = sc.next();
		if (family.includes(c)) return out;
		if (c === "\\") {
			if (sc.done) sc.fail("A string in the tag is not closed.", "End the string with the same kind of quote that starts it, for example \"PLN\".", start);
			out += sc.next();
			continue;
		}
		out += c;
	}
}
function readNumber(sc) {
	const start = sc.pos;
	const digits = () => {
		const from = sc.pos;
		while (sc.peek() >= "0" && sc.peek() <= "9") sc.next();
		return sc.pos - from;
	};
	if (sc.peek() === "-") sc.next();
	if (sc.peek() === "0") sc.next();
	else if (digits() === 0) sc.unexpected("a number");
	if (sc.peek() === ".") {
		sc.next();
		if (digits() === 0) sc.unexpected("a digit");
	}
	if (sc.peek() === "e" || sc.peek() === "E") {
		sc.next();
		if (sc.peek() === "+" || sc.peek() === "-") sc.next();
		if (digits() === 0) sc.unexpected("a digit");
	}
	const value = Number(sc.s.slice(start, sc.pos));
	if (!Number.isFinite(value)) sc.fail("A number in the tag is too large.", "Use a number between -1e308 and 1e308.", start);
	return value;
}
function readLiteral(sc) {
	const c = sc.peek();
	if (DOUBLE_QUOTES.includes(c) || SINGLE_QUOTES.includes(c)) return readString(sc);
	if (c === "-" || c >= "0" && c <= "9") return readNumber(sc);
	if (isIdStart(c)) {
		const start = sc.pos;
		while (isIdContinue(sc.peek())) sc.next();
		const word = sc.s.slice(start, sc.pos);
		if (word === "true") return true;
		if (word === "false") return false;
		if (word === "null") return null;
		sc.fail(`"${word}" is not a value an argument can have.`, "Arguments of formatters are strings in quotes, numbers, true, false or null, for example currency(\"PLN\").", start);
	}
	return sc.unexpected("a string, a number, true, false or null");
}
function readFormatter(sc) {
	const name = readName(sc, "the name of a formatter after \"|\"");
	const save = sc.pos;
	sc.skipWs();
	if (sc.peek() !== "(") {
		sc.pos = save;
		return {
			name,
			args: []
		};
	}
	sc.next();
	sc.skipWs();
	const args = [];
	if (sc.peek() === ")") {
		sc.next();
		return {
			name,
			args
		};
	}
	for (;;) {
		args.push(readLiteral(sc));
		sc.skipWs();
		const c = sc.peek();
		if (c === ")") {
			sc.next();
			return {
				name,
				args
			};
		}
		if (c !== ",") sc.unexpected("\",\" or \")\"");
		sc.next();
		sc.skipWs();
	}
}
const COMPARE_OPS = [
	"==",
	"!=",
	"<=",
	">=",
	"<",
	">"
];
var ExprParser = class {
	sc;
	maxDepth;
	depth = 0;
	/** "&&" and "||" build a tree as deep as the chain is long, so they count too (D-071). */
	operators = 0;
	constructor(sc, maxDepth) {
		this.sc = sc;
		this.maxDepth = maxDepth;
	}
	enter() {
		this.depth++;
		if (this.depth > this.maxDepth || this.depth > 256) throw new Problem({
			code: "TV_E_LIMIT_EXCEEDED",
			message: `The expression is nested more than ${String(Math.min(this.maxDepth, 256))} levels deep.`,
			hint: "Use fewer parentheses and \"!\". If you trust this template, raise limits.maxExpressionDepth.",
			position: this.sc.pos
		});
	}
	chain() {
		if (++this.operators > 256) throw new Problem({
			code: "TV_E_LIMIT_EXCEEDED",
			message: `The expression has more than ${String(256)} "&&" and "||".`,
			hint: "Split the condition into nested sections, or compute the value in the data.",
			position: this.sc.pos
		});
	}
	/** Tries `token` after optional spaces; leaves the position unchanged when it is not there. */
	take(token, notFollowedBy = "") {
		const save = this.sc.pos;
		this.sc.skipWs();
		const s = this.sc.s;
		if (s.startsWith(token, this.sc.pos) && (notFollowedBy === "" || !notFollowedBy.includes(s[this.sc.pos + token.length] ?? "\0"))) {
			this.sc.pos += token.length;
			this.sc.skipWs();
			return true;
		}
		this.sc.pos = save;
		return false;
	}
	pipe() {
		const e = this.logical();
		const formatters = [];
		while (this.take("|", "|")) formatters.push(readFormatter(this.sc));
		return formatters.length === 0 ? e : {
			t: "pipe",
			e,
			formatters
		};
	}
	logical() {
		let l = this.and();
		while (this.take("||")) {
			this.chain();
			l = {
				t: "or",
				l,
				r: this.and()
			};
		}
		return l;
	}
	and() {
		let l = this.compare();
		while (this.take("&&")) {
			this.chain();
			l = {
				t: "and",
				l,
				r: this.compare()
			};
		}
		return l;
	}
	operator() {
		const save = this.sc.pos;
		this.sc.skipWs();
		const s = this.sc.s;
		for (const op of COMPARE_OPS) if (s.startsWith(op, this.sc.pos)) {
			this.sc.pos += op.length;
			this.sc.skipWs();
			return op;
		}
		if (s[this.sc.pos] === "=") this.sc.fail("\"=\" is not an operator.", "Compare with \"==\", for example {#status == \"paid\"}.");
		this.sc.pos = save;
		return null;
	}
	compare() {
		const l = this.unary();
		const op = this.operator();
		if (op === null) return l;
		const r = this.unary();
		const at = this.sc.pos;
		if (this.operator() !== null) this.sc.fail("Comparisons cannot be chained.", "Join two comparisons with &&, for example {#a > 1 && a < 9}.", at);
		return {
			t: "compare",
			op,
			l,
			r
		};
	}
	unary() {
		const sc = this.sc;
		if (sc.peek() === "!" && sc.s[sc.pos + 1] !== "=") {
			sc.next();
			sc.skipWs();
			this.enter();
			const e = this.unary();
			this.depth--;
			return {
				t: "not",
				e
			};
		}
		return this.operand();
	}
	operand() {
		const sc = this.sc;
		const c = sc.peek();
		if (c === "(") {
			sc.next();
			sc.skipWs();
			this.enter();
			const e = this.pipe();
			sc.skipWs();
			if (sc.peek() !== ")") sc.unexpected("\")\"");
			sc.next();
			this.depth--;
			return e;
		}
		if (DOUBLE_QUOTES.includes(c) || SINGLE_QUOTES.includes(c) || c === "-" || c >= "0" && c <= "9") return {
			t: "literal",
			value: readLiteral(sc)
		};
		if (isIdStart(c)) {
			const save = sc.pos;
			while (isIdContinue(sc.peek())) sc.next();
			const word = sc.s.slice(save, sc.pos);
			if (word === "true" || word === "false" || word === "null") return {
				t: "literal",
				value: word === "null" ? null : word === "true"
			};
			sc.pos = save;
		}
		return {
			t: "path",
			segments: readPath(sc, true)
		};
	}
};
/**
* Parses the inside of a tag. Never throws: a tag that breaks the grammar, or nests deeper than
* `maxDepth`, gives a {@link TagProblem}.
*/
function parseTag(inner, maxDepth = 32) {
	const sc = new Scanner(inner);
	try {
		sc.skipWs();
		if (sc.peek() === "@") sc.fail("Raw XML tags such as {@rawXml} are not supported: tplvane escapes every value, so data cannot change the structure of a document.", "Use a value tag; for formatted text, use Markdown with {~name}.");
		if (sc.peek() === "-" && /^-[\w]+:[\w]+\s/.test(inner.trimStart())) sc.fail("Loops over a named element such as {-w:p loop} are not supported.", "Use {#loop} … {/loop}: a loop whose tags are alone in their paragraphs repeats the paragraphs, and one in a table row repeats the row.");
		const kind = SIGILS.get(sc.peek()) ?? "value";
		if (kind !== "value") {
			sc.next();
			sc.skipWs();
		}
		if (sc.done && kind === "close") return {
			kind,
			expr: {
				t: "literal",
				value: null
			},
			path: "",
			segments: null,
			formatters: [],
			paths: []
		};
		if (sc.done) return {
			code: "TV_E_TAG_EMPTY",
			message: kind === "value" ? "The tag is empty." : "The tag has no name.",
			hint: "Write a name between the braces, for example {customer.name}, or delete the braces if they are meant as text.",
			position: sc.pos
		};
		const parser = new ExprParser(sc, maxDepth);
		let full;
		if (kind === "image" || kind === "markdown") full = {
			t: "path",
			segments: readPath(sc, true)
		};
		else if (kind === "close") {
			if (sc.peek() === ".") readPath(sc, false);
			full = parser.logical();
		} else full = parser.pipe();
		sc.skipWs();
		if (!sc.done) {
			if (kind !== "value" && kind !== "section" && kind !== "inverted" && sc.peek() === "|" && inner[sc.pos + 1] !== "|") sc.fail(kind === "close" ? "A closing tag cannot have formatters." : "Image and Markdown tags cannot have formatters.", kind === "close" ? "Close the section with its name only, for example {/items}." : "Write only the name, for example {%logo} or {~summary}.");
			sc.unexpected(full.t === "path" && full.segments.length > 0 ? "\".\", an operator, \"|\" or the end of the tag" : "an operator, \"|\" or the end of the tag");
		}
		const expr = full.t === "pipe" ? full.e : full;
		const formatters = full.t === "pipe" ? full.formatters : [];
		return {
			kind,
			expr,
			path: canonical(expr),
			segments: expr.t === "path" ? expr.segments : null,
			formatters,
			paths: pathsOf$1(full)
		};
	} catch (error) {
		if (error instanceof Problem) return error.problem;
		throw error;
	}
}
// src/xlsx/ranges.ts
const isDigit = (c) => c >= 48 && c <= 57;
const isLetter = (c) => c >= 65 && c <= 90 || c >= 97 && c <= 122;
/** Characters of an unquoted name, cell reference or function name. */
const isWordChar = (c) => isLetter(c) || isDigit(c) || c === 95 || c === 46 || c === 36 || c === 92 || c === 63 || c > 127;
/** Reads `text[from, to)` as a cell reference, or returns null. */
function cellRef(text, from, to) {
	let i = from;
	if (text.charCodeAt(i) === 36) i++;
	const letters = i;
	while (i < to && isLetter(text.charCodeAt(i))) i++;
	if (i === letters || i - letters > 3) return null;
	if (text.charCodeAt(i) === 36) i++;
	const rowStart = i;
	while (i < to && isDigit(text.charCodeAt(i))) i++;
	if (i !== to || i === rowStart || text.charCodeAt(rowStart) === 48 || i - rowStart > 7) return null;
	return {
		rowStart,
		rowEnd: i,
		row: Number(text.slice(rowStart, i))
	};
}
/** The end of the word starting at `from`. */
function wordEnd(text, from) {
	let i = from;
	while (i < text.length && isWordChar(text.charCodeAt(i))) i++;
	return i;
}
/**
* Returns `formula` (without the leading "=") with every A1 area of `change.sheet` whose lower end
* is on `change.row` extended to `change.to`. Everything else is kept character for character.
*
* @example
* ```ts
* extendAreas('SUM(C5:C5)', { sheet: 'Data', formulaSheet: 'Data', row: 5, to: 9 }); // 'SUM(C5:C9)'
* ```
*/
function extendAreas(formula, change) {
	if (change.to === change.row) return formula;
	const target = change.sheet.toLowerCase();
	const own = change.formulaSheet.toLowerCase() === target;
	const replacement = String(change.to);
	const out = [];
	let copied = 0;
	let prefix = null;
	let i = 0;
	const n = formula.length;
	while (i < n) {
		const c = formula.charCodeAt(i);
		if (c === 34) {
			i++;
			while (i < n) if (formula.charCodeAt(i) === 34) {
				if (formula.charCodeAt(i + 1) === 34) i += 2;
				else break;
			} else i++;
			i++;
			prefix = null;
			continue;
		}
		if (c === 39) {
			let j = i + 1;
			const parts = [];
			let from = j;
			let closed = false;
			while (j < n) if (formula.charCodeAt(j) === 39) {
				parts.push(formula.slice(from, j));
				if (formula.charCodeAt(j + 1) === 39) {
					parts.push("'");
					j += 2;
					from = j;
				} else {
					closed = true;
					break;
				}
			} else j++;
			if (!closed) break;
			const name = parts.join("");
			if (formula.charCodeAt(j + 1) === 33) {
				prefix = name.toLowerCase() === target ? "target" : "other";
				i = j + 2;
			} else {
				prefix = null;
				i = j + 1;
			}
			continue;
		}
		if (c === 91) {
			let depth = 0;
			while (i < n) {
				const d = formula.charCodeAt(i);
				if (d === 91) depth++;
				else if (d === 93 && --depth === 0) break;
				i++;
			}
			i++;
			const end = wordEnd(formula, i);
			if (end > i && formula.charCodeAt(end) === 33) {
				prefix = "other";
				i = end + 1;
			} else prefix = null;
			continue;
		}
		if (isWordChar(c)) {
			const end = wordEnd(formula, i);
			const next = formula.charCodeAt(end);
			if (next === 33) {
				prefix = !(i > 0 && formula.charCodeAt(i - 1) === 58) && formula.slice(i, end).toLowerCase() === target ? "target" : "other";
				i = end + 1;
				continue;
			}
			const first = cellRef(formula, i, end);
			if (first !== null && next === 58) {
				const secondStart = end + 1;
				const secondEnd = wordEnd(formula, secondStart);
				const second = cellRef(formula, secondStart, secondEnd);
				const after = formula.charCodeAt(secondEnd);
				if (second !== null && after !== 33 && after !== 40) {
					const belongs = prefix === null ? own : prefix === "target";
					const lower = second.row >= first.row ? second : first;
					if (belongs && lower.row === change.row) {
						out.push(formula.slice(copied, lower.rowStart), replacement);
						copied = lower.rowEnd;
					}
					i = secondEnd;
					prefix = null;
					continue;
				}
			}
			i = end;
			prefix = null;
			continue;
		}
		i++;
		prefix = null;
	}
	if (copied === 0) return formula;
	out.push(formula.slice(copied));
	return out.join("");
}
// src/xlsx/errors.ts
const SAME = /* @__PURE__ */ new Map([
	["XLV_E_LIMIT_EXCEEDED", "TV_E_LIMIT_EXCEEDED"],
	["XLV_E_NOT_ZIP", "TV_E_NOT_ZIP"],
	["XLV_E_UNSUPPORTED_CONTAINER", "TV_E_UNSUPPORTED_CONTAINER"],
	["XLV_E_ZIP_CORRUPT", "TV_E_ZIP_CORRUPT"],
	["XLV_E_ZIP_UNSUPPORTED", "TV_E_ZIP_UNSUPPORTED"],
	["XLV_E_XML_MALFORMED", "TV_E_XML_MALFORMED"],
	["XLV_E_XML_DTD_FORBIDDEN", "TV_E_XML_DTD_FORBIDDEN"],
	["XLV_E_XML_ENTITY_FORBIDDEN", "TV_E_XML_ENTITY_FORBIDDEN"],
	["XLV_E_XML_ENCODING", "TV_E_XML_ENCODING"],
	["XLV_E_PART_MISSING", "TV_E_PART_MISSING"],
	["XLV_E_NOT_SUPPORTED", "TV_E_NOT_SUPPORTED"],
	["XLV_E_INVALID_VALUE", "TV_E_INVALID_VALUE"],
	["XLV_E_INTERNAL", "TV_E_INTERNAL"]
]);
/** The place of an xlvane error in tplvane's terms. */
function locationOf(error) {
	const { part, sheet, cell } = error.location;
	const where = sheet === void 0 ? {} : {
		story: "sheet",
		sheet,
		description: cell === void 0 ? `sheet ${sheet}` : `sheet ${sheet}, cell ${cell}`
	};
	return {
		...part === void 0 ? {} : { part },
		...where,
		...cell === void 0 ? {} : { cell }
	};
}
/**
* Runs `fn`, turning an XlvError into a TvError with the same message and hint. `context` says
* what tplvane was doing, for errors that the template author cannot read as they are.
*/
function viaXlvane(fn, context) {
	try {
		return fn();
	} catch (error) {
		if (!isXlvError(error)) throw error;
		const code = SAME.get(error.code) ?? "TV_E_PART_INVALID";
		const details = { xlvaneCode: error.code };
		const limit = error.details["limit"];
		if (typeof limit === "string") details["limit"] = limit;
		const value = error.details["value"];
		if (typeof value === "number") details["value"] = value;
		throw new TvError(code, `${context}: ${error.message}`, {
			location: locationOf(error),
			hint: error.hint,
			details,
			cause: error
		});
	}
}
// src/xlsx/analyze.ts
const NO_ELEMENT = makeElement({
	ns: "",
	local: "c",
	qname: "c"
}, "", "");
/** The location of a cell. */
function cellLocation(sheet, ref) {
	const location = {
		story: "sheet",
		sheet: sheet.name,
		cell: ref,
		description: `sheet ${sheet.name}, cell ${ref}`
	};
	return sheet.part === null ? location : {
		part: sheet.part,
		...location
	};
}
/** The read limits of tplvane in the shape xlvane takes them; xlvane's own cell limits keep their defaults. */
function xlvaneLimits(limits) {
	const { maxInputBytes, maxEntries, maxEntryBytes, maxTotalBytes, maxCompressionRatio, compressionRatioMinBytes, maxXmlDepth, maxXmlAttributes, maxXmlNameLength, maxXmlTextLength, maxXmlTreeElements } = limits;
	return {
		maxInputBytes,
		maxEntries,
		maxEntryBytes,
		maxTotalBytes,
		maxCompressionRatio,
		compressionRatioMinBytes,
		maxXmlDepth,
		maxXmlAttributes,
		maxXmlNameLength,
		maxXmlTextLength,
		maxXmlTreeElements
	};
}
function placement$1(tag, message, hint) {
	return new TvError("TV_E_SECTION_PLACEMENT", message, {
		location: tag.info.location,
		hint,
		details: { tag: clip(tag.info.raw) }
	});
}
/** Tags that need what an Excel template cannot do through xlvane yet (D-043). */
function unsupported(tag) {
	const raw = clip(tag.info.raw);
	if (tag.info.kind === "image" || tag.info.kind === "markdown") return new TvError("TV_E_NOT_SUPPORTED", `${tag.info.kind === "image" ? "Image" : "Markdown"} tags such as "${raw}" are not supported in Excel templates yet.`, {
		location: tag.info.location,
		hint: tag.info.kind === "image" ? "Remove the tag, or use a Word or PowerPoint template for images." : `Write the text with a value tag instead, for example {${tag.info.path}}.`,
		details: { tag: raw }
	});
	if ([...formattersOf(tag.parsed.expr), ...tag.parsed.formatters].some((f) => f.name === "link")) return new TvError("TV_E_NOT_SUPPORTED", `The link formatter in "${raw}" is not supported in Excel templates yet.`, {
		location: tag.info.location,
		hint: "Write the address as text without \"| link\", or add the hyperlink to the cell in Excel.",
		details: { tag: raw }
	});
	return null;
}
/** Every tag of a part of the tree, sections included. */
function collect(children, out) {
	for (const child of children) if ("tag" in child) out.add(child.tag);
	else {
		out.add(child.section.open);
		out.add(child.section.close);
		collect(child.section.children, out);
	}
}
/** The section whose rows include `row`; `sections` are sorted and do not overlap. */
function sectionAt(sections, row) {
	let lo = 0;
	let hi = sections.length - 1;
	while (lo <= hi) {
		const mid = lo + hi >> 1;
		const s = sections[mid];
		if (row < s.open) hi = mid - 1;
		else if (row > s.close) lo = mid + 1;
		else return s;
	}
	return null;
}
/** The first tag of a tree child. */
const firstTag = (child) => "tag" in child ? child.tag : child.section.open;
/**
* Reads the sheets of a workbook. `scan` lexes and parses paragraphs into tag records (the same
* code as for documents) and `report` collects problems.
*/
function analyzeWorkbook(bytes, limits, scan, report, delimiters = DEFAULT_DELIMITERS) {
	const workbook = viaXlvane(() => readWorkbook(bytes, { limits: xlvaneLimits(limits) }), "Reading the workbook failed");
	const cellOf = /* @__PURE__ */ new Map();
	const sheets = [];
	for (const sheet of workbook.sheets) {
		if (sheet.kind !== "worksheet") continue;
		const strings = [];
		for (const cell of sheet.cells()) if (cell.formula === null && cell.type === "string" && typeof cell.value === "string") strings.push(cell);
		const candidate = (cell) => cell.value.includes(delimiters.start) || cell.value.includes(delimiters.end) || cell.value.includes("{");
		const cells = strings.some((cell) => candidate(cell) && /=\S+\s+\S+=/.test(cell.value)) ? strings : strings.filter(candidate);
		cells.sort((a, b) => a.row - b.row || a.col - b.col);
		const paragraphs = [];
		for (const cell of cells) {
			const paragraph = {
				element: NO_ELEMENT,
				text: cell.value,
				pieces: [],
				location: cellLocation(sheet, cell.ref),
				alternates: []
			};
			paragraphs.push(paragraph);
			cellOf.set(paragraph, {
				sheet,
				cell,
				paragraph
			});
		}
		const tags = [];
		for (const tag of scan(paragraphs)) {
			const problem = unsupported(tag);
			if (problem === null) tags.push(tag);
			else report(problem);
		}
		const tree = buildTree(tags, limits);
		for (const problem of tree.problems) report(problem);
		const rowSections = classify(sheet, tree.children, cellOf, report);
		sheets.push({
			sheet,
			tags,
			tree: tree.children,
			rowSections
		});
	}
	checkArrayFormulas(workbook, sheets, report);
	return {
		workbook,
		sheets,
		cellOf
	};
}
/**
* Array formulas and data tables cannot be copied into repeated rows, and xlvane writes no array
* formula, so one that the tplvane rule would extend (D-041) cannot be either. Both are template
* problems, reported by inspect and validate as well as render (D-071).
*/
function checkArrayFormulas(workbook, sheets, report) {
	for (const { sheet, rowSections } of sheets) for (const rows of rowSections) {
		if (rows.last < rows.first) continue;
		for (const other of workbook.sheets) {
			if (other.kind !== "worksheet") continue;
			for (const cell of other.cells()) {
				const { formula } = cell;
				if (formula === null || formula.kind !== "array" && formula.kind !== "dataTable") continue;
				const inside = other === sheet && cell.row >= rows.first && cell.row <= rows.last;
				const extended = !inside && formula.kind === "array" && formula.text !== null && extendAreas(formula.text, {
					sheet: sheet.name,
					formulaSheet: other.name,
					row: rows.last,
					to: rows.last + 1
				}) !== formula.text;
				if (!inside && !extended) continue;
				report(new TvError("TV_E_NOT_SUPPORTED", inside ? `The ${formula.kind === "array" ? "array formula" : "data table"} in ${other.name}!${cell.ref} is in rows that repeat.` : `The array formula in ${other.name}!${cell.ref} refers to the rows of ${clip(rows.section.open.info.raw)}, which grow, and array formulas cannot be extended yet.`, {
					location: cellLocation(other, cell.ref),
					hint: inside ? "Move it out of the rows that repeat, or write it as an ordinary formula." : "Write it as an ordinary formula, such as SUMPRODUCT instead of an array SUM, so that it grows with the rows.",
					details: { kind: formula.kind }
				}));
			}
		}
	}
}
/** Finds the sections over rows of one sheet and checks where they are. */
function classify(sheet, tree, cellOf, report) {
	const cellOfTag = (tag) => cellOf.get(tag.paragraph).cell;
	const candidates = [];
	const nested = (children) => {
		for (const child of children) {
			if ("tag" in child) continue;
			const { section } = child;
			if (cellOfTag(section.open) !== cellOfTag(section.close)) report(placement$1(section.open, `The section "${clip(section.open.info.raw)}" repeats rows, but it is inside another section that repeats rows.`, "Rows repeat on one level in this version. Keep the inner section within one cell, or move it out of the outer section."));
			else nested(section.children);
		}
	};
	for (const child of tree) {
		if ("tag" in child) continue;
		const { section } = child;
		const open = cellOfTag(section.open);
		const close = cellOfTag(section.close);
		if (open === close) continue;
		candidates.push({
			section,
			open,
			close
		});
		nested(section.children);
	}
	if (candidates.length === 0) return [];
	const wanted = /* @__PURE__ */ new Set();
	for (const c of candidates) {
		wanted.add(c.open.row);
		wanted.add(c.close.row);
	}
	const filled = /* @__PURE__ */ new Map();
	for (const cell of sheet.cells()) if (wanted.has(cell.row) && (cell.type !== "empty" || cell.formula !== null)) filled.set(cell.row, (filled.get(cell.row) ?? 0) + 1);
	const alone = (tag, cell) => cell.value.trim() === tag.info.raw && filled.get(cell.row) === 1;
	const sections = [];
	const inside = /* @__PURE__ */ new Set();
	for (const c of candidates) {
		const markers = c.open.row < c.close.row && alone(c.section.open, c.open) && alone(c.section.close, c.close);
		const previous = sections[sections.length - 1];
		if (previous !== void 0 && c.open.row <= previous.close) {
			report(placement$1(c.section.open, `The section "${clip(c.section.open.info.raw)}" starts in row ${String(c.open.row)}, which the section before it repeats.`, "Start each section that repeats rows in a row below the end of the one before it."));
			collect([{ section: c.section }], inside);
			continue;
		}
		sections.push({
			section: c.section,
			open: c.open.row,
			close: c.close.row,
			markers,
			first: markers ? c.open.row + 1 : c.open.row,
			last: markers ? c.close.row - 1 : c.close.row
		});
	}
	for (const s of sections) collect([{ section: s.section }], inside);
	for (const child of tree) {
		const tag = firstTag(child);
		if (inside.has(tag)) continue;
		const row = cellOfTag(tag).row;
		const around = sectionAt(sections, row);
		if (around === null) continue;
		report(placement$1(tag, `The tag "${clip(tag.info.raw)}" is in row ${String(row)}, which the section "${clip(around.section.open.info.raw)}" repeats, but it is outside the section.`, `Move ${clip(around.section.open.info.raw)} before this tag, or move the tag to another row.`));
	}
	return sections;
}
// src/template/analyze.ts
function lexProblem(problem, location, limits) {
	switch (problem.code) {
		case "TV_E_TAG_UNCLOSED": return new TvError("TV_E_TAG_UNCLOSED", `The tag "${problem.raw}" is not closed.`, {
			location,
			hint: "End the tag with \"}\". A tag must start and end in the same paragraph and cannot contain another \"{\"; braces are reserved for tags.",
			details: { tag: problem.raw }
		});
		case "TV_E_TAG_UNOPENED": return new TvError("TV_E_TAG_UNOPENED", "The text has a \"}\" without a \"{\" before it.", {
			location,
			hint: "Delete the \"}\" or add the missing \"{\"; braces are reserved for tags. A tag must start and end in the same paragraph.",
			details: { tag: problem.raw }
		});
		case "TV_E_LIMIT_EXCEEDED": return limitError("maxTagLength", limits.maxTagLength, `The tag "${clip(problem.raw.slice(0, 40))}" is longer than ${String(limits.maxTagLength)} characters.`, location);
		case "TV_E_TAG_SYNTAX": return new TvError("TV_E_TAG_SYNTAX", `The tag "${problem.raw}" does not set valid delimiters.`, {
			location,
			hint: "A tag between equals signs changes the delimiters: write the new start and end separated by a space, for example {=[[ ]]=}. They must not be empty, contain spaces or contain each other.",
			details: { tag: problem.raw }
		});
	}
}
/**
* Reads the template. Throws `TV_E_TEMPLATE_INVALID` with every problem of its tags and sections
* (up to `limits.maxErrors`), and the package, ZIP, XML and limit errors of the layers below.
*/
function analyzeTemplate(bytes, limits, delimiters = DEFAULT_DELIMITERS) {
	const docx = openDocx(bytes, limits);
	const problems = [];
	const full = () => problems.length >= limits.maxErrors;
	const report = (error) => {
		if (!full()) problems.push(error);
	};
	let tagCount = 0;
	/** Lexes and parses the tags of paragraphs, in paragraph order. */
	const scan = (paragraphs) => {
		const records = [];
		let current = delimiters;
		for (const paragraph of paragraphs) {
			if (full()) break;
			const lexed = lexText(paragraph.text, limits.maxTagLength, limits.maxErrors - problems.length, current);
			current = lexed.delimiters;
			const at = (offset) => ({
				...paragraph.location,
				character: offset + 1
			});
			for (const problem of lexed.problems) report(lexProblem(problem, at(problem.start), limits));
			for (const raw of lexed.tags) {
				if (++tagCount > limits.maxTags) throw limitError("maxTags", limits.maxTags, `The template has more than ${String(limits.maxTags)} tags.`, at(raw.start));
				const parsed = parseTag(raw.changes === void 0 ? raw.inner : "\"\"", limits.maxExpressionDepth);
				if ("code" in parsed) {
					const details = {
						tag: clip(raw.raw),
						position: parsed.position + 1 + Math.max(0, raw.raw.indexOf(raw.inner))
					};
					report(parsed.code === "TV_E_LIMIT_EXCEEDED" ? limitError("maxExpressionDepth", limits.maxExpressionDepth, `${parsed.message} (in "${clip(raw.raw)}")`, at(raw.start)) : new TvError(parsed.code, `${parsed.message} (in "${clip(raw.raw)}")`, {
						location: at(raw.start),
						hint: parsed.hint,
						details
					}));
					continue;
				}
				let formatterOk = true;
				for (const call of [...formattersOf(parsed.expr), ...parsed.formatters]) {
					const problem = checkFormatter(call);
					if (problem === null) continue;
					formatterOk = false;
					report(new TvError(problem.code, `${problem.message} (in "${clip(raw.raw)}")`, {
						location: at(raw.start),
						hint: problem.hint,
						details: {
							tag: clip(raw.raw),
							formatter: call.name
						}
					}));
				}
				if (!formatterOk) continue;
				const info = {
					kind: parsed.kind,
					raw: raw.raw,
					path: parsed.path,
					paths: parsed.paths,
					formatters: parsed.formatters,
					location: at(raw.start)
				};
				records.push({
					info,
					parsed,
					paragraph,
					start: raw.start,
					end: raw.end
				});
			}
		}
		return records;
	};
	const stories = [];
	const workbook = docx.format === "xlsx" ? analyzeWorkbook(bytes, limits, scan, report, delimiters) : null;
	for (const story of docx.stories) {
		if (full()) break;
		const xml = docx.pkg.readXml(story.part);
		if (xml === void 0) continue;
		const doc = parseRawTree(xml, {
			part: story.part,
			limits
		});
		const paragraphs = docx.format === "pptx" ? walkSlideParagraphs(doc, {
			part: story.part,
			story: story.story,
			slide: story.slide ?? 0
		}) : walkParagraphs(doc, story);
		const tags = scan(paragraphs).map((record) => ({
			record,
			at: sourceOffset(record.paragraph, record.start)
		}));
		tags.sort((a, b) => a.at - b.at);
		const ordered = tags.map((t) => t.record);
		if (docx.format === "pptx") {
			stories.push({
				...story,
				doc,
				paragraphs,
				tags: ordered,
				tree: []
			});
			continue;
		}
		const tree = buildTree(ordered, limits);
		for (const problem of tree.problems) report(problem);
		stories.push({
			...story,
			doc,
			paragraphs,
			tags: ordered,
			tree: tree.children
		});
	}
	let trees = workbook === null ? stories.map((s) => s.tree) : workbook.sheets.map((s) => s.tree);
	const slideSections = /* @__PURE__ */ new Set();
	const droppedShapes = /* @__PURE__ */ new Set();
	if (docx.format === "pptx" && !full()) {
		const tree = buildTree(stories.flatMap((s) => s.tags), limits);
		for (const problem of tree.problems) report(problem);
		const local = presentationTrees(tree.children, stories, slideSections, droppedShapes, report);
		for (let i = 0; i < stories.length; i++) stories[i] = {
			...stories[i],
			tree: local.get(stories[i]) ?? []
		};
		trees = [tree.children];
	}
	if (problems.length > 0) {
		const first = problems[0];
		const count = problems.length;
		throw new TvError("TV_E_TEMPLATE_INVALID", count === 1 ? `The template has a problem: ${first.message}` : `The template has ${String(count)}${full() ? " or more" : ""} problems. The first: ${first.message}`, {
			location: first.location,
			hint: `Fix the problems listed in "errors"; each one has its place in the document and a hint. The first: ${first.hint}`,
			details: {
				count,
				truncated: full()
			},
			errors: problems
		});
	}
	const tags = workbook === null ? stories.flatMap((s) => s.tags) : workbook.sheets.flatMap((s) => s.tags);
	return {
		docx,
		stories,
		trees,
		slideSections,
		droppedShapes,
		tags,
		workbook
	};
}
/**
* Splits the tree of a presentation (D-036). A section whose opening and closing tags each fill a
* shape of their own on a slide repeats the slides from the one of its opening tag to the one of
* its closing tag; its two shapes are left out of the result, and what is inside it belongs to the
* slides. Every other section must open and close in one part. Returns the tree of each story,
* with the slide sections taken out.
*/
function presentationTrees(children, stories, slideSections, droppedShapes, report) {
	const storyOf = /* @__PURE__ */ new Map();
	for (const s of stories) for (const t of s.tags) storyOf.set(t, s);
	const local = /* @__PURE__ */ new Map();
	const list = (s) => {
		let l = local.get(s);
		if (l === void 0) {
			l = [];
			local.set(s, l);
		}
		return l;
	};
	const aloneInShape = (tag, story) => {
		if (story.story !== "slide") return void 0;
		const shape = shapeOf(tag.paragraph.element);
		return shape !== void 0 && shapeText(shape, story.paragraphs).trim() === tag.info.raw ? shape : void 0;
	};
	const walk = (nodes, insideOther) => {
		let lastSlide = 0;
		for (const child of nodes) {
			if ("tag" in child) {
				list(storyOf.get(child.tag)).push(child);
				continue;
			}
			const section = child.section;
			const openStory = storyOf.get(section.open);
			const closeStory = storyOf.get(section.close);
			const openShape = aloneInShape(section.open, openStory);
			const closeShape = aloneInShape(section.close, closeStory);
			if (openShape !== void 0 && closeShape !== void 0 && openShape !== closeShape) {
				if (insideOther) {
					report(new TvError("TV_E_SECTION_PLACEMENT", `The section "${section.open.info.raw}" repeats slides, but it is inside a section that does not.`, {
						location: section.open.info.location,
						hint: "A section that repeats slides can only be inside other sections that repeat slides. Put the outer section in shapes of its own too.",
						details: { tag: section.open.info.raw }
					}));
					continue;
				}
				const from = openStory.slide ?? 0;
				const to = closeStory.slide ?? 0;
				if (from <= lastSlide || to < from) {
					report(new TvError("TV_E_SECTION_PLACEMENT", `The section "${section.open.info.raw}" repeats slides that another section of the same level also repeats.`, {
						location: section.open.info.location,
						hint: "Let each section that repeats slides start on a slide after the last slide of the one before it, or nest one inside the other.",
						details: { tag: section.open.info.raw }
					}));
					continue;
				}
				lastSlide = to;
				slideSections.add(section);
				droppedShapes.add(openShape);
				droppedShapes.add(closeShape);
				walk(section.children, false);
				continue;
			}
			if (openStory !== closeStory) {
				report(new TvError("TV_E_SECTION_PLACEMENT", `The section "${section.open.info.raw}" opens and closes on different slides, but its tags do not each fill a shape of their own.`, {
					location: section.open.info.location,
					hint: `To repeat slides, put ${section.open.info.raw} alone in a shape on the first slide and {/${section.open.info.path}} alone in another shape on the last slide; otherwise keep the section within one slide or its notes.`,
					details: { tag: section.open.info.raw }
				}));
				continue;
			}
			list(openStory).push(child);
			checkNested(section.children, true);
		}
	};
	const checkNested = (nodes, insideOther) => {
		for (const child of nodes) {
			if ("tag" in child) continue;
			const section = child.section;
			const openStory = storyOf.get(section.open);
			const openShape = aloneInShape(section.open, openStory);
			const closeShape = aloneInShape(section.close, storyOf.get(section.close));
			if (insideOther && openShape !== void 0 && closeShape !== void 0 && openShape !== closeShape) report(new TvError("TV_E_SECTION_PLACEMENT", `The section "${section.open.info.raw}" repeats slides, but it is inside a section that does not.`, {
				location: section.open.info.location,
				hint: "A section that repeats slides can only be inside other sections that repeat slides.",
				details: { tag: section.open.info.raw }
			}));
			checkNested(section.children, insideOther);
		}
	};
	walk(children, false);
	return local;
}
// src/inspect.ts
const OPTION_NAMES$1 = ["limits", "delimiters"];
/** The `delimiters` option, checked. */
function delimitersOption(options) {
	const value = options === null || typeof options !== "object" ? void 0 : options.delimiters;
	if (value === void 0) return DEFAULT_DELIMITERS;
	const start = value !== null && typeof value === "object" ? value.start : void 0;
	const end = value !== null && typeof value === "object" ? value.end : void 0;
	if (typeof start !== "string" || typeof end !== "string") throw new TvError("TV_E_INVALID_OPTION", "Option \"delimiters\" must be an object with \"start\" and \"end\" strings.", {
		hint: "Pass for example { delimiters: { start: '{{', end: '}}' } }.",
		details: { option: "delimiters" }
	});
	const problem = delimiterProblem({
		start,
		end
	});
	if (problem !== null) throw new TvError("TV_E_INVALID_OPTION", `Option "delimiters" is not valid: ${problem}.`, {
		hint: "Use two different strings without spaces, such as { start: '[[', end: ']]' }.",
		details: { option: "delimiters" }
	});
	return {
		start,
		end
	};
}
/** Validates the options shared by inspect and render, and returns the limits. */
function checkOptions(options, names = OPTION_NAMES$1, what = "inspect") {
	if (options === void 0) return resolveLimits(void 0);
	if (options === null || typeof options !== "object" || Array.isArray(options)) throw new TvError("TV_E_INVALID_OPTION", `The options of ${what} must be an object.`, { hint: "Pass an object such as { limits: { maxTags: 1000 } }, or no options." });
	for (const key of Object.keys(options)) if (!names.includes(key)) throw new TvError("TV_E_INVALID_OPTION", `Unknown option "${clip(key)}".`, {
		hint: `Known options: ${names.join(", ")}.`,
		details: {
			option: clip(key),
			known: names
		}
	});
	return resolveLimits(options.limits);
}
/**
* Reads a template and returns its tags, its sections (loops and conditions) and the JSON Schema
* of the data it needs. Tags are found however the word processor split their text, in the main
* body, tables, text boxes, headers, footers, footnotes and endnotes of a document, in the shapes,
* tables and notes of a presentation, and in the text cells of every sheet of a workbook. Nothing
* is rendered.
*
* Throws {@link TvError}: `TV_E_TEMPLATE_INVALID` lists every problem with a tag or section in
* `errors` (up to `limits.maxErrors`), each with its code, place and hint;
* `TV_E_NOT_SUPPORTED` for other formats; ZIP, XML, package and limit errors for damaged or
* hostile files.
*
* @example
* ```ts
* import { readFileSync } from 'node:fs';
* import { inspect } from 'tplvane';
*
* const result = inspect(readFileSync('offer.docx'));
* for (const tag of result.tags) console.log(tag.raw, '-', tag.location.description);
* console.log(JSON.stringify(result.schema, null, 2));
* ```
*/
function inspect(template, options) {
	if (!(template instanceof Uint8Array)) throw new TvError("TV_E_INVALID_INPUT", "The template must be a Uint8Array (a Node.js Buffer is one).", { hint: "Read the file into bytes first, for example inspect(readFileSync(\"offer.docx\")) or inspect(new Uint8Array(await file.arrayBuffer()))." });
	const analysis = analyzeTemplate(template, checkOptions(options), delimitersOption(options));
	const tags = [];
	const sections = [];
	for (const tag of analysis.tags) tags.push(tag.info);
	for (const tree of analysis.trees) for (const child of tree) if ("section" in child) sections.push(toSectionNode(child.section));
	const { schema, warnings } = inferSchema(analysis.trees);
	return {
		format: analysis.docx.format,
		tags,
		sections,
		schema,
		warnings: [...analysis.docx.warnings, ...warnings]
	};
}
// src/docx/comments.ts
const COMMENTS = /\/comments$/;
const COMMENT_PARTS = /\/(?:commentsExtended|commentsIds|commentsExtensible)$/;
/** An attribute by local name, whatever its namespace. */
function attr$1(e, local) {
	for (const [key, value] of e.attrs) if (key === local || key.endsWith(`}${local}`)) return value;
}
function* elements$1(e) {
	for (const c of e.children) {
		if (c.kind !== "element") continue;
		yield c;
		yield* elements$1(c);
	}
}
/** `xml` without the children of its root for which `drop` holds; null when there are none. */
function without(xml, part, limits, drop) {
	const doc = parseRawTree(xml, {
		part,
		limits
	});
	const kept = doc.root.children.filter((c) => c.kind !== "element" || !drop(c));
	if (kept.length === doc.root.children.length) return null;
	doc.root.children = kept;
	return serialize$2(doc);
}
/**
* The comments parts of a document without the comments that no story refers to. `stories` gives
* the final XML of the parts that hold comment references, by partKey; it is called only when the
* document has comments.
*/
function pruneComments(pkg, main, stories, limits) {
	const out = /* @__PURE__ */ new Map();
	const rels = pkg.relationships(main);
	const comments = rels.find((r) => !r.external && COMMENTS.test(r.type) && pkg.hasPart(r.target));
	if (comments === void 0) return out;
	const referenced = /* @__PURE__ */ new Set();
	for (const [key, xml] of stories()) {
		if (!xml.includes("commentReference")) continue;
		for (const e of elements$1(parseRawTree(normalizeLineEnds(xml), {
			part: key,
			limits
		}).root)) if (e.local === "commentReference") {
			const id = attr$1(e, "id");
			if (id !== void 0) referenced.add(id);
		}
	}
	const xml = normalizeLineEnds(pkg.readXml(comments.target) ?? "");
	const doc = parseRawTree(xml, {
		part: comments.target,
		limits
	});
	const paraIds = /* @__PURE__ */ new Set();
	const durable = /* @__PURE__ */ new Set();
	let dropped = false;
	for (const c of doc.root.children) {
		if (c.kind !== "element" || c.local !== "comment" || referenced.has(attr$1(c, "id") ?? "")) continue;
		dropped = true;
		for (const e of elements$1(c)) {
			const paraId = e.local === "p" ? attr$1(e, "paraId") : void 0;
			if (paraId !== void 0) paraIds.add(paraId);
		}
		const id = attr$1(c, "durableId");
		if (id !== void 0) durable.add(id);
	}
	if (!dropped) return out;
	const pruned = without(xml, comments.target, limits, (e) => e.local === "comment" && !referenced.has(attr$1(e, "id") ?? ""));
	if (pruned !== null) out.set(partKey(comments.target), {
		name: comments.target,
		xml: pruned
	});
	const extensions = rels.filter((r) => !r.external && COMMENT_PARTS.test(r.type) && pkg.hasPart(r.target));
	for (const rel of extensions.filter((r) => !r.type.endsWith("/commentsExtensible"))) {
		const text = normalizeLineEnds(pkg.readXml(rel.target) ?? "");
		const tree = parseRawTree(text, {
			part: rel.target,
			limits
		});
		for (const e of tree.root.children) {
			const id = e.kind === "element" && paraIds.has(attr$1(e, "paraId") ?? "") ? attr$1(e, "durableId") : void 0;
			if (id !== void 0) durable.add(id);
		}
		const kept = without(text, rel.target, limits, (e) => paraIds.has(attr$1(e, "paraId") ?? ""));
		if (kept !== null) out.set(partKey(rel.target), {
			name: rel.target,
			xml: kept
		});
	}
	for (const rel of extensions.filter((r) => r.type.endsWith("/commentsExtensible"))) {
		const kept = without(normalizeLineEnds(pkg.readXml(rel.target) ?? ""), rel.target, limits, (e) => durable.has(attr$1(e, "durableId") ?? ""));
		if (kept !== null) out.set(partKey(rel.target), {
			name: rel.target,
			xml: kept
		});
	}
	return out;
}
// src/expr/evaluate.ts
/** Thrown when an expression needs a value that the data does not have. */
var MissingField = class extends Error {
	path;
	constructor(path) {
		super(`missing ${path}`);
		this.path = path;
	}
};
/** The value of an own data property, or undefined. Getters are not called. */
function ownValue(obj, key) {
	if (obj === null || typeof obj !== "object") return void 0;
	const d = Object.getOwnPropertyDescriptor(obj, key);
	return d !== void 0 && "value" in d ? d.value : void 0;
}
/** Whether `obj` has `key` as an own data property. */
function hasOwnValue(obj, key) {
	if (obj === null || typeof obj !== "object") return false;
	const d = Object.getOwnPropertyDescriptor(obj, key);
	return d !== void 0 && "value" in d;
}
/**
* Looks a path up in the scopes, innermost first: the first name is searched in each scope
* outwards, the rest only below where it was found. Returns undefined when it is missing.
*/
function lookup(scopes, segments) {
	if (segments.length === 0) return scopes[scopes.length - 1];
	const first = segments[0];
	for (let i = scopes.length - 1; i >= 0; i--) {
		const scope = scopes[i];
		if (!hasOwnValue(scope, first)) continue;
		let value = ownValue(scope, first);
		for (let j = 1; j < segments.length; j++) {
			value = ownValue(value, segments[j]);
			if (value === void 0) return void 0;
		}
		return value;
	}
}
/** The truth of a value in sections and logical operators. */
function truthy(value) {
	if (Array.isArray(value)) return value.length > 0;
	return Boolean(value);
}
function kind(value) {
	if (value === null) return "null";
	if (Array.isArray(value)) return "a list";
	return typeof value === "object" ? "an object" : typeof value;
}
/**
* Evaluates `e`. A missing path gives undefined, except where the value must exist (an operand
* of a comparison), where it throws {@link MissingField}. Formatters throw ValueError.
*/
function evaluate(e, scopes, fmt) {
	switch (e.t) {
		case "literal": return e.value;
		case "path": return lookup(scopes, e.segments);
		case "not": return !truthy(evaluate(e.e, scopes, fmt));
		case "and": {
			const l = evaluate(e.l, scopes, fmt);
			return truthy(l) ? evaluate(e.r, scopes, fmt) : l;
		}
		case "or": {
			const l = evaluate(e.l, scopes, fmt);
			return truthy(l) ? l : evaluate(e.r, scopes, fmt);
		}
		case "compare": {
			const l = operand(e.l, scopes, fmt);
			const r = operand(e.r, scopes, fmt);
			switch (e.op) {
				case "==": return l === r;
				case "!=": return l !== r;
				default: {
					if (!(typeof l === "number" && typeof r === "number") && !(typeof l === "string" && typeof r === "string")) throw new ValueError(`"${e.op}" compares two numbers or two texts, not ${kind(l)} and ${kind(r)}.`, "Make both sides numbers (or both texts) in the data, or compare with == instead.");
					const a = l;
					const b = r;
					return e.op === "<" ? a < b : e.op === "<=" ? a <= b : e.op === ">" ? a > b : a >= b;
				}
			}
		}
		case "pipe": {
			let value = evaluate(e.e, scopes, fmt);
			for (const f of e.formatters) value = applyFormatter(f, value, fmt);
			return value;
		}
	}
}
function operand(e, scopes, fmt) {
	const value = evaluate(e, scopes, fmt);
	if (value === void 0) throw new MissingField(firstMissing(e, scopes) ?? "?");
	return value;
}
/** The first path of `e` that the data does not have. */
function firstMissing(e, scopes) {
	switch (e.t) {
		case "path": return lookup(scopes, e.segments) === void 0 ? e.segments.length === 0 ? "." : e.segments.join(".") : null;
		case "not":
		case "pipe": return firstMissing(e.e, scopes);
		case "and":
		case "or":
		case "compare": return firstMissing(e.l, scopes) ?? firstMissing(e.r, scopes);
		default: return null;
	}
}
// src/media/image-info.ts
const DEFAULT_DPI = 96;
const u16be = (b, at) => (b[at] ?? 0) << 8 | (b[at + 1] ?? 0);
const u32be = (b, at) => ((b[at] ?? 0) << 24 | (b[at + 1] ?? 0) << 16 | (b[at + 2] ?? 0) << 8 | (b[at + 3] ?? 0)) >>> 0;
const u16le = (b, at) => (b[at] ?? 0) | (b[at + 1] ?? 0) << 8;
function starts(b, sig) {
	return b.length >= sig.length && sig.every((x, i) => b[i] === x);
}
const PNG = [
	137,
	80,
	78,
	71,
	13,
	10,
	26,
	10
];
function png(b) {
	if (b.length < 33 || u32be(b, 8) !== 13 || String.fromCharCode(...b.subarray(12, 16)) !== "IHDR") return "The PNG image has no valid header (IHDR).";
	const width = u32be(b, 16);
	const height = u32be(b, 20);
	let dpiX = DEFAULT_DPI;
	let dpiY = DEFAULT_DPI;
	for (let at = 8; at + 12 <= b.length;) {
		const length = u32be(b, at);
		const type = String.fromCharCode(...b.subarray(at + 4, at + 8));
		if (type === "IDAT" || type === "IEND") break;
		if (type === "pHYs" && length === 9 && at + 17 <= b.length && b[at + 16] === 1) {
			dpiX = Math.round(u32be(b, at + 8) * .0254) || DEFAULT_DPI;
			dpiY = Math.round(u32be(b, at + 12) * .0254) || DEFAULT_DPI;
		}
		at += 12 + length;
	}
	return {
		type: "png",
		extension: "png",
		contentType: "image/png",
		width,
		height,
		dpiX,
		dpiY
	};
}
function jpeg(b) {
	let dpiX = DEFAULT_DPI;
	let dpiY = DEFAULT_DPI;
	for (let at = 2; at + 4 <= b.length;) {
		if (b[at] !== 255) return "The JPEG image has a damaged structure.";
		const marker = b[at + 1];
		if (marker === 255) {
			at++;
			continue;
		}
		if (marker === 216 || marker >= 208 && marker <= 215 || marker === 1) {
			at += 2;
			continue;
		}
		const length = u16be(b, at + 2);
		if (length < 2) return "The JPEG image has a damaged structure.";
		if (marker === 224 && length >= 16 && String.fromCharCode(...b.subarray(at + 4, at + 9)) === "JFIF\0") {
			const units = b[at + 11];
			const x = u16be(b, at + 12);
			const y = u16be(b, at + 14);
			if ((units === 1 || units === 2) && x > 0 && y > 0) {
				dpiX = units === 1 ? x : Math.round(x * 2.54);
				dpiY = units === 1 ? y : Math.round(y * 2.54);
			}
		}
		if (marker >= 192 && marker <= 207 && marker !== 196 && marker !== 200 && marker !== 204) {
			if (at + 9 > b.length) return "The JPEG image is truncated.";
			return {
				type: "jpeg",
				extension: "jpeg",
				contentType: "image/jpeg",
				width: u16be(b, at + 7),
				height: u16be(b, at + 5),
				dpiX,
				dpiY
			};
		}
		if (marker === 218 || marker === 217) break;
		at += 2 + length;
	}
	return "The JPEG image has no frame header.";
}
function gif(b) {
	if (b.length < 10) return "The GIF image is truncated.";
	return {
		type: "gif",
		extension: "gif",
		contentType: "image/gif",
		width: u16le(b, 6),
		height: u16le(b, 8),
		dpiX: DEFAULT_DPI,
		dpiY: DEFAULT_DPI
	};
}
/** The type, size and resolution of an image, or a message saying why it cannot be used. */
function imageInfo(b) {
	let info;
	if (starts(b, PNG)) info = png(b);
	else if (starts(b, [
		255,
		216,
		255
	])) info = jpeg(b);
	else if (starts(b, [
		71,
		73,
		70,
		56
	]) && (b[4] === 55 || b[4] === 57) && b[5] === 97) info = gif(b);
	else return "The image is not a PNG, JPEG or GIF file.";
	if (typeof info !== "string" && (info.width === 0 || info.height === 0)) return "The image has a width or height of 0 pixels.";
	return info;
}
// src/markdown/parse.ts
const PUNCTUATION = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";
const MAX_LIST_LEVEL = 2;
/** The heading level and text of an ATX heading line, or null. */
function heading(line) {
	let i = 0;
	while (i < 3 && line[i] === " ") i++;
	let level = 0;
	while (line[i + level] === "#") level++;
	if (level < 1 || level > 6) return null;
	const after = line[i + level];
	if (after !== void 0 && after !== " " && after !== "	") return null;
	let text = line.slice(i + level).trim();
	let end = text.length;
	while (end > 0 && text[end - 1] === "#") end--;
	if (end === 0 || text[end - 1] === " " || text[end - 1] === "	") text = text.slice(0, end).trimEnd();
	return {
		level,
		text
	};
}
/** The level, kind, number and text of a list item line, or null. */
function listItem(line) {
	let spaces = 0;
	while (line[spaces] === " ") spaces++;
	const c = line[spaces];
	let at = spaces;
	let ordered = false;
	let number = 1;
	if (c === "-" || c === "*" || c === "+") at++;
	else {
		let digits = 0;
		while (digits < 9 && line[at + digits] !== void 0 && line[at + digits] >= "0" && line[at + digits] <= "9") digits++;
		if (digits === 0 || line[at + digits] !== "." && line[at + digits] !== ")") return null;
		number = Number(line.slice(at, at + digits));
		ordered = true;
		at += digits + 1;
	}
	if (line[at] !== " " && line[at] !== "	") return null;
	return {
		level: Math.min(MAX_LIST_LEVEL, Math.floor(spaces / 2)),
		ordered,
		number,
		text: line.slice(at + 1).trim()
	};
}
/** The index of the next occurrence of each character, for every position, from the end. */
function nextOf(text, char) {
	const next = new Int32Array(text.length + 1).fill(-1);
	for (let i = text.length - 1; i >= 0; i--) next[i] = text[i] === char ? i : next[i + 1];
	return next;
}
/** Parses inline Markdown. */
function parseInline(text) {
	const tokens = [];
	const nextTick = nextOf(text, "`");
	const nextBracket = nextOf(text, "]");
	const nextParen = nextOf(text, ")");
	let buf = "";
	let pendingBreak = false;
	const flush = () => {
		if (buf !== "" || pendingBreak) {
			tokens.push(pendingBreak ? {
				kind: "text",
				text: buf,
				breakBefore: true
			} : {
				kind: "text",
				text: buf
			});
			buf = "";
			pendingBreak = false;
		}
	};
	const isWord = (ch) => ch !== void 0 && /[\p{L}\p{N}]/u.test(ch);
	for (let i = 0; i < text.length;) {
		const c = text[i];
		if (c === "\\" && i + 1 < text.length && PUNCTUATION.includes(text[i + 1])) {
			buf += text[i + 1];
			i += 2;
			continue;
		}
		if (c === "\n") {
			flush();
			pendingBreak = true;
			i++;
			continue;
		}
		if (c === "`") {
			let run = 1;
			while (text[i + run] === "`") run++;
			if (run === 1) {
				const end = nextTick[i + 1] ?? -1;
				if (end !== -1) {
					flush();
					tokens.push({
						kind: "code",
						text: text.slice(i + 1, end)
					});
					i = end + 1;
					continue;
				}
			} else {
				const fence = "`".repeat(run);
				let end = text.indexOf(fence, i + run);
				while (end !== -1 && end - i < SEARCH && text[end + run] === "`") end = text.indexOf(fence, end + run + 1);
				if (end !== -1 && end - i < SEARCH) {
					flush();
					const inner = text.slice(i + run, end);
					tokens.push({
						kind: "code",
						text: inner.startsWith(" ") && inner.endsWith(" ") && inner.trim() !== "" ? inner.slice(1, -1) : inner
					});
					i = end + run;
					continue;
				}
			}
			buf += "`".repeat(run);
			i += run;
			continue;
		}
		const isImage = c === "!" && text[i + 1] === "[";
		if (c === "[" || isImage) {
			const start = isImage ? i + 1 : i;
			const close = nextBracket[start + 1] ?? -1;
			const end = close !== -1 && text[close + 1] === "(" ? linkEnd(text, close + 2, nextParen) : -1;
			if (end !== -1) {
				const label = text.slice(start + 1, close);
				const url = text.slice(close + 2, end).trim();
				if (isImage || url === "") buf += text.slice(i, end + 1);
				else {
					flush();
					tokens.push({
						kind: "link",
						text: label === "" ? url : label,
						url
					});
				}
				i = end + 1;
				continue;
			}
		}
		if (c === "*" || c === "_") {
			let run = 1;
			while (text[i + run] === c) run++;
			const before = text[i - 1];
			const after = text[i + run];
			const size = run >= 2 ? 2 : 1;
			const intraword = isWord(before) && isWord(after);
			const open = after !== void 0 && after !== " " && !(c === "_" && intraword);
			const close = before !== void 0 && before !== " " && !(c === "_" && intraword);
			if (open || close) {
				flush();
				for (let k = 0; k < Math.floor(run / size); k++) tokens.push({
					kind: "delim",
					char: c,
					size,
					open,
					close
				});
				if (run % size === 1) tokens.push({
					kind: "delim",
					char: c,
					size: 1,
					open,
					close
				});
				i += run;
				continue;
			}
			buf += text.slice(i, i + run);
			i += run;
			continue;
		}
		buf += c;
		i++;
	}
	flush();
	const role = new Array(tokens.length).fill(null);
	const stacks = /* @__PURE__ */ new Map();
	tokens.forEach((t, i) => {
		if (t.kind !== "delim") return;
		const key = `${t.char}${String(t.size)}`;
		const stack = stacks.get(key) ?? [];
		if (t.close && stack.length > 0) {
			role[stack.pop()] = "open";
			role[i] = "close";
		} else if (t.open) stack.push(i);
		stacks.set(key, stack);
	});
	const out = [];
	let bold = 0;
	let italic = 0;
	let breakNext = false;
	const push = (inline) => {
		const withBreak = breakNext ? {
			...inline,
			breakBefore: true
		} : inline;
		breakNext = false;
		out.push({
			bold: bold > 0,
			italic: italic > 0,
			...withBreak
		});
	};
	tokens.forEach((t, i) => {
		switch (t.kind) {
			case "text":
				if (t.breakBefore === true) breakNext = true;
				if (t.text !== "") push({
					text: t.text,
					code: false
				});
				return;
			case "code":
				push({
					text: t.text,
					code: true
				});
				return;
			case "link":
				push({
					text: t.text === "" ? t.url : t.text,
					code: false,
					link: t.url
				});
				return;
			case "delim":
				if (role[i] === null) {
					push({
						text: t.char.repeat(t.size),
						code: false
					});
					return;
				}
				if (t.size === 2) bold += role[i] === "open" ? 1 : -1;
				else italic += role[i] === "open" ? 1 : -1;
		}
	});
	return out;
}
/** Parses Markdown into blocks. Lines are normalised to LF first. */
function parseMarkdown(markdown) {
	const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
	const blocks = [];
	let paragraph = [];
	let item = null;
	const joinLines = (parts) => parts.map((line, i) => {
		if (i === parts.length - 1) return line.trim();
		if (line.endsWith("  ")) return `${line.trim()}\n`;
		if (line.endsWith("\\")) return `${line.slice(0, -1).trim()}\n`;
		return `${line.trim()} `;
	}).join("");
	const end = () => {
		if (paragraph.length > 0) blocks.push({
			kind: "paragraph",
			content: parseInline(joinLines(paragraph))
		});
		if (item !== null) blocks.push({
			kind: "item",
			ordered: item.ordered,
			level: item.level,
			number: item.number,
			content: parseInline(joinLines(item.lines))
		});
		paragraph = [];
		item = null;
	};
	for (const line of lines) {
		if (line.trim() === "") {
			end();
			continue;
		}
		const h = heading(line);
		if (h !== null) {
			end();
			blocks.push({
				kind: "heading",
				level: h.level,
				content: parseInline(h.text)
			});
			continue;
		}
		const li = listItem(line);
		if (li !== null) {
			end();
			item = {
				ordered: li.ordered,
				level: li.level,
				number: li.number,
				lines: [li.text]
			};
			continue;
		}
		if (item !== null) item.lines.push(line);
		else paragraph.push(line);
	}
	end();
	return blocks;
}
/** How far a closing backtick run or parenthesis is looked for (characters). */
const SEARCH = 4096;
/**
* The index of the ")" that ends a link address starting at `from`: parentheses inside it are
* balanced (http://x.y/a_(b)), within SEARCH characters; -1 when there is none (D-071).
*/
function linkEnd(text, from, nextParen) {
	const first = nextParen[from] ?? -1;
	if (first === -1) return -1;
	const open = text.indexOf("(", from);
	if (open === -1 || open > first) return first;
	let depth = 0;
	for (let i = from; i < text.length && i - from < SEARCH; i++) {
		const ch = text[i];
		if (ch === "\n") return -1;
		if (ch === "(") depth++;
		else if (ch === ")") {
			if (depth === 0) return i;
			depth--;
		}
	}
	return first;
}
// src/xml/writer.ts
/** Escapes character data for element content. */
function escapeText(s) {
	return /[&<>\r]/.test(s) ? s.replace(/[&<>\r]/g, (c) => c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&#13;") : s;
}
const ATTR_ESCAPES = /* @__PURE__ */ new Map([
	["&", "&amp;"],
	["<", "&lt;"],
	[">", "&gt;"],
	["\"", "&quot;"],
	["	", "&#9;"],
	["\n", "&#10;"],
	["\r", "&#13;"]
]);
/** Escapes an attribute value that will be enclosed in double quotes. */
function escapeAttr(s) {
	return /[&<>"\t\n\r]/.test(s) ? s.replace(/[&<>"\t\n\r]/g, (c) => ATTR_ESCAPES.get(c) ?? c) : s;
}
/**
* Index (in UTF-16 code units) of the first character of `s` that XML 1.0 cannot hold at all, or
* -1 if there is none: a code point outside the Char production (§2.2), such as U+0001 or U+FFFE,
* or an unpaired surrogate. `escapeText` and `escapeAttr` leave such characters in place, so
* text from untrusted data must be checked with this before it is written.
*/
function firstInvalidXmlChar(s) {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c >= 55296 && c <= 56319) {
			const next = s.charCodeAt(i + 1);
			if (next >= 56320 && next <= 57343) {
				if (!isXmlChar((c - 55296 << 10) + (next - 56320) + 65536)) return i;
				i++;
				continue;
			}
			return i;
		}
		if (!isXmlChar(c)) return i;
	}
	return -1;
}
// src/util/data-path.ts
/** A data path as JSON Pointer (RFC 6901): `/items/1/name`. */
function pointerOf(path) {
	return path.map((p) => `/${String(p).replace(/~/g, "~0").replace(/\//g, "~1")}`).join("");
}
/** A data path as people write it: `items[1].name`. */
function dataPathOf(path) {
	let out = "";
	for (const p of path) out += typeof p === "number" ? `[${String(p)}]` : out === "" ? p : `.${p}`;
	return out;
}
// src/template/evaluator.ts
/** The paths of each scope of `scopes`. */
function pathsOf(ctx, scopes) {
	return ctx.scopePaths.get(scopes) ?? scopes.map(() => []);
}
/** The index of the scope where `first` is found (the innermost with it as an own data field), or the innermost scope. */
function scopeOf(scopes, first) {
	if (first !== void 0) for (let i = scopes.length - 1; i >= 0; i--) {
		const scope = scopes[i];
		if (scope !== null && typeof scope === "object" && ownValue(scope, first) !== void 0) return i;
	}
	return scopes.length - 1;
}
/**
* Where a path of a tag is in the data: the path up to its first missing field, and the names that
* the object holding that missing field offers (none when nothing is missing).
*/
function locateInData(ctx, scopes, segments) {
	const at = scopeOf(scopes, segments[0]);
	const path = [...pathsOf(ctx, scopes)[at] ?? []];
	let holder = scopes[at];
	for (const name of segments) {
		path.push(name);
		const next = ownValue(holder, name);
		if (next === void 0) return {
			path,
			available: namesOf(holder)
		};
		holder = next;
	}
	return {
		path,
		available: []
	};
}
/** The names of the own data fields of an object, up to 20. */
function namesOf(obj) {
	if (obj === null || typeof obj !== "object" || Array.isArray(obj)) return [];
	const names = [];
	for (const key in obj) {
		if (ownValue(obj, key) === void 0) continue;
		names.push(key);
		if (names.length >= 20) break;
	}
	return names;
}
/** The expression of a tag with its formatters. */
function fullExpr(tag) {
	return tag.parsed.formatters.length === 0 ? tag.parsed.expr : {
		t: "pipe",
		e: tag.parsed.expr,
		formatters: tag.parsed.formatters
	};
}
/** The names that the data offers at this point, for "did you mean" hints. */
function availableNames(scopes) {
	const names = /* @__PURE__ */ new Set();
	for (let i = scopes.length - 1; i >= 0 && names.size < 200; i--) {
		const scope = scopes[i];
		if (scope === null || typeof scope !== "object" || Array.isArray(scope)) continue;
		for (const key in scope) {
			if (names.size >= 200) break;
			if (!names.has(key) && ownValue(scope, key) !== void 0) names.add(key);
		}
	}
	return [...names];
}
/** Evaluates tags; each format's renderer extends it. */
var TagEvaluator = class {
	ctx;
	constructor(ctx) {
		this.ctx = ctx;
	}
	problem(error) {
		if (this.ctx.problems.length < this.ctx.limits.maxErrors) this.ctx.problems.push(error);
	}
	missing(tag, path, scopes) {
		if (this.ctx.missing === "empty") return;
		if (this.ctx.problems.length >= this.ctx.limits.maxErrors) return;
		const where = locateInData(this.ctx, scopes, path === "." ? [] : path.split("."));
		const missingName = String(where.path[where.path.length - 1] ?? path);
		const first = path.split(".")[0] ?? path;
		const close = suggest(missingName, !scopes.some((scope) => ownValue(scope, first) !== void 0) ? availableNames(scopes) : where.available);
		const here = where.available.length > 0 ? ` The data has here: ${where.available.map(clip).join(", ")}.` : "";
		this.problem(new TvError("TV_E_MISSING_FIELD", `The data has no "${clip(path)}", which ${tag.info.raw.length > 60 ? "a tag" : `"${clip(tag.info.raw)}"`} needs.`, {
			location: tag.info.location,
			hint: `${close.length > 0 ? `Did you mean ${close.map((c) => `"${c}"`).join(" or ")}? ` : ""}Add "${clip(path)}" to the data, ${tag.info.kind === "value" ? "give the tag a default (| default(\"-\")), " : ""}or render with { missing: "empty" }.${here}`,
			details: {
				path: clip(path),
				tag: clip(tag.info.raw),
				dataPath: clip(dataPathOf(where.path)),
				pointer: clip(pointerOf(where.path)),
				available: where.available.map(clip)
			}
		}));
	}
	/** The tag last evaluated and its scopes, to give problems found afterwards the place of the value in the data. */
	current = null;
	invalid(tag, message, hint) {
		const segments = tag.parsed.segments;
		const place = this.current !== null && this.current.tag === tag && segments !== null ? (() => {
			const { path } = locateInData(this.ctx, this.current.scopes, segments);
			return {
				dataPath: clip(dataPathOf(path)),
				pointer: clip(pointerOf(path))
			};
		})() : {};
		this.problem(new TvError("TV_E_INVALID_VALUE", message, {
			location: tag.info.location,
			hint,
			details: {
				tag: clip(tag.info.raw),
				...place
			}
		}));
	}
	/** Evaluates the expression of a tag; problems are recorded and give undefined. */
	evaluate(tag, scopes) {
		this.current = {
			tag,
			scopes
		};
		try {
			return {
				ok: true,
				value: evaluate(fullExpr(tag), scopes, this.ctx.formatting)
			};
		} catch (error) {
			if (error instanceof MissingField) this.missing(tag, error.path, scopes);
			else if (error instanceof ValueError) this.invalid(tag, `${error.message} (in "${clip(tag.info.raw)}")`, error.hint);
			else throw error;
			return {
				ok: false,
				value: void 0
			};
		}
	}
	/** The text of a value tag (or a link from the link formatter), or null when it has a problem. */
	valueText(tag, scopes) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return null;
		if (value === void 0) {
			this.missing(tag, firstMissing(fullExpr(tag), scopes) ?? tag.info.path, scopes);
			return null;
		}
		if (value instanceof LinkValue) return this.checkText(tag, value.text) === null ? null : value;
		let text;
		if (value === null) text = "";
		else if (typeof value === "string") text = value;
		else if (typeof value === "number") {
			if (!Number.isFinite(value)) {
				this.invalid(tag, `The value of "${clip(tag.info.raw)}" is ${String(value)}, which cannot be written.`, "Pass a finite number.");
				return null;
			}
			text = String(value);
		} else if (typeof value === "boolean" || typeof value === "bigint") text = String(value);
		else {
			const what = Array.isArray(value) ? "a list" : value instanceof Date ? "a date" : typeof value === "object" ? "an object" : `a ${typeof value}`;
			this.invalid(tag, `The value of "${clip(tag.info.raw)}" is ${what}, which cannot be written as text.`, Array.isArray(value) ? `Use a section to repeat the list: {#${tag.info.path}}…{/${tag.info.path}}.` : value instanceof Date ? `Format it with a formatter, for example {${tag.info.path} | date("long")}.` : "Write one of its fields, for example {customer.name} instead of {customer}.");
			return null;
		}
		return this.checkText(tag, text);
	}
	/** `text` when a document can hold all its characters; otherwise records a problem and gives null. */
	checkText(tag, text) {
		const bad = firstInvalidXmlChar(text);
		if (bad === -1) return text;
		this.invalid(tag, `The value of "${clip(tag.info.raw)}" contains the character U+${(text.codePointAt(bad) ?? 0).toString(16).toUpperCase().padStart(4, "0")} at position ${String(bad + 1)}, which a document cannot hold.`, "Remove control characters (other than tab and line breaks) and unpaired surrogates from the data.");
		return null;
	}
	/** The scopes of each repetition of a section, or an empty list when its content is not shown. */
	iterations(section, scopes) {
		const open = section.open;
		const { ok, value } = this.evaluate(open, scopes);
		if (!ok) return [];
		if (open.info.kind === "inverted") return truthy(value) ? [] : [scopes];
		if (value === void 0) {
			if (section.kind === "loop") this.missing(open, firstMissing(fullExpr(open), scopes) ?? open.info.path, scopes);
			return [];
		}
		const parents = pathsOf(this.ctx, scopes);
		const segments = open.parsed.segments;
		const base = segments === null ? parents[parents.length - 1] ?? [] : [...parents[scopeOf(scopes, segments[0])] ?? [], ...segments];
		const scoped = (item, path) => {
			const next = [...scopes, item];
			this.ctx.scopePaths.set(next, [...parents, path]);
			return next;
		};
		if (Array.isArray(value)) {
			if (this.ctx.iterations + value.length > this.ctx.limits.maxIterations) throw limitError("maxIterations", this.ctx.limits.maxIterations, `Sections repeat their content more than ${String(this.ctx.limits.maxIterations)} times in all.`, open.info.location);
			const out = [];
			for (let i = 0; i < value.length; i++) out.push(scoped(ownValue(value, String(i)), [...base, i]));
			return out;
		}
		return truthy(value) ? [scoped(value, base)] : [];
	}
};
// src/opc/conformance.ts
const TRANSITIONAL_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";
/** Transitional names (ECMA-376 Part 4). */
const TRANSITIONAL = {
	strict: false,
	r: TRANSITIONAL_R,
	a: "http://schemas.openxmlformats.org/drawingml/2006/main",
	pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
	wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
	imageRel: `${TRANSITIONAL_R}/image`,
	hyperlinkRel: `${TRANSITIONAL_R}/hyperlink`,
	slideRel: `${TRANSITIONAL_R}/slide`
};
/** Strict names (ECMA-376 Part 1). */
const STRICT = {
	strict: true,
	r: STRICT_R,
	a: "http://purl.oclc.org/ooxml/drawingml/main",
	pic: "http://purl.oclc.org/ooxml/drawingml/picture",
	wp: "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
	imageRel: `${STRICT_R}/image`,
	hyperlinkRel: `${STRICT_R}/hyperlink`,
	slideRel: `${STRICT_R}/slide`
};
/** The vocabulary of a package, from the type of its officeDocument relationship. */
function vocabularyOf(officeDocumentType) {
	return officeDocumentType.startsWith(`${STRICT_R}/`) ? STRICT : TRANSITIONAL;
}
// src/docx/content.ts
/** A node that is written as it is: for content made by the renderer. */
function rawXml(xml) {
	return {
		kind: "text",
		raw: xml,
		text: "",
		parent: null,
		start: -1
	};
}
function serialize$1(node) {
	const out = [];
	serializeInto(node, out);
	return out.join("");
}
const isWml$1 = (n, local) => n.kind === "element" && WML_NAMESPACES.has(n.ns) && n.local === local;
/**
* The w:rPr for a run: the properties of the tag's run (`base`) with `style` applied. w:rStyle comes
* first, w:b and w:i after w:rStyle and w:rFonts, as LibreOffice writes them.
*/
function runProperties(p, base, style = {}) {
	const children = (base?.children ?? []).filter((c) => c.kind === "element");
	const drop = (c) => style.rStyle !== void 0 && isWml$1(c, "rStyle") || style.bold === true && (isWml$1(c, "b") || isWml$1(c, "bCs")) || style.italic === true && (isWml$1(c, "i") || isWml$1(c, "iCs"));
	const kept = children.filter((c) => !drop(c));
	const head = [];
	const rest = [];
	for (const c of kept) (isWml$1(c, "rStyle") || isWml$1(c, "rFonts") ? head : rest).push(serialize$1(c));
	const added = [...style.bold === true ? [`<${p}b/><${p}bCs/>`] : [], ...style.italic === true ? [`<${p}i/><${p}iCs/>`] : []];
	const inner = [
		...style.rStyle === void 0 ? [] : [`<${p}rStyle ${p}val="${escapeAttr(style.rStyle)}"/>`],
		...head,
		...added,
		...rest
	].join("");
	return inner === "" ? "" : `<${p}rPr>${inner}</${p}rPr>`;
}
/** The w:t, w:tab and w:br elements for a text: tabs and line breaks become elements. */
function textContent(p, text) {
	const out = [];
	text.replace(/\r\n?/g, "\n").split("\n").forEach((line, i) => {
		if (i > 0) out.push(`<${p}br/>`);
		line.split("	").forEach((piece, j) => {
			if (j > 0) out.push(`<${p}tab/>`);
			if (piece !== "") out.push(`<${p}t${/^\s|\s$|\s\s/.test(piece) ? " xml:space=\"preserve\"" : ""}>${escapeText(piece)}</${p}t>`);
		});
	});
	return out.join("");
}
/** A run with the start tag of `model` (the tag's run), properties and text. */
function runXml(model, rPr, text, breakBefore = false) {
	const p = model.qname.includes(":") ? model.qname.slice(0, model.qname.indexOf(":") + 1) : "";
	const open = model.open.endsWith("/>") ? `<${model.qname}>` : model.open;
	const close = model.close === "" ? `</${model.qname}>` : model.close;
	return `${open}${rPr}${breakBefore ? `<${p}br/>` : ""}${textContent(p, text)}${close}`;
}
/** A hyperlink around runs, to an external relationship. */
function hyperlinkXml(p, rId, runs, names = TRANSITIONAL) {
	return `<${p}hyperlink xmlns:r="${names.r}" r:id="${escapeAttr(rId)}">${runs}</${p}hyperlink>`;
}
/** A run with an inline picture of cx × cy EMU, as LibreOffice writes one (F-015). */
function pictureRunXml(model, rPr, picture, names = TRANSITIONAL) {
	const p = model.qname.includes(":") ? model.qname.slice(0, model.qname.indexOf(":") + 1) : "";
	const open = model.open.endsWith("/>") ? `<${model.qname}>` : model.open;
	const close = model.close === "" ? `</${model.qname}>` : model.close;
	const { rId, cx, cy, id } = picture;
	const name = escapeAttr(picture.name);
	const alt = escapeAttr(picture.alt);
	return `${open}${rPr}<${p}drawing><wp:inline xmlns:wp="${names.wp}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${String(cx)}" cy="${String(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${String(id)}" name="${name}" descr="${alt}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${names.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${names.a}"><a:graphicData uri="${names.pic}"><pic:pic xmlns:pic="${names.pic}"><pic:nvPicPr><pic:cNvPr id="${String(id)}" name="${name}" descr="${alt}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="${names.r}" r:embed="${escapeAttr(rId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></${p}drawing>${close}`;
}
/**
* The w:pPr of a paragraph made from `template`: its properties, with the paragraph style replaced
* by `styleId` (and direct numbering removed, which would override the style's own) when given.
*/
function paragraphProperties(p, template, styleId, numbering) {
	const pPr = template.children.find((c) => isWml$1(c, "pPr"));
	if (styleId === void 0) return pPr === void 0 ? "" : serialize$1(pPr);
	const children = (pPr?.children ?? []).filter((c) => !(isWml$1(c, "pStyle") || isWml$1(c, "numPr")));
	const kept = children.map(serialize$1);
	if (numbering !== void 0) {
		const before = /* @__PURE__ */ new Set([
			"keepNext",
			"keepLines",
			"pageBreakBefore",
			"framePr",
			"widowControl"
		]);
		let at = 0;
		children.forEach((c, i) => {
			if (c.kind === "element" && WML_NAMESPACES.has(c.ns) && before.has(c.local)) at = i + 1;
		});
		kept.splice(at, 0, `<${p}numPr><${p}ilvl ${p}val="${String(numbering.level)}"/><${p}numId ${p}val="${escapeAttr(numbering.numId)}"/></${p}numPr>`);
	}
	const open = pPr === void 0 || pPr.open.endsWith("/>") ? `<${p}pPr>` : pPr.open;
	const close = pPr === void 0 || pPr.close === "" ? `</${p}pPr>` : pPr.close;
	return `${open}<${p}pStyle ${p}val="${escapeAttr(styleId)}"/>${kept.join("")}${close}`;
}
/** A paragraph with the start tag of `template`, properties and content. */
function paragraphXml(template, pPr, content) {
	return `${template.open.endsWith("/>") ? `<${template.qname}>` : template.open}${pPr}${content}${template.close === "" ? `</${template.qname}>` : template.close}`;
}
/** The w:rPr child of a run, if any. */
function runPropertiesOf(run) {
	return run.children.find((c) => isWml$1(c, "rPr"));
}
// src/docx/normalize.ts
const WML_DIALECT = {
	namespaces: WML_NAMESPACES,
	preserveSpace: true
};
const isIn = (d) => (e, local) => e.kind === "element" && d.namespaces.has(e.ns) && e.local === local;
/**
* Collects replacements of nodes and applies them with one pass over the children of each parent,
* so that a paragraph with many runs is rewritten in linear time (D-062).
*/
var Replacements = class {
	byParent = /* @__PURE__ */ new Map();
	/** Replaces `node` in its parent by `replacements` at the next flush. */
	add(node, replacements) {
		const parent = node.parent;
		if (parent === null) return;
		const map = this.byParent.get(parent) ?? /* @__PURE__ */ new Map();
		map.set(node, replacements);
		this.byParent.set(parent, map);
	}
	flush() {
		for (const [parent, map] of this.byParent) {
			const children = [];
			for (const child of parent.children) {
				const replacement = map.get(child);
				if (replacement === void 0) {
					children.push(child);
					continue;
				}
				for (const r of replacement) {
					r.parent = parent;
					children.push(r);
				}
			}
			parent.children = children;
		}
		this.byParent.clear();
	}
};
/**
* A new run with the start and end tags of `run`, a shared copy of its properties, and `content`,
* which is moved into it (its nodes keep their identity, so records that point to them stay valid).
*/
function runWith(run, rPr, content) {
	const copy = makeElement(run, run.open.endsWith("/>") ? run.open.replace(/\s*\/>$/, ">") : run.open, run.close === "" ? `</${run.qname}>` : run.close, run.attrs);
	const children = [];
	if (rPr !== void 0) children.push({
		...rPr,
		parent: copy
	});
	for (const c of content) {
		c.parent = copy;
		children.push(c);
	}
	copy.children = children;
	return copy;
}
/** A new w:t with `text`, written like `model`. */
function textElement(model, text, dialect = WML_DIALECT) {
	const p = prefixOf(model.qname);
	const t = makeElement(model, `<${p}t${dialect.preserveSpace && /^\s|\s$|\s\s/.test(text) ? " xml:space=\"preserve\"" : ""}>`, `</${p}t>`);
	t.children = [{
		kind: "text",
		raw: escapeText(text),
		text,
		parent: t,
		start: -1
	}];
	return t;
}
/**
* Isolates the tags of one paragraph. Returns the anchor run of each tag. The pieces of the record
* describe the tree before this call; the tree is changed in place.
*/
function isolateTags(paragraph, tags, dialect = WML_DIALECT) {
	const isWml = isIn(dialect);
	const anchors = /* @__PURE__ */ new Map();
	if (tags.length === 0) return anchors;
	const replacements = new Replacements();
	const runs = /* @__PURE__ */ new Set();
	for (const piece of paragraph.pieces) if (piece.kind === "text") runs.add(piece.run);
	for (const run of runs) {
		const elements = run.children.filter((c) => c.kind === "element");
		const rPr = elements.find((c) => isWml(c, "rPr"));
		const content = elements.filter((c) => c !== rPr);
		if (content.length <= 1) continue;
		replacements.add(run, content.map((c) => runWith(run, rPr, [c])));
	}
	replacements.flush();
	const cuts = [...new Set(tags.flatMap((tag) => [tag.start, tag.end]))].sort((a, b) => a - b);
	/** The index of the first cut greater than `value`. */
	const firstAbove = (value) => {
		let lo = 0;
		let hi = cuts.length;
		while (lo < hi) {
			const mid = lo + hi >> 1;
			if (cuts[mid] <= value) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	};
	const groups = /* @__PURE__ */ new Map();
	for (const piece of paragraph.pieces) {
		if (piece.kind !== "text") continue;
		const g = groups.get(piece.t);
		if (g === void 0) groups.set(piece.t, {
			start: piece.textStart,
			end: piece.textEnd,
			text: piece.node.text
		});
		else {
			g.end = piece.textEnd;
			g.text += piece.node.text;
		}
	}
	const textRuns = [];
	for (const [t, g] of groups) {
		const run = t.parent;
		const points = [];
		for (let i = firstAbove(g.start); i < cuts.length && cuts[i] < g.end; i++) points.push(cuts[i]);
		if (points.length === 0) {
			textRuns.push({
				run,
				start: g.start,
				end: g.end
			});
			continue;
		}
		const rPr = run.children.find((c) => isWml(c, "rPr"));
		const edges = [
			g.start,
			...points,
			g.end
		];
		const parts = [];
		for (let i = 0; i < edges.length - 1; i++) {
			const from = edges[i];
			const to = edges[i + 1];
			const part = runWith(run, rPr, [textElement(t, g.text.slice(from - g.start, to - g.start), dialect)]);
			parts.push(part);
			textRuns.push({
				run: part,
				start: from,
				end: to
			});
		}
		replacements.add(run, parts);
	}
	replacements.flush();
	textRuns.sort((a, b) => a.start - b.start);
	/** The index of the first text run that starts at or after `value`. */
	const firstRunFrom = (value) => {
		let lo = 0;
		let hi = textRuns.length;
		while (lo < hi) {
			const mid = lo + hi >> 1;
			if (textRuns[mid].start < value) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	};
	for (const tag of tags) {
		const covered = [];
		for (let i = firstRunFrom(tag.start); i < textRuns.length && textRuns[i].end <= tag.end; i++) covered.push(textRuns[i]);
		const first = covered[0];
		if (first === void 0) continue;
		const t = first.run.children.find((c) => isWml(c, "t"));
		if (t !== void 0) {
			const fresh = textElement(t, tag.info.raw, dialect);
			fresh.parent = first.run;
			first.run.children[first.run.children.indexOf(t)] = fresh;
		}
		for (const other of covered.slice(1)) {
			const parent = other.run.parent;
			if (parent !== null) parent.children.splice(parent.children.indexOf(other.run), 1);
		}
		anchors.set(tag, first.run);
	}
	return anchors;
}
// src/docx/styles.ts
const STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";
const TRANSITIONAL_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
/** Reads the styles part of the main document; an empty map when there is none. */
function readStyles(pkg, main, limits) {
	const rel = pkg.relationships(main).find((r) => !r.external && (r.type === TRANSITIONAL_REL || r.type === STRICT_REL));
	const xml = rel === void 0 ? void 0 : pkg.readXml(rel.target);
	const out = /* @__PURE__ */ new Map();
	if (xml === void 0 || rel === void 0) return out;
	const root = parseXmlTree(xml, {
		part: rel.target,
		limits
	});
	for (const style of root.children) {
		if (!WML_NAMESPACES.has(style.ns) || style.local !== "style") continue;
		const id = style.attrs.get(`{${style.ns}}styleId`);
		const type = style.attrs.get(`{${style.ns}}type`) ?? "paragraph";
		const name = style.children.find((c) => WML_NAMESPACES.has(c.ns) && c.local === "name")?.attrs.get(`{${style.ns}}val`);
		if (id === void 0 || name === void 0) continue;
		const key = name.toLowerCase();
		const numId = (style.children.find((c) => WML_NAMESPACES.has(c.ns) && c.local === "pPr")?.children.find((c) => WML_NAMESPACES.has(c.ns) && c.local === "numPr"))?.children.find((c) => WML_NAMESPACES.has(c.ns) && c.local === "numId")?.attrs.get(`{${style.ns}}val`);
		if (!out.has(key)) out.set(key, numId === void 0 ? {
			id,
			type
		} : {
			id,
			type,
			numId
		});
	}
	return out;
}
/** A style of the given type by name, or undefined. */
function styleNamed(styles, name, type) {
	const style = styles.get(name.toLowerCase());
	return style !== void 0 && style.type === type ? style : void 0;
}
// src/docx/render.ts
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const WP = /* @__PURE__ */ new Set([TRANSITIONAL.wp, STRICT.wp]);
/** Elements left out of the second and later copies of repeated content (D-027). */
const NOT_IN_COPIES = /* @__PURE__ */ new Set([
	"bookmarkStart",
	"bookmarkEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"commentReference"
]);
const EMU_PER_INCH = 914400;
const EMU_PER_CM = 36e4;
const EMU_PER_TWIP = 635;
const MAX_IMAGE_CM = 1e3;
const isWml = (e, local) => WML_NAMESPACES.has(e.ns) && (local === void 0 || e.local === local);
function placement(tag, message, hint) {
	return new TvError("TV_E_SECTION_PLACEMENT", message, {
		location: tag.info.location,
		hint,
		details: { tag: tag.info.raw }
	});
}
/** The ancestors of `node`, from the root down to `node` itself. */
function ancestors(node) {
	const out = [];
	for (let n = node; n !== null; n = n.parent) out.push(n);
	return out.reverse();
}
/**
* Whether a row of a DrawingML table has a cell merged down: rowSpan gives the number of rows, which
* copies would not change. (In documents, w:vMerge marks each cell as the start or a continuation
* of a merge, so copies stay consistent and are allowed.)
*/
function hasVerticalMerge(row) {
	for (const cell of row.children) {
		if (cell.kind !== "element" || cell.local !== "tc") continue;
		if (Number(cell.attrs.get("rowSpan") ?? "1") > 1 || cell.attrs.get("vMerge") === "1" || cell.attrs.get("vMerge") === "true") return true;
	}
	return false;
}
/** Elements that must hold at least one paragraph (see StoryRenderer.keepParagraph). */
const PARAGRAPH_HOLDERS = /* @__PURE__ */ new Set([
	"tc",
	"hdr",
	"ftr",
	"footnote",
	"endnote",
	"comment",
	"txbxContent",
	"txBody"
]);
/** Where a section repeats, from the anchors of its opening and closing tags. */
function unitOf(section, anchors, isText) {
	const open = anchors.get(section.open);
	const close = anchors.get(section.close);
	const a = ancestors(open);
	const b = ancestors(close);
	let i = 0;
	while (i < a.length && i < b.length && a[i] === b[i]) i++;
	const container = a[i - 1];
	const x = a[i];
	const y = b[i];
	const from = container.children.indexOf(x);
	const to = container.children.indexOf(y);
	const blockish = (e) => isText(e, "p") || isText(e, "tbl") || isText(e, "sdt");
	const where = `{${section.open.info.kind === "inverted" ? "^" : "#"}${section.open.info.path}}`;
	if (isText(x, "tr") && isText(y, "tr")) return {
		tree: section,
		container,
		from,
		to,
		inclusive: true
	};
	if (isText(x, "tc") && isText(y, "tc")) {
		const table = container.parent;
		const row = table.children.indexOf(container);
		return {
			tree: section,
			container: table,
			from: row,
			to: row,
			inclusive: true
		};
	}
	if (blockish(x) || blockish(y)) {
		if (isText(x, "sdt") && x !== open || isText(y, "sdt") && y !== close) return placement(section.open, `The section "${where}" has one tag inside a content control and the other outside it.`, `Put ${where} and {/${section.open.info.path}} both inside the content control or both outside it.`);
		if (isText(x, "tbl") && x !== open || isText(y, "tbl") && y !== close) return placement(section.open, `The section "${where}" starts and ends in different tables, or one of its tags is in a table and the other outside it.`, `Put ${where} and {/${section.open.info.path}} in the same table row, in two rows of the same table, or both outside the table.`);
		const alone = (tag, element) => tag.paragraph.element === element && tag.paragraph.text.trim() === tag.info.raw;
		if (!alone(section.open, x) || !alone(section.close, y)) return placement(section.open, `The section "${where}" spans several paragraphs, so its opening and closing tags must each be alone in their paragraph.`, `Put ${where} and {/${section.open.info.path}} each on a line of their own, or keep the whole section within one paragraph.`);
		return {
			tree: section,
			container,
			from,
			to,
			inclusive: false
		};
	}
	if (x === open && y === close) return {
		tree: section,
		container,
		from,
		to,
		inclusive: false
	};
	if ((x.local === "sp" || x.local === "graphicFrame") && (y.local === "sp" || y.local === "graphicFrame")) return placement(section.open, `The tags of the section "${where}" are in different shapes of the slide.`, `To repeat slides, put ${where} alone in a shape and {/${section.open.info.path}} alone in another; otherwise keep the whole section within one shape.`);
	return placement(section.open, `The tags of the section "${where}" are in different parts of a paragraph (for example one inside a hyperlink and one outside).`, "Put both tags at the same level: both inside the hyperlink or field, or both outside it.");
}
/** Collects the sections of a tree, outermost first. */
function sectionsOf(children, out = []) {
	for (const child of children) if ("section" in child) {
		out.push(child.section);
		sectionsOf(child.section.children, out);
	}
	return out;
}
/** One story being rendered. Presentations extend it (../pptx/render.ts). */
var StoryRenderer = class extends TagEvaluator {
	story;
	ctx;
	anchors = /* @__PURE__ */ new Map();
	valueAnchors = /* @__PURE__ */ new Map();
	imageAnchors = /* @__PURE__ */ new Map();
	/** Markdown tags that share their paragraph with other content: rendered inline. */
	markdownAnchors = /* @__PURE__ */ new Map();
	/** Paragraphs that hold only a Markdown tag: replaced by the paragraphs of the Markdown. */
	markdownParagraphs = /* @__PURE__ */ new Map();
	droppedAnchors = /* @__PURE__ */ new Set();
	units = /* @__PURE__ */ new Map();
	prunedAlternates = /* @__PURE__ */ new Set();
	sizes = /* @__PURE__ */ new Map();
	/** Elements left out of the result, such as the shapes that hold the tags of a slide section. */
	droppedElements = /* @__PURE__ */ new Set();
	/** The part being written: the story's own part, or a copy of it (a repeated slide). */
	targetPart = "";
	constructor(story, ctx) {
		super(ctx);
		this.story = story;
		this.ctx = ctx;
	}
	/** Whether an element is of the text vocabulary of this story (WordprocessingML here). */
	isText(e, local) {
		return isWml(e, local);
	}
	/** Whether an element is left out of the second and later copies of repeated content. */
	dropInCopy(e) {
		return isWml(e) && NOT_IN_COPIES.has(e.local);
	}
	/** Moves each tag of a paragraph into a run of its own. */
	isolate(paragraph, tags) {
		return isolateTags(paragraph, tags);
	}
	/** The scopes of each repetition of a section, for sections that repeat slides. */
	sectionScopes(section, scopes) {
		return this.iterations(section, scopes);
	}
	/** Leaves `e` out of every rendering. */
	drop(e) {
		this.droppedElements.add(e);
	}
	/** Normalises the template and finds where each section repeats. Returns problems of placement. */
	prepare() {
		const byParagraph = /* @__PURE__ */ new Map();
		for (const tag of this.story.tags) {
			const list = byParagraph.get(tag.paragraph) ?? [];
			list.push(tag);
			byParagraph.set(tag.paragraph, list);
		}
		for (const paragraph of this.story.paragraphs) {
			const tags = byParagraph.get(paragraph);
			if (tags === void 0) continue;
			for (const [tag, anchor] of this.isolate(paragraph, tags)) this.anchors.set(tag, anchor);
			for (const alternate of paragraph.alternates) this.prunedAlternates.add(alternate);
		}
		for (const tag of this.story.tags) {
			const anchor = this.anchors.get(tag);
			if (anchor === void 0) continue;
			if (tag.info.kind === "value") this.valueAnchors.set(anchor, tag);
			else if (tag.info.kind === "image") this.imageAnchors.set(anchor, tag);
			else if (tag.info.kind === "markdown") {
				if (tag.paragraph.text.trim() === tag.info.raw) this.markdownParagraphs.set(tag.paragraph.element, tag);
				else this.markdownAnchors.set(anchor, tag);
			}
		}
		const problems = [];
		for (const section of sectionsOf(this.story.tree)) {
			const unit = unitOf(section, this.anchors, (e, local) => this.isText(e, local));
			if (unit instanceof TvError) {
				problems.push(unit);
				continue;
			}
			if (unit.inclusive && this.isText(unit.container.children[unit.from], "tr")) {
				if (unit.container.children.slice(unit.from, unit.to + 1).some((row) => row.kind === "element" && hasVerticalMerge(row))) {
					problems.push(placement(section.open, `The rows of the section "{#${section.open.info.path}}" have cells merged down across rows, which the copies would overlap.`, "Unmerge the cells of the rows that repeat, or keep merged cells out of them."));
					continue;
				}
			}
			const list = this.units.get(unit.container) ?? [];
			if (unit.inclusive && list.some((u) => u.inclusive && u.from === unit.from && u.to === unit.to)) {
				problems.push(placement(section.open, `The section "{#${section.open.info.path}}" repeats the same table rows as the section around it.`, "Move one of the two sections so that each repeats different rows, or join them into one."));
				continue;
			}
			list.push(unit);
			this.units.set(unit.container, list);
			if (unit.inclusive) {
				this.droppedAnchors.add(this.anchors.get(section.open));
				this.droppedAnchors.add(this.anchors.get(section.close));
			}
		}
		for (const list of this.units.values()) list.sort((p, q) => p.from - q.from || q.to - q.from - (p.to - p.from));
		return problems;
	}
	render(scopes, copy = false, part = this.story.part) {
		this.targetPart = part;
		const { doc } = this.story;
		const root = this.renderElement(doc.root, scopes, /* @__PURE__ */ new Set(), copy);
		return {
			prolog: doc.prolog,
			root,
			epilog: doc.epilog
		};
	}
	count(node) {
		let size = this.sizes.get(node);
		if (size === void 0) {
			const out = [];
			serializeInto(node, out);
			size = out.reduce((n, s) => n + s.length, 0);
			this.sizes.set(node, size);
		}
		this.ctx.output += size;
		if (this.ctx.output > this.ctx.limits.maxOutputBytes) throw limitError("maxOutputBytes", this.ctx.limits.maxOutputBytes, `The result would be larger than ${String(this.ctx.limits.maxOutputBytes)} characters of XML.`, { part: this.story.part });
	}
	/**
	* Counts text built for a value against maxOutputBytes while it is built (Markdown can grow much
	* larger than its value, D-071). The caller gives the amount back before the result is counted.
	*/
	charge(size) {
		this.ctx.output += size;
		if (this.ctx.output > this.ctx.limits.maxOutputBytes) throw limitError("maxOutputBytes", this.ctx.limits.maxOutputBytes, `The result would be larger than ${String(this.ctx.limits.maxOutputBytes)} characters of XML.`, { part: this.story.part });
	}
	warn(code, message, hint, tag) {
		if (this.ctx.warnings.some((w) => w.code === code && w.message === message)) return;
		this.ctx.warnings.push({
			code,
			message,
			location: tag.info.location,
			hint
		});
	}
	/** A style of the template by name, with a warning when the template has none. */
	style(name, type, tag) {
		const style = styleNamed(this.ctx.styles, name, type);
		if (style === void 0) this.warn("TV_W_STYLE_MISSING", `The template has no ${type} style "${name}", so ${type === "paragraph" ? "these paragraphs keep the style of the tag’s paragraph" : "this text keeps the formatting of the tag"}.`, `Add a ${type} style named "${name}" to the template (for example by using it once in Word or LibreOffice) to control how it looks.`, tag);
		return style?.id;
	}
	/** A hyperlink run in place of the anchor. */
	linkNodes(anchor, link, tag, style = {}) {
		const p = prefixOf(anchor.qname);
		const rStyle = this.style("Hyperlink", "character", tag);
		const run = runXml(anchor, runProperties(p, runPropertiesOf(anchor), rStyle === void 0 ? style : {
			...style,
			rStyle
		}), link.text);
		if (anchor.parent !== null && ancestors(anchor).some((a) => isWml(a, "hyperlink"))) return [rawXml(run)];
		return [rawXml(hyperlinkXml(p, this.ctx.resources.relationship(this.targetPart, this.ctx.resources.vocabulary.hyperlinkRel, link.url, true), run, this.ctx.resources.vocabulary))];
	}
	/** The bytes, size and description of the value of an image tag, or null with a problem recorded. */
	imageValue(tag, value) {
		const bad = (message, hint) => {
			this.invalid(tag, `${message} (in "${clip(tag.info.raw)}")`, hint);
			return null;
		};
		let source = value;
		let width;
		let height;
		let alt = "";
		if (source !== null && typeof source === "object" && !(source instanceof Uint8Array) && !(source instanceof ArrayBuffer)) {
			source = ownValue(value, "data");
			const w = ownValue(value, "width");
			const h = ownValue(value, "height");
			const a = ownValue(value, "alt");
			for (const [name, x] of [["width", w], ["height", h]]) if (x !== void 0 && !(typeof x === "number" && Number.isFinite(x) && x > 0 && x <= MAX_IMAGE_CM)) return bad(`The ${name} of an image must be a number of centimetres above 0 and at most ${String(MAX_IMAGE_CM)}.`, `Pass for example { data, ${name}: 5 }.`);
			if (a !== void 0 && typeof a !== "string") return bad("The alt text of an image must be text.", "Pass for example { data, alt: \"Company logo\" }.");
			width = w;
			height = h;
			alt = a ?? "";
			if (firstInvalidXmlChar(alt) !== -1) return bad("The alt text of the image contains a character that a document cannot hold.", "Remove control characters from the alt text.");
		}
		let bytes;
		if (source instanceof Uint8Array) bytes = source;
		else if (source instanceof ArrayBuffer) bytes = new Uint8Array(source);
		else if (typeof source === "string") {
			const lower = source.slice(0, 16).toLowerCase();
			const decoded = lower.startsWith("data:") ? this.ctx.resources.dataUrls.get(source) : void 0;
			if (decoded !== void 0) bytes = decoded;
			else if (lower.startsWith("data:")) {
				const m = /^data:image\/(?:png|jpeg|jpg|gif);base64,/i.exec(source.slice(0, 40));
				if (m === null) return bad("The data URL of the image is not a base64 PNG, JPEG or GIF.", "Pass for example \"data:image/png;base64,iVBORw0…\".");
				const b64 = source.slice(m[0].length).replace(/\s+/g, "");
				if (b64.length * 3 / 4 > this.ctx.limits.maxImageBytes) throw limitError("maxImageBytes", this.ctx.limits.maxImageBytes, `An image is larger than ${String(this.ctx.limits.maxImageBytes)} bytes.`, tag.info.location);
				if (!/^[A-Za-z0-9+/]*={0,2}$/.test(b64) || b64.replace(/=+$/, "").length % 4 === 1) return bad("The data URL of the image is not valid base64.", "Encode the image bytes with base64.");
				let binary;
				try {
					binary = atob(b64);
				} catch {
					return bad("The data URL of the image is not valid base64.", "Encode the image bytes with base64.");
				}
				bytes = new Uint8Array(binary.length);
				for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
				this.ctx.resources.dataUrls.set(source, bytes);
			} else if (lower.startsWith("http://") || lower.startsWith("https://")) {
				const fetched = this.ctx.fetched.get(source);
				if (fetched === void 0) {
					if (this.ctx.collect !== void 0) {
						this.ctx.collect.add(source);
						return null;
					}
					return bad(`The image "${clip(source)}" is an address, and tplvane does not download anything by itself.`, "Pass the bytes or a data URL, or use renderAsync with a fetchImage function that you control.");
				}
				bytes = fetched;
			} else return bad("The value of an image tag is text, but not a data URL or an http(s) address.", "Pass the bytes (Uint8Array), a data URL, or { data, width, height, alt }.");
		} else return bad("The value of an image tag must be the bytes, a data URL, an address or { data, width, height, alt }.", "Pass for example the bytes of a PNG file as a Uint8Array.");
		if (bytes.length > this.ctx.limits.maxImageBytes) throw limitError("maxImageBytes", this.ctx.limits.maxImageBytes, `An image is larger than ${String(this.ctx.limits.maxImageBytes)} bytes.`, tag.info.location);
		return {
			bytes,
			...width === void 0 ? {} : { width },
			...height === void 0 ? {} : { height },
			alt
		};
	}
	/** Width available to an image at the anchor: its table cell, else the text column, in EMU. */
	availableWidth(anchor) {
		for (let e = anchor.parent; e !== null; e = e.parent) {
			if (!isWml(e, "tc")) continue;
			const tcW = e.children.find((c) => c.kind === "element" && isWml(c, "tcPr"))?.children.find((c) => c.kind === "element" && isWml(c, "tcW"));
			const w = Number(tcW?.attrs.get(`{${tcW.ns}}w`));
			if ((tcW?.attrs.get(`{${tcW.ns}}type`) ?? "dxa") === "dxa" && Number.isFinite(w) && w > 0) return Math.min(w * EMU_PER_TWIP, this.ctx.textWidth);
			break;
		}
		return this.ctx.textWidth;
	}
	/** An inline picture in place of the anchor. */
	imageNodes(anchor, tag, scopes) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return [];
		if (value === void 0 || value === null) {
			if (value === void 0) this.missing(tag, tag.info.path, scopes);
			return [];
		}
		const image = this.imageValue(tag, value);
		if (image === null) return [];
		const info = imageInfo(image.bytes);
		if (typeof info === "string") {
			this.invalid(tag, `${info} (in "${clip(tag.info.raw)}")`, "Pass a PNG, JPEG or GIF image.");
			return [];
		}
		if (info.width * info.height > this.ctx.limits.maxImagePixels) throw limitError("maxImagePixels", this.ctx.limits.maxImagePixels, `An image has ${String(info.width)} × ${String(info.height)} pixels, more than the limit.`, tag.info.location);
		const ratio = info.height / info.dpiY / (info.width / info.dpiX);
		let cx;
		let cy;
		if (image.width !== void 0 || image.height !== void 0) {
			cx = image.width !== void 0 ? image.width * EMU_PER_CM : image.height * EMU_PER_CM / ratio;
			cy = image.height !== void 0 ? image.height * EMU_PER_CM : cx * ratio;
		} else {
			cx = info.width / info.dpiX * EMU_PER_INCH;
			cy = info.height / info.dpiY * EMU_PER_INCH;
			const max = this.availableWidth(anchor);
			if (cx > max) {
				cy = cy * (max / cx);
				cx = max;
			}
		}
		const media = this.ctx.resources.image(image.bytes, info);
		const rId = this.ctx.resources.imageRelationship(this.targetPart, media);
		const id = this.ctx.nextDrawingId++;
		const p = prefixOf(anchor.qname);
		const largest = MAX_IMAGE_CM * EMU_PER_CM;
		if (cx > largest || cy > largest) {
			const scale = largest / Math.max(cx, cy);
			cx *= scale;
			cy *= scale;
		}
		return [rawXml(pictureRunXml(anchor, runProperties(p, runPropertiesOf(anchor)), {
			rId,
			cx: Math.max(1, Math.round(cx)),
			cy: Math.max(1, Math.round(cy)),
			id,
			name: `Image ${String(id)}`,
			alt: image.alt
		}, this.ctx.resources.vocabulary))];
	}
	/** The runs of inline Markdown, with the formatting of the anchor. */
	inlineRuns(anchor, tag, content) {
		const p = prefixOf(anchor.qname);
		const base = runPropertiesOf(anchor);
		const out = [];
		let charged = 0;
		const push = (xml) => {
			this.charge(xml.length);
			charged += xml.length;
			out.push(xml);
		};
		for (const inline of content) {
			const style = {
				...inline.bold ? { bold: true } : {},
				...inline.italic ? { italic: true } : {}
			};
			const code = inline.code ? this.style("HTML Code", "character", tag) : void 0;
			const runStyle = code === void 0 ? style : {
				...style,
				rStyle: code
			};
			if (inline.link !== void 0) {
				if (checkLinkUrl(inline.link) !== null) {
					push(runXml(anchor, runProperties(p, base, runStyle), inline.text, inline.breakBefore === true));
					continue;
				}
				if (inline.breakBefore === true) push(runXml(anchor, runProperties(p, base), "", true));
				const [node] = this.linkNodes(anchor, new LinkValue(inline.link, inline.text), tag, runStyle);
				push(node.raw);
				continue;
			}
			push(runXml(anchor, runProperties(p, base, runStyle), inline.text, inline.breakBefore === true));
		}
		this.ctx.output -= charged;
		return out.join("");
	}
	/** Markdown in place of the anchor (inline), or of the paragraph that holds only the tag (blocks). */
	markdownNodes(node, tag, scopes, blocks) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return [];
		if (value === void 0) {
			this.missing(tag, tag.info.path, scopes);
			return [];
		}
		if (value === null || value === "") return [];
		if (typeof value !== "string") {
			this.invalid(tag, `The value of "${clip(tag.info.raw)}" must be Markdown text.`, "Pass the Markdown as a string.");
			return [];
		}
		if (value.length > this.ctx.limits.maxMarkdownLength) throw limitError("maxMarkdownLength", this.ctx.limits.maxMarkdownLength, `A Markdown value is longer than ${String(this.ctx.limits.maxMarkdownLength)} characters.`, tag.info.location);
		if (this.checkText(tag, value) === null) return [];
		const parsed = parseMarkdown(value);
		if (!blocks) {
			if (parsed.length > 1 || parsed[0] !== void 0 && parsed[0].kind !== "paragraph") {
				this.invalid(tag, `The Markdown of "${clip(tag.info.raw)}" has headings, lists or several paragraphs, but the tag shares its paragraph with other text.`, `Put ${tag.info.raw} alone in a paragraph to get headings, lists and paragraphs.`);
				return [];
			}
			return parsed[0] === void 0 ? [] : [rawXml(this.inlineRuns(node, tag, parsed[0].content))];
		}
		const anchor = this.anchors.get(tag);
		const p = prefixOf(node.qname);
		const numbers = [
			0,
			0,
			0
		];
		const kinds = [
			void 0,
			void 0,
			void 0
		];
		const out = [];
		let charged = 0;
		let list;
		let start = 1;
		for (const block of parsed) {
			let styleId;
			let prefix = "";
			let numbering;
			if (block.kind === "heading") styleId = this.style(`heading ${String(block.level)}`, "paragraph", tag);
			if (block.kind === "item") {
				const name = `${block.ordered ? "List Number" : "List Bullet"}${block.level === 0 ? "" : ` ${String(block.level + 1)}`}`;
				styleId = this.style(name, "paragraph", tag);
				if (kinds[block.level] !== block.ordered) numbers[block.level] = 0;
				kinds[block.level] = block.ordered;
				numbers[block.level] = (numbers[block.level] ?? 0) + 1;
				for (let l = block.level + 1; l < numbers.length; l++) {
					numbers[l] = 0;
					kinds[l] = void 0;
				}
				if (block.level === 0 && numbers[0] === 1) start = block.number;
				if (styleId === void 0) prefix = `${"    ".repeat(block.level)}${block.ordered ? `${String(block.level === 0 ? start + (numbers[0] ?? 1) - 1 : numbers[block.level])}. ` : "• "}`;
				if (block.level === 0) {
					if (!block.ordered) list = void 0;
					else if (numbers[0] === 1) {
						const numId = styleNamed(this.ctx.styles, name, "paragraph")?.numId;
						list = numId === void 0 ? void 0 : this.ctx.numbering.restart(numId, block.number);
					}
				}
				if (block.ordered && block.level === 0 && list !== void 0) numbering = {
					numId: list,
					level: 0
				};
			} else {
				numbers.fill(0);
				kinds.fill(void 0);
				list = void 0;
			}
			const content = (prefix === "" ? "" : runXml(anchor, runProperties(p, runPropertiesOf(anchor)), prefix)) + this.inlineRuns(anchor, tag, block.content);
			const paragraph = paragraphXml(node, paragraphProperties(p, node, styleId, numbering), content);
			this.charge(paragraph.length);
			charged += paragraph.length;
			out.push(paragraph);
		}
		this.ctx.output -= charged;
		return [rawXml(out.join(""))];
	}
	/** The runs that show `text` in place of the anchor: line breaks become w:br, tabs w:tab. */
	valueRuns(anchor, text) {
		if (text === "") return [];
		const p = prefixOf(anchor.qname);
		const rPr = anchor.children.find((c) => c.kind === "element" && isWml(c, "rPr"));
		const t = anchor.children.find((c) => c.kind === "element" && isWml(c, "t"));
		const run = makeElement(anchor, anchor.open, anchor.close, anchor.attrs);
		const children = rPr === void 0 ? [] : [rPr];
		text.replace(/\r\n?/g, "\n").split("\n").forEach((line, i) => {
			if (i > 0) children.push(makeElement({
				ns: anchor.ns,
				local: "br",
				qname: `${p}br`
			}, `<${p}br/>`, ""));
			line.split("	").forEach((piece, j) => {
				if (j > 0) children.push(makeElement({
					ns: anchor.ns,
					local: "tab",
					qname: `${p}tab`
				}, `<${p}tab/>`, ""));
				if (piece !== "") children.push(textElement(t, piece));
			});
		});
		run.children = children;
		return [run];
	}
	renderElement(e, scopes, active, copy) {
		if (this.valueAnchors.get(e) !== void 0) return null;
		if (this.droppedAnchors.has(e) || this.droppedElements.has(e)) return null;
		if (copy && this.dropInCopy(e)) return null;
		if (copy && isWml(e) && (e.local === "footnoteReference" || e.local === "endnoteReference")) throw new TvError("TV_E_NOT_SUPPORTED", `A ${e.local === "footnoteReference" ? "footnote" : "endnote"} is in content that repeats; notes are not copied with repeated content.`, {
			location: { part: this.story.part },
			hint: "Move the note out of the repeated rows or paragraphs, or write its text in the repeated content itself.",
			details: { element: e.local }
		});
		if (!(this.units.get(e) !== void 0 || this.containsWork(e)) && !copy) {
			this.count(e);
			return e;
		}
		let open = e.open;
		if (copy) {
			if (e.attrs.has(`{${W14}}paraId`) || e.attrs.has(`{${W14}}textId`)) open = open.replace(/\s+[\w.-]+:(?:paraId|textId)="[^"]*"/g, "");
			if (WP.has(e.ns) && e.local === "docPr" && e.attrs.has("id")) open = open.replace(/(\sid=")[^"]*(")/, `$1${String(this.ctx.nextDrawingId++)}$2`);
		}
		const out = makeElement(e, open, e.close, e.attrs);
		this.ctx.output += open.length + e.close.length;
		out.children = this.renderChildren(e, 0, e.children.length - 1, scopes, active, copy);
		this.keepParagraph(e, out);
		return out;
	}
	/**
	* A table cell, header, footer, note, comment, text box or text body must hold a paragraph
	* (ECMA-376 Part 1 §17.4.66; a:txBody in DrawingML): when a section shown no time or an empty
	* Markdown value took all of them, an empty one is put back, with the properties of the first
	* paragraph of the template (D-071).
	*/
	keepParagraph(template, out) {
		if (!PARAGRAPH_HOLDERS.has(template.local)) return;
		const isBlock = (c) => c.kind === "element" && (c.local === "p" || c.local === "tbl");
		if (out.children.some((c) => isBlock(c) || c.kind === "text" && c.raw.includes("<") && /<(?:\w+:)?(?:p|tbl)[\s>/]/.test(c.raw))) return;
		const model = template.children.find((c) => c.kind === "element" && c.local === "p");
		if (model === void 0) return;
		const pPr = model.children.find((c) => c.kind === "element" && c.local === "pPr");
		const props = pPr === void 0 ? "" : rawXmlText(pPr);
		const empty = makeElement(model, props === "" ? `<${model.qname}/>` : `<${model.qname}>`, props === "" ? "" : `</${model.qname}>`);
		if (props !== "") empty.children = [{
			kind: "text",
			raw: props,
			text: "",
			parent: empty,
			start: -1
		}];
		out.children.push(empty);
		this.ctx.output += props.length + model.qname.length * 2 + 5;
	}
	/** Whether an element holds anchors or repeated sections below it (computed once per element). */
	work = /* @__PURE__ */ new Map();
	containsWork(e) {
		let known = this.work.get(e);
		if (known !== void 0) return known;
		known = this.valueAnchors.has(e) || this.imageAnchors.has(e) || this.markdownAnchors.has(e) || this.markdownParagraphs.has(e) || this.droppedAnchors.has(e) || this.droppedElements.has(e) || this.units.has(e) || this.prunedAlternates.has(e) || e.children.some((c) => c.kind === "element" && this.containsWork(c));
		this.work.set(e, known);
		return known;
	}
	/** The units of a container by the index they start at, built once: a lookup per child stays constant (D-071). */
	starts = /* @__PURE__ */ new Map();
	unitsByStart(container) {
		const units = this.units.get(container);
		if (units === void 0) return void 0;
		let byStart = this.starts.get(container);
		if (byStart === void 0) {
			byStart = /* @__PURE__ */ new Map();
			for (const u of units) {
				const list = byStart.get(u.from);
				if (list === void 0) byStart.set(u.from, [u]);
				else list.push(u);
			}
			this.starts.set(container, byStart);
		}
		return byStart;
	}
	renderChildren(container, from, to, scopes, active, copy) {
		const out = [];
		const byStart = this.unitsByStart(container);
		const pruned = this.prunedAlternates.has(container) ? selectedBranch(container) : void 0;
		for (let i = from; i <= to;) {
			const unit = byStart?.get(i)?.find((u) => u.to <= to && !active.has(u));
			if (unit !== void 0) {
				const repeats = this.iterations(unit.tree, scopes);
				const inner = new Set(active).add(unit);
				repeats.forEach((itemScopes, k) => {
					if (++this.ctx.iterations > this.ctx.limits.maxIterations) throw limitError("maxIterations", this.ctx.limits.maxIterations, `Sections repeat their content more than ${String(this.ctx.limits.maxIterations)} times in all.`, unit.tree.open.info.location);
					const a = unit.inclusive ? unit.from : unit.from + 1;
					const b = unit.inclusive ? unit.to : unit.to - 1;
					out.push(...this.renderChildren(container, a, b, itemScopes, inner, copy || k > 0));
				});
				i = unit.to + 1;
				continue;
			}
			const child = container.children[i];
			i++;
			if (child.kind === "text") {
				this.count(child);
				out.push(child);
				continue;
			}
			if (pruned !== void 0 && child.ns === "http://schemas.openxmlformats.org/markup-compatibility/2006" && child !== pruned) continue;
			const valueTag = this.valueAnchors.get(child);
			if (valueTag !== void 0) {
				const value = this.valueText(valueTag, scopes);
				const nodes = value instanceof LinkValue ? this.linkNodes(child, value, valueTag) : this.valueRuns(child, value ?? "");
				for (const run of nodes) {
					this.count(run);
					out.push(run);
				}
				continue;
			}
			const imageTag = this.imageAnchors.get(child);
			if (imageTag !== void 0) {
				for (const node of this.imageNodes(child, imageTag, scopes)) {
					this.count(node);
					out.push(node);
				}
				continue;
			}
			const markdownTag = this.markdownAnchors.get(child) ?? this.markdownParagraphs.get(child);
			if (markdownTag !== void 0) {
				for (const node of this.markdownNodes(child, markdownTag, scopes, this.markdownParagraphs.has(child))) {
					this.count(node);
					out.push(node);
				}
				continue;
			}
			const rendered = this.renderElement(child, scopes, active, copy);
			if (rendered !== null) out.push(rendered);
		}
		return out;
	}
};
/** Prepares one story; problems of data found while rendering are added to `ctx.problems`. */
function prepareStory(story, ctx, create = (s, c) => new StoryRenderer(s, c), dropped = []) {
	const renderer = create(story, ctx);
	for (const e of dropped) renderer.drop(e);
	return {
		problems: renderer.prepare(),
		renderer,
		render(scopes, copy = false, part = story.part) {
			const doc = renderer.render(scopes, copy, part);
			const out = [doc.prolog];
			serializeInto(doc.root, out);
			out.push(doc.epilog);
			return out.join("");
		}
	};
}
/** The largest wp:docPr id in a tree, so that copies of drawings get new ones above it. */
function maxDrawingId(doc) {
	let max = 0;
	const work = [doc.root];
	while (work.length > 0) {
		const node = work.pop();
		if (node.kind !== "element") continue;
		if (WP.has(node.ns) && node.local === "docPr") {
			const id = Number(node.attrs.get("id"));
			if (Number.isSafeInteger(id) && id > max) max = id;
		}
		for (const c of node.children) work.push(c);
	}
	return max;
}
/**
* The width of the text column of the main document in EMU: the page width minus the left and
* right margins of the last w:sectPr of the body (F-015), or 16 cm when they are not given.
*/
function textWidthOf(doc) {
	const sectPr = [...doc.root.children.find((c) => c.kind === "element" && isWml(c, "body"))?.children ?? []].reverse().find((c) => c.kind === "element" && isWml(c, "sectPr"));
	const child = (local) => sectPr?.children.find((c) => c.kind === "element" && isWml(c, local));
	const pgSz = child("pgSz");
	const pgMar = child("pgMar");
	const num = (e, name) => Number(e?.attrs.get(`{${e.ns}}${name}`));
	const width = num(pgSz, "w") - (num(pgMar, "left") || num(pgMar, "start") || 0) - (num(pgMar, "right") || num(pgMar, "end") || 0);
	return Number.isFinite(width) && width > 0 ? width * EMU_PER_TWIP : 16 * EMU_PER_CM;
}
// src/docx/resources.ts
/** The target of a relationship from `from` to `to`, relative to the folder of `from` (as "../media/x.png"). */
function relativeTarget(from, to) {
	const a = from.split("/").slice(1, -1);
	const b = to.split("/").slice(1);
	let i = 0;
	while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
	return [...a.slice(i).map(() => ".."), ...b.slice(i)].join("/");
}
/** A Relationship element. */
function relationshipXml(id, type, target, external) {
	return `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(type)}" Target="${escapeAttr(target)}"${external ? " TargetMode=\"External\"" : ""}/>`;
}
/** The Id of a Relationship element written by relationshipXml. */
const idOf = (xml) => (/\bId="([^"]*)"/.exec(xml)?.[1] ?? "").replace(/&(lt|gt|quot|apos|amp);/g, (_, e) => ({
	lt: "<",
	gt: ">",
	quot: "\"",
	apos: "'",
	amp: "&"
})[e]);
/** The parts and relationships one render adds. */
var Resources = class {
	pkg;
	zip;
	limits;
	mediaFolder;
	vocabulary;
	rels = /* @__PURE__ */ new Map();
	media = /* @__PURE__ */ new Map();
	mediaParts = [];
	usedNames = /* @__PURE__ */ new Set();
	overrides = /* @__PURE__ */ new Map();
	removedOverrides = /* @__PURE__ */ new Set();
	removedParts = /* @__PURE__ */ new Set();
	next = 1;
	constructor(pkg, zip, limits, mediaFolder = "/word/media/", vocabulary = TRANSITIONAL) {
		this.pkg = pkg;
		this.zip = zip;
		this.limits = limits;
		this.mediaFolder = mediaFolder;
		this.vocabulary = vocabulary;
	}
	partRels(part) {
		const key = partKey(part);
		let r = this.rels.get(key);
		if (r === void 0) {
			r = {
				part,
				removed: /* @__PURE__ */ new Set(),
				ids: new Set(this.pkg.relationships(part).map((x) => x.id)),
				added: [],
				byKey: /* @__PURE__ */ new Map()
			};
			this.rels.set(key, r);
		}
		return r;
	}
	/** A new part name `<folder><stem>N.<ext>` that is not in the package and not given out before. */
	newPartName(folder, stem, ext) {
		let name;
		let n = this.usedNames.size + 1;
		do
			name = `${folder}${stem}${String(n++)}.${ext}`;
		while (this.pkg.hasPart(name) || this.usedNames.has(partKey(name)));
		this.usedNames.add(partKey(name));
		return name;
	}
	/** Starts the relationships of a new part `part` from the XML of another part's relationships. */
	seedRelationships(part, entries, ids) {
		const r = {
			part,
			base: entries,
			removed: /* @__PURE__ */ new Set(),
			ids: new Set(ids),
			added: [],
			byKey: /* @__PURE__ */ new Map()
		};
		this.rels.set(partKey(part), r);
	}
	/** Removes a relationship of `part` by Id. */
	removeRelationship(part, id) {
		this.partRels(part).removed.add(id);
	}
	/** Adds an Override to the Media Types stream for a new part. */
	addOverride(partName, contentType) {
		this.overrides.set(partName, contentType);
	}
	/** Leaves a part, its Relationships part and its Override out of the result. */
	removePart(partName) {
		this.removedParts.add(partKey(partName));
		this.removedParts.add(partKey(relationshipsPartName(partName)));
		this.removedOverrides.add(partKey(partName));
	}
	/** The parts left out of the result, by partKey. */
	get removed() {
		return this.removedParts;
	}
	/** A relationship from `part`, reused for the same target; returns its Id. */
	relationship(part, type, target, external) {
		const r = this.partRels(part);
		const key = `${type} ${target}`;
		const known = r.byKey.get(key);
		if (known !== void 0) return known;
		let id;
		do
			id = `rIdDv${String(this.next++)}`;
		while (r.ids.has(id));
		r.ids.add(id);
		r.byKey.set(key, id);
		r.added.push(`<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(type)}" Target="${escapeAttr(target)}"${external ? " TargetMode=\"External\"" : ""}/>`);
		return id;
	}
	/** The media part for these image bytes, added once per distinct image. */
	/** Media parts by the identity of the bytes, so that a value repeated in a loop costs nothing more (D-062). */
	byIdentity = /* @__PURE__ */ new WeakMap();
	/** Decoded data URLs by their text, for the same reason. */
	dataUrls = /* @__PURE__ */ new Map();
	image(bytes, info) {
		const known = this.byIdentity.get(bytes);
		if (known !== void 0) return known;
		const name = this.imageByContent(bytes, info);
		this.byIdentity.set(bytes, name);
		return name;
	}
	imageByContent(bytes, info) {
		const key = `${info.extension}:${String(bytes.length)}:${String(crc32(bytes))}`;
		const list = this.media.get(key) ?? [];
		const same = list.find((m) => m.bytes.length === bytes.length && m.bytes.every((b, i) => b === bytes[i]));
		if (same !== void 0) return same.partName;
		if (this.mediaParts.length >= this.limits.maxImages) throw limitError("maxImages", this.limits.maxImages, `The document would get more than ${String(this.limits.maxImages)} different images.`);
		const name = this.newPartName(this.mediaFolder, "tplvane", info.extension);
		const media = {
			partName: name,
			bytes,
			info
		};
		list.push(media);
		this.media.set(key, list);
		this.mediaParts.push(media);
		return name;
	}
	/** A relationship from `part` to the media part `media`, as a target relative to the part's folder. */
	imageRelationship(part, media) {
		const target = relativeTarget(part, media);
		return this.relationship(part, this.vocabulary.imageRel, target, false);
	}
	/**
	* The new and changed parts: media, Relationships parts (existing ones with the new entries
	* added, missing ones created) and the Media Types stream with a Default for each new extension.
	* Changed parts are keyed by partKey, like the parts a render rewrites.
	*/
	parts() {
		const added = /* @__PURE__ */ new Map();
		const changed = /* @__PURE__ */ new Map();
		for (const m of this.mediaParts) added.set(m.partName, m.bytes);
		for (const r of this.rels.values()) {
			if (r.added.length === 0 && r.removed.size === 0 && r.base === void 0) continue;
			const relsPart = relationshipsPartName(r.part);
			const kept = r.base ?? this.pkg.relationships(r.part).filter((x) => !r.removed.has(x.id)).map((x) => relationshipXml(x.id, x.type, x.external ? x.target : relativeTarget(r.part, x.target), x.external));
			const entries = r.base === void 0 ? kept : kept.filter((x) => !r.removed.has(idOf(x)));
			changed.set(partKey(relsPart), {
				name: relsPart,
				xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join("")}${r.added.join("")}</Relationships>`
			});
		}
		const missing = /* @__PURE__ */ new Map();
		for (const m of this.mediaParts) if (this.pkg.contentType(m.partName) === void 0) missing.set(m.info.extension, m.info.contentType);
		if (missing.size > 0 || this.overrides.size > 0 || this.removedOverrides.size > 0) {
			const defaults = new Map(this.pkg.mediaTypeDefaults);
			for (const [ext, type] of missing) if (!defaults.has(ext.toLowerCase())) defaults.set(ext.toLowerCase(), type);
			const out = [];
			for (const [ext, type] of defaults) out.push(`<Default Extension="${escapeAttr(ext)}" ContentType="${escapeAttr(type)}"/>`);
			for (const name of this.pkg.partNames) {
				const type = this.pkg.overrideType(name);
				if (type === void 0 || this.removedOverrides.has(partKey(name)) || this.removedParts.has(partKey(name)) || this.overrides.has(name)) continue;
				out.push(`<Override PartName="${escapeAttr(name)}" ContentType="${escapeAttr(type)}"/>`);
			}
			for (const [name, type] of this.overrides) out.push(`<Override PartName="${escapeAttr(name)}" ContentType="${escapeAttr(type)}"/>`);
			const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${out.join("")}</Types>`;
			changed.set(partKey("/[Content_Types].xml"), {
				name: "/[Content_Types].xml",
				xml
			});
		}
		return {
			added,
			changed
		};
	}
	/** The total size of the media added, for the output limit. */
	get mediaBytes() {
		return this.mediaParts.reduce((n, m) => n + m.bytes.length, 0);
	}
};
// src/docx/numbering.ts
const NUMBERING = /\/numbering$/;
/** The numbering part of a document and the instances a render adds to it. */
var Numbering = class {
	part;
	xml;
	abstractOf = /* @__PURE__ */ new Map();
	added = [];
	next = 1;
	/** Where the new instances go: before w:numIdMacAtCleanup, or before the end tag. */
	at = -1;
	prefix = "w:";
	constructor(pkg, main, limits) {
		const rel = pkg.relationships(main).find((r) => !r.external && NUMBERING.test(r.type));
		this.part = rel !== void 0 && pkg.hasPart(rel.target) ? rel.target : void 0;
		this.xml = this.part === void 0 ? "" : normalizeLineEnds(pkg.readXml(this.part) ?? "");
		if (this.part === void 0 || this.xml === "") return;
		const root = parseRawTree(this.xml, {
			part: this.part,
			limits
		}).root;
		this.prefix = root.qname.includes(":") ? root.qname.slice(0, root.qname.indexOf(":") + 1) : "";
		let cleanup;
		for (const c of root.children) {
			if (c.kind !== "element") continue;
			if (c.local === "num") {
				const id = attr(c, "numId");
				const abstract = c.children.find((x) => x.kind === "element" && x.local === "abstractNumId");
				const abstractId = abstract === void 0 ? void 0 : attr(abstract, "val");
				if (id !== void 0 && abstractId !== void 0) this.abstractOf.set(id, abstractId);
				if (id !== void 0 && Number(id) >= this.next) this.next = Number(id) + 1;
			}
			if (c.local === "numIdMacAtCleanup") cleanup = c;
		}
		this.at = cleanup !== void 0 ? cleanup.start : root.close === "" ? -1 : this.xml.lastIndexOf(root.close);
	}
	/** A new numbering instance like `numId`, starting at `start`; undefined when there is no such numbering. */
	restart(numId, start) {
		const abstract = this.abstractOf.get(numId);
		if (abstract === void 0 || this.at < 0) return void 0;
		const id = String(this.next++);
		const p = this.prefix;
		this.added.push(`<${p}num ${p}numId="${id}"><${p}abstractNumId ${p}val="${escapeAttr(abstract)}"/><${p}lvlOverride ${p}ilvl="0"><${p}startOverride ${p}val="${String(Math.max(0, Math.floor(start)))}"/></${p}lvlOverride></${p}num>`);
		return id;
	}
	/** The numbering part with the new instances, or null when none were added. */
	changed() {
		if (this.part === void 0 || this.added.length === 0) return null;
		return {
			name: this.part,
			xml: this.xml.slice(0, this.at) + this.added.join("") + this.xml.slice(this.at)
		};
	}
};
/** An attribute in the WordprocessingML namespace (w:val, w:numId). */
function attr(e, local) {
	for (const [key, value] of e.attrs) if (key === local || key.endsWith(`}${local}`)) return value;
}
// src/pptx/render.ts
const DML_DIALECT = {
	namespaces: DML_NAMESPACES,
	preserveSpace: false
};
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
function serialize(node) {
	const out = [];
	serializeInto(node, out);
	return out.join("");
}
const child = (e, test) => e?.children.find((c) => c.kind === "element" && test(c));
/** The position and size of a shape (x, y, cx, cy in EMU), or null when it has none of its own. */
function frameOf(shape) {
	const spPr = child(shape, (c) => isPml(c, "spPr"));
	const xfrm = child(spPr, (c) => isDml(c, "xfrm"));
	const off = child(xfrm, (c) => isDml(c, "off"));
	const ext = child(xfrm, (c) => isDml(c, "ext"));
	const num = (e, name) => Number(e?.attrs.get(name));
	const frame = [
		num(off, "x"),
		num(off, "y"),
		num(ext, "cx"),
		num(ext, "cy")
	];
	return frame.every(Number.isFinite) && frame[2] > 0 && frame[3] > 0 ? frame : null;
}
/** a:rPr of a run with attributes set (b, i) and an element added before the extension list (a:hlinkClick). */
function runProps(p, base, attrs, extra = "") {
	let open = base === void 0 ? `<${p}rPr>` : base.open.endsWith("/>") ? base.open.replace(/\s*\/>$/, ">") : base.open;
	for (const [name, value] of Object.entries(attrs)) {
		open = open.replace(new RegExp(`\\s${name}\\s*=\\s*(?:"[^"]*"|'[^']*')`), "");
		open = open.replace(/>$/, ` ${name}="${value}">`);
	}
	const kids = (base?.children ?? []).filter((c) => c.kind === "element");
	const tail = kids.findIndex((c) => isDml(c, "hlinkMouseOver") || isDml(c, "rtl") || isDml(c, "extLst"));
	const before = (tail === -1 ? kids : kids.slice(0, tail)).filter((c) => !(extra !== "" && isDml(c, "hlinkClick"))).map(serialize);
	const after = tail === -1 ? [] : kids.slice(tail).map(serialize);
	return `${open}${before.join("")}${extra}${after.join("")}</${p}rPr>`;
}
/** A DrawingML run with properties and text. */
function run(model, rPr, text) {
	const p = prefixOf(model.qname);
	return `${model.open.endsWith("/>") ? `<${model.qname}>` : model.open}${rPr}<${p}t>${escapeText(text)}</${p}t></${model.qname}>`;
}
/** A line break between runs, with the run properties of the run before it. */
function lineBreak(p, rPr) {
	return `<${p}br>${rPr}</${p}br>`;
}
/** Story renderer for slides and notes. */
var SlideRenderer = class extends StoryRenderer {
	isText(e, local) {
		return isDml(e, local);
	}
	dropInCopy(e) {
		return e.ns === P14 && e.local === "creationId";
	}
	isolate(paragraph, tags) {
		return isolateTags(paragraph, tags, DML_DIALECT);
	}
	prepare() {
		const problems = super.prepare();
		for (const [anchor, tag] of [...this.imageAnchors]) {
			this.imageAnchors.delete(anchor);
			const shape = shapeOf(tag.paragraph.element);
			if (shape === void 0 || shapeText(shape, this.story.paragraphs).trim() !== tag.info.raw) {
				problems.push(new TvError("TV_E_TAG_PLACEMENT", `The image tag "${clip(tag.info.raw)}" shares its shape with other text.`, {
					location: tag.info.location,
					hint: "On a slide, an image takes the place of a shape: put the image tag alone in a shape of the size the image should have.",
					details: { tag: clip(tag.info.raw) }
				}));
				continue;
			}
			if (frameOf(shape) === null) {
				problems.push(new TvError("TV_E_TAG_PLACEMENT", `The shape of the image tag "${clip(tag.info.raw)}" has no position and size of its own.`, {
					location: tag.info.location,
					hint: "Use a text box, or move or resize the placeholder once, so that the shape has its own position and size.",
					details: { tag: clip(tag.info.raw) }
				}));
				continue;
			}
			this.imageAnchors.set(shape, tag);
		}
		return problems;
	}
	valueRuns(anchor, text) {
		if (text === "") return [];
		const p = prefixOf(anchor.qname);
		const rPr = child(anchor, (c) => isDml(c, "rPr"));
		const props = rPr === void 0 ? "" : serialize(rPr);
		return [rawXml(text.replace(/\r\n?/g, "\n").split("\n").map((line) => line === "" ? "" : run(anchor, props, line)).join(lineBreak(p, props)))];
	}
	linkNodes(anchor, link, _tag, style = {}) {
		const p = prefixOf(anchor.qname);
		const rId = this.ctx.resources.relationship(this.targetPart, this.ctx.resources.vocabulary.hyperlinkRel, link.url, true);
		const attrs = {};
		if (style.bold === true) attrs["b"] = "1";
		if (style.italic === true) attrs["i"] = "1";
		return [rawXml(run(anchor, runProps(p, child(anchor, (c) => isDml(c, "rPr")), attrs, `<${p}hlinkClick xmlns:r="${this.ctx.resources.vocabulary.r}" r:id="${escapeAttr(rId)}"/>`), link.text))];
	}
	/** A picture in the frame of the shape, keeping the image's proportions, centred (D-037). */
	imageNodes(shape, tag, scopes) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return [];
		if (value === void 0 || value === null) {
			if (value === void 0) this.missing(tag, tag.info.path, scopes);
			return [];
		}
		const image = this.imageValue(tag, value);
		if (image === null) return [];
		const info = imageInfo(image.bytes);
		if (typeof info === "string") {
			this.invalid(tag, `${info} (in "${clip(tag.info.raw)}")`, "Pass a PNG, JPEG or GIF image.");
			return [];
		}
		if (info.width * info.height > this.ctx.limits.maxImagePixels) throw limitError("maxImagePixels", this.ctx.limits.maxImagePixels, `An image has ${String(info.width)} × ${String(info.height)} pixels, more than the limit.`, tag.info.location);
		const [x, y, w, h] = frameOf(shape);
		const ratio = info.height / info.dpiY / (info.width / info.dpiX);
		let cx = w;
		let cy = Math.round(w * ratio);
		if (cy > h) {
			cy = h;
			cx = Math.round(h / ratio);
		}
		const ox = Math.round(x + (w - cx) / 2);
		const oy = Math.round(y + (h - cy) / 2);
		const media = this.ctx.resources.image(image.bytes, info);
		const rId = this.ctx.resources.imageRelationship(this.targetPart, media);
		const nv = child(shape, (c) => isPml(c, "nvSpPr"));
		const cNvPr = child(nv, (c) => isPml(c, "cNvPr"));
		const id = cNvPr?.attrs.get("id") ?? "1";
		const name = cNvPr?.attrs.get("name") ?? "Image";
		const p = prefixOf(shape.qname);
		return [rawXml(`<${p}pic><${p}nvPicPr><${p}cNvPr id="${escapeAttr(id)}" name="${escapeAttr(name)}" descr="${escapeAttr(image.alt)}"/><${p}cNvPicPr><a:picLocks xmlns:a="${this.ctx.resources.vocabulary.a}" noChangeAspect="1"/></${p}cNvPicPr><${p}nvPr/></${p}nvPicPr><${p}blipFill><a:blip xmlns:a="${this.ctx.resources.vocabulary.a}" xmlns:r="${this.ctx.resources.vocabulary.r}" r:embed="${escapeAttr(rId)}"/><a:stretch xmlns:a="${this.ctx.resources.vocabulary.a}"><a:fillRect/></a:stretch></${p}blipFill><${p}spPr><a:xfrm xmlns:a="${this.ctx.resources.vocabulary.a}"><a:off x="${String(ox)}" y="${String(oy)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm><a:prstGeom xmlns:a="${this.ctx.resources.vocabulary.a}" prst="rect"><a:avLst/></a:prstGeom></${p}spPr></${p}pic>`)];
	}
	/** Inline Markdown as runs with the properties of the anchor. */
	inlineXml(anchor, tag, content) {
		const p = prefixOf(anchor.qname);
		const base = child(anchor, (c) => isDml(c, "rPr"));
		const plain = base === void 0 ? "" : serialize(base);
		const out = [];
		let charged = 0;
		const push = (xml) => {
			this.charge(xml.length);
			charged += xml.length;
			out.push(xml);
		};
		for (const inline of content) {
			if (inline.breakBefore === true) push(lineBreak(p, plain));
			const attrs = {};
			if (inline.bold) attrs["b"] = "1";
			if (inline.italic) attrs["i"] = "1";
			if (inline.link !== void 0) {
				if (checkLinkUrl(inline.link) === null) {
					const rId = this.ctx.resources.relationship(this.targetPart, this.ctx.resources.vocabulary.hyperlinkRel, inline.link, true);
					push(run(anchor, runProps(p, base, attrs, `<${p}hlinkClick xmlns:r="${this.ctx.resources.vocabulary.r}" r:id="${escapeAttr(rId)}"/>`), inline.text));
					continue;
				}
			}
			push(run(anchor, Object.keys(attrs).length === 0 ? plain : runProps(p, base, attrs), inline.text));
		}
		this.ctx.output -= charged;
		return out.join("");
	}
	/** Markdown in a slide: paragraphs with indentation levels for lists, bold paragraphs for headings (D-038). */
	markdownNodes(node, tag, scopes, blocks) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return [];
		if (value === void 0) {
			this.missing(tag, tag.info.path, scopes);
			return [];
		}
		if (value === null || value === "") return [];
		if (typeof value !== "string") {
			this.invalid(tag, `The value of "${clip(tag.info.raw)}" must be Markdown text.`, "Pass the Markdown as a string.");
			return [];
		}
		if (value.length > this.ctx.limits.maxMarkdownLength) throw limitError("maxMarkdownLength", this.ctx.limits.maxMarkdownLength, `A Markdown value is longer than ${String(this.ctx.limits.maxMarkdownLength)} characters.`, tag.info.location);
		if (this.checkText(tag, value) === null) return [];
		const parsed = parseMarkdown(value);
		const anchor = this.anchors.get(tag);
		if (!blocks) {
			if (parsed.length > 1 || parsed[0] !== void 0 && parsed[0].kind !== "paragraph") {
				this.invalid(tag, `The Markdown of "${clip(tag.info.raw)}" has headings, lists or several paragraphs, but the tag shares its paragraph with other text.`, `Put ${tag.info.raw} alone in a paragraph to get headings, lists and paragraphs.`);
				return [];
			}
			return parsed[0] === void 0 ? [] : [rawXml(this.inlineXml(node, tag, parsed[0].content))];
		}
		const p = prefixOf(node.qname);
		const pPr = child(node, (c) => isDml(c, "pPr"));
		const end = child(node, (c) => isDml(c, "endParaRPr"));
		const open = node.open.endsWith("/>") ? `<${node.qname}>` : node.open;
		const close = node.close === "" ? `</${node.qname}>` : node.close;
		const out = [];
		let charged = 0;
		for (const block of parsed) {
			let props = pPr === void 0 ? "" : serialize(pPr);
			if (block.kind === "item") props = (pPr === void 0 ? `<${p}pPr/>` : serialize(pPr)).replace(/\slvl\s*=\s*(?:"[^"]*"|'[^']*')/, "").replace(/^<([^\s/>]+)/, `<$1 lvl="${String(block.level)}"`);
			const content = block.kind === "heading" ? block.content.map((i) => ({
				...i,
				bold: true
			})) : block.content;
			const paragraph = `${open}${props}${this.inlineXml(anchor, tag, content)}${end === void 0 ? "" : serialize(end)}${close}`;
			this.charge(paragraph.length);
			charged += paragraph.length;
			out.push(paragraph);
		}
		this.ctx.output -= charged;
		return [rawXml(out.join(""))];
	}
};
// src/prepare.ts
/** A new render context for a document or presentation (not a workbook). */
function createContext(analysis, limits, options) {
	const { zip, pkg, main } = analysis.docx;
	const mainStory = analysis.stories.find((s) => s.story === "body");
	return {
		limits,
		formatting: new Formatting({
			locale: options.locale,
			timeZone: options.timeZone
		}),
		missing: options.missing,
		problems: [],
		iterations: 0,
		output: 0,
		nextDrawingId: Math.max(0, ...analysis.stories.map((s) => maxDrawingId(s.doc))) + 1,
		resources: new Resources(pkg, zip, limits, analysis.docx.format === "pptx" ? "/ppt/media/" : "/word/media/", vocabularyOf(pkg.relationships(null).find((r) => r.type.endsWith("/officeDocument"))?.type ?? "")),
		styles: readStyles(pkg, main, limits),
		numbering: new Numbering(pkg, main, limits),
		textWidth: mainStory === void 0 ? 576e4 : textWidthOf(mainStory.doc),
		fetched: options.fetched,
		collect: options.collect,
		warnings: [],
		scopePaths: /* @__PURE__ */ new WeakMap()
	};
}
/** Prepares every story; returns them with the problems of placement. */
function prepareStories(analysis, ctx) {
	const pptx = analysis.docx.format === "pptx";
	const prepared = /* @__PURE__ */ new Map();
	for (const story of analysis.stories) prepared.set(story, pptx ? prepareStory(story, ctx, (s, c) => new SlideRenderer(s, c), analysis.droppedShapes) : prepareStory(story, ctx));
	return {
		prepared,
		problems: [...prepared.values()].flatMap((p) => p.problems)
	};
}
// src/pptx/assemble.ts
const MAX_SLIDE_ID = 2147483647;
/**
* Renders the stories of a presentation and returns the XML of every changed or new part, keyed
* by part name. Relationships, Overrides and removed parts are recorded in `ctx.resources`.
*/
function renderPresentation(analysis, ctx, prepared, data) {
	const { pkg, main } = analysis.docx;
	const { slides } = presentationSlides(pkg, main, ctx.limits);
	const slideStory = /* @__PURE__ */ new Map();
	const notesStory = /* @__PURE__ */ new Map();
	for (const s of analysis.stories) (s.story === "notes" ? notesStory : slideStory).set(s.slide ?? 0, s);
	const slideOf = (tagStory) => tagStory?.slide ?? 0;
	const storyOf = /* @__PURE__ */ new Map();
	for (const s of analysis.stories) for (const t of s.tags) storyOf.set(t, s);
	const collect = (children) => {
		const out = [];
		for (const c of children) {
			if (!("section" in c)) continue;
			if (analysis.slideSections.has(c.section)) out.push({
				tree: c.section,
				from: slideOf(storyOf.get(c.section.open)),
				to: slideOf(storyOf.get(c.section.close)),
				inner: collect(c.section.children)
			});
		}
		return out;
	};
	const sections = collect(analysis.trees[0] ?? []);
	const expand = (from, to, secs, scopes, out) => {
		for (let i = from; i <= to;) {
			const sec = secs.find((s) => s.from === i);
			if (sec === void 0) {
				out.push({
					slide: i,
					scopes
				});
				i++;
				continue;
			}
			const story = storyOf.get(sec.tree.open);
			const renderer = prepared.get(story).renderer;
			for (const itemScopes of renderer.sectionScopes(sec.tree, scopes)) {
				if (++ctx.iterations > ctx.limits.maxIterations) throw limitError("maxIterations", ctx.limits.maxIterations, `Sections repeat their content more than ${String(ctx.limits.maxIterations)} times in all.`, sec.tree.open.info.location);
				expand(sec.from, sec.to, sec.inner, itemScopes, out);
			}
			i = sec.to + 1;
		}
	};
	const instances = [];
	expand(1, slides.length, sections, [data], instances);
	const changed = /* @__PURE__ */ new Map();
	const used = /* @__PURE__ */ new Map();
	const entries = [];
	const presentation = parseRawTree(pkg.readXml(main) ?? "", {
		part: main,
		limits: ctx.limits
	});
	const list = presentation.root.children.find((c) => c.kind === "element" && isPml(c, "sldIdLst"));
	const originals = /* @__PURE__ */ new Map();
	for (const c of list?.children ?? []) if (c.kind === "element" && isPml(c, "sldId")) originals.set(c.attrs.get("id") ?? "", rawXmlText(c));
	let nextId = Math.max(255, ...slides.map((s) => Number(s.id)).filter(Number.isSafeInteger)) + 1;
	const prefix = list === void 0 ? "p:" : list.qname.slice(0, list.qname.indexOf(":") + 1);
	for (const instance of instances) {
		const ref = slides[instance.slide - 1];
		const story = slideStory.get(instance.slide);
		const notes = notesStory.get(instance.slide);
		const k = used.get(instance.slide) ?? 0;
		used.set(instance.slide, k + 1);
		const renderSlide = prepared.get(story).render;
		if (k === 0) {
			if (story.tags.length > 0 || analysis.droppedShapes.size > 0) changed.set(ref.part, renderSlide(instance.scopes));
			if (notes !== void 0 && notes.tags.length > 0) changed.set(notes.part, prepared.get(notes).render(instance.scopes));
			entries.push(originals.get(ref.id) ?? "");
			continue;
		}
		const part = ctx.resources.newPartName("/ppt/slides/", "tplvane", "xml");
		ctx.resources.addOverride(part, pkg.contentType(ref.part) ?? "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
		const owned = pkg.relationships(ref.part).find((r) => !r.external && SLIDE_OWNED.test(r.type));
		if (owned !== void 0) throw new TvError("TV_E_NOT_SUPPORTED", `Slide ${String(instance.slide)} cannot be repeated: it has a part of its own (${owned.type.slice(owned.type.lastIndexOf("/") + 1)}) that a copy would share.`, {
			location: {
				part: ref.part,
				slide: instance.slide
			},
			hint: "Repeat slides without charts, SmartArt, comments or embedded objects; put those on slides that are shown once.",
			details: { relationship: owned.type }
		});
		const moved = /* @__PURE__ */ new Map();
		let notesPart;
		if (notes !== void 0) {
			notesPart = ctx.resources.newPartName("/ppt/notesSlides/", "tplvane", "xml");
			ctx.resources.addOverride(notesPart, pkg.contentType(notes.part) ?? "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml");
			moved.set(partKey(notes.part), notesPart);
			const notesRels = copiedRelationships(pkg, notes.part, notesPart, /* @__PURE__ */ new Map([[partKey(ref.part), part]]));
			ctx.resources.seedRelationships(notesPart, notesRels, pkg.relationships(notes.part).map((r) => r.id));
		}
		const slideRels = copiedRelationships(pkg, ref.part, part, moved);
		ctx.resources.seedRelationships(part, slideRels, pkg.relationships(ref.part).map((r) => r.id));
		ctx.output += slideRels.join("").length + (notes === void 0 ? 0 : copiedRelationships(pkg, notes.part, part, /* @__PURE__ */ new Map()).join("").length);
		if (ctx.output > ctx.limits.maxOutputBytes) throw limitError("maxOutputBytes", ctx.limits.maxOutputBytes, `The result would be larger than ${String(ctx.limits.maxOutputBytes)} characters of XML.`, { part: ref.part });
		changed.set(part, renderSlide(instance.scopes, true, part));
		if (notes !== void 0 && notesPart !== void 0) changed.set(notesPart, prepared.get(notes).render(instance.scopes, true, notesPart));
		const rId = ctx.resources.relationship(main, ctx.resources.vocabulary.slideRel, relativeTarget(main, part), false);
		if (nextId > MAX_SLIDE_ID) throw limitError("maxIterations", ctx.limits.maxIterations, "The presentation would get more slides than slide ids allow.");
		entries.push(`<${prefix}sldId xmlns:r="${ctx.resources.vocabulary.r}" id="${String(nextId++)}" r:id="${escapeAttr(rId)}"/>`);
	}
	slides.forEach((ref, i) => {
		if (used.has(i + 1)) return;
		ctx.resources.removeRelationship(main, ref.rId);
		ctx.resources.removePart(ref.part);
		if (ref.notes !== void 0) ctx.resources.removePart(ref.notes);
		changed.delete(ref.part);
	});
	removeUnreferencedParts(pkg, ctx, slides.filter((_, i) => !used.has(i + 1)));
	if (list !== void 0 && (sections.length > 0 || instances.length !== slides.length)) {
		const xml = rawXmlText(presentation.root);
		const listXml = rawXmlText(list);
		const open = list.open.endsWith("/>") ? `<${list.qname}>` : list.open;
		const close = list.close === "" ? `</${list.qname}>` : list.close;
		const newList = `${open}${entries.join("")}${close}`;
		changed.set(main, presentation.prolog + xml.replace(listXml, () => newList) + presentation.epilog);
	}
	return changed;
}
/**
* Leaves out the parts that only removed slides and notes used: their images, comments, charts and
* the workbooks embedded in those charts. Every image part shall be the target of a relationship
* (ECMA-376 Part 1 §15.2.14, F-025), and the text of a removed slide's comments or chart must not
* stay in the file (D-071). A part another kept part refers to (a layout, a shared image) stays,
* and so do parts the template already left unused.
*/
function removeUnreferencedParts(pkg, ctx, removedSlides) {
	if (removedSlides.length === 0) return;
	const removed = ctx.resources.removed;
	const referencedByKept = /* @__PURE__ */ new Set();
	const keptSources = () => {
		referencedByKept.clear();
		for (const part of pkg.partNames) {
			if (removed.has(partKey(part))) continue;
			for (const rel of pkg.relationships(part)) if (!rel.external) referencedByKept.add(partKey(rel.target));
		}
		for (const rel of pkg.relationships(null)) if (!rel.external) referencedByKept.add(partKey(rel.target));
	};
	let sources = removedSlides.flatMap((ref) => ref.notes === void 0 ? [ref.part] : [ref.part, ref.notes]);
	for (let round = 0; round < 8 && sources.length > 0; round++) {
		keptSources();
		const next = [];
		for (const source of sources) for (const rel of pkg.relationships(source)) {
			const key = partKey(rel.target);
			if (rel.external || removed.has(key) || referencedByKept.has(key) || !pkg.hasPart(rel.target)) continue;
			ctx.resources.removePart(rel.target);
			next.push(rel.target);
		}
		sources = next;
	}
}
/** Relationship types of parts that belong to one slide. */
const SLIDE_OWNED = /\/(?:chart|chartUserShapes|comments|oleObject|package|diagramData|diagramLayout|diagramQuickStyle|diagramColors|diagramDrawing|control|activeXControl|ctrlProp)$/;
/**
* The relationships of `source` for its copy `copy`, written from the parsed relationships rather
* than edited as text (a target may be written relative or absolute, D-071): internal targets
* relative to the copy, and targets in `moved` replaced by their new parts.
*/
function copiedRelationships(pkg, source, copy, moved) {
	return pkg.relationships(source).map((r) => relationshipXml(r.id, r.type, r.external ? r.target : relativeTarget(copy, moved.get(partKey(r.target)) ?? r.target), r.external));
}
/** Largest offset or size a ZIP without ZIP64 can hold (4-byte fields, Z-04). */
const MAX_32 = 4294967294;
const SIG_LOCAL = 67324752;
const SIG_CENTRAL = 33639248;
const SIG_EOCD = 101010256;
/** Version 2.0: Deflate (Z-06); upper byte 0 = MS-DOS host. */
const VERSION = 20;
/**
* MS-DOS date for 1980-01-01: (year - 1980) << 9 | month << 5 | day. APPNOTE only says
* "standard MS-DOS format" (Z-19); the bit layout is the one fflate writes
* (node_modules/fflate/esm/index.mjs, zip header writer). Time 00:00:00 is 0.
*/
const DOS_DATE = 33;
const DOS_TIME = 0;
const encoder = new TextEncoder();
function tooLarge(what) {
	return new TvError("TV_E_LIMIT_EXCEEDED", `The document is too large to write: ${what}.`, {
		hint: "Split the data into several documents. Writing ZIP64 archives is not supported yet.",
		details: { limit: "zip32" }
	});
}
/**
* Writes a ZIP archive. Throws `TV_E_INVALID_INPUT` for names that are not printable ASCII or
* that repeat (ignoring case, as OPC requires, E-58), and `TV_E_LIMIT_EXCEEDED` when the
* archive would need ZIP64. `level` is the Deflate level 0–9 (0 stores every entry).
*/
function writeZip(entries, level = 6) {
	if (entries.length > 65534) throw tooLarge(`${String(entries.length)} parts`);
	const seen = /* @__PURE__ */ new Set();
	const chunks = [];
	const central = [];
	let offset = 0;
	for (const entry of entries) {
		const name = entry.name;
		if (!/^[\x21-\x7e][\x20-\x7e]*$/.test(name) || name.startsWith("/") || name.length > 65535) throw new TvError("TV_E_INVALID_INPUT", `"${name.slice(0, 60)}" is not a valid ZIP item name.`, {
			hint: "Item names must be printable ASCII without a leading slash. This is an internal error; please report it.",
			details: { name: name.slice(0, 60) }
		});
		const key = name.toLowerCase();
		if (seen.has(key)) throw new TvError("TV_E_INVALID_INPUT", `The ZIP item name "${name.slice(0, 60)}" is used twice.`, {
			hint: "Item names must be unique ignoring case. This is an internal error; please report it.",
			details: { name: name.slice(0, 60) }
		});
		seen.add(key);
		const nameBytes = encoder.encode(name);
		let crc;
		let body;
		let method;
		let size;
		if (entry.raw !== void 0) ({crc32: crc, body, method, uncompressedSize: size} = entry.raw);
		else {
			crc = crc32(entry.data);
			const deflated = entry.data.length === 0 || level === 0 ? null : deflateSync(entry.data, { level });
			const useDeflate = deflated !== null && deflated.length < entry.data.length;
			body = useDeflate ? deflated : entry.data;
			method = useDeflate ? 8 : 0;
			size = entry.data.length;
		}
		if (size > MAX_32 || body.length > MAX_32 || offset > MAX_32) throw tooLarge("more than 4 GiB");
		const local = new Uint8Array(30 + nameBytes.length);
		const lv = new DataView(local.buffer);
		lv.setUint32(0, SIG_LOCAL, true);
		lv.setUint16(4, VERSION, true);
		lv.setUint16(6, 0, true);
		lv.setUint16(8, method, true);
		lv.setUint16(10, DOS_TIME, true);
		lv.setUint16(12, DOS_DATE, true);
		lv.setUint32(14, crc, true);
		lv.setUint32(18, body.length, true);
		lv.setUint32(22, size, true);
		lv.setUint16(26, nameBytes.length, true);
		lv.setUint16(28, 0, true);
		local.set(nameBytes, 30);
		const cd = new Uint8Array(46 + nameBytes.length);
		const cv = new DataView(cd.buffer);
		cv.setUint32(0, SIG_CENTRAL, true);
		cv.setUint16(4, VERSION, true);
		cv.setUint16(6, VERSION, true);
		cv.setUint16(8, 0, true);
		cv.setUint16(10, method, true);
		cv.setUint16(12, DOS_TIME, true);
		cv.setUint16(14, DOS_DATE, true);
		cv.setUint32(16, crc, true);
		cv.setUint32(20, body.length, true);
		cv.setUint32(24, size, true);
		cv.setUint16(28, nameBytes.length, true);
		cv.setUint32(42, offset, true);
		cd.set(nameBytes, 46);
		chunks.push(local, body);
		central.push(cd);
		offset += local.length + body.length;
	}
	const cdOffset = offset;
	const cdSize = central.reduce((n, c) => n + c.length, 0);
	if (cdOffset > MAX_32 || cdOffset + cdSize > MAX_32) throw tooLarge("more than 4 GiB");
	const eocd = /* @__PURE__ */ new Uint8Array(22);
	const ev = new DataView(eocd.buffer);
	ev.setUint32(0, SIG_EOCD, true);
	ev.setUint16(8, entries.length, true);
	ev.setUint16(10, entries.length, true);
	ev.setUint32(12, cdSize, true);
	ev.setUint32(16, cdOffset, true);
	const out = new Uint8Array(cdOffset + cdSize + eocd.length);
	let pos = 0;
	for (const part of [
		...chunks,
		...central,
		eocd
	]) {
		out.set(part, pos);
		pos += part.length;
	}
	return out;
}
// src/xlsx/shared-strings.ts
const SHARED_STRINGS = /\/sharedStrings$/;
const WORKSHEET = /\/worksheet$/;
const elements = (e) => e.children.filter((c) => c.kind === "element");
/** The workbook with its shared strings reduced to those in use; the bytes as they are when nothing changes. */
function compactSharedStrings(bytes, limits) {
	const zip = openZip(bytes, limits);
	const pkg = openPackage(zip, limits);
	const book = pkg.relationships(null).find((r) => !r.external && r.type.endsWith("/officeDocument"));
	if (book === void 0) return bytes;
	const rels = pkg.relationships(book.target);
	const table = rels.find((r) => !r.external && SHARED_STRINGS.test(r.type));
	if (table === void 0 || !pkg.hasPart(table.target)) return bytes;
	const sst = parseRawTree(pkg.readXml(table.target) ?? "", {
		part: table.target,
		limits
	});
	const items = elements(sst.root).filter((e) => e.local === "si");
	const sheets = rels.filter((r) => !r.external && WORKSHEET.test(r.type) && pkg.hasPart(r.target)).map((r) => r.target);
	const used = /* @__PURE__ */ new Map();
	const edits = /* @__PURE__ */ new Map();
	for (const sheet of sheets) {
		const xml = normalizeLineEnds(pkg.readXml(sheet) ?? "");
		const doc = parseRawTree(xml, {
			part: sheet,
			limits
		});
		const at = [];
		const visit = (e) => {
			if (e.local === "c" && e.attrs.get("t") === "s") {
				const text = elements(e).find((c) => c.local === "v")?.children.find((c) => c.kind === "text");
				const index = text === void 0 ? NaN : Number(text.text.trim());
				if (text !== void 0 && Number.isInteger(index) && index >= 0 && index < items.length) {
					if (!used.has(index)) used.set(index, used.size);
					at.push({
						start: text.start,
						length: text.raw.length,
						index
					});
				}
				return;
			}
			for (const c of elements(e)) visit(c);
		};
		visit(doc.root);
		edits.set(partKey(sheet), {
			xml,
			at
		});
	}
	if (used.size === items.length && [...used].every(([from, to]) => from === to)) return bytes;
	const kept = [...used.keys()].map((i) => rawXmlText(items[i]));
	const count = [...edits.values()].reduce((n, e) => n + e.at.length, 0);
	const root = sst.root;
	const open = root.open.replace(/\s(count|uniqueCount)\s*=\s*(?:"[^"]*"|'[^']*')/g, "").replace(/\s*\/?>$/, ` count="${escapeAttr(String(count))}" uniqueCount="${escapeAttr(String(kept.length))}">`);
	const close = root.close === "" ? `</${root.qname}>` : root.close;
	const changed = /* @__PURE__ */ new Map([[partKey(table.target), `${sst.prolog}${open}${kept.join("")}${close}${sst.epilog}`]]);
	for (const [key, { xml, at }] of edits) {
		if (at.length === 0) continue;
		let out = xml;
		for (const { start, length, index } of [...at].sort((a, b) => b.start - a.start)) out = out.slice(0, start) + String(used.get(index)) + out.slice(start + length);
		changed.set(key, out);
	}
	const encoder = new TextEncoder();
	return writeZip(zip.entries.map((entry) => {
		const xml = changed.get(partKey(zipNameToPartName(entry.name)));
		if (xml !== void 0) return {
			name: partNameToZipName(zipNameToPartName(entry.name)) ?? entry.name,
			data: encoder.encode(xml)
		};
		const raw = zip.raw(entry);
		return raw === null ? {
			name: entry.name,
			data: zip.read(entry)
		} : {
			name: entry.name,
			data: /* @__PURE__ */ new Uint8Array(0),
			raw
		};
	}));
}
// src/xlsx/tables.ts
const REF = /^(\$?[A-Z]{1,3}\$?)(\d+):(\$?[A-Z]{1,3}\$?)(\d+)$/;
const REL_ID = /(?:^|\})id$/;
/** The ref "A1:B2" with its last row moved from `last` to `last + added`; null when it does not end on `last`. */
function grow(ref, last, added) {
	const m = REF.exec(ref);
	if (m === null || Number(m[4]) !== last || Number(m[2]) > last) return null;
	return `${m[1]}${m[2]}:${m[3]}${String(last + added)}`;
}
/** Replaces the value of the attribute `ref` in an open tag, either quote. */
function withRef(open, ref) {
	return open.replace(/(\sref\s*=\s*)(?:"[^"]*"|'[^']*')/, `$1"${ref}"`);
}
/** The workbook with the tables of `sheet` that end on row `last` grown by `added` rows. */
function extendTables(bytes, sheet, last, added, limits) {
	if (added <= 0) return bytes;
	const zip = openZip(bytes, limits);
	const pkg = openPackage(zip, limits);
	const book = pkg.relationships(null).find((r) => !r.external && r.type.endsWith("/officeDocument"));
	if (book === void 0) return bytes;
	const workbook = parseRawTree(pkg.readXml(book.target) ?? "", {
		part: book.target,
		limits
	});
	let id;
	const find = (e) => {
		if (e.local === "sheet" && e.attrs.get("name") === sheet) {
			for (const [key, value] of e.attrs) if (REL_ID.test(key) && key !== "id" && key !== "sheetId") id = value;
		}
		for (const c of e.children) if (c.kind === "element") find(c);
	};
	find(workbook.root);
	const target = pkg.relationships(book.target).find((r) => r.id === id && !r.external)?.target;
	if (target === void 0) return bytes;
	const changed = /* @__PURE__ */ new Map();
	for (const rel of pkg.relationships(target)) {
		if (rel.external || !rel.type.endsWith("/table") || !pkg.hasPart(rel.target)) continue;
		const xml = normalizeLineEnds(pkg.readXml(rel.target) ?? "");
		const doc = parseRawTree(xml, {
			part: rel.target,
			limits
		});
		const ref = doc.root.attrs.get("ref");
		const grown = ref === void 0 ? null : grow(ref, last, added);
		if (grown === null) continue;
		let out = xml.slice(0, doc.root.start) + withRef(doc.root.open, grown) + xml.slice(doc.root.start + doc.root.open.length);
		const filter = doc.root.children.find((c) => c.kind === "element" && c.local === "autoFilter");
		const filterRef = filter?.attrs.get("ref");
		const filterGrown = filter === void 0 || filterRef === void 0 ? null : grow(filterRef, last, added);
		if (filter !== void 0 && filterGrown !== null) {
			const shift = out.length - xml.length;
			const at = filter.start + shift;
			out = out.slice(0, at) + withRef(filter.open, filterGrown) + out.slice(at + filter.open.length);
		}
		changed.set(partKey(rel.target), out);
	}
	if (changed.size === 0) return bytes;
	const encoder = new TextEncoder();
	return writeZip(zip.entries.map((entry) => {
		const xml = changed.get(partKey(zipNameToPartName(entry.name)));
		if (xml !== void 0) return {
			name: partNameToZipName(zipNameToPartName(entry.name)) ?? entry.name,
			data: encoder.encode(xml)
		};
		const raw = zip.raw(entry);
		return raw === null ? {
			name: entry.name,
			data: zip.read(entry)
		} : {
			name: entry.name,
			data: /* @__PURE__ */ new Uint8Array(0),
			raw
		};
	}));
}
// src/xlsx/workbook-props.ts
/** The attribute with the white space before it, in either quotes. */
const DATE_COMPATIBILITY = /\s+dateCompatibility\s*=\s*(?:"[^"]*"|'[^']*')/;
/** The workbook without workbookPr/@dateCompatibility; the bytes as they are when it has none. */
function dropDateCompatibility(bytes, limits) {
	const zip = openZip(bytes, limits);
	const pkg = openPackage(zip, limits);
	const book = pkg.relationships(null).find((r) => !r.external && r.type.endsWith("/officeDocument"));
	if (book === void 0 || !pkg.hasPart(book.target)) return bytes;
	const xml = normalizeLineEnds(pkg.readXml(book.target) ?? "");
	if (!xml.includes("dateCompatibility")) return bytes;
	const props = parseRawTree(xml, {
		part: book.target,
		limits
	}).root.children.find((c) => c.kind === "element" && c.local === "workbookPr");
	if (props === void 0 || !props.attrs.has("dateCompatibility")) return bytes;
	const open = props.open.replace(DATE_COMPATIBILITY, "");
	const changed = xml.slice(0, props.start) + open + xml.slice(props.start + props.open.length);
	const key = partKey(book.target);
	const encoder = new TextEncoder();
	return writeZip(zip.entries.map((entry) => {
		if (partKey(zipNameToPartName(entry.name)) === key) return {
			name: partNameToZipName(zipNameToPartName(entry.name)) ?? entry.name,
			data: encoder.encode(changed)
		};
		const raw = zip.raw(entry);
		return raw === null ? {
			name: entry.name,
			data: zip.read(entry)
		} : {
			name: entry.name,
			data: /* @__PURE__ */ new Uint8Array(0),
			raw
		};
	}));
}
// src/xlsx/render.ts
/** The most characters a cell can hold (observed in xlvane 1.3.3, which refuses more: F-021). */
const MAX_CELL_TEXT = 32767;
/** The tree children in each cell, with the tags that are to show nothing (those of a section over rows). */
function byCell(children, cellOf, out = /* @__PURE__ */ new Map()) {
	for (const child of children) {
		const tag = "tag" in child ? child.tag : child.section.open;
		const cell = cellOf.get(tag.paragraph);
		const list = out.get(cell) ?? [];
		list.push(child);
		out.set(cell, list);
	}
	return out;
}
const startOf = (child) => "tag" in child ? child.tag.start : child.section.open.start;
var CellRenderer = class extends TagEvaluator {
	constructor(ctx) {
		super(ctx);
	}
	/** The scopes of each repetition of a section, counted against maxIterations. */
	repetitions(section, scopes) {
		const list = this.iterations(section, scopes);
		this.ctx.iterations += list.length;
		if (this.ctx.iterations > this.ctx.limits.maxIterations) throw limitError("maxIterations", this.ctx.limits.maxIterations, `Sections repeat their content more than ${String(this.ctx.limits.maxIterations)} times in all.`, section.open.info.location);
		return list;
	}
	/**
	* The value of a cell. `items` are the tree children in the cell; tags in `silent` show nothing.
	* A cell that holds one value tag and nothing else (but spaces) gets the type of the value.
	* Returns undefined when a tag has a problem, which is recorded.
	*/
	cell(text, items, silent, scopes) {
		const sorted = [...items].sort((a, b) => startOf(a) - startOf(b));
		const shown = sorted.filter((c) => !("tag" in c && silent.has(c.tag)));
		const only = shown.length === 1 ? shown[0] : void 0;
		if (only !== void 0 && "tag" in only && only.tag.info.kind === "value") {
			let rest = "";
			let pos = 0;
			for (const c of sorted) {
				rest += text.slice(pos, startOf(c));
				pos = "tag" in c ? c.tag.end : c.section.close.end;
			}
			rest += text.slice(pos);
			if (rest.trim() === "") return this.typed(only.tag, scopes);
		}
		const out = this.range(text, 0, text.length, sorted, silent, scopes);
		if (out.length > MAX_CELL_TEXT) {
			const tag = shown[0] === void 0 ? void 0 : "tag" in shown[0] ? shown[0].tag : shown[0].section.open;
			if (tag !== void 0) this.invalid(tag, `The text of the cell would have ${String(out.length)} characters.`, `A cell holds at most ${String(MAX_CELL_TEXT)} characters. Shorten the value.`);
			return;
		}
		this.count(out.length, items);
		return out === "" ? null : out;
	}
	count(length, items) {
		this.ctx.output += length;
		if (this.ctx.output > this.ctx.limits.maxOutputBytes) {
			const first = items[0];
			throw limitError("maxOutputBytes", this.ctx.limits.maxOutputBytes, `The result would be larger than ${String(this.ctx.limits.maxOutputBytes)} characters of text.`, first === void 0 ? {} : ("tag" in first ? first.tag : first.section.open).info.location);
		}
	}
	/** The text of `text[from, to)` with its tags and sections rendered; a tag with a problem shows nothing (the problem is recorded). */
	range(text, from, to, children, silent, scopes) {
		let out = "";
		let pos = from;
		for (const child of children) {
			out += text.slice(pos, startOf(child));
			if ("tag" in child) {
				pos = child.tag.end;
				if (silent.has(child.tag)) continue;
				const value = this.valueText(child.tag, scopes);
				if (value !== null) out += typeof value === "string" ? value : value.text;
				continue;
			}
			const { section } = child;
			for (const inner of this.repetitions(section, scopes)) {
				out += this.range(text, section.open.end, section.close.start, section.children, silent, inner);
				if (out.length > MAX_CELL_TEXT) break;
			}
			pos = section.close.end;
		}
		return out + text.slice(pos, to);
	}
	/** The typed value of a value tag that fills its cell. */
	typed(tag, scopes) {
		const { ok, value } = this.evaluate(tag, scopes);
		if (!ok) return void 0;
		if (typeof value === "number" && Number.isFinite(value)) return value;
		if (typeof value === "boolean") return value;
		if (value instanceof Date) {
			const year = value.getUTCFullYear();
			if (!Number.isNaN(value.getTime()) && (year < 1900 || year > 9999)) {
				this.invalid(tag, `The date of "${clip(tag.info.raw)}" is outside the years 1900 to 9999 that Excel can hold.`, "Pass a date from 1900 to 9999, or format it as text with | date.");
				return;
			}
			if (!Number.isNaN(value.getTime())) return wallClock(value, this.ctx.formatting.ctx.timeZone);
			this.invalid(tag, `The value of "${clip(tag.info.raw)}" is an invalid date.`, "Pass a valid Date.");
			return;
		}
		if (value === null) return null;
		if (value === void 0 && this.ctx.missing === "empty") {
			this.valueText(tag, scopes);
			return null;
		}
		const text = this.valueText(tag, scopes);
		if (text === null) return void 0;
		const s = typeof text === "string" ? text : text.text;
		if (s.length > MAX_CELL_TEXT) {
			this.invalid(tag, `The value of "${clip(tag.info.raw)}" has ${String(s.length)} characters.`, `A cell holds at most ${String(MAX_CELL_TEXT)} characters. Shorten the value.`);
			return;
		}
		this.count(s.length, [{ tag }]);
		return s === "" ? null : s;
	}
};
/** A Date whose UTC fields are the date and time of `date` in `timeZone`. */
function wallClock(date, timeZone) {
	if (timeZone === "UTC") return date;
	const parts = {};
	for (const p of new Intl.DateTimeFormat("en-US", {
		timeZone,
		year: "numeric",
		month: "numeric",
		day: "numeric",
		hour: "numeric",
		minute: "numeric",
		second: "numeric",
		hourCycle: "h23"
	}).formatToParts(date)) if (p.type !== "literal") parts[p.type] = Number(p.value);
	const out = new Date(Date.UTC(2e3, 0, 1, parts["hour"] ?? 0, parts["minute"] ?? 0, parts["second"] ?? 0, date.getUTCMilliseconds()));
	out.setUTCFullYear(parts["year"] ?? 1970, (parts["month"] ?? 1) - 1, parts["day"] ?? 1);
	return out;
}
/** Evaluates every tag of the workbook against the data. Problems go to `ctx.problems`. */
function planWorkbook(analysis, data, ctx) {
	const renderer = new CellRenderer(ctx);
	const edits = [];
	const sections = [];
	const none = /* @__PURE__ */ new Set();
	for (const sheet of analysis.sheets) {
		const inRows = new Set(sheet.rowSections.map((r) => r.section));
		const plain = sheet.tree.filter((c) => !("section" in c && inRows.has(c.section)));
		for (const [cell, items] of byCell(plain, analysis.cellOf)) {
			const value = renderer.cell(cell.paragraph.text, items, none, [data]);
			if (value !== void 0) edits.push({
				sheet: sheet.sheet.name,
				cell: cell.cell.ref,
				value
			});
		}
		for (const rows of sheet.rowSections) {
			const { section } = rows;
			const silent = /* @__PURE__ */ new Set([section.open, section.close]);
			const cells = byCell(section.children, analysis.cellOf);
			for (const tag of silent) {
				const cell = analysis.cellOf.get(tag.paragraph);
				const list = cells.get(cell) ?? [];
				list.push({ tag });
				cells.set(cell, list);
			}
			const repetitions = renderer.repetitions(section, [data]);
			const values = /* @__PURE__ */ new Map();
			for (const [cell, items] of cells) {
				if (rows.markers && (cell.cell.row === rows.open || cell.cell.row === rows.close)) continue;
				const list = [];
				for (const scopes of repetitions) list.push(renderer.cell(cell.paragraph.text, items, silent, scopes) ?? null);
				values.set(cell, list);
			}
			if (repetitions.length > 1 && rows.last >= rows.first) {
				const perCopy = rowCells(sheet.sheet, rows.first, rows.last).reduce((n, cell) => n + 32 + String(cell.formula?.text ?? cell.value ?? "").length, 0);
				ctx.output += perCopy * (repetitions.length - 1);
				if (ctx.output > ctx.limits.maxOutputBytes) throw limitError("maxOutputBytes", ctx.limits.maxOutputBytes, `The repeated rows would be larger than ${String(ctx.limits.maxOutputBytes)} characters.`, section.open.info.location);
			}
			sections.push({
				sheet: sheet.sheet.name,
				rows,
				count: repetitions.length,
				cells: values
			});
		}
	}
	return {
		edits,
		sections
	};
}
/** A cell value as it is in the file, to write it again in a copy. */
function copyValue(cell, sheet, offset) {
	if (cell.formula !== null) {
		const { formula } = cell;
		const unsupported = (why) => new TvError("TV_E_NOT_SUPPORTED", `The formula in ${sheet.name}!${cell.ref} cannot be copied to the repeated rows: ${why}`, {
			location: cellLocation(sheet, cell.ref),
			hint: "Move the formula out of the rows that repeat, or write it as an ordinary formula with A1 references.",
			details: { kind: formula.kind }
		});
		if (formula.kind === "array" || formula.kind === "dataTable") throw unsupported(formula.kind === "array" ? "it is an array formula." : "it is a data table.");
		if (formula.text === null) throw unsupported("it is a shared formula that cannot be read for this cell.");
		const moved = shiftFormula(formula.text, offset, 0);
		if (moved === null) throw unsupported("its references cannot be moved safely, or would leave the sheet.");
		return { formula: moved };
	}
	switch (cell.type) {
		case "empty": return;
		case "error": return { error: cell.value };
		case "string": return cell.richText === void 0 ? cell.value : { richText: cell.richText };
		default: return cell.value;
	}
}
/** Limits for files that tplvane itself wrote in an earlier step: their size grew with the repetitions, which the limits of rendering already bound. */
function ownLimits(limits) {
	return { limits: {
		...xlvaneLimits(limits),
		maxCompressionRatio: Number.MAX_SAFE_INTEGER
	} };
}
/** The cells of rows `first` to `last`, with values or formats. */
function rowCells(sheet, first, last) {
	const out = [];
	for (const cell of sheet.cells()) if (cell.row >= first && cell.row <= last) out.push(cell);
	for (const cell of sheet.styledCells()) if (cell.row >= first && cell.row <= last) out.push(cell);
	return out;
}
/** The least work the sections of a plan will take (see applySection), checked before any of it is done. */
function checkSectionWork(plan, size, limits) {
	const first = plan.sections[0];
	if (first !== void 0 && plan.sections.length * size * 48 > limits.maxOutputBytes) throw limitError("maxOutputBytes", limits.maxOutputBytes, `The workbook has so many repeated sections that writing them would process more than ${String(limits.maxOutputBytes)} bytes.`, first.rows.section.open.info.location);
}
/** Writes the plan into the workbook. */
function applyWorkbook(bytes, plan, limits) {
	const own = ownLimits(limits);
	let current = bytes;
	if (plan.edits.length > 0) current = viaXlvane(() => editCells(current, plan.edits, { limits: xlvaneLimits(limits) }).bytes, "Writing the cells failed");
	const budget = {
		output: 0,
		work: 0,
		limits
	};
	checkSectionWork(plan, current.length, limits);
	for (const planned of [...plan.sections].reverse()) current = applySection(current, planned, own, budget);
	return dropDateCompatibility(compactSharedStrings(dropCachedResults(current, own), limits), limits);
}
/** Writes every formula that has a cached result again without it, so that applications compute it on open. */
function dropCachedResults(bytes, own) {
	const workbook = viaXlvane(() => readWorkbook(bytes, own), "Reading the workbook failed");
	const edits = [];
	for (const sheet of workbook.sheets) {
		if (sheet.kind !== "worksheet") continue;
		for (const cell of sheet.cells()) {
			const { formula } = cell;
			if (formula === null || formula.text === null || formula.kind === "dataTable" || formula.kind === "array" || cell.type === "empty") continue;
			edits.push({
				sheet: sheet.name,
				cell: cell.ref,
				value: { formula: formula.text }
			});
		}
	}
	return edits.length === 0 ? bytes : viaXlvane(() => editCells(bytes, edits, own).bytes, "Writing the formulas failed");
}
function applySection(bytes, planned, own, budget) {
	const { rows, count, sheet: name } = planned;
	let current = bytes;
	budget.work += current.length * 48;
	if (budget.work > budget.limits.maxOutputBytes) throw limitError("maxOutputBytes", budget.limits.maxOutputBytes, `The workbook has so many repeated sections that writing them would process more than ${String(budget.limits.maxOutputBytes)} bytes.`, rows.section.open.info.location);
	if (count === 0) return viaXlvane(() => deleteRows(current, {
		sheet: name,
		at: rows.open,
		count: rows.close - rows.open + 1,
		...own
	}).bytes, `Deleting the rows of ${clip(rows.section.open.info.raw)} failed`);
	const height = rows.last - rows.first + 1;
	const added = (count - 1) * height;
	const sheet = viaXlvane(() => readWorkbook(current, own), "Reading the workbook failed").getSheet(name);
	const template = height > 0 ? rowCells(sheet, rows.first, rows.last) : [];
	const tagCells = /* @__PURE__ */ new Map();
	for (const [cell, values] of planned.cells) tagCells.set(cell.cell.ref, values);
	const edits = [];
	for (const [ref, values] of tagCells) edits.push({
		sheet: name,
		cell: ref,
		value: values[0] ?? null
	});
	if (added > 0) {
		const perCopy = template.reduce((n, cell) => n + 32 + String(cell.formula?.text ?? cell.value ?? "").length, 0);
		budget.output += perCopy * (count - 1);
		if (budget.output > budget.limits.maxOutputBytes) throw limitError("maxOutputBytes", budget.limits.maxOutputBytes, `The repeated rows would be larger than ${String(budget.limits.maxOutputBytes)} characters.`, rows.section.open.info.location);
		current = viaXlvane(() => insertRows(current, {
			sheet: name,
			at: rows.last + 1,
			count: added,
			style: height === 1 ? "above" : "none",
			...own
		}).bytes, `Inserting the rows of ${clip(rows.section.open.info.raw)} failed`);
		current = extendTables(current, name, rows.last, added, budget.limits);
		const copies = template.map((cell) => ({
			cell,
			value: void 0
		}));
		for (let k = 1; k < count; k++) {
			const offset = k * height;
			for (const { cell } of copies) {
				const values = tagCells.get(cell.ref);
				const value = values === void 0 ? copyValue(cell, sheet, offset) : values[k] ?? null;
				if (value === void 0) continue;
				const ref = `${cell.ref.replace(/[0-9]+$/, "")}${String(cell.row + offset)}`;
				const style = height > 1 && Object.keys(cell.style).length > 0 ? cell.style : void 0;
				edits.push(style === void 0 ? {
					sheet: name,
					cell: ref,
					value
				} : {
					sheet: name,
					cell: ref,
					value,
					style
				});
			}
		}
		const after = viaXlvane(() => readWorkbook(current, own), "Reading the workbook failed");
		const change = {
			sheet: name,
			row: rows.last,
			to: rows.last + added
		};
		for (const other of after.sheets) {
			if (other.kind !== "worksheet") continue;
			for (const cell of other.cells()) {
				const { formula } = cell;
				if (formula === null || formula.text === null || formula.kind === "dataTable" || formula.kind === "array") continue;
				if (other.name === name && cell.row >= rows.first && cell.row <= rows.last + added) continue;
				const text = extendAreas(formula.text, {
					...change,
					formulaSheet: other.name
				});
				if (text === formula.text) continue;
				edits.push({
					sheet: other.name,
					cell: cell.ref,
					value: { formula: text }
				});
			}
		}
	}
	if (edits.length > 0) current = viaXlvane(() => editCells(current, edits, own).bytes, "Writing the repeated rows failed");
	if (rows.markers) {
		current = viaXlvane(() => deleteRows(current, {
			sheet: name,
			at: rows.close + added,
			...own
		}).bytes, "Deleting the row of the closing tag failed");
		current = viaXlvane(() => deleteRows(current, {
			sheet: name,
			at: rows.open,
			...own
		}).bytes, "Deleting the row of the opening tag failed");
	}
	return current;
}
// src/render.ts
const OPTION_NAMES = [
	"limits",
	"delimiters",
	"locale",
	"timeZone",
	"missing",
	"onWarning"
];
const ASYNC_OPTION_NAMES = [...OPTION_NAMES, "fetchImage"];
function option(options, name, fallback) {
	const value = options?.[name];
	if (value === void 0) return fallback;
	if (typeof value !== "string") throw new TvError("TV_E_INVALID_OPTION", `Option "${name}" must be a string.`, {
		hint: name === "locale" ? "Pass a BCP 47 locale such as \"pl-PL\"." : "Pass an IANA time zone such as \"Europe/Warsaw\".",
		details: { option: name }
	});
	try {
		if (name === "locale") Intl.getCanonicalLocales(value);
		else new Intl.DateTimeFormat("en-US", { timeZone: value });
	} catch (cause) {
		throw new TvError("TV_E_INVALID_OPTION", `"${clip(value)}" is not a valid ${name === "locale" ? "locale" : "time zone"}.`, {
			hint: name === "locale" ? "Pass a BCP 47 locale such as \"pl-PL\" or \"en-US\"." : "Pass an IANA time zone such as \"Europe/Warsaw\" or \"UTC\".",
			details: { option: name },
			cause
		});
	}
	return value;
}
function functionOption(options, name) {
	const value = options?.[name];
	if (value === void 0) return void 0;
	if (typeof value !== "function") throw new TvError("TV_E_INVALID_OPTION", `Option "${name}" must be a function.`, {
		hint: `Pass a function, or leave "${name}" out.`,
		details: { option: name }
	});
	return value;
}
/**
* The shared core of render, renderAsync and validate. With `collect`, only collects image
* addresses. With `dry`, renders without writing a file: image addresses are accepted unfetched,
* and the problems of the data go into `dry` instead of being thrown.
*/
function renderCore(template, data, options, names, fetched, collect, dry) {
	if (!(template instanceof Uint8Array)) throw new TvError("TV_E_INVALID_INPUT", "The template must be a Uint8Array (a Node.js Buffer is one).", { hint: "Read the file into bytes first, for example render(readFileSync(\"offer.docx\"), data)." });
	const limits = checkOptions(options, names, dry?.what ?? "render");
	const locale = option(options, "locale", "en-US");
	const timeZone = option(options, "timeZone", "UTC");
	const onWarning = functionOption(options, "onWarning");
	const missing = options?.missing ?? "error";
	if (missing !== "error" && missing !== "empty") throw new TvError("TV_E_INVALID_OPTION", "Option \"missing\" must be \"error\" or \"empty\".", {
		hint: "Leave it out to report missing fields, or pass \"empty\" to show nothing for them.",
		details: { option: "missing" }
	});
	const analysis = analyzeTemplate(template, limits, delimitersOption(options));
	if (dry !== void 0) dry.analysis = analysis;
	if (analysis.workbook !== null) {
		if (collect !== void 0 && dry === void 0) return /* @__PURE__ */ new Uint8Array(0);
		const evalCtx = {
			limits,
			formatting: new Formatting({
				locale,
				timeZone
			}),
			missing,
			problems: [],
			iterations: 0,
			output: 0,
			scopePaths: /* @__PURE__ */ new WeakMap()
		};
		const plan = planWorkbook(analysis.workbook, data, evalCtx);
		checkSectionWork(plan, template.length, limits);
		if (dry !== void 0) {
			dry.problems = evalCtx.problems;
			return /* @__PURE__ */ new Uint8Array(0);
		}
		if (evalCtx.problems.length > 0) throw aggregate("TV_E_DATA_INVALID", evalCtx.problems, evalCtx.problems.length >= limits.maxErrors);
		return applyWorkbook(template, plan, limits);
	}
	const { zip } = analysis.docx;
	const ctx = createContext(analysis, limits, {
		locale,
		timeZone,
		missing,
		fetched,
		collect: dry === void 0 ? collect : /* @__PURE__ */ new Set()
	});
	const pptx = analysis.docx.format === "pptx";
	const { prepared, problems: placement } = prepareStories(analysis, ctx);
	if (placement.length > 0) throw aggregate("TV_E_TEMPLATE_INVALID", placement, false);
	const changed = /* @__PURE__ */ new Map();
	const partNames = /* @__PURE__ */ new Map();
	if (pptx) for (const [part, xml] of renderPresentation(analysis, ctx, prepared, data)) {
		changed.set(partKey(part), xml);
		partNames.set(partKey(part), part);
	}
	else for (const [story, p] of prepared) {
		if (story.tags.length === 0) continue;
		changed.set(partKey(story.part), p.render([data]));
	}
	if (dry !== void 0) {
		dry.problems = ctx.problems;
		dry.warnings = ctx.warnings;
		return /* @__PURE__ */ new Uint8Array(0);
	}
	if (collect !== void 0 && ctx.problems.length === 0) return /* @__PURE__ */ new Uint8Array(0);
	if (ctx.problems.length > 0) throw aggregate("TV_E_DATA_INVALID", ctx.problems, ctx.problems.length >= limits.maxErrors);
	const extra = ctx.resources.parts();
	const numbering = ctx.numbering.changed();
	if (numbering !== null) {
		changed.set(partKey(numbering.name), numbering.xml);
		partNames.set(partKey(numbering.name), numbering.name);
	}
	for (const [key, part] of extra.changed) {
		changed.set(key, part.xml);
		partNames.set(key, part.name);
	}
	if (!pptx) {
		const stories = () => new Map(analysis.stories.map((s) => [partKey(s.part), changed.get(partKey(s.part)) ?? analysis.docx.pkg.readXml(s.part) ?? ""]));
		for (const [key, part] of pruneComments(analysis.docx.pkg, analysis.docx.main, stories, limits)) {
			changed.set(key, part.xml);
			partNames.set(key, part.name);
		}
	}
	ctx.output += ctx.resources.mediaBytes;
	if (ctx.output > limits.maxOutputBytes) throw new TvError("TV_E_LIMIT_EXCEEDED", `The result would be larger than ${String(limits.maxOutputBytes)} (XML characters and image bytes).`, {
		hint: `If you trust this template and data, raise the limit, for example { limits: { maxOutputBytes: ${String(limits.maxOutputBytes * 4)} } }.`,
		details: {
			limit: "maxOutputBytes",
			value: limits.maxOutputBytes
		}
	});
	const encoder = new TextEncoder();
	const written = /* @__PURE__ */ new Set();
	const removed = ctx.resources.removed;
	const kept = (name) => name.toLowerCase() === "[content_types].xml" || isPartName(zipNameToPartName(name)) && !name.includes(":");
	const entries = zip.entries.filter((entry) => kept(entry.name) && !removed.has(partKey(zipNameToPartName(entry.name)))).map((entry) => {
		const partName = zipNameToPartName(entry.name);
		const name = /^[\x20-\x7e]*$/.test(entry.name) ? entry.name : partNameToZipName(partName) ?? entry.name;
		const key = partKey(partName);
		written.add(key);
		const xml = changed.get(key);
		if (xml !== void 0) return {
			name,
			data: encoder.encode(xml)
		};
		const raw = zip.raw(entry);
		return raw === null ? {
			name,
			data: zip.read(entry)
		} : {
			name,
			data: /* @__PURE__ */ new Uint8Array(0),
			raw
		};
	});
	for (const [key, xml] of changed) {
		if (written.has(key)) continue;
		const name = partNameToZipName(partNames.get(key) ?? key);
		if (name !== null) entries.push({
			name,
			data: encoder.encode(xml)
		});
	}
	for (const [partName, bytes] of extra.added) {
		const name = partNameToZipName(partName);
		if (name !== null) entries.push({
			name,
			data: bytes
		});
	}
	for (const warning of ctx.warnings) onWarning?.(warning);
	return writeZip(entries);
}
/**
* Fills a template with data and returns the new file, in the format of the template (DOCX, PPTX
* or XLSX).
*
* Values are written with the formatting of the tag's first character; `{#list}…{/list}` repeats
* its content for each item (within a paragraph, over paragraphs, table rows or ranges of rows);
* `{#flag}…{/flag}` and `{^flag}…{/flag}` show content by the truth of a value; expressions and
* formatters work as described in the syntax reference. `{%logo}` inserts an image given as bytes
* or a data URL, `{~summary}` Markdown with the template's heading and list styles, and
* `{site | link}` a hyperlink. In a workbook, a cell that holds one value tag gets the type of the
* value, a section over cells of different rows repeats those rows with their formulas, and ranges
* that end on the last repeated row grow to the last copy. Template and data are untrusted: every
* value is escaped, nothing in the template can run code, nothing is downloaded, and the output is
* limited.
*
* Throws {@link TvError}: `TV_E_TEMPLATE_INVALID` for problems of the template (the same as
* {@link inspect} reports, plus `TV_E_SECTION_PLACEMENT`), `TV_E_DATA_INVALID` listing every
* missing field and invalid value in `errors` (an image given as an address is one: use
* {@link renderAsync}), `TV_E_LIMIT_EXCEEDED`, and the file errors of `inspect`.
*
* @example
* ```ts
* import { readFileSync, writeFileSync } from 'node:fs';
* import { render } from 'tplvane';
*
* // The template and image of examples/offer in the tplvane repository.
* const result = render(readFileSync('offer.docx'), {
*   logo: readFileSync('logo.png'),
*   number: '2026/11/001',
*   customer: { name: 'Anna Nowak', company: 'Nowak Design' },
*   date: '2026-11-06',
*   items: [{ name: 'Desk', qty: 1, price: 1200, total: 1200 }],
*   sum: 1200,
*   vip: true,
*   notes: 'We offer **free delivery**.\n\n- Assembly included',
*   author: 'Jan Kowalski',
* }, { locale: 'pl-PL' });
* writeFileSync('offer-anna.docx', result);
*
* // The same call fills a workbook: one row per item, and the totals below follow.
* writeFileSync('prices.xlsx', render(readFileSync('prices-template.xlsx'), { items: [{ name: 'Desk', price: 1200 }] }));
* ```
*/
function render(template, data, options) {
	return renderCore(template, data, options, OPTION_NAMES, /* @__PURE__ */ new Map());
}
/**
* Like {@link render}, but images given as http(s) addresses are fetched first with
* `options.fetchImage`, a function you pass and control, once per distinct address; then the
* document is rendered as by `render`. tplvane makes no connection itself: without `fetchImage`,
* an address in the data is reported as `TV_E_INVALID_VALUE`. A failing `fetchImage` gives
* `TV_E_IMAGE_FETCH` with the address; its bytes are checked like any image.
*
* @example
* ```ts
* import { readFileSync } from 'node:fs';
* import { renderAsync } from 'tplvane';
*
* const result = await renderAsync(readFileSync('offer.docx'), { logo: 'https://cdn.example.com/logo.png' }, {
*   async fetchImage(url) {
*     if (!url.startsWith('https://cdn.example.com/')) throw new Error('address not allowed');
*     return new Uint8Array(await (await fetch(url)).arrayBuffer());
*   },
* });
* console.log(result.length);
* ```
*/
async function renderAsync(template, data, options) {
	const fetchImage = functionOption(options, "fetchImage");
	const fetched = /* @__PURE__ */ new Map();
	if (fetchImage !== void 0) {
		const urls = /* @__PURE__ */ new Set();
		renderCore(template, data, options, ASYNC_OPTION_NAMES, fetched, urls);
		const limits = checkOptions(options, ASYNC_OPTION_NAMES, "render");
		if (urls.size > limits.maxImages) throw new TvError("TV_E_LIMIT_EXCEEDED", `The data has more than ${String(limits.maxImages)} image addresses.`, {
			hint: "Pass fewer images, or raise limits.maxImages if you trust the data.",
			details: {
				limit: "maxImages",
				value: limits.maxImages
			}
		});
		let total = 0;
		for (const url of urls) {
			let bytes;
			try {
				bytes = await fetchImage(url);
			} catch (cause) {
				throw new TvError("TV_E_IMAGE_FETCH", `fetchImage failed for "${clip(url)}".`, {
					hint: "Check the address and your fetchImage function; the error it threw is in \"cause\".",
					details: { url: clip(url) },
					cause
				});
			}
			if (!(bytes instanceof Uint8Array) && !(bytes instanceof ArrayBuffer)) throw new TvError("TV_E_IMAGE_FETCH", `fetchImage returned no bytes for "${clip(url)}".`, {
				hint: "Return a Uint8Array or an ArrayBuffer from fetchImage.",
				details: { url: clip(url) }
			});
			if (bytes.byteLength > limits.maxImageBytes) throw new TvError("TV_E_LIMIT_EXCEEDED", `The image fetched from "${clip(url)}" has more than ${String(limits.maxImageBytes)} bytes.`, {
				hint: "Use a smaller image, or raise limits.maxImageBytes if you trust the source.",
				details: {
					limit: "maxImageBytes",
					value: limits.maxImageBytes,
					url: clip(url)
				}
			});
			total += bytes.byteLength;
			if (total > limits.maxInputBytes) throw new TvError("TV_E_LIMIT_EXCEEDED", `The images fetched have more than ${String(limits.maxInputBytes)} bytes in all.`, {
				hint: "Use fewer or smaller images, or raise limits.maxInputBytes if you trust the sources.",
				details: {
					limit: "maxInputBytes",
					value: limits.maxInputBytes
				}
			});
			fetched.set(url, bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
		}
	}
	return renderCore(template, data, options, ASYNC_OPTION_NAMES, fetched);
}
/** Renders without writing a file, for validate (D-047). Throws only for problems of the input, the options, the file and the template. */
function dryRender(template, data, options, what) {
	if (!(template instanceof Uint8Array)) throw new TvError("TV_E_INVALID_INPUT", "The template must be a Uint8Array (a Node.js Buffer is one).", { hint: `Read the file into bytes first, for example ${what}(readFileSync("offer.docx"), data).` });
	const dry = {
		problems: [],
		warnings: [],
		analysis: null,
		what
	};
	renderCore(template, data, options, [
		"limits",
		"delimiters",
		"locale",
		"timeZone",
		"missing"
	], /* @__PURE__ */ new Map(), void 0, dry);
	return dry;
}
// src/preview.ts
const OBJECT = "￼";
function cellText(cell) {
	if (cell.formula !== null && cell.type === "empty") return `=${cell.formula.text ?? ""}`;
	const v = cell.value;
	if (v === null) return "";
	if (v instanceof Date) {
		const iso = v.toISOString();
		return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso.slice(0, 19).replace(/\.000$/, "");
	}
	if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
	return String(v);
}
/**
* Returns the text of a Word document, a PowerPoint presentation or an Excel workbook, part by part:
* the paragraphs of the body, headers, footers and notes of a document (tables cell by cell);
* the paragraphs of each slide and its notes; the rows of each sheet with cells separated by tabs,
* dates as ISO 8601 text, and formulas without a stored result (as `render` writes them) as
* `=formula`. Pictures and other objects in text show as `[object]`. Use it to check a rendered
* result without an office application.
*
* Throws {@link TvError} for input that is not such a file, as {@link inspect} does.
*
* @example
* ```ts
* import { readFileSync } from 'node:fs';
* import { preview } from 'tplvane';
*
* console.log(preview(readFileSync('offer-filled.docx')).text);
* ```
*/
function preview(document, options) {
	if (!(document instanceof Uint8Array)) throw new TvError("TV_E_INVALID_INPUT", "The document must be a Uint8Array (a Node.js Buffer is one).", { hint: "Read the file into bytes first, for example preview(readFileSync(\"offer-filled.docx\"))." });
	const limits = checkOptions(options, ["limits"], "preview");
	const docx = openDocx(document, limits);
	const parts = [];
	let size = 0;
	const add = (n) => {
		size += n;
		if (size > limits.maxOutputBytes) throw limitError("maxOutputBytes", limits.maxOutputBytes, `The text would be longer than ${String(limits.maxOutputBytes)} characters.`);
	};
	if (docx.format === "xlsx") {
		const workbook = viaXlvane(() => readWorkbook(document, { limits: xlvaneLimits(limits) }), "Reading the workbook failed");
		for (const sheet of workbook.sheets) {
			if (sheet.kind !== "worksheet") continue;
			const lines = [];
			for (const row of sheet.rows()) {
				const first = row[0];
				if (first === void 0) continue;
				while (lines.length < first.row - 1) {
					add(1);
					lines.push("");
				}
				const cells = [];
				for (const cell of row) {
					while (cells.length < cell.col - 1) {
						add(1);
						cells.push("");
					}
					const text = cellText(cell);
					add(text.length + 1);
					cells.push(text);
				}
				while (cells.length > 0 && cells[cells.length - 1] === "") cells.pop();
				lines.push(cells.join("	"));
			}
			while (lines.length > 0 && lines[lines.length - 1] === "") lines.pop();
			parts.push({
				name: `sheet ${sheet.name}`,
				lines
			});
		}
	} else for (const story of docx.stories) {
		const xml = docx.pkg.readXml(story.part);
		if (xml === void 0) continue;
		const doc = parseRawTree(xml, {
			part: story.part,
			limits
		});
		const paragraphs = docx.format === "pptx" ? walkSlideParagraphs(doc, {
			part: story.part,
			story: story.story,
			slide: story.slide ?? 0
		}) : walkParagraphs(doc, story);
		const name = story.story === "slide" ? `slide ${String(story.slide ?? 0)}` : story.story === "notes" ? `slide ${String(story.slide ?? 0)} notes` : story.story === "header" || story.story === "footer" ? `${story.story} ${story.part}` : story.story === "footnote" || story.story === "endnote" ? `${story.story}s` : story.story;
		const lines = paragraphs.map((p) => p.text.split(OBJECT).join("[object]"));
		for (const line of lines) add(line.length + 1);
		parts.push({
			name,
			lines
		});
	}
	const text = parts.map((part) => `== ${part.name}\n${part.lines.map((l) => `${l}\n`).join("")}`).join("");
	return {
		format: docx.format,
		parts,
		text
	};
}
// src/validate/json-schema.ts
/** The JSON kind of a JavaScript value, or `other` for functions and symbols. */
function kindOf(value) {
	if (value === null) return "null";
	if (Array.isArray(value)) return "array";
	if (value instanceof Date || value instanceof Uint8Array || value instanceof ArrayBuffer) return "string";
	switch (typeof value) {
		case "string": return "string";
		case "number":
		case "bigint": return "number";
		case "boolean": return "boolean";
		case "object": return "object";
		default: return "other";
	}
}
function words(list) {
	return list.length <= 1 ? list[0] ?? "anything" : `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
}
/** The kinds a schema allows; an empty list allows anything. */
function typesOf(schema) {
	if (schema.type !== void 0) return typeof schema.type === "string" ? [schema.type] : schema.type;
	if (schema.anyOf !== void 0) return [...new Set(schema.anyOf.flatMap((s) => s.type === void 0 ? [] : typeof s.type === "string" ? [s.type] : s.type))];
	return [];
}
/** Checks `value` against `schema` and returns the issues, in the order of the schema. */
function checkSchema(value, schema, options = {}) {
	const maxIssues = options.maxIssues ?? 100;
	let budget = options.maxNodes ?? 1e6;
	const issues = [];
	const walk = (v, s, path, out) => {
		if (out.length >= maxIssues || --budget < 0) return;
		if (s.anyOf !== void 0) {
			for (const option of s.anyOf) {
				const trial = [];
				walk(v, option, path, trial);
				if (trial.length === 0) return;
			}
			out.push({
				path,
				kind: "type",
				expected: words(typesOf(s)),
				found: kindOf(v),
				schema: s
			});
			return;
		}
		const allowed = typesOf(s);
		const kind = kindOf(v);
		if (allowed.length > 0 && !allowed.includes(kind)) {
			out.push({
				path,
				kind: "type",
				expected: words(allowed),
				found: kind,
				schema: s
			});
			return;
		}
		if (kind === "object" && s.properties !== void 0) {
			const required = new Set(s.required ?? []);
			for (const name of Object.keys(s.properties)) {
				const field = ownValue(v, name);
				const fieldSchema = s.properties[name];
				if (field === void 0) {
					if (required.has(name) && out.length < maxIssues) out.push({
						path: [...path, name],
						kind: "missing",
						expected: words(typesOf(fieldSchema)),
						schema: fieldSchema
					});
					continue;
				}
				walk(field, fieldSchema, [...path, name], out);
			}
		}
		if (kind === "array" && s.items !== void 0) {
			const list = v;
			for (let i = 0; i < list.length && out.length < maxIssues && budget > 0; i++) walk(ownValue(list, String(i)), s.items, [...path, i], out);
		}
	};
	walk(value, schema, [], issues);
	return issues;
}
// src/validate.ts
function schemaWarning(issue, where) {
	const path = dataPathOf(issue.path);
	const name = path === "" ? "The data" : `"${clip(path)}"`;
	return {
		code: "TV_W_DATA_SCHEMA",
		message: issue.kind === "missing" ? `${name} is missing; the schema of the template needs it (${issue.expected}), but render does not use it with this data.` : `${name} is ${issue.found === "null" ? "null" : `${/^[aeiou]/.test(issue.found ?? "value") ? "an" : "a"} ${issue.found ?? "value"}`}, but the schema of the template expects ${issue.expected}; render accepts it with this data.`,
		location: where,
		hint: issue.kind === "missing" ? "Nothing fails now. Add the field if the section that uses it can be shown with other data." : "Nothing fails now. Pass the kind the schema names, so that the data also fits other templates and tools that read the schema.",
		details: {
			pointer: clip(pointerOf(issue.path)),
			dataPath: clip(path),
			kind: issue.kind,
			expected: issue.expected
		}
	};
}
/** Whether the last name of `path` is an own field of an object around the list item that lacks it. */
function foundOutside(data, path) {
	const name = path[path.length - 1];
	if (typeof name !== "string" || !path.some((p) => typeof p === "number")) return false;
	const holders = [data];
	let at = data;
	for (const step of path.slice(0, -1)) {
		at = ownValue(at, String(step));
		holders.push(at);
	}
	const lastIndex = path.map((p) => typeof p).lastIndexOf("number");
	return holders.slice(0, lastIndex).some((h) => h !== null && typeof h === "object" && !Array.isArray(h) && ownValue(h, name) !== void 0);
}
/**
* Checks data against a template without writing a file. `errors` are exactly the problems that
* {@link render} would report for this data (a dry run, with the same options), each with the place
* of its tag in the template, the path of the field in the data (`details.dataPath` such as
* `items[1].name`, `details.pointer` such as `/items/1/name`) and, for a missing field, the fields
* that are there (`details.available`). The data is also checked against the JSON Schema that
* {@link inspect} returns; what does not fit it but does not fail render is a `TV_W_DATA_SCHEMA`
* warning. Images given as addresses are not fetched and count as valid.
*
* Throws {@link TvError} only for problems of the input, the options, the file and the template
* (`TV_E_TEMPLATE_INVALID` and the file errors of `inspect`), and for limits.
*
* @example
* ```ts
* import { readFileSync } from 'node:fs';
* import { validate } from 'tplvane';
*
* const result = validate(readFileSync('offer.docx'), { customer: {}, items: [{ name: 'Desk' }] });
* if (!result.valid) {
*   for (const error of result.errors) console.log(`${String(error.details['dataPath'])}: ${error.message} ${error.hint}`);
* }
* ```
*/
function validate(template, data, options) {
	const dry = dryRender(template, data, options, "validate");
	const analysis = dry.analysis;
	const { schema, locations } = inferSchema(analysis === null ? [] : analysis.trees);
	const errors = dry.problems;
	const confirmed = new Set(errors.map((e) => e.details["pointer"]).filter((p) => typeof p === "string"));
	const covered = (pointer) => {
		for (const p of confirmed) if (p === pointer || p.startsWith(`${pointer}/`) || pointer.startsWith(`${p}/`)) return true;
		return false;
	};
	const limits = { maxIssues: options?.limits?.maxErrors ?? 100 };
	const warnings = [];
	for (const issue of checkSchema(data, schema, limits)) {
		if (covered(pointerOf(issue.path))) continue;
		if (issue.kind === "missing" && foundOutside(data, issue.path)) continue;
		warnings.push(schemaWarning(issue, locations.get(issue.schema) ?? {}));
	}
	warnings.push(...dry.warnings);
	return {
		valid: errors.length === 0,
		errors,
		warnings,
		schema
	};
}
/**
* Returns the template as a {@link TemplateSchema}: a Standard Schema of its data with the JSON
* Schema of {@link inspect}. Problems of the template are thrown here, at once.
*
* @example
* ```ts
* import { readFileSync } from 'node:fs';
* import { templateSchema } from 'tplvane';
*
* const schema = templateSchema(readFileSync('offer.docx'), { missing: 'empty' });
* console.log(schema['~standard'].vendor, Object.keys(schema.jsonSchema.properties ?? {}));
* ```
*/
function templateSchema(template, options) {
	const { schema } = validate(template, void 0, options);
	return {
		jsonSchema: schema,
		"~standard": {
			version: 1,
			vendor: "tplvane",
			validate(value) {
				const result = validate(template, value, options);
				if (result.valid) return { value };
				return { issues: result.errors.map((e) => {
					const pointer = e.details["pointer"];
					return {
						message: `${e.message} ${e.hint}`,
						...typeof pointer === "string" && pointer !== "" ? { path: pointer.slice(1).split("/").map(decodeSegment) } : {}
					};
				}) };
			}
		}
	};
}
/** A segment of a JSON Pointer as a key: digits are list indexes. */
function decodeSegment(segment) {
	const text = segment.replace(/~1/g, "/").replace(/~0/g, "~");
	return /^(0|[1-9][0-9]{0,8})$/.test(text) ? Number(text) : text;
}
export { renderAsync as a, delimiterProblem as c, suggest as d, ERROR_CODES as f, VERSION$1 as g, isTvError as h, render as i, clip as l, WARNING_CODES as m, validate as n, ownValue as o, TvError as p, preview as r, inspect as s, templateSchema as t, DEFAULT_LIMITS as u };
