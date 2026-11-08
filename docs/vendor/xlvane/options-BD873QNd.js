import { Inflate, deflateSync } from "../fflate/browser.js";
// src/errors.ts
/** Base URL of the error catalogue. Every code has a heading (and so an anchor) there. */
const DOCS_URL = "https://github.com/DevXord/xlvane/blob/main/docs/errors.md";
/**
* Every error code xlvane can throw. Codes are stable: a code is never reused for a different
* condition. Each one is described in docs/errors.md.
*
* @example
* ```ts
* import { ERROR_CODES } from 'xlvane';
* console.log(ERROR_CODES.includes('XLV_E_SHEET_NOT_FOUND')); // true
* ```
*/
const ERROR_CODES = [
	"XLV_E_INVALID_INPUT",
	"XLV_E_INVALID_OPTION",
	"XLV_E_INVALID_ARGUMENT",
	"XLV_E_LIMIT_EXCEEDED",
	"XLV_E_NOT_ZIP",
	"XLV_E_UNSUPPORTED_CONTAINER",
	"XLV_E_ZIP_CORRUPT",
	"XLV_E_ZIP_UNSUPPORTED",
	"XLV_E_XML_MALFORMED",
	"XLV_E_XML_DTD_FORBIDDEN",
	"XLV_E_XML_ENTITY_FORBIDDEN",
	"XLV_E_XML_ENCODING",
	"XLV_E_PART_MISSING",
	"XLV_E_PART_INVALID",
	"XLV_E_INVALID_CELL_REF",
	"XLV_E_SHEET_NOT_FOUND",
	"XLV_E_TABLE_NOT_FOUND",
	"XLV_E_FILE_READ",
	"XLV_E_INVALID_VALUE",
	"XLV_E_INVALID_STYLE",
	"XLV_E_INVALID_SHEET_NAME",
	"XLV_E_DUPLICATE_SHEET_NAME",
	"XLV_E_INVALID_DEFINED_NAME",
	"XLV_E_INVALID_HYPERLINK",
	"XLV_E_INVALID_TABLE",
	"XLV_E_INVALID_IMAGE",
	"XLV_E_INVALID_CHART",
	"XLV_E_OVERLAPPING_MERGE",
	"XLV_E_FILE_EXISTS",
	"XLV_E_FILE_WRITE",
	"XLV_E_HEADER_NOT_FOUND",
	"XLV_E_SCHEMA_INVALID",
	"XLV_E_ASYNC_SCHEMA",
	"XLV_E_INVALID_ROWS",
	"XLV_E_JSON_INVALID",
	"XLV_E_CSV_MALFORMED",
	"XLV_E_NOT_SUPPORTED",
	"XLV_E_USAGE",
	"XLV_E_NOT_INSTALLED",
	"XLV_E_INTERNAL",
	"XLV_E_PATH_NOT_ALLOWED"
];
/**
* Every warning code xlvane can report. Warnings never stop reading; they describe something
* in the file that the caller may want to know about. Each one is described in docs/errors.md.
*
* @example
* ```ts
* import { WARNING_CODES } from 'xlvane';
* console.log(WARNING_CODES.includes('XLV_W_MACROS_PRESENT')); // true
* ```
*/
const WARNING_CODES = [
	"XLV_W_DIMENSION_MISMATCH",
	"XLV_W_SHARED_FORMULA_NOT_EXPANDED",
	"XLV_W_FORMULA_NO_CACHED_VALUE",
	"XLV_W_VALUES_IN_MERGED_CELLS",
	"XLV_W_VERY_HIDDEN_SHEET",
	"XLV_W_MACROS_PRESENT",
	"XLV_W_EXTERNAL_LINKS",
	"XLV_W_DUPLICATE_CELL",
	"XLV_W_INVALID_VALUE",
	"XLV_W_SHEET_PART_MISSING",
	"XLV_W_SHEET_PART_SHARED",
	"XLV_W_PART_MISSING",
	"XLV_W_FORMULAS_NOT_RECALCULATED",
	"XLV_W_INVALID_CELL_REF",
	"XLV_W_UNUSUAL_HYPERLINK",
	"XLV_W_THREADED_COMMENTS",
	"XLV_W_DRAWING_CONTENT_NOT_READ",
	"XLV_W_MERGED_CELL_WRITE",
	"XLV_W_LIMIT_REACHED"
];
/**
* The only error type thrown by xlvane for problems with input files or options. It says what
* happened (`message`), where (`location`), what to do next (`hint`) and links to the
* description of the code (`docsUrl`). `toJSON()` returns the same information as data.
*
* @example
* ```ts
* import { readWorkbook, isXlvError } from 'xlvane';
*
* try {
*   readWorkbook(new Uint8Array([1, 2, 3]));
* } catch (error) {
*   if (isXlvError(error)) console.log(error.code, error.hint); // "XLV_E_NOT_ZIP", "..."
* }
* ```
*/
var XlvError = class extends Error {
	name = "XlvError";
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
	constructor(code, message, init) {
		super(message, init.cause === void 0 ? void 0 : { cause: init.cause });
		this.code = code;
		this.location = { ...init.location };
		this.hint = init.hint;
		this.docsUrl = `${DOCS_URL}#${code.toLowerCase()}`;
		this.details = { ...init.details };
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
			details: { ...this.details }
		};
	}
};
/**
* Returns `true` if `value` is an {@link XlvError}.
*
* @example
* ```ts
* import { isXlvError } from 'xlvane';
* console.log(isXlvError(new Error('x'))); // false
* ```
*/
function isXlvError(value) {
	return value instanceof XlvError;
}
// src/sml/xstring.ts
const UNDERSCORE$1 = 95;
const LOWER_X = 120;
/** Length of one escape sequence, `_xHHHH_`. */
const ESCAPE_LENGTH = 7;
function hexValue$2(c) {
	if (c >= 48 && c <= 57) return c - 48;
	if (c >= 65 && c <= 70) return c - 55;
	if (c >= 97 && c <= 102) return c - 87;
	return -1;
}
/** The code unit escaped at `i` (`_xHHHH_`), or -1 if there is no escape at `i`. */
function escapeAt(s, i) {
	if (s.charCodeAt(i) !== UNDERSCORE$1 || s.charCodeAt(i + 1) !== LOWER_X || s.charCodeAt(i + 6) !== UNDERSCORE$1) return -1;
	let value = 0;
	for (let k = 2; k < 6; k++) {
		const d = hexValue$2(s.charCodeAt(i + k));
		if (d < 0) return -1;
		value = value * 16 + d;
	}
	return value;
}
/**
* Decodes the `_xHHHH_` escapes of an ST_Xstring value (ECMA-376 Part 1 §22.9.2.19), for
* example `"a_x000D__x000A_b"` → `"a\r\nb"` and `"_x005F_x0041_"` → `"_x0041_"`. Sequences that
* are not an escape, including escapes of lone surrogates, are kept as text. Runs in linear
* time.
*/
function decodeXstring(value) {
	let at = value.indexOf("_x");
	if (at === -1) return value;
	let out = "";
	let last = 0;
	while (at !== -1) {
		const unit = escapeAt(value, at);
		let length = 0;
		let decoded = "";
		if (unit >= 55296 && unit <= 56319) {
			const low = escapeAt(value, at + ESCAPE_LENGTH);
			if (low >= 56320 && low <= 57343) {
				decoded = String.fromCharCode(unit, low);
				length = 14;
			}
		} else if (unit !== -1 && (unit < 56320 || unit > 57343)) {
			decoded = String.fromCharCode(unit);
			length = ESCAPE_LENGTH;
		}
		if (length === 0) {
			at = value.indexOf("_x", at + 1);
			continue;
		}
		out += value.slice(last, at) + decoded;
		last = at + length;
		at = value.indexOf("_x", last);
	}
	return out + value.slice(last);
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
function hex4(code) {
	return code.toString(16).toUpperCase().padStart(4, "0");
}
/**
* Encodes a string as an ST_Xstring value: characters that XML 1.0 cannot represent become
* `_xHHHH_`, and an underscore that starts `_x` + four hex digits + `_` becomes `_x005F_` (also
* when the character after the digits is one that is escaped, since its escape starts with `_`), so
* `decodeXstring(encodeXstring(s)) === s` for every well-formed string. Lone surrogates must be
* rejected by the caller first ({@link firstLoneSurrogate}); they are encoded here as
* escapes, which the reader keeps as text (D-XS-3).
*/
function encodeXstring(s) {
	if (!/[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF￾￿]|_x[0-9A-Fa-f]{4}_/.test(s)) return s;
	/** Whether the character at `j` is written as an escape (so its output starts with `_`). */
	const escaped = (j) => {
		const c = s.charCodeAt(j);
		if (c >= 55296 && c <= 56319) return !(s.charCodeAt(j + 1) >= 56320 && s.charCodeAt(j + 1) <= 57343);
		return c < 32 && c !== 9 && c !== 10 && c !== 13 || c >= 56320 && c <= 57343 || c === 65534 || c === 65535;
	};
	let out = "";
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c >= 55296 && c <= 56319) {
			const next = s.charCodeAt(i + 1);
			if (next >= 56320 && next <= 57343) {
				out += s[i] + s[i + 1];
				i++;
				continue;
			}
			out += `_x${hex4(c)}_`;
		} else if (c < 32 && c !== 9 && c !== 10 && c !== 13 || c >= 56320 && c <= 57343 || c === 65534 || c === 65535) out += `_x${hex4(c)}_`;
		else if (c === 95 && /^_x[0-9A-Fa-f]{4}/.test(s.slice(i, i + 6)) && (s[i + 6] === "_" || i + 6 < s.length && escaped(i + 6))) out += "_x005F_";
		else out += s[i];
	}
	return out;
}
/** Index of the first unpaired surrogate code unit in `s`, or -1 if `s` is well-formed UTF-16. */
function firstLoneSurrogate(s) {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c < 55296 || c > 57343) continue;
		if (c <= 56319) {
			const next = s.charCodeAt(i + 1);
			if (next >= 56320 && next <= 57343) {
				i++;
				continue;
			}
		}
		return i;
	}
	return -1;
}
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
/** The Char production of XML 1.0 for a code point. */
function isXmlChar(c) {
	return c === TAB || c === LF || c === CR || c >= 32 && c <= 55295 || c >= 57344 && c <= 65533 || c >= 65536 && c <= 1114111;
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
* Throws {@link XlvError} with `XLV_E_XML_MALFORMED`, `XLV_E_XML_DTD_FORBIDDEN`,
* `XLV_E_XML_ENTITY_FORBIDDEN` or `XLV_E_LIMIT_EXCEEDED`.
*/
function parseXml(xml, handler, options) {
	const parser = new Parser(handler, options, true);
	parser.write(xml);
	parser.end();
}
/**
* Creates a parser for a document given in pieces of text. Memory is bounded by the largest
* single token (a tag, comment, processing instruction, CDATA section or text run) and not by
* the size of the document: consumed input is dropped.
*/
function createXmlStreamParser(handler, options) {
	const parser = new Parser(handler, options, false);
	return {
		write: (chunk) => parser.write(chunk),
		end: () => parser.end(),
		get buffered() {
			return parser.buffered;
		}
	};
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
				throw new XlvError("XLV_E_LIMIT_EXCEEDED", message, {
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
					throw new XlvError("XLV_E_XML_ENTITY_FORBIDDEN", `Reference to entity "&${body};", which is not one of the five predefined XML entities.`, {
						location: this.location(base + amp),
						hint: "Only &lt; &gt; &amp; &apos; &quot; and character references such as &#160; are allowed, because xlvane does not process DTDs. Re-save the file in a spreadsheet application.",
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
		throw new XlvError("XLV_E_XML_DTD_FORBIDDEN", "The part contains a document type declaration (DTD).", {
			location: this.location(this.pos),
			hint: "xlvane rejects DTDs to protect against external entities (XXE) and entity expansion attacks. Office Open XML parts do not use DTDs; if the file was generated by a program, report the problem to its authors."
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
		throw new XlvError("XLV_E_XML_MALFORMED", message, {
			location: this.location(offset),
			hint: "The part is not well-formed XML. Re-save the file in a spreadsheet application; if it was generated by a program, report the problem to its authors."
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
		throw new XlvError("XLV_E_LIMIT_EXCEEDED", message, {
			location: this.location(offset),
			hint: `If you trust this file, raise the limit, for example { limits: { ${limit}: ${String(this.options.limits[limit] * 4)} } }.`,
			details
		});
	}
};
/**
* Decodes the bytes of an XML part. Office Open XML packaging parts are UTF-8 or UTF-16
* (docs/facts/ecma-376.md E-55); a byte order mark selects UTF-16, otherwise UTF-8 is used. Invalid byte sequences
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
		throw new XlvError("XLV_E_XML_ENCODING", `The part is not valid ${encoding.toUpperCase()}.`, {
			location: { part },
			hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.",
			cause
		});
	}
	const declared = declaredEncoding(text);
	if (declared !== null) {
		const normalized = declared.toLowerCase();
		if (!(encoding === "utf-8" ? normalized === "utf-8" : normalized === "utf-16" || normalized === encoding)) throw new XlvError("XLV_E_XML_ENCODING", `The part declares encoding "${clip(declared)}", which is not supported.`, {
			location: { part },
			hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.",
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
/**
* Creates a parser for the bytes of a part given in pieces: the streaming equivalent of
* {@link decodeXml} followed by {@link parseXml}. A byte order mark selects UTF-16, otherwise
* UTF-8 is used; invalid byte sequences and a declared encoding other than the detected one
* throw `XLV_E_XML_ENCODING` with the same messages as `decodeXml`.
*/
function createXmlByteStreamParser(handler, options) {
	const text = createXmlStreamParser(handler, options);
	const part = options.part;
	let encoding = null;
	let decoder = null;
	/** Bytes held until the encoding is known (at most the first two). */
	let head = /* @__PURE__ */ new Uint8Array(0);
	/**
	* Decoded pieces held until the XML declaration can be checked (null once it has been), with
	* their length, their first five characters and whether "?>" has been seen. Each piece is
	* searched once, so a long declaration costs linear time (D-H01).
	*/
	let pending = [];
	let pendingLength = 0;
	let pendingHead = "";
	let pendingLast = "";
	let declarationEnd = false;
	const encodingError = (cause) => new XlvError("XLV_E_XML_ENCODING", `The part is not valid ${(encoding ?? "utf-8").toUpperCase()}.`, {
		location: { part },
		hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.",
		cause
	});
	const decode = (bytes, stream) => {
		try {
			return (decoder ?? new TextDecoder("utf-8", { fatal: true })).decode(bytes, { stream });
		} catch (cause) {
			throw encodingError(cause);
		}
	};
	/** Passes decoded text on, after checking the declaration once enough text has arrived. */
	const feed = (chunk, final) => {
		if (pending === null) {
			if (chunk.length > 0) text.write(chunk);
			return;
		}
		if (chunk.length > 0) {
			pending.push(chunk);
			pendingLength += chunk.length;
			if (pendingHead.length < 5) pendingHead += chunk.slice(0, 5 - pendingHead.length);
			declarationEnd ||= (pendingLast + chunk).includes("?>");
			pendingLast = chunk.slice(-1);
		}
		const known = pendingLength >= 5 && (!pendingHead.startsWith("<?xml") || declarationEnd);
		if (!known && pendingLength - 2 > options.limits.maxXmlTextLength + 1) throw new XlvError("XLV_E_LIMIT_EXCEEDED", PI_TOO_LONG, {
			location: {
				part,
				line: 1,
				column: 1
			},
			hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTextLength: ${String(options.limits.maxXmlTextLength * 4)} } }.`,
			details: {
				limit: "maxXmlTextLength",
				value: options.limits.maxXmlTextLength
			}
		});
		if (!known && !final) return;
		const all = pending.join("");
		pending = null;
		const declared = declaredEncoding(all);
		if (declared !== null) {
			const normalized = declared.toLowerCase();
			if (!(encoding === "utf-8" ? normalized === "utf-8" : normalized === "utf-16" || normalized === encoding)) throw new XlvError("XLV_E_XML_ENCODING", `The part declares encoding "${declared}", which is not supported.`, {
				location: { part },
				hint: "Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.",
				details: { encoding: declared }
			});
		}
		if (all.length > 0) text.write(all);
	};
	const start = (bytes) => {
		if (bytes[0] === 255 && bytes[1] === 254) encoding = "utf-16le";
		else if (bytes[0] === 254 && bytes[1] === 255) encoding = "utf-16be";
		else encoding = "utf-8";
		decoder = new TextDecoder(encoding, { fatal: true });
		return bytes;
	};
	return {
		write(bytes) {
			if (encoding === null) {
				const joined = new Uint8Array(head.length + bytes.length);
				joined.set(head);
				joined.set(bytes, head.length);
				if (joined.length < 2) {
					head = joined;
					return;
				}
				head = /* @__PURE__ */ new Uint8Array(0);
				feed(decode(start(joined), true), false);
				return;
			}
			feed(decode(bytes, true), false);
		},
		end() {
			if (encoding === null) start(head);
			feed(decode(head, false), true);
			head = /* @__PURE__ */ new Uint8Array(0);
			text.end();
		}
	};
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
			if (count > options.limits.maxXmlTreeElements) throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The part has more elements and attributes than the limit.", {
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
/** Returns the direct children of `element` with the given namespace and local name. */
function childrenNamed(element, ns, local) {
	return element.children.filter((c) => c.local === local && c.ns === ns);
}
/** Returns the first direct child of `element` with the given namespace and local name. */
function childNamed(element, ns, local) {
	return element.children.find((c) => c.local === local && c.ns === ns);
}
// src/doc-properties.ts
/** The core properties namespace (Part 2 Table E.1). */
const CORE_PROPERTIES_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
/** Namespace of the ISO 15836-1 elements (Part 2 §8.3.4.2). */
const DC_NS = "http://purl.org/dc/elements/1.1/";
/** Namespace of the ISO 15836-2 elements (Part 2 §8.3.4.3). */
const DCTERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
/** The core properties relationship type (Part 2 Table E.3; [MS-OI29500] §2.1.30). */
const CORE_PROPERTIES_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
/** The type ECMA-376 Part 4 §13.2.10 lists for Transitional documents; accepted when reading. */
const CORE_PROPERTIES_REL_P4 = "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";
/** Media type of the Core Properties part (Part 2 Table E.2). */
const CORE_PROPERTIES_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
/** Part name used by the writer (a convention; parts are found through relationships, E-08). */
const CORE_PROPERTIES_PART = "docProps/core.xml";
/** `true` for a relationship type that points to the Core Properties part. */
function isCorePropertiesRel(type) {
	return type === "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" || type === CORE_PROPERTIES_REL_P4;
}
/** Property name, element namespace prefix and local name, in the order of CT_CoreProperties. */
const ELEMENTS = [
	["category", "cp"],
	["contentStatus", "cp"],
	["created", "dcterms"],
	["creator", "dc"],
	["description", "dc"],
	["identifier", "dc"],
	["keywords", "cp"],
	["language", "dc"],
	["lastModifiedBy", "cp"],
	["lastPrinted", "cp"],
	["modified", "dcterms"],
	["revision", "cp"],
	["subject", "dc"],
	["title", "dc"],
	["version", "cp"]
];
const DATE_KEYS = [
	"created",
	"modified",
	"lastPrinted"
];
/** Extended properties and their element names (Part 1 §22.2.2.5, .15, .1, .2). */
const APP_ELEMENTS = [
	["application", "Application"],
	["appVersion", "AppVersion"],
	["company", "Company"],
	["manager", "Manager"]
];
const KEYS$1 = [...ELEMENTS.map(([key]) => key), ...APP_ELEMENTS.map(([key]) => key)];
/** Office's pattern for AppVersion ([MS-OI29500] §2.1.1717 a; Part 1 §22.2.2.2: `XX.YYYY`). */
const APP_VERSION_RE = /^[0-9]{1,2}\.[0-9]{1,4}$/;
const PREFIX_NS = {
	cp: CORE_PROPERTIES_NS,
	dc: DC_NS,
	dcterms: DCTERMS_NS
};
function invalidProperty(key, message, hint) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		hint,
		details: { option: `properties.${key}` }
	});
}
/**
* Checks document properties given by a caller and returns a copy with only the given ones.
* Text must be a string that XML can hold; dates must be valid `Date` objects in the years
* 0–9999 (W3CDTF uses four-digit years).
*/
function normalizeProperties(value) {
	if (value === void 0) return {};
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"properties\" must be an object.", {
		hint: "For example { properties: { title: \"Q3 sales\", creator: \"Finance\" } }.",
		details: { option: "properties" }
	});
	const out = {};
	for (const key of Object.keys(value)) {
		if (!KEYS$1.includes(key)) throw new XlvError("XLV_E_INVALID_OPTION", `Unknown document property "${key.slice(0, 40)}".`, {
			hint: `Known properties: ${KEYS$1.join(", ")}.`,
			details: {
				option: `properties.${key.slice(0, 40)}`,
				known: KEYS$1
			}
		});
		const v = value[key];
		if (v === void 0) continue;
		if (DATE_KEYS.includes(key)) {
			if (!(v instanceof Date) || Number.isNaN(v.getTime()) || v.getUTCFullYear() < 0 || v.getUTCFullYear() > 9999) throw invalidProperty(key, `The document property "${key}" must be a valid Date between the years 0 and 9999.`, `For example ${key}: new Date().`);
			out[key] = new Date(v.getTime());
		} else {
			if (typeof v !== "string") throw invalidProperty(key, `The document property "${key}" must be a string.`, `For example ${key}: "text".`);
			if (firstLoneSurrogate(v) !== -1 || /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/.test(v)) throw invalidProperty(key, `The document property "${key}" contains a character that XML cannot store.`, "Remove control characters (other than tab, line feed and carriage return) and unpaired surrogates.");
			if (key === "appVersion" && !APP_VERSION_RE.test(v)) throw invalidProperty(key, `The document property "appVersion" is "${v.slice(0, 20)}"; it must be one or two digits, a dot and one to four digits.`, "For example appVersion: \"1.0\".");
			out[key] = v;
		}
	}
	return out;
}
/** A date as W3CDTF in UTC with whole seconds, e.g. `2026-10-30T12:00:00Z` (F-OPC-19). */
function w3cdtf(d) {
	return `${d.toISOString().slice(0, 19)}Z`;
}
/** The Core Properties part, or `null` when no property is set. */
function corePropertiesXml(properties) {
	let body = "";
	for (const [key, prefix] of ELEMENTS) {
		const v = properties[key];
		if (v === void 0) continue;
		const name = `${prefix}:${key}`;
		if (v instanceof Date) body += `<${name}${prefix === "dcterms" ? " xsi:type=\"dcterms:W3CDTF\"" : ""}>${w3cdtf(v)}</${name}>`;
		else body += `<${name}>${escapeText(v)}</${name}>`;
	}
	if (body === "") return null;
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:xsi="${XSI_NS}">${body}</cp:coreProperties>`;
}
const DATE_RE = /^(\d{4})(?:-(\d{2})(?:-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?)?)?$/;
/**
* Parses a W3CDTF / xsd:dateTime value: `YYYY`, `YYYY-MM`, `YYYY-MM-DD` or a date with time,
* optional seconds and fraction, and `Z` or `±hh:mm`. A value without a zone is read as UTC.
* Returns `null` for anything else.
*/
function parseW3cdtf(text) {
	const m = DATE_RE.exec(text.trim());
	if (m === null) return null;
	const num = (s, fallback) => s === void 0 ? fallback : Number(s);
	const year = num(m[1], 0);
	const month = num(m[2], 1);
	const day = num(m[3], 1);
	const hour = num(m[4], 0);
	const minute = num(m[5], 0);
	const second = num(m[6], 0);
	const ms = m[7] === void 0 ? 0 : Math.floor(Number(`0${m[7]}`) * 1e3);
	if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second > 59) return null;
	const date = /* @__PURE__ */ new Date(0);
	date.setUTCFullYear(year, month - 1, day);
	date.setUTCHours(hour, minute, second, ms);
	if (date.getUTCDate() !== day) return null;
	const zone = m[8];
	if (zone !== void 0 && zone !== "Z") {
		const sign = zone.startsWith("-") ? -1 : 1;
		const zh = Number(zone.slice(1, 3));
		const zm = Number(zone.slice(4, 6));
		if (zh > 23 || zm > 59) return null;
		date.setTime(date.getTime() - sign * (zh * 60 + zm) * 6e4);
	}
	return date;
}
/** Text of an element including the text of its children (keywords may wrap words in `value`). */
function allText(el) {
	if (el.children.length === 0) return el.text;
	return [el.text, ...el.children.map(allText)].map((s) => s.trim()).filter((s) => s !== "").join(" ");
}
/**
* Reads the Core Properties part. Unknown elements are ignored; a date that cannot be read is
* left out with `XLV_W_INVALID_VALUE`; malformed XML or another root element gives a warning and
* no properties, because the part is optional metadata.
*/
function parseCoreProperties(xml, ctx) {
	let root;
	try {
		root = parseXmlTree(xml, ctx);
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			ctx.warnings.add("XLV_W_INVALID_VALUE", "The document properties part is not well-formed XML; its properties are ignored.", { part: ctx.part });
			return {};
		}
		throw error;
	}
	if (root.ns !== "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" || root.local !== "coreProperties") {
		ctx.warnings.add("XLV_W_INVALID_VALUE", `The document properties part has the root element <${root.local}>; its properties are ignored.`, { part: ctx.part });
		return {};
	}
	const out = {};
	for (const [key, prefix] of ELEMENTS) {
		const el = root.children.find((c) => c.ns === PREFIX_NS[prefix] && c.local === key);
		if (el === void 0) continue;
		const text = allText(el);
		if (DATE_KEYS.includes(key)) {
			if (text.trim() === "") continue;
			const date = parseW3cdtf(text);
			if (date === null) {
				ctx.warnings.add("XLV_W_INVALID_VALUE", `The document property ${prefix}:${key} has the value "${text.trim().slice(0, 40)}", which is not a date; it is ignored.`, { part: ctx.part });
				continue;
			}
			out[key] = date;
		} else out[key] = text;
	}
	return out;
}
/** Root namespaces of the Extended File Properties part: Transitional (Part 4 §13.2.12), Strict (Part 1 §15.2.12.3). */
const APP_PROPERTIES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const APP_PROPERTIES_NS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";
/** Variant types namespace of the Transitional schema (shared-documentPropertiesExtended.xsd). */
const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
/** Relationship types of the Extended File Properties part: Transitional (Part 4 §13.2.12), Strict (Part 1 §15.2.12.3). */
const APP_PROPERTIES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const APP_PROPERTIES_REL_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";
/** Media type of the Extended File Properties part (Part 1 §15.2.12.3). */
const APP_PROPERTIES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
/** Part name used by the writer (the name in Part 1 §12.2's example; parts are found through relationships). */
const APP_PROPERTIES_PART = "docProps/app.xml";
/** `true` for a relationship type that points to the Extended File Properties part. */
function isAppPropertiesRel(type) {
	return type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" || type === APP_PROPERTIES_REL_STRICT;
}
/** A `vt:lpstr` with the text escaped as ST_Xstring requires for this part (Part 1 §22.2.2.21). */
function lpstr(s) {
	return `<vt:lpstr>${escapeText(encodeXstring(s))}</vt:lpstr>`;
}
/**
* The Extended File Properties part, or `null` when it is not needed. It is written when one of
* `company`, `manager`, `application`, `appVersion` is set, or when the workbook has defined
* names that are not hidden: Part 1 §18.17.6.5 requires each name to be an `lpstr` of this part.
* `TitlesOfParts` then lists the sheet names followed by those names, and `HeadingPairs` counts
* them as "Worksheets" and "Named Ranges" ([MS-OI29500] §2.1.1720 b, §2.1.1722; W2-52).
*/
function appPropertiesXml(properties, sheetNames, nameTitles) {
	if (APP_ELEMENTS.filter(([key]) => properties[key] !== void 0).length === 0 && nameTitles.length === 0) return null;
	const text = (key, element) => {
		const v = properties[key];
		return v === void 0 ? "" : `<${element}>${escapeText(encodeXstring(v))}</${element}>`;
	};
	let body = text("application", "Application");
	const pairs = [["Worksheets", sheetNames.length], ["Named Ranges", nameTitles.length]].filter(([, n]) => n !== 0);
	if (pairs.length > 0) {
		body += `<HeadingPairs><vt:vector size="${String(pairs.length * 2)}" baseType="variant">`;
		for (const [heading, n] of pairs) body += `<vt:variant>${lpstr(heading)}</vt:variant><vt:variant><vt:i4>${String(n)}</vt:i4></vt:variant>`;
		body += "</vt:vector></HeadingPairs>";
		const titles = [...sheetNames, ...nameTitles];
		body += `<TitlesOfParts><vt:vector size="${String(titles.length)}" baseType="lpstr">${titles.map(lpstr).join("")}</vt:vector></TitlesOfParts>`;
	}
	body += text("manager", "Manager") + text("company", "Company") + text("appVersion", "AppVersion");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="${APP_PROPERTIES_NS}" xmlns:vt="${VT_NS}">${body}</Properties>`;
}
/**
* Reads `Company`, `Manager`, `Application` and `AppVersion` from the Extended File Properties
* part (Transitional or Strict namespace), decoding `_xHHHH_` escapes. Other elements are
* ignored; malformed XML or another root element gives a warning and no properties.
*/
function parseAppProperties(xml, ctx) {
	let root;
	try {
		root = parseXmlTree(xml, ctx);
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			ctx.warnings.add("XLV_W_INVALID_VALUE", "The extended document properties part is not well-formed XML; its properties are ignored.", { part: ctx.part });
			return {};
		}
		throw error;
	}
	if (root.ns !== "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" && root.ns !== APP_PROPERTIES_NS_STRICT || root.local !== "Properties") {
		ctx.warnings.add("XLV_W_INVALID_VALUE", `The extended document properties part has the root element <${root.local}>; its properties are ignored.`, { part: ctx.part });
		return {};
	}
	const out = {};
	for (const [key, element] of APP_ELEMENTS) {
		const el = root.children.find((c) => c.ns === root.ns && c.local === element);
		if (el !== void 0) out[key] = decodeXstring(el.text);
	}
	return out;
}
// src/hyperlinks.ts
const KEYS = [
	"url",
	"location",
	"tooltip",
	"display"
];
/** [MS-OI29500] §2.1.626 a–c (W2-43). */
const MAX_LENGTH = {
	url: 65536,
	location: 2084,
	display: 2084,
	tooltip: 255
};
function invalid$6(message, where, hint, property) {
	return new XlvError("XLV_E_INVALID_HYPERLINK", message, {
		location: where,
		hint,
		details: {
			cell: where.cell,
			...property === void 0 ? {} : { property }
		}
	});
}
/** Checks a hyperlink given to the writer; throws `XLV_E_INVALID_HYPERLINK`. */
function normalizeHyperlink(value, where) {
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw invalid$6(`The hyperlink for ${where.cell} must be an object.`, where, "Pass { url: \"https://…\" } or { location: \"Sheet1!A1\" }.");
	const out = {
		url: null,
		location: null,
		tooltip: null,
		display: null
	};
	for (const key of Object.keys(value)) {
		if (!KEYS.includes(key)) throw invalid$6(`Unknown hyperlink property "${key.slice(0, 40)}" for ${where.cell}.`, where, `Known properties: ${KEYS.join(", ")}.`, key.slice(0, 40));
		const v = value[key];
		if (v === void 0) continue;
		const k = key;
		if (typeof v !== "string") throw invalid$6(`The hyperlink property "${k}" for ${where.cell} must be a string.`, where, `For example ${k}: "text".`, k);
		if (v.length > MAX_LENGTH[k]) throw invalid$6(`The hyperlink property "${k}" for ${where.cell} has ${String(v.length)} characters; at most ${String(MAX_LENGTH[k])} are allowed.`, where, `Shorten "${k}" to ${String(MAX_LENGTH[k])} characters or fewer.`, k);
		if (firstLoneSurrogate(v) !== -1) throw invalid$6(`The hyperlink property "${k}" for ${where.cell} contains an unpaired surrogate.`, where, "Remove the broken character.", k);
		if (k === "url" && (v === "" || /[\x00-\x1F\x7F\uFFFE\uFFFF]/.test(v))) throw invalid$6(`The hyperlink URL for ${where.cell} is empty or contains control characters or non-characters.`, where, "Pass the address as text, for example url: \"https://example.com\"; percent-encode unusual characters.", k);
		out[k] = v;
	}
	if (out["url"] === null && (out["location"] === null || out["location"] === "")) throw invalid$6(`The hyperlink for ${where.cell} has neither a url nor a location.`, where, "Pass { url: \"https://…\" } for a web page or file, or { location: \"Sheet1!A1\" } for a place in the workbook.");
	return out;
}
/** Schemes that are ordinary for links in a spreadsheet; others are reported (D-A08). */
const USUAL_SCHEME = /^(?:https?|mailto):/i;
/** Whether an external target uses a scheme other than http, https or mailto. */
function isUnusualTarget(target) {
	return /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target) ? !USUAL_SCHEME.test(target) : true;
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
function hexValue$1(c) {
	if (c >= 48 && c <= 57) return c - 48;
	if (c >= 65 && c <= 70) return c - 55;
	if (c >= 97 && c <= 102) return c - 87;
	return -1;
}
/** The octet encoded by `%HH` at `at`, or -1 if `at` does not start a valid percent-encoding. */
function octetAt(s, at) {
	if (s.charCodeAt(at) !== PERCENT) return -1;
	const hi = hexValue$1(s.charCodeAt(at + 1));
	const lo = hexValue$1(s.charCodeAt(at + 2));
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
// src/sml/cell-ref.ts
/** Number of rows in a worksheet: rows 1 to 1 048 576 (ECMA-376 Part 1 §18.17.5.1). */
const MAX_ROWS = 1048576;
/** Number of columns in a worksheet: columns A to XFD (ECMA-376 Part 1 §18.17.5.1). */
const MAX_COLUMNS = 16384;
/** "XFD" has three letters and "1048576" seven digits, so longer inputs are never valid. */
const MAX_LETTERS = 3;
const MAX_DIGITS = 7;
const MAX_CELL_REF_LENGTH = 10;
const CODE_A = 65;
const CODE_Z$1 = 90;
const CODE_LOWER_A = 97;
const CODE_LOWER_Z = 122;
const CODE_0$1 = 48;
const CODE_9$1 = 57;
const CODE_COLON$1 = 58;
/** Value 1–26 of an ASCII letter (either case), or 0 for any other UTF-16 code unit. */
function letterValue(code) {
	if (code >= CODE_A && code <= CODE_Z$1) return code - CODE_A + 1;
	if (code >= CODE_LOWER_A && code <= CODE_LOWER_Z) return code - CODE_LOWER_A + 1;
	return 0;
}
function isRow(row) {
	return Number.isInteger(row) && row >= 1 && row <= 1048576;
}
function isColumn(col) {
	return Number.isInteger(col) && col >= 1 && col <= 16384;
}
/**
* Returns the column letters for a 1-based column number: 1 → "A", 27 → "AA", 16384 → "XFD".
* Throws `RangeError` outside 1..MAX_COLUMNS.
*/
function columnLetters(col) {
	if (!isColumn(col)) throw new RangeError(`Column number must be a whole number from 1 to ${MAX_COLUMNS}, got ${col}.`);
	let letters = "";
	let n = col;
	while (n > 0) {
		const rem = (n - 1) % 26;
		letters = String.fromCharCode(CODE_A + rem) + letters;
		n = (n - 1 - rem) / 26;
	}
	return letters;
}
/**
* Parses the letters at `ref[from..to)` as a column. Returns the 1-based column number, or 0 if
* the slice is empty, longer than three letters, contains a non-letter or is beyond XFD.
*/
function scanColumn(ref, from, to) {
	const length = to - from;
	if (length < 1 || length > MAX_LETTERS) return 0;
	let col = 0;
	for (let i = from; i < to; i++) {
		const value = letterValue(ref.charCodeAt(i));
		if (value === 0) return 0;
		col = col * 26 + value;
	}
	return col <= 16384 ? col : 0;
}
/**
* Returns the 1-based column number for column letters ("A" → 1, "xfd" → 16384; case does not
* matter), or `null` if `letters` is not 1–3 ASCII letters or names a column beyond XFD.
*/
function columnNumber(letters) {
	const col = scanColumn(letters, 0, letters.length);
	return col === 0 ? null : col;
}
/** Parses `ref[from..to)` as a single cell reference; returns `null` if it is not one. */
function scanCellRef(ref, from, to) {
	const length = to - from;
	if (length < 2 || length > MAX_CELL_REF_LENGTH) return null;
	let i = from;
	while (i < to && letterValue(ref.charCodeAt(i)) !== 0) i++;
	const col = scanColumn(ref, from, i);
	if (col === 0) return null;
	const digitsStart = i;
	if (to - digitsStart < 1 || to - digitsStart > MAX_DIGITS) return null;
	let row = 0;
	for (; i < to; i++) {
		const code = ref.charCodeAt(i);
		if (code < CODE_0$1 || code > CODE_9$1) return null;
		row = row * 10 + (code - CODE_0$1);
	}
	if (row < 1 || row > 1048576) return null;
	return {
		row,
		col
	};
}
/**
* Parses an A1-style cell reference such as `"B7"` into `{ row: 7, col: 2 }`. Lowercase letters
* are accepted; `$` markers, ranges, whitespace and references outside the grid give `null`.
*/
function parseCellRef(ref) {
	return scanCellRef(ref, 0, ref.length);
}
/** Formats a 1-based row and column as an A1 reference. Throws `RangeError` outside the grid. */
function formatCellRef(row, col) {
	if (!isRow(row)) throw new RangeError(`Row number must be a whole number from 1 to ${MAX_ROWS}, got ${row}.`);
	return columnLetters(col) + String(row);
}
/**
* Parses a range such as `"A1:C3"`, or a single cell such as `"B2"` (a one-cell range). The
* corners are normalised so that `start` is the top-left and `end` the bottom-right cell.
* Whole-row or whole-column references, `$` markers and malformed input give `null`.
*/
function parseRange(ref) {
	const length = ref.length;
	if (length > 21) return null;
	let colon = -1;
	for (let i = 0; i < length; i++) if (ref.charCodeAt(i) === CODE_COLON$1) {
		if (colon !== -1) return null;
		colon = i;
	}
	if (colon === -1) {
		const cell = scanCellRef(ref, 0, length);
		return cell === null ? null : {
			start: cell,
			end: cell
		};
	}
	const a = scanCellRef(ref, 0, colon);
	if (a === null) return null;
	const b = scanCellRef(ref, colon + 1, length);
	if (b === null) return null;
	return {
		start: {
			row: Math.min(a.row, b.row),
			col: Math.min(a.col, b.col)
		},
		end: {
			row: Math.max(a.row, b.row),
			col: Math.max(a.col, b.col)
		}
	};
}
/**
* Formats a range as `"A1:C3"`, or as a single reference such as `"B2"` when both corners are the
* same cell. The corners are written as given. Throws `RangeError` for corners outside the grid.
*/
function formatRange(range) {
	const start = formatCellRef(range.start.row, range.start.col);
	if (range.start.row === range.end.row && range.start.col === range.end.col) return start;
	return `${start}:${formatCellRef(range.end.row, range.end.col)}`;
}
// src/sml/formula-shift.ts
/** Number of decimal digits of a positive whole number. */
const digits = (n) => n < 10 ? 1 : n < 100 ? 2 : n < 1e3 ? 3 : n < 1e4 ? 4 : n < 1e5 ? 5 : n < 1e6 ? 6 : 7;
/** Number of letters of a column number from 1 to 16384 (A to XFD). */
const letters = (c) => c <= 26 ? 1 : c <= 702 ? 2 : 3;
/** Error constants (§18.17.3, fact E-33). */
const ERROR_CONSTANTS = [
	"#NULL!",
	"#DIV/0!",
	"#VALUE!",
	"#REF!",
	"#NAME?",
	"#NUM!",
	"#N/A",
	"#GETTING_DATA"
];
/**
* Characters copied as operators or separators (F2-07). `@` is not among them: it is in no
* documented grammar. CR and LF are white space ([MS-XLSX] §2.2.2 `whitespace`; an XML parser
* turns the CR LF pair into LF).
*/
const OPERATORS = "+-*/^&=<>%,;(){}: 	\r\n";
function isLetter(c) {
	return c >= 65 && c <= 90 || c >= 97 && c <= 122;
}
function isDigit(c) {
	return c >= 48 && c <= 57;
}
/** Characters allowed inside a name or reference token (F2-05). */
function isIdentChar(c) {
	return isLetter(c) || isDigit(c) || c === 95 || c === 46 || c === 92 || c === 63 || c === 36 || c >= 128;
}
function isIdentStart(c) {
	return isLetter(c) || c === 95 || c === 92 || c === 36 || c >= 128;
}
/** Column number of 1–3 ASCII letters at s[from..to), or 0. */
function columnOf(s, from, to) {
	if (to - from < 1 || to - from > 3) return 0;
	let col = 0;
	for (let i = from; i < to; i++) {
		const c = s.charCodeAt(i);
		if (!isLetter(c)) return 0;
		col = col * 26 + ((c | 32) - 96);
	}
	return col <= 16384 ? col : 0;
}
/** Row number of the decimal digits at s[from..to) (1 … MAX_ROWS), or 0. */
function rowOf(s, from, to) {
	if (to - from < 1 || to - from > 7) return 0;
	let row = 0;
	for (let i = from; i < to; i++) {
		const c = s.charCodeAt(i);
		if (!isDigit(c)) return 0;
		row = row * 10 + (c - 48);
	}
	return row >= 1 && row <= 1048576 ? row : 0;
}
/** Parses `$?letters$?digits` (the whole of `id`) as a cell reference. */
function cellRef(id) {
	let i = 0;
	const colAbs = id.charCodeAt(i) === 36;
	if (colAbs) i++;
	const colStart = i;
	while (i < id.length && isLetter(id.charCodeAt(i))) i++;
	const col = columnOf(id, colStart, i);
	if (col === 0) return null;
	const rowAbs = id.charCodeAt(i) === 36;
	if (rowAbs) i++;
	const rowStart = i;
	while (i < id.length && isDigit(id.charCodeAt(i))) i++;
	if (i !== id.length) return null;
	const row = rowOf(id, rowStart, i);
	return row === 0 ? null : {
		kind: "cell",
		col,
		colAbs,
		row,
		rowAbs
	};
}
/** Whether `id` has the shape `$?letters$?digits` with 1–3 letters but lies outside the grid. */
function beyondGrid(id) {
	let i = id.charCodeAt(0) === 36 ? 1 : 0;
	const colStart = i;
	while (i < id.length && isLetter(id.charCodeAt(i))) i++;
	if (i - colStart < 1 || i - colStart > 3) return false;
	if (id.charCodeAt(i) === 36) i++;
	const rowStart = i;
	while (i < id.length && isDigit(id.charCodeAt(i))) i++;
	return i > rowStart && i === id.length && cellRef(id) === null;
}
/** Parses `$?letters` (the whole of `id`) as a column; returns 0 if it is not one. */
function columnRef(id) {
	const abs = id.charCodeAt(0) === 36;
	const col = columnOf(id, abs ? 1 : 0, id.length);
	return col === 0 ? null : {
		col,
		abs
	};
}
/** Names that look like R1C1 references (`R`, `C`, `R1`, `C2`, `RC`, `R1C1`): ambiguous, F2-09. */
function looksR1C1(id) {
	let i = 0;
	const upper = (k) => id.charCodeAt(k) & -33;
	if (upper(i) === 82) {
		i++;
		while (i < id.length && isDigit(id.charCodeAt(i))) i++;
		if (i === id.length) return true;
	}
	if (upper(i) !== 67) return false;
	i++;
	while (i < id.length && isDigit(id.charCodeAt(i))) i++;
	return i === id.length;
}
/**
* Splits a formula into literal text and references. Returns `null` when the text contains
* anything that cannot be classified with confidence.
*/
function tokenize(s) {
	const parts = [];
	let literal = "";
	let lastName = null;
	const flush = () => {
		if (literal !== "") {
			parts.push(literal);
			literal = "";
		}
		lastName = null;
	};
	const n = s.length;
	let i = 0;
	while (i < n) {
		const c = s.charCodeAt(i);
		if (c === 34) {
			let j = i + 1;
			for (;;) {
				if (j >= n) return null;
				if (s.charCodeAt(j) === 34) {
					if (s.charCodeAt(j + 1) === 34) j += 2;
					else break;
				} else j++;
			}
			literal += s.slice(i, j + 1);
			i = j + 1;
			continue;
		}
		if (c === 39) {
			let j = i + 1;
			for (;;) {
				if (j >= n) return null;
				if (s.charCodeAt(j) === 39) {
					if (s.charCodeAt(j + 1) === 39) j += 2;
					else break;
				} else j++;
			}
			if (s.charCodeAt(j + 1) !== 33) return null;
			const inner = s.slice(i + 1, j).replace(/''/g, "'");
			const external = inner.startsWith("[") || literal.endsWith("]");
			flush();
			parts.push({
				kind: "prefix",
				text: s.slice(i, j + 2),
				sheets: inner.replace(/^\[[^\]]*\]/, "").split(":"),
				external
			});
			i = j + 2;
			continue;
		}
		if (c === 91) {
			let depth = 0;
			let j = i;
			while (j < n) {
				const d = s.charCodeAt(j);
				if (d === 39) j++;
				else if (d === 91) depth++;
				else if (d === 93 && --depth === 0) break;
				j++;
			}
			if (j >= n) return null;
			literal += s.slice(i, j + 1);
			i = j + 1;
			continue;
		}
		if (c === 93) return null;
		if (c === 33) {
			literal += "!";
			i++;
			continue;
		}
		if (c === 35) {
			const constant = ERROR_CONSTANTS.find((e) => s.startsWith(e, i));
			if (constant === void 0) return null;
			literal += constant;
			i += constant.length;
			continue;
		}
		if (isDigit(c) || c === 46 || c === 36 && isDigit(s.charCodeAt(i + 1))) {
			const rows = rowRange(s, i);
			if (rows !== null) {
				flush();
				parts.push(rows.part);
				i = rows.end;
				continue;
			}
			if (c === 36) return null;
			let j = i;
			while (j < n && isDigit(s.charCodeAt(j))) j++;
			if (s.charCodeAt(j) === 46) {
				j++;
				while (j < n && isDigit(s.charCodeAt(j))) j++;
			}
			if (j === i + 1 && c === 46) return null;
			const e = s.charCodeAt(j);
			if (e === 69 || e === 101) {
				j++;
				if (s.charCodeAt(j) === 43 || s.charCodeAt(j) === 45) j++;
				const digits = j;
				while (j < n && isDigit(s.charCodeAt(j))) j++;
				if (j === digits) return null;
			}
			if (j < n && (isIdentChar(s.charCodeAt(j)) || s.charCodeAt(j) === 58)) return null;
			literal += s.slice(i, j);
			i = j;
			continue;
		}
		if (isIdentStart(c)) {
			let j = i;
			while (j < n && isIdentChar(s.charCodeAt(j))) j++;
			const id = s.slice(i, j);
			const next = s.charCodeAt(j);
			let k = j;
			while (k < n && (s.charCodeAt(k) === 32 || s.charCodeAt(k) === 10 || s.charCodeAt(k) === 13)) k++;
			const cell = cellRef(id);
			const column = cell === null ? columnRef(id) : null;
			const refLike = cell !== null || looksR1C1(id);
			const dot = id.indexOf(".");
			if (dot > 0 && cellRef(id.slice(0, dot)) !== null) return null;
			if (next === 40) {
				literal += id;
				i = j;
				continue;
			}
			if (s.charCodeAt(k) === 40 && refLike) return null;
			if (next === 33 || next === 91) {
				if (refLike) return null;
				if (next === 91) {
					literal += id;
					i = j;
					continue;
				}
				const external = literal.endsWith("]");
				let text = id + "!";
				let sheets = [id];
				if (lastName !== null && literal.length === lastName.end + 1 && literal.endsWith(":")) {
					literal = literal.slice(0, lastName.end - lastName.id.length);
					text = `${lastName.id}:${text}`;
					sheets = [lastName.id, id];
				}
				flush();
				parts.push({
					kind: "prefix",
					text,
					sheets,
					external
				});
				i = j + 1;
				continue;
			}
			if (cell === null && beyondGrid(id)) return null;
			if (cell !== null) {
				flush();
				parts.push(cell);
				i = j;
				continue;
			}
			if (column !== null && next === 58) {
				let m = j + 1;
				while (m < n && isIdentChar(s.charCodeAt(m))) m++;
				const second = columnRef(s.slice(j + 1, m));
				if (second !== null) {
					if (s.charCodeAt(m) === 33) return null;
					flush();
					parts.push({
						kind: "cols",
						c1: column.col,
						c1Abs: column.abs,
						c2: second.col,
						c2Abs: second.abs
					});
					i = m;
					continue;
				}
			}
			if (id.includes("$") || looksR1C1(id)) return null;
			literal += id;
			lastName = {
				id,
				end: literal.length
			};
			i = j;
			continue;
		}
		if (OPERATORS.includes(s[i] ?? "")) {
			literal += s[i];
			i++;
			continue;
		}
		return null;
	}
	flush();
	return parts;
}
/** Parses `$?digits:$?digits` at `i` as a whole-row range, or returns `null`. */
function rowRange(s, i) {
	const one = (at) => {
		let j = at;
		const abs = s.charCodeAt(j) === 36;
		if (abs) j++;
		const start = j;
		while (j < s.length && isDigit(s.charCodeAt(j))) j++;
		const row = rowOf(s, start, j);
		return row === 0 ? null : {
			row,
			abs,
			end: j
		};
	};
	const a = one(i);
	if (a === null || s.charCodeAt(a.end) !== 58) return null;
	const b = one(a.end + 1);
	if (b === null || b.end < s.length && isIdentChar(s.charCodeAt(b.end))) return null;
	return {
		part: {
			kind: "rows",
			r1: a.row,
			r1Abs: a.abs,
			r2: b.row,
			r2Abs: b.abs
		},
		end: b.end
	};
}
var Compiled = class {
	parts;
	minRow = Infinity;
	maxRow = -Infinity;
	minCol = Infinity;
	maxCol = -Infinity;
	/** Length of the literal text and sheet prefixes, which do not change when moved. */
	literalLength = 0;
	/** The references, for {@link renderedLength}. */
	refs;
	constructor(parts) {
		this.parts = parts;
		const refs = [];
		for (const p of parts) if (typeof p === "string") this.literalLength += p.length;
		else if (p.kind === "prefix") this.literalLength += p.text.length;
		else refs.push(p);
		this.refs = refs;
		const row = (r, abs) => {
			if (abs) return;
			if (r < this.minRow) this.minRow = r;
			if (r > this.maxRow) this.maxRow = r;
		};
		const col = (c, abs) => {
			if (abs) return;
			if (c < this.minCol) this.minCol = c;
			if (c > this.maxCol) this.maxCol = c;
		};
		for (const p of parts) {
			if (typeof p === "string" || p.kind === "prefix") continue;
			if (p.kind === "cell") {
				row(p.row, p.rowAbs);
				col(p.col, p.colAbs);
			} else if (p.kind === "cols") {
				col(p.c1, p.c1Abs);
				col(p.c2, p.c2Abs);
			} else {
				row(p.r1, p.r1Abs);
				row(p.r2, p.r2Abs);
			}
		}
	}
	canShift(dRow, dCol) {
		if (!Number.isInteger(dRow) || !Number.isInteger(dCol)) return false;
		if (this.minRow !== Infinity && (this.minRow + dRow < 1 || this.maxRow + dRow > 1048576)) return false;
		if (this.minCol !== Infinity && (this.minCol + dCol < 1 || this.maxCol + dCol > 16384)) return false;
		return true;
	}
	renderedLength(dRow, dCol) {
		let length = this.literalLength;
		const col = (c, abs) => abs ? 1 + letters(c) : letters(c + dCol);
		const row = (r, abs) => abs ? 1 + digits(r) : digits(r + dRow);
		for (const p of this.refs) if (p.kind === "cell") length += col(p.col, p.colAbs) + row(p.row, p.rowAbs);
		else if (p.kind === "cols") length += col(p.c1, p.c1Abs) + 1 + col(p.c2, p.c2Abs);
		else length += row(p.r1, p.r1Abs) + 1 + row(p.r2, p.r2Abs);
		return length;
	}
	render(dRow, dCol) {
		const out = [];
		const col = (c, abs) => (abs ? "$" : "") + columnLetters(abs ? c : c + dCol);
		const row = (r, abs) => (abs ? "$" : "") + String(abs ? r : r + dRow);
		for (const p of this.parts) if (typeof p === "string") out.push(p);
		else if (p.kind === "prefix") out.push(p.text);
		else if (p.kind === "cell") out.push(col(p.col, p.colAbs) + row(p.row, p.rowAbs));
		else if (p.kind === "cols") out.push(`${col(p.c1, p.c1Abs)}:${col(p.c2, p.c2Abs)}`);
		else out.push(`${row(p.r1, p.r1Abs)}:${row(p.r2, p.r2Abs)}`);
		return out.join("");
	}
};
/**
* Splits a formula (without the leading "=") into text and references once, so that it can be
* moved to many cells cheaply. Returns `null` if the formula cannot be classified with
* confidence; such a formula must not be moved.
*/
function compileFormula(text) {
	const parts = tokenize(text);
	return parts === null ? null : new Compiled(parts);
}
/**
* Moves a formula by `dRow` rows and `dCol` columns: relative parts of A1 references change,
* `$`-marked parts, strings, names, functions, numbers and structured references do not.
* Returns `null` if the text cannot be classified with confidence or a reference would leave
* A1:XFD1048576. This is how a formula changes when it is copied to another cell; to adjust
* formulas for inserted or deleted rows or columns use `adjustFormula`.
*
* @example
* ```ts
* import { shiftFormula } from 'xlvane';
*
* console.log(shiftFormula('SUM(A1:A3)*$B$1', 2, 1)); // "SUM(B3:B5)*$B$1"
* ```
*/
function shiftFormula(text, dRow, dCol) {
	const compiled = compileFormula(text);
	return compiled !== null && compiled.canShift(dRow, dCol) ? compiled.render(dRow, dCol) : null;
}
// src/charts.ts
/** Content type of a chart part (CH-01). */
const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
/** Chart namespace (Transitional) and the `graphicData` uri of a chart frame (CH-01, CH-04). */
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS$1 = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS$1 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAX_CHART_TEXT = 255;
/** Most values of a pie series: each slice gets a `dPt` with its colour (CH-21). */
const MAX_PIE_POINTS = 1e3;
/**
* Most cells the caches of all charts of one workbook written by `writeWorkbook` cover (a limit
* of xlvane). A cached point takes about 42 bytes of XML besides its text (a 200 000-point
* scatter chart is a 16.8 MB part), so the chart parts stay below about 170 MB plus the cached
* text ({@link MAX_CHART_CACHE_TEXT}), within the reader's default `maxEntryBytes`. The streaming
* writer writes no caches (CH-20) and has no such limit.
*/
const MAX_CHART_CACHE_CELLS = 4e6;
/**
* Most characters of cached text (the `v` of string caches, and number format codes, counted as
* written, after escaping) of all charts of one workbook written by `writeWorkbook` (a limit of
* xlvane): 8 Mi characters, at most 24 MiB of UTF-8. With {@link MAX_CHART_CACHE_CELLS} a chart
* part stays below about 200 MB, under the reader's default `maxEntryBytes` (256 MiB), and its
* text far below the longest string of V8 (536 870 888 characters in Node.js 22,
* `buffer.constants.MAX_STRING_LENGTH`). Without it, 20 000 cells of 32 767 characters made a
* string too long for V8.
*/
const MAX_CHART_CACHE_TEXT = 8388608;
/** Longest reference text accepted, far above any sheet name (31) plus two cells. */
const MAX_REF_LENGTH = 512;
/**
* Default series colours: accent1 … accent6 of the theme part of `chartTitle_withTitle.xlsx`
* (Microsoft Excel 15.0300) in the POI corpus (CH-35). Series and pie slices take them in order
* and start again after the sixth.
*/
const CHART_PALETTE = [
	"5B9BD5",
	"ED7D31",
	"A5A5A5",
	"FFC000",
	"4472C4",
	"70AD47"
];
const TYPES = /* @__PURE__ */ new Set([
	"column",
	"bar",
	"line",
	"pie",
	"scatter",
	"area"
]);
const LEGENDS = /* @__PURE__ */ new Set([
	"right",
	"bottom",
	"top",
	"left"
]);
const CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/;
/** Titles and literal series names are one line: every C0 control is refused (paragraphs are R6-11). */
const LINE_CONTROL = /[\u0000-\u001f\ufffe\uffff]/;
function chartError(message, hint, sheet, option) {
	return new XlvError("XLV_E_INVALID_CHART", message, {
		location: { sheet },
		hint,
		details: { option }
	});
}
function keysOf(value, allowed, what, example, sheet) {
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw chartError(`The ${what} must be an object.`, `For example ${example}.`, sheet, what);
	for (const key of Object.keys(value)) if (!allowed.includes(key)) throw chartError(`Unknown ${what} option "${key.slice(0, 40)}".`, `Use ${allowed.join(", ")}.`, sheet, key.slice(0, 40));
	return value;
}
/**
* Parses a sheet-qualified reference such as `'Sales'!B2:B13` or `Sales!$B$1` (an optional
* leading `=` is ignored) with the formula tokenizer (F2-11); `null` if the text is not one. An
* unquoted sheet name of ASCII letters, digits and `_` that the tokenizer reads as a cell or R1C1
* reference, such as `S1!$A$1` as LibreOffice writes it back, gives a reference marked
* `unquoted`, which `resolveChartSheets` accepts only for a sheet of exactly that name.
*/
function parseChartRef(text) {
	if (text.length > MAX_REF_LENGTH) return null;
	const body = text.startsWith("=") ? text.slice(1) : text;
	const ref = refOfTokens(tokenize(body));
	if (ref !== null) return ref;
	const bare = /^([A-Za-z_][A-Za-z0-9_]*)!/.exec(body);
	if (bare === null) return null;
	const quoted = refOfTokens(tokenize(`'${bare[1]}'${body.slice(bare[1].length)}`));
	return quoted === null ? null : {
		...quoted,
		unquoted: true
	};
}
function refOfTokens(parts) {
	if (parts === null || parts.length !== 2 && parts.length !== 4) return null;
	const [prefix, a, colon, b] = parts;
	if (typeof prefix !== "object" || prefix.kind !== "prefix" || prefix.external || prefix.sheets.length !== 1) return null;
	if (typeof a !== "object" || a.kind !== "cell") return null;
	if (parts.length === 4 && (colon !== ":" || typeof b !== "object" || b.kind !== "cell")) return null;
	const end = b ?? a;
	return {
		sheet: prefix.sheets[0],
		range: {
			start: {
				row: Math.min(a.row, end.row),
				col: Math.min(a.col, end.col)
			},
			end: {
				row: Math.max(a.row, end.row),
				col: Math.max(a.col, end.col)
			}
		}
	};
}
function checkText(value, what, max, sheet) {
	if (typeof value !== "string" || value === "" || value.length > max || LINE_CONTROL.test(value) || firstLoneSurrogate(value) !== -1) throw chartError(`The chart ${what} must be a text of 1 to ${String(max)} characters on one line without control characters.`, `For example ${what}: "Sales".`, sheet, what);
	return value;
}
function checkRange(value, what, sheet) {
	const ref = typeof value === "string" ? parseChartRef(value.trim()) : null;
	if (ref === null) throw chartError(`The chart series ${what} ${typeof value === "string" ? `"${value.slice(0, 60)}"` : ""} is not a sheet-qualified range.`, `Give the sheet and the cells, for example ${what}: "'Sales'!B2:B13".`, sheet, what);
	if (ref.range.start.row !== ref.range.end.row && ref.range.start.col !== ref.range.end.col) throw chartError(`The chart series ${what} must be one row or one column, not a block of cells.`, `For example "'Sales'!B2:B13" (a column) or "'Sales'!B2:M2" (a row).`, sheet, what);
	return ref;
}
function cellCount(r) {
	return (r.range.end.row - r.range.start.row + 1) * (r.range.end.col - r.range.start.col + 1);
}
function normalizeSeries(raw, t, sheet) {
	const s = keysOf(raw, [
		"values",
		"categories",
		"name",
		"color"
	], "series", "{ values: \"Sales!B2:B13\" }", sheet);
	const values = checkRange(s["values"], "values", sheet);
	const categories = s["categories"] === void 0 ? null : checkRange(s["categories"], "categories", sheet);
	let name = null;
	if (s["name"] !== void 0) {
		const ref = typeof s["name"] === "string" ? parseChartRef(s["name"].trim()) : null;
		if (ref !== null && cellCount(ref) !== 1) throw chartError("A series name reference must be one cell.", "For example name: \"Sales!B1\", or a literal name such as name: \"Sales\".", sheet, "name");
		name = ref === null ? { text: checkText(s["name"], "name", MAX_CHART_TEXT, sheet) } : { ref };
	}
	let color = null;
	if (s["color"] !== void 0) {
		if (t === "pie") throw chartError("The slices of a pie chart take their colours from the palette; a pie series has no color.", "Leave color out of the pie series.", sheet, "color");
		const c = s["color"];
		if (typeof c !== "string" || !/^#?[0-9A-Fa-f]{6}$/.test(c)) throw chartError("A series color must be six hex digits RRGGBB.", "For example color: \"2E75B6\".", sheet, "color");
		color = c.replace("#", "").toUpperCase();
	}
	if (t === "pie" && cellCount(values) > MAX_PIE_POINTS) throw chartError(`A pie chart has at most ${String(MAX_PIE_POINTS)} values.`, "Use fewer cells, or a column chart for many values.", sheet, "values");
	return {
		name,
		values,
		categories,
		color
	};
}
/**
* Validates the options of `NewSheet.addChart` except the anchor (checked by the caller) and
* the sheets of the references (checked when the workbook is written). Throws
* `XLV_E_INVALID_CHART`.
*/
function normalizeChart(input, sheet) {
	const o = keysOf(input, [
		"type",
		"anchor",
		"series",
		"title",
		"legend",
		"stacked",
		"axes",
		"name",
		"description"
	], "chart", "{ type: \"column\", anchor: \"E2:L18\", series: [{ values: \"Sales!B2:B13\" }] }", sheet);
	const type = o["type"];
	if (typeof type !== "string" || !TYPES.has(type)) throw chartError(`The chart type ${JSON.stringify(String(type).slice(0, 20))} is not supported.`, "Use \"column\", \"bar\", \"line\", \"pie\", \"scatter\" or \"area\".", sheet, "type");
	const t = type;
	const list = o["series"];
	if (!Array.isArray(list) || list.length === 0 || list.length > 255) throw chartError(`A chart needs 1 to ${String(255)} series.`, "For example series: [{ name: \"Sales!B1\", categories: \"Sales!A2:A13\", values: \"Sales!B2:B13\" }].", sheet, "series");
	if (t === "pie" && list.length !== 1) throw chartError("A pie chart has exactly one series.", "Give one series, or use a column or bar chart to compare several.", sheet, "series");
	const series = [];
	for (let i = 0; i < list.length; i++) {
		if (!Object.hasOwn(list, i)) throw chartError(`Series ${String(i + 1)} of the chart is missing (the series array has a hole).`, "Give an object such as { values: \"Sales!B2:B13\" } for every series.", sheet, "series");
		series.push(normalizeSeries(list[i], t, sheet));
	}
	const title = o["title"] === void 0 ? null : checkText(o["title"], "title", MAX_CHART_TEXT, sheet);
	const legend = o["legend"] === void 0 ? "right" : o["legend"];
	if (legend !== false && (typeof legend !== "string" || !LEGENDS.has(legend))) throw chartError("The chart legend must be \"right\", \"bottom\", \"top\", \"left\" or false.", "For example legend: \"bottom\", or legend: false for no legend.", sheet, "legend");
	const stacked = o["stacked"] === void 0 ? false : o["stacked"];
	if (stacked !== false && stacked !== true && stacked !== "percent") throw chartError("stacked must be true, false or \"percent\".", "For example stacked: true.", sheet, "stacked");
	if (stacked !== false && (t === "pie" || t === "scatter")) throw chartError(`A ${t} chart cannot be stacked.`, "Use stacked with column, bar, line or area charts.", sheet, "stacked");
	let xTitle = null;
	let yTitle = null;
	let min = null;
	let max = null;
	if (o["axes"] !== void 0) {
		if (t === "pie") throw chartError("A pie chart has no axes.", "Leave axes out of a pie chart.", sheet, "axes");
		const axes = keysOf(o["axes"], ["x", "y"], "axes", "{ y: { title: \"EUR\", min: 0 } }", sheet);
		if (axes["x"] !== void 0) {
			const x = keysOf(axes["x"], ["title"], "x axis", "{ title: \"Month\" }", sheet);
			if (x["title"] !== void 0) xTitle = checkText(x["title"], "axis title", MAX_CHART_TEXT, sheet);
		}
		if (axes["y"] !== void 0) {
			const y = keysOf(axes["y"], [
				"title",
				"min",
				"max"
			], "y axis", "{ title: \"EUR\", min: 0, max: 100 }", sheet);
			if (y["title"] !== void 0) yTitle = checkText(y["title"], "axis title", MAX_CHART_TEXT, sheet);
			for (const k of ["min", "max"]) {
				const v = y[k];
				if (v === void 0) continue;
				if (typeof v !== "number" || !Number.isFinite(v)) throw chartError(`The y axis ${k} must be a finite number.`, `For example ${k}: 0.`, sheet, k);
				if (k === "min") min = v;
				else max = v;
			}
			if (min !== null && max !== null && min >= max) throw chartError("The y axis min must be less than its max.", "For example { min: 0, max: 100 }.", sheet, "min");
		}
	}
	const objectText = (key) => {
		const v = o[key];
		if (v === void 0) return null;
		if (typeof v !== "string" || v.length > 2084 || CONTROL.test(v) || firstLoneSurrogate(v) !== -1 || key === "name" && v === "") throw chartError(`The chart ${key} must be a text of ${key === "name" ? "1" : "0"} to 2084 characters without control characters.`, `For example ${key}: "Sales chart".`, sheet, key);
		return v;
	};
	return {
		type: t,
		title,
		legend: legend === false ? null : legend,
		stacked,
		series,
		xTitle,
		yTitle,
		min,
		max,
		name: objectText("name"),
		description: objectText("description")
	};
}
/** Builds the lookup of {@link resolveChartSheets} for the sheets of a workbook, in order. */
function chartSheetNames(names) {
	return {
		names,
		exact: new Set(names),
		byKey: new Map(names.map((s) => [s.toLowerCase(), s]))
	};
}
/**
* Replaces the sheet of every reference by the workbook's sheet of that name (names compare
* without regard to case; an `unquoted` reference must name a sheet exactly). `lookup` is built
* once per workbook ({@link chartSheetNames}), so resolving the charts of a workbook takes time
* linear in its sheets and charts. Throws `XLV_E_INVALID_CHART` for a sheet the workbook does
* not have.
*/
function resolveChartSheets(chart, lookup, sheet) {
	const { names: sheets, exact, byKey } = lookup;
	const fix = (r) => {
		if (r.unquoted === true) {
			if (exact.has(r.sheet)) return {
				sheet: r.sheet,
				range: r.range
			};
			const quoted = chartRefText({
				sheet: r.sheet,
				range: r.range
			});
			throw new XlvError("XLV_E_INVALID_CHART", `A chart on sheet "${sheet}" refers to sheet "${r.sheet.slice(0, 100)}" without quotes. Such a name reads as a cell or R1C1 reference, so it is accepted only as exactly the name of a sheet of the workbook, and the workbook has no sheet "${r.sheet.slice(0, 100)}".`, {
				location: { sheet },
				hint: `Quote the sheet name, for example "${quoted}", or use the exact name of one of the sheets: ${sheets.slice(0, 10).map((s) => `"${s}"`).join(", ")}${sheets.length > 10 ? ", …" : ""}.`,
				details: {
					option: "series",
					sheet: r.sheet.slice(0, 100)
				}
			});
		}
		const name = byKey.get(r.sheet.toLowerCase());
		if (name === void 0) throw new XlvError("XLV_E_INVALID_CHART", `A chart on sheet "${sheet}" refers to sheet "${r.sheet.slice(0, 100)}", which the workbook does not have.`, {
			location: { sheet },
			hint: `Use one of the sheets: ${sheets.slice(0, 10).map((s) => `"${s}"`).join(", ")}${sheets.length > 10 ? ", …" : ""}.`,
			details: {
				option: "series",
				sheet: r.sheet.slice(0, 100)
			}
		});
		return {
			sheet: name,
			range: r.range
		};
	};
	return {
		...chart,
		series: chart.series.map((s) => ({
			...s,
			values: fix(s.values),
			categories: s.categories === null ? null : fix(s.categories),
			name: s.name !== null && "ref" in s.name ? { ref: fix(s.name.ref) } : s.name
		}))
	};
}
/**
* Checks the cells the caches of `charts` (all charts of a workbook, in order) would cover against
* {@link MAX_CHART_CACHE_CELLS} before any chart part is built. Throws `XLV_E_INVALID_CHART`.
*/
function checkChartCacheCells(charts) {
	let total = 0;
	for (const { chart, sheet } of charts) {
		for (const s of chart.series) total += cellCount(s.values) + (s.categories === null ? 0 : cellCount(s.categories)) + (s.name !== null && "ref" in s.name ? 1 : 0);
		if (total > 4e6) throw new XlvError("XLV_E_INVALID_CHART", `The charts of the workbook would hold cached values of ${String(total)} cells (counted up to a chart on sheet "${sheet}"), more than the limit of ${String(MAX_CHART_CACHE_CELLS)} cells.`, {
			location: { sheet },
			hint: "Refer charts to fewer cells, or write the workbook with writeWorkbookStream, which writes charts without cached values.",
			details: {
				option: "series",
				limit: MAX_CHART_CACHE_CELLS,
				value: total
			}
		});
	}
}
/** Counts `written` characters of cached text; throws `XLV_E_INVALID_CHART` over {@link MAX_CHART_CACHE_TEXT}. */
function countCacheText(source, written) {
	source.text.used += written.length;
	if (source.text.used > 8388608) throw new XlvError("XLV_E_INVALID_CHART", `The caches of the charts of the workbook would hold more than ${String(MAX_CHART_CACHE_TEXT)} characters of text (counted up to a chart on sheet "${source.sheet}"), the limit of xlvane.`, {
		location: { sheet: source.sheet },
		hint: "Refer charts to cells with shorter text (for example the categories), or write the workbook with writeWorkbookStream, which writes charts without cached values.",
		details: {
			option: "series",
			limit: MAX_CHART_CACHE_TEXT,
			value: source.text.used
		}
	});
	return written;
}
/**
* Quotes a sheet name for a reference (F2-11): unquoted only for ASCII letters, digits and `_`
* starting with a letter or `_`, not `TRUE` or `FALSE`, not letters followed by digits (such as
* `XFE1` or `Sheet1`, which can read as a cell even beyond the grid) and not read as a cell or
* R1C1 reference by the formula tokenizer, the form Excel writes for names such as `Data`
* (CH-18); quoted otherwise, which is valid for every name (F2-11).
*/
function chartSheetPrefix(name) {
	if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^(true|false)$/i.test(name) && !/^[A-Za-z]+[0-9]+$/.test(name)) {
		const parts = tokenize(`${name}!A1`);
		const prefix = parts?.[0];
		if (parts?.length === 2 && typeof prefix === "object" && prefix.kind === "prefix" && prefix.sheets[0] === name) return `${name}!`;
	}
	return `'${name.replace(/'/g, "''")}'!`;
}
/** The `f` text of a reference: sheet prefix and absolute cells, e.g. `Sales!$B$2:$B$13` (CH-18). */
function chartRefText(r) {
	const abs = (row, col) => formatCellRef(row, col).replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");
	const a = abs(r.range.start.row, r.range.start.col);
	const single = r.range.start.row === r.range.end.row && r.range.start.col === r.range.end.col;
	return `${chartSheetPrefix(r.sheet)}${single ? a : `${a}:${abs(r.range.end.row, r.range.end.col)}`}`;
}
const val = (local, v) => `<c:${local} val="${String(v)}"/>`;
const solid = (color) => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
const numText = (n) => String(n).replace("e+", "E").replace("e-", "E-");
/** Literal rich text (CT_Tx `rich`, CH-22): one paragraph with one run. */
function richXml(text, rotate) {
	return `<c:tx><c:rich><a:bodyPr${rotate ? " rot=\"-5400000\" vert=\"horz\"" : ""}/><a:lstStyle/><a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}
/** CT_Title: `tx`, `overlay` (CH-22, CH-25); the vertical axis title is rotated as excelize and LibreOffice write it. */
function titleXml(text, rotate = false) {
	return text === null ? "" : `<c:title>${richXml(text, rotate)}${val("overlay", 0)}</c:title>`;
}
/** The cells of a reference in order (one row or one column). */
function cellsOf(r, { cells }) {
	const out = [];
	for (let row = r.range.start.row; row <= r.range.end.row; row++) for (let col = r.range.start.col; col <= r.range.end.col; col++) out.push(cells(r.sheet, row, col));
	return out;
}
/** CT_StrData: `ptCount`, then a `pt` for every cell with a value (CH-19). */
function strCache(values, source) {
	let xml = `<c:strCache>${val("ptCount", values.length)}`;
	values.forEach((v, i) => {
		if (v !== null) xml += `<c:pt idx="${String(i)}"><c:v>${countCacheText(source, escapeText(encodeXstring("s" in v ? v.s : numText(v.n))))}</c:v></c:pt>`;
	});
	return `${xml}</c:strCache>`;
}
/** CT_NumData: `formatCode` (of the first number), `ptCount`, then a `pt` for every number (CH-19). */
function numCache(values, source) {
	let xml = `<c:numCache><c:formatCode>${countCacheText(source, escapeText(encodeXstring(values.find((v) => v !== null && "n" in v)?.format ?? "General")))}</c:formatCode>${val("ptCount", values.length)}`;
	values.forEach((v, i) => {
		if (v !== null && "n" in v) xml += `<c:pt idx="${String(i)}"><c:v>${numText(v.n)}</c:v></c:pt>`;
	});
	return `${xml}</c:numCache>`;
}
/**
* A data source: `numRef` or `strRef` with the reference and, when `cells` is given, its cache.
* `kind` `'auto'` (categories) is a `numRef` when every non-empty cell holds a number and a
* `strRef` otherwise; without cells (the streaming writer) it is a `strRef`.
*/
function dataXml(local, r, kind, cells) {
	const values = cells === null ? null : cellsOf(r, cells);
	const k = kind !== "auto" ? kind : values !== null && values.some((v) => v !== null) && values.every((v) => v === null || "n" in v) ? "num" : "str";
	const cache = values === null || cells === null ? "" : k === "num" ? numCache(values, cells) : strCache(values, cells);
	return `<c:${local}><c:${k}Ref><c:f>${escapeText(chartRefText(r))}</c:f>${cache}</c:${k}Ref></c:${local}>`;
}
/** CT_SerTx: literal `v` or a `strRef` to one cell (CH-17). */
function seriesNameXml(name, cells) {
	if (name === null) return "";
	if ("text" in name) return `<c:tx><c:v>${escapeText(encodeXstring(name.text))}</c:v></c:tx>`;
	const cache = cells === null ? "" : strCache(cellsOf(name.ref, cells), cells);
	return `<c:tx><c:strRef><c:f>${escapeText(chartRefText(name.ref))}</c:f>${cache}</c:strRef></c:tx>`;
}
/** One series in the order of its type's sequence (CH-16), with an explicit colour (CH-21). */
function seriesXml(type, s, i, cells) {
	const color = s.color ?? CHART_PALETTE[i % CHART_PALETTE.length];
	let xml = `<c:ser>${val("idx", i)}${val("order", i)}${seriesNameXml(s.name, cells)}`;
	const cat = s.categories === null ? "" : dataXml(type === "scatter" ? "xVal" : "cat", s.categories, type === "scatter" ? "num" : "auto", cells);
	const values = dataXml(type === "scatter" ? "yVal" : "val", s.values, "num", cells);
	if (type === "line") xml += `<c:spPr><a:ln>${solid(color)}</a:ln></c:spPr><c:marker>${val("symbol", "none")}</c:marker>${cat}${values}${val("smooth", 0)}`;
	else if (type === "scatter") xml += `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker>${val("symbol", "circle")}${val("size", 5)}<c:spPr>${solid(color)}<a:ln>${solid(color)}</a:ln></c:spPr></c:marker>${cat}${values}${val("smooth", 0)}`;
	else if (type === "pie") {
		let points = "";
		for (let k = 0; k < cellCount(s.values); k++) points += `<c:dPt>${val("idx", k)}<c:spPr>${solid(CHART_PALETTE[k % CHART_PALETTE.length])}</c:spPr></c:dPt>`;
		xml += `${points}${cat}${values}`;
	} else xml += `<c:spPr>${solid(color)}</c:spPr>${cat}${values}`;
	return `${xml}</c:ser>`;
}
/** Axis ids of the one chart group of a part: fixed, unique within the part (CH-24, PROVISIONAL). */
const AXIS_IDS = [500000001, 500000002];
function axisXml(local, id, cross, pos, title, rotate, opts) {
	const scaling = `<c:scaling>${val("orientation", "minMax")}${opts.max == null ? "" : val("max", numText(opts.max))}${opts.min == null ? "" : val("min", numText(opts.min))}</c:scaling>`;
	let xml = `<c:${local}>${val("axId", id)}${scaling}${val("delete", 0)}${val("axPos", pos)}${opts.gridlines ? "<c:majorGridlines/>" : ""}${titleXml(title, rotate)}`;
	xml += `<c:numFmt formatCode="General" sourceLinked="1"/>${val("tickLblPos", "nextTo")}${val("crossAx", cross)}${val("crosses", "autoZero")}`;
	xml += local === "catAx" ? `${val("auto", 1)}${val("lblAlgn", "ctr")}${val("lblOffset", 100)}` : val("crossBetween", opts.crossBetween ?? "between");
	return `${xml}</c:${local}>`;
}
/**
* The chart part `xl/charts/chartN.xml` (CH-06 … CH-28): `roundedCorners`, then `chart` with the
* title, `autoTitleDeleted`, the plot area (one chart group and its axes), the legend,
* `plotVisOnly` and `dispBlanksAs`. `cells` fills the caches (CH-19) and counts their text
* against {@link MAX_CHART_CACHE_TEXT} (`XLV_E_INVALID_CHART` over it); the streaming writer
* passes `null` and writes references without caches (CH-20). No Office extensions (CH-32).
*/
function chartPartXml(chart, cells) {
	const t = chart.type;
	const series = chart.series.map((s, i) => seriesXml(t, s, i, cells)).join("");
	const grouping = chart.stacked === "percent" ? "percentStacked" : chart.stacked ? "stacked" : t === "column" || t === "bar" ? "clustered" : "standard";
	const [x, y] = AXIS_IDS;
	const axIds = val("axId", x) + val("axId", y);
	let group;
	if (t === "column" || t === "bar") group = `<c:barChart>${val("barDir", t === "bar" ? "bar" : "col")}${val("grouping", grouping)}${val("varyColors", 0)}${series}${val("gapWidth", 150)}${chart.stacked ? val("overlap", 100) : ""}${axIds}</c:barChart>`;
	else if (t === "line") group = `<c:lineChart>${val("grouping", grouping)}${val("varyColors", 0)}${series}${val("marker", 1)}${axIds}</c:lineChart>`;
	else if (t === "area") group = `<c:areaChart>${val("grouping", grouping)}${val("varyColors", 0)}${series}${axIds}</c:areaChart>`;
	else if (t === "scatter") group = `<c:scatterChart>${val("scatterStyle", "lineMarker")}${val("varyColors", 0)}${series}${axIds}</c:scatterChart>`;
	else group = `<c:pieChart>${val("varyColors", 1)}${series}${val("firstSliceAng", 0)}</c:pieChart>`;
	let axes = "";
	if (t !== "pie") {
		const between = t === "area" || t === "scatter" ? "midCat" : "between";
		axes = t === "scatter" ? axisXml("valAx", x, y, "b", chart.xTitle, false, { crossBetween: between }) + axisXml("valAx", y, x, "l", chart.yTitle, true, {
			gridlines: true,
			crossBetween: between,
			min: chart.min,
			max: chart.max
		}) : axisXml("catAx", x, y, "b", chart.xTitle, t === "bar", {}) + axisXml("valAx", y, x, "l", chart.yTitle, t !== "bar", {
			gridlines: true,
			crossBetween: between,
			min: chart.min,
			max: chart.max
		});
	}
	const legend = chart.legend === null ? "" : `<c:legend>${val("legendPos", {
		right: "r",
		bottom: "b",
		top: "t",
		left: "l"
	}[chart.legend])}${val("overlay", 0)}</c:legend>`;
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${A_NS$1}" xmlns:r="${R_NS$1}">${val("roundedCorners", 0)}<c:chart>${titleXml(chart.title)}${val("autoTitleDeleted", chart.title === null ? 1 : 0)}<c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}${val("plotVisOnly", 1)}${val("dispBlanksAs", "gap")}</c:chart></c:chartSpace>`;
}
/** Part name of chart `n` (1-based, numbered across the workbook; CH-01). */
function chartPartName(n) {
	return `xl/charts/chart${String(n)}.xml`;
}
/** Chart group elements of CT_PlotArea (CH-09). */
const GROUPS = /* @__PURE__ */ new Set([
	"areaChart",
	"area3DChart",
	"lineChart",
	"line3DChart",
	"stockChart",
	"radarChart",
	"scatterChart",
	"pieChart",
	"pie3DChart",
	"doughnutChart",
	"barChart",
	"bar3DChart",
	"ofPieChart",
	"surfaceChart",
	"surface3DChart",
	"bubbleChart"
]);
const KINDS = /* @__PURE__ */ new Map([
	["lineChart", "line"],
	["pieChart", "pie"],
	["scatterChart", "scatter"],
	["areaChart", "area"]
]);
const LEGEND_POSITIONS = /* @__PURE__ */ new Map([
	["r", "right"],
	["b", "bottom"],
	["t", "top"],
	["l", "left"],
	["tr", "topRight"]
]);
const AXES = /* @__PURE__ */ new Set([
	"catAx",
	"valAx",
	"dateAx",
	"serAx"
]);
const valOf = (el) => el?.attrs.get("val");
/** The `f` of a `numRef`, `strRef` or `multiLvlStrRef` child of a data source (CH-18), or `null`. */
function refOf(source) {
	if (source === void 0) return null;
	for (const local of [
		"numRef",
		"strRef",
		"multiLvlStrRef"
	]) {
		const ref = childNamed(source, CHART_NS, local);
		const f = ref === void 0 ? void 0 : childNamed(ref, CHART_NS, "f");
		if (f !== void 0) return f.text;
	}
	return null;
}
/**
* The text of a CT_Tx or CT_SerTx (CH-17, CH-22): `rich` text (the `a:t` of every run and field,
* paragraphs joined by a line feed, PROVISIONAL R6-11), a literal `v`, or the `f` of a `strRef`.
*/
function textOf(tx) {
	if (tx === void 0) return null;
	const v = childNamed(tx, CHART_NS, "v");
	if (v !== void 0) return { text: decodeXstring(v.text) };
	const ref = childNamed(tx, CHART_NS, "strRef");
	const f = ref === void 0 ? void 0 : childNamed(ref, CHART_NS, "f");
	if (f !== void 0) return { ref: f.text };
	const rich = childNamed(tx, CHART_NS, "rich");
	if (rich === void 0) return null;
	return { text: childrenNamed(rich, A_NS$1, "p").map((p) => p.children.map((run) => run.ns !== A_NS$1 ? "" : run.local === "br" ? "\n" : run.local === "r" || run.local === "fld" ? childNamed(run, A_NS$1, "t")?.text ?? "" : "").join("")).join("\n") };
}
function titleOf(el) {
	return el === void 0 ? null : textOf(childNamed(el, CHART_NS, "tx"));
}
function numberOf(el) {
	const t = (valOf(el) ?? "").trim();
	return /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(t) && Number.isFinite(Number(t)) ? Number(t) : null;
}
function axisOf(el) {
	if (el === void 0) return null;
	const scaling = childNamed(el, CHART_NS, "scaling");
	return Object.freeze({
		title: titleOf(childNamed(el, CHART_NS, "title")),
		min: scaling === void 0 ? null : numberOf(childNamed(scaling, CHART_NS, "min")),
		max: scaling === void 0 ? null : numberOf(childNamed(scaling, CHART_NS, "max"))
	});
}
/** What `chartOf` returns when a chart has more series and chart groups than the budget left. */
const OVER_BUDGET = Symbol("over budget");
/**
* The model of a `chartSpace` element; children are found by name, so order does not matter
* (CH-29). Each chart group and series counts one against `remaining`; over it, nothing more is
* built and {@link OVER_BUDGET} is returned. `used` receives the count.
*/
function chartOf(root, remaining, used) {
	const chart = childNamed(root, CHART_NS, "chart");
	const plotArea = chart === void 0 ? void 0 : childNamed(chart, CHART_NS, "plotArea");
	if (chart === void 0 || plotArea === void 0) return null;
	const axes = /* @__PURE__ */ new Map();
	for (const el of plotArea.children) {
		if (el.ns !== "http://schemas.openxmlformats.org/drawingml/2006/chart" || !AXES.has(el.local)) continue;
		const id = valOf(childNamed(el, CHART_NS, "axId"));
		if (id !== void 0 && !axes.has(id)) axes.set(id, el);
	}
	const types = [];
	const series = [];
	let stacked = null;
	let axisPair = null;
	for (const group of plotArea.children) {
		if (group.ns !== "http://schemas.openxmlformats.org/drawingml/2006/chart" || !GROUPS.has(group.local)) continue;
		if (++used.count > remaining) return OVER_BUDGET;
		const kind = group.local === "barChart" ? valOf(childNamed(group, "http://schemas.openxmlformats.org/drawingml/2006/chart", "barDir")) === "bar" ? "bar" : "column" : KINDS.get(group.local) ?? `other:${group.local}`;
		types.push(kind);
		if (stacked === null && (group.local === "barChart" || group.local === "lineChart" || group.local === "areaChart")) {
			const g = valOf(childNamed(group, CHART_NS, "grouping"));
			stacked = g === "stacked" ? true : g === "percentStacked" ? "percent" : false;
		}
		const scatter = group.local === "scatterChart" || group.local === "bubbleChart";
		for (const ser of childrenNamed(group, CHART_NS, "ser")) {
			if (++used.count > remaining) return OVER_BUDGET;
			series.push(Object.freeze({
				type: kind,
				name: textOf(childNamed(ser, CHART_NS, "tx")),
				categories: refOf(childNamed(ser, CHART_NS, scatter ? "xVal" : "cat")),
				values: refOf(childNamed(ser, CHART_NS, scatter ? "yVal" : "val"))
			}));
		}
		const ids = childrenNamed(group, CHART_NS, "axId").map((a) => valOf(a));
		if (axisPair === null && ids.length >= 2) {
			const pair = ids.slice(0, 2).map((id) => id === void 0 ? void 0 : axes.get(id));
			const xi = pair.findIndex((a) => a !== void 0 && a.local !== "valAx");
			const xIndex = xi === -1 ? 0 : xi;
			axisPair = {
				x: axisOf(pair[xIndex]),
				y: axisOf(pair[1 - xIndex])
			};
		}
	}
	const legend = childNamed(chart, CHART_NS, "legend");
	return Object.freeze({
		types: Object.freeze(types),
		stacked,
		title: titleOf(childNamed(chart, CHART_NS, "title")),
		legend: legend === void 0 ? null : LEGEND_POSITIONS.get(valOf(childNamed(legend, "http://schemas.openxmlformats.org/drawingml/2006/chart", "legendPos")) ?? "r") ?? "right",
		axes: Object.freeze(axisPair ?? {
			x: null,
			y: null
		}),
		series: Object.freeze(series)
	});
}
/**
* Elements of a chart part that are never read: the cached and literal values of data sources
* (CH-18, CH-19) and extensions. They are parsed under the XML limits but not built, so a chart
* over many cells reads in time and memory linear in the part and within `maxXmlTreeElements`.
*/
const NOT_READ = /* @__PURE__ */ new Set([
	"numCache",
	"strCache",
	"multiLvlStrCache",
	"numLit",
	"strLit",
	"extLst"
]);
const skipNotRead = (ns, local) => ns === "http://schemas.openxmlformats.org/drawingml/2006/chart" && NOT_READ.has(local);
/**
* Reads the chart part `target` once per workbook (`cache`). A chart is optional content, so a
* part that cannot be read is skipped with a warning and `null`, and the workbook is read:
* - `XLV_W_PART_MISSING`: the part does not exist;
* - `XLV_W_INVALID_VALUE`: the part's content type is not the chart content type (CH-01), its
*   bytes are not valid UTF-8 or UTF-16 or declare another encoding, it is not well-formed XML,
*   or it has no `chartSpace` with a chart and a plot area;
* - `XLV_W_LIMIT_REACHED`: the part's XML is over a limit of `ctx.limits` that applies to one
*   part (`maxXmlTreeElements`, `maxXmlDepth`, `maxXmlAttributes`, `maxXmlNameLength`,
*   `maxXmlTextLength`); or the chart groups and series of the charts read so far, each counted
*   once per part, are over what the cells of the workbook leave of `maxCells`
*   (`cache.budget`): this chart and every chart part read after it are skipped, the later ones
*   without being parsed. Cached and literal values are parsed under the XML limits but do not
*   count against `maxXmlTreeElements`.
*
* Errors of the package still throw as for every part: ZIP damage (`XLV_E_ZIP_CORRUPT`), the
* budgets of the whole file (`maxEntryBytes`, `maxTotalBytes`, `maxCompressionRatio`, as
* `XLV_E_LIMIT_EXCEEDED`), and a DTD or entity declaration (`XLV_E_XML_DTD_FORBIDDEN`,
* `XLV_E_XML_ENTITY_FORBIDDEN`).
*/
function readChartPart(pkg, target, sheet, ctx, cache) {
	const key = partKey(target);
	const cached = cache.parts.get(key);
	if (cached !== void 0 || cache.parts.has(key)) return cached ?? null;
	cache.parts.set(key, null);
	const where = {
		part: target,
		sheet
	};
	const overBudget = () => {
		cache.budget.remaining = 0;
		ctx.warnings.add("XLV_W_LIMIT_REACHED", "The charts of the workbook have more chart groups and series than the cells leave of the limit maxCells; this chart is skipped, as are the charts read after it. If you trust this file, raise limits.maxCells to read them.", where);
		return null;
	};
	if (cache.budget.remaining <= 0) return overBudget();
	if (!pkg.hasPart(target)) {
		ctx.warnings.add("XLV_W_PART_MISSING", "A chart of the sheet refers to a chart part that does not exist; the chart is skipped.", where);
		return null;
	}
	const type = pkg.contentType(target);
	if (type !== "application/vnd.openxmlformats-officedocument.drawingml.chart+xml") {
		ctx.warnings.add("XLV_W_INVALID_VALUE", `A chart of the sheet refers to a part of content type "${clip(type ?? "(none)")}", not a chart part; the chart is skipped.`, where);
		return null;
	}
	const bytes = pkg.readBytes(target);
	let root;
	try {
		root = parseXmlTree(decodeXml(bytes, target), {
			part: target,
			limits: ctx.limits,
			skip: skipNotRead
		});
	} catch (error) {
		if (error instanceof XlvError && (error.code === "XLV_E_XML_MALFORMED" || error.code === "XLV_E_XML_ENCODING")) {
			ctx.warnings.add("XLV_W_INVALID_VALUE", `A chart part is not ${error.code === "XLV_E_XML_ENCODING" ? "encoded in UTF-8 or UTF-16" : "well-formed XML"}; the chart is skipped.`, where);
			return null;
		}
		if (error instanceof XlvError && error.code === "XLV_E_LIMIT_EXCEEDED") {
			const limit = typeof error.details["limit"] === "string" ? error.details["limit"] : "a read limit";
			ctx.warnings.add("XLV_W_LIMIT_REACHED", `A chart part is over the limit ${limit}; the chart is skipped. If you trust this file, raise limits.${limit} to read it.`, where);
			return null;
		}
		throw error;
	}
	const used = { count: 0 };
	const parsed = root.ns === "http://schemas.openxmlformats.org/drawingml/2006/chart" && root.local === "chartSpace" ? chartOf(root, cache.budget.remaining, used) : null;
	if (parsed === OVER_BUDGET) return overBudget();
	cache.budget.remaining -= used.count;
	if (parsed === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `A chart part has no chart (root <${root.local.slice(0, 40)}>, or no chart with a plot area); the chart is skipped.`, where);
	cache.parts.set(key, parsed);
	return parsed;
}
// src/sml/namespaces.ts
/** SpreadsheetML main namespace. */
const SML_NS = {
	strict: "http://purl.oclc.org/ooxml/spreadsheetml/main",
	transitional: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
};
/** Namespace of `r:id` and similar relationship-reference attributes. */
const REL_NS = {
	strict: "http://purl.oclc.org/ooxml/officeDocument/relationships",
	transitional: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
};
/** OPC namespaces (ECMA-376 Part 2 schemas opc-contentTypes.xsd and opc-relationships.xsd). */
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
/** Full relationship type URIs for a name, in both conformance classes. */
function relationshipTypes(name) {
	return [STRICT_REL + name, TRANSITIONAL_REL + name];
}
/** Returns the relationship name and conformance class of a relationship type URI, if known. */
function classifyRelationship(type) {
	if (type.startsWith(STRICT_REL)) return {
		name: type.slice(56),
		conformance: "strict"
	};
	if (type.startsWith(TRANSITIONAL_REL)) return {
		name: type.slice(68),
		conformance: "transitional"
	};
	return null;
}
const MACROSHEET_CONTENT_TYPES = ["application/vnd.ms-excel.macrosheet+xml", "application/vnd.ms-excel.intlmacrosheet+xml"];
// src/images.ts
/** Content type of a Drawings part (Part 1 §12.3.8, TD-30). */
const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
/**
* Content types of the image formats (Part 1 §15.2.14 note, TD-31); EMF and WMF as Excel declares
* them in the files of the POI corpus (TD-31a).
*/
const IMAGE_CONTENT_TYPES = {
	png: "image/png",
	jpeg: "image/jpeg",
	gif: "image/gif",
	emf: "image/x-emf",
	wmf: "image/x-wmf"
};
/**
* EMF or WMF recognised from the first bytes, as observed in the image parts of the POI corpus
* (TD-31a): EMF starts with the record type 1 (`01 00 00 00`) and has " EMF" at offset 40; a
* placeable WMF starts with `D7 CD C6 9A`.
*/
function metafileFormat(bytes) {
	const b = (i) => bytes[i] ?? -1;
	if (b(0) === 1 && b(1) === 0 && b(2) === 0 && b(3) === 0 && b(40) === 32 && b(41) === 69 && b(42) === 77 && b(43) === 70) return "emf";
	if (b(0) === 215 && b(1) === 205 && b(2) === 198 && b(3) === 154) return "wmf";
	return null;
}
const XDR_NS = {
	transitional: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
	strict: "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
};
const A_NS = {
	transitional: "http://schemas.openxmlformats.org/drawingml/2006/main",
	strict: "http://purl.oclc.org/ooxml/drawingml/main"
};
/** EMU per pixel at 96 pixels per inch: 914 400 / 96 (TD-33; the pixel size is a convention). */
const EMU_PER_PIXEL = 9525;
/** Last 0-based column and row of a drawing marker ([MS-OI29500] §2.1.1358 ff., TD-32). */
const MAX_MARKER_COL = 16383;
const MAX_MARKER_ROW = 1048575;
/** ST_PositiveCoordinate maximum (dml-main.xsd L204–209). */
const MAX_COORDINATE = 27273042316900;
/**
* The format of an image from its first bytes (TD-31): PNG `89 50 4E 47 0D 0A 1A 0A`, JPEG
* `FF D8 FF`, GIF `GIF87a` or `GIF89a`. These signatures were observed in the image files of the
* POI corpus and of LibreOffice 24.2.7.2; the format specifications were not available.
*
* @example
* ```ts
* import { imageFormat } from 'xlvane';
*
* console.log(imageFormat(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]))); // 'gif'
* ```
*/
function imageFormat(bytes) {
	const b = (i) => bytes[i] ?? -1;
	if (b(0) === 137 && b(1) === 80 && b(2) === 78 && b(3) === 71 && b(4) === 13 && b(5) === 10 && b(6) === 26 && b(7) === 10) return "png";
	if (b(0) === 255 && b(1) === 216 && b(2) === 255) return "jpeg";
	if (b(0) === 71 && b(1) === 73 && b(2) === 70 && b(3) === 56 && (b(4) === 55 || b(4) === 57) && b(5) === 97) return "gif";
	return null;
}
function mediaOf(pkg, rel, media) {
	if (rel === void 0 || rel.external) return null;
	const cached = media.get(rel.target);
	if (cached !== void 0) return cached;
	const bytes = pkg.readBytes(rel.target);
	const image = bytes === void 0 ? null : Object.freeze({
		part: rel.target,
		format: imageFormat(bytes),
		contentType: pkg.contentType(rel.target) ?? null,
		bytes
	});
	media.set(rel.target, image);
	return image;
}
/** Parses an ST_Coordinate: an integer in EMU, or a number with in, cm, mm or pt (TD-33). */
function coordinate(text) {
	const t = (text ?? "").trim();
	if (/^-?\d+$/.test(t)) return Number(t);
	const m = /^(-?\d+(?:\.\d+)?)(in|cm|mm|pt)$/.exec(t);
	if (m === null) return null;
	const per = {
		in: 914400,
		cm: 36e4,
		mm: 36e3,
		pt: 12700
	}[m[2]];
	return Math.round(Number(m[1]) * per);
}
/**
* Reads the Drawings part and the background picture of a worksheet. Pictures and charts (a
* `graphicFrame` whose `graphicData` holds a `c:chart`, CH-04) are read; other objects (shapes,
* connectors, group shapes, other graphic frames, content parts, alternate content) are counted
* and reported once with `XLV_W_DRAWING_CONTENT_NOT_READ`. Broken markers skip their object with
* a warning; a malformed or foreign drawing part gives a warning and no pictures or charts. A
* drawing part belongs to the first worksheet whose `<drawing>` shows it (`owners`, keyed by the
* case-folded part name, D-H01, D-K03): a later sheet that shows it gets no pictures or charts
* from it and `XLV_W_SHEET_PART_SHARED`, and the part is not read again. A chart part is parsed once per workbook
* (`chartCache`), however many frames show it, within the budget of `chartCache.budget`.
*/
function readSheetDrawing(pkg, sheetPart, sheet, ids, ctx, media, chartCache, owners) {
	if (ids.drawing === null && ids.picture === null) return {
		images: [],
		charts: [],
		background: null
	};
	const rels = pkg.relationships(sheetPart);
	const byId = (id, name) => id === null ? void 0 : rels.find((r) => r.id === id && classifyRelationship(r.type)?.name === name);
	const background = mediaOf(pkg, byId(ids.picture, "image"), media);
	const drawing = byId(ids.drawing, "drawing");
	if (drawing === void 0 || drawing.external) {
		if (ids.drawing !== null) ctx.warnings.add("XLV_W_PART_MISSING", "The sheet refers to a drawing that is not a drawing part of the package; its pictures are not read.", {
			part: sheetPart,
			sheet
		});
		return {
			images: [],
			charts: [],
			background
		};
	}
	const owner = owners.get(partKey(drawing.target));
	if (owner === void 0) owners.set(partKey(drawing.target), sheet);
	else if (owner !== sheet) {
		ctx.warnings.add("XLV_W_SHEET_PART_SHARED", `Sheet "${sheet}" shows the drawing ${drawing.target}, which belongs to sheet "${owner}"; it is read for "${owner}" only.`, {
			sheet,
			part: drawing.target
		});
		return {
			images: [],
			charts: [],
			background
		};
	}
	const xml = pkg.readXml(drawing.target);
	if (xml === void 0) {
		ctx.warnings.add("XLV_W_PART_MISSING", "The drawing part of the sheet does not exist; its pictures are not read.", {
			part: drawing.target,
			sheet
		});
		return {
			images: [],
			charts: [],
			background
		};
	}
	const where = {
		part: drawing.target,
		sheet
	};
	const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, where);
	let root;
	try {
		root = parseXmlTree(xml, {
			part: drawing.target,
			limits: ctx.limits
		});
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			warn("The drawing part of the sheet is not well-formed XML; its pictures are not read.");
			return {
				images: [],
				charts: [],
				background
			};
		}
		throw error;
	}
	const xdr = root.ns === XDR_NS.strict ? XDR_NS.strict : XDR_NS.transitional;
	const a = xdr === XDR_NS.strict ? A_NS.strict : A_NS.transitional;
	const r = xdr === XDR_NS.strict ? REL_NS.strict : REL_NS.transitional;
	if (root.local !== "wsDr" || root.ns !== xdr) {
		warn(`The drawing part of the sheet has the root element <${clip(root.local)}>; its pictures are not read.`);
		return {
			images: [],
			charts: [],
			background
		};
	}
	const drawingRels = pkg.relationships(drawing.target);
	const drawingRelsById = /* @__PURE__ */ new Map();
	for (const rel of drawingRels) if (!drawingRelsById.has(rel.id)) drawingRelsById.set(rel.id, rel);
	const relOfType = (id, name) => {
		const rel = id === void 0 ? void 0 : drawingRelsById.get(id);
		return rel !== void 0 && classifyRelationship(rel.type)?.name === name ? rel : void 0;
	};
	const images = [];
	const charts = [];
	let other = 0;
	const marker = (el) => {
		if (el === void 0) return null;
		const int = (local, max) => {
			const t = (childNamed(el, xdr, local)?.text ?? "").trim();
			return /^\d+$/.test(t) && Number(t) <= max ? Number(t) : null;
		};
		const col = int("col", MAX_MARKER_COL);
		const row = int("row", MAX_MARKER_ROW);
		const colOffset = coordinate(childNamed(el, xdr, "colOff")?.text);
		const rowOffset = coordinate(childNamed(el, xdr, "rowOff")?.text);
		if (col === null || row === null || colOffset === null || rowOffset === null) return null;
		return Object.freeze({
			cell: formatCellRef(row + 1, col + 1),
			col,
			row,
			colOffset,
			rowOffset
		});
	};
	const size = (el) => {
		const w = coordinate(el?.attrs.get("cx"));
		const h = coordinate(el?.attrs.get("cy"));
		return w === null || h === null || w < 0 || h < 0 ? null : Object.freeze({
			width: w,
			height: h
		});
	};
	for (const el of root.children) {
		if (el.ns !== xdr || el.local !== "twoCellAnchor" && el.local !== "oneCellAnchor" && el.local !== "absoluteAnchor") {
			if (el.ns !== xdr || el.local !== "extLst") other++;
			continue;
		}
		const pic = childNamed(el, xdr, "pic");
		const frame = pic === void 0 ? childNamed(el, xdr, "graphicFrame") : void 0;
		const graphicData = frame === void 0 ? void 0 : childNamed(childNamed(frame, a, "graphic") ?? frame, a, "graphicData");
		const chartEl = graphicData?.attrs.get("uri") === "http://schemas.openxmlformats.org/drawingml/2006/chart" ? childNamed(graphicData, CHART_NS, "chart") : void 0;
		if (pic === void 0 && chartEl === void 0) {
			other++;
			continue;
		}
		let anchor = null;
		if (el.local === "twoCellAnchor") {
			const from = marker(childNamed(el, xdr, "from"));
			const to = marker(childNamed(el, xdr, "to"));
			const editAs = el.attrs.get("editAs") ?? "twoCell";
			if (from !== null && to !== null) anchor = {
				type: "twoCell",
				from,
				to,
				editAs: editAs === "oneCell" || editAs === "absolute" ? editAs : "twoCell"
			};
		} else if (el.local === "oneCellAnchor") {
			const from = marker(childNamed(el, xdr, "from"));
			const ext = size(childNamed(el, xdr, "ext"));
			if (from !== null && ext !== null) anchor = {
				type: "oneCell",
				from,
				size: ext
			};
		} else {
			const pos = childNamed(el, xdr, "pos");
			const x = coordinate(pos?.attrs.get("x"));
			const y = coordinate(pos?.attrs.get("y"));
			const ext = size(childNamed(el, xdr, "ext"));
			if (x !== null && y !== null && ext !== null) anchor = {
				type: "absolute",
				position: Object.freeze({
					x,
					y
				}),
				size: ext
			};
		}
		if (anchor === null) {
			warn(`A ${pic === void 0 ? "chart" : "picture"} of the sheet has an invalid position; it is skipped.`);
			continue;
		}
		if (pic === void 0) {
			const chartRel = relOfType(chartEl.attrs.get(`{${r}}id`), "chart");
			if (chartRel === void 0 || chartRel.external) {
				warn("A chart of the sheet does not refer to a chart part of the package; it is skipped.");
				continue;
			}
			const parsed = readChartPart(pkg, chartRel.target, sheet, ctx, chartCache);
			if (parsed === null) continue;
			const nvFrame = childNamed(childNamed(frame, xdr, "nvGraphicFramePr") ?? frame, xdr, "cNvPr");
			charts.push(Object.freeze({
				name: nvFrame?.attrs.get("name") ?? "",
				description: nvFrame?.attrs.get("descr") ?? null,
				anchor: Object.freeze(anchor),
				...parsed
			}));
			continue;
		}
		const nv = childNamed(childNamed(pic, xdr, "nvPicPr") ?? pic, xdr, "cNvPr");
		const embed = childNamed(childNamed(pic, xdr, "blipFill") ?? pic, a, "blip")?.attrs.get(`{${r}}embed`);
		const imageRel = relOfType(embed, "image");
		const linkRel = relOfType(nv === void 0 ? void 0 : childNamed(nv, a, "hlinkClick")?.attrs.get(`{${r}}id`), "hyperlink");
		images.push(Object.freeze({
			name: nv?.attrs.get("name") ?? "",
			description: nv?.attrs.get("descr") ?? null,
			anchor: Object.freeze(anchor),
			image: mediaOf(pkg, imageRel, media),
			hyperlink: linkRel?.target ?? null
		}));
	}
	if (other > 0) ctx.warnings.add("XLV_W_DRAWING_CONTENT_NOT_READ", `The drawing of the sheet has ${String(other)} object${other === 1 ? "" : "s"} other than pictures and charts (shapes, groups or other objects); xlvane reports only pictures and charts, and editing keeps the drawing unchanged.`, where);
	return {
		images,
		charts,
		background
	};
}
function invalid$5(message, hint, details, kind = "picture") {
	return new XlvError(kind === "chart" ? "XLV_E_INVALID_CHART" : "XLV_E_INVALID_OPTION", kind === "chart" ? message.replace("picture", "chart") : message, {
		hint,
		...details ? { details } : {}
	});
}
/** Checks the image bytes given to `NewWorkbook.addImage` and returns their format. */
function checkImage(bytes) {
	const data = bytes instanceof Uint8Array ? bytes : bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : null;
	if (data === null) throw new XlvError("XLV_E_INVALID_IMAGE", "addImage expects the bytes of an image (a Uint8Array).", { hint: "Pass the file contents, for example new Uint8Array(await file.arrayBuffer())." });
	const format = imageFormat(data) ?? metafileFormat(data);
	if (format === null) throw new XlvError("XLV_E_INVALID_IMAGE", "The image is not a PNG, JPEG, GIF, EMF or WMF file (recognised from its first bytes).", {
		hint: "Convert the image to PNG or JPEG first; xlvane writes PNG, JPEG and GIF images, and EMF and WMF images as they are.",
		details: { size: data.length }
	});
	return {
		bytes: data,
		format
	};
}
/** Validates the options of `NewSheet.addImage` (the anchor is resolved when the sheet is written). */
function normalizeImage(image, count, input, options) {
	if (typeof image !== "number" || !Number.isInteger(image) || image < 0 || image >= count) throw invalid$5(`Image ${String(image).slice(0, 20)} was not added to this workbook.`, "Pass the number that workbook.addImage(bytes) returned.", { image: String(image).slice(0, 20) });
	const anchor = copyAnchor(input);
	checkAnchor(anchor);
	let name = null;
	let description = null;
	let hyperlink = null;
	if (options !== void 0) {
		if (typeof options !== "object" || options === null) throw invalid$5("The image options must be an object.", "For example { description: \"Logo\" }.");
		for (const [key, value] of Object.entries(options)) {
			if (key !== "name" && key !== "description" && key !== "hyperlink") throw invalid$5(`Unknown image option "${key.slice(0, 40)}".`, "Use name, description or hyperlink.", { option: key.slice(0, 40) });
			if (value === void 0) continue;
			const control = key === "hyperlink" ? /[\u0000-\u001f\u007f\ufffe\uffff]/ : /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/;
			if (typeof value !== "string" || value.length > 2084 || control.test(value) || firstLoneSurrogate(value) !== -1 || key !== "description" && value === "") throw invalid$5(`The image option "${key}" must be a text of 1 to 2084 characters without control or non-characters.`, key === "hyperlink" ? "For example hyperlink: \"https://example.com\"." : "For example name: \"Logo\".", { option: key });
			if (key === "name") name = value;
			else if (key === "description") description = value;
			else hyperlink = value;
		}
	}
	return {
		image,
		anchor,
		name,
		description,
		hyperlink
	};
}
function position(p, what, kind = "picture") {
	if (typeof p === "string") {
		const at = parseCellRef(p.trim().toUpperCase());
		if (at === null) throw invalid$5(`The picture ${what} "${p.slice(0, 20)}" is not a cell.`, "Use a cell such as \"B2\" or a position such as { col: 1, row: 1 }.", { option: what }, kind);
		return {
			col: at.col - 1,
			row: at.row - 1
		};
	}
	if (typeof p === "object" && p !== null) {
		const { col, row } = p;
		if (Object.keys(p).every((k) => k === "col" || k === "row") && typeof col === "number" && typeof row === "number" && Number.isFinite(col) && Number.isFinite(row) && col >= 0 && row >= 0 && col < 16384 && row < 1048576) return {
			col,
			row
		};
	}
	throw invalid$5(`The picture ${what} is not a valid position.`, "Use a cell such as \"B2\", or { col, row } with 0-based numbers (col below 16384, row below 1048576).", { option: what }, kind);
}
/**
* A copy of a position as given, taken when the picture or chart is added, so that later changes
* to the caller's objects do not change it (and cannot bypass {@link checkAnchor}). Objects are
* copied two levels deep (the anchor and its `from`, `to` and `size`); unknown keys are kept, so
* that `checkAnchor` refuses them as before.
*/
function copyAnchor(anchor) {
	const copy = (value) => typeof value === "object" && value !== null && !Array.isArray(value) ? Object.freeze(Object.fromEntries(Object.keys(value).map((k) => [k, value[k]]))) : value;
	if (typeof anchor !== "object" || anchor === null || Array.isArray(anchor)) return anchor;
	return Object.freeze(Object.fromEntries(Object.keys(anchor).map((k) => [k, copy(anchor[k])])));
}
/** Checks the position of a picture (`XLV_E_INVALID_OPTION`) or a chart (`XLV_E_INVALID_CHART`). */
function checkAnchor(anchor, kind = "picture") {
	const bad = (message, hint, option) => invalid$5(message, hint, { option }, kind);
	if (typeof anchor === "string") {
		if (parseRange(anchor.trim().toUpperCase()) === null) throw bad(`The picture range "${anchor.slice(0, 40)}" is not valid.`, "Use a range such as \"B2:D6\".", "anchor");
		return;
	}
	if (typeof anchor !== "object" || anchor === null) throw bad("The picture position must be a range or an object.", "For example \"B2:D6\", { from: \"B2\", to: \"E8\" } or { from: \"B2\", size: { width: 160, height: 90 } }.", "anchor");
	const o = anchor;
	for (const key of Object.keys(o)) if (![
		"from",
		"to",
		"size",
		"editAs"
	].includes(key)) throw bad(`Unknown picture position option "${key.slice(0, 40)}".`, "Use from with to (and editAs), or from with size.", key.slice(0, 40));
	const from = position(o["from"], "from", kind);
	if (o["to"] === void 0 === (o["size"] === void 0)) throw bad("A picture position needs either \"to\" or \"size\".", "For example { from: \"B2\", to: \"E8\" } or { from: \"B2\", size: { width: 160, height: 90 } }.", "anchor");
	if (o["to"] !== void 0) {
		const to = position(o["to"], "to", kind);
		if (to.col < from.col || to.row < from.row) throw bad("The picture corner \"to\" is above or left of \"from\".", "Give the top-left corner as from and the bottom-right corner as to.", "to");
		const editAs = o["editAs"];
		if (editAs !== void 0 && editAs !== "twoCell" && editAs !== "oneCell" && editAs !== "absolute") throw bad("editAs must be \"twoCell\", \"oneCell\" or \"absolute\".", "Use editAs: \"oneCell\" for an object that moves but does not resize with its cells.", "editAs");
	} else {
		if (o["editAs"] !== void 0) throw bad("editAs applies to positions with two corners only.", "Use { from, to, editAs } or leave editAs out.", "editAs");
		const s = o["size"];
		const ok = (v) => typeof v === "number" && Number.isFinite(v) && Math.round(v * 9525) >= 1 && v * 9525 <= MAX_COORDINATE;
		if (typeof s !== "object" || s === null || !ok(s.width) || !ok(s.height)) throw bad("The picture size must be { width, height } in pixels, both at least 1 EMU (1/9525 pixel).", "For example size: { width: 160, height: 90 }.", "size");
	}
}
/** Metrics from the column widths and row heights of a sheet being written (1-based maps). */
function metricsOf(columns, rows) {
	return {
		colPixels(col) {
			const c = columns.get(col + 1);
			if (c?.hidden) return 0;
			return c?.width == null ? 64 : Math.round(c.width * 7);
		},
		rowPixels(row) {
			const r = rows.get(row + 1);
			if (r?.hidden) return 0;
			return r?.height == null ? 20 : Math.round(r.height * 96 / 72);
		}
	};
}
function toMarker(p, m) {
	const col = Math.floor(p.col);
	const row = Math.floor(p.row);
	return {
		col,
		row,
		colOff: Math.round((p.col - col) * m.colPixels(col) * EMU_PER_PIXEL),
		rowOff: Math.round((p.row - row) * m.rowPixels(row) * EMU_PER_PIXEL)
	};
}
/** Resolves the anchor of a stored picture with the sheet's column widths and row heights. */
function resolveImage(s, m) {
	return {
		image: s.image,
		name: s.name,
		description: s.description,
		hyperlink: s.hyperlink,
		...resolveAnchor(s.anchor, m)
	};
}
/** Resolves a checked anchor (of a picture or a chart) with the sheet's column widths and row heights. */
function resolveAnchor(anchor, m) {
	if (typeof anchor === "string") {
		const range = parseRange(anchor.trim().toUpperCase());
		const toCol = range.end.col > MAX_MARKER_COL ? {
			col: MAX_MARKER_COL,
			colOff: m.colPixels(MAX_MARKER_COL) * EMU_PER_PIXEL
		} : {
			col: range.end.col,
			colOff: 0
		};
		const toRow = range.end.row > MAX_MARKER_ROW ? {
			row: MAX_MARKER_ROW,
			rowOff: m.rowPixels(MAX_MARKER_ROW) * EMU_PER_PIXEL
		} : {
			row: range.end.row,
			rowOff: 0
		};
		return {
			from: {
				col: range.start.col - 1,
				row: range.start.row - 1,
				colOff: 0,
				rowOff: 0
			},
			to: {
				...toCol,
				...toRow
			},
			ext: null,
			editAs: null
		};
	}
	const from = toMarker(position(anchor.from, "from"), m);
	if ("to" in anchor && anchor.to !== void 0) return {
		from,
		to: toMarker(position(anchor.to, "to"), m),
		ext: null,
		editAs: anchor.editAs ?? null
	};
	const size = anchor.size;
	return {
		from,
		to: null,
		ext: {
			cx: Math.round(size.width * EMU_PER_PIXEL),
			cy: Math.round(size.height * EMU_PER_PIXEL)
		},
		editAs: null
	};
}
const XML_DECL$1 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const REL_BASE$1 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const attr$1 = (name, value) => ` ${name}="${typeof value === "number" ? String(value) : escapeAttr(value)}"`;
/**
* The Image parts of a package: each distinct image (same format and bytes) is stored once, as
* `xl/media/imageN.<png|jpeg|gif>` (TD-31), in the order the pictures use them.
*/
var MediaParts = class {
	images;
	byIndex = /* @__PURE__ */ new Map();
	byKey = /* @__PURE__ */ new Map();
	list = [];
	constructor(images) {
		this.images = images;
	}
	/** Target of image `index` relative to `xl/drawings/` or `xl/worksheets/`. */
	target(index) {
		let n = this.byIndex.get(index);
		if (n === void 0) {
			const img = this.images[index];
			let h = 2166136261;
			for (const b of img.bytes) h = Math.imul(h ^ b, 16777619);
			const key = `${img.format}:${String(img.bytes.length)}:${String(h >>> 0)}`;
			const same = (this.byKey.get(key) ?? []).find((k) => sameBytes$1(this.list[k - 1].bytes, img.bytes));
			if (same !== void 0) n = same;
			else {
				this.list.push(img);
				n = this.list.length;
				this.byKey.set(key, [...this.byKey.get(key) ?? [], n]);
			}
			this.byIndex.set(index, n);
		}
		return `../media/image${String(n)}.${this.list[n - 1].format}`;
	}
	/** The Image parts as ZIP entries. */
	entries() {
		return this.list.map((img, i) => ({
			name: `xl/media/image${String(i + 1)}.${img.format}`,
			data: img.bytes
		}));
	}
	/** The formats used, for the `Default` content types. */
	formats() {
		return [...new Set(this.list.map((img) => img.format))];
	}
};
function sameBytes$1(a, b) {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
	return true;
}
function markerXml(local, m) {
	return `<xdr:${local}><xdr:col>${String(m.col)}</xdr:col><xdr:colOff>${String(m.colOff)}</xdr:colOff><xdr:row>${String(m.row)}</xdr:row><xdr:rowOff>${String(m.rowOff)}</xdr:rowOff></xdr:${local}>`;
}
/** An anchor element around an object (TD-32): `twoCellAnchor` with its markers or `oneCellAnchor` with its size. */
function anchorXml(p, body) {
	if (p.to !== null) return `<xdr:twoCellAnchor${p.editAs === null ? "" : attr$1("editAs", p.editAs)}>${markerXml("from", p.from)}${markerXml("to", p.to)}${body}<xdr:clientData/></xdr:twoCellAnchor>`;
	return `<xdr:oneCellAnchor>${markerXml("from", p.from)}<xdr:ext${attr$1("cx", p.ext.cx)}${attr$1("cy", p.ext.cy)}/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
}
/**
* Builds the Drawings part `xl/drawings/drawing<n>.xml` (Part 1 §20.5, TD-35) for the pictures and
* charts of a sheet, adds the `drawing` and background `image` relationships to `rels`, and
* returns the sheet elements. Each picture is an `xdr:pic` in a `twoCellAnchor` or
* `oneCellAnchor`, with a unique `cNvPr` id, `a:blip r:embed` to its image and
* `a:stretch/a:fillRect`, as in the Part 1 Annex L picture example; one image relationship per
* image and drawing. Each chart, after the pictures, is an `xdr:graphicFrame` as Excel writes it
* (CH-04) with a `chart` relationship from the same counter (CH-02).
*/
function drawingParts(pictures, background, media, n, rels, charts = []) {
	const picture = background === null ? "" : `<picture${attr$1("r:id", rels.add("image", media.target(background)))}/>`;
	if (pictures.length === 0 && charts.length === 0) return {
		drawing: "",
		picture,
		part: null
	};
	const items = [];
	const imageIds = /* @__PURE__ */ new Map();
	const add = (type, target, external = false) => {
		const id = `rId${String(items.length + 1)}`;
		items.push(`<Relationship${attr$1("Id", id)}${attr$1("Type", REL_BASE$1 + type)}${attr$1("Target", target)}${external ? " TargetMode=\"External\"" : ""}/>`);
		return id;
	};
	let xml = `${XML_DECL$1}<xdr:wsDr xmlns:xdr="${XDR_NS.transitional}" xmlns:a="${A_NS.transitional}" xmlns:r="${REL_BASE$1.slice(0, -1)}">`;
	pictures.forEach((p, i) => {
		let embed = imageIds.get(p.image);
		if (embed === void 0) {
			embed = add("image", media.target(p.image));
			imageIds.set(p.image, embed);
		}
		const link = p.hyperlink === null ? "" : `<a:hlinkClick${attr$1("r:id", add("hyperlink", p.hyperlink, true))}/>`;
		const body = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr${attr$1("id", i + 1)}${attr$1("name", p.name ?? `Picture ${String(i + 1)}`)}${p.description === null ? "" : attr$1("descr", p.description)}>${link}</xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip${attr$1("r:embed", embed)}/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
		xml += anchorXml(p, body);
	});
	charts.forEach((c, k) => {
		const id = add("chart", c.target);
		const body = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr${attr$1("id", pictures.length + k + 1)}${attr$1("name", c.name ?? `Chart ${String(k + 1)}`)}${c.description === null ? "" : attr$1("descr", c.description)}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" xmlns:r="${REL_BASE$1.slice(0, -1)}"${attr$1("r:id", id)}/></a:graphicData></a:graphic></xdr:graphicFrame>`;
		xml += anchorXml(c, body);
	});
	xml += "</xdr:wsDr>";
	const name = `xl/drawings/drawing${String(n)}.xml`;
	return {
		drawing: `<drawing${attr$1("r:id", rels.add("drawing", `../drawings/drawing${String(n)}.xml`))}/>`,
		picture,
		part: {
			name,
			xml,
			relsName: drawingRelsPartName(n),
			rels: `${XML_DECL$1}<Relationships xmlns="${PACKAGE_REL_NS}">${items.join("")}</Relationships>`
		}
	};
}
/** Part name of the relationships part of drawing `n`. */
function drawingRelsPartName(n) {
	return `xl/drawings/_rels/drawing${String(n)}.xml.rels`;
}
// src/util/sha.ts
/** @internal */
const SHA_ALGORITHMS = [
	"SHA-1",
	"SHA-256",
	"SHA-384",
	"SHA-512"
];
let tables = null;
function iroot(n, k) {
	let x = 1n << BigInt(n.toString(2).length) / k + 1n;
	for (;;) {
		const y = ((k - 1n) * x + n / x ** (k - 1n)) / k;
		if (y >= x) return x;
		x = y;
	}
}
function getTables() {
	if (tables) return tables;
	const primes = [];
	for (let n = 2; primes.length < 80; n++) if (primes.every((p) => n % p !== 0)) primes.push(n);
	const mask = (1n << 64n) - 1n;
	const pair = (out, i, v) => {
		out[2 * i] = Number(v >> 32n & 4294967295n);
		out[2 * i + 1] = Number(v & 4294967295n);
	};
	const k512 = /* @__PURE__ */ new Uint32Array(160);
	const h512 = /* @__PURE__ */ new Uint32Array(16);
	const h384 = /* @__PURE__ */ new Uint32Array(16);
	const k256 = /* @__PURE__ */ new Uint32Array(64);
	const h256 = /* @__PURE__ */ new Uint32Array(8);
	primes.forEach((p, i) => pair(k512, i, iroot(BigInt(p) << 192n, 3n) & mask));
	for (let i = 0; i < 8; i++) {
		pair(h512, i, iroot(BigInt(primes[i]) << 128n, 2n) & mask);
		pair(h384, i, iroot(BigInt(primes[i + 8]) << 128n, 2n) & mask);
	}
	for (let i = 0; i < 64; i++) k256[i] = k512[2 * i];
	for (let i = 0; i < 8; i++) h256[i] = h512[2 * i];
	tables = {
		k512,
		h512,
		h384,
		k256,
		h256
	};
	return tables;
}
function pad(data, blockBytes, lengthBytes) {
	const total = Math.ceil((data.length + 1 + lengthBytes) / blockBytes) * blockBytes;
	const out = new Uint8Array(total);
	out.set(data);
	out[data.length] = 128;
	const bits = data.length * 8;
	const view = new DataView(out.buffer);
	view.setUint32(total - 4, bits >>> 0);
	view.setUint32(total - 8, Math.floor(bits / 4294967296));
	return out;
}
function sha1(data) {
	const m = pad(data, 64, 8);
	const view = new DataView(m.buffer);
	const h = new Uint32Array([
		1732584193,
		4023233417,
		2562383102,
		271733878,
		3285377520
	]);
	const w = /* @__PURE__ */ new Uint32Array(80);
	const k = [
		1518500249,
		1859775393,
		2400959708,
		3395469782
	];
	for (let off = 0; off < m.length; off += 64) {
		for (let t = 0; t < 16; t++) w[t] = view.getUint32(off + 4 * t);
		for (let t = 16; t < 80; t++) {
			const x = w[t - 3] ^ w[t - 8] ^ w[t - 14] ^ w[t - 16];
			w[t] = x << 1 | x >>> 31;
		}
		let a = h[0], b = h[1], c = h[2], d = h[3], e = h[4];
		for (let t = 0; t < 80; t++) {
			const f = t < 20 ? b & c | ~b & d : t < 40 ? b ^ c ^ d : t < 60 ? b & c | b & d | c & d : b ^ c ^ d;
			const tmp = (a << 5 | a >>> 27) + f + e + k[t / 20 | 0] + w[t] >>> 0;
			e = d;
			d = c;
			c = b << 30 | b >>> 2;
			b = a;
			a = tmp;
		}
		h[0] = h[0] + a;
		h[1] = h[1] + b;
		h[2] = h[2] + c;
		h[3] = h[3] + d;
		h[4] = h[4] + e;
	}
	return toBytes$1(h, 5);
}
function sha256(data) {
	const { k256, h256 } = getTables();
	const m = pad(data, 64, 8);
	const view = new DataView(m.buffer);
	const h = Uint32Array.from(h256);
	const w = /* @__PURE__ */ new Uint32Array(64);
	const rotr = (x, n) => x >>> n | x << 32 - n;
	for (let off = 0; off < m.length; off += 64) {
		for (let t = 0; t < 16; t++) w[t] = view.getUint32(off + 4 * t);
		for (let t = 16; t < 64; t++) {
			const x = w[t - 15];
			const y = w[t - 2];
			const s0 = rotr(x, 7) ^ rotr(x, 18) ^ x >>> 3;
			const s1 = rotr(y, 17) ^ rotr(y, 19) ^ y >>> 10;
			w[t] = w[t - 16] + s0 + w[t - 7] + s1;
		}
		let a = h[0], b = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], hh = h[7];
		for (let t = 0; t < 64; t++) {
			const t1 = hh + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + (e & f ^ ~e & g) + k256[t] + w[t] >>> 0;
			const t2 = (rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + (a & b ^ a & c ^ b & c) >>> 0;
			hh = g;
			g = f;
			f = e;
			e = d + t1 >>> 0;
			d = c;
			c = b;
			b = a;
			a = t1 + t2 >>> 0;
		}
		h[0] = h[0] + a;
		h[1] = h[1] + b;
		h[2] = h[2] + c;
		h[3] = h[3] + d;
		h[4] = h[4] + e;
		h[5] = h[5] + f;
		h[6] = h[6] + g;
		h[7] = h[7] + hh;
	}
	return toBytes$1(h, 8);
}
function sha512(data, init, words) {
	const { k512 } = getTables();
	const m = pad(data, 128, 16);
	const view = new DataView(m.buffer);
	const h = Uint32Array.from(init);
	const w = /* @__PURE__ */ new Uint32Array(160);
	const s = /* @__PURE__ */ new Uint32Array(16);
	for (let off = 0; off < m.length; off += 128) {
		for (let t = 0; t < 32; t++) w[t] = view.getUint32(off + 4 * t);
		for (let t = 16; t < 80; t++) {
			const xh = w[2 * (t - 15)], xl = w[2 * (t - 15) + 1];
			const s0h = (xh >>> 1 | xl << 31) ^ (xh >>> 8 | xl << 24) ^ xh >>> 7;
			const s0l = (xl >>> 1 | xh << 31) ^ (xl >>> 8 | xh << 24) ^ (xl >>> 7 | xh << 25);
			const yh = w[2 * (t - 2)], yl = w[2 * (t - 2) + 1];
			const s1h = (yh >>> 19 | yl << 13) ^ (yl >>> 29 | yh << 3) ^ yh >>> 6;
			const s1l = (yl >>> 19 | yh << 13) ^ (yh >>> 29 | yl << 3) ^ (yl >>> 6 | yh << 26);
			const lo = (s0l >>> 0) + (s1l >>> 0) + w[2 * (t - 16) + 1] + w[2 * (t - 7) + 1];
			w[2 * t + 1] = lo;
			w[2 * t] = (s0h >>> 0) + (s1h >>> 0) + w[2 * (t - 16)] + w[2 * (t - 7)] + Math.floor(lo / 4294967296);
		}
		s.set(h);
		for (let t = 0; t < 80; t++) {
			const ah = s[0], al = s[1], eh = s[8], el = s[9];
			const S1h = (eh >>> 14 | el << 18) ^ (eh >>> 18 | el << 14) ^ (el >>> 9 | eh << 23);
			const S1l = (el >>> 14 | eh << 18) ^ (el >>> 18 | eh << 14) ^ (eh >>> 9 | el << 23);
			const chh = eh & s[10] ^ ~eh & s[12];
			const chl = el & s[11] ^ ~el & s[13];
			const t1l = s[15] + (S1l >>> 0) + (chl >>> 0) + k512[2 * t + 1] + w[2 * t + 1];
			const t1h = s[14] + (S1h >>> 0) + (chh >>> 0) + k512[2 * t] + w[2 * t] + Math.floor(t1l / 4294967296);
			const S0h = (ah >>> 28 | al << 4) ^ (al >>> 2 | ah << 30) ^ (al >>> 7 | ah << 25);
			const S0l = (al >>> 28 | ah << 4) ^ (ah >>> 2 | al << 30) ^ (ah >>> 7 | al << 25);
			const mjh = ah & s[2] ^ ah & s[4] ^ s[2] & s[4];
			const mjl = al & s[3] ^ al & s[5] ^ s[3] & s[5];
			const t2l = (S0l >>> 0) + (mjl >>> 0);
			const t2h = (S0h >>> 0) + (mjh >>> 0) + Math.floor(t2l / 4294967296);
			s.copyWithin(2, 0, 14);
			const el2 = s[9] + (t1l >>> 0);
			s[9] = el2;
			s[8] = s[8] + t1h + Math.floor(el2 / 4294967296);
			const al2 = (t1l >>> 0) + (t2l >>> 0);
			s[1] = al2;
			s[0] = t1h + t2h + Math.floor(al2 / 4294967296);
		}
		for (let i = 0; i < 8; i++) {
			const lo = h[2 * i + 1] + s[2 * i + 1];
			h[2 * i + 1] = lo;
			h[2 * i] = h[2 * i] + s[2 * i] + Math.floor(lo / 4294967296);
		}
	}
	return toBytes$1(h, words);
}
function toBytes$1(h, words) {
	const out = new Uint8Array(words * 4);
	const view = new DataView(out.buffer);
	for (let i = 0; i < words; i++) view.setUint32(4 * i, h[i]);
	return out;
}
/** @internal The digest of `data` with `algorithm`. */
function sha(algorithm, data) {
	switch (algorithm) {
		case "SHA-1": return sha1(data);
		case "SHA-256": return sha256(data);
		case "SHA-384": return sha512(data, getTables().h384, 12);
		case "SHA-512": return sha512(data, getTables().h512, 16);
	}
}
// src/password.ts
const MAX_SPIN = 1e7;
const MAX_SALT = 65536;
const MAX_PASSWORD = 255;
function invalid$4(message, hint, option) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		hint,
		details: { option }
	});
}
const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
/** @internal Base64 (RFC 4648 §4) of `bytes`. */
function toBase64(bytes) {
	let out = "";
	for (let i = 0; i < bytes.length; i += 3) {
		const n = bytes[i] << 16 | (bytes[i + 1] ?? 0) << 8 | (bytes[i + 2] ?? 0);
		out += B64[n >> 18 & 63] + B64[n >> 12 & 63];
		out += i + 1 < bytes.length ? B64[n >> 6 & 63] : "=";
		out += i + 2 < bytes.length ? B64[n & 63] : "=";
	}
	return out;
}
/** @internal The bytes of base64 text (white space ignored), or `null` when it is not base64. */
function fromBase64(text) {
	const s = text.replace(/[\t\n\r ]/g, "");
	if (!/^[A-Za-z0-9+/]*={0,2}$/.test(s) || s.length % 4 !== 0) return null;
	const pad = s.endsWith("==") ? 2 : s.endsWith("=") ? 1 : 0;
	const out = new Uint8Array(s.length / 4 * 3 - pad);
	let o = 0;
	for (let i = 0; i < s.length; i += 4) {
		let n = 0;
		for (let j = 0; j < 4; j++) n = n << 6 | (s[i + j] === "=" ? 0 : B64.indexOf(s[i + j]));
		if (o < out.length) out[o++] = n >> 16 & 255;
		if (o < out.length) out[o++] = n >> 8 & 255;
		if (o < out.length) out[o++] = n & 255;
	}
	return out;
}
/** The password as hashed: UTF-16LE without a leading byte order mark (ECMA-376 Part 1, SR-51). */
function passwordBytes(password) {
	const p = password.startsWith("﻿") ? password.slice(1) : password;
	const out = new Uint8Array(p.length * 2);
	for (let i = 0; i < p.length; i++) {
		const c = p.charCodeAt(i);
		out[2 * i] = c & 255;
		out[2 * i + 1] = c >> 8;
	}
	return out;
}
/** [MS-OFFCRYPTO] §2.4.2.4: H0 = H(salt + password); Hn = H(Hn-1 + iterator), iterator 0 … spinCount − 1, 4 bytes little-endian. */
function isoHash(algorithm, salt, password, spinCount) {
	const first = new Uint8Array(salt.length + password.length);
	first.set(salt);
	first.set(password, salt.length);
	let h = sha(algorithm, first);
	const buf = new Uint8Array(h.length + 4);
	const view = new DataView(buf.buffer);
	for (let i = 0; i < spinCount; i++) {
		buf.set(h);
		view.setUint32(h.length, i, true);
		h = sha(algorithm, buf);
	}
	return h;
}
function checkPassword(password) {
	if (typeof password !== "string") throw invalid$4("The password must be a string.", "For example hashPassword('secret').", "password");
	const p = password.startsWith("﻿") ? password.slice(1) : password;
	if (p.length < 1 || p.length > MAX_PASSWORD) throw invalid$4(`The password has ${String(p.length)} characters; it must have 1 to 255.`, "Use a password of 1 to 255 characters ([MS-OFFCRYPTO] §2.4.2.4).", "password");
	return p;
}
function randomSalt() {
	const c = globalThis.crypto;
	if (typeof c?.getRandomValues !== "function") throw new XlvError("XLV_E_NOT_SUPPORTED", "This runtime has no crypto.getRandomValues, so a random salt cannot be made.", {
		hint: "Pass your own random salt: hashPassword(password, { salt }), 16 bytes from a secure random source.",
		details: { option: "salt" }
	});
	return c.getRandomValues(/* @__PURE__ */ new Uint8Array(16));
}
/**
* Computes the verifier that protects a sheet or workbook with `password` ([MS-OFFCRYPTO]
* §2.4.2.4, SR-50 … SR-53): the password as UTF-16LE, prefixed by the salt and hashed, then
* hashed again `spinCount` times with a 4-byte little-endian counter. Pass the result to
* `NewSheet.protect({ password })` or `NewWorkbook.protect({ password })`. Without a `salt` a
* random one is made, so each call gives a different (equally valid) hash.
*
* A protection password only asks spreadsheet applications to refuse changes; it is not
* encryption, and the file's content stays readable.
*
* @example
* ```ts
* import { createWorkbook, hashPassword, writeWorkbook } from 'xlvane';
*
* const wb = createWorkbook();
* wb.addSheet('Prices').set('A1', 10).protect({ password: hashPassword('secret') });
* const bytes = writeWorkbook(wb);
* ```
*/
function hashPassword(password, options = {}) {
	const p = checkPassword(password);
	if (typeof options !== "object" || options === null) throw invalid$4("The hashPassword options must be an object.", "For example { algorithm: 'SHA-512' }.", "options");
	for (const k of Object.keys(options)) if (k !== "algorithm" && k !== "spinCount" && k !== "salt") throw invalid$4(`hashPassword has no option "${k}".`, "Use algorithm, spinCount or salt.", k);
	const algorithm = options.algorithm ?? "SHA-512";
	if (!SHA_ALGORITHMS.includes(algorithm)) throw new XlvError("XLV_E_NOT_SUPPORTED", `xlvane cannot compute the hash algorithm ${JSON.stringify(String(algorithm).slice(0, 40))}.`, {
		hint: `Use one of ${SHA_ALGORITHMS.map((a) => `'${a}'`).join(", ")}; 'SHA-512' is what Excel writes.`,
		details: { option: "algorithm" }
	});
	const spinCount = options.spinCount ?? 1e5;
	if (typeof spinCount !== "number" || !Number.isInteger(spinCount) || spinCount < 0 || spinCount > MAX_SPIN) throw invalid$4("The option \"spinCount\" must be a whole number from 0 to 10000000.", "For example { spinCount: 100000 }, the value Excel writes.", "spinCount");
	let salt;
	if (options.salt === void 0) salt = randomSalt();
	else if (options.salt instanceof Uint8Array) salt = options.salt;
	else if (typeof options.salt === "string" && fromBase64(options.salt) !== null) salt = fromBase64(options.salt);
	else throw invalid$4("The option \"salt\" must be a Uint8Array or base64 text.", "For example { salt: 'c2FsdHNhbHRzYWx0c2FsdA==' }.", "salt");
	if (salt.length > MAX_SALT) throw invalid$4(`The salt has ${String(salt.length)} bytes; at most 65536 are allowed.`, "Use a salt of 16 bytes, as Excel does.", "salt");
	return Object.freeze({
		algorithmName: algorithm,
		hashValue: toBase64(isoHash(algorithm, salt, passwordBytes(p), spinCount)),
		saltValue: toBase64(salt),
		spinCount
	});
}
/**
* The legacy 16-bit verifier of Transitional files (`password` attributes; ECMA-376 Part 4
* §15.3.3, [MS-OFFCRYPTO] §2.3.7.1; SR-54), as 4 uppercase hexadecimal digits, high byte first
* as LibreOffice writes it (SR-55). Characters above U+00FF have no single-byte form here and
* count as `?` (0x3F), the replacement [MS-XLS] §4 describes; the real conversion uses the
* system's ANSI code page, which a file does not record. Weak by design: many passwords share a
* verifier. Use {@link hashPassword} for new files.
*
* @example
* ```ts
* import { legacyPasswordHash } from 'xlvane';
*
* console.log(legacyPasswordHash('secret')); // "DAA7"
* ```
*/
function legacyPasswordHash(password) {
	return legacyVerifier(password).toString(16).toUpperCase().padStart(4, "0");
}
function legacyVerifier(password) {
	if (typeof password !== "string") throw invalid$4("The password must be a string.", "For example legacyPasswordHash('secret').", "password");
	const bytes = [...password].map((ch) => {
		const c = ch.codePointAt(0);
		return c <= 255 ? c : 63;
	});
	if (bytes.length === 0) return 0;
	let h = 0;
	for (let i = bytes.length - 1; i >= 0; i--) {
		h = h >> 14 & 1 | h << 1 & 32767;
		h ^= bytes[i];
	}
	h = h >> 14 & 1 | h << 1 & 32767;
	h ^= bytes.length;
	return h ^ 52811;
}
function verifierOf(v) {
	if ("lockStructure" in v) {
		const w = v;
		return {
			algorithmName: w.workbookAlgorithmName,
			hashValue: w.workbookHashValue,
			saltValue: w.workbookSaltValue,
			spinCount: w.workbookSpinCount,
			password: w.workbookPassword
		};
	}
	return v;
}
/**
* Checks `password` against a stored verifier: a sheet's `protection`, the workbook's
* `protection` (its workbook password) or a {@link PasswordHash}. Strong verifiers are
* recomputed per [MS-OFFCRYPTO] §2.4.2.4; as [MS-XLSX] §2.6.56 and [MS-XLSB] §2.2.10 require,
* the variant that hashes the 16-bit legacy verifier as a single character is tried too
* (SR-56). A legacy `password` is compared as stored, in either byte order. Returns `false` when
* there is no verifier, or when the stored hash has the wrong length for its algorithm. A check
* helper only: xlvane does not need passwords to read or edit protected sheets. Cost: a strong
* verifier is hashed `spinCount + 1` times, twice for a wrong password (the two variants); a
* file may ask for up to 10,000,000 iterations, which takes minutes in plain JavaScript.
*
* @example
* ```ts
* import { hashPassword, verifyPassword } from 'xlvane';
*
* const hash = hashPassword('secret', { spinCount: 1000 });
* console.log(verifyPassword(hash, 'secret'), verifyPassword(hash, 'Secret')); // true false
* ```
*/
function verifyPassword(verifier, password) {
	if (typeof verifier !== "object" || verifier === null) throw invalid$4("verifyPassword needs a protection or a password hash.", "Pass sheet.protection, workbook.protection or the result of hashPassword.", "verifier");
	if (typeof password !== "string") throw invalid$4("The password must be a string.", "For example verifyPassword(sheet.protection, 'secret').", "password");
	const v = verifierOf(verifier);
	if (typeof v.hashValue === "string" && typeof v.algorithmName === "string") {
		const algorithm = SHA_ALGORITHMS.find((a) => a.toLowerCase() === v.algorithmName.toLowerCase().replace(/^sha(\d)/, "sha-$1"));
		if (algorithm === void 0) throw new XlvError("XLV_E_NOT_SUPPORTED", `xlvane cannot compute the hash algorithm ${JSON.stringify(v.algorithmName.slice(0, 40))}.`, {
			hint: `Only ${SHA_ALGORITHMS.join(", ")} are supported.`,
			details: { algorithmName: v.algorithmName.slice(0, 40) }
		});
		const salt = typeof v.saltValue === "string" ? fromBase64(v.saltValue) : /* @__PURE__ */ new Uint8Array(0);
		const expected = fromBase64(v.hashValue);
		const spin = v.spinCount ?? 0;
		const digest = {
			"SHA-1": 20,
			"SHA-256": 32,
			"SHA-384": 48,
			"SHA-512": 64
		}[algorithm];
		if (salt === null || expected === null || expected.length !== digest || !Number.isInteger(spin) || spin < 0 || spin > MAX_SPIN || salt.length > MAX_SALT) return false;
		const p = password.startsWith("﻿") ? password.slice(1) : password;
		const same = (h) => h.length === expected.length && h.every((b, i) => b === expected[i]);
		if (same(isoHash(algorithm, salt, passwordBytes(p), spin))) return true;
		const legacy = legacyVerifier(p);
		return same(isoHash(algorithm, salt, new Uint8Array([legacy & 255, legacy >> 8]), spin));
	}
	if (typeof v.password === "string" && /^[0-9a-f]{4}$/i.test(v.password)) {
		if (password.length === 0) return false;
		const stored = parseInt(v.password, 16);
		const computed = legacyVerifier(password);
		return stored === computed || stored === ((computed & 255) << 8 | computed >> 8);
	}
	return false;
}
// src/printing.ts
const PROTECTION_FLAGS = [
	"objects",
	"scenarios",
	"formatCells",
	"formatColumns",
	"formatRows",
	"insertColumns",
	"insertRows",
	"insertHyperlinks",
	"deleteColumns",
	"deleteRows",
	"selectLockedCells",
	"sort",
	"autoFilter",
	"pivotTables",
	"selectUnlockedCells"
];
/** @internal */
const DEFAULT_PROTECTION = Object.freeze({
	objects: false,
	scenarios: false,
	formatCells: true,
	formatColumns: true,
	formatRows: true,
	insertColumns: true,
	insertRows: true,
	insertHyperlinks: true,
	deleteColumns: true,
	deleteRows: true,
	selectLockedCells: false,
	sort: true,
	autoFilter: true,
	pivotTables: true,
	selectUnlockedCells: false
});
/** Margins of the Part 1 §18.3.1.62 example (L78045), used for values not given (SR-24). */
const EXAMPLE_MARGINS = Object.freeze({
	left: .7,
	right: .7,
	top: .75,
	bottom: .75,
	header: .3,
	footer: .3
});
/** @internal The page setup of a sheet without print elements. */
const DEFAULT_PAGE_SETUP = Object.freeze({
	margins: null,
	orientation: "default",
	paperSize: 1,
	scale: 100,
	fitToPage: false,
	fitToWidth: 1,
	fitToHeight: 1,
	pageOrder: "downThenOver",
	blackAndWhite: false,
	draft: false,
	cellComments: "none",
	errors: "displayed",
	firstPageNumber: null,
	horizontalDpi: null,
	verticalDpi: null,
	copies: 1,
	printOptions: Object.freeze({
		gridLines: false,
		headings: false,
		horizontalCentered: false,
		verticalCentered: false
	})
});
const DEFAULT_HEADER_FOOTER = Object.freeze({
	oddHeader: "",
	oddFooter: "",
	evenHeader: "",
	evenFooter: "",
	firstHeader: "",
	firstFooter: "",
	differentOddEven: false,
	differentFirst: false,
	scaleWithDoc: true,
	alignWithMargins: true
});
/** @internal */
const EMPTY_PRINT_DATA = Object.freeze({
	protection: null,
	pageSetup: DEFAULT_PAGE_SETUP,
	pageSetupSet: false,
	headerFooter: null,
	rowBreaks: [],
	colBreaks: [],
	printArea: null,
	printTitles: null
});
/** [MS-OI29500] §2.1.618 b, §2.1.619 c, §2.1.621 a, §2.1.622 a, §2.1.633 a, §2.1.634 a. */
const MAX_HEADER_FOOTER_LENGTH = 255;
/** [MS-OI29500] §2.1.599 a (colBreaks ≤ 1023) and §2.1.645 a (rowBreaks < 1023). */
const MAX_COL_BREAKS = 1023;
const MAX_ROW_BREAKS = 1022;
function invalid$3(message, hint, option) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		hint,
		details: { option }
	});
}
function readObject$1(value, known, what) {
	const out = /* @__PURE__ */ new Map();
	if (value === void 0) return out;
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw invalid$3(`${what} must be an object.`, `Pass an object with some of these properties: ${known.join(", ")}.`, what);
	for (const key of Object.keys(value)) {
		if (!known.includes(key)) throw invalid$3(`Unknown option "${key.slice(0, 40)}" in ${what}.`, `Known options: ${known.join(", ")}.`, key.slice(0, 40));
		const v = value[key];
		if (v !== void 0) out.set(key, v);
	}
	return out;
}
function bool$2(name, v) {
	if (typeof v !== "boolean") throw invalid$3(`The option "${name}" must be true or false.`, `For example { ${name}: true }.`, name);
	return v;
}
function integer$1(name, v, min, max, example) {
	if (typeof v !== "number" || !Number.isInteger(v) || v < min || v > max) throw invalid$3(`The option "${name}" must be a whole number from ${String(min)} to ${String(max)}.`, `For example { ${name}: ${String(example)} }.`, name);
	return v;
}
function oneOf$1(name, v, allowed) {
	if (typeof v === "string" && allowed.includes(v)) return v;
	throw invalid$3(`The option "${name}" has the value ${JSON.stringify(String(v).slice(0, 40))}.`, `Use ${allowed.map((a) => `"${a}"`).join(", ")}.`, name);
}
/** @internal Checks the options of `NewSheet.protect`. */
function normalizeProtection(options, sheet) {
	const o = readObject$1(options, [...PROTECTION_FLAGS, "password"], "protect options");
	const out = {
		...DEFAULT_PROTECTION,
		hash: null
	};
	for (const [k, v] of o) if (k === "password") out["hash"] = normalizePasswordHash(v, sheet);
	else out[k] = bool$2(k, v);
	return out;
}
/** @internal A password string (hashed with the defaults) or a checked `PasswordHash`. */
function normalizePasswordHash(value, sheet) {
	if (typeof value === "string") return hashPassword(value);
	const bad = (what) => new XlvError("XLV_E_INVALID_OPTION", `The option "password" ${what}.`, {
		...sheet === void 0 ? {} : { location: { sheet } },
		hint: "Pass the password as text, or the result of hashPassword('…').",
		details: { option: "password" }
	});
	if (typeof value !== "object" || value === null) throw bad("must be text or a password hash");
	const h = value;
	for (const k of Object.keys(value)) if (![
		"algorithmName",
		"hashValue",
		"saltValue",
		"spinCount"
	].includes(k)) throw bad(`has an unknown property "${k}"`);
	if (typeof h.algorithmName !== "string" || h.algorithmName.length === 0 || h.algorithmName.length > 255) throw bad("needs an algorithmName");
	if (typeof h.hashValue !== "string" || fromBase64(h.hashValue) === null || h.hashValue.length === 0) throw bad("needs a base64 hashValue");
	if (typeof h.saltValue !== "string" || fromBase64(h.saltValue) === null) throw bad("needs a base64 saltValue");
	if (typeof h.spinCount !== "number" || !Number.isInteger(h.spinCount) || h.spinCount < 0 || h.spinCount > 1e7) throw bad("needs a spinCount from 0 to 10000000");
	return Object.freeze({
		algorithmName: h.algorithmName,
		hashValue: h.hashValue,
		saltValue: h.saltValue,
		spinCount: h.spinCount
	});
}
/** @internal Checks the options of `NewWorkbook.protect`. */
function normalizeWorkbookProtection(options) {
	const o = readObject$1(options, [
		"lockStructure",
		"lockWindows",
		"password"
	], "workbook protect options");
	return {
		lockStructure: o.has("lockStructure") ? bool$2("lockStructure", o.get("lockStructure")) : true,
		lockWindows: o.has("lockWindows") ? bool$2("lockWindows", o.get("lockWindows")) : false,
		hash: o.has("password") ? normalizePasswordHash(o.get("password")) : null
	};
}
/** @internal `workbookProtection` (after `workbookPr`, before `bookViews`; sml.xsd T4097–T4104). */
function workbookProtectionXml(p) {
	if (p === null) return "";
	return `<workbookProtection${p.lockStructure ? " lockStructure=\"1\"" : ""}${p.lockWindows ? " lockWindows=\"1\"" : ""}${passwordHashAttrs(p.hash, "workbook")}/>`;
}
/** @internal The verifier attributes of a protection element, with an optional name prefix. */
function passwordHashAttrs(h, prefix = "") {
	if (h === null) return "";
	const n = (name) => prefix === "" ? name : prefix + name[0].toUpperCase() + name.slice(1);
	return `${a$1(n("algorithmName"), h.algorithmName)}${a$1(n("hashValue"), h.hashValue)}${a$1(n("saltValue"), h.saltValue)}${a$1(n("spinCount"), h.spinCount)}`;
}
const ORIENTATIONS = [
	"default",
	"portrait",
	"landscape"
];
const COMMENTS = [
	"none",
	"asDisplayed",
	"atEnd"
];
const ERRORS = [
	"displayed",
	"blank",
	"dash",
	"NA"
];
const ORDERS = ["downThenOver", "overThenDown"];
const SETUP_KEYS = [
	"margins",
	"orientation",
	"paperSize",
	"scale",
	"fitToPage",
	"fitToWidth",
	"fitToHeight",
	"pageOrder",
	"blackAndWhite",
	"draft",
	"cellComments",
	"errors",
	"firstPageNumber",
	"horizontalDpi",
	"verticalDpi",
	"copies",
	"printOptions"
];
/** @internal Applies the options of `NewSheet.pageSetup` to `base`. */
function normalizePageSetup(options, base) {
	const o = readObject$1(options, SETUP_KEYS, "pageSetup options");
	const next = { ...base };
	const margins = o.get("margins");
	if (margins !== void 0) {
		const m = readObject$1(margins, [
			"left",
			"right",
			"top",
			"bottom",
			"header",
			"footer"
		], "margins");
		const out = { ...base.margins ?? EXAMPLE_MARGINS };
		for (const [k, v] of m) {
			if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v >= 49) throw invalid$3(`The margin "${k}" must be a number of inches from 0 to below 49.`, `For example { margins: { ${k}: 0.75 } }.`, k);
			out[k] = v;
		}
		next.margins = out;
	}
	if (o.has("orientation")) next.orientation = oneOf$1("orientation", o.get("orientation"), ORIENTATIONS);
	if (o.has("paperSize")) next.paperSize = integer$1("paperSize", o.get("paperSize"), 1, 2147483647, 9);
	if (o.has("scale")) {
		const s = o.get("scale");
		if (typeof s !== "number" || !Number.isInteger(s) || s !== 0 && (s < 10 || s > 400)) throw invalid$3("The option \"scale\" must be a whole number from 10 to 400 (percent), or 0 for automatic.", "For example { scale: 80 }.", "scale");
		next.scale = s;
	}
	if (o.has("fitToPage")) next.fitToPage = bool$2("fitToPage", o.get("fitToPage"));
	if (o.has("fitToWidth")) next.fitToWidth = integer$1("fitToWidth", o.get("fitToWidth"), 0, 32767, 1);
	if (o.has("fitToHeight")) next.fitToHeight = integer$1("fitToHeight", o.get("fitToHeight"), 0, 32767, 1);
	if (o.has("pageOrder")) next.pageOrder = oneOf$1("pageOrder", o.get("pageOrder"), ORDERS);
	if (o.has("blackAndWhite")) next.blackAndWhite = bool$2("blackAndWhite", o.get("blackAndWhite"));
	if (o.has("draft")) next.draft = bool$2("draft", o.get("draft"));
	if (o.has("cellComments")) next.cellComments = oneOf$1("cellComments", o.get("cellComments"), COMMENTS);
	if (o.has("errors")) next.errors = oneOf$1("errors", o.get("errors"), ERRORS);
	if (o.has("firstPageNumber")) {
		const v = o.get("firstPageNumber");
		next.firstPageNumber = v === null ? null : integer$1("firstPageNumber", v, 0, 32767, 1);
	}
	for (const k of ["horizontalDpi", "verticalDpi"]) {
		if (!o.has(k)) continue;
		const v = o.get(k);
		next[k] = v === null ? null : integer$1(k, v, 1, 4294967295, 600);
	}
	if (o.has("copies")) next.copies = integer$1("copies", o.get("copies"), 1, 32767, 2);
	const po = o.get("printOptions");
	if (po !== void 0) {
		const p = readObject$1(po, [
			"gridLines",
			"headings",
			"horizontalCentered",
			"verticalCentered"
		], "printOptions");
		const out = { ...base.printOptions };
		for (const [k, v] of p) out[k] = bool$2(k, v);
		next.printOptions = out;
	}
	return next;
}
/** @internal Applies the options of `NewSheet.headerFooter` to `base` (or the defaults). */
function normalizeHeaderFooter(options, base) {
	const texts = [
		"oddHeader",
		"oddFooter",
		"evenHeader",
		"evenFooter",
		"firstHeader",
		"firstFooter"
	];
	const flags = [
		"differentOddEven",
		"differentFirst",
		"scaleWithDoc",
		"alignWithMargins"
	];
	const o = readObject$1(options, [...texts, ...flags], "headerFooter options");
	const out = { ...base ?? DEFAULT_HEADER_FOOTER };
	for (const k of texts) {
		if (!o.has(k)) continue;
		const v = o.get(k);
		if (typeof v !== "string" || firstLoneSurrogate(v) !== -1) throw invalid$3(`The option "${k}" must be text.`, `For example { ${k}: '&CPage &P' }.`, k);
		if (v.length > MAX_HEADER_FOOTER_LENGTH) throw invalid$3(`The option "${k}" has ${String(v.length)} characters; Office allows at most ${String(MAX_HEADER_FOOTER_LENGTH)}.`, "Shorten the text.", k);
		out[k] = v;
	}
	for (const k of flags) if (o.has(k)) out[k] = bool$2(k, o.get(k));
	return out;
}
/** @internal Checks the options of `NewSheet.pageBreaks`; returns sorted, unique row and column numbers. */
function normalizePageBreaks(options) {
	const o = readObject$1(options, ["rows", "columns"], "pageBreaks options");
	const list = (key) => {
		const v = o.get(key) ?? [];
		if (!Array.isArray(v)) throw invalid$3(`The option "${key}" must be an array.`, key === "rows" ? "For example { rows: [25, 49] }." : "For example { columns: [\"F\", 12] }.", key);
		return v;
	};
	const rows = [...new Set(list("rows").map((r) => integer$1("rows", r, 2, MAX_ROWS, 25)))].sort((a, b) => a - b);
	const cols = [...new Set(list("columns").map((c) => {
		const n = typeof c === "string" ? columnNumber(c.toUpperCase()) : c;
		if (typeof n !== "number" || !Number.isInteger(n) || n < 2 || n > 16384) throw invalid$3(`The page break column ${JSON.stringify(String(c).slice(0, 20))} is not valid.`, "Use a column from \"B\" to \"XFD\" (or 2 to 16384); the column starts a new page.", "columns");
		return n;
	}))].sort((a, b) => a - b);
	if (rows.length > MAX_ROW_BREAKS) throw invalid$3(`A sheet can have at most ${String(MAX_ROW_BREAKS)} row page breaks; ${String(rows.length)} were given.`, "Remove some breaks.", "rows");
	if (cols.length > MAX_COL_BREAKS) throw invalid$3(`A sheet can have at most ${String(MAX_COL_BREAKS)} column page breaks; ${String(cols.length)} were given.`, "Remove some breaks.", "columns");
	return {
		rows,
		cols
	};
}
/** @internal Parses the ranges of `NewSheet.printArea`. */
function normalizePrintArea(value) {
	if (value === null) return null;
	const list = typeof value === "string" ? [value] : value;
	if (!Array.isArray(list) || list.length === 0) throw invalid$3("The print area must be a range or a non-empty array of ranges.", "For example printArea(\"A1:H40\") or printArea([\"A1:D20\", \"F1:H20\"]).", "printArea");
	return list.map((r) => {
		const range = typeof r === "string" ? parseRange(r.replace(/\$/g, "").toUpperCase()) : null;
		if (range === null) throw invalid$3(`The print area range ${JSON.stringify(String(r).slice(0, 40))} is not valid.`, "Use A1-style ranges such as \"A1:H40\".", "printArea");
		return range;
	});
}
function span(value, kind) {
	if (value === null || value === void 0) return null;
	const parse = (s) => {
		if (kind === "rows") {
			if (!/^\d{1,7}$/.test(s)) return null;
			const n = Number(s);
			return n >= 1 && n <= 1048576 ? n : null;
		}
		return columnNumber(s.toUpperCase());
	};
	const parts = typeof value === "string" ? value.replace(/\$/g, "").split(":") : [];
	const a = parts.length >= 1 && parts.length <= 2 ? parse(parts[0]) : null;
	const b = parts.length === 2 ? parse(parts[1]) : a;
	if (a === null || b === null) throw invalid$3(`The print titles ${kind} ${JSON.stringify(String(value).slice(0, 40))} are not valid.`, kind === "rows" ? "Use rows such as \"1\" or \"1:2\"." : "Use columns such as \"A\" or \"A:B\".", kind);
	return [Math.min(a, b), Math.max(a, b)];
}
/** @internal Parses the options of `NewSheet.printTitles`. */
function normalizePrintTitles(value) {
	if (value === null) return null;
	const o = readObject$1(value, ["rows", "columns"], "printTitles options");
	const rows = span(o.get("rows"), "rows");
	const cols = span(o.get("columns"), "columns");
	return rows === null && cols === null ? null : {
		rows,
		cols
	};
}
function a$1(name, value) {
	return ` ${name}="${typeof value === "number" ? String(value) : escapeAttr(value)}"`;
}
/** @internal `sheetProtection` with `sheet="1"` and the flags that differ from the schema defaults. */
function sheetProtectionXml(p) {
	if (p === null) return "";
	let xml = `<sheetProtection${passwordHashAttrs(p.hash)} sheet="1"`;
	for (const k of PROTECTION_FLAGS) if (p[k] !== DEFAULT_PROTECTION[k]) xml += a$1(k, p[k] ? 1 : 0);
	return `${xml}/>`;
}
/**
* @internal The elements after `hyperlinks` in CT_Worksheet order (sml.xsd T2192–T2197, SR-23):
* printOptions, pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks.
*/
function printXml(d) {
	let xml = "";
	const s = d.pageSetup;
	if (d.pageSetupSet) {
		const po = s.printOptions;
		if (po.horizontalCentered || po.verticalCentered || po.headings || po.gridLines) {
			xml += "<printOptions";
			if (po.horizontalCentered) xml += " horizontalCentered=\"1\"";
			if (po.verticalCentered) xml += " verticalCentered=\"1\"";
			if (po.headings) xml += " headings=\"1\"";
			if (po.gridLines) xml += " gridLines=\"1\"";
			xml += "/>";
		}
		if (s.margins !== null) {
			const m = s.margins;
			xml += `<pageMargins${a$1("left", m.left)}${a$1("right", m.right)}${a$1("top", m.top)}${a$1("bottom", m.bottom)}${a$1("header", m.header)}${a$1("footer", m.footer)}/>`;
		}
		let ps = "";
		if (s.paperSize !== 1) ps += a$1("paperSize", s.paperSize);
		if (s.scale !== 100) ps += a$1("scale", s.scale);
		if (s.firstPageNumber !== null) ps += a$1("firstPageNumber", s.firstPageNumber);
		if (s.fitToWidth !== 1) ps += a$1("fitToWidth", s.fitToWidth);
		if (s.fitToHeight !== 1) ps += a$1("fitToHeight", s.fitToHeight);
		if (s.pageOrder !== "downThenOver") ps += a$1("pageOrder", s.pageOrder);
		if (s.orientation !== "default") ps += a$1("orientation", s.orientation);
		if (s.blackAndWhite) ps += " blackAndWhite=\"1\"";
		if (s.draft) ps += " draft=\"1\"";
		if (s.cellComments !== "none") ps += a$1("cellComments", s.cellComments);
		if (s.firstPageNumber !== null) ps += " useFirstPageNumber=\"1\"";
		if (s.errors !== "displayed") ps += a$1("errors", s.errors);
		if (s.horizontalDpi !== null) ps += a$1("horizontalDpi", s.horizontalDpi);
		if (s.verticalDpi !== null) ps += a$1("verticalDpi", s.verticalDpi);
		if (s.copies !== 1) ps += a$1("copies", s.copies);
		if (ps !== "") xml += `<pageSetup${ps}/>`;
	}
	const hf = d.headerFooter;
	if (hf !== null) {
		let attrs = "";
		if (hf.differentOddEven) attrs += " differentOddEven=\"1\"";
		if (hf.differentFirst) attrs += " differentFirst=\"1\"";
		if (!hf.scaleWithDoc) attrs += " scaleWithDoc=\"0\"";
		if (!hf.alignWithMargins) attrs += " alignWithMargins=\"0\"";
		let body = "";
		for (const k of [
			"oddHeader",
			"oddFooter",
			"evenHeader",
			"evenFooter",
			"firstHeader",
			"firstFooter"
		]) if (hf[k] !== "") body += `<${k}>${escapeText(encodeXstring(hf[k]))}</${k}>`;
		if (attrs !== "" || body !== "") xml += `<headerFooter${attrs}>${body}</headerFooter>`;
	}
	const breaks = (element, list, max) => list.length === 0 ? "" : `<${element}${a$1("count", list.length)}${a$1("manualBreakCount", list.length)}>${list.map((n) => `<brk${a$1("id", n - 1)}${a$1("max", max)} man="1"/>`).join("")}</${element}>`;
	xml += breaks("rowBreaks", d.rowBreaks, 16383);
	xml += breaks("colBreaks", d.colBreaks, 1048575);
	return xml;
}
function quoteSheet(name) {
	return `'${name.replace(/'/g, "''")}'`;
}
function abs(row, col) {
	return `$${columnLetters(col)}$${String(row)}`;
}
/**
* @internal The `_xlnm.Print_Area` and `_xlnm.Print_Titles` names of one sheet (SR-30). Their
* text follows the files written by Excel in the Apache POI corpus (no document defines it):
* `'Sheet'!$A$1:$J$20` (several areas comma-separated) and `'Sheet'!$A:$B,'Sheet'!$1:$2`.
*/
function printNames(sheet, d) {
	const out = [];
	const q = quoteSheet(sheet);
	if (d.printArea !== null) {
		const text = d.printArea.map((r) => {
			const one = r.start.row === r.end.row && r.start.col === r.end.col;
			return `${q}!${abs(r.start.row, r.start.col)}${one ? "" : `:${abs(r.end.row, r.end.col)}`}`;
		}).join(",");
		out.push({
			name: "_xlnm.Print_Area",
			refersTo: text,
			sheet,
			hidden: false,
			comment: null
		});
	}
	const t = d.printTitles;
	if (t !== null) {
		const parts = [];
		if (t.cols !== null) parts.push(`${q}!$${columnLetters(t.cols[0])}:$${columnLetters(t.cols[1])}`);
		if (t.rows !== null) parts.push(`${q}!$${String(t.rows[0])}:$${String(t.rows[1])}`);
		out.push({
			name: "_xlnm.Print_Titles",
			refersTo: parts.join(","),
			sheet,
			hidden: false,
			comment: null
		});
	}
	return out;
}
/** Splits formula text at commas outside quotes. */
function splitList(text) {
	const out = [];
	let start = 0;
	let quoted = false;
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (c === "'") quoted = !quoted;
		else if (c === "," && !quoted) {
			out.push(text.slice(start, i));
			start = i + 1;
		}
	}
	out.push(text.slice(start));
	return out;
}
/** Removes the sheet prefix of `part` if it names `sheet`; `null` if it names another sheet or none. */
function withoutSheet(part, sheet) {
	const bang = part.lastIndexOf("!");
	if (bang < 0) return null;
	let prefix = part.slice(0, bang).trim();
	if (prefix.startsWith("'") && prefix.endsWith("'") && prefix.length >= 2) prefix = prefix.slice(1, -1).replace(/''/g, "'");
	return prefix.toLowerCase() === sheet.toLowerCase() ? part.slice(bang + 1).trim().replace(/\$/g, "") : null;
}
/** The ranges of a `_xlnm.Print_Area` name of `sheet`, or `null` if its text is not a list of plain ranges. */
function parsePrintArea(text, sheet) {
	const out = [];
	for (const part of splitList(text)) {
		const ref = withoutSheet(part, sheet);
		if ((ref === null ? null : parseRange(ref.toUpperCase())) === null || ref === null) return null;
		out.push(ref.toUpperCase());
	}
	return out;
}
/** The rows and columns of a `_xlnm.Print_Titles` name of `sheet`, or `null` if its text has another form. */
function parsePrintTitles(text, sheet) {
	let rows = null;
	let columns = null;
	for (const part of splitList(text)) {
		const ref = withoutSheet(part, sheet);
		if (ref === null) return null;
		if (/^\d{1,7}:\d{1,7}$/.test(ref)) rows = ref;
		else if (/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(ref)) columns = ref.toUpperCase();
		else return null;
	}
	return rows === null && columns === null ? null : {
		rows,
		columns
	};
}
Object.freeze({
	protection: null,
	pageSetup: DEFAULT_PAGE_SETUP,
	headerFooter: null,
	rowBreaks: [],
	colBreaks: []
});
const HF_TEXTS = /* @__PURE__ */ new Set([
	"oddHeader",
	"oddFooter",
	"evenHeader",
	"evenFooter",
	"firstHeader",
	"firstFooter"
]);
/** Breaks read per collection; the file's own limit is about a thousand ([MS-OI29500] §2.1.599). */
const MAX_READ_BREAKS = 65536;
/**
* @internal Collects `sheetProtection`, `printOptions`, `pageMargins`, `pageSetup`,
* `sheetPr/pageSetUpPr`, `headerFooter`, `rowBreaks` and `colBreaks` of a worksheet while it is
* parsed. `start` returns `true` for elements it handled. Invalid values warn and keep the default.
*/
function createPrintCollector(warn) {
	let protection = null;
	const setup = { ...DEFAULT_PAGE_SETUP };
	let headerFooter = null;
	let textTarget = null;
	let text = "";
	const rowBreaks = [];
	const colBreaks = [];
	let breaks = null;
	const boolAttr = (attrs, name, fallback, what) => {
		const v = attrs.get(name);
		if (v === void 0) return fallback;
		if (v === "1" || v === "true") return true;
		if (v === "0" || v === "false") return false;
		warn(`The ${what} attribute ${name}="${v.slice(0, 20)}" is not a boolean; the default is used.`);
		return fallback;
	};
	const uintAttr = (attrs, name, what) => {
		const v = attrs.get(name);
		if (v === void 0) return null;
		if (/^\d{1,10}$/.test(v.trim()) && Number(v) <= 4294967295) return Number(v);
		warn(`The ${what} attribute ${name}="${v.slice(0, 20)}" is not a whole number; the default is used.`);
		return null;
	};
	const textAttr = (attrs, name) => {
		const v = attrs.get(name);
		return v === void 0 ? null : v;
	};
	const enumAttr = (attrs, name, allowed, fallback) => {
		const v = attrs.get(name);
		if (v === void 0) return fallback;
		if (allowed.includes(v)) return v;
		warn(`The page setup attribute ${name}="${v.slice(0, 30)}" is not known; the default is used.`);
		return fallback;
	};
	return {
		start(depth, local, parent, attrs) {
			if (depth === 2 && local === "sheetProtection") {
				const flags = {};
				for (const k of PROTECTION_FLAGS) flags[k] = boolAttr(attrs, k, DEFAULT_PROTECTION[k], "sheet protection");
				protection = Object.freeze({
					...flags,
					sheet: boolAttr(attrs, "sheet", false, "sheet protection"),
					algorithmName: textAttr(attrs, "algorithmName"),
					hashValue: textAttr(attrs, "hashValue"),
					saltValue: textAttr(attrs, "saltValue"),
					spinCount: uintAttr(attrs, "spinCount", "sheet protection"),
					password: textAttr(attrs, "password")
				});
				return true;
			}
			if (depth === 2 && local === "printOptions") {
				setup.printOptions = Object.freeze({
					gridLines: boolAttr(attrs, "gridLines", false, "print options") && boolAttr(attrs, "gridLinesSet", true, "print options"),
					headings: boolAttr(attrs, "headings", false, "print options"),
					horizontalCentered: boolAttr(attrs, "horizontalCentered", false, "print options"),
					verticalCentered: boolAttr(attrs, "verticalCentered", false, "print options")
				});
				return true;
			}
			if (depth === 2 && local === "pageMargins") {
				const m = {};
				for (const k of [
					"left",
					"right",
					"top",
					"bottom",
					"header",
					"footer"
				]) {
					const v = attrs.get(k);
					const n = v === void 0 ? NaN : Number(v);
					if (v === void 0 || v.trim() === "" || !Number.isFinite(n) || n < 0) {
						warn(`The page margin ${k}="${(v ?? "").slice(0, 20)}" is missing or not a valid number; the margins are ignored.`);
						return true;
					}
					m[k] = n;
				}
				setup.margins = Object.freeze(m);
				return true;
			}
			if (depth === 2 && local === "pageSetup") {
				const what = "page setup";
				setup.paperSize = uintAttr(attrs, "paperSize", what) ?? 1;
				setup.scale = uintAttr(attrs, "scale", what) ?? 100;
				setup.fitToWidth = uintAttr(attrs, "fitToWidth", what) ?? 1;
				setup.fitToHeight = uintAttr(attrs, "fitToHeight", what) ?? 1;
				setup.pageOrder = enumAttr(attrs, "pageOrder", ORDERS, "downThenOver");
				setup.orientation = enumAttr(attrs, "orientation", ORIENTATIONS, "default");
				setup.blackAndWhite = boolAttr(attrs, "blackAndWhite", false, what);
				setup.draft = boolAttr(attrs, "draft", false, what);
				setup.cellComments = enumAttr(attrs, "cellComments", COMMENTS, "none");
				setup.errors = enumAttr(attrs, "errors", ERRORS, "displayed");
				setup.firstPageNumber = boolAttr(attrs, "useFirstPageNumber", false, what) ? uintAttr(attrs, "firstPageNumber", what) ?? 1 : null;
				setup.horizontalDpi = uintAttr(attrs, "horizontalDpi", what);
				setup.verticalDpi = uintAttr(attrs, "verticalDpi", what);
				setup.copies = uintAttr(attrs, "copies", what) ?? 1;
				return true;
			}
			if (depth === 3 && local === "pageSetUpPr" && parent === "sheetPr") {
				setup.fitToPage = boolAttr(attrs, "fitToPage", false, "page setup");
				return true;
			}
			if (depth === 2 && local === "headerFooter") {
				headerFooter = {
					...DEFAULT_HEADER_FOOTER,
					differentOddEven: boolAttr(attrs, "differentOddEven", false, "header/footer"),
					differentFirst: boolAttr(attrs, "differentFirst", false, "header/footer"),
					scaleWithDoc: boolAttr(attrs, "scaleWithDoc", true, "header/footer"),
					alignWithMargins: boolAttr(attrs, "alignWithMargins", true, "header/footer")
				};
				return true;
			}
			if (depth === 3 && parent === "headerFooter" && HF_TEXTS.has(local)) {
				textTarget = local;
				text = "";
				return true;
			}
			if (depth === 2 && (local === "rowBreaks" || local === "colBreaks")) {
				breaks = local === "rowBreaks" ? rowBreaks : colBreaks;
				return true;
			}
			if (depth === 3 && local === "brk" && (parent === "rowBreaks" || parent === "colBreaks") && breaks !== null) {
				const id = uintAttr(attrs, "id", "page break") ?? 0;
				const limit = parent === "rowBreaks" ? MAX_ROWS : MAX_COLUMNS;
				if (id + 1 > limit) warn(`The page break id="${String(id)}" is beyond the sheet; it is ignored.`);
				else if (breaks.length >= MAX_READ_BREAKS) warn(`The sheet has more than ${String(MAX_READ_BREAKS)} page breaks; the rest are ignored.`);
				else breaks.push(id + 1);
				return true;
			}
			return false;
		},
		end(local) {
			if (textTarget !== null && local === textTarget) {
				if (headerFooter !== null) headerFooter[textTarget] = decodeXstring(text);
				textTarget = null;
				text = "";
			} else if (local === "rowBreaks" || local === "colBreaks") breaks = null;
		},
		text(value) {
			if (textTarget === null) return false;
			if (text.length < 1048576) text += value;
			return true;
		},
		result() {
			const unique = (list) => [...new Set(list)].sort((x, y) => x - y);
			return {
				protection,
				pageSetup: Object.freeze({ ...setup }),
				headerFooter: headerFooter === null ? null : Object.freeze(headerFooter),
				rowBreaks: unique(rowBreaks),
				colBreaks: unique(colBreaks)
			};
		}
	};
}
/** @internal Reads the `workbookProtection` attributes (SR-22). */
function workbookProtectionOf(attrs, warn) {
	const b = (name) => {
		const v = attrs.get(name);
		if (v === void 0) return false;
		if (v === "1" || v === "true") return true;
		if (v !== "0" && v !== "false") warn(`The workbook protection attribute ${name}="${v.slice(0, 20)}" is not a boolean; false is used.`);
		return false;
	};
	const t = (name) => attrs.get(name) ?? null;
	const n = (name) => {
		const v = attrs.get(name);
		return v !== void 0 && /^\d{1,10}$/.test(v.trim()) && Number(v) <= 4294967295 ? Number(v) : null;
	};
	return Object.freeze({
		lockStructure: b("lockStructure"),
		lockWindows: b("lockWindows"),
		lockRevision: b("lockRevision"),
		workbookAlgorithmName: t("workbookAlgorithmName"),
		workbookHashValue: t("workbookHashValue"),
		workbookSaltValue: t("workbookSaltValue"),
		workbookSpinCount: n("workbookSpinCount"),
		revisionsAlgorithmName: t("revisionsAlgorithmName"),
		revisionsHashValue: t("revisionsHashValue"),
		revisionsSaltValue: t("revisionsSaltValue"),
		revisionsSpinCount: n("revisionsSpinCount"),
		workbookPassword: t("workbookPassword"),
		revisionsPassword: t("revisionsPassword")
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
* The default {@link ReadLimits}. Frozen; pass `{ limits: { ... } }` to change individual values.
*
* @example
* ```ts
* import { DEFAULT_READ_LIMITS } from 'xlvane';
* console.log(DEFAULT_READ_LIMITS.maxEntries); // 10000
* ```
*/
const DEFAULT_READ_LIMITS = Object.freeze({
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
	maxCells: 5e6,
	maxSharedStringChars: 512 * MiB
});
const LIMIT_NAMES = Object.keys(DEFAULT_READ_LIMITS);
/**
* Validates user-supplied limits and merges them over the defaults. Only own, known keys are
* read, so keys such as `__proto__` cannot change anything; unknown keys are rejected.
*/
function resolveLimits(input) {
	const limits = { ...DEFAULT_READ_LIMITS };
	if (input === void 0) return limits;
	if (input === null || typeof input !== "object" || Array.isArray(input)) throw new XlvError("XLV_E_INVALID_OPTION", `Option "limits" must be an object, got ${describe$1(input)}.`, { hint: "Pass an object such as { limits: { maxEntries: 10000 } }, or leave \"limits\" out to use the defaults." });
	for (const key of Object.keys(input)) if (!LIMIT_NAMES.includes(key)) {
		const close = suggest(key, LIMIT_NAMES);
		throw new XlvError("XLV_E_INVALID_OPTION", `Unknown limit "${key}".`, {
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
		limits[name] = validate(name, value);
	}
	return limits;
}
function validate(name, value) {
	const isRatio = name === "maxCompressionRatio";
	if (!(typeof value === "number" && (value === Number.POSITIVE_INFINITY || (isRatio ? Number.isFinite(value) && value >= 1 : Number.isSafeInteger(value) && value > 0)))) throw new XlvError("XLV_E_INVALID_OPTION", isRatio ? `Limit "${name}" must be a number of at least 1 or Infinity, got ${describe$1(value)}.` : `Limit "${name}" must be a whole number greater than 0 or Infinity, got ${describe$1(value)}.`, {
		hint: `Pass for example { limits: { ${name}: ${String(DEFAULT_READ_LIMITS[name])} } }, or leave it out to use the default.`,
		details: { option: name }
	});
	return value;
}
function describe$1(value) {
	if (value === null) return "null";
	if (Array.isArray(value)) return "an array";
	if (typeof value === "string") return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
	if (typeof value === "object" || typeof value === "function") return `a value of type ${typeof value}`;
	return String(value);
}
/**
* Checks the number of positions an output built from a sheet would have against
* max(`maxCells`, OUTPUT_PER_CELL × cells read). Throws `XLV_E_LIMIT_EXCEEDED` above it.
* `hint` replaces the default hint, which suits outputs read from a file.
*/
function checkOutputPositions(positions, cells, maxCells, what, sheet, hint) {
	const budget = Math.max(maxCells, 16 * cells);
	if (positions <= budget) return;
	throw new XlvError("XLV_E_LIMIT_EXCEEDED", `${what} of sheet "${sheet}" would have ${String(positions)} positions (rows × columns) for ${String(cells)} cells, more than the limit of ${String(budget)}.`, {
		location: { sheet },
		hint: hint ?? `The cells are spread over a far larger area than they fill. Pass a smaller range (for example { range: "A1:H1000" } or --range), or raise limits.maxCells if the file is trusted.`,
		details: {
			limit: "maxCells",
			value: budget
		}
	});
}
/**
* Every row object repeats the keys of all columns, so a long header makes a small file give
* gigabytes of keys (and JSON.stringify of the rows fail). The key text of all rows is bounded
* by `limits.maxSharedStringChars`, the bound on text that cells repeat (D-H01).
*/
function checkKeyText(columns, rows, limit, sheet) {
	let perRow = 0;
	for (const c of columns) perRow += c.key.length;
	if (perRow * rows <= limit) return;
	throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The rows of the table of sheet "${sheet}" would repeat ${String(perRow)} characters of column names in each of ${String(rows)} rows, more than the limit of ${String(limit)} characters.`, {
		location: { sheet },
		hint: `Each row object holds the text of every column name. Read fewer columns or rows with the range option (for example { range: "A1:H1000" } or --range), or raise limits.maxSharedStringChars if the file is trusted.`,
		details: {
			limit: "maxSharedStringChars",
			value: limit
		}
	});
}
// src/sml/numfmt.ts
/** A `Map` whose contents cannot change after construction. */
var FrozenMap = class extends Map {
	constructor(entries) {
		super();
		for (const [key, value] of entries) super.set(key, value);
		Object.freeze(this);
	}
	set(_key, _value) {
		throw new TypeError("This map is read-only.");
	}
	delete(_key) {
		throw new TypeError("This map is read-only.");
	}
	clear() {
		throw new TypeError("This map is read-only.");
	}
};
/**
* Number formats implied by their `numFmtId` without a `numFmt` element: the "All Languages"
* table of ECMA-376 Part 1 §18.8.30, verbatim (ids 37 and 38 have a space before ";" in the
* spec). Ids not listed, such as 5–8, have no implied format and need an explicit `numFmt`.
* The locale-dependent ids (27–36, 50–81) have different codes per language and are not
* included; see {@link isDateFormatId}.
*/
const BUILTIN_NUMBER_FORMATS = new FrozenMap([
	[0, "General"],
	[1, "0"],
	[2, "0.00"],
	[3, "#,##0"],
	[4, "#,##0.00"],
	[9, "0%"],
	[10, "0.00%"],
	[11, "0.00E+00"],
	[12, "# ?/?"],
	[13, "# ??/??"],
	[14, "mm-dd-yy"],
	[15, "d-mmm-yy"],
	[16, "d-mmm"],
	[17, "mmm-yy"],
	[18, "h:mm AM/PM"],
	[19, "h:mm:ss AM/PM"],
	[20, "h:mm"],
	[21, "h:mm:ss"],
	[22, "m/d/yy h:mm"],
	[37, "#,##0 ;(#,##0)"],
	[38, "#,##0 ;[Red](#,##0)"],
	[39, "#,##0.00;(#,##0.00)"],
	[40, "#,##0.00;[Red](#,##0.00)"],
	[45, "mm:ss"],
	[46, "[h]:mm:ss"],
	[47, "mmss.0"],
	[48, "##0.0E+0"],
	[49, "@"]
]);
/** Built-in ids whose implied formats are dates or times (ECMA-376 Part 1 §18.8.30). */
const DATE_FORMAT_IDS = /* @__PURE__ */ new Set([
	14,
	15,
	16,
	17,
	18,
	19,
	20,
	21,
	22,
	45,
	46,
	47,
	27,
	28,
	29,
	30,
	31,
	32,
	33,
	34,
	35,
	36,
	50,
	51,
	52,
	53,
	54,
	55,
	56,
	57,
	58,
	71,
	72,
	73,
	74,
	75,
	76,
	77,
	78,
	79,
	80,
	81
]);
/**
* Returns `true` if `id` is a built-in number format id whose implied format displays a date or
* time, including the locale-dependent date ids 27–36 and 50–58 (Chinese, Japanese, Korean) and
* 71–81 (Thai) listed in ECMA-376 Part 1 §18.8.30. Use it for an `xf@numFmtId` that has no
* `numFmt` element; use {@link isDateFormat} for explicit format codes.
*/
function isDateFormatId(id) {
	return DATE_FORMAT_IDS.has(id);
}
const CODE_QUOTE = 34;
const CODE_DOLLAR = 36;
const CODE_ASTERISK = 42;
const CODE_PLUS = 43;
const CODE_HYPHEN$1 = 45;
const CODE_LEFT_BRACKET = 91;
const CODE_BACKSLASH = 92;
const CODE_UNDERSCORE = 95;
const CODE_UPPER_G = 71;
const CODE_LOWER_D = 100;
const CODE_LOWER_E = 101;
const CODE_LOWER_G = 103;
const CODE_LOWER_H = 104;
const CODE_LOWER_M = 109;
const CODE_LOWER_R = 114;
const CODE_LOWER_S = 115;
const CODE_LOWER_Y = 121;
/** Language-info values meaning "system long date format" and "system time format". */
const LANGUAGE_SYSTEM_LONG_DATE = 63488;
const LANGUAGE_SYSTEM_TIME = 62464;
const MAX_LANGUAGE_INFO_DIGITS = 8;
function hexValue(code) {
	if (code >= 48 && code <= 57) return code - 48;
	if (code >= 65 && code <= 70) return code - 65 + 10;
	if (code >= 97 && code <= 102) return code - 97 + 10;
	return -1;
}
/** True if the bracket content `code[from..to)` makes the format display a date or time. */
function isDateTimeBracket(code, from, to) {
	if (from >= to) return false;
	const first = code.charCodeAt(from);
	if (first === CODE_LOWER_H || first === CODE_LOWER_M || first === CODE_LOWER_S) {
		for (let i = from + 1; i < to; i++) if (code.charCodeAt(i) !== first) return false;
		return true;
	}
	if (first === CODE_DOLLAR) {
		let hyphen = -1;
		for (let i = to - 1; i > from; i--) if (code.charCodeAt(i) === CODE_HYPHEN$1) {
			hyphen = i;
			break;
		}
		const digits = to - hyphen - 1;
		if (hyphen === -1 || digits < 1 || digits > MAX_LANGUAGE_INFO_DIGITS) return false;
		let value = 0;
		for (let i = hyphen + 1; i < to; i++) {
			const digit = hexValue(code.charCodeAt(i));
			if (digit < 0) return false;
			value = value * 16 + digit;
		}
		const languageId = value % 65536;
		return languageId === LANGUAGE_SYSTEM_LONG_DATE || languageId === LANGUAGE_SYSTEM_TIME;
	}
	return false;
}
/** True if `code` has the keyword "General" (any case) at `at`. */
function isGeneralAt(code, at) {
	return at + 7 <= code.length && code.slice(at, at + 7).toLowerCase() === "general";
}
/**
* Returns `true` if the number format code displays a date or time in any of its sections,
* following the format-code grammar of ECMA-376 Part 1 §18.8.31. Date and time codes are `y`,
* `m`, `d`, `h`, `s`, the elapsed-time brackets `[h]`, `[mm]`, `[ss]`, the era codes `e`, `g`,
* `r`, and the language-info values for the system date and time formats (`[$-F800]`,
* `[$-F400]`); they are lowercase in the spec and matched case-sensitively. Text in double
* quotes, a character after `\`, `_` or `*`, colour, condition and locale brackets, the keyword
* `General` and the exponent codes `E+`, `E-`, `e+`, `e-` are not dates. An unterminated quote
* or bracket runs to the end of the code. Runs in linear time.
*/
function isDateFormat(code) {
	const length = code.length;
	let i = 0;
	while (i < length) switch (code.charCodeAt(i)) {
		case CODE_QUOTE: {
			const end = code.indexOf("\"", i + 1);
			if (end === -1) return false;
			i = end + 1;
			break;
		}
		case CODE_BACKSLASH:
		case CODE_UNDERSCORE:
		case CODE_ASTERISK:
			i += 2;
			break;
		case CODE_LEFT_BRACKET: {
			const end = code.indexOf("]", i + 1);
			if (end === -1) return false;
			if (isDateTimeBracket(code, i + 1, end)) return true;
			i = end + 1;
			break;
		}
		case CODE_UPPER_G:
			i += isGeneralAt(code, i) ? 7 : 1;
			break;
		case CODE_LOWER_G:
			if (isGeneralAt(code, i)) {
				i += 7;
				break;
			}
			return true;
		case CODE_LOWER_E: {
			const next = code.charCodeAt(i + 1);
			if (next === CODE_PLUS || next === CODE_HYPHEN$1) {
				i += 2;
				break;
			}
			return true;
		}
		case CODE_LOWER_Y:
		case CODE_LOWER_M:
		case CODE_LOWER_D:
		case CODE_LOWER_H:
		case CODE_LOWER_S:
		case CODE_LOWER_R: return true;
		default: i++;
	}
	return false;
}
// src/style.ts
const UNDERLINES = [
	"single",
	"double",
	"singleAccounting",
	"doubleAccounting"
];
const PATTERNS = [
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625"
];
const BORDER_STYLES = [
	"thin",
	"medium",
	"dashed",
	"dotted",
	"thick",
	"double",
	"hair",
	"mediumDashed",
	"dashDot",
	"mediumDashDot",
	"dashDotDot",
	"mediumDashDotDot",
	"slantDashDot"
];
const HORIZONTAL = [
	"general",
	"left",
	"center",
	"right",
	"fill",
	"justify",
	"centerContinuous",
	"distributed"
];
const VERTICAL = [
	"top",
	"center",
	"bottom",
	"justify",
	"distributed"
];
const EDGES = [
	"left",
	"right",
	"top",
	"bottom"
];
const VERT_ALIGN_RUNS = [
	"baseline",
	"superscript",
	"subscript"
];
const FONT_SCHEMES = [
	"major",
	"minor",
	"none"
];
const READING_ORDERS = [
	"context",
	"leftToRight",
	"rightToLeft"
];
/** Font boolean properties beyond bold, italic and strike, in the order Office writes them. */
const FONT_FLAGS = [
	"condense",
	"extend",
	"outline",
	"shadow"
];
/** Largest font size Office accepts, in points ([MS-OI29500] §2.1.677 c). */
const MAX_FONT_SIZE = 409.55;
/** Most cell format (`xf`) records in `cellXfs` ([MS-OI29500] §2.1.700 a, c). */
const MAX_CELL_XFS = 65430;
/** Most `font` records ([MS-OI29500] §2.1.707 a). */
const MAX_FONTS = 65491;
/** Most `fill` records ([MS-OI29500] §2.1.705 a). */
const MAX_FILLS = 65430;
/** Most `border` records ([MS-OI29500] §2.1.696). */
const MAX_BORDERS = 65430;
/** Caps for styles read from a file that is being edited: the earlier implementation caps (D-115, D-207). */
const EXISTING_MAX_FONT_NAME_LENGTH = 255;
const EXISTING_MAX_NUMFMT_LENGTH = 1024;
const STYLE_COUNT_LIMITS = {
	customNumFmts: {
		max: 206,
		what: "custom number formats",
		hint: "Use fewer distinct number format codes, for example one code such as \"#,##0.00\" for all amounts."
	},
	cellXfs: {
		max: MAX_CELL_XFS,
		what: "cell formats (cellXfs records)",
		hint: "Use fewer distinct styles, for example one style object such as { font: { bold: true } } for all header cells."
	},
	fonts: {
		max: MAX_FONTS,
		what: "fonts",
		hint: "Use fewer distinct fonts, for example one font such as { name: \"Arial\", size: 11 } for all body cells."
	},
	fills: {
		max: MAX_FILLS,
		what: "fills",
		hint: "Use fewer distinct fills, for example one fill such as { color: \"DDEBF7\" } for all highlighted cells."
	},
	borders: {
		max: MAX_BORDERS,
		what: "borders",
		hint: "Use fewer distinct borders, for example one border such as { bottom: { style: \"thin\" } } for all underlined cells."
	}
};
/**
* Throws `XLV_E_LIMIT_EXCEEDED` when a styles part would have more records of one kind than
* Office accepts ([MS-OI29500] §2.1.696, §2.1.700, §2.1.705, §2.1.707, §2.1.713 g).
*/
function checkStyleCount(limit, count) {
	const l = STYLE_COUNT_LIMITS[limit];
	if (count <= l.max) return;
	throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The workbook needs ${String(count)} ${l.what}; Office accepts at most ${String(l.max)}.`, {
		hint: l.hint,
		details: {
			limit,
			value: l.max
		}
	});
}
function invalid$2(path, message, hint) {
	return new XlvError("XLV_E_INVALID_STYLE", `Invalid style: ${path} ${message}`, {
		hint,
		details: { property: path }
	});
}
/** Reads the own enumerable properties of `value` once each, rejecting unknown keys. */
function readObject(value, path, known) {
	if (typeof value !== "object" || value === null || Array.isArray(value) || value instanceof Date) throw invalid$2(path, "must be an object.", `Pass an object with some of these properties: ${known.join(", ")}.`);
	const out = /* @__PURE__ */ new Map();
	for (const key of Object.keys(value)) {
		if (!known.includes(key)) {
			const close = suggest(key, known);
			throw invalid$2(`${path}.${key.slice(0, 40)}`, "is not a known style property.", `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known properties of ${path}: ${known.join(", ")}.`);
		}
		const v = value[key];
		if (v !== void 0) out.set(key, v);
	}
	return out;
}
function oneOf(value, path, allowed) {
	if (typeof value === "string" && allowed.includes(value)) return value;
	const close = typeof value === "string" ? suggest(value, allowed) : [];
	throw invalid$2(path, `has the value ${describe(value)}.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Allowed values: ${allowed.join(", ")}.`);
}
function bool$1(value, path) {
	if (typeof value === "boolean") return value;
	throw invalid$2(path, `has the value ${describe(value)}.`, "Use true or false.");
}
function integer(value, path, min, max, extra, example) {
	if (typeof value === "number" && Number.isInteger(value) && (value >= min && value <= max || value === extra)) return value;
	throw invalid$2(path, `has the value ${describe(value)}.`, `Use a whole number from ${String(min)} to ${String(max)}${extra === void 0 ? "" : ` or ${String(extra)}`}${example === void 0 ? "" : `, such as ${String(example)}`}.`);
}
function describe(value) {
	if (typeof value === "string") return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
	if (typeof value === "number" || typeof value === "boolean" || value === null || value === void 0) return String(value);
	return `of type ${typeof value}`;
}
/** Validates a colour and returns its normalised form (ARGB upper case for strings). */
function normalizeColor(value, path) {
	if (typeof value === "string") {
		if (/^[0-9A-Fa-f]{6}$/.test(value)) return `FF${value.toUpperCase()}`;
		if (/^[0-9A-Fa-f]{8}$/.test(value)) return value.toUpperCase();
		throw invalid$2(path, `has the value ${describe(value)}.`, "Use 6 hex digits (RRGGBB, for example \"FF0000\") or 8 (AARRGGBB).");
	}
	const hint = "Use a hex string such as \"FF0000\", or { theme: 1 }, { indexed: 8 } or { auto: true }.";
	if (typeof value !== "object" || value === null) throw invalid$2(path, `has the value ${describe(value)}.`, hint);
	const keys = Object.keys(value);
	if (keys.includes("theme")) {
		const o = readObject(value, path, ["theme", "tint"]);
		const theme = integer(o.get("theme"), `${path}.theme`, 0, 65535);
		const tint = o.get("tint");
		if (tint === void 0 || tint === 0) return { theme };
		if (typeof tint !== "number" || !(tint >= -1 && tint <= 1)) throw invalid$2(`${path}.tint`, `has the value ${describe(tint)}.`, "Use a number from -1 to 1.");
		return {
			theme,
			tint
		};
	}
	if (keys.includes("indexed")) return { indexed: integer(readObject(value, path, ["indexed"]).get("indexed"), `${path}.indexed`, 0, 65535) };
	if (keys.includes("auto")) {
		if (readObject(value, path, ["auto"]).get("auto") !== true) throw invalid$2(`${path}.auto`, "must be true.", hint);
		return { auto: true };
	}
	throw invalid$2(path, "is not a colour.", hint);
}
function normalizeFont(value, source) {
	const o = readObject(value, "font", [
		"name",
		"size",
		"bold",
		"italic",
		"underline",
		"strike",
		"color",
		"vertAlign",
		"family",
		"scheme",
		"charset",
		...FONT_FLAGS
	]);
	const out = {};
	const name = o.get("name");
	if (name !== void 0) {
		const max = source === "input" ? 31 : EXISTING_MAX_FONT_NAME_LENGTH;
		if (typeof name !== "string" || name.length === 0 || name.length > max) throw invalid$2("font.name", `has the value ${describe(name)}.`, `Use a font name of 1 to ${String(max)} characters, such as "Arial".`);
		out.name = name;
	}
	const size = o.get("size");
	if (size !== void 0) {
		if (!(typeof size === "number" && (source === "input" ? size >= 1 && size <= 409.55 : size > 0 && Number.isFinite(size)))) throw invalid$2("font.size", `has the value ${describe(size)}.`, `Use a size in points from ${String(1)} to ${String(MAX_FONT_SIZE)}, such as 11.`);
		out.size = size;
	}
	for (const key of [
		"bold",
		"italic",
		"strike",
		...FONT_FLAGS
	]) {
		const v = o.get(key);
		if (v !== void 0 && bool$1(v, `font.${key}`)) out[key] = true;
	}
	const underline = o.get("underline");
	if (underline !== void 0) {
		if (typeof underline === "boolean") {
			if (underline) out.underline = "single";
		} else out.underline = oneOf(underline, "font.underline", UNDERLINES);
	}
	const color = o.get("color");
	if (color !== void 0) out.color = normalizeColor(color, "font.color");
	const vertAlign = o.get("vertAlign");
	if (vertAlign !== void 0) {
		const v = oneOf(vertAlign, "font.vertAlign", VERT_ALIGN_RUNS);
		if (v !== "baseline") out.vertAlign = v;
	}
	const family = o.get("family");
	if (family !== void 0) out.family = integer(family, "font.family", 0, 14, void 0, 2);
	const scheme = o.get("scheme");
	if (scheme !== void 0) out.scheme = oneOf(scheme, "font.scheme", FONT_SCHEMES);
	const charset = o.get("charset");
	if (charset !== void 0) out.charset = integer(charset, "font.charset", 0, 255, void 0, 1);
	return out;
}
function finite(value, path, min, max, example) {
	if (typeof value === "number" && Number.isFinite(value) && value >= min && value <= max) return value;
	throw invalid$2(path, `has the value ${describe(value)}.`, `Use a number from ${String(min)} to ${String(max)}, such as ${String(example)}.`);
}
function normalizeGradient(value) {
	const o = readObject(value, "fill", [
		"type",
		"gradient",
		"degree",
		"left",
		"right",
		"top",
		"bottom",
		"stops"
	]);
	const out = {
		type: "gradient",
		gradient: oneOf(o.get("gradient"), "fill.gradient", ["linear", "path"]),
		stops: []
	};
	const degree = o.get("degree");
	if (degree !== void 0) {
		const d = finite(degree, "fill.degree", -Number.MAX_VALUE, Number.MAX_VALUE, 90);
		if (d !== 0) out.degree = d;
	}
	for (const side of [
		"left",
		"right",
		"top",
		"bottom"
	]) {
		const v = o.get(side);
		if (v === void 0) continue;
		const n = finite(v, `fill.${side}`, -Number.MAX_VALUE, Number.MAX_VALUE, .5);
		if (n !== 0) out[side] = n;
	}
	const stops = o.get("stops");
	if (!Array.isArray(stops)) throw invalid$2("fill.stops", "must be an array of stops.", "Use stops such as [{ position: 0, color: \"FFFFFF\" }, { position: 1, color: \"4472C4\" }].");
	stops.forEach((stop, i) => {
		const so = readObject(stop, `fill.stops[${String(i)}]`, ["position", "color"]);
		out.stops.push({
			position: finite(so.get("position"), `fill.stops[${String(i)}].position`, -Number.MAX_VALUE, Number.MAX_VALUE, .5),
			color: normalizeColor(so.get("color"), `fill.stops[${String(i)}].color`)
		});
	});
	return out;
}
function normalizeFill(value) {
	if (typeof value === "object" && value !== null && Object.keys(value).includes("type")) {
		if (value.type !== "gradient") throw invalid$2("fill.type", `has the value ${describe(value.type)}.`, "Use \"gradient\" for a gradient fill, or leave type out for a pattern fill.");
		return normalizeGradient(value);
	}
	if (typeof value === "object" && value !== null && Object.keys(value).includes("color")) return {
		pattern: "solid",
		fgColor: normalizeColor(readObject(value, "fill", ["color"]).get("color"), "fill.color")
	};
	const o = readObject(value, "fill", [
		"pattern",
		"fgColor",
		"bgColor"
	]);
	const out = { pattern: oneOf(o.get("pattern"), "fill.pattern", PATTERNS) };
	const fg = o.get("fgColor");
	if (fg !== void 0) out.fgColor = normalizeColor(fg, "fill.fgColor");
	const bg = o.get("bgColor");
	if (bg !== void 0) out.bgColor = normalizeColor(bg, "fill.bgColor");
	return out;
}
function normalizeBorder(value) {
	const o = readObject(value, "border", [
		...EDGES,
		"diagonal",
		"diagonalUp",
		"diagonalDown"
	]);
	const out = {};
	for (const flag of ["diagonalUp", "diagonalDown"]) {
		const v = o.get(flag);
		if (v !== void 0 && bool$1(v, `border.${flag}`)) out[flag] = true;
	}
	for (const edge of [...EDGES, "diagonal"]) {
		const e = o.get(edge);
		if (e === void 0) continue;
		const eo = readObject(e, `border.${edge}`, ["style", "color"]);
		const style = oneOf(eo.get("style"), `border.${edge}.style`, BORDER_STYLES);
		const color = eo.get("color");
		out[edge] = color === void 0 ? { style } : {
			style,
			color: normalizeColor(color, `border.${edge}.color`)
		};
	}
	return out;
}
function normalizeAlignment(value, source) {
	const o = readObject(value, "alignment", [
		"horizontal",
		"vertical",
		"wrapText",
		"indent",
		"textRotation",
		"shrinkToFit",
		"justifyLastLine",
		"readingOrder"
	]);
	const out = {};
	const h = o.get("horizontal");
	if (h !== void 0) out.horizontal = oneOf(h, "alignment.horizontal", HORIZONTAL);
	const v = o.get("vertical");
	if (v !== void 0) out.vertical = oneOf(v, "alignment.vertical", VERTICAL);
	for (const key of [
		"wrapText",
		"shrinkToFit",
		"justifyLastLine"
	]) {
		const v = o.get(key);
		if (v !== void 0 && bool$1(v, `alignment.${key}`)) out[key] = true;
	}
	const order = o.get("readingOrder");
	if (order !== void 0) {
		const r = oneOf(order, "alignment.readingOrder", READING_ORDERS);
		if (r !== "context") out.readingOrder = r;
	}
	const indent = o.get("indent");
	if (indent !== void 0) {
		const n = integer(indent, "alignment.indent", 0, source === "input" ? 255 : Number.MAX_SAFE_INTEGER, void 0, 2);
		if (n !== 0) out.indent = n;
	}
	const rotation = o.get("textRotation");
	if (rotation !== void 0) {
		const n = integer(rotation, "alignment.textRotation", 0, 180, 255, 45);
		if (n !== 0) out.textRotation = n;
	}
	return out;
}
/**
* Validates a style and returns a normalised copy. Throws `XLV_E_INVALID_STYLE` for unknown keys
* (including `__proto__`), wrong types and values outside the allowed sets. Each property of the
* input is read exactly once. With `source` `'input'` (the default) the values must be within the
* limits Office accepts; `'existing'` is for styles read from a file being edited, which are kept
* even when they are outside those limits.
*/
function normalizeStyle(value, source = "input") {
	const o = readObject(value, "style", [
		"font",
		"fill",
		"border",
		"alignment",
		"protection",
		"numFmt"
	]);
	const out = {};
	const font = o.get("font");
	if (font !== void 0) {
		const f = normalizeFont(font, source);
		if (Object.keys(f).length > 0) out.font = f;
	}
	const fill = o.get("fill");
	if (fill !== void 0) out.fill = normalizeFill(fill);
	const border = o.get("border");
	if (border !== void 0) {
		const b = normalizeBorder(border);
		if (Object.keys(b).length > 0) out.border = b;
	}
	const alignment = o.get("alignment");
	if (alignment !== void 0) {
		const a = normalizeAlignment(alignment, source);
		if (Object.keys(a).length > 0) out.alignment = a;
	}
	const protection = o.get("protection");
	if (protection !== void 0) {
		const po = readObject(protection, "protection", ["locked", "hidden"]);
		const p = {};
		for (const key of ["locked", "hidden"]) {
			const v = po.get(key);
			if (v !== void 0) p[key] = bool$1(v, `protection.${key}`);
		}
		if (Object.keys(p).length > 0) out.protection = p;
	}
	const numFmt = o.get("numFmt");
	if (typeof numFmt === "number") {
		if (!BUILTIN_NUMBER_FORMATS.has(numFmt)) throw invalid$2("numFmt", `has the value ${describe(numFmt)}, which is not a built-in number format id.`, `Use a built-in id (${[...BUILTIN_NUMBER_FORMATS.keys()].join(", ")}), such as 14 for a date, or a format code such as "yyyy-mm-dd".`);
		if (numFmt !== 0) out.numFmt = numFmt;
	} else if (numFmt !== void 0) {
		const max = source === "input" ? 254 : EXISTING_MAX_NUMFMT_LENGTH;
		if (typeof numFmt !== "string" || numFmt.length === 0 || numFmt.length > max) throw invalid$2("numFmt", `has the value ${describe(numFmt)}.`, `Use a number format code of 1 to ${String(max)} characters, such as "#,##0.00" or "yyyy-mm-dd".`);
		if (numFmt !== "General") out.numFmt = numFmt;
	}
	return out;
}
/** Merges `patch` into `base` part by part (font, fill, border, alignment, protection, numFmt). */
function mergeStyles(base, patch) {
	const out = { ...base };
	if (patch.font) out.font = {
		...base.font,
		...patch.font
	};
	if (patch.fill) out.fill = patch.fill;
	if (patch.border) out.border = {
		...base.border,
		...patch.border
	};
	if (patch.alignment) out.alignment = {
		...base.alignment,
		...patch.alignment
	};
	if (patch.protection) out.protection = {
		...base.protection,
		...patch.protection
	};
	if (patch.numFmt !== void 0) out.numFmt = patch.numFmt;
	return out;
}
/** A stable string key for a value made of plain objects, arrays and primitives. */
function stableKey(value) {
	if (typeof value !== "object" || value === null) return JSON.stringify(value) ?? "undefined";
	if (Array.isArray(value)) return `[${value.map(stableKey).join(",")}]`;
	return `{${Object.keys(value).sort().map((k) => `${JSON.stringify(k)}:${stableKey(value[k])}`).join(",")}}`;
}
// src/sml/styles.ts
const EMPTY_STYLE = Object.freeze({});
/** The number format id of General (§18.8.30, built-in id 0). */
const GENERAL = 0;
const MAX_UNSIGNED_INT = 4294967295;
var NumberFormats = class {
	ids;
	codes;
	xfs;
	fonts;
	fills;
	borders;
	dxfs;
	cellFormatCount;
	resolved = /* @__PURE__ */ new Map();
	constructor(ids, codes, xfs = [], fonts = [], fills = [], borders = [], dxfs = []) {
		this.ids = ids;
		this.codes = codes;
		this.xfs = xfs;
		this.fonts = fonts;
		this.fills = fills;
		this.borders = borders;
		this.dxfs = dxfs;
		this.cellFormatCount = ids.length;
	}
	defaultFont() {
		return this.fonts[0];
	}
	dxf(index) {
		return Number.isInteger(index) && index >= 0 ? this.dxfs[index] : void 0;
	}
	cellStyle(styleIndex, numFmt) {
		const cached = this.resolved.get(styleIndex);
		if (cached !== void 0) return cached;
		const xf = Number.isInteger(styleIndex) && styleIndex >= 0 ? this.xfs[styleIndex] : void 0;
		if (xf === void 0 && numFmt === null) return EMPTY_STYLE;
		const style = {};
		if (xf !== void 0) {
			const font = xf.fontId === 0 ? void 0 : this.fonts[xf.fontId];
			if (font !== void 0) style.font = font;
			const fill = this.fills[xf.fillId];
			if (fill !== void 0) style.fill = fill;
			const border = this.borders[xf.borderId];
			if (border !== void 0) style.border = border;
			if (xf.alignment !== void 0) style.alignment = xf.alignment;
			if (xf.protection !== void 0) style.protection = xf.protection;
		}
		if (numFmt !== null) style.numFmt = numFmt;
		const frozen = Object.freeze(style);
		if (xf !== void 0) this.resolved.set(styleIndex, frozen);
		return frozen;
	}
	numFmtId(styleIndex) {
		if (!Number.isInteger(styleIndex) || styleIndex < 0 || styleIndex >= this.ids.length) return GENERAL;
		return this.ids[styleIndex] ?? GENERAL;
	}
	customFormatCode(numFmtId) {
		return this.codes.get(numFmtId);
	}
};
/** Styles of a workbook without a Styles part: no cell formats, everything General. */
const EMPTY_STYLES = new NumberFormats([], /* @__PURE__ */ new Map());
/**
* Parses the Styles part. Throws `XLV_E_PART_INVALID` if the root element is not `styleSheet`
* in a SpreadsheetML namespace, and the XML parser's errors for malformed XML. Recoverable
* problems are reported to `ctx.warnings` as `XLV_W_INVALID_VALUE`.
*/
function parseStyles(xml, ctx) {
	const root = parseXmlTree(xml, ctx);
	if (root.local !== "styleSheet" || root.ns !== SML_NS.strict && root.ns !== SML_NS.transitional) throw new XlvError("XLV_E_PART_INVALID", `The styles part has the root element <${root.local}> in namespace "${root.ns}"; expected <styleSheet> in the SpreadsheetML namespace.`, {
		location: { part: ctx.part },
		hint: "The part the workbook names as its styles is not a SpreadsheetML style sheet. Re-save the file in a spreadsheet application; if it was generated by a program, report the problem to its authors.",
		details: {
			element: root.local,
			namespace: root.ns
		}
	});
	const ns = root.ns;
	const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, { part: ctx.part });
	/** The first child named `local`; further ones (the schema allows one) are reported and ignored. */
	const single = (local) => {
		const found = childrenNamed(root, ns, local);
		if (found.length > 1) warn(`The styles part has ${String(found.length)} <${local}> elements; only the first is used.`);
		return found[0];
	};
	const codes = /* @__PURE__ */ new Map();
	const numFmts = single("numFmts");
	for (const el of numFmts === void 0 ? [] : childrenNamed(numFmts, ns, "numFmt")) {
		const rawId = el.attrs.get("numFmtId");
		const formatCode = el.attrs.get("formatCode");
		if (rawId === void 0 || formatCode === void 0) {
			warn(`A numFmt element has no ${rawId === void 0 ? "numFmtId" : "formatCode"} attribute; it is ignored.`);
			continue;
		}
		const id = parseUnsignedInt(rawId);
		if (id === null) {
			warn(`numFmt@numFmtId has the invalid value ${quote(rawId)}; the number format is ignored.`);
			continue;
		}
		if (codes.has(id)) {
			warn(`Number format ${String(id)} is declared more than once; the first declaration is used.`);
			continue;
		}
		codes.set(id, decodeXstring(formatCode));
	}
	const listOf = (container, item, parse) => {
		const el = single(container);
		return el === void 0 ? [] : childrenNamed(el, ns, item).map(parse);
	};
	const colorOf = (el, what) => el === void 0 ? void 0 : parseColor(el, what, warn);
	const fonts = listOf("fonts", "font", (el) => parseFont(el, ns, colorOf, warn));
	const fills = listOf("fills", "fill", (el) => parseFill(el, ns, colorOf, warn));
	const borders = listOf("borders", "border", (el) => parseBorder(el, ns, colorOf, warn));
	const cellXfs = single("cellXfs");
	const xfElements = cellXfs === void 0 ? [] : childrenNamed(cellXfs, ns, "xf");
	const index = (xf, name) => {
		const raw = xf.attrs.get(name);
		if (raw === void 0) return 0;
		const n = parseUnsignedInt(raw);
		if (n === null) warn(`cellXfs/xf@${name} has the invalid value ${quote(raw)}; using 0.`);
		return n ?? 0;
	};
	const xfs = xfElements.map((xf) => ({
		fontId: index(xf, "fontId"),
		fillId: index(xf, "fillId"),
		borderId: index(xf, "borderId"),
		alignment: parseAlignment(childNamed(xf, ns, "alignment"), warn),
		protection: parseProtection(childNamed(xf, ns, "protection"))
	}));
	return new NumberFormats(xfElements.map((xf) => {
		const raw = xf.attrs.get("numFmtId");
		if (raw === void 0) return GENERAL;
		const id = parseUnsignedInt(raw);
		if (id === null) {
			warn(`cellXfs/xf@numFmtId has the invalid value ${quote(raw)}; using the General number format.`);
			return GENERAL;
		}
		return id;
	}), codes, xfs, fonts, fills, borders, listOf("dxfs", "dxf", (el) => parseDxf(el, ns, colorOf, warn)));
}
/**
* Parses a differential format (CT_Dxf, SR-16): font, number format, fill and border; the
* properties a dxf leaves out are not part of the result. A pattern fill without `patternType`
* is read as solid, and a solid fill whose background equals its foreground as its foreground
* only (implementation choice, SR-16: the standard does not say how a dxf's fill colours apply).
*/
function parseDxf(el, ns, colorOf, warn) {
	const style = {};
	const fontEl = childNamed(el, ns, "font");
	if (fontEl !== void 0) {
		const font = parseFont(fontEl, ns, colorOf, warn);
		if (Object.keys(font).length > 0) style.font = font;
	}
	const code = childNamed(el, ns, "numFmt")?.attrs.get("formatCode");
	if (code !== void 0) style.numFmt = decodeXstring(code);
	const fillEl = childNamed(el, ns, "fill");
	const pf = fillEl === void 0 ? void 0 : childNamed(fillEl, ns, "patternFill");
	if (fillEl !== void 0 && pf !== void 0 && pf.attrs.get("patternType") === void 0) {
		const fg = colorOf(childNamed(pf, ns, "fgColor"), "fill");
		const bg = colorOf(childNamed(pf, ns, "bgColor"), "fill");
		if (fg !== void 0 || bg !== void 0) style.fill = Object.freeze({
			pattern: "solid",
			...fg === void 0 ? {} : { fgColor: fg },
			...bg === void 0 ? {} : { bgColor: bg }
		});
	} else if (fillEl !== void 0) {
		const fill = parseFill(fillEl, ns, colorOf, warn);
		if (fill !== void 0) {
			const f = fill;
			style.fill = f.pattern === "solid" && f.fgColor !== void 0 && f.bgColor !== void 0 && JSON.stringify(f.fgColor) === JSON.stringify(f.bgColor) && f.fgColor !== void 0 ? Object.freeze({
				pattern: "solid",
				fgColor: f.fgColor
			}) : fill;
		}
	}
	const alignment = parseAlignment(childNamed(el, ns, "alignment"), warn);
	if (alignment !== void 0) style.alignment = alignment;
	const borderEl = childNamed(el, ns, "border");
	const border = borderEl === void 0 ? void 0 : parseBorder(borderEl, ns, colorOf, warn);
	if (border !== void 0) style.border = border;
	const protection = parseProtection(childNamed(el, ns, "protection"));
	if (protection !== void 0) style.protection = protection;
	return Object.freeze(style);
}
/**
* Reads a differential format element outside the styles part, such as `x14:dxf` of an extended
* conditional formatting rule ([MS-XLSX] §2.6.27: of type CT_Dxf, with children in `ns`).
*/
function parseDxfElement(el, ns, warn) {
	return parseDxf(el, ns, (c, what) => c === void 0 ? void 0 : parseColor(c, what, warn), warn);
}
/**
* Parses CT_Color (W2-16): rgb as ARGB, or theme (with tint), indexed or auto. An rgb of exactly
* 6 hex digits is not valid (ST_UnsignedIntHex has 4 bytes) but some writers produce it; it is
* read as an opaque RGB colour, alpha FF, with a warning (W2-16a).
*/
function parseColor(el, what, warn) {
	const rgb = el.attrs.get("rgb");
	if (rgb !== void 0) {
		if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return rgb.toUpperCase();
		if (/^[0-9A-Fa-f]{6}$/.test(rgb)) {
			const argb = `FF${rgb.toUpperCase()}`;
			warn(`The ${what} colour ${quote(rgb)} is not an ARGB value; it was read as the opaque RGB colour "${argb}".`);
			return argb;
		}
		warn(`The ${what} colour ${quote(rgb)} is not an ARGB value; it is ignored.`);
		return;
	}
	const theme = el.attrs.get("theme");
	if (theme !== void 0) {
		const n = parseUnsignedInt(theme);
		if (n === null) return void 0;
		const rawTint = el.attrs.get("tint");
		const tint = rawTint === void 0 ? 0 : Number(rawTint.trim());
		return Number.isFinite(tint) && tint !== 0 && tint >= -1 && tint <= 1 ? Object.freeze({
			theme: n,
			tint
		}) : Object.freeze({ theme: n });
	}
	const indexed = el.attrs.get("indexed");
	if (indexed !== void 0) {
		const n = parseUnsignedInt(indexed);
		return n === null ? void 0 : Object.freeze({ indexed: n });
	}
	const auto = el.attrs.get("auto");
	if (auto !== void 0 && parseBoolean(auto) === true) return Object.freeze({ auto: true });
}
/** xsd:boolean: "true", "false", "1", "0" with surrounding white space. */
function parseBoolean(raw) {
	const s = raw.trim();
	if (s === "true" || s === "1") return true;
	if (s === "false" || s === "0") return false;
	return null;
}
/** CT_BooleanProperty: `val` defaults to true (W2-11). */
function booleanProperty(el) {
	if (el === void 0) return false;
	const val = el.attrs.get("val");
	return val === void 0 ? true : parseBoolean(val) === true;
}
/**
* Parses the run properties of a rich text run (CT_RPrElt, F-SST-08): the font properties of
* CT_Font, with the font name in `rFont` instead of `name`. Unknown or invalid properties are left
* out and reported through `warn`.
*/
function parseRunFont(el, ns, warn) {
	return parseFont(el, ns, (c, what) => c === void 0 ? void 0 : parseColor(c, what, warn), warn, "rFont");
}
/** Parses CT_Font (W2-11). Unknown or invalid properties are left out. */
function parseFont(el, ns, colorOf, warn, nameElement = "name") {
	const font = {};
	const name = childNamed(el, ns, nameElement)?.attrs.get("val");
	if (name !== void 0) font.name = decodeXstring(name);
	const sz = childNamed(el, ns, "sz")?.attrs.get("val");
	if (sz !== void 0) {
		const n = Number(sz.trim());
		if (sz.trim() !== "" && Number.isFinite(n) && n > 0) font.size = n;
		else warn(`The font size ${quote(sz)} is not valid; it is ignored.`);
	}
	if (booleanProperty(childNamed(el, ns, "b"))) font.bold = true;
	if (booleanProperty(childNamed(el, ns, "i"))) font.italic = true;
	if (booleanProperty(childNamed(el, ns, "strike"))) font.strike = true;
	const u = childNamed(el, ns, "u");
	if (u !== void 0) {
		const val = u.attrs.get("val") ?? "single";
		const found = UNDERLINES.find((x) => x === val);
		if (found !== void 0) font.underline = found;
		else if (val !== "none") warn(`The underline style ${quote(val)} is unknown; it is ignored.`);
	}
	const color = colorOf(childNamed(el, ns, "color"), "font");
	if (color !== void 0) font.color = color;
	for (const key of FONT_FLAGS) if (booleanProperty(childNamed(el, ns, key))) font[key] = true;
	const vertAlign = childNamed(el, ns, "vertAlign")?.attrs.get("val");
	if (vertAlign !== void 0) {
		const found = VERT_ALIGN_RUNS.find((x) => x === vertAlign);
		if (found === void 0) warn(`The vertical text alignment ${quote(vertAlign)} is unknown; it is ignored.`);
		else if (found !== "baseline") font.vertAlign = found;
	}
	const scheme = childNamed(el, ns, "scheme")?.attrs.get("val");
	if (scheme !== void 0) {
		const found = FONT_SCHEMES.find((x) => x === scheme);
		if (found === void 0) warn(`The font scheme ${quote(scheme)} is unknown; it is ignored.`);
		else font.scheme = found;
	}
	for (const [key, max] of [["family", 14], ["charset", 255]]) {
		const raw = childNamed(el, ns, key)?.attrs.get("val");
		if (raw === void 0) continue;
		const n = parseUnsignedInt(raw);
		if (n === null || n > max) warn(`The font ${key} ${quote(raw)} is not valid; it is ignored.`);
		else font[key] = n;
	}
	return Object.freeze(font);
}
/** Parses CT_CellProtection (W2-63): the values given in the file, as booleans. */
function parseProtection(el) {
	if (el === void 0) return void 0;
	const p = {};
	for (const key of ["locked", "hidden"]) {
		const raw = el.attrs.get(key);
		const b = raw === void 0 ? null : parseBoolean(raw);
		if (b !== null) p[key] = b;
	}
	return Object.keys(p).length > 0 ? Object.freeze(p) : void 0;
}
/** Parses CT_GradientFill (W2-60): the type, the numbers other than their default 0, and the stops. */
function parseGradient(el, ns, colorOf, warn) {
	const number = (name) => {
		const raw = el.attrs.get(name);
		if (raw === void 0) return void 0;
		const n = Number(raw.trim());
		if (raw.trim() === "" || !Number.isFinite(n)) {
			warn(`gradientFill@${name} has the invalid value ${quote(raw)}; it is ignored.`);
			return;
		}
		return n === 0 ? void 0 : n;
	};
	const stops = [];
	for (const stop of childrenNamed(el, ns, "stop")) {
		const raw = stop.attrs.get("position") ?? "";
		const position = Number(raw.trim());
		const color = colorOf(childNamed(stop, ns, "color"), "gradient stop");
		if (raw.trim() === "" || !Number.isFinite(position) || color === void 0) {
			warn("A gradient stop has no valid position or colour; it is ignored.");
			continue;
		}
		stops.push(Object.freeze({
			position,
			color
		}));
	}
	const fill = {};
	for (const k of [
		"degree",
		"left",
		"right",
		"top",
		"bottom"
	]) {
		const n = number(k);
		if (n !== void 0) fill[k] = n;
	}
	return Object.freeze({
		type: "gradient",
		gradient: el.attrs.get("type") === "path" ? "path" : "linear",
		...fill,
		stops: Object.freeze(stops)
	});
}
/** Parses CT_Fill (W2-14, W2-60): gradient fills, and pattern fills other than "none". */
function parseFill(el, ns, colorOf, warn) {
	const gf = childNamed(el, ns, "gradientFill");
	if (gf !== void 0) return parseGradient(gf, ns, colorOf, warn);
	const pf = childNamed(el, ns, "patternFill");
	if (pf === void 0) return void 0;
	const type = pf.attrs.get("patternType");
	if (type === void 0 || type === "none") return void 0;
	const pattern = PATTERNS.find((p) => p === type);
	if (pattern === void 0) {
		warn(`The fill pattern ${quote(type)} is unknown; it is ignored.`);
		return;
	}
	const fill = { pattern };
	const fg = colorOf(childNamed(pf, ns, "fgColor"), "fill");
	if (fg !== void 0) fill.fgColor = fg;
	const bg = colorOf(childNamed(pf, ns, "bgColor"), "fill");
	if (bg !== void 0) fill.bgColor = bg;
	return Object.freeze(fill);
}
/**
* Parses CT_Border (W2-15, W2-61): the left, right, top, bottom and diagonal edges whose style is
* not "none", and the diagonal directions.
*/
function parseBorder(el, ns, colorOf, warn) {
	const border = {};
	let any = false;
	for (const flag of ["diagonalUp", "diagonalDown"]) {
		const raw = el.attrs.get(flag);
		if (raw !== void 0 && parseBoolean(raw) === true) {
			border[flag] = true;
			any = true;
		}
	}
	for (const edge of [...EDGES, "diagonal"]) {
		const e = childNamed(el, ns, edge);
		const raw = e?.attrs.get("style");
		if (e === void 0 || raw === void 0 || raw === "none") continue;
		const style = BORDER_STYLES.find((b) => b === raw);
		if (style === void 0) {
			warn(`The border style ${quote(raw)} is unknown; it is ignored.`);
			continue;
		}
		const color = colorOf(childNamed(e, ns, "color"), "border");
		border[edge] = Object.freeze(color === void 0 ? { style } : {
			style,
			color
		});
		any = true;
	}
	return any ? Object.freeze(border) : void 0;
}
/** Parses CT_CellAlignment (W2-17): the properties xlvane models, defaults left out. */
function parseAlignment(el, warn) {
	if (el === void 0) return void 0;
	const a = {};
	const h = el.attrs.get("horizontal");
	if (h !== void 0) {
		const found = HORIZONTAL.find((x) => x === h);
		if (found !== void 0) a.horizontal = found;
		else warn(`The horizontal alignment ${quote(h)} is unknown; it is ignored.`);
	}
	const v = el.attrs.get("vertical");
	if (v !== void 0) {
		const found = VERTICAL.find((x) => x === v);
		if (found !== void 0) a.vertical = found;
		else warn(`The vertical alignment ${quote(v)} is unknown; it is ignored.`);
	}
	for (const key of [
		"wrapText",
		"shrinkToFit",
		"justifyLastLine"
	]) {
		const raw = el.attrs.get(key);
		if (raw !== void 0 && parseBoolean(raw) === true) a[key] = true;
	}
	const order = el.attrs.get("readingOrder");
	if (order !== void 0) {
		const n = parseUnsignedInt(order);
		if (n === null || n > 2) warn(`The reading order ${quote(order)} is not valid; it is ignored.`);
		else if (n > 0) a.readingOrder = READING_ORDERS[n] ?? "context";
	}
	const indent = el.attrs.get("indent");
	const indentN = indent === void 0 ? null : parseUnsignedInt(indent);
	if (indentN !== null && indentN > 0) a.indent = indentN;
	const rotation = el.attrs.get("textRotation");
	const rotationN = rotation === void 0 ? null : parseUnsignedInt(rotation);
	if (rotationN !== null && rotationN > 180 && rotationN !== 255) warn(`The text rotation ${quote(rotation ?? "")} is not 0 to 180 or 255; it is ignored.`);
	else if (rotationN !== null && rotationN > 0) a.textRotation = rotationN;
	return Object.keys(a).length > 0 ? Object.freeze(a) : void 0;
}
/**
* Parses an `xsd:unsignedInt` value: ASCII digits (leading zeros allowed) with optional
* surrounding white space, at most 4294967295. Returns `null` for anything else. Linear time.
*/
function parseUnsignedInt(raw) {
	const s = raw.trim();
	if (s.length === 0) return null;
	let value = 0;
	for (let i = 0; i < s.length; i++) {
		const d = s.charCodeAt(i) - 48;
		if (d < 0 || d > 9) return null;
		value = value * 10 + d;
		if (value > MAX_UNSIGNED_INT) return null;
	}
	return value;
}
/** Quotes a value from the file for a message, shortened to keep messages small. */
function quote(value) {
	return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
}
// src/notes.ts
/** Comments part content type (Part 1 §12.3.3, W2-46). */
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
/** VML Drawing part content type (Part 4 §8.2, W2-47). */
const VML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing";
/** Most characters of an author ([MS-OI29500] §2.1.689 a). */
const MAX_AUTHOR_LENGTH = 65535;
/**
* Reads a Comments part (CT_Comments, W2-45). A comment whose `ref` is not a single cell, or
* whose `authorId` names no author, is read with a warning (skipped, or with author `null`).
* Malformed XML or another root element gives a warning and no notes: notes are annotations, so
* the sheet stays readable without them. Limits and DTD rules of the XML parser still apply.
*/
function parseComments(xml, ctx) {
	const warn = (message, cell) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, {
		part: ctx.part,
		sheet: ctx.sheet,
		...cell === void 0 ? {} : { cell }
	});
	let root;
	try {
		root = parseXmlTree(xml, ctx);
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			warn("The comments part of the sheet is not well-formed XML; its notes are ignored.");
			return [];
		}
		throw error;
	}
	const ns = root.ns;
	if (root.local !== "comments" || ns !== SML_NS.strict && ns !== SML_NS.transitional) {
		warn(`The comments part of the sheet has the root element <${root.local}>; its notes are ignored.`);
		return [];
	}
	const authors = childrenNamed(childNamed(root, ns, "authors") ?? root, ns, "author").map((a) => decodeXstring(a.text));
	const list = childNamed(root, ns, "commentList");
	const notes = [];
	if (list === void 0) return notes;
	for (const c of childrenNamed(list, ns, "comment")) {
		const rawRef = c.attrs.get("ref") ?? "";
		const address = parseCellRef(rawRef.trim().toUpperCase());
		if (address === null) {
			warn(`A note refers to "${rawRef.slice(0, 40)}", which is not a single cell; it is skipped.`);
			continue;
		}
		const ref = formatCellRef(address.row, address.col);
		const idText = c.attrs.get("authorId");
		const author = authors[idText === void 0 || !/^\d+$/.test(idText.trim()) ? -1 : Number(idText.trim())];
		if (author === void 0) warn(`The note of ${ref} names author ${String(idText ?? "(none)").slice(0, 20)}, which is not in the list of authors; its author is unknown.`, ref);
		const { text, runs } = readRst(childNamed(c, ns, "text"), ns, (m) => warn(m, ref));
		notes.push(Object.freeze(runs === void 0 ? {
			ref,
			text,
			author: author ?? null
		} : {
			ref,
			text,
			author: author ?? null,
			runs
		}));
	}
	return notes;
}
/**
* Reads CT_Rst (`t?`, `r*`, `rPh*`): the text is the own `t` and the `t` of every run;
* phonetic runs are not part of it (as for shared strings, D-SST-2).
*/
function readRst(el, ns, warn) {
	if (el === void 0) return { text: "" };
	let text = "";
	const runs = [];
	let rich = false;
	for (const child of el.children) {
		if (child.ns !== ns) continue;
		if (child.local === "t") {
			const t = decodeXstring(child.text);
			text += t;
			runs.push({ text: t });
		} else if (child.local === "r") {
			rich = true;
			const t = decodeXstring(childNamed(child, ns, "t")?.text ?? "");
			text += t;
			const rPr = childNamed(child, ns, "rPr");
			const font = rPr === void 0 ? {} : parseRunFont(rPr, ns, warn);
			runs.push(Object.keys(font).length > 0 ? Object.freeze({
				text: t,
				font
			}) : Object.freeze({ text: t }));
		}
	}
	return rich ? {
		text,
		runs: Object.freeze(runs)
	} : { text };
}
const VML_NS = "urn:schemas-microsoft-com:vml";
const OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const EXCEL_NS = "urn:schemas-microsoft-com:office:excel";
/** Millimetres per unit of the VML length units (Part 4 §19.1; px at 96 per inch, a convention, W2-69). */
const MM_PER_UNIT = {
	mm: 1,
	cm: 10,
	in: 25.4,
	pt: 25.4 / 72,
	pc: 25.4 / 6,
	px: 25.4 / 96
};
/** Part 4 §19.1.2.22: the default inset, "0.1in, 0.05in, 0.1in, 0.05in", in millimetres. */
const DEFAULT_INSET_MM = [
	2.54,
	1.27,
	2.54,
	1.27
];
/** ST_TrueFalseBlank (Part 4 §20.1.2.6): blank and t/true mean true. */
function trueFalseBlank(text) {
	const v = text.trim().toLowerCase();
	if (v === "" || v === "t" || v === "true") return true;
	if (v === "f" || v === "false") return false;
	return null;
}
/** Parses the `inset` of a text box into millimetres; `null` when a value has no or an unknown unit. */
function parseInset(inset) {
	const parts = inset.split(/\s*,\s*|\s+/);
	const out = [];
	for (let i = 0; i < 4; i++) {
		const p = (parts[i] ?? "").trim();
		if (p === "") {
			out.push(DEFAULT_INSET_MM[i]);
			continue;
		}
		const m = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(mm|cm|in|pt|pc|px)$/.exec(p);
		if (m === null) return null;
		out.push(Math.round(Number(m[1]) * MM_PER_UNIT[m[2]] * 1e4) / 1e4);
	}
	return {
		left: out[0],
		top: out[1],
		right: out[2],
		bottom: out[3]
	};
}
/**
* Reads the note boxes of a VML Drawing part (Part 4 §19.4; W2-69): for every shape with an
* `x:ClientData` of type `Note`, its cell (`x:Row`, `x:Column`, 0-based) and its settings.
* Malformed VML gives a warning and no boxes; the notes stay readable without them.
*/
function parseNoteBoxes(xml, ctx) {
	const boxes = /* @__PURE__ */ new Map();
	const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, {
		part: ctx.part,
		sheet: ctx.sheet
	});
	let root;
	try {
		root = parseXmlTree(xml, ctx);
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			warn("The VML drawing of the sheet is not well-formed XML; the settings of its note boxes are not read.");
			return boxes;
		}
		throw error;
	}
	const visit = (el) => {
		if (el.ns === VML_NS && el.local === "shape") {
			const data = el.children.find((c) => c.ns === EXCEL_NS && c.local === "ClientData");
			if (data !== void 0 && (data.attrs.get("ObjectType") ?? "").trim() === "Note") readShape(el, data);
			return;
		}
		for (const c of el.children) visit(c);
	};
	const readShape = (shape, data) => {
		const child = (name) => data.children.find((c) => c.ns === EXCEL_NS && c.local === name);
		const row = Number((child("Row")?.text ?? "").trim());
		const col = Number((child("Column")?.text ?? "").trim());
		if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || row >= 1048576 || col >= 16384 || child("Row") === void 0 || child("Column") === void 0) return;
		const flag = (name) => {
			const el = child(name);
			if (el === void 0) return true;
			const v = trueFalseBlank(el.text);
			if (v === null) warn(`The note box of row ${String(row + 1)}, column ${String(col + 1)} has x:${name} "${el.text.trim().slice(0, 20)}"; it is read as true.`);
			return v ?? true;
		};
		const move = child("MoveWithCells") !== void 0;
		const size = child("SizeWithCells") !== void 0;
		const textbox = shape.children.find((c) => c.ns === VML_NS && c.local === "textbox");
		const insetText = textbox?.attrs.get("inset");
		let margins = null;
		if (insetText !== void 0) {
			margins = parseInset(insetText);
			if (margins === null) warn(`The note box of row ${String(row + 1)}, column ${String(col + 1)} has margins "${insetText.slice(0, 40)}" without a known unit; they are not read.`);
		}
		const mode = (shape.attrs.get(`{${OFFICE_NS}}insetmode`) ?? textbox?.attrs.get(`{${OFFICE_NS}}insetmode`) ?? "custom").trim();
		boxes.set(formatCellRef(row + 1, col + 1), Object.freeze({
			autoMargins: mode === "auto",
			margins,
			locked: flag("Locked"),
			lockText: flag("LockText"),
			anchor: move && size ? "absolute" : move || size ? "oneCell" : "twoCell"
		}));
	};
	visit(root);
	return boxes;
}
// src/defined-names.ts
/** [MS-OI29500] §2.1.568 k (name) and j (comment). */
const MAX_NAME_LENGTH = 255;
const MAX_COMMENT_LENGTH = 255;
/** The writer's formula limit, as for cell formulas (write.ts). */
const MAX_REFERS_TO_LENGTH = 8192;
const NAME_RE = /^(?:[A-Za-z_\\]|[^\x00-\x7F])(?:[A-Za-z0-9_\\.?]|[^\x00-\x7F])*$/u;
/** A1 cell references (also beyond the grid, to stay on the safe side) and R1C1 forms (F2-09). */
const A1_NAME_RE = /^[A-Za-z]{1,3}\d+$/;
const R1C1_NAME_RE = /^(?:[Rr]\d*(?:[Cc]\d*)?|[Cc]\d*(?:[Rr]\d*)?)$/;
function invalidName(name, message, hint) {
	return new XlvError("XLV_E_INVALID_DEFINED_NAME", message, {
		hint,
		details: { name: name.slice(0, 100) }
	});
}
/**
* Checks a defined name given by a caller. `sheets` holds the workbook's sheet names in sheet
* order, or is `null` when the sheets are not known yet (the sheet is then checked later by the
* caller).
*/
function checkDefinedName(name, refersTo, options, sheets) {
	if (typeof name !== "string") throw invalidName(String(name), "A defined name must be a string.", "For example workbook.defineName(\"Rates\", \"Sheet1!$A$1:$A$10\").");
	const example = "Use letters, digits, \"_\", \".\" and \"\\\", starting with a letter, \"_\" or \"\\\", for example \"Sales_2026\".";
	if (name.length < 1 || name.length > MAX_NAME_LENGTH) throw invalidName(name, `The defined name "${name.slice(0, 40)}" has ${String(name.length)} characters; names have 1 to ${String(MAX_NAME_LENGTH)}.`, example);
	if (!NAME_RE.test(name) || firstLoneSurrogate(name) !== -1) throw invalidName(name, `"${name.slice(0, 40)}" is not a valid defined name.`, `${example} Spaces and characters such as "-", "!" or "$" are not allowed.`);
	const upper = name.toUpperCase();
	if (upper === "TRUE" || upper === "FALSE" || A1_NAME_RE.test(name) || R1C1_NAME_RE.test(name)) throw invalidName(name, `"${name}" cannot be a defined name because it reads as ${upper === "TRUE" || upper === "FALSE" ? "a boolean" : "a cell reference"}.`, "Add a word or an underscore, for example \"Rate_A1\" or \"_A1\".");
	if (upper.startsWith("_XL")) throw invalidName(name, `Defined names starting with "_xl" are reserved for built-in names ("${name.slice(0, 40)}").`, "Choose another name. Built-in names such as print areas are set with their own options.");
	if (name.length === 2 && name.startsWith("\\")) throw invalidName(name, `"${name}" cannot be a defined name: Office does not load names of a backslash and one character.`, "Use a longer name.");
	const formula = typeof refersTo === "string" && refersTo.startsWith("=") ? refersTo.slice(1) : refersTo;
	if (typeof formula !== "string" || formula.trim() === "" || formula.length > MAX_REFERS_TO_LENGTH || firstLoneSurrogate(formula) !== -1) throw invalidName(name, `The formula of the defined name "${name}" is empty, too long or not valid text.`, `Pass what the name refers to, with or without "=", for example "Sheet1!$A$1:$B$10" or "0.19"; at most ${String(MAX_REFERS_TO_LENGTH)} characters.`);
	let sheet = null;
	let hidden = false;
	let comment = null;
	if (options !== void 0) {
		if (typeof options !== "object" || options === null || Array.isArray(options)) throw new XlvError("XLV_E_INVALID_OPTION", "Defined name options must be an object.", {
			hint: "For example { sheet: \"Data\", hidden: true }.",
			details: { option: "options" }
		});
		for (const key of Object.keys(options)) if (key !== "sheet" && key !== "hidden" && key !== "comment") throw new XlvError("XLV_E_INVALID_OPTION", `Unknown option "${key.slice(0, 40)}" for the defined name "${name}".`, {
			hint: "Known options: sheet, hidden, comment.",
			details: {
				option: key.slice(0, 40),
				known: [
					"sheet",
					"hidden",
					"comment"
				]
			}
		});
		const o = options;
		if (o["sheet"] !== void 0) {
			if (typeof o["sheet"] !== "string") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"sheet\" must be a sheet name.", {
				hint: "For example { sheet: \"Data\" }.",
				details: { option: "sheet" }
			});
			sheet = o["sheet"];
		}
		if (o["hidden"] !== void 0) {
			if (typeof o["hidden"] !== "boolean") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"hidden\" must be true or false.", {
				hint: "Use hidden: true or leave it out.",
				details: { option: "hidden" }
			});
			hidden = o["hidden"];
		}
		if (o["comment"] !== void 0) {
			const c = o["comment"];
			if (typeof c !== "string" || c.length > MAX_COMMENT_LENGTH || firstLoneSurrogate(c) !== -1) throw invalidName(name, `The comment of the defined name "${name}" must be text of at most ${String(MAX_COMMENT_LENGTH)} characters.`, "Shorten the comment.");
			comment = c;
		}
	}
	if (sheet !== null && sheets !== null) checkNameSheet(name, sheet, sheets);
	return {
		name,
		refersTo: formula,
		sheet,
		hidden,
		comment
	};
}
/** Throws `XLV_E_SHEET_NOT_FOUND` if the sheet of a local name does not exist. */
function checkNameSheet(name, sheet, sheets) {
	if (sheets.has(sheet)) return;
	throw new XlvError("XLV_E_SHEET_NOT_FOUND", `The defined name "${name}" belongs to the sheet "${sheet.slice(0, 100)}", which the workbook does not have.`, {
		location: { sheet: sheet.slice(0, 100) },
		hint: `Add the sheet first, or use one of: ${sheets.size > 0 ? [...sheets].map((s) => `"${s}"`).join(", ") : "(no sheets yet)"}.`,
		details: {
			name,
			available: [...sheets]
		}
	});
}
/**
* The key of a defined name in its scope: names are not case-sensitive (P1 §18.17.2.5) and are
* unique per sheet for local names and among workbook-level names otherwise ([MS-OI29500]
* §2.1.568 n). A name has no NUL character, so the last one separates the scope from the name.
*/
function nameScopeKey(n) {
	return `${n.sheet === null ? "w" : `s${n.sheet}`}\u0000${n.name.toLowerCase()}`;
}
/**
* Throws `XLV_E_INVALID_DEFINED_NAME` if `added` has the same name and scope as one of
* `existing`, which maps {@link nameScopeKey} to the first name with that key (D-H01: a map,
* not a search of every name for each one added).
*/
function checkUniqueName(added, existing) {
	const clash = existing.get(nameScopeKey(added));
	if (clash === void 0) return;
	throw invalidName(added.name, `The ${added.sheet === null ? "workbook" : `sheet "${added.sheet}"`} already has a defined name "${clash.name}".`, "Names are not case-sensitive and must be unique in their scope. Choose another name, or make it local to a sheet with { sheet: \"…\" }.");
}
/** One `definedName` element. `localSheetId` is the 0-based index of the name's sheet. */
function definedNameXml(n, localSheetId) {
	let xml = `<definedName name="${escapeAttr(encodeXstring(n.name))}"`;
	if (n.comment !== null) xml += ` comment="${escapeAttr(encodeXstring(n.comment))}"`;
	if (localSheetId !== null) xml += ` localSheetId="${String(localSheetId)}"`;
	if (n.hidden) xml += " hidden=\"1\"";
	return `${xml}>${escapeText(encodeXstring(n.refersTo))}</definedName>`;
}
// src/tables.ts
/** Content type of a Table Definition part (Part 1 §12.3.21, TD-01). */
const TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
/** SUBTOTAL function numbers by totals function ([MS-OI29500] §2.1.1094, TD-06). */
const SUBTOTAL = {
	average: 101,
	countNums: 102,
	count: 103,
	max: 104,
	min: 105,
	stdDev: 107,
	sum: 109,
	var: 110
};
const FUNCTIONS = [
	"none",
	"sum",
	"min",
	"max",
	"average",
	"count",
	"countNums",
	"stdDev",
	"var",
	"custom"
];
/** Built-in table style names of Part 1 Annex G.1 / presetTableStyles.xml (TD-09). */
const THEME_RE = /^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/;
/** [MS-OI29500] §2.1.680 b, §2.1.681 a, j (TD-04). */
const MAX_NAME = 255;
const MAX_LABEL = 32767;
const MAX_FORMULA$1 = 8192;
/** A sheet holds at most this many tables in the reader (an implementation limit). */
const MAX_TABLES_PER_SHEET = 4096;
function invalidTable(sheet, message, hint, details = {}) {
	return new XlvError("XLV_E_INVALID_TABLE", message, {
		location: { sheet },
		hint,
		details
	});
}
/** Checks a table or display name (P1 §18.5.1.2, §18.17.2.5; [MS-OI29500] §2.1.680 b; TD-03). */
function checkTableName(value, what, sheet) {
	const hint = "Use letters, digits, \"_\" and \".\", starting with a letter or \"_\", for example \"Sales_2026\"; no spaces.";
	if (typeof value !== "string" || value.length < 1 || value.length > MAX_NAME) throw invalidTable(sheet, `The table ${what} must be text of 1 to ${String(MAX_NAME)} characters.`, hint, { name: String(value).slice(0, 40) });
	const upper = value.toUpperCase();
	if (!NAME_RE.test(value) || firstLoneSurrogate(value) !== -1 || upper === "TRUE" || upper === "FALSE" || A1_NAME_RE.test(value) || R1C1_NAME_RE.test(value)) throw invalidTable(sheet, `"${value.slice(0, 40)}" is not a valid table ${what}.`, `${hint} It must not read as a cell reference such as "A1" or "R1C1".`, { name: value.slice(0, 40) });
	if (BUILT_IN_NAMES.has(upper.replace(/^_XLNM\./, "")) || upper.startsWith("_XL")) throw invalidTable(sheet, `"${value.slice(0, 40)}" is a built-in name of Excel and cannot name a table.`, `${hint} Built-in names such as Print_Area, Print_Titles, Criteria or Database are reserved.`, { name: value.slice(0, 40) });
	return value;
}
/** The built-in defined names of ECMA-376 Part 1 §18.2.5 without the `_xlnm.` prefix (W2-06), upper case. */
const BUILT_IN_NAMES = /* @__PURE__ */ new Set([
	"PRINT_AREA",
	"PRINT_TITLES",
	"CRITERIA",
	"_FILTERDATABASE",
	"EXTRACT",
	"CONSOLIDATE_AREA",
	"DATABASE",
	"SHEET_TITLE"
]);
function text(value, sheet, what, max) {
	if (typeof value !== "string" || value.length > max || firstLoneSurrogate(value) !== -1) throw invalidTable(sheet, `The ${what} must be text of at most ${String(max)} characters.`, `Shorten the ${what}.`);
	return value;
}
function formula(value, sheet, what) {
	const f = typeof value === "string" && value.startsWith("=") ? value.slice(1) : value;
	if (typeof f !== "string" || f.trim() === "" || f.length > MAX_FORMULA$1 || firstLoneSurrogate(f) !== -1) throw invalidTable(sheet, `The ${what} must be a formula of at most ${String(MAX_FORMULA$1)} characters.`, "Pass the formula text with or without \"=\", for example \"SUM(Sales[Units])\".");
	return f;
}
function bool(o, key, def, sheet) {
	const v = o[key];
	if (v === void 0) return def;
	if (typeof v !== "boolean") throw invalidTable(sheet, `The table option "${key}" must be true or false.`, `Use ${key}: true or leave it out.`, { option: key });
	return v;
}
function known(o, keys, what, sheet) {
	if (typeof o !== "object" || o === null || Array.isArray(o)) throw invalidTable(sheet, `The ${what} must be an object.`, `For example { ${keys[0] ?? ""}: … }.`);
	for (const key of Object.keys(o)) if (!keys.includes(key)) throw invalidTable(sheet, `Unknown ${what} option "${key.slice(0, 40)}".`, `Known options: ${keys.join(", ")}.`, { option: key.slice(0, 40) });
	return o;
}
/**
* Checks a table given to `NewSheet.addTable` (or the streaming writer, `streaming = true`: then
* `ref` must be the whole range and `rows` must not be given). Names, columns and the range are
* checked here; uniqueness and overlaps are checked by the caller.
*/
function normalizeTable(input, sheet, streaming = false) {
	const o = known(input, [
		"name",
		"displayName",
		"ref",
		"headerRow",
		"totalsRow",
		"style",
		"columns",
		"rows"
	], "table", sheet);
	const name = checkTableName(o["name"], "name", sheet);
	const displayName = o["displayName"] === void 0 ? name : checkTableName(o["displayName"], "display name", sheet);
	const headerRow = bool(o, "headerRow", true, sheet);
	const totalsRow = bool(o, "totalsRow", false, sheet);
	const s = o["style"] === void 0 ? {} : known(o["style"], [
		"theme",
		"showFirstColumn",
		"showLastColumn",
		"showRowStripes",
		"showColumnStripes"
	], "table style", sheet);
	const theme = s["theme"] === void 0 ? "TableStyleMedium2" : s["theme"];
	if (theme !== null && (typeof theme !== "string" || !THEME_RE.test(theme))) throw invalidTable(sheet, `"${String(theme).slice(0, 40)}" is not a built-in table style.`, "Use TableStyleLight1 … TableStyleLight21, TableStyleMedium1 … TableStyleMedium28, TableStyleDark1 … TableStyleDark11, or null.", { option: "theme" });
	const style = {
		theme,
		showFirstColumn: bool(s, "showFirstColumn", false, sheet),
		showLastColumn: bool(s, "showLastColumn", false, sheet),
		showRowStripes: bool(s, "showRowStripes", false, sheet),
		showColumnStripes: bool(s, "showColumnStripes", false, sheet)
	};
	const rawColumns = o["columns"];
	if (!Array.isArray(rawColumns) || rawColumns.length === 0 || rawColumns.length > 16384) throw invalidTable(sheet, `The table "${name}" needs a list of 1 to ${String(MAX_COLUMNS)} columns.`, "For example columns: [{ name: 'Region' }, { name: 'Revenue' }].", { name });
	const seen = /* @__PURE__ */ new Set();
	const columns = rawColumns.map((raw, i) => {
		const c = known(raw, [
			"name",
			"totalsRowFunction",
			"totalsRowLabel",
			"totalsRowFormula",
			"calculatedColumnFormula",
			"filterButton"
		], `column ${String(i + 1)} of table "${name}"`, sheet);
		const columnName = c["name"];
		if (typeof columnName !== "string" || columnName.length < 1 || columnName.length > MAX_NAME || columnName.trim() === "" || firstLoneSurrogate(columnName) !== -1 || /[\r\n]/.test(columnName)) throw invalidTable(sheet, `Column ${String(i + 1)} of table "${name}" needs a name of 1 to ${String(MAX_NAME)} characters on one line.`, "Give every column a short, unique name, for example \"Revenue\".", {
			name,
			column: i + 1
		});
		const key = columnName.toLowerCase();
		if (seen.has(key)) throw invalidTable(sheet, `Table "${name}" has two columns named "${columnName.slice(0, 40)}".`, "Column names must be unique in a table (case is ignored), for example \"Revenue\" and \"Revenue_2\".", {
			name,
			column: i + 1
		});
		seen.add(key);
		const fn = c["totalsRowFunction"] ?? "none";
		if (typeof fn !== "string" || !FUNCTIONS.includes(fn)) throw invalidTable(sheet, `"${String(fn).slice(0, 20)}" is not a totals function.`, `Use one of ${FUNCTIONS.join(", ")}.`, {
			name,
			column: i + 1
		});
		const label = c["totalsRowLabel"] === void 0 ? null : text(c["totalsRowLabel"], sheet, `totals label of column "${columnName}"`, MAX_LABEL);
		if (label !== null && fn !== "none") throw invalidTable(sheet, `Column "${columnName}" of table "${name}" has both a totals label and the function "${fn}".`, "A totals cell shows either a label (with totalsRowFunction \"none\") or a function, not both (ECMA-376 Part 1 §18.5.1.3).", {
			name,
			column: i + 1
		});
		const totalsFormula = c["totalsRowFormula"] === void 0 ? null : formula(c["totalsRowFormula"], sheet, `totals formula of column "${columnName}"`);
		if (fn === "custom" !== (totalsFormula !== null)) throw invalidTable(sheet, `Column "${columnName}" of table "${name}": totalsRowFormula goes with totalsRowFunction "custom" and only with it.`, "For example { totalsRowFunction: 'custom', totalsRowFormula: 'SUM(Sales[Units])*2' }.", {
			name,
			column: i + 1
		});
		return {
			name: columnName,
			totalsRowFunction: fn,
			totalsRowLabel: label,
			totalsRowFormula: totalsFormula,
			calculatedColumnFormula: c["calculatedColumnFormula"] === void 0 ? null : formula(c["calculatedColumnFormula"], sheet, `calculated formula of column "${columnName}"`),
			filterButton: bool(c, "filterButton", true, sheet)
		};
	});
	const rawRows = o["rows"] ?? [];
	if (!Array.isArray(rawRows) || rawRows.some((r) => !Array.isArray(r))) throw invalidTable(sheet, `The rows of table "${name}" must be an array of arrays.`, "For example rows: [['North', 1200], ['South', 980]].", { name });
	if (streaming && rawRows.length > 0) throw invalidTable(sheet, `The streaming writer takes no rows for table "${name}".`, "Stream the header, data and totals rows as ordinary rows of the sheet and give the whole range in ref.", { name });
	for (const [i, r] of rawRows.entries()) if (r.length > columns.length) throw invalidTable(sheet, `Row ${String(i + 1)} of table "${name}" has ${String(r.length)} values for ${String(columns.length)} columns.`, "Give at most one value per column.", {
		name,
		row: i + 1
	});
	const rows = rawRows;
	const rawRef = o["ref"];
	const given = typeof rawRef === "string" ? parseRange(rawRef) : null;
	if (given === null) throw invalidTable(sheet, `The ref of table "${name}" must be a cell such as "B2" or a range such as "B2:D10".`, "Give the top-left cell, or the whole range including the header and totals rows.", { name });
	const single = !String(rawRef).includes(":");
	if (streaming && single) throw invalidTable(sheet, `The streaming writer needs the whole range of table "${name}", not only its top-left cell.`, "For example ref: \"A1:C101\" for a header row and 100 data rows.", { name });
	const extra = (headerRow ? 1 : 0) + (totalsRow ? 1 : 0);
	let range;
	if (single) {
		const height = extra + Math.max(rows.length, 1);
		range = {
			start: given.start,
			end: {
				row: given.start.row + height - 1,
				col: given.start.col + columns.length - 1
			}
		};
		if (range.end.row > 1048576 || range.end.col > 16384) throw invalidTable(sheet, `Table "${name}" at ${String(rawRef)} would reach beyond the last row or column of the sheet.`, "Start the table higher or further left, or use fewer rows or columns.", { name });
	} else {
		range = given;
		const width = range.end.col - range.start.col + 1;
		if (width !== columns.length) throw invalidTable(sheet, `The range ${formatRange(range)} of table "${name}" is ${String(width)} columns wide, but ${String(columns.length)} columns are given.`, "Give one column for each column of the range, or only the top-left cell.", { name });
		const dataRows = range.end.row - range.start.row + 1 - extra;
		if (dataRows < 1) throw invalidTable(sheet, `The range ${formatRange(range)} of table "${name}" leaves no data row.`, "A table needs at least one data row between its header and totals rows.", { name });
		if (rows.length > dataRows) throw invalidTable(sheet, `Table "${name}" has ${String(rows.length)} rows, but its range ${formatRange(range)} holds ${String(dataRows)}.`, "Make the range larger, or give only the top-left cell.", { name });
	}
	return {
		name,
		displayName,
		range,
		headerRow,
		totalsRow,
		style,
		columns,
		rows
	};
}
/** True when two ranges share a cell. */
function rangesOverlap(a, b) {
	return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}
/**
* A structured reference to a whole column of a table, as Excel writes it in totals rows (TD-07):
* `Sales[Units]`; names with a leading or trailing space or with `'`, `[`, `]`, `#` get double
* brackets, and those four characters are escaped with `'` ([MS-OI29500] §3.2.3.1.1).
*/
function columnReference(table, column) {
	const escaped = column.replace(/['[\]#]/g, "'$&");
	return escaped !== column || column.trim() !== column ? `${table}[[${escaped}]]` : `${table}[${column}]`;
}
/** The formula of a column's totals cell, or `null` when it has none. */
function totalsFormula(table, column) {
	if (column.totalsRowFunction === "custom") return column.totalsRowFormula;
	const n = SUBTOTAL[column.totalsRowFunction];
	return n === void 0 ? null : `SUBTOTAL(${String(n)},${columnReference(table, column.name)})`;
}
const a = (name, value) => ` ${name}="${typeof value === "number" ? String(value) : escapeAttr(value)}"`;
/** The Table Definition part of a table (CT_Table; element order of sml.xsd T3946; TD-02). */
function tableXml(t, id) {
	const ref = `${formatCellRef(t.range.start.row, t.range.start.col)}:${formatCellRef(t.range.end.row, t.range.end.col)}`;
	let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${SML_NS.transitional}"${a("id", id)}${a("name", encodeXstring(t.name))}${a("displayName", encodeXstring(t.displayName))}${a("ref", ref)}`;
	if (!t.headerRow) xml += a("headerRowCount", 0);
	xml += t.totalsRow ? a("totalsRowCount", 1) : a("totalsRowShown", 0);
	xml += ">";
	if (t.headerRow) {
		const filterRange = {
			start: t.range.start,
			end: {
				row: t.range.end.row - (t.totalsRow ? 1 : 0),
				col: t.range.end.col
			}
		};
		const hidden = t.columns.map((c, i) => c.filterButton ? "" : `<filterColumn${a("colId", i)} hiddenButton="1"/>`).join("");
		xml += `<autoFilter${a("ref", formatRange(filterRange))}${hidden === "" ? "/>" : `>${hidden}</autoFilter>`}`;
	}
	xml += `<tableColumns${a("count", t.columns.length)}>`;
	t.columns.forEach((c, i) => {
		xml += `<tableColumn${a("id", i + 1)}${a("name", encodeXstring(c.name))}`;
		if (t.totalsRow && c.totalsRowFunction !== "none") xml += a("totalsRowFunction", c.totalsRowFunction);
		if (t.totalsRow && c.totalsRowLabel !== null) xml += a("totalsRowLabel", encodeXstring(c.totalsRowLabel));
		let children = "";
		if (c.calculatedColumnFormula !== null) children += `<calculatedColumnFormula>${escapeText(encodeXstring(c.calculatedColumnFormula))}</calculatedColumnFormula>`;
		if (t.totalsRow && c.totalsRowFormula !== null) children += `<totalsRowFormula>${escapeText(encodeXstring(c.totalsRowFormula))}</totalsRowFormula>`;
		xml += children === "" ? "/>" : `>${children}</tableColumn>`;
	});
	xml += "</tableColumns>";
	const s = t.style;
	xml += `<tableStyleInfo${s.theme === null ? "" : a("name", s.theme)}${a("showFirstColumn", s.showFirstColumn ? 1 : 0)}${a("showLastColumn", s.showLastColumn ? 1 : 0)}${a("showRowStripes", s.showRowStripes ? 1 : 0)}${a("showColumnStripes", s.showColumnStripes ? 1 : 0)}/>`;
	return `${xml}</table>`;
}
/** Part name of the n-th table of a workbook (1-based), as Excel names them. */
function tablePartName(n) {
	return `xl/tables/table${String(n)}.xml`;
}
function boolAttr(el, name, def) {
	const v = el.attrs.get(name)?.trim();
	return v === void 0 ? def : v === "1" || v === "true";
}
function uintAttr(el, name, def) {
	const v = el.attrs.get(name)?.trim();
	return v !== void 0 && /^\d{1,10}$/.test(v) ? Number(v) : def;
}
/**
* Reads a Table Definition part. Returns `null` with a warning when the part is not a table or
* its range is not valid; unknown attributes and elements are ignored.
*/
function parseTable(xml, ctx) {
	const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, {
		part: ctx.part,
		sheet: ctx.sheet
	});
	let root;
	try {
		root = parseXmlTree(xml, ctx);
	} catch (error) {
		if (error instanceof XlvError && error.code === "XLV_E_XML_MALFORMED") {
			warn("A table part of the sheet is not well-formed XML; the table is ignored.");
			return null;
		}
		throw error;
	}
	const ns = root.ns;
	if (root.local !== "table" || ns !== SML_NS.strict && ns !== SML_NS.transitional) {
		warn(`A table part of the sheet has the root element <${root.local}>; it is ignored.`);
		return null;
	}
	const range = parseRange((root.attrs.get("ref") ?? "").trim().toUpperCase());
	if (range === null) {
		warn(`A table of the sheet has the invalid range "${(root.attrs.get("ref") ?? "").slice(0, 40)}"; it is ignored.`);
		return null;
	}
	const displayName = decodeXstring(root.attrs.get("displayName") ?? "");
	const objectName = decodeXstring(root.attrs.get("name") ?? "");
	const name = objectName !== "" ? objectName : displayName;
	const autoFilterEl = childNamed(root, ns, "autoFilter");
	const filterRef = autoFilterEl?.attrs.get("ref")?.trim().toUpperCase();
	const filterRange = filterRef === void 0 ? null : parseRange(filterRef);
	const hidden = /* @__PURE__ */ new Set();
	if (autoFilterEl) {
		for (const fc of childrenNamed(autoFilterEl, ns, "filterColumn")) if (boolAttr(fc, "hiddenButton", false) || !boolAttr(fc, "showButton", true)) hidden.add(uintAttr(fc, "colId", -1));
	}
	const columnsEl = childNamed(root, ns, "tableColumns");
	const columns = (columnsEl === void 0 ? [] : childrenNamed(columnsEl, ns, "tableColumn")).slice(0, MAX_COLUMNS).map((c, i) => {
		const fn = c.attrs.get("totalsRowFunction")?.trim() ?? "none";
		const label = c.attrs.get("totalsRowLabel");
		const calc = childNamed(c, ns, "calculatedColumnFormula");
		const total = childNamed(c, ns, "totalsRowFormula");
		return Object.freeze({
			id: uintAttr(c, "id", i + 1),
			name: decodeXstring(c.attrs.get("name") ?? ""),
			totalsRowFunction: FUNCTIONS.includes(fn) ? fn : "none",
			totalsRowLabel: label === void 0 ? null : decodeXstring(label),
			totalsRowFormula: total === void 0 ? null : decodeXstring(total.text),
			calculatedColumnFormula: calc === void 0 ? null : decodeXstring(calc.text),
			filterButton: !hidden.has(i)
		});
	});
	const styleEl = childNamed(root, ns, "tableStyleInfo");
	const type = root.attrs.get("tableType")?.trim();
	const comment = root.attrs.get("comment");
	return Object.freeze({
		name,
		displayName: displayName !== "" ? displayName : name,
		ref: formatRange(range),
		headerRowCount: Math.min(uintAttr(root, "headerRowCount", 1), 1),
		totalsRowCount: Math.min(uintAttr(root, "totalsRowCount", 0), 1),
		totalsRowShown: boolAttr(root, "totalsRowShown", true),
		tableType: type === "xml" || type === "queryTable" ? type : "worksheet",
		comment: comment === void 0 ? null : decodeXstring(comment),
		columns,
		style: styleEl === void 0 ? null : Object.freeze({
			theme: styleEl.attrs.has("name") ? decodeXstring(styleEl.attrs.get("name") ?? "") : null,
			showFirstColumn: boolAttr(styleEl, "showFirstColumn", false),
			showLastColumn: boolAttr(styleEl, "showLastColumn", false),
			showRowStripes: boolAttr(styleEl, "showRowStripes", false),
			showColumnStripes: boolAttr(styleEl, "showColumnStripes", false)
		}),
		autoFilter: filterRange === null ? null : formatRange(filterRange),
		part: ctx.part
	});
}
// src/opc/package.ts
/** ZIP item name of the Media Types stream (Part 2 §7.3.7, F-OPC-09). */
const CONTENT_TYPES_ITEM = "[content_types].xml";
/** Used in error locations; the stream is not a part, but this is where users find it. */
const CONTENT_TYPES_LOCATION = "/[Content_Types].xml";
/** Media type of Relationships parts (Part 2 §6.5.2.1, F-OPC-07). */
const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const PACKAGE_RELATIONSHIPS = "/_rels/.rels";
const INVALID_HINT = "The package structure is damaged or was written by a program that does not follow ECMA-376 Part 2. Re-save the file in a spreadsheet application; if a program generated it, report the problem to its authors.";
const EMPTY = Object.freeze([]);
function invalid$1(part, message, reason, extra = {}) {
	return new XlvError("XLV_E_PART_INVALID", message, {
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
	if (value === void 0) throw invalid$1(part, `A ${element.local} element has no ${name} attribute.`, "missing-attribute", {
		element: element.local,
		attribute: name
	});
	return value;
}
/**
* Returns the name of the Relationships part for `source`: "/_rels/.rels" for the package (null),
* otherwise "_rels/" inserted before the last segment and ".rels" appended
* (Part 2 §6.5.2.2 and §6.5.2.3, F-OPC-07). "/xl/workbook.xml" gives "/xl/_rels/workbook.xml.rels".
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
* Throws `XLV_E_PART_MISSING` without `[Content_Types].xml`, `XLV_E_PART_INVALID` for two items
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
		if (previous !== void 0) throw invalid$1(previous.name, `ZIP items "${clip(previous.entry.name)}" and "${clip(entry.name)}" both map to part ${clip(previous.name)}.`, "equivalent-part-names", { items: [clip(previous.entry.name), clip(entry.name)] });
		parts.set(key, {
			name,
			entry
		});
		partNames.push(name);
	}
	if (contentTypes === void 0) throw new XlvError("XLV_E_PART_MISSING", interleaved ? "The package stores [Content_Types].xml in interleaved pieces, which Office Open XML documents must not use." : "The package has no [Content_Types].xml, so it is not an Office Open XML package.", {
		location: { part: CONTENT_TYPES_LOCATION },
		hint: "Check that the file is an .xlsx workbook and not another ZIP-based format (such as .ods or .zip). If it is, re-save it in a spreadsheet application.",
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
	if (root.ns !== "http://schemas.openxmlformats.org/package/2006/content-types" || root.local !== "Types") throw invalid$1(part, `The root element of [Content_Types].xml is not Types in namespace ${CONTENT_TYPES_NS}.`, "root-element", {
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
			if (defaults.has(key)) throw invalid$1(part, `[Content_Types].xml has more than one Default for extension "${clip(extension)}".`, "duplicate-default", { extension: clip(extension) });
			defaults.set(key, type);
		} else if (child.local === "Override") {
			const partName = requireAttribute(child, "PartName", part);
			const type = requireAttribute(child, "ContentType", part);
			const key = partKey(canonicalReference(partName));
			if (overrides.has(key)) throw invalid$1(part, `[Content_Types].xml has more than one Override for part "${clip(partName)}".`, "duplicate-override", { partName: clip(partName) });
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
		if (root.ns !== "http://schemas.openxmlformats.org/package/2006/relationships" || root.local !== "Relationships") throw invalid$1(part, `The root element of ${clip(part)} is not Relationships in namespace ${PACKAGE_REL_NS}.`, "root-element", {
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
			if (ids.has(id)) throw invalid$1(part, `Relationship Id "${clip(id)}" is used more than once.`, "duplicate-id", { id: clip(id) });
			ids.add(id);
			if (mode !== void 0 && mode !== "Internal" && mode !== "External") throw invalid$1(part, `Relationship "${clip(id)}" has TargetMode "${clip(mode)}"; expected Internal or External.`, "target-mode", {
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
			if (resolved === null) throw invalid$1(part, `Relationship "${clip(id)}" has the internal target "${clip(target)}", which is not a relative reference to a part.`, "internal-target", {
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
// src/sml/shared-strings.ts
const SI_DEPTH = 2;
const ITEM_CHILD_DEPTH = 3;
const RUN_TEXT_DEPTH = 4;
const RUN_PROPERTY_DEPTH = 5;
/**
* Parses the shared strings part and returns the plain text of every string item (`si`), in
* order, so that index `i` is the value of a cell with `t="s"` and `<v>i</v>`. When `rich` is
* given, the runs of every rich text item are stored in it under the item's index.
* Throws `XLV_E_PART_INVALID` if the root element is not `sst` in a SpreadsheetML namespace,
* and the XML parser's errors for malformed XML.
*/
function parseSharedStrings(xml, ctx, rich) {
	const strings = [];
	const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, { part: ctx.part });
	let ns = "";
	let depth = 0;
	let inItem = false;
	let inRun = false;
	/** Depth of the open t element whose text is collected, or 0. */
	let textDepth = 0;
	let item = [];
	let raw = [];
	let runs = [];
	let hasRuns = false;
	let runText = "";
	let runProperties = null;
	let runFont;
	parseXml(xml, {
		start(elementNs, local, attrs) {
			depth++;
			if (depth === 1) {
				if (local !== "sst" || elementNs !== SML_NS.strict && elementNs !== SML_NS.transitional) throw new XlvError("XLV_E_PART_INVALID", `The shared strings part has the root element <${local}> in namespace "${elementNs}"; expected <sst> in the SpreadsheetML namespace.`, {
					location: { part: ctx.part },
					hint: "The part the workbook names as its shared string table is not one. Re-save the file in a spreadsheet application; if it was generated by a program, report the problem to its authors.",
					details: {
						element: local,
						namespace: elementNs
					}
				});
				ns = elementNs;
				return;
			}
			if (elementNs !== ns) return;
			if (depth === SI_DEPTH) {
				if (local === "si") {
					inItem = true;
					item = [];
					runs = [];
					hasRuns = false;
				}
				return;
			}
			if (!inItem) return;
			if (depth === ITEM_CHILD_DEPTH) {
				if (local === "t") textDepth = depth;
				else if (local === "r") {
					inRun = true;
					runText = "";
					runFont = void 0;
				}
			} else if (depth === RUN_TEXT_DEPTH && inRun) {
				if (local === "t") textDepth = depth;
				else if (local === "rPr" && rich !== void 0) runProperties = [];
			} else if (depth === RUN_PROPERTY_DEPTH && runProperties !== null) runProperties.push({
				ns: elementNs,
				local,
				attrs,
				children: [],
				text: ""
			});
		},
		end(_ns, local) {
			if (textDepth !== 0 && depth === textDepth) {
				const text = decodeXstring(raw.length === 1 ? raw[0] ?? "" : raw.join(""));
				item.push(text);
				if (inRun) runText += text;
				else if (rich !== void 0) runs.push({ text });
				raw = [];
				textDepth = 0;
			} else if (depth === RUN_TEXT_DEPTH && runProperties !== null && local === "rPr") {
				const font = parseRunFont({
					ns,
					local,
					attrs: /* @__PURE__ */ new Map(),
					children: runProperties,
					text: ""
				}, ns, warn);
				if (Object.keys(font).length > 0) runFont = font;
				runProperties = null;
			} else if (depth === ITEM_CHILD_DEPTH && inRun) {
				if (local === "r" && rich !== void 0) {
					runs.push(runFont === void 0 ? { text: runText } : {
						text: runText,
						font: runFont
					});
					hasRuns = true;
				}
				inRun = false;
			} else if (depth === SI_DEPTH && inItem) {
				if (rich !== void 0 && hasRuns) rich.set(strings.length, Object.freeze(runs.map((r) => Object.freeze(r))));
				strings.push(item.length === 1 ? item[0] ?? "" : item.join(""));
				inItem = false;
			}
			depth--;
		},
		text(value) {
			if (textDepth !== 0 && depth === textDepth) raw.push(value);
		}
	}, ctx);
	return strings;
}
// src/sml/workbook.ts
const SHEET_STATES = [
	"visible",
	"hidden",
	"veryHidden"
];
/** Parses an `xsd:boolean` lexical value; `null` if it is not one. */
function parseXsdBoolean(value) {
	const v = value.trim();
	if (v === "true" || v === "1") return true;
	if (v === "false" || v === "0") return false;
	return null;
}
/** Parses the workbook part. */
function parseWorkbookPart(xml, ctx) {
	const root = parseXmlTree(xml, ctx);
	let conformance;
	if (root.local === "workbook" && root.ns === SML_NS.strict) conformance = "strict";
	else if (root.local === "workbook" && root.ns === SML_NS.transitional) conformance = "transitional";
	else throw new XlvError("XLV_E_PART_INVALID", `The workbook part has the root element <${root.local}> in namespace "${root.ns}".`, {
		location: { part: ctx.part },
		hint: "The main part of the package is not a SpreadsheetML workbook. Check that the file is an Excel workbook (.xlsx/.xlsm), not a Word or PowerPoint file.",
		details: {
			element: root.local,
			namespace: root.ns
		}
	});
	const ns = root.ns;
	const relIdKey = `{${REL_NS[conformance]}}id`;
	let date1904 = false;
	const rawDate1904 = childNamed(root, ns, "workbookPr")?.attrs.get("date1904");
	if (rawDate1904 !== void 0) {
		const parsed = parseXsdBoolean(rawDate1904);
		if (parsed === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `workbookPr@date1904 has the invalid value "${clip(rawDate1904)}"; using the 1900 date system.`, { part: ctx.part });
		else date1904 = parsed;
	}
	const sheetsElement = childNamed(root, ns, "sheets");
	if (sheetsElement === void 0) throw new XlvError("XLV_E_PART_INVALID", "The workbook part has no <sheets> element.", {
		location: { part: ctx.part },
		hint: "Every workbook must list its sheets. Re-save the file in a spreadsheet application."
	});
	const seenIds = /* @__PURE__ */ new Set();
	const sheets = childrenNamed(sheetsElement, ns, "sheet").map((el) => {
		const name = decodeXstring(el.attrs.get("name") ?? "");
		const rawState = el.attrs.get("state") ?? "visible";
		let state = SHEET_STATES.find((s) => s === rawState);
		if (state === void 0) {
			ctx.warnings.add("XLV_W_INVALID_VALUE", `Sheet "${name}" has the unknown state "${clip(rawState)}"; treating it as visible.`, {
				part: ctx.part,
				sheet: name
			});
			state = "visible";
		}
		const rawId = el.attrs.get("sheetId");
		let sheetId = rawId !== void 0 && /^\d{1,10}$/.test(rawId.trim()) ? Number(rawId.trim()) : null;
		if (sheetId !== null && sheetId > 4294967295) sheetId = null;
		if (sheetId === null) ctx.warnings.add("XLV_W_INVALID_VALUE", rawId === void 0 ? `Sheet "${name}" has no sheetId.` : `Sheet "${name}" has the invalid sheetId "${rawId.slice(0, 20)}".`, {
			part: ctx.part,
			sheet: name
		});
		else if (seenIds.has(sheetId)) {
			ctx.warnings.add("XLV_W_INVALID_VALUE", `Sheet "${name}" repeats the sheetId ${String(sheetId)} of an earlier sheet; its sheetId is read as null.`, {
				part: ctx.part,
				sheet: name
			});
			sheetId = null;
		} else seenIds.add(sheetId);
		return {
			name,
			sheetId,
			state,
			relId: el.attrs.get(relIdKey) ?? null
		};
	});
	const namesElement = childNamed(root, ns, "definedNames");
	const definedNames = (namesElement === void 0 ? [] : childrenNamed(namesElement, ns, "definedName")).map((el) => {
		const name = decodeXstring(el.attrs.get("name") ?? "");
		const rawId = el.attrs.get("localSheetId");
		let localSheetId = null;
		if (rawId !== void 0) {
			if (/^\d{1,10}$/.test(rawId.trim())) localSheetId = Number(rawId.trim());
			else ctx.warnings.add("XLV_W_INVALID_VALUE", `The defined name "${name.slice(0, 100)}" has the invalid localSheetId "${rawId.slice(0, 20)}"; it is read as a workbook-level name.`, { part: ctx.part });
		}
		const rawHidden = el.attrs.get("hidden");
		const comment = el.attrs.get("comment");
		return {
			name,
			refersTo: decodeXstring(el.text),
			localSheetId,
			hidden: rawHidden !== void 0 && parseXsdBoolean(rawHidden) === true,
			comment: comment === void 0 ? null : decodeXstring(comment)
		};
	});
	const protectionElement = childNamed(root, ns, "workbookProtection");
	const protection = protectionElement === void 0 ? null : workbookProtectionOf(protectionElement.attrs, (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, { part: ctx.part }));
	const views = [];
	const bookViews = childNamed(root, ns, "bookViews");
	for (const v of bookViews?.children ?? []) {
		if (v.local !== "workbookView" || v.ns !== ns || views.length >= 1e3) continue;
		const num = (name, signed) => {
			const raw = v.attrs.get(name)?.trim();
			if (raw === void 0 || !/^[+-]?\d{1,10}$/.test(raw)) return null;
			const n = Number(raw);
			return signed ? n >= -2147483648 && n <= 2147483647 ? n : null : n >= 0 && n <= 4294967295 ? n : null;
		};
		const visibility = v.attrs.get("visibility");
		views.push(Object.freeze({
			x: num("xWindow", true),
			y: num("yWindow", true),
			width: num("windowWidth", false),
			height: num("windowHeight", false),
			firstSheet: num("firstSheet", false),
			activeTab: num("activeTab", false),
			visibility: visibility === "hidden" || visibility === "veryHidden" ? visibility : "visible"
		}));
	}
	const calc = childNamed(root, ns, "calcPr");
	const rawFull = calc?.attrs.get("fullCalcOnLoad");
	const calculation = calc === void 0 ? null : Object.freeze({ fullCalcOnLoad: rawFull !== void 0 && parseXsdBoolean(rawFull) === true });
	return {
		conformance,
		date1904,
		sheets,
		definedNames,
		protection,
		views,
		calculation
	};
}
// src/sml/dates.ts
const DAY_MS$1 = 864e5;
/** Serial 0 of the Strict 1900 system, and of the Transitional one from 1900-03-01 onwards. */
const EPOCH_1899_12_30 = Date.UTC(1899, 11, 30);
/** Serial 0 of the Transitional 1900 system for 1900-01-01 to 1900-02-28 (Part 4 §15.3.4). */
const EPOCH_1899_12_31 = Date.UTC(1899, 11, 31);
/** Serial 0 of the 1904 system (Part 1 §18.17.4.1). */
const EPOCH_1904_01_01 = Date.UTC(1904, 0, 1);
const TIME_1900_03_01 = Date.UTC(1900, 2, 1);
/** Earliest and latest permitted instants (Part 1 §18.17.4 and §18.17.4.1). */
const MIN_TIME = (/* @__PURE__ */ new Date(0)).setUTCFullYear(1, 0, 1);
const MAX_TIME = Date.UTC(9999, 11, 31, 23, 59, 59, 999);
/** No date system reaches beyond these serials; checked before multiplying, to keep precision. */
const MIN_ANY_SERIAL = -695056;
const MAX_ANY_SERIAL = 2958467;
/**
* Converts a serial date-time to a `Date` whose UTC fields are the spreadsheet's wall-clock date
* and time. The time of day is rounded to the nearest millisecond, and the result is `null` when
* the rounded value is outside the date system's valid range (ECMA-376 Part 1 §18.17.4.1, or
* Part 4 §15.3.4 for Transitional documents) or is the Transitional serial 60 (the nonexistent
* 29 February 1900). `conformance` selects the 1900 date system's variant; it defaults to
* `'transitional'`, the class of documents without `workbook@conformance` (Part 1 §18.2.27).
*/
function serialToDate(serial, system, conformance = "transitional") {
	if (!Number.isFinite(serial) || serial < MIN_ANY_SERIAL || serial > MAX_ANY_SERIAL) return null;
	const ms = Math.round(serial * DAY_MS$1);
	let time;
	if (system === "1904") time = EPOCH_1904_01_01 + ms;
	else if (conformance === "strict") time = EPOCH_1899_12_30 + ms;
	else {
		if (ms < 0) return null;
		if (ms < 60 * DAY_MS$1) time = EPOCH_1899_12_31 + ms;
		else if (ms < 61 * DAY_MS$1) return null;
		else time = EPOCH_1899_12_30 + ms;
	}
	if (time < MIN_TIME || time > MAX_TIME) return null;
	return new Date(time);
}
/**
* Converts a `Date` (read through its UTC fields as wall-clock time) to a serial date-time in the
* given date system; `conformance` defaults to `'transitional'` as in {@link serialToDate}.
* Throws `RangeError` for an invalid `Date` or one outside the date system's valid range.
*/
function dateToSerial(date, system, conformance = "transitional") {
	const time = date.getTime();
	if (Number.isNaN(time)) throw new RangeError("Cannot convert an invalid Date to a serial date-time.");
	if (time < MIN_TIME || time > MAX_TIME) throw new RangeError(`Date ${date.toISOString()} is outside 0001-01-01T00:00:00.000 to 9999-12-31T23:59:59.999.`);
	if (system === "1904") return (time - EPOCH_1904_01_01) / DAY_MS$1;
	if (conformance === "strict") return (time - EPOCH_1899_12_30) / DAY_MS$1;
	if (time < EPOCH_1899_12_31) throw new RangeError(`Date ${date.toISOString()} is before 1899-12-31 (serial 0, used for time-only values), the start of the Transitional 1900 date system.`);
	if (time < TIME_1900_03_01) return (time - EPOCH_1899_12_31) / DAY_MS$1;
	return (time - EPOCH_1899_12_30) / DAY_MS$1;
}
const CODE_0 = 48;
const CODE_9 = 57;
const CODE_HYPHEN = 45;
const CODE_COLON = 58;
const CODE_PERIOD = 46;
const CODE_T = 84;
const CODE_Z = 90;
/** Longest accepted text: "YYYY-MM-DDThh:mm:ss." + 9 fraction digits + "Z". */
const MAX_ISO_LENGTH = 30;
const MAX_FRACTION_DIGITS = 9;
/** The number formed by `count` ASCII digits at `text[at]`, or -1 if any of them is not a digit. */
function digitsAt(text, at, count) {
	if (at + count > text.length) return -1;
	let value = 0;
	for (let i = at; i < at + count; i++) {
		const code = text.charCodeAt(i);
		if (code < CODE_0 || code > CODE_9) return -1;
		value = value * 10 + (code - CODE_0);
	}
	return value;
}
function daysInMonth(year, month) {
	if (month === 2) return year % 4 === 0 && year % 100 !== 0 || year % 400 === 0 ? 29 : 28;
	return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}
/**
* Parses the ISO 8601 lexical forms of Part 1 §18.17.4 used by `t="d"` cells: a date
* (`1975-10-05`), a time of day (`08:30`, `08:30:00`, `08:30:00.000`) or both joined by `T`
* (`1976-11-22T08:30:00.000`). A trailing `Z` after the time is accepted (as in the Annex L
* examples); other time-zone designators are not. Fractional seconds (1–9 digits) are rounded to
* the nearest millisecond. A time without a date is placed on 1899-12-30, the day of serial 0 in
* the Strict 1900 date system. Returns `null` for any other text, an impossible date or time
* (hour 24 and leap seconds are not permitted), or a value outside 0001-01-01 to
* 9999-12-31T23:59:59.999. The result's UTC fields hold the wall-clock value.
*/
function parseIsoDateTime(text) {
	const length = text.length;
	if (length < 5 || length > MAX_ISO_LENGTH) return null;
	let year = 1899;
	let month = 12;
	let day = 30;
	let i = 0;
	const hasDate = length >= 10 && text.charCodeAt(4) === CODE_HYPHEN;
	if (hasDate) {
		year = digitsAt(text, 0, 4);
		month = digitsAt(text, 5, 2);
		day = digitsAt(text, 8, 2);
		if (year < 1 || text.charCodeAt(7) !== CODE_HYPHEN || month < 1 || month > 12) return null;
		if (day < 1 || day > daysInMonth(year, month)) return null;
		i = 10;
	}
	let hour = 0;
	let minute = 0;
	let second = 0;
	let fractionMs = 0;
	if (!hasDate || i < length) {
		if (hasDate) {
			if (text.charCodeAt(i) !== CODE_T) return null;
			i++;
		}
		hour = digitsAt(text, i, 2);
		minute = digitsAt(text, i + 3, 2);
		if (hour < 0 || text.charCodeAt(i + 2) !== CODE_COLON || minute < 0) return null;
		i += 5;
		if (i < length && text.charCodeAt(i) === CODE_COLON) {
			second = digitsAt(text, i + 1, 2);
			if (second < 0) return null;
			i += 3;
			if (i < length && text.charCodeAt(i) === CODE_PERIOD) {
				i++;
				const fractionStart = i;
				let fraction = 0;
				while (i < length) {
					const code = text.charCodeAt(i);
					if (code < CODE_0 || code > CODE_9) break;
					fraction = fraction * 10 + (code - CODE_0);
					i++;
				}
				const digits = i - fractionStart;
				if (digits < 1 || digits > MAX_FRACTION_DIGITS) return null;
				fractionMs = Math.round(fraction * 1e3 / 10 ** digits);
			}
		}
		if (i < length && text.charCodeAt(i) === CODE_Z) i++;
		if (i !== length) return null;
		if (hour > 23 || minute > 59 || second > 59) return null;
	}
	const time = (/* @__PURE__ */ new Date(0)).setUTCFullYear(year, month - 1, day) + ((hour * 60 + minute) * 60 + second) * 1e3 + fractionMs;
	if (time < MIN_TIME || time > MAX_TIME) return null;
	return new Date(time);
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
/** Largest offset or size a ZIP without ZIP64 can hold (4-byte fields, Z-04). */
const MAX_32 = 4294967294;
const SIG_LOCAL$1 = 67324752;
const SIG_CENTRAL$1 = 33639248;
const SIG_EOCD$1 = 101010256;
/** Version 2.0: Deflate (Z-06); upper byte 0 = MS-DOS host. */
const VERSION = 20;
/**
* MS-DOS date for 1980-01-01: (year - 1980) << 9 | month << 5 | day. APPNOTE only says
* "standard MS-DOS format" (Z-19); the bit layout is the one fflate writes
* (node_modules/fflate/esm/index.mjs, zip header writer). Time 00:00:00 is 0.
*/
const DOS_DATE = 33;
const DOS_TIME = 0;
const encoder$1 = new TextEncoder();
function tooLarge(what) {
	return new XlvError("XLV_E_LIMIT_EXCEEDED", `The workbook is too large to write: ${what}.`, {
		hint: "Split the data into several workbooks. Writing ZIP64 archives is not supported yet.",
		details: { limit: "zip32" }
	});
}
/**
* Writes a ZIP archive. Throws `XLV_E_INVALID_INPUT` for names that are not printable ASCII or
* that repeat (ignoring case, as OPC requires, E-58), and `XLV_E_LIMIT_EXCEEDED` when the
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
		if (!/^[\x21-\x7e][\x20-\x7e]*$/.test(name) || name.startsWith("/") || name.length > 65535) throw new XlvError("XLV_E_INVALID_INPUT", `"${name.slice(0, 60)}" is not a valid ZIP item name.`, {
			hint: "Item names must be printable ASCII without a leading slash. This is an internal error; please report it.",
			details: { name: name.slice(0, 60) }
		});
		const key = name.toLowerCase();
		if (seen.has(key)) throw new XlvError("XLV_E_INVALID_INPUT", `The ZIP item name "${name.slice(0, 60)}" is used twice.`, {
			hint: "Item names must be unique ignoring case. This is an internal error; please report it.",
			details: { name: name.slice(0, 60) }
		});
		seen.add(key);
		const nameBytes = encoder$1.encode(name);
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
		lv.setUint32(0, SIG_LOCAL$1, true);
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
		cv.setUint32(0, SIG_CENTRAL$1, true);
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
	ev.setUint32(0, SIG_EOCD$1, true);
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
// src/sml/formula-groups.ts
const inside = (r, row, col) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
const overlaps = (a, b) => a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
function invalid(sheet, cell, message, hint) {
	return new XlvError("XLV_E_INVALID_VALUE", message, {
		location: {
			sheet,
			cell
		},
		hint,
		details: { cell }
	});
}
/** Tracks the formula groups of one worksheet while its cells are written in order. */
var FormulaGroups = class {
	sheet;
	active = [];
	masters = /* @__PURE__ */ new Map();
	next = 0;
	constructor(sheet) {
		this.sheet = sheet;
	}
	/**
	* The attributes of the cell's `f` element (`t`, `ref`, `si`), or `''` for a plain formula and
	* for other values. Throws `XLV_E_INVALID_VALUE` when a rule above is broken.
	*/
	attributes(row, col, cell, value) {
		for (let i = this.active.length - 1; i >= 0; i--) if (this.active[i].range.end.row < row) this.active.splice(i, 1);
		const array = this.active.find((g) => g.t === "array" && inside(g.range, row, col));
		if (value === null || value.kind !== "formula") return "";
		if (array !== void 0) throw invalid(this.sheet, cell, `${cell} lies in the range ${array.ref} of an array formula, so it cannot hold its own formula.`, `Give ${cell} its cached result as a plain value; the array formula in the top-left cell of ${array.ref} covers it.`);
		if (value.master !== void 0) {
			const master = this.masters.get(value.master);
			if (master === void 0) throw invalid(this.sheet, cell, `${cell} shares the formula of ${value.master}, which holds no shared formula master written before it.`, `Set ${value.master} to { formula, shared: "${value.master}:…" }, a range whose top-left cell is ${value.master} and that contains ${cell}.`);
			if (!inside(master.range, row, col)) throw invalid(this.sheet, cell, `${cell} lies outside the range ${master.ref} of the shared formula in ${value.master}.`, `Extend the shared range of ${value.master} to cover ${cell}, or give ${cell} its own formula.`);
			return ` t="shared" si="${String(master.si)}"`;
		}
		const g = value.group;
		if (g === void 0) return "";
		const range = parseRange(g.ref);
		const clash = this.active.find((a) => overlaps(a.range, range));
		if (clash !== void 0) throw invalid(this.sheet, cell, `The ${g.t} formula range ${g.ref} of ${cell} overlaps the ${clash.t} formula range ${clash.ref}.`, "Array and shared formula ranges of a sheet must not overlap. Shrink one of the ranges.");
		const group = {
			t: g.t,
			ref: g.ref,
			range,
			si: g.t === "shared" ? this.next++ : -1
		};
		this.active.push(group);
		if (g.t === "shared") {
			this.masters.set(cell, group);
			return ` t="shared" ref="${g.ref}" si="${String(group.si)}"`;
		}
		return ` t="array" ref="${g.ref}"`;
	}
};
// src/sml/write-parts.ts
const XML_DECL = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const NS = SML_NS.transitional;
const R_NS = REL_NS.transitional;
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const CT_BASE = "application/vnd.openxmlformats-officedocument.spreadsheetml.";
/** Default font of new workbooks (D-205). */
const DEFAULT_FONT = {
	name: "Calibri",
	size: 11
};
/** First id used for number formats declared in the styles part (D-204). */
const FIRST_CUSTOM_NUMFMT_ID = 164;
/** Built-in ids reused when a style's format code equals theirs: the non-date ones (D-204). */
const REUSABLE_BUILTINS = new Map([...BUILTIN_NUMBER_FORMATS].filter(([id]) => id !== 0 && !(id >= 14 && id <= 22) && !(id >= 45 && id <= 47)).map(([id, code]) => [code, id]));
const encoder = new TextEncoder();
/** Formats a finite number as an xsd:double lexical value that keeps every bit. */
function formatNumber(n) {
	return String(n).replace("e+", "E").replace("e-", "E-");
}
function attr(name, value) {
	return ` ${name}="${typeof value === "number" ? formatNumber(value) : escapeAttr(value)}"`;
}
function rangeRef(r) {
	const a = formatCellRef(r.start.row, r.start.col);
	return r.start.row === r.end.row && r.start.col === r.end.col ? a : `${a}:${formatCellRef(r.end.row, r.end.col)}`;
}
function absoluteRangeRef(r) {
	const abs = (row, col) => formatCellRef(row, col).replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");
	return `${abs(r.start.row, r.start.col)}:${abs(r.end.row, r.end.col)}`;
}
/** Quotes a sheet name for use in a formula: 'name' with apostrophes doubled (W2-08). */
function quoteSheetName(name) {
	return `'${name.replace(/'/g, "''")}'`;
}
function colorXml(element, color) {
	if (typeof color === "string") return `<${element}${attr("rgb", color)}/>`;
	if ("theme" in color) return `<${element}${attr("theme", color.theme)}${color.tint === void 0 ? "" : attr("tint", color.tint)}/>`;
	if ("indexed" in color) return `<${element}${attr("indexed", color.indexed)}/>`;
	return `<${element} auto="1"/>`;
}
var Table = class {
	items = [];
	ids = /* @__PURE__ */ new Map();
	add(xml) {
		let id = this.ids.get(xml);
		if (id === void 0) {
			id = this.items.length;
			this.items.push(xml);
			this.ids.set(xml, id);
		}
		return id;
	}
};
/**
* A `font` record: the font completed with `defaults` (name and size). Children are written in the
* order b, i, strike, u, sz, color, name: Office requires `i, strike, …, u, vertAlign, sz, color,
* name, family, charset, scheme` ([MS-OI29500] §2.1.706 a, which omits `b`) and puts `b` first in
* the run properties ([MS-OI29500] §2.1.674 a); CT_Font in sml.xsd allows any order (W2-11).
*/
function fontXml(font, defaults = DEFAULT_FONT) {
	return `<font>${fontChildrenXml({
		...defaults,
		...font
	}, "name")}</font>`;
}
/**
* The children of a font (CT_Font) or of run properties (CT_RPrElt, with the name in `rFont`),
* only those given, in the order Office requires: b, i, strike, condense, extend, outline,
* shadow, u, vertAlign, sz, color, name, family, charset, scheme ([MS-OI29500] §2.1.674 a,
* §2.1.706 a; W2-11, F-SST-08).
*/
function fontChildrenXml(f, nameElement) {
	let xml = "";
	if (f.bold) xml += "<b/>";
	if (f.italic) xml += "<i/>";
	if (f.strike) xml += "<strike/>";
	if (f.condense) xml += "<condense/>";
	if (f.extend) xml += "<extend/>";
	if (f.outline) xml += "<outline/>";
	if (f.shadow) xml += "<shadow/>";
	if (f.underline) xml += f.underline === "single" ? "<u/>" : `<u${attr("val", f.underline)}/>`;
	if (f.vertAlign !== void 0) xml += `<vertAlign${attr("val", f.vertAlign)}/>`;
	if (f.size !== void 0) xml += `<sz${attr("val", f.size)}/>`;
	if (f.color !== void 0) xml += colorXml("color", f.color);
	if (f.name !== void 0) xml += `<${nameElement}${attr("val", encodeXstring(f.name))}/>`;
	if (f.family !== void 0) xml += `<family${attr("val", f.family)}/>`;
	if (f.charset !== void 0) xml += `<charset${attr("val", f.charset)}/>`;
	if (f.scheme !== void 0) xml += `<scheme${attr("val", f.scheme)}/>`;
	return xml;
}
/** A `fill` record with a pattern fill. */
function fillXml(fill) {
	if ("gradient" in fill) {
		let xml = "<gradientFill";
		if (fill.gradient === "path") xml += " type=\"path\"";
		for (const k of [
			"degree",
			"left",
			"right",
			"top",
			"bottom"
		]) if (fill[k] !== void 0) xml += attr(k, fill[k]);
		xml += ">";
		for (const stop of fill.stops) xml += `<stop${attr("position", stop.position)}>${colorXml("color", stop.color)}</stop>`;
		return `<fill>${xml}</gradientFill></fill>`;
	}
	const { pattern, fgColor, bgColor } = fill;
	let inner = "";
	if (fgColor !== void 0) inner += colorXml("fgColor", fgColor);
	if (bgColor !== void 0) inner += colorXml("bgColor", bgColor);
	return `<fill><patternFill${attr("patternType", pattern)}${inner === "" ? "/>" : `>${inner}</patternFill>`}</fill>`;
}
/** A `border` record with the left, right, top, bottom and diagonal edges in schema order (W2-15, W2-61). */
function borderXml(border) {
	let xml = `<border${border.diagonalUp ? " diagonalUp=\"1\"" : ""}${border.diagonalDown ? " diagonalDown=\"1\"" : ""}>`;
	for (const edge of [
		"left",
		"right",
		"top",
		"bottom",
		"diagonal"
	]) {
		const e = border[edge];
		xml += e === void 0 ? `<${edge}/>` : `<${edge}${attr("style", e.style)}>${e.color === void 0 ? "" : colorXml("color", e.color)}</${edge}>`;
	}
	return `${xml}</border>`;
}
const READING_ORDER_VALUES = {
	context: 0,
	leftToRight: 1,
	rightToLeft: 2
};
/** An `alignment` element (CT_CellAlignment, W2-17, W2-62), or '' when there is nothing to write. */
function alignmentXml(a) {
	if (!a) return "";
	let al = "<alignment";
	if (a.horizontal !== void 0) al += attr("horizontal", a.horizontal);
	if (a.vertical !== void 0) al += attr("vertical", a.vertical);
	if (a.textRotation !== void 0) al += attr("textRotation", a.textRotation);
	if (a.wrapText) al += " wrapText=\"1\"";
	if (a.indent !== void 0) al += attr("indent", a.indent);
	if (a.justifyLastLine) al += " justifyLastLine=\"1\"";
	if (a.shrinkToFit) al += " shrinkToFit=\"1\"";
	if (a.readingOrder !== void 0) al += attr("readingOrder", READING_ORDER_VALUES[a.readingOrder]);
	return `${al}/>`;
}
/**
* A `protection` element (CT_CellProtection, W2-63) with the values that differ from Office's
* defaults, locked and not hidden ([MS-OI29500] §2.1.714), or '' when both are the defaults.
*/
function protectionXml(p) {
	if (!p) return "";
	let xml = "";
	if (p.locked === false) xml += " locked=\"0\"";
	if (p.hidden === true) xml += " hidden=\"1\"";
	return xml === "" ? "" : `<protection${xml}/>`;
}
/** The built-in id reused for a format code, if any (D-204). */
function reusableBuiltinNumFmt(code) {
	return REUSABLE_BUILTINS.get(code);
}
/**
* Builds the styles part. With `identity`, cell format `i` is interned style `i` (no merging of
* styles that give the same record), so a streaming writer can use style ids as `s` values
* before the part is built. `defaultFont` is font 0, the workbook's default font (W2-12); its
* name and size, or those of {@link DEFAULT_FONT}, complete the other fonts.
*/
function buildStyles(styles, identity = false, dxfs = [], defaultFont = null) {
	const fonts = new Table();
	const fills = new Table();
	const borders = new Table();
	const xfs = new Table();
	const numFmts = /* @__PURE__ */ new Map();
	const fontDefaults = {
		name: defaultFont?.name ?? DEFAULT_FONT.name,
		size: defaultFont?.size ?? DEFAULT_FONT.size
	};
	fonts.add(fontXml(defaultFont ?? {}, fontDefaults));
	fills.add("<fill><patternFill patternType=\"none\"/></fill>");
	fills.add("<fill><patternFill patternType=\"gray125\"/></fill>");
	borders.add("<border><left/><right/><top/><bottom/><diagonal/></border>");
	const numFmtId = (code) => {
		if (code === void 0) return 0;
		if (typeof code === "number") return code;
		const builtin = REUSABLE_BUILTINS.get(code);
		if (builtin !== void 0) return builtin;
		let id = numFmts.get(code);
		if (id === void 0) {
			id = FIRST_CUSTOM_NUMFMT_ID + numFmts.size;
			numFmts.set(code, id);
		}
		return id;
	};
	const xfOf = styles.map((style) => {
		const fontId = style.font ? fonts.add(fontXml(style.font, fontDefaults)) : 0;
		let fillId = 0;
		if (style.fill) fillId = fills.add(fillXml(style.fill));
		let borderId = 0;
		if (style.border) borderId = borders.add(borderXml(style.border));
		const nf = numFmtId(style.numFmt);
		let xf = `<xf${attr("numFmtId", nf)}${attr("fontId", fontId)}${attr("fillId", fillId)}${attr("borderId", borderId)} xfId="0"`;
		if (nf !== 0) xf += " applyNumberFormat=\"1\"";
		if (fontId !== 0) xf += " applyFont=\"1\"";
		if (fillId !== 0) xf += " applyFill=\"1\"";
		if (borderId !== 0) xf += " applyBorder=\"1\"";
		const al = alignmentXml(style.alignment);
		const pr = protectionXml(style.protection);
		if (al !== "") xf += " applyAlignment=\"1\"";
		if (pr !== "") xf += " applyProtection=\"1\"";
		xf += al === "" && pr === "" ? "/>" : `>${al}${pr}</xf>`;
		if (identity) {
			xfs.items.push(xf);
			return xfs.items.length - 1;
		}
		return xfs.add(xf);
	});
	const dxfXml = dxfs.map((d) => dxfRecordXml(d, numFmtId));
	checkStyleCount("customNumFmts", numFmts.size);
	checkStyleCount("fonts", fonts.items.length);
	checkStyleCount("fills", fills.items.length);
	checkStyleCount("borders", borders.items.length);
	checkStyleCount("cellXfs", xfs.items.length);
	let xml = `${XML_DECL}<styleSheet xmlns="${NS}">`;
	if (numFmts.size > 0) {
		xml += `<numFmts${attr("count", numFmts.size)}>`;
		for (const [code, id] of numFmts) xml += `<numFmt${attr("numFmtId", id)}${attr("formatCode", encodeXstring(code))}/>`;
		xml += "</numFmts>";
	}
	xml += `<fonts${attr("count", fonts.items.length)}>${fonts.items.join("")}</fonts>`;
	xml += `<fills${attr("count", fills.items.length)}>${fills.items.join("")}</fills>`;
	xml += `<borders${attr("count", borders.items.length)}>${borders.items.join("")}</borders>`;
	xml += "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>";
	xml += `<cellXfs${attr("count", xfs.items.length)}>${xfs.items.join("")}</cellXfs>`;
	xml += "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>";
	if (dxfXml.length > 0) xml += `<dxfs${attr("count", dxfXml.length)}>${dxfXml.join("")}</dxfs>`;
	xml += "</styleSheet>";
	return {
		xml,
		xfOf
	};
}
var SharedStrings = class {
	index = /* @__PURE__ */ new Map();
	list = [];
	total = 0;
	add(item) {
		this.total++;
		let i = this.index.get(item);
		if (i === void 0) {
			i = this.list.length;
			this.list.push(item);
			this.index.set(item, i);
		}
		return i;
	}
	xml() {
		let xml = `${XML_DECL}<sst xmlns="${NS}"${attr("count", this.total)}${attr("uniqueCount", this.list.length)}>`;
		for (const item of this.list) xml += `<si>${item}</si>`;
		return `${xml}</sst>`;
	}
};
/**
* Run properties (`rPr`, CT_RPrElt, F-SST-08) of a rich text run, or '' for a run without font.
* Only the given properties are written, in the order Office requires: b, i, strike, condense,
* extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme
* ([MS-OI29500] §2.1.674 a).
*/
function runPropertiesXml(font) {
	if (font === void 0) return "";
	const xml = fontChildrenXml(font, "rFont");
	return xml === "" ? "" : `<rPr>${xml}</rPr>`;
}
/**
* The content of a string item (CT_Rst, F-SST-02): one `t` for plain text, or one `r` per rich
* text run (`rPr` then `t`). Used for `si` in the shared string table and `is` in a cell.
*/
function stringItemXml(v) {
	if (v.runs === void 0) return textElement(v.s);
	let xml = "";
	for (const run of v.runs) xml += `<r>${runPropertiesXml(run.font)}${textElement(run.text)}</r>`;
	return xml;
}
/** Never used for strings: inline strings are handled before valueXml is called. */
const INLINE_ONLY = { add() {
	throw new Error("internal: inline strings are written by cellXml");
} };
/** A `t` element; `xml:space="preserve"` when white space at the ends or inside could be lost (W2-07). */
function textElement(s) {
	return `<t${/^[\s]|[\s]$|[\t\n\r]/.test(s) ? " xml:space=\"preserve\"" : ""}>${escapeText(encodeXstring(s))}</t>`;
}
function valueXml(v, sst, inFormula) {
	switch (v.kind) {
		case "number": return {
			t: "",
			v: formatNumber(v.n)
		};
		case "string": return inFormula ? {
			t: "str",
			v: escapeText(encodeXstring(v.s))
		} : {
			t: "s",
			v: String(sst.add(stringItemXml(v)))
		};
		case "boolean": return {
			t: "b",
			v: v.b ? "1" : "0"
		};
		case "error": return {
			t: "e",
			v: escapeText(v.e)
		};
	}
}
/**
* One `c` element. With `sst === null`, plain strings are written inline (`t="inlineStr"`,
* E-22, E-29) instead of into the shared string table.
*/
function cellXml(ref, cell, xfOf, sst, fAttributes = "") {
	const xf = xfOf === null ? cell.style : xfOf[cell.style] ?? 0;
	const s = xf === 0 ? "" : attr("s", xf);
	const v = cell.value;
	if (v === null) return `<c r="${ref}"${s}/>`;
	if (sst === null && v.kind === "string") return `<c r="${ref}"${s} t="inlineStr"><is>${stringItemXml(v)}</is></c>`;
	if (v.kind === "formula") {
		const f = v.master !== void 0 ? `<f${fAttributes}/>` : `<f${fAttributes}>${escapeText(encodeXstring(v.text))}</f>`;
		if (v.result === null) return `<c r="${ref}"${s}>${f}</c>`;
		const r = valueXml(v.result, sst ?? INLINE_ONLY, true);
		return `<c r="${ref}"${s}${r.t === "" ? "" : attr("t", r.t)}>${f}<v>${r.v}</v></c>`;
	}
	const r = valueXml(v, sst ?? INLINE_ONLY, false);
	return `<c r="${ref}"${s}${r.t === "" ? "" : attr("t", r.t)}><v>${r.v}</v></c>`;
}
/** Attributes of `sheetView` that differ from the CT_SheetView defaults (W2-51). */
function viewAttributes(view) {
	if (view === void 0) return "";
	let xml = "";
	if (!view.showGridLines) xml += " showGridLines=\"0\"";
	if (!view.showRowColHeaders) xml += " showRowColHeaders=\"0\"";
	if (view.rightToLeft) xml += " rightToLeft=\"1\"";
	if (!view.showRuler) xml += " showRuler=\"0\"";
	if (view.layout !== "normal") xml += attr("view", view.layout);
	if (view.topLeftCell !== null) xml += attr("topLeftCell", view.topLeftCell);
	if (view.zoom !== 100) xml += attr("zoomScale", view.zoom);
	if (view.zoomScaleNormal !== 0) xml += attr("zoomScaleNormal", view.zoomScaleNormal);
	return xml;
}
/**
* The start of a worksheet part, up to and including `<sheetData>`; `dimension` is optional
* (W2-01). `xfOf` maps interned style ids to cell format indexes (identity when omitted).
*/
function sheetHead(sheet, selected, dimension, xfOf) {
	let xml = `${XML_DECL}<worksheet xmlns="${NS}" xmlns:r="${R_NS}">`;
	const fitToPage = sheet.print?.pageSetup.fitToPage === true;
	const props = sheet.properties;
	const outlinePr = props && (!props.summaryBelow || !props.summaryRight) ? `<outlinePr${props.summaryBelow ? "" : " summaryBelow=\"0\""}${props.summaryRight ? "" : " summaryRight=\"0\""}/>` : "";
	if (sheet.tabColor || fitToPage || outlinePr) xml += `<sheetPr>${sheet.tabColor ? colorXml("tabColor", sheet.tabColor) : ""}${outlinePr}${fitToPage ? "<pageSetUpPr fitToPage=\"1\"/>" : ""}</sheetPr>`;
	if (dimension !== null) xml += `<dimension${attr("ref", dimension)}/>`;
	const attrsView = sheet.freeze && sheet.view ? {
		...sheet.view,
		topLeftCell: null
	} : sheet.view;
	xml += `<sheetViews><sheetView${selected ? " tabSelected=\"1\"" : ""}${viewAttributes(attrsView)} workbookViewId="0"`;
	const split = sheet.view?.split ?? null;
	if (split !== null && !sheet.freeze) {
		const pane = split.x > 0 && split.y > 0 ? "bottomRight" : split.y > 0 ? "bottomLeft" : "topRight";
		xml += "><pane";
		if (split.x > 0) xml += attr("xSplit", split.x * 20);
		if (split.y > 0) xml += attr("ySplit", split.y * 20);
		if (split.topLeftCell !== void 0) xml += attr("topLeftCell", split.topLeftCell);
		xml += `${attr("activePane", pane)} state="split"/>`;
		const active = sheet.view?.activeCell;
		xml += active ? `<selection${attr("pane", pane)}${attr("activeCell", active)}${attr("sqref", active)}/></sheetView>` : "</sheetView>";
	} else if (sheet.freeze) {
		const { rows, cols } = sheet.freeze;
		const pane = rows > 0 && cols > 0 ? "bottomRight" : rows > 0 ? "bottomLeft" : "topRight";
		const topLeft = formatCellRef(rows + 1, cols + 1);
		xml += "><pane";
		if (cols > 0) xml += attr("xSplit", cols);
		if (rows > 0) xml += attr("ySplit", rows);
		xml += `${attr("topLeftCell", topLeft)}${attr("activePane", pane)} state="frozen"/>`;
		const active = sheet.view?.activeCell ?? topLeft;
		xml += `<selection${attr("pane", pane)}${attr("activeCell", active)}${attr("sqref", active)}/></sheetView>`;
	} else if (sheet.view?.activeCell) xml += `><selection${attr("activeCell", sheet.view.activeCell)}${attr("sqref", sheet.view.activeCell)}/></sheetView>`;
	else xml += "/>";
	xml += "</sheetViews>";
	let rowLevel = props?.outlineLevelRow ?? 0;
	for (const r of sheet.rows?.values() ?? []) rowLevel = Math.max(rowLevel, r.outlineLevel ?? 0);
	let colLevel = props?.outlineLevelCol ?? 0;
	for (const c of sheet.columns.values()) colLevel = Math.max(colLevel, c.outlineLevel ?? 0);
	xml += `<sheetFormatPr${attr("defaultRowHeight", props?.defaultRowHeight ?? 15)}${rowLevel > 0 ? attr("outlineLevelRow", rowLevel) : ""}${colLevel > 0 ? attr("outlineLevelCol", colLevel) : ""}/>`;
	if (sheet.columns.size > 0) {
		xml += "<cols>";
		for (const index of [...sheet.columns.keys()].sort((a, b) => a - b)) {
			const c = sheet.columns.get(index);
			if (!c) continue;
			xml += `<col${attr("min", index)}${attr("max", index)}`;
			if (c.width !== null) xml += `${attr("width", c.width)} customWidth="1"`;
			if (c.hidden) xml += " hidden=\"1\"";
			const xf = c.style === 0 ? 0 : xfOf?.[c.style] ?? c.style;
			if (xf !== 0) xml += attr("style", xf);
			if (c.outlineLevel) xml += attr("outlineLevel", c.outlineLevel);
			if (c.collapsed) xml += " collapsed=\"1\"";
			xml += "/>";
		}
		xml += "</cols>";
	}
	return `${xml}<sheetData>`;
}
/**
* The start tag of a row (without `>` for rows that have cells: callers add `>` or `/>`). `xfOf`
* maps interned style ids to cell format indexes (identity when omitted).
*/
function rowStart(r, info, xfOf) {
	let xml = `<row${attr("r", r)}`;
	const style = info?.style ?? 0;
	const xf = style === 0 ? 0 : xfOf?.[style] ?? style;
	if (xf !== 0) xml += `${attr("s", xf)} customFormat="1"`;
	if (info?.height != null) xml += `${attr("ht", info.height)} customHeight="1"`;
	if (info?.hidden) xml += " hidden=\"1\"";
	if (info?.outlineLevel) xml += attr("outlineLevel", info.outlineLevel);
	if (info?.collapsed) xml += " collapsed=\"1\"";
	return xml;
}
/**
* The end of a worksheet part, from `</sheetData>` (W2-01 order: autoFilter, mergeCells). `after`
* is the XML of the elements that follow `mergeCells` in CT_Worksheet, already in schema order
* (`hyperlinks`, …, `legacyDrawing`; W2-42).
*/
function sheetTail(autoFilter, merges, after = "", protection = "") {
	let xml = `</sheetData>${protection}`;
	if (autoFilter) xml += `<autoFilter${attr("ref", rangeRef(autoFilter))}/>`;
	if (merges.length > 0) {
		xml += `<mergeCells${attr("count", merges.length)}>`;
		for (const m of merges) xml += `<mergeCell${attr("ref", rangeRef(m))}/>`;
		xml += "</mergeCells>";
	}
	return `${xml}${after}</worksheet>`;
}
/**
* The relationships of one worksheet part, collected while its XML is written. Each feature that
* points from the sheet to another part or to an external resource (hyperlinks; comments and
* their VML drawing) calls `add` and writes the returned id into the sheet's XML; `xml()` gives
* the part `xl/worksheets/_rels/sheetN.xml.rels`, or `null` when nothing was added, in which case
* no part is written. Relationship types are Transitional (Part 4 §13.3, W2-52), like every other
* relationship the writer emits. The `.rels` extension is covered by the `Default` content type.
*/
var SheetRels = class {
	items = [];
	/** Adds a relationship of type `REL_BASE + type` and returns its id (`rId1`, `rId2`, …). */
	add(type, target, external = false) {
		const id = `rId${String(this.items.length + 1)}`;
		this.items.push(`<Relationship${attr("Id", id)}${attr("Type", REL_BASE + type)}${attr("Target", target)}${external ? " TargetMode=\"External\"" : ""}/>`);
		return id;
	}
	xml() {
		return this.items.length === 0 ? null : `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">${this.items.join("")}</Relationships>`;
	}
};
/** Part name of the relationships part of worksheet `n` (1-based). */
function sheetRelsPartName(n) {
	return `xl/worksheets/_rels/sheet${String(n)}.xml.rels`;
}
/**
* The Table Definition parts of one worksheet and its `tableParts` element (written last before
* `extLst`, CT_Worksheet sml.xsd T2210; TD-02). Tables are numbered across the workbook from
* `first` (1-based): that number is both the part name and the table `id`, unique per workbook
* (Part 1 §18.5.1.2). Each part is the target of a `table` relationship of the sheet.
*/
function tableParts(tables, first, rels) {
	if (tables.length === 0) return {
		element: "",
		parts: []
	};
	const parts = tables.map((t, i) => ({
		name: tablePartName(first + i),
		xml: tableXml(t, first + i)
	}));
	const ids = parts.map((p) => rels.add("table", `../tables/${p.name.slice(10)}`));
	return {
		element: `<tableParts${attr("count", ids.length)}>${ids.map((id) => `<tablePart${attr("r:id", id)}/>`).join("")}</tableParts>`,
		parts
	};
}
/**
* A differential format (CT_Dxf: font, numFmt, fill, alignment, border, protection; SR-16). The
* font is written without the workbook defaults: only what the format changes. A solid fill with
* one colour is written with that colour as both foreground and background, as in every dxf
* example of ECMA-376 Part 1 (Annex L.2.7.3.8, L.2.7.4.3); the standard does not say which of
* the two a spreadsheet application uses in a dxf (SR-16). `numFmtId` is required by the schema
* and ignored by Office outside `numFmts` ([MS-OI29500] §2.1.712 c).
*/
function dxfRecordXml(style, numFmtId) {
	let xml = "<dxf>";
	const f = style.font;
	if (f) xml += `<font>${fontChildrenXml(f, "name")}</font>`;
	if (style.numFmt !== void 0) {
		const id = numFmtId(style.numFmt);
		const code = typeof style.numFmt === "number" ? BUILTIN_NUMBER_FORMATS.get(style.numFmt) ?? "General" : style.numFmt;
		xml += `<numFmt${attr("numFmtId", id)}${attr("formatCode", encodeXstring(code))}/>`;
	}
	if (style.fill && "gradient" in style.fill) xml += fillXml(style.fill);
	else if (style.fill) {
		const { pattern, fgColor, bgColor } = style.fill;
		const solidOne = pattern === "solid" && fgColor === void 0 !== (bgColor === void 0);
		const fg = solidOne ? fgColor ?? bgColor : fgColor;
		const bg = solidOne ? bgColor ?? fgColor : bgColor;
		xml += fillXml({
			pattern,
			...fg === void 0 ? {} : { fgColor: fg },
			...bg === void 0 ? {} : { bgColor: bg }
		});
	}
	xml += alignmentXml(style.alignment);
	if (style.border) xml += borderXml(style.border);
	xml += protectionXml(style.protection);
	return `${xml}</dxf>`;
}
function cfValuesXml(values) {
	return values.map((v) => {
		const type = v.type === "autoMin" ? "min" : v.type === "autoMax" ? "max" : v.type;
		return `<cfvo${attr("type", type)}${v.value === null || type !== v.type ? "" : attr("val", encodeXstring(v.value))}${v.gte ? "" : " gte=\"0\""}/>`;
	}).join("");
}
/** Value objects of the extension: the value in `xm:f`, `gte` only in icon sets ([MS-XLSX] §2.6.26). */
function x14CfValuesXml(values, iconSet) {
	return values.map((v) => {
		const open = `<x14:cfvo${attr("type", v.type)}${iconSet && !v.gte ? " gte=\"0\"" : ""}`;
		return v.value === null || v.type === "min" || v.type === "max" || v.type === "autoMin" || v.type === "autoMax" ? `${open}/>` : `${open}><xm:f>${escapeText(encodeXstring(v.value ?? ""))}</xm:f></x14:cfvo>`;
	}).join("");
}
/** Icon sets that exist only in the extension ([MS-XLSX] §2.7.10, SR-43). */
const X14_ICON_SETS$1 = /* @__PURE__ */ new Set([
	"3Stars",
	"3Triangles",
	"5Boxes",
	"NoIcons"
]);
/** A rule that only the [MS-XLSX] extension can hold: an extension icon set or custom icons (SR-43). */
function extensionOnly(r) {
	return r.iconSet !== null && (r.iconSet.custom !== null || X14_ICON_SETS$1.has(r.iconSet.iconSet));
}
/**
* The `id` that links a main data bar rule to its extension (ST_Guid, [MS-XLSX] §2.4.28,
* §2.6.27). [MS-XLSX] requires only that the main rule and the extension carry the same value;
* xlvane derives it from the sheet number and the rule's position on the sheet, so the output is
* the same for the same input (SR-41).
*/
function ruleGuid(sheet, rule) {
	const hex = (n, width) => n.toString(16).toUpperCase().padStart(width, "0");
	return `{${hex(sheet, 8)}-0000-0000-0000-${hex(rule, 12)}}`;
}
/** `x14:dataBar` of a rule (§2.6.30); `fill` only for a rule with its own priority. */
function x14DataBarXml(d, fill) {
	let xml = "<x14:dataBar";
	if (d.minLength !== 10) xml += attr("minLength", d.minLength);
	if (d.maxLength !== 90) xml += attr("maxLength", d.maxLength);
	if (!d.showValue) xml += " showValue=\"0\"";
	if (d.border) xml += " border=\"1\"";
	if (!d.gradient) xml += " gradient=\"0\"";
	if (d.direction !== "context") xml += attr("direction", d.direction);
	if (d.negativeFillColor === null) xml += " negativeBarColorSameAsPositive=\"1\"";
	if (d.negativeBorderColor !== null) xml += " negativeBarBorderColorSameAsPositive=\"0\"";
	if (d.axisPosition !== "automatic") xml += attr("axisPosition", d.axisPosition);
	xml += `>${x14CfValuesXml(d.cfvo, false)}`;
	if (fill) xml += colorXml("x14:fillColor", d.color);
	if (d.borderColor !== null) xml += colorXml("x14:borderColor", d.borderColor);
	if (d.negativeFillColor !== null) xml += colorXml("x14:negativeFillColor", d.negativeFillColor);
	if (d.negativeBorderColor !== null) xml += colorXml("x14:negativeBorderColor", d.negativeBorderColor);
	if (d.axisColor !== null) xml += colorXml("x14:axisColor", d.axisColor);
	return `${xml}</x14:dataBar>`;
}
/** `x14:iconSet` of a rule (§2.6.28, §2.6.36); `percent` is undefined there and not written. */
function x14IconSetXml(s) {
	let xml = "<x14:iconSet";
	if (s.iconSet !== "3TrafficLights1") xml += attr("iconSet", s.iconSet);
	if (!s.showValue) xml += " showValue=\"0\"";
	if (s.reverse && s.custom === null) xml += " reverse=\"1\"";
	if (s.custom !== null) xml += " custom=\"1\"";
	xml += `>${x14CfValuesXml(s.cfvo, true)}`;
	for (const icon of s.custom ?? []) xml += `<x14:cfIcon${attr("iconSet", icon.iconSet)}${attr("iconId", icon.iconId)}/>`;
	return `${xml}</x14:iconSet>`;
}
/**
* The `conditionalFormatting` elements and the `dataValidations` element of a sheet, in this
* order (CT_Worksheet: … mergeCells, phoneticPr, conditionalFormatting*, dataValidations,
* hyperlinks …, sml.xsd T2186–T2191; SR-01, SR-07), and the worksheet `extLst` with the
* [MS-XLSX] extensions (SR-40 … SR-43): the extension of a data bar that uses extension
* features, linked to its main rule by `id`, and rules only the extension can hold, with their
* own priority. Attributes are written only where they differ from the schema defaults. `sheet`
* is the 1-based sheet number, used for the rule ids.
*/
function sheetRulesXml(sheet, sheetNumber = 1) {
	let xml = "";
	let ext = "";
	let ruleNumber = 0;
	for (const cf of sheet.conditionalFormats ?? []) {
		let main = "";
		let extended = "";
		cf.rules.forEach((r, i) => {
			ruleNumber++;
			if (extensionOnly(r) && r.iconSet !== null) {
				extended += `<x14:cfRule type="iconSet"${attr("priority", r.priority)}${r.stopIfTrue ? " stopIfTrue=\"1\"" : ""}>${x14IconSetXml(r.iconSet)}</x14:cfRule>`;
				return;
			}
			const dxf = cf.dxfIds[i];
			main += `<cfRule${attr("type", r.type)}${dxf === null || dxf === void 0 ? "" : attr("dxfId", dxf)}${attr("priority", r.priority)}`;
			if (r.stopIfTrue) main += " stopIfTrue=\"1\"";
			if (r.type === "aboveAverage" && !r.aboveAverage) main += " aboveAverage=\"0\"";
			if (r.percent) main += " percent=\"1\"";
			if (r.bottom) main += " bottom=\"1\"";
			if (r.operator !== null) main += attr("operator", r.operator);
			if (r.text !== null) main += attr("text", r.text);
			if (r.timePeriod !== null) main += attr("timePeriod", r.timePeriod);
			if (r.rank !== null) main += attr("rank", r.rank);
			if (r.stdDev !== null) main += attr("stdDev", r.stdDev);
			if (r.equalAverage) main += " equalAverage=\"1\"";
			let inner = r.formulas.map((f) => `<formula>${escapeText(encodeXstring(f))}</formula>`).join("");
			if (r.colorScale) inner += `<colorScale>${cfValuesXml(r.colorScale.cfvo)}${r.colorScale.colors.map((c) => colorXml("color", c)).join("")}</colorScale>`;
			if (r.dataBar) {
				const d = r.dataBar;
				const full = d.extended && d.minLength === 0 && d.maxLength === 100;
				const min = full ? 10 : d.minLength;
				const max = full ? 90 : d.maxLength;
				inner += `<dataBar${min === 10 ? "" : attr("minLength", min)}${max === 90 ? "" : attr("maxLength", max)}${d.showValue ? "" : " showValue=\"0\""}>${cfValuesXml(d.cfvo)}${colorXml("color", d.color)}</dataBar>`;
				if (d.extended) {
					const id = ruleGuid(sheetNumber, ruleNumber);
					inner += `<extLst><ext${attr("uri", X14_ID_URI)}${attr("xmlns:x14", X14_NS)}><x14:id>${id}</x14:id></ext></extLst>`;
					extended += `<x14:cfRule type="dataBar"${attr("id", id)}>${x14DataBarXml(d, false)}</x14:cfRule>`;
				}
			}
			if (r.iconSet) {
				const s = r.iconSet;
				inner += `<iconSet${s.iconSet === "3TrafficLights1" ? "" : attr("iconSet", s.iconSet)}${s.showValue ? "" : " showValue=\"0\""}${s.percent ? "" : " percent=\"0\""}${s.reverse ? " reverse=\"1\"" : ""}>${cfValuesXml(s.cfvo)}</iconSet>`;
			}
			main += inner === "" ? "/>" : `>${inner}</cfRule>`;
		});
		if (main !== "") xml += `<conditionalFormatting${attr("sqref", cf.sqref)}>${main}</conditionalFormatting>`;
		if (extended !== "") ext += `<x14:conditionalFormatting>${extended}<xm:sqref>${cf.sqref}</xm:sqref></x14:conditionalFormatting>`;
	}
	const dvs = sheet.dataValidations ?? [];
	if (dvs.length > 0) {
		xml += `<dataValidations${attr("count", dvs.length)}>`;
		for (const d of dvs) {
			xml += `<dataValidation${d.type === "none" ? "" : attr("type", d.type)}`;
			if (d.errorStyle !== "stop") xml += attr("errorStyle", d.errorStyle);
			if (d.imeMode !== "noControl") xml += attr("imeMode", d.imeMode);
			if (d.operator !== null && d.operator !== "between") xml += attr("operator", d.operator);
			if (d.allowBlank) xml += " allowBlank=\"1\"";
			if (d.showDropDown) xml += " showDropDown=\"1\"";
			if (d.showInputMessage) xml += " showInputMessage=\"1\"";
			if (d.showErrorMessage) xml += " showErrorMessage=\"1\"";
			for (const k of [
				"errorTitle",
				"error",
				"promptTitle",
				"prompt"
			]) {
				const v = d[k];
				if (v !== null) xml += attr(k, encodeXstring(v));
			}
			xml += attr("sqref", d.sqref);
			const f = (d.formula1 === null ? "" : `<formula1>${escapeText(encodeXstring(d.formula1))}</formula1>`) + (d.formula2 === null ? "" : `<formula2>${escapeText(encodeXstring(d.formula2))}</formula2>`);
			xml += f === "" ? "/>" : `>${f}</dataValidation>`;
		}
		xml += "</dataValidations>";
	}
	return {
		main: xml,
		ext: ext === "" ? "" : `<extLst><ext${attr("uri", X14_CF_URI)}${attr("xmlns:x14", X14_NS)}${attr("xmlns:xm", XM_NS)}><x14:conditionalFormattings>${ext}</x14:conditionalFormattings></ext></extLst>`
	};
}
/**
* The `hyperlinks` element (§18.3.1.47, §18.3.1.48; W2-42). An external target gets a
* relationship with TargetMode="External" and `r:id`; a place in the workbook is written in
* `location` only, without a relationship.
*/
function hyperlinksXml(links, rels) {
	if (links.length === 0) return "";
	let xml = "<hyperlinks>";
	for (const l of links) {
		xml += `<hyperlink${attr("ref", l.ref)}`;
		if (l.url !== null) xml += attr("r:id", rels.add("hyperlink", l.url, true));
		if (l.location !== null) xml += attr("location", encodeXstring(l.location));
		if (l.tooltip !== null) xml += attr("tooltip", encodeXstring(l.tooltip));
		if (l.display !== null) xml += attr("display", encodeXstring(l.display));
		xml += "/>";
	}
	return `${xml}</hyperlinks>`;
}
/** Part name of the Comments part of worksheet `n` (1-based). */
function commentsPartName(n) {
	return `xl/comments${String(n)}.xml`;
}
/**
* Builds the Comments part (W2-45, W2-46) and the VML Drawing part (W2-47) of worksheet `n` and
* adds their relationships to `rels`; `null` when the sheet has no notes. Shape ids use the
* 1024-id blocks from `block` on, written in `o:idmap`, as in the examples of Part 1 Annex
* L.5.5.2 and Part 4 §19.4 (`data="1"` with ids from 1025).
*/
function noteParts(notes, n, rels, block) {
	if (notes.length === 0) return null;
	const commentsName = commentsPartName(n);
	const vmlName = `xl/drawings/vmlDrawing${String(n)}.vml`;
	rels.add("comments", `../comments${String(n)}.xml`);
	const vmlId = rels.add("vmlDrawing", `../drawings/vmlDrawing${String(n)}.vml`);
	const authors = /* @__PURE__ */ new Map();
	let list = "";
	for (const note of notes) {
		let id = authors.get(note.author);
		if (id === void 0) {
			id = authors.size;
			authors.set(note.author, id);
		}
		list += `<comment${attr("ref", formatCellRef(note.row, note.col))}${attr("authorId", id)}><text>${stringItemXml(note.value)}</text></comment>`;
	}
	let comments = `${XML_DECL}<comments xmlns="${NS}"><authors>`;
	for (const author of authors.keys()) comments += `<author>${escapeText(encodeXstring(author))}</author>`;
	comments += `</authors><commentList>${list}</commentList></comments>`;
	const blocks = Math.ceil(notes.length / 1024);
	return {
		legacyDrawing: `<legacyDrawing${attr("r:id", vmlId)}/>`,
		comments: {
			name: commentsName,
			xml: comments
		},
		vml: {
			name: vmlName,
			xml: vmlXml(notes, block, blocks)
		},
		blocks
	};
}
/** Points per default column (64 px) and row (20 px) and per pixel, for the box position. */
const COLUMN_PT = 48;
const ROW_PT = 15;
const PX_PT = .75;
/**
* The VML Drawing part (Part 4 §8.2, §19.1.2.19, §19.4; W2-47): one hidden text box shape per
* note, with the `x:ClientData` of the Part 4 §19.4 example minus `x:Visible` (so the note is
* hidden, §19.4.2.65). The box lies right of the cell; its anchor follows the Part 1 Annex
* L.5.5.2 example, and the CSS position is computed for default column widths and row heights.
*/
function vmlXml(notes, block, blocks) {
	let xml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit"${attr("data", Array.from({ length: blocks }, (_, i) => String(block + i)).join(","))}/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`;
	notes.forEach((note, k) => {
		const row = note.row - 1;
		const col = note.col - 1;
		const top = Math.max(row - 1, 0);
		const style = `position:absolute;margin-left:${String((col + 1) * COLUMN_PT + 15 * PX_PT)}pt;margin-top:${String(top * ROW_PT + 10 * PX_PT)}pt;width:${String(96)}pt;height:${String(60 - 6 * PX_PT)}pt;z-index:${String(k + 1)};visibility:hidden`;
		const box = note.box;
		const inset = box.margins === null ? "" : attr("inset", [
			box.margins.left,
			box.margins.top,
			box.margins.right,
			box.margins.bottom
		].map((v) => `${String(v)}mm`).join(","));
		const anchorFlags = box.anchor === "absolute" ? "<x:MoveWithCells/><x:SizeWithCells/>" : box.anchor === "oneCell" ? "<x:SizeWithCells/>" : "";
		xml += `<v:shape${attr("id", `_x0000_s${String(block * 1024 + k + 1)}`)} type="#_x0000_t202"${attr("style", style)} fillcolor="#ffffe1"${attr("o:insetmode", box.autoMargins ? "auto" : "custom")}><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"${inset}><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note">${anchorFlags}<x:Anchor>${[
			col + 1,
			15,
			top,
			10,
			col + 3,
			15,
			top + 4,
			4
		].join(", ")}</x:Anchor>${box.locked ? "" : "<x:Locked>False</x:Locked>"}<x:AutoFill>False</x:AutoFill>${box.lockText ? "" : "<x:LockText>False</x:LockText>"}<x:Row>${String(row)}</x:Row><x:Column>${String(col)}</x:Column></x:ClientData></v:shape>`;
	});
	return `${xml}</xml>`;
}
function sheetXml(sheet, selected, xfOf, sst, rels, after = "", sheetNumber = 1) {
	const sortedRows = [.../* @__PURE__ */ new Set([...sheet.cells.keys(), ...sheet.rows.keys()])].sort((a, b) => a - b);
	let minRow = Infinity;
	let maxRow = 0;
	let minCol = Infinity;
	let maxCol = 0;
	for (const [r, cells] of sheet.cells) for (const c of cells.keys()) {
		if (r < minRow) minRow = r;
		if (r > maxRow) maxRow = r;
		if (c < minCol) minCol = c;
		if (c > maxCol) maxCol = c;
	}
	let xml = sheetHead(sheet, selected, maxRow === 0 ? "A1" : rangeRef({
		start: {
			row: minRow,
			col: minCol
		},
		end: {
			row: maxRow,
			col: maxCol
		}
	}), xfOf);
	const groups = new FormulaGroups(sheet.name);
	for (const r of sortedRows) {
		const cells = sheet.cells.get(r);
		const start = rowStart(r, sheet.rows.get(r), xfOf);
		if (!cells || cells.size === 0) {
			xml += `${start}/>`;
			continue;
		}
		xml += `${start}>`;
		for (const c of [...cells.keys()].sort((a, b) => a - b)) {
			const cell = cells.get(c);
			if (cell) {
				const ref = formatCellRef(r, c);
				xml += cellXml(ref, cell, xfOf, sst, groups.attributes(r, c, ref, cell.value));
			}
		}
		xml += "</row>";
	}
	const print = sheet.print;
	const rules = sheetRulesXml(sheet, sheetNumber);
	return xml + sheetTail(sheet.autoFilter, sheet.merges, rules.main + hyperlinksXml(sheet.hyperlinks ?? [], rels) + (print ? printXml(print) : "") + after + rules.ext, print ? sheetProtectionXml(print.protection) : "");
}
/** Builds all parts and returns the zipped package. */
function writePackage(input) {
	const { xml: stylesXml, xfOf } = buildStyles(input.styles, false, input.dxfs, input.defaultFont);
	const shared = input.sharedStrings !== false;
	const sst = shared ? new SharedStrings() : null;
	const active = input.sheets.findIndex((s) => s.state === "visible");
	const sheetRels = input.sheets.map(() => new SheetRels());
	const notes = [];
	let block = 1;
	const tables = [];
	let tableCount = 0;
	const media = new MediaParts(input.images ?? []);
	const drawings = [];
	const charts = [];
	const sheetNames = chartSheetNames(input.sheets.map((s) => s.name));
	checkChartCacheCells(input.sheets.flatMap((s) => (s.charts ?? []).map((c) => ({
		chart: c.chart,
		sheet: s.name
	}))));
	const cells = chartCells(input.sheets, input.styles);
	const cacheText = { used: 0 };
	const sheetParts = input.sheets.map((sheet, i) => {
		const rels = sheetRels[i];
		const parts = noteParts(sheet.notes ?? [], i + 1, rels, block);
		notes.push(parts);
		if (parts !== null) block += parts.blocks;
		const frames = (sheet.charts ?? []).map((c) => {
			charts.push({
				name: chartPartName(charts.length + 1),
				xml: chartPartXml(resolveChartSheets(c.chart, sheetNames, sheet.name), {
					cells,
					text: cacheText,
					sheet: sheet.name
				})
			});
			return {
				from: c.from,
				to: c.to,
				ext: c.ext,
				editAs: c.editAs,
				name: c.chart.name,
				description: c.chart.description,
				target: `../charts/chart${String(charts.length)}.xml`
			};
		});
		const drawing = drawingParts(sheet.images ?? [], sheet.background ?? null, media, drawings.length + 1, rels, frames);
		if (drawing.part !== null) drawings.push(drawing.part);
		const sheetTables = tableParts(sheet.tables ?? [], tableCount + 1, rels);
		tableCount += sheetTables.parts.length;
		tables.push(...sheetTables.parts);
		return sheetXml(sheet, i === active, xfOf, sst, rels, drawing.drawing + (parts?.legacyDrawing ?? "") + drawing.picture + sheetTables.element, i + 1);
	});
	const n = input.sheets.length;
	const workbook = workbookXml(input.sheets, input.dateSystem, input.definedNames, input.workbookProtection ?? null, {
		bookViews: input.bookViews ?? [],
		calculation: input.calculation ?? null
	});
	const workbookRels = workbookRelsXml(n, shared);
	const core = corePropertiesXml(input.properties ?? {});
	const app = appPropertiesXml(input.properties ?? {}, input.sheets.map((sheet) => sheet.name), nameTitles(input.definedNames ?? []));
	const rootRels = rootRelsXml(core !== null, app !== null);
	const noteSheets = notes.flatMap((p, i) => p === null ? [] : [i + 1]);
	const mediaEntries = media.entries();
	const types = contentTypesXml(n, shared, core !== null, app !== null, noteSheets, [
		...tables.map((t) => [t.name, TABLE_CONTENT_TYPE]),
		...drawings.map((d) => [d.name, DRAWING_CONTENT_TYPE]),
		...charts.map((c) => [c.name, CHART_CONTENT_TYPE]),
		...mediaEntries.map((m) => [m.name, IMAGE_CONTENT_TYPES[m.name.slice(m.name.lastIndexOf(".") + 1)]])
	]);
	const entries = [
		{
			name: "[Content_Types].xml",
			data: encoder.encode(types)
		},
		{
			name: "_rels/.rels",
			data: encoder.encode(rootRels)
		},
		{
			name: "xl/workbook.xml",
			data: encoder.encode(workbook)
		},
		{
			name: "xl/_rels/workbook.xml.rels",
			data: encoder.encode(workbookRels)
		},
		...sheetParts.map((xml, i) => ({
			name: `xl/worksheets/sheet${String(i + 1)}.xml`,
			data: encoder.encode(xml)
		})),
		{
			name: "xl/styles.xml",
			data: encoder.encode(stylesXml)
		},
		...sst ? [{
			name: "xl/sharedStrings.xml",
			data: encoder.encode(sst.xml())
		}] : []
	];
	sheetRels.forEach((rels, i) => {
		const xml = rels.xml();
		if (xml !== null) entries.push({
			name: sheetRelsPartName(i + 1),
			data: encoder.encode(xml)
		});
		const parts = notes[i];
		if (parts) entries.push({
			name: parts.comments.name,
			data: encoder.encode(parts.comments.xml)
		}, {
			name: parts.vml.name,
			data: encoder.encode(parts.vml.xml)
		});
	});
	for (const t of tables) entries.push({
		name: t.name,
		data: encoder.encode(t.xml)
	});
	for (const d of drawings) entries.push({
		name: d.name,
		data: encoder.encode(d.xml)
	}, {
		name: d.relsName,
		data: encoder.encode(d.rels)
	});
	for (const c of charts) entries.push({
		name: c.name,
		data: encoder.encode(c.xml)
	});
	entries.push(...mediaEntries);
	if (core !== null) entries.push({
		name: CORE_PROPERTIES_PART,
		data: encoder.encode(core)
	});
	if (app !== null) entries.push({
		name: APP_PROPERTIES_PART,
		data: encoder.encode(app)
	});
	return writeZip(entries, input.compressionLevel ?? 6);
}
/**
* The values of the workbook's cells for chart caches (CH-19): numbers with the code of their
* number format, text, and booleans as TRUE or FALSE; a formula gives its cached result; errors
* and empty cells give `null`.
*/
function chartCells(sheets, styles) {
	const bySheet = new Map(sheets.map((s) => [s.name, s.cells]));
	const formatOf = (style) => {
		const f = styles[style]?.numFmt;
		return f === void 0 ? "General" : typeof f === "number" ? BUILTIN_NUMBER_FORMATS.get(f) ?? "General" : f;
	};
	const valueOf = (v, style) => {
		if (v === null) return null;
		if (v.kind === "number") return {
			n: v.n,
			format: formatOf(style)
		};
		if (v.kind === "string") return { s: v.s };
		if (v.kind === "boolean") return { s: v.b ? "TRUE" : "FALSE" };
		if (v.kind === "formula") return valueOf(v.result, style);
		return null;
	};
	return (sheet, row, col) => {
		const cell = bySheet.get(sheet)?.get(row)?.get(col);
		return cell === void 0 ? null : valueOf(cell.value, cell.style);
	};
}
/**
* The defined names as listed in `TitlesOfParts` of `docProps/app.xml` (Part 1 §18.17.6.5): names
* that are not hidden, a sheet-level name prefixed with its quoted sheet, e.g. `'Data'!Total`.
* This follows the files written by Excel in the Apache POI corpus (W2-52).
*/
function nameTitles(definedNames) {
	return definedNames.filter((n) => !n.hidden).map((n) => {
		const name = n.name.startsWith("_xlnm.") ? n.name.slice(6) : n.name;
		return n.sheet === null ? name : `${quoteSheetName(n.sheet)}!${name}`;
	});
}
/**
* The sheet ids to write: the given ones, and the lowest unused ids for the other sheets in sheet
* order, so that sheets without given ids get 1, 2, 3, … (unique, as E-09 requires; D-F03).
*/
function sheetIds(sheets) {
	const used = /* @__PURE__ */ new Set();
	for (const s of sheets) if (typeof s.sheetId === "number") used.add(s.sheetId);
	let next = 1;
	return sheets.map((s) => {
		if (typeof s.sheetId === "number") return s.sheetId;
		while (used.has(next)) next++;
		used.add(next);
		return next;
	});
}
/** The workbook part: date system, book view, sheets, and the autofilter and user defined names. */
function workbookXml(sheets, dateSystem, definedNames = [], protection = null, extras = {}) {
	const active = sheets.findIndex((s) => s.state === "visible");
	let workbook = `${XML_DECL}<workbook xmlns="${NS}" xmlns:r="${R_NS}">`;
	if (dateSystem === "1904") workbook += "<workbookPr date1904=\"1\"/>";
	workbook += workbookProtectionXml(protection);
	const views = extras.bookViews ?? [];
	if (views.length === 0) workbook += `<bookViews><workbookView${active > 0 ? `${attr("firstSheet", active)}${attr("activeTab", active)}` : ""}/></bookViews>`;
	else {
		workbook += "<bookViews>";
		for (const v of views) {
			workbook += "<workbookView";
			if (v.visibility !== void 0 && v.visibility !== "visible") workbook += attr("visibility", v.visibility);
			if (v.x !== void 0) workbook += attr("xWindow", v.x);
			if (v.y !== void 0) workbook += attr("yWindow", v.y);
			if (v.width !== void 0) workbook += attr("windowWidth", v.width);
			if (v.height !== void 0) workbook += attr("windowHeight", v.height);
			if (v.firstSheet !== void 0) workbook += attr("firstSheet", v.firstSheet);
			if (v.activeTab !== void 0) workbook += attr("activeTab", v.activeTab);
			workbook += "/>";
		}
		workbook += "</bookViews>";
	}
	workbook += "<sheets>";
	const ids = sheetIds(sheets);
	sheets.forEach((sheet, i) => {
		workbook += `<sheet${attr("name", encodeXstring(sheet.name))}${attr("sheetId", ids[i] ?? i + 1)}`;
		if (sheet.state !== "visible") workbook += attr("state", sheet.state);
		workbook += `${attr("r:id", `rId${String(i + 1)}`)}/>`;
	});
	workbook += "</sheets>";
	const names = sheets.map((sheet, i) => sheet.autoFilter ? `<definedName name="_xlnm._FilterDatabase"${attr("localSheetId", i)} hidden="1">${escapeText(encodeXstring(`${quoteSheetName(sheet.name)}!${absoluteRangeRef(sheet.autoFilter)}`))}</definedName>` : "").join("") + definedNames.map((n) => {
		const id = n.sheet === null ? null : sheets.findIndex((s) => s.name === n.sheet);
		return definedNameXml(n, id === null || id < 0 ? null : id);
	}).join("");
	if (names !== "") workbook += `<definedNames>${names}</definedNames>`;
	if (extras.calculation) workbook += `<calcPr${extras.calculation.fullCalcOnLoad ? " fullCalcOnLoad=\"1\"" : ""}/>`;
	return `${workbook}</workbook>`;
}
/** Relationships of the workbook part: sheets rId1…rIdN, then styles and (optionally) shared strings. */
function workbookRelsXml(sheetCount, sharedStrings) {
	let xml = `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">`;
	for (let i = 1; i <= sheetCount; i++) xml += `<Relationship Id="rId${String(i)}" Type="${REL_BASE}worksheet" Target="worksheets/sheet${String(i)}.xml"/>`;
	xml += `<Relationship Id="rId${String(sheetCount + 1)}" Type="${REL_BASE}styles" Target="styles.xml"/>`;
	if (sharedStrings) xml += `<Relationship Id="rId${String(sheetCount + 2)}" Type="${REL_BASE}sharedStrings" Target="sharedStrings.xml"/>`;
	return `${xml}</Relationships>`;
}
/** The package relationships part, with the Core and Extended File Properties relationships when asked. */
function rootRelsXml(coreProperties = false, appProperties = false) {
	return `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${REL_BASE}officeDocument" Target="xl/workbook.xml"/>` + (coreProperties ? `<Relationship Id="rId2" Type="${CORE_PROPERTIES_REL}" Target="${CORE_PROPERTIES_PART}"/>` : "") + (appProperties ? `<Relationship Id="rId3" Type="${APP_PROPERTIES_REL}" Target="${APP_PROPERTIES_PART}"/>` : "") + "</Relationships>";
}
/**
* The Media Types stream (W2-30). `noteSheets` lists the worksheets (1-based) that have notes:
* each gets an Override for its Comments part, and the VML parts are covered by a `vml` Default.
*/
function contentTypesXml(sheetCount, sharedStrings, coreProperties = false, appProperties = false, noteSheets = [], overrides = []) {
	let types = `${XML_DECL}<Types xmlns="${CONTENT_TYPES_NS}">`;
	types += "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
	types += "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
	if (noteSheets.length > 0) types += `<Default Extension="vml"${attr("ContentType", VML_CONTENT_TYPE)}/>`;
	for (const i of noteSheets) types += `<Override${attr("PartName", `/${commentsPartName(i)}`)}${attr("ContentType", COMMENTS_CONTENT_TYPE)}/>`;
	types += `<Override PartName="/xl/workbook.xml" ContentType="${CT_BASE}sheet.main+xml"/>`;
	for (let i = 1; i <= sheetCount; i++) types += `<Override PartName="/xl/worksheets/sheet${String(i)}.xml" ContentType="${CT_BASE}worksheet+xml"/>`;
	types += `<Override PartName="/xl/styles.xml" ContentType="${CT_BASE}styles+xml"/>`;
	if (sharedStrings) types += `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_BASE}sharedStrings+xml"/>`;
	if (coreProperties) types += `<Override PartName="/${CORE_PROPERTIES_PART}" ContentType="${CORE_PROPERTIES_CONTENT_TYPE}"/>`;
	if (appProperties) types += `<Override PartName="/${APP_PROPERTIES_PART}" ContentType="${APP_PROPERTIES_CONTENT_TYPE}"/>`;
	for (const [part, type] of overrides) types += `<Override${attr("PartName", `/${part}`)}${attr("ContentType", type)}/>`;
	return `${types}</Types>`;
}
/** A shared string table that grows as strings are added; `xml()` writes the part. */
function createSharedStrings() {
	return new SharedStrings();
}
// src/sheet-rules.ts
const DV_TYPES = [
	"none",
	"whole",
	"decimal",
	"list",
	"date",
	"time",
	"textLength",
	"custom"
];
const DV_OPERATORS = [
	"between",
	"notBetween",
	"equal",
	"notEqual",
	"lessThan",
	"lessThanOrEqual",
	"greaterThan",
	"greaterThanOrEqual"
];
const DV_ERROR_STYLES = [
	"stop",
	"warning",
	"information"
];
const DV_IME_MODES = [
	"noControl",
	"off",
	"on",
	"disabled",
	"hiragana",
	"fullKatakana",
	"halfKatakana",
	"fullAlpha",
	"halfAlpha",
	"fullHangul",
	"halfHangul"
];
const CF_TYPES = [
	"expression",
	"cellIs",
	"colorScale",
	"dataBar",
	"iconSet",
	"top10",
	"uniqueValues",
	"duplicateValues",
	"containsText",
	"notContainsText",
	"beginsWith",
	"endsWith",
	"containsBlanks",
	"notContainsBlanks",
	"containsErrors",
	"notContainsErrors",
	"timePeriod",
	"aboveAverage"
];
const CF_OPERATORS = [
	"lessThan",
	"lessThanOrEqual",
	"equal",
	"notEqual",
	"greaterThanOrEqual",
	"greaterThan",
	"between",
	"notBetween",
	"containsText",
	"notContains",
	"beginsWith",
	"endsWith"
];
const TEXT_OPERATORS = [
	"containsText",
	"notContains",
	"beginsWith",
	"endsWith"
];
const TIME_PERIODS = [
	"today",
	"yesterday",
	"tomorrow",
	"last7Days",
	"thisMonth",
	"lastMonth",
	"nextMonth",
	"thisWeek",
	"lastWeek",
	"nextWeek"
];
/** ST_CfvoType of Part 1 (§18.18.13); the main `cfvo` element cannot hold `autoMin` / `autoMax`. */
const CFVO_TYPES = [
	"num",
	"percent",
	"max",
	"min",
	"formula",
	"percentile"
];
const ICON_SETS = [
	"3Arrows",
	"3ArrowsGray",
	"3Flags",
	"3TrafficLights1",
	"3TrafficLights2",
	"3Signs",
	"3Symbols",
	"3Symbols2",
	"4Arrows",
	"4ArrowsGray",
	"4RedToBlack",
	"4Rating",
	"4TrafficLights",
	"5Arrows",
	"5ArrowsGray",
	"5Rating",
	"5Quarters"
];
/** The icon sets that exist only in the [MS-XLSX] extension (§2.7.10, SR-43). */
const X14_ICON_SETS = [
	"3Stars",
	"3Triangles",
	"5Boxes",
	"NoIcons"
];
const ALL_ICON_SETS = [...ICON_SETS, ...X14_ICON_SETS];
const X14_CFVO_TYPES = [
	"num",
	"percent",
	"max",
	"min",
	"formula",
	"percentile",
	"autoMin",
	"autoMax"
];
const BAR_DIRECTIONS = [
	"context",
	"leftToRight",
	"rightToLeft"
];
const AXIS_POSITIONS = [
	"automatic",
	"middle",
	"none"
];
/**
* The number of icons of an icon set, taken from the digit its name starts with. [MS-XLSX] does
* not state it for 3Stars, 3Triangles and 5Boxes (SR-43); NoIcons has one "no icon".
*/
const iconCount = (set) => set === "NoIcons" ? 1 : Number(set[0]);
/** Types whose `operator` gives the comparison and whose `formula2` is the upper bound (SR-03). */
const DV_COMPARING = [
	"whole",
	"decimal",
	"date",
	"time",
	"textLength"
];
/** Rule types whose criteria are given entirely by formulas (SR-11; generated formulas: C2). */
const CF_NEEDS_FORMULA = [
	"expression",
	"cellIs",
	"containsText",
	"notContainsText",
	"beginsWith",
	"endsWith",
	"containsBlanks",
	"notContainsBlanks",
	"containsErrors",
	"notContainsErrors",
	"timePeriod"
];
/** An example condition per rule type, for the hint when `formulas` is missing (A1 = top-left cell). */
const CF_FORMULA_EXAMPLE = {
	containsText: "\"ISNUMBER(SEARCH(\\\"x\\\",A1))\"",
	notContainsText: "\"ISERROR(SEARCH(\\\"x\\\",A1))\"",
	beginsWith: "\"LEFT(A1,1)=\\\"x\\\"\"",
	endsWith: "\"RIGHT(A1,1)=\\\"x\\\"\"",
	containsBlanks: "\"ISBLANK(A1)\"",
	notContainsBlanks: "\"NOT(ISBLANK(A1))\"",
	containsErrors: "\"ISERROR(A1)\"",
	notContainsErrors: "\"NOT(ISERROR(A1))\"",
	timePeriod: "\"A1=TODAY()\""
};
const CF_TEXT_TYPES = [
	"containsText",
	"notContainsText",
	"beginsWith",
	"endsWith"
];
/** Data bar properties that exist only in the [MS-XLSX] extension (§2.6.30, SR-41). */
const BAR_EXTENSION_KEYS = [
	"gradient",
	"borderColor",
	"negativeFillColor",
	"negativeBorderColor",
	"axisPosition",
	"axisColor",
	"direction"
];
const CF_GRAPHIC_TYPES = [
	"colorScale",
	"dataBar",
	"iconSet"
];
/** `prompt` of a data validation: at most 255 characters ([MS-OI29500] §2.1.613 a). */
const MAX_PROMPT = 255;
/** `error`: at most 224 characters (§2.1.613 c). */
const MAX_ERROR = 224;
/** `errorTitle` and `promptTitle`: at most 32 characters (§2.1.613 d, e). */
const MAX_TITLE = 32;
/** References in `dataValidation@sqref`: 1 to 32767 (§2.1.613 b). */
const MAX_DV_REFS = 32767;
/** References in any other ST_Sqref, such as `conditionalFormatting@sqref`: 8192 (§2.1.1092 a). */
const MAX_SQREF_REFS = 8192;
/** `dataValidation` elements of a sheet: at most 65534 (§2.1.614 g). */
const MAX_DATA_VALIDATIONS = 65534;
/** Length of an ST_Formula: at most 8192 characters (§2.1.1085 a). */
const MAX_FORMULA = 8192;
/** `dataBar@minLength` and `@maxLength`: at most 100 (§2.1.611). */
const MAX_BAR_LENGTH = 100;
/** cfvo and color children of a colour scale: 2 (schema minimum) to 3 (§2.1.601 a, b). */
const MAX_SCALE_POINTS = 3;
/** `cfRule@priority` is an xsd:int; [MS-XLSX] §2.6.27 requires it to be greater than 0. */
const MAX_PRIORITY = 2147483647;
function fail(where, option, message, hint) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		location: { sheet: where.sheet },
		hint,
		details: { option }
	});
}
function readKeys(value, known, what, where) {
	if (typeof value !== "object" || value === null || Array.isArray(value) || value instanceof Date) throw fail(where, what, `${what} must be an object.`, `Pass an object with some of these properties: ${known.join(", ")}.`);
	const out = /* @__PURE__ */ new Map();
	for (const key of Object.keys(value)) {
		if (!known.includes(key)) {
			const close = suggest(key, known);
			throw fail(where, `${what}.${key.slice(0, 40)}`, `Unknown property "${key.slice(0, 40)}" in ${what}.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known properties: ${known.join(", ")}.`);
		}
		const v = value[key];
		if (v !== void 0) out.set(key, v);
	}
	return out;
}
function enumValue(value, allowed, option, where) {
	if (typeof value === "string" && allowed.includes(value)) return value;
	const shown = typeof value === "string" ? JSON.stringify(value.slice(0, 40)) : String(value);
	const close = typeof value === "string" ? suggest(value, allowed) : [];
	throw fail(where, option, `The ${option} ${shown} is not valid.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Allowed values: ${allowed.join(", ")}.`);
}
function boolValue(value, option, where) {
	if (typeof value === "boolean") return value;
	throw fail(where, option, `The ${option} must be true or false.`, `Use ${option}: true.`);
}
function intValue(value, option, min, max, where) {
	if (typeof value === "number" && Number.isInteger(value) && value >= min && value <= max) return value;
	throw fail(where, option, `The ${option} ${String(value).slice(0, 40)} is not valid.`, `Use a whole number from ${String(min)} to ${String(max)}.`);
}
function textValue(value, option, max, where) {
	if (typeof value !== "string") throw fail(where, option, `The ${option} must be a string.`, `For example ${option}: "text".`);
	if (value.length > max) throw fail(where, option, `The ${option} has ${String(value.length)} characters; at most ${String(max)} are allowed.`, `Shorten the ${option} to ${String(max)} characters or fewer.`);
	if (firstLoneSurrogate(value) !== -1) throw fail(where, option, `The ${option} contains an unpaired surrogate.`, "Remove the broken character.");
	return value;
}
/** Formula text from a caller value: text without a leading "=", a number, or a date (serial). */
function formulaValue(value, option, where) {
	let text;
	if (typeof value === "number") {
		if (!Number.isFinite(value)) throw fail(where, option, `The ${option} must be a finite number.`, `Pass a number, or formula text such as "A1*2".`);
		text = formatNumber(value);
	} else if (value instanceof Date) {
		const time = value.getTime();
		if (Number.isNaN(time)) throw fail(where, option, `The ${option} is an invalid Date.`, "Pass a valid Date.");
		try {
			text = formatNumber(dateToSerial(value, where.dateSystem));
		} catch {
			throw fail(where, option, `The date ${value.toISOString()} in ${option} is outside the ${where.dateSystem} date system.`, "Use a date from 1900 (or 1904 with that date system) to 9999.");
		}
	} else if (typeof value === "string") {
		text = value.startsWith("=") ? value.slice(1) : value;
		if (text.trim() === "") throw fail(where, option, `The ${option} is empty.`, "Pass formula text such as \"A1>0\", a number, or \"\\\"Yes,No\\\"\" for a list.");
	} else throw fail(where, option, `The ${option} must be formula text, a number or a Date.`, "For example \"A1>0\", 10 or new Date(Date.UTC(2024, 0, 31)).");
	if (text.length > MAX_FORMULA) throw fail(where, option, `The ${option} has ${String(text.length)} characters; at most ${String(MAX_FORMULA)} are allowed.`, `Shorten the formula to ${String(MAX_FORMULA)} characters or fewer.`);
	if (firstLoneSurrogate(text) !== -1) throw fail(where, option, `The ${option} contains an unpaired surrogate.`, "Remove the broken character.");
	return text;
}
/**
* Normalises a range list (ST_Sqref, §18.18.76: space-separated references) to ranges in
* canonical form, at most `max` of them. Throws `XLV_E_INVALID_CELL_REF` for a bad range.
*/
function normalizeSqref(sqref, max, where) {
	const parts = typeof sqref === "string" ? sqref.trim().split(/\s+/) : [];
	if (parts.length === 0 || parts[0] === "") throw new XlvError("XLV_E_INVALID_CELL_REF", `"${String(sqref).slice(0, 40)}" is not a range.`, {
		location: { sheet: where.sheet },
		hint: "Pass a cell or range such as \"B2:B100\", or several separated by spaces: \"B2:B10 D2:D10\".",
		details: { ref: String(sqref).slice(0, 40) }
	});
	if (parts.length > max) throw fail(where, "sqref", `The range list has ${String(parts.length)} ranges; at most ${String(max)} are allowed.`, "Combine neighbouring cells into larger ranges, such as \"A1:A100\".");
	return parts.map((p) => {
		const range = parseRange(p.toUpperCase().replace(/\$/g, ""));
		if (range === null) throw new XlvError("XLV_E_INVALID_CELL_REF", `"${p.slice(0, 40)}" is not a valid range.`, {
			location: { sheet: where.sheet },
			hint: "Use A1-style ranges such as \"B2:B100\" (columns A to XFD, rows 1 to 1048576), separated by spaces.",
			details: { ref: p.slice(0, 40) }
		});
		return formatRange(range);
	}).join(" ");
}
const DV_KEYS = [
	"type",
	"operator",
	"formula1",
	"formula2",
	"allowBlank",
	"showDropDown",
	"showInputMessage",
	"showErrorMessage",
	"errorStyle",
	"imeMode",
	"errorTitle",
	"error",
	"promptTitle",
	"prompt"
];
/** Validates a data validation rule (SR-02 … SR-06); throws `XLV_E_INVALID_OPTION`. */
function normalizeDataValidation(sqref, rule, where) {
	const range = normalizeSqref(sqref, MAX_DV_REFS, where);
	const o = readKeys(rule, DV_KEYS, "data validation", where);
	if (!o.has("type")) throw fail(where, "type", "The data validation has no type.", `Pass type: one of ${DV_TYPES.join(", ")}.`);
	const type = enumValue(o.get("type"), DV_TYPES, "type", where);
	const comparing = DV_COMPARING.includes(type);
	let operator = null;
	if (o.has("operator")) {
		if (!comparing) throw fail(where, "operator", `A data validation of type "${type}" takes no operator.`, `Remove the operator; it applies to ${DV_COMPARING.join(", ")}.`);
		operator = enumValue(o.get("operator"), DV_OPERATORS, "operator", where);
	} else if (comparing) operator = "between";
	const f1 = o.get("formula1");
	const f2 = o.get("formula2");
	if (type === "none") {
		if (f1 !== void 0 || f2 !== void 0) throw fail(where, f1 !== void 0 ? "formula1" : "formula2", "A data validation of type \"none\" takes no formula1 or formula2.", "Remove the formulas, or choose a type such as \"list\" or \"whole\".");
	} else if (f1 === void 0) throw fail(where, "formula1", `A data validation of type "${type}" needs formula1.`, type === "list" ? "For a fixed list pass formula1: \"\\\"Yes,No\\\"\", for a list from cells formula1: \"Lists!$A$1:$A$5\"." : "Pass the value or formula to compare with, for example formula1: 10.");
	const ranged = operator === "between" || operator === "notBetween";
	if (ranged && f2 === void 0) throw fail(where, "formula2", `The operator "${String(operator)}" needs formula2, the upper bound.`, "Pass formula2, or choose another operator such as \"greaterThan\".");
	if (!ranged && f2 !== void 0) throw fail(where, "formula2", `formula2 is used only with the operators "between" and "notBetween".`, "Remove formula2, or use operator: \"between\".");
	const flag = (k) => o.has(k) ? boolValue(o.get(k), k, where) : false;
	const text = (k, max) => o.has(k) ? textValue(o.get(k), k, max, where) : null;
	return Object.freeze({
		sqref: range,
		type,
		operator,
		formula1: f1 === void 0 ? null : formulaValue(f1, "formula1", where),
		formula2: f2 === void 0 ? null : formulaValue(f2, "formula2", where),
		allowBlank: flag("allowBlank"),
		showDropDown: flag("showDropDown"),
		showInputMessage: flag("showInputMessage"),
		showErrorMessage: flag("showErrorMessage"),
		errorStyle: o.has("errorStyle") ? enumValue(o.get("errorStyle"), DV_ERROR_STYLES, "errorStyle", where) : "stop",
		imeMode: o.has("imeMode") ? enumValue(o.get("imeMode"), DV_IME_MODES, "imeMode", where) : "noControl",
		errorTitle: text("errorTitle", MAX_TITLE),
		error: text("error", MAX_ERROR),
		promptTitle: text("promptTitle", MAX_TITLE),
		prompt: text("prompt", MAX_PROMPT)
	});
}
const CF_KEYS = [
	"type",
	"priority",
	"stopIfTrue",
	"style",
	"formulas",
	"operator",
	"text",
	"timePeriod",
	"rank",
	"percent",
	"bottom",
	"aboveAverage",
	"equalAverage",
	"stdDev",
	"colorScale",
	"dataBar",
	"iconSet"
];
function cfValues(value, option, min, max, where, inIconSet, allowAuto = false) {
	if (!Array.isArray(value) || value.length < min || value.length > max) throw fail(where, option, `${option} must be an array of ${min === max ? String(min) : `${String(min)} to ${String(max)}`} value objects.`, "For example [{ type: \"min\" }, { type: \"max\" }].");
	return value.map((v, i) => {
		const o = readKeys(v, [
			"type",
			"value",
			"gte"
		], `${option}[${String(i)}]`, where);
		const type = enumValue(o.get("type"), allowAuto ? X14_CFVO_TYPES : CFVO_TYPES, `${option}[${String(i)}].type`, where);
		if (inIconSet && (type === "min" || type === "max")) throw fail(where, `${option}[${String(i)}].type`, `An icon set threshold cannot be of type "${type}".`, "Use { type: \"percent\", value: 0 } for the lowest threshold.");
		const raw = o.get("value");
		if ((type === "num" || type === "percent" || type === "percentile" || type === "formula") && raw === void 0) throw fail(where, `${option}[${String(i)}].value`, `A value object of type "${type}" needs a value.`, `For example { type: "${type}", value: ${type === "formula" ? "\"$A$1\"" : "50"} }.`);
		const out = Object.freeze({
			type,
			value: raw === void 0 ? null : formulaValue(raw, `${option}[${String(i)}].value`, where),
			gte: o.has("gte") ? boolValue(o.get("gte"), `${option}[${String(i)}].gte`, where) : true
		});
		if (!inIconSet && !out.gte) throw fail(where, `${option}[${String(i)}].gte`, "gte applies to icon sets only.", "Remove gte.");
		return out;
	});
}
/** Validates one conditional formatting rule (SR-08 … SR-15); `priority` is assigned by the caller. */
function normalizeCfRule(rule, index, where) {
	const what = `rules[${String(index)}]`;
	const o = readKeys(rule, CF_KEYS, what, where);
	const type = enumValue(o.get("type"), CF_TYPES, `${what}.type`, where);
	const only = (key, types) => {
		if (!o.has(key)) return false;
		if (!types.includes(type)) throw fail(where, `${what}.${key}`, `The rule property "${key}" does not apply to rules of type "${type}".`, `Remove "${key}"; it applies to ${types.join(", ")}.`);
		return true;
	};
	const graphic = CF_GRAPHIC_TYPES.includes(type);
	let style = null;
	if (o.has("style")) {
		if (graphic) throw fail(where, `${what}.style`, `A rule of type "${type}" takes no style.`, "Remove the style; colour scales, data bars and icon sets draw their own colours.");
		const raw = o.get("style");
		if (typeof raw === "object" && raw !== null) {
			for (const k of Object.keys(raw)) if (k !== "font" && k !== "fill" && k !== "border" && k !== "numFmt") throw fail(where, `${what}.style.${k.slice(0, 40)}`, `The style of a conditional format cannot set "${k.slice(0, 40)}".`, "Use font, fill, border and numFmt.");
		}
		style = normalizeStyle(raw);
	}
	let formulas = [];
	if (o.has("formulas")) {
		if (graphic) throw fail(where, `${what}.formulas`, `A rule of type "${type}" takes no formulas.`, "Give the thresholds in the cfvo value objects instead.");
		const raw = o.get("formulas");
		if (!Array.isArray(raw) || raw.length < 1 || raw.length > 3) throw fail(where, `${what}.formulas`, "formulas must be an array of 1 to 3 formulas.", "For example formulas: [\"$B2>100\"].");
		formulas = raw.map((f, i) => formulaValue(f, `${what}.formulas[${String(i)}]`, where));
	} else if (CF_NEEDS_FORMULA.includes(type)) throw fail(where, `${what}.formulas`, `A rule of type "${type}" needs its criteria in formulas.`, type === "cellIs" ? "For example formulas: [100] with operator: \"greaterThan\"." : `Pass the condition as a formula for the top-left cell of the range, for example formulas: [${CF_FORMULA_EXAMPLE[type] ?? "\"$B2>100\""}].`);
	let operator = null;
	if (type === "cellIs") {
		if (!o.has("operator")) throw fail(where, `${what}.operator`, "A rule of type \"cellIs\" needs an operator.", `Use one of ${CF_OPERATORS.slice(0, 8).join(", ")}.`);
		operator = enumValue(o.get("operator"), CF_OPERATORS.slice(0, 8), `${what}.operator`, where);
	} else if (only("operator", CF_TEXT_TYPES)) operator = enumValue(o.get("operator"), TEXT_OPERATORS, `${what}.operator`, where);
	let timePeriod = null;
	if (type === "timePeriod") {
		if (!o.has("timePeriod")) throw fail(where, `${what}.timePeriod`, "A rule of type \"timePeriod\" needs timePeriod.", `Use one of ${TIME_PERIODS.join(", ")}.`);
		timePeriod = enumValue(o.get("timePeriod"), TIME_PERIODS, `${what}.timePeriod`, where);
	} else only("timePeriod", ["timePeriod"]);
	const text = only("text", CF_TEXT_TYPES) ? textValue(o.get("text"), `${what}.text`, MAX_FORMULA, where) : null;
	if (text !== null && /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(text)) throw fail(where, `${what}.text`, `The ${what}.text contains a control character or non-character that XML cannot carry.`, "Remove the character; tab, line feed and carriage return are allowed.");
	const rank = only("rank", ["top10"]) ? intValue(o.get("rank"), `${what}.rank`, 0, 4294967295, where) : null;
	const flag = (k, types, dflt) => only(k, types) ? boolValue(o.get(k), `${what}.${k}`, where) : dflt;
	const percent = flag("percent", ["top10"], false);
	const bottom = flag("bottom", ["top10"], false);
	const aboveAverage = flag("aboveAverage", ["aboveAverage"], true);
	const equalAverage = flag("equalAverage", ["aboveAverage"], false);
	const stdDev = only("stdDev", ["aboveAverage"]) ? intValue(o.get("stdDev"), `${what}.stdDev`, -2147483648, 2147483647, where) : null;
	for (const key of CF_GRAPHIC_TYPES) if (type !== key) only(key, [key]);
	else if (!o.has(key)) throw fail(where, `${what}.${key}`, `A rule of type "${type}" needs ${key}.`, `Pass ${key}: { cfvo: [...], ... } (see ConditionalFormatRuleInput).`);
	let colorScale = null;
	if (type === "colorScale") {
		const c = readKeys(o.get("colorScale"), ["cfvo", "colors"], `${what}.colorScale`, where);
		const cfvo = cfValues(c.get("cfvo"), `${what}.colorScale.cfvo`, 2, MAX_SCALE_POINTS, where, false);
		const colors = c.get("colors");
		if (!Array.isArray(colors) || colors.length !== cfvo.length) throw fail(where, `${what}.colorScale.colors`, `colorScale.colors must have one colour per value object (${String(cfvo.length)}).`, "For example colors: [\"F8696B\", \"63BE7B\"] for two value objects.");
		colorScale = Object.freeze({
			cfvo,
			colors: colors.map((x, i) => normalizeColor(x, `${what}.colorScale.colors[${String(i)}]`))
		});
	}
	let dataBar = null;
	if (type === "dataBar") {
		const d = readKeys(o.get("dataBar"), [
			"cfvo",
			"color",
			"minLength",
			"maxLength",
			"showValue",
			...BAR_EXTENSION_KEYS
		], `${what}.dataBar`, where);
		const cfvo = cfValues(d.get("cfvo"), `${what}.dataBar.cfvo`, 2, 2, where, false, true);
		if (!d.has("color")) throw fail(where, `${what}.dataBar.color`, "A data bar needs a color.", "For example color: \"638EC6\".");
		const len = (k, dflt) => d.has(k) ? intValue(d.get(k), `${what}.dataBar.${k}`, 0, MAX_BAR_LENGTH, where) : dflt;
		const color = (k) => d.has(k) ? normalizeColor(d.get(k), `${what}.dataBar.${k}`) : null;
		const extended = BAR_EXTENSION_KEYS.some((k) => d.has(k)) || cfvo.some((v) => v.type === "autoMin" || v.type === "autoMax");
		const minLength = len("minLength", 10);
		const maxLength = len("maxLength", 90);
		const borderColor = color("borderColor");
		const negativeBorderColor = color("negativeBorderColor");
		const axisPosition = d.has("axisPosition") ? enumValue(d.get("axisPosition"), AXIS_POSITIONS, `${what}.dataBar.axisPosition`, where) : "automatic";
		const axisColor = color("axisColor");
		if (extended) {
			const [first, last] = cfvo;
			if (first.type === "max" || first.type === "autoMax") throw fail(where, `${what}.dataBar.cfvo[0].type`, `The first value object of a data bar cannot be "${first.type}".`, "Use \"min\", \"autoMin\", \"num\", \"percent\", \"percentile\" or \"formula\".");
			if (last.type === "min" || last.type === "autoMin") throw fail(where, `${what}.dataBar.cfvo[1].type`, `The second value object of a data bar cannot be "${last.type}".`, "Use \"max\", \"autoMax\", \"num\", \"percent\", \"percentile\" or \"formula\".");
			if (minLength > maxLength) throw fail(where, `${what}.dataBar.minLength`, `minLength (${String(minLength)}) is greater than maxLength (${String(maxLength)}).`, "Give a minLength from 0 to maxLength.");
			if (negativeBorderColor !== null && borderColor === null) throw fail(where, `${what}.dataBar.negativeBorderColor`, "negativeBorderColor needs borderColor: negative bars have a border only when the bars have one.", "Add borderColor, for example borderColor: \"638EC6\".");
			if (axisPosition !== "none" && axisColor === null) throw fail(where, `${what}.dataBar.axisColor`, "A data bar with an axis needs axisColor.", "Add axisColor, for example axisColor: \"000000\", or set axisPosition: \"none\".");
		}
		if (!extended && cfvo.some((v) => v.type === "autoMin" || v.type === "autoMax")) throw fail(where, `${what}.dataBar.cfvo`, "autoMin and autoMax need the data bar extension.", "Use \"min\" and \"max\".");
		dataBar = Object.freeze({
			cfvo,
			color: normalizeColor(d.get("color"), `${what}.dataBar.color`),
			minLength,
			maxLength,
			showValue: d.has("showValue") ? boolValue(d.get("showValue"), `${what}.dataBar.showValue`, where) : true,
			extended,
			gradient: d.has("gradient") ? boolValue(d.get("gradient"), `${what}.dataBar.gradient`, where) : true,
			border: borderColor !== null,
			borderColor,
			negativeFillColor: color("negativeFillColor"),
			negativeBorderColor,
			axisPosition,
			axisColor: axisPosition === "none" ? null : axisColor,
			direction: d.has("direction") ? enumValue(d.get("direction"), BAR_DIRECTIONS, `${what}.dataBar.direction`, where) : "context"
		});
	}
	let iconSet = null;
	if (type === "iconSet") {
		const s = readKeys(o.get("iconSet"), [
			"iconSet",
			"cfvo",
			"reverse",
			"showValue",
			"percent",
			"custom"
		], `${what}.iconSet`, where);
		const b = (k, dflt) => s.has(k) ? boolValue(s.get(k), `${what}.iconSet.${k}`, where) : dflt;
		const set = s.has("iconSet") ? enumValue(s.get("iconSet"), ALL_ICON_SETS, `${what}.iconSet.iconSet`, where) : "3TrafficLights1";
		const extended = X14_ICON_SETS.includes(set) || s.has("custom");
		const cfvo = cfValues(s.get("cfvo"), `${what}.iconSet.cfvo`, extended ? 3 : 2, 5, where, true);
		let custom = null;
		if (s.has("custom")) {
			const raw = s.get("custom");
			if (!Array.isArray(raw) || raw.length !== cfvo.length) throw fail(where, `${what}.iconSet.custom`, `custom must list one icon per value object (${String(cfvo.length)}).`, "For example custom: [{ iconSet: '3Flags', iconId: 0 }, …].");
			custom = raw.map((icon, i) => {
				const at = `${what}.iconSet.custom[${String(i)}]`;
				const c = readKeys(icon, ["iconSet", "iconId"], at, where);
				const iconSetOf = enumValue(c.get("iconSet"), ALL_ICON_SETS, `${at}.iconSet`, where);
				return Object.freeze({
					iconSet: iconSetOf,
					iconId: intValue(c.get("iconId"), `${at}.iconId`, 0, iconCount(iconSetOf) - 1, where)
				});
			});
		}
		iconSet = Object.freeze({
			iconSet: set,
			cfvo,
			reverse: b("reverse", false),
			showValue: b("showValue", true),
			percent: b("percent", true),
			custom: custom === null ? null : Object.freeze(custom)
		});
	}
	return {
		type,
		priority: o.has("priority") ? intValue(o.get("priority"), `${what}.priority`, 1, MAX_PRIORITY, where) : null,
		stopIfTrue: o.has("stopIfTrue") ? boolValue(o.get("stopIfTrue"), `${what}.stopIfTrue`, where) : false,
		style,
		formulas,
		operator,
		text,
		timePeriod,
		rank,
		percent,
		bottom,
		aboveAverage,
		equalAverage,
		stdDev,
		colorScale,
		dataBar,
		iconSet
	};
}
/**
* Validates the rules of one `conditionalFormat` call and gives each a priority: the given one,
* which must not be used on the sheet yet, or the next number after the highest one used.
*/
function normalizeConditionalFormat(sqref, rules, where, used) {
	const range = normalizeSqref(sqref, MAX_SQREF_REFS, where);
	if (!Array.isArray(rules) || rules.length === 0) throw fail(where, "rules", "conditionalFormat needs an array with at least one rule.", "For example [{ type: \"cellIs\", operator: \"greaterThan\", formulas: [100], style: { font: { bold: true } } }].");
	const normalized = rules.map((r, i) => normalizeCfRule(r, i, where));
	const taken = new Set(used);
	for (const [i, r] of normalized.entries()) {
		if (r.priority === null) continue;
		if (taken.has(r.priority)) throw fail(where, `rules[${String(i)}].priority`, `The priority ${String(r.priority)} is already used on the sheet.`, "Leave priority out to get the next free one, or choose a number not used yet.");
		taken.add(r.priority);
	}
	let next = Math.max(0, ...taken) + 1;
	const out = normalized.map((r) => {
		if (r.priority !== null) return Object.freeze({
			...r,
			priority: r.priority
		});
		if (next > MAX_PRIORITY) throw fail(where, "priority", "No priority is left on the sheet.", "Give the rules priorities explicitly.");
		taken.add(next);
		return Object.freeze({
			...r,
			priority: next++
		});
	});
	for (const p of taken) used.add(p);
	return {
		sqref: range,
		rules: out
	};
}
/** Namespace of the Excel 2010 extensions, prefix `x14` ([MS-XLSX] §2.2, §5.3; SR-40). */
const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
/** Namespace of `f`, `sqref` and `ref` of the extensions, prefix `xm` ([MS-XLSX] §2.4.3–§2.4.5, §5.1). */
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
/** `ext@uri` of the worksheet extension with `x14:conditionalFormattings` ([MS-XLSX] §2.2.4.12). */
const X14_CF_URI = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}";
/** `ext@uri` of the worksheet extension with `x14:dataValidations` (§2.2.4.12). */
const X14_DV_URI = "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}";
/** `ext@uri` of the cfRule extension with `x14:id` (§2.2.4.12). */
const X14_ID_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
const parseBool = (raw, dflt) => {
	if (raw === void 0) return dflt;
	const s = raw.trim();
	if (s === "true" || s === "1") return true;
	if (s === "false" || s === "0") return false;
	return dflt;
};
const parseInteger = (raw, min, max) => {
	if (raw === void 0 || !/^\s*[-+]?\d{1,10}\s*$/.test(raw)) return null;
	const n = Number(raw.trim());
	return n >= min && n <= max ? n : null;
};
/** A space-separated range list as read, canonical; `null` (with a warning) when invalid or too long. */
function readSqref(raw, max, what, warn) {
	const parts = (raw ?? "").trim().split(/\s+/).filter((p) => p !== "");
	if (parts.length === 0) {
		warn(`A ${what} has no range (sqref); it is ignored.`);
		return null;
	}
	if (parts.length > max) {
		warn(`A ${what} lists ${String(parts.length)} ranges, more than the ${String(max)} Office allows; it is ignored.`);
		return null;
	}
	const out = [];
	for (const p of parts) {
		const r = parseRange(p.replace(/\$/g, ""));
		if (r === null) {
			warn(`A ${what} has the invalid range "${p.slice(0, 40)}"; it is ignored.`);
			return null;
		}
		out.push(formatRange(r));
	}
	return out.join(" ");
}
/** The text of the `xm:f` child of an element, decoded; `null` without one. */
function xmFormula(el) {
	const f = el === void 0 ? void 0 : childrenNamed(el, XM_NS, "f")[0];
	return f === void 0 ? null : decodeXstring(f.text);
}
/**
* One data validation from its element: `sqref` is the range list, `formula(name)` the text of
* `formula1` / `formula2` (main element or [MS-XLSX] `xm:f`). `null`, with a warning, when the
* range or the type is not valid.
*/
function readDataValidation(dv, rawSqref, formula, warn) {
	const sqref = readSqref(rawSqref, MAX_DV_REFS, "data validation", warn);
	if (sqref === null) return null;
	const rawType = dv.attrs.get("type") ?? "none";
	const type = DV_TYPES.find((t) => t === rawType);
	if (type === void 0) {
		warn(`A data validation has the unknown type "${rawType.slice(0, 40)}"; it is ignored.`);
		return null;
	}
	const pick = (name, allowed, dflt) => {
		const raw = dv.attrs.get(name);
		if (raw === void 0) return dflt;
		const found = allowed.find((a) => a === raw);
		if (found === void 0) warn(`A data validation has the unknown ${name} "${raw.slice(0, 40)}"; "${dflt}" is used.`);
		return found ?? dflt;
	};
	const text = (name) => {
		const raw = dv.attrs.get(name);
		return raw === void 0 ? null : decodeXstring(raw);
	};
	const operator = DV_COMPARING.includes(type) ? pick("operator", DV_OPERATORS, "between") : null;
	return Object.freeze({
		sqref,
		type,
		operator,
		formula1: type === "none" ? null : formula("formula1"),
		formula2: operator === "between" || operator === "notBetween" ? formula("formula2") : null,
		allowBlank: parseBool(dv.attrs.get("allowBlank"), false),
		showDropDown: parseBool(dv.attrs.get("showDropDown"), false),
		showInputMessage: parseBool(dv.attrs.get("showInputMessage"), false),
		showErrorMessage: parseBool(dv.attrs.get("showErrorMessage"), false),
		errorStyle: pick("errorStyle", DV_ERROR_STYLES, "stop"),
		imeMode: pick("imeMode", DV_IME_MODES, "noControl"),
		errorTitle: text("errorTitle"),
		error: text("error"),
		promptTitle: text("promptTitle"),
		prompt: text("prompt")
	});
}
/** Reads `dataValidations` (§18.3.1.33); invalid entries are skipped with a warning. */
function parseDataValidations(el, ns, warn) {
	const out = [];
	for (const dv of childrenNamed(el, ns, "dataValidation")) {
		if (out.length >= 65534) {
			warn(`The sheet has more than ${String(MAX_DATA_VALIDATIONS)} data validations; the rest are ignored.`);
			break;
		}
		const read = readDataValidation(dv, dv.attrs.get("sqref"), (name) => {
			const f = childrenNamed(dv, ns, name)[0];
			return f === void 0 ? null : decodeXstring(f.text);
		}, warn);
		if (read !== null) out.push(read);
	}
	return out;
}
/**
* Reads `x14:dataValidations` ([MS-XLSX] §2.6.3–§2.6.5, SR-44): the same attributes as the main
* element, formulas in `x14:formula1/xm:f` and the range in `xm:sqref`.
*/
function parseX14DataValidations(el, warn) {
	const out = [];
	for (const dv of childrenNamed(el, X14_NS, "dataValidation")) {
		if (out.length >= 65534) {
			warn(`The sheet has more than ${String(MAX_DATA_VALIDATIONS)} data validations; the rest are ignored.`);
			break;
		}
		const sqref = childrenNamed(dv, XM_NS, "sqref")[0]?.text;
		const read = readDataValidation(dv, sqref, (name) => xmFormula(childrenNamed(dv, X14_NS, name)[0]), warn);
		if (read !== null) out.push(read);
	}
	return out;
}
/**
* The value objects of a main (`x14` false) or extension `cfvo` list. Main values are in `val`,
* extension values in `xm:f` ([MS-XLSX] §2.6.26). `null`, with a warning, for an unknown type.
*/
function readCfValues(parent, ns, warn, x14 = false) {
	const out = [];
	for (const v of childrenNamed(parent, ns, "cfvo")) {
		const raw = v.attrs.get("type") ?? "";
		const type = (x14 ? X14_CFVO_TYPES : CFVO_TYPES).find((t) => t === raw);
		if (type === void 0) {
			warn(`A conditional format value object has the unknown type "${raw.slice(0, 40)}"; the rule is ignored.`);
			return null;
		}
		const val = x14 ? xmFormula(v) ?? void 0 : v.attrs.get("val");
		const unused = type === "min" || type === "max" || type === "autoMin" || type === "autoMax";
		out.push(Object.freeze({
			type,
			value: val === void 0 || unused ? null : x14 ? val : decodeXstring(val),
			gte: parseBool(v.attrs.get("gte"), true)
		}));
	}
	return out;
}
function readColors(parent, ns, warn, readColor) {
	const out = [];
	for (const c of childrenNamed(parent, ns, "color")) {
		const color = readColor(c);
		if (color === void 0) {
			warn("A conditional format colour is not valid; the rule is ignored.");
			return null;
		}
		out.push(color);
	}
	return out;
}
/** The extension fields of a data bar without extension data (the [MS-XLSX] schema defaults). */
const PLAIN_BAR = {
	extended: false,
	gradient: true,
	border: false,
	borderColor: null,
	negativeFillColor: null,
	negativeBorderColor: null,
	axisPosition: "automatic",
	axisColor: null,
	direction: "context"
};
/** Reads `x14:dataBar` (§2.6.30); `fillColor` is the bar colour of a rule with its own priority. */
function readX14DataBar(db, warn, readColor) {
	const cfvo = readCfValues(db, X14_NS, warn, true);
	if (cfvo === null) return null;
	const color = (name) => {
		const el = childrenNamed(db, X14_NS, name)[0];
		return el === void 0 ? null : readColor(el) ?? null;
	};
	const pick = (name, allowed, dflt) => allowed.find((a) => a === db.attrs.get(name)) ?? dflt;
	const border = parseBool(db.attrs.get("border"), false);
	const axisPosition = pick("axisPosition", AXIS_POSITIONS, "automatic");
	const negativeSame = parseBool(db.attrs.get("negativeBarColorSameAsPositive"), false);
	const negativeBorderSame = parseBool(db.attrs.get("negativeBarBorderColorSameAsPositive"), true);
	return {
		cfvo,
		minLength: parseInteger(db.attrs.get("minLength"), 0, MAX_BAR_LENGTH) ?? 10,
		maxLength: parseInteger(db.attrs.get("maxLength"), 0, MAX_BAR_LENGTH) ?? 90,
		showValue: parseBool(db.attrs.get("showValue"), true),
		gradient: parseBool(db.attrs.get("gradient"), true),
		border,
		borderColor: border ? color("borderColor") : null,
		negativeFillColor: negativeSame ? null : color("negativeFillColor"),
		negativeBorderColor: negativeBorderSame || !border ? null : color("negativeBorderColor"),
		axisPosition,
		axisColor: axisPosition === "none" ? null : color("axisColor"),
		direction: pick("direction", BAR_DIRECTIONS, "context"),
		fill: color("fillColor")
	};
}
/** Reads `x14:iconSet` (§2.6.28, §2.6.36). */
function readX14IconSet(is, warn) {
	const cfvo = readCfValues(is, X14_NS, warn, true);
	if (cfvo === null) return null;
	const raw = is.attrs.get("iconSet");
	const set = raw === void 0 ? "3TrafficLights1" : ALL_ICON_SETS.find((x) => x === raw);
	if (set === void 0) {
		warn(`A conditional formatting rule has the unknown icon set "${String(raw).slice(0, 40)}"; it is ignored.`);
		return null;
	}
	const custom = parseBool(is.attrs.get("custom"), false);
	const icons = [];
	if (custom) for (const c of childrenNamed(is, X14_NS, "cfIcon")) {
		const ofSet = ALL_ICON_SETS.find((x) => x === c.attrs.get("iconSet"));
		const id = ofSet === void 0 ? null : parseInteger(c.attrs.get("iconId"), 0, iconCount(ofSet) - 1);
		if (ofSet === void 0 || id === null) {
			warn("A custom icon of an icon set is not valid; the rule is ignored.");
			return null;
		}
		icons.push(Object.freeze({
			iconSet: ofSet,
			iconId: id
		}));
	}
	return Object.freeze({
		iconSet: set,
		cfvo,
		reverse: parseBool(is.attrs.get("reverse"), false),
		showValue: parseBool(is.attrs.get("showValue"), true),
		percent: true,
		custom: custom ? Object.freeze(icons) : null
	});
}
/** The attributes of a cfRule element that depend on its type (§18.3.1.10, SR-09). */
function ruleFromAttributes(r, type, priority, style, formulas, graphics, warn) {
	const enumAttr = (name, allowed) => {
		const raw = r.attrs.get(name);
		if (raw === void 0) return null;
		const found = allowed.find((a) => a === raw);
		if (found === void 0) warn(`A conditional formatting rule has the unknown ${name} "${raw.slice(0, 40)}"; it is ignored.`);
		return found ?? null;
	};
	const top = type === "top10";
	const avg = type === "aboveAverage";
	const textType = CF_TEXT_TYPES.includes(type);
	const stdDev = avg ? r.attrs.get("stdDev") : void 0;
	return Object.freeze({
		type,
		priority,
		stopIfTrue: parseBool(r.attrs.get("stopIfTrue"), false),
		style,
		formulas: Object.freeze(formulas.slice(0, 3)),
		operator: type === "cellIs" || textType ? enumAttr("operator", CF_OPERATORS) : null,
		text: textType ? r.attrs.get("text") ?? null : null,
		timePeriod: type === "timePeriod" ? enumAttr("timePeriod", TIME_PERIODS) : null,
		rank: top ? parseInteger(r.attrs.get("rank"), 0, 4294967295) : null,
		percent: top && parseBool(r.attrs.get("percent"), false),
		bottom: top && parseBool(r.attrs.get("bottom"), false),
		aboveAverage: !avg || parseBool(r.attrs.get("aboveAverage"), true),
		equalAverage: avg && parseBool(r.attrs.get("equalAverage"), false),
		stdDev: stdDev === void 0 ? null : parseInteger(stdDev, -2147483648, 2147483647),
		...graphics
	});
}
/** The `x14:id` in a main cfRule's `extLst` (§2.2.4.12, §2.6.30), upper-cased; `null` without one. */
function extensionId(r, ns) {
	for (const list of childrenNamed(r, ns, "extLst")) for (const ext of childrenNamed(list, ns, "ext")) {
		if (ext.attrs.get("uri")?.toUpperCase() !== "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}") continue;
		const id = childrenNamed(ext, X14_NS, "id")[0]?.text.trim();
		if (id !== void 0 && id !== "") return id.toUpperCase();
	}
	return null;
}
function readConditionalFormatting(el, ns, dxf, readColor, warn) {
	const sqref = readSqref(el.attrs.get("sqref"), MAX_SQREF_REFS, "conditional format", warn);
	if (sqref === null) return null;
	const rules = [];
	const ids = [];
	for (const r of childrenNamed(el, ns, "cfRule")) {
		const rawType = r.attrs.get("type") ?? "";
		const type = CF_TYPES.find((t) => t === rawType);
		if (type === void 0) {
			warn(`A conditional formatting rule has the unknown type "${rawType.slice(0, 40)}"; it is ignored.`);
			continue;
		}
		const priority = parseInteger(r.attrs.get("priority"), -2147483648, MAX_PRIORITY);
		if (priority === null) {
			warn("A conditional formatting rule has no valid priority; it is ignored.");
			continue;
		}
		let style = null;
		const rawDxf = r.attrs.get("dxfId");
		if (rawDxf !== void 0) {
			const id = parseInteger(rawDxf, 0, 4294967294);
			const found = id === null ? void 0 : dxf(id);
			if (found === void 0) warn(`A conditional formatting rule refers to the differential format "${rawDxf.slice(0, 20)}", which does not exist; the rule is read without a format.`);
			else style = found;
		}
		let colorScale = null;
		let dataBar = null;
		let iconSet = null;
		const sub = (name) => childrenNamed(r, ns, name)[0];
		const cs = type === "colorScale" ? sub("colorScale") : void 0;
		const db = type === "dataBar" ? sub("dataBar") : void 0;
		const is = type === "iconSet" ? sub("iconSet") : void 0;
		if (cs !== void 0) {
			const cfvo = readCfValues(cs, ns, warn);
			const colors = readColors(cs, ns, warn, readColor);
			if (cfvo === null || colors === null) continue;
			colorScale = Object.freeze({
				cfvo,
				colors
			});
		} else if (db !== void 0) {
			const cfvo = readCfValues(db, ns, warn);
			const colors = readColors(db, ns, warn, readColor);
			if (cfvo === null || colors === null) continue;
			const color = colors[0];
			if (color === void 0) {
				warn("A data bar has no colour; the rule is ignored.");
				continue;
			}
			dataBar = Object.freeze({
				cfvo,
				color,
				minLength: parseInteger(db.attrs.get("minLength"), 0, 4294967295) ?? 10,
				maxLength: parseInteger(db.attrs.get("maxLength"), 0, 4294967295) ?? 90,
				showValue: parseBool(db.attrs.get("showValue"), true),
				...PLAIN_BAR
			});
		} else if (is !== void 0) {
			const cfvo = readCfValues(is, ns, warn);
			if (cfvo === null) continue;
			const raw = is.attrs.get("iconSet");
			const set = raw === void 0 ? "3TrafficLights1" : ICON_SETS.find((s) => s === raw);
			if (set === void 0) {
				warn(`A conditional formatting rule has the unknown icon set "${String(raw).slice(0, 40)}"; it is ignored.`);
				continue;
			}
			iconSet = Object.freeze({
				iconSet: set,
				cfvo,
				reverse: parseBool(is.attrs.get("reverse"), false),
				showValue: parseBool(is.attrs.get("showValue"), true),
				percent: parseBool(is.attrs.get("percent"), true),
				custom: null
			});
		}
		const formulas = childrenNamed(r, ns, "formula").slice(0, 3).map((f) => decodeXstring(f.text));
		rules.push(ruleFromAttributes(r, type, priority, style, formulas, {
			colorScale,
			dataBar,
			iconSet
		}, warn));
		ids.push(type === "dataBar" ? extensionId(r, ns) : null);
	}
	return rules.length === 0 ? null : {
		format: Object.freeze({
			sqref,
			rules: Object.freeze(rules)
		}),
		ids
	};
}
/**
* Reads `x14:conditionalFormattings`. A rule with a `priority` stands alone; a rule without one
* extends the main data bar rule whose `x14:id` equals its `id` (§2.6.27). `mainNs` is the
* namespace of the main elements of a differential format (`x14:dxf` is CT_Dxf).
*/
function parseX14ConditionalFormattings(el, mainNs, readDxf, readColor, warn) {
	const formats = [];
	const bars = /* @__PURE__ */ new Map();
	for (const cf of childrenNamed(el, X14_NS, "conditionalFormatting")) {
		const sqrefEl = childrenNamed(cf, XM_NS, "sqref")[0];
		const rules = [];
		for (const r of childrenNamed(cf, X14_NS, "cfRule")) {
			const rawPriority = r.attrs.get("priority");
			const db = childrenNamed(r, X14_NS, "dataBar")[0];
			if (rawPriority === void 0) {
				const id = r.attrs.get("id")?.trim().toUpperCase();
				if (id === void 0 || id === "" || db === void 0) continue;
				const bar = readX14DataBar(db, warn, readColor);
				if (bar !== null) {
					const { fill: _fill, ...rest } = bar;
					bars.set(id, rest);
				}
				continue;
			}
			const priority = parseInteger(rawPriority, 1, MAX_PRIORITY);
			const rawType = r.attrs.get("type") ?? "";
			const type = CF_TYPES.find((t) => t === rawType);
			if (priority === null || type === void 0) {
				warn(`An extended conditional formatting rule has ${priority === null ? "no valid priority" : `the unknown type "${rawType.slice(0, 40)}"`}; it is ignored.`);
				continue;
			}
			const dxfEl = childrenNamed(r, X14_NS, "dxf")[0];
			const style = dxfEl === void 0 ? null : readDxf(dxfEl);
			let colorScale = null;
			let dataBar = null;
			let iconSet = null;
			if (type === "colorScale") {
				const cs = childrenNamed(r, X14_NS, "colorScale")[0];
				const cfvo = cs === void 0 ? null : readCfValues(cs, X14_NS, warn, true);
				const colors = cs === void 0 ? null : readColors(cs, X14_NS, warn, readColor);
				if (cfvo === null || colors === null) continue;
				colorScale = Object.freeze({
					cfvo,
					colors
				});
			} else if (type === "dataBar") {
				const bar = db === void 0 ? null : readX14DataBar(db, warn, readColor);
				if (bar === null || bar.fill === null) {
					warn("An extended data bar has no fill colour; the rule is ignored.");
					continue;
				}
				const { fill, ...rest } = bar;
				dataBar = Object.freeze({
					...rest,
					color: fill,
					extended: true
				});
			} else if (type === "iconSet") {
				const is = childrenNamed(r, X14_NS, "iconSet")[0];
				iconSet = is === void 0 ? null : readX14IconSet(is, warn);
				if (iconSet === null) continue;
			}
			const formulas = childrenNamed(r, XM_NS, "f").map((f) => decodeXstring(f.text));
			rules.push(ruleFromAttributes(r, type, priority, style, CF_GRAPHIC_TYPES.includes(type) ? [] : formulas, {
				colorScale,
				dataBar,
				iconSet
			}, warn));
		}
		if (rules.length === 0) continue;
		const sqref = readSqref(sqrefEl?.text, MAX_SQREF_REFS, "conditional format", warn);
		if (sqref !== null) formats.push(Object.freeze({
			sqref,
			rules: Object.freeze(rules)
		}));
	}
	return {
		formats,
		bars
	};
}
/**
* Collects the `conditionalFormatting` and `dataValidations` elements of a worksheet, and the
* [MS-XLSX] extensions of the same in the worksheet `extLst` (SR-40), while it is streamed
* through the SAX parser (they follow `sheetData`, so this costs no memory per row). Each element
* is kept as a small tree until its end tag; elements and attributes count against
* `maxElements` for the whole sheet, so a hostile part cannot grow it without bound. Other
* extensions in `extLst` are skipped without being kept.
*/
function createRuleCollector(ctx) {
	const main = [];
	const extended = [];
	const bars = /* @__PURE__ */ new Map();
	const dataValidations = [];
	let stack = null;
	let count = 0;
	let inExtList = false;
	let skip = 0;
	const addValidations = (list) => {
		const room = MAX_DATA_VALIDATIONS - dataValidations.length;
		if (list.length > room) ctx.warn(`The sheet has more than ${String(MAX_DATA_VALIDATIONS)} data validations; the rest are ignored.`);
		dataValidations.push(...list.slice(0, Math.max(0, room)));
	};
	return {
		start(elementNs, local, depth, attrs) {
			if (skip > 0) {
				skip++;
				return true;
			}
			if (stack === null) {
				const isMain = elementNs === ctx.ns();
				if (isMain && depth === 2 && local === "extLst") {
					inExtList = true;
					return true;
				}
				if (inExtList && depth === 3) {
					const uri = attrs.get("uri")?.toUpperCase();
					if (!isMain || local !== "ext" || uri !== "{78C0D931-6437-407D-A8EE-F0AAD7539E65}" && uri !== "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}") {
						skip = 1;
						return true;
					}
				} else if (!isMain || depth !== 2 || local !== "conditionalFormatting" && local !== "dataValidations") return false;
				stack = [];
			}
			count += 1 + attrs.size;
			if (count > ctx.maxElements) throw ctx.tooLarge();
			const el = {
				ns: elementNs,
				local,
				attrs,
				children: [],
				text: ""
			};
			stack[stack.length - 1]?.children.push(el);
			stack.push(el);
			return true;
		},
		end(elementNs, local) {
			if (skip > 0) {
				skip--;
				return true;
			}
			if (stack === null) {
				if (inExtList && local === "extLst" && elementNs === ctx.ns()) {
					inExtList = false;
					return true;
				}
				return false;
			}
			const done = stack.pop();
			if (stack.length === 0 && done !== void 0) {
				stack = null;
				if (done.local === "dataValidations") addValidations(parseDataValidations(done, done.ns, ctx.warn));
				else if (done.local === "conditionalFormatting") {
					const cf = readConditionalFormatting(done, done.ns, ctx.dxf, ctx.readColor, ctx.warn);
					if (cf !== null) main.push(cf);
				} else for (const child of done.children) {
					if (child.ns !== "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main") continue;
					if (child.local === "dataValidations") addValidations(parseX14DataValidations(child, ctx.warn));
					else if (child.local === "conditionalFormattings") {
						const x = parseX14ConditionalFormattings(child, ctx.ns(), ctx.readDxf, ctx.readColor, ctx.warn);
						extended.push(...x.formats);
						for (const [id, bar] of x.bars) if (!bars.has(id)) bars.set(id, bar);
					}
				}
			}
			return true;
		},
		text(value) {
			if (skip > 0) return true;
			if (stack === null) return false;
			const top = stack[stack.length - 1];
			if (top !== void 0) top.text += value;
			return true;
		},
		result() {
			return {
				conditionalFormats: [...main.map(({ format, ids }) => {
					if (!ids.some((id) => id !== null && bars.has(id))) return format;
					const rules = format.rules.map((rule, i) => {
						const ext = ids[i] === null || ids[i] === void 0 ? void 0 : bars.get(ids[i]);
						return ext === void 0 || rule.dataBar === null ? rule : Object.freeze({
							...rule,
							dataBar: Object.freeze({
								...rule.dataBar,
								...ext,
								extended: true
							})
						});
					});
					return Object.freeze({
						sqref: format.sqref,
						rules: Object.freeze(rules)
					});
				}), ...extended],
				dataValidations
			};
		}
	};
}
// src/sml/worksheet.ts
/**
* Format codes longer than this are not classified (and so never treated as dates); this
* bounds the work an oversized code can cause. Implementation choice, see docs/decisions.md D-115.
*/
const MAX_FORMAT_CODE_LENGTH = 1024;
/** The view of a sheet without a `sheetView`, from the defaults of CT_SheetView (W2-51). */
const DEFAULT_SHEET_VIEW = Object.freeze({
	zoom: 100,
	showGridLines: true,
	showRowColHeaders: true,
	rightToLeft: false,
	layout: "normal",
	split: null,
	topLeftCell: null,
	activeCell: null,
	zoomScaleNormal: 0,
	showRuler: true
});
/** Sheet properties without `sheetFormatPr`/`outlinePr`: the CT_SheetFormatPr and CT_OutlinePr defaults, row height 15 (W2-09, W2-65). */
const DEFAULT_SHEET_PROPERTIES = Object.freeze({
	defaultRowHeight: 15,
	outlineLevelRow: 0,
	outlineLevelCol: 0,
	summaryBelow: true,
	summaryRight: true
});
/** Hyperlinks kept per sheet; more are skipped with a warning (implementation limit, D-A08). */
const MAX_HYPERLINKS = 65536;
const FORMULA_KINDS = [
	"normal",
	"shared",
	"array",
	"dataTable"
];
/** Parses a non-negative integer of at most `maxDigits` digits; `null` otherwise. */
function parseUint(text, maxDigits) {
	const s = text.trim();
	if (s.length === 0 || s.length > maxDigits) return null;
	let n = 0;
	for (let i = 0; i < s.length; i++) {
		const d = s.charCodeAt(i) - 48;
		if (d < 0 || d > 9) return null;
		n = n * 10 + d;
	}
	return n;
}
/**
* Parses the lexical form of `xsd:double` restricted to finite values: optional sign, digits
* with an optional fraction, optional exponent. Returns `null` for anything else (including
* INF and NaN, which have no JSON representation).
*/
function parseXsdDouble(text) {
	const s = text.trim();
	if (s.length === 0 || s.length > 64) return null;
	let i = 0;
	if (s[i] === "+" || s[i] === "-") i++;
	let digits = 0;
	while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) {
		i++;
		digits++;
	}
	if (s[i] === ".") {
		i++;
		while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) {
			i++;
			digits++;
		}
	}
	if (digits === 0) return null;
	if (s[i] === "e" || s[i] === "E") {
		i++;
		if (s[i] === "+" || s[i] === "-") i++;
		const start = i;
		while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
		if (i === start) return null;
	}
	if (i !== s.length) return null;
	const n = Number(s);
	return Number.isFinite(n) ? n : null;
}
/** Parses a worksheet part. */
function parseWorksheet(xml, ctx) {
	const { handler, finish } = createWorksheetParser(ctx);
	parseXml(xml, handler, {
		part: ctx.part,
		limits: ctx.limits
	});
	return finish();
}
/**
* The worksheet parser as an XML handler, for whole-text and streaming parsing. With `onRow`
* (streaming, docs/decisions.md D-402), each row's cells are passed on when its `</row>` is
* read and then dropped, so memory does not grow with the sheet; rows come in file order, and
* shared formulas whose master comes later in the file cannot be expanded.
*/
/** The cells of one row's style-only list (`[col, styleIndex, …]`), in column order. */
function styledCellsOf(row, list, make) {
	const pairs = [];
	for (let i = 0; i + 1 < list.length; i += 2) if (list[i + 1] !== -1) pairs.push([list[i], list[i + 1]]);
	pairs.sort((a, b) => a[0] - b[0]);
	return pairs.map(([col, styleIndex]) => make(row, col, styleIndex));
}
function createWorksheetParser(ctx, onRow) {
	const rows = /* @__PURE__ */ new Map();
	const styleOnly = /* @__PURE__ */ new Map();
	const styleOnlyMax = /* @__PURE__ */ new Map();
	const styleOnlyIndex = /* @__PURE__ */ new Map();
	const merges = [];
	const formats = ctx.formatCache;
	const path = [];
	let ns = null;
	let dimension = null;
	let cellCount = 0;
	const masters = /* @__PURE__ */ new Map();
	const dependents = [];
	const compiled = /* @__PURE__ */ new Map();
	const columns = [];
	const rowInfo = /* @__PURE__ */ new Map();
	let freeze = null;
	let autoFilter = null;
	let tabColor = null;
	let view = DEFAULT_SHEET_VIEW;
	let properties = DEFAULT_SHEET_PROPERTIES;
	const selections = /* @__PURE__ */ new Map();
	let activePane = "topLeft";
	const hyperlinks = [];
	let drawingRelId = null;
	let pictureRelId = null;
	const print = createPrintCollector((message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc()));
	let buffered = 0;
	const bufferCell = (ref) => {
		if (!onRow || ++buffered <= 16384) return;
		throw new XlvError("XLV_E_LIMIT_EXCEEDED", `One row element of the sheet holds more than ${String(MAX_COLUMNS)} cells, more than a row can have.`, {
			location: loc(ref),
			hint: `A row has at most ${String(MAX_COLUMNS)} cells; the cells of this row element name other rows, which the streaming reader would have to hold in memory. The file is damaged or was made to exhaust memory. Re-save it in a spreadsheet application, or read it with readWorkbook, which bounds all cells with limits.maxCells.`,
			details: {
				limit: "maxCells",
				value: MAX_COLUMNS
			}
		});
	};
	const rules = createRuleCollector({
		ns: () => ns ?? SML_NS.transitional,
		maxElements: ctx.limits.maxXmlTreeElements,
		dxf: (index) => ctx.styles.dxf(index),
		readDxf: (el) => parseDxfElement(el, ns ?? SML_NS.transitional, (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc())),
		readColor: (el) => parseColor(el, "conditional format", (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc())),
		warn: (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc()),
		tooLarge: () => new XlvError("XLV_E_LIMIT_EXCEEDED", "The conditional formatting and data validations of the sheet have more elements and attributes than the limit.", {
			location: {
				part: ctx.part,
				sheet: ctx.sheet
			},
			hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTreeElements: ${String(ctx.limits.maxXmlTreeElements * 4)} } }.`,
			details: {
				limit: "maxXmlTreeElements",
				value: ctx.limits.maxXmlTreeElements
			}
		})
	});
	const relIdKey = `{${REL_NS[ctx.conformance]}}id`;
	let sheetViews = 0;
	let currentRow = 0;
	let lastCol = 0;
	/** The current row has an invalid number; its cells are skipped (D-903). */
	let skipRow = false;
	let cell = null;
	/** Where character data currently goes. */
	let target = null;
	let phoneticDepth = 0;
	/** Raw text of the inline-string `t` element being read. */
	let inlineText = "";
	/** The inline rich text run being read: its text, font and run properties (rPr children). */
	let inlineRun = null;
	const loc = (ref) => ref === void 0 ? {
		part: ctx.part,
		sheet: ctx.sheet
	} : {
		part: ctx.part,
		sheet: ctx.sheet,
		cell: ref
	};
	const numberFormat = (styleIndex) => {
		const id = ctx.styles.numFmtId(styleIndex);
		let format = formats.get(id);
		if (format === void 0) {
			const code = ctx.styles.customFormatCode(id) ?? BUILTIN_NUMBER_FORMATS.get(id) ?? null;
			let isDate;
			if (code === null) isDate = isDateFormatId(id);
			else if (code.length > MAX_FORMAT_CODE_LENGTH) {
				isDate = false;
				ctx.warnings.add("XLV_W_INVALID_VALUE", `Number format ${String(id)} is ${String(code.length)} characters long; it is not checked for dates.`, {
					part: ctx.part,
					sheet: ctx.sheet
				});
			} else isDate = isDateFormat(code);
			format = {
				code: id === 0 || code === "General" ? null : code,
				isDate
			};
			formats.set(id, format);
		}
		return format;
	};
	/** Position of `col` in the style-only list of `row`, or -1; searched only below the largest column seen. */
	const styledIndex = (row, col) => {
		const max = styleOnlyMax.get(row);
		if (max === void 0 || col > max) return -1;
		const list = styleOnly.get(row);
		if (list === void 0) return -1;
		if (list[list.length - 2] === col && list[list.length - 1] !== -1) return list.length - 2;
		let index = styleOnlyIndex.get(row);
		if (index === void 0) {
			index = /* @__PURE__ */ new Map();
			for (let i = 0; i + 1 < list.length; i += 2) if (list[i + 1] !== -1) index.set(list[i], i);
			styleOnlyIndex.set(row, index);
		}
		return index.get(col) ?? -1;
	};
	const styledCell = (row, col, styleIndex) => {
		const code = numberFormat(styleIndex).code;
		return {
			ref: formatCellRef(row, col),
			row,
			col,
			type: "empty",
			value: null,
			formula: null,
			styleIndex,
			numFmt: code,
			style: ctx.styles.cellStyle(styleIndex, code)
		};
	};
	const finishCell = (p) => {
		const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc(p.ref));
		const format = numberFormat(p.styleIndex);
		let type = "empty";
		let value = null;
		const raw = p.v;
		let richText;
		switch (p.t) {
			case "n": {
				if (raw === null) break;
				const n = parseXsdDouble(raw);
				if (n === null) {
					warn(`A numeric cell holds "${raw.slice(0, 40)}", which is not a number; it is returned as text.`);
					type = "string";
					value = raw;
				} else if (format.isDate) {
					const date = serialToDate(n, ctx.dateSystem, ctx.conformance);
					if (date === null) {
						warn(`The date serial ${String(n)} is outside the range of the ${ctx.dateSystem} date system; it is returned as a number.`);
						type = "number";
						value = n;
					} else {
						type = "date";
						value = date;
					}
				} else {
					type = "number";
					value = n;
				}
				break;
			}
			case "s": {
				if (raw === null) break;
				const index = parseUint(raw, 10);
				const text = index === null ? void 0 : ctx.sharedStrings[index];
				if (text === void 0) warn(`Shared string index "${raw.slice(0, 40)}" does not exist; the cell is treated as empty.`);
				else {
					ctx.budget.sharedText -= text.length;
					if (ctx.budget.sharedText < 0) throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The cells refer to more shared-string text than the limit allows.", {
						location: loc(p.ref),
						hint: `Small files can repeat one huge shared string in many cells. If you trust this file, raise the limit, for example { limits: { maxSharedStringChars: ${String(ctx.limits.maxSharedStringChars * 4)} } }.`,
						details: {
							limit: "maxSharedStringChars",
							value: ctx.limits.maxSharedStringChars
						}
					});
					type = "string";
					value = text;
					richText = index === null ? void 0 : ctx.richStrings.get(index);
				}
				break;
			}
			case "str":
				if (raw !== null) {
					type = "string";
					value = decodeXstring(raw);
				}
				break;
			case "inlineStr": break;
			case "b": {
				if (raw === null) break;
				const b = parseXsdBoolean(raw);
				if (b === null) {
					warn(`A boolean cell holds "${raw.slice(0, 40)}"; it is returned as text.`);
					type = "string";
					value = raw;
				} else {
					type = "boolean";
					value = b;
				}
				break;
			}
			case "e":
				if (raw !== null) {
					type = "error";
					value = raw.trim();
				}
				break;
			case "d": {
				if (raw === null) break;
				const date = parseIsoDateTime(raw.trim());
				if (date === null) {
					warn(`A date cell holds "${raw.slice(0, 40)}", which is not an ISO 8601 date; it is returned as text.`);
					type = "string";
					value = raw;
				} else {
					type = "date";
					value = date;
				}
				break;
			}
			default:
				warn(`The cell type "${p.t.slice(0, 20)}" is unknown; the value is returned as text.`);
				if (raw !== null) {
					type = "string";
					value = raw;
				}
		}
		if (value === null && p.inline !== null) {
			type = "string";
			value = p.inline;
			if (p.inlineRuns !== null) richText = Object.freeze(p.inlineRuns.map((r) => Object.freeze(r)));
		}
		let formula = null;
		if (p.formulaAttrs !== null) {
			const rawKind = p.formulaAttrs.get("t") ?? "normal";
			let kind = FORMULA_KINDS.find((k) => k === rawKind) ?? "normal";
			const si = p.formulaAttrs.get("si");
			const sharedIndex = si === void 0 ? null : parseUint(si, 10);
			const rawRef = p.formulaAttrs.get("ref");
			const ref = rawRef === void 0 ? null : rawRef.replace(/\$/g, "");
			if (kind === "shared" && sharedIndex === null) kind = "normal";
			let text = decodeXstring(p.formulaText);
			if (kind === "shared" && sharedIndex !== null) {
				if (ref === null) {
					text = null;
					dependents.push({
						row: p.row,
						col: p.col,
						ref: p.ref,
						si: sharedIndex
					});
				} else if (!masters.has(sharedIndex)) masters.set(sharedIndex, {
					row: p.row,
					col: p.col,
					text,
					range: parseRange(ref)
				});
			}
			formula = {
				kind,
				text,
				ref,
				sharedIndex: kind === "shared" ? sharedIndex : null
			};
			if (p.v === null && p.inline === null) ctx.warnings.add("XLV_W_FORMULA_NO_CACHED_VALUE", "Formula cells without a cached value are returned with value null; xlvane does not calculate formulas.", loc(p.ref));
		}
		if (value === null && formula === null) {
			if (p.styleIndex === 0) return;
			const valued = rows.get(p.row);
			if (valued?.has(p.col) === true) {
				valued.delete(p.col);
				cellCount--;
				ctx.budget.remaining++;
				buffered--;
				ctx.warnings.add("XLV_W_DUPLICATE_CELL", "The sheet defines the same cell more than once; the last definition is used.", loc(p.ref));
			}
			let styled = styleOnly.get(p.row);
			if (styled === void 0) {
				styled = [];
				styleOnly.set(p.row, styled);
			}
			const at = styledIndex(p.row, p.col);
			if (at >= 0) {
				styled[at + 1] = p.styleIndex;
				ctx.warnings.add("XLV_W_DUPLICATE_CELL", "The sheet defines the same cell more than once; the last definition is used.", loc(p.ref));
				return;
			}
			if (--ctx.budget.remaining < 0) throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The workbook has more cells than the limit.", {
				location: loc(p.ref),
				hint: `Cells that hold only a style count too. If you trust this file, raise the limit, for example { limits: { maxCells: ${String(ctx.limits.maxCells * 4)} } }.`,
				details: {
					limit: "maxCells",
					value: ctx.limits.maxCells
				}
			});
			bufferCell(p.ref);
			styleOnlyIndex.get(p.row)?.set(p.col, styled.length);
			styled.push(p.col, p.styleIndex);
			if (p.col > (styleOnlyMax.get(p.row) ?? 0)) styleOnlyMax.set(p.row, p.col);
			return;
		}
		const at = styledIndex(p.row, p.col);
		if (at >= 0) {
			const list = styleOnly.get(p.row);
			if (list !== void 0) list[at + 1] = -1;
			styleOnlyIndex.get(p.row)?.delete(p.col);
			ctx.budget.remaining++;
			buffered--;
			ctx.warnings.add("XLV_W_DUPLICATE_CELL", "The sheet defines the same cell more than once; the last definition is used.", loc(p.ref));
		}
		let rowCells = rows.get(p.row);
		if (rowCells === void 0) {
			rowCells = /* @__PURE__ */ new Map();
			rows.set(p.row, rowCells);
		}
		if (rowCells.has(p.col)) ctx.warnings.add("XLV_W_DUPLICATE_CELL", "The sheet defines the same cell more than once; the last definition is used.", loc(p.ref));
		else {
			if (--ctx.budget.remaining < 0) throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The workbook has more cells than the limit.", {
				location: loc(p.ref),
				hint: `If you trust this file, raise the limit, for example { limits: { maxCells: ${String(ctx.limits.maxCells * 4)} } }.`,
				details: {
					limit: "maxCells",
					value: ctx.limits.maxCells
				}
			});
			cellCount++;
			bufferCell(p.ref);
		}
		rowCells.set(p.col, {
			ref: p.ref,
			row: p.row,
			col: p.col,
			type,
			value,
			formula,
			styleIndex: p.styleIndex,
			numFmt: format.code,
			style: ctx.styles.cellStyle(p.styleIndex, format.code),
			...richText === void 0 ? {} : { richText }
		});
	};
	const spendFormulaText = (chars, ref) => {
		ctx.budget.formulaText -= chars;
		if (ctx.budget.formulaText >= 0) return;
		throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The cells that share formulas would hold more formula text than the limit allows.", {
			location: loc(ref),
			hint: `Small files can share one huge formula with many cells; the limit counts the formula of each cell that shares it. If you trust this file, raise the limit, for example { limits: { maxSharedStringChars: ${String(ctx.limits.maxSharedStringChars * 4)} } }.`,
			details: {
				limit: "maxSharedStringChars",
				value: ctx.limits.maxSharedStringChars
			}
		});
	};
	const warnNotExpanded = (ref) => ctx.warnings.add("XLV_W_SHARED_FORMULA_NOT_EXPANDED", "Some cells reuse a shared formula that could not be moved to them (missing master, cell outside the master range, or formula text xlvane cannot translate safely); their formula.text is null and their cached values are available.", loc(ref));
	/** Streaming: expands the pending shared formulas and hands the buffered cells on. */
	const flushRows = () => {
		if (!onRow) return;
		buffered = 0;
		const info = rowInfo.get(currentRow);
		rowInfo.delete(currentRow);
		if (rows.size === 0 && styleOnly.size === 0) return;
		expandSharedFormulas(rows, masters, compiled, dependents, warnNotExpanded, spendFormulaText);
		dependents.length = 0;
		const out = [];
		for (const r of [...rows.keys()].sort((a, b) => a - b)) {
			const cells = rows.get(r);
			if (cells) for (const c of [...cells.keys()].sort((a, b) => a - b)) out.push(cells.get(c));
		}
		rows.clear();
		const styled = [];
		for (const r of [...styleOnly.keys()].sort((a, b) => a - b)) styled.push(...styledCellsOf(r, styleOnly.get(r) ?? [], styledCell));
		styleOnly.clear();
		styleOnlyMax.clear();
		styleOnlyIndex.clear();
		if (out.length > 0 || styled.length > 0) onRow(currentRow, out, info, styled);
	};
	const handler = {
		start(elementNs, local, attrs) {
			if (ns === null) {
				if (local !== "worksheet" || elementNs !== SML_NS.strict && elementNs !== SML_NS.transitional) throw new XlvError("XLV_E_PART_INVALID", `The worksheet part has the root element <${local}>.`, {
					location: {
						part: ctx.part,
						sheet: ctx.sheet
					},
					hint: "The sheet part is not a SpreadsheetML worksheet. Re-save the file in a spreadsheet application.",
					details: {
						element: local,
						namespace: elementNs
					}
				});
				ns = elementNs;
			}
			const parent = path[path.length - 1];
			path.push(elementNs === ns ? local : "");
			const depth = path.length;
			if (cell === null && rules.start(elementNs, local, depth, attrs)) return;
			if (elementNs !== ns) return;
			if (cell !== null) {
				if (phoneticDepth > 0 || local === "rPh") {
					phoneticDepth++;
					return;
				}
				if (depth === 5 && local === "v") target = "v";
				else if (depth === 5 && local === "f") {
					target = "f";
					cell.formulaAttrs = attrs;
				} else if (depth === 5 && local === "is") cell.inline ??= "";
				else if (depth === 6 && local === "r" && parent === "is") {
					cell.inlineRuns ??= cell.inline === null || cell.inline === "" ? [] : [{ text: cell.inline }];
					inlineRun = {
						text: "",
						font: void 0,
						properties: null
					};
				} else if (depth === 7 && local === "rPr" && inlineRun !== null) inlineRun.properties = [];
				else if (depth === 8 && parent === "rPr" && inlineRun?.properties) inlineRun.properties.push({
					ns: elementNs,
					local,
					attrs,
					children: [],
					text: ""
				});
				else if (local === "t" && cell.inline !== null && (parent === "is" || parent === "r")) target = "is";
				return;
			}
			if (print.start(depth, local, parent, attrs)) return;
			if (depth === 2 && local === "dimension") {
				const ref = attrs.get("ref") ?? "";
				const range = parseRange(ref);
				if (range === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The sheet dimension "${ref.slice(0, 40)}" is not a valid range.`, loc());
				else dimension = formatRange(range);
			} else if (depth === 3 && local === "mergeCell" && parent === "mergeCells") {
				const ref = attrs.get("ref") ?? "";
				const range = parseRange(ref);
				if (range === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The merged range "${ref.slice(0, 40)}" is not a valid range.`, loc());
				else merges.push(formatRange(range));
			} else if (depth === 3 && local === "hyperlink" && parent === "hyperlinks") {
				const ref = attrs.get("ref") ?? "";
				const range = parseRange(ref);
				const text = (name) => {
					const v = attrs.get(name);
					return v === void 0 ? null : decodeXstring(v);
				};
				if (range === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The hyperlink range "${ref.slice(0, 40)}" is not a valid range; the link is skipped.`, loc());
				else if (hyperlinks.length >= MAX_HYPERLINKS) ctx.warnings.add("XLV_W_INVALID_VALUE", `The sheet has more than ${String(MAX_HYPERLINKS)} hyperlinks; the rest are skipped.`, loc());
				else hyperlinks.push({
					ref: formatRange(range),
					relId: attrs.get(relIdKey) ?? null,
					location: text("location"),
					tooltip: text("tooltip"),
					display: text("display")
				});
			} else if (depth === 2 && (local === "drawing" || local === "picture")) {
				if (local === "drawing") drawingRelId ??= attrs.get(relIdKey) ?? null;
				else pictureRelId ??= attrs.get(relIdKey) ?? null;
			} else if (depth === 2 && local === "autoFilter") {
				const ref = attrs.get("ref");
				const range = ref === void 0 ? null : parseRange(ref);
				if (range === null) {
					if (ref !== void 0) ctx.warnings.add("XLV_W_INVALID_VALUE", `The autofilter range "${ref.slice(0, 40)}" is not a valid range.`, loc());
				} else autoFilter = formatRange(range);
			} else if (depth === 3 && local === "sheetView" && parent === "sheetViews") {
				sheetViews++;
				if (sheetViews === 1) view = parseSheetView(attrs, (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc()));
			} else if (depth === 4 && local === "pane" && parent === "sheetView" && sheetViews === 1) {
				const warn = (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc());
				activePane = attrs.get("activePane") ?? "topLeft";
				freeze = parsePane(attrs, warn);
				const split = parseSplit(attrs, warn);
				if (split !== null) view = Object.freeze({
					...view,
					split
				});
			} else if (depth === 4 && local === "selection" && parent === "sheetView" && sheetViews === 1) {
				const cell = parseCellRef((attrs.get("activeCell") ?? "").trim());
				if (cell !== null) selections.set(attrs.get("pane") ?? "topLeft", formatCellRef(cell.row, cell.col));
			} else if (depth === 2 && local === "sheetFormatPr") properties = parseSheetFormat(attrs, properties, (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc()));
			else if (depth === 3 && local === "outlinePr" && parent === "sheetPr") {
				const flag = (name) => {
					const raw = attrs.get(name);
					return raw === void 0 ? true : parseXsdBoolean(raw) !== false;
				};
				properties = Object.freeze({
					...properties,
					summaryBelow: flag("summaryBelow"),
					summaryRight: flag("summaryRight")
				});
			} else if (depth === 3 && local === "tabColor" && parent === "sheetPr") tabColor = parseColor({
				ns: elementNs,
				local,
				attrs,
				children: [],
				text: ""
			}, "sheet tab", (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc())) ?? null;
			else if (depth === 3 && local === "col" && parent === "cols") {
				const col = parseColumn(attrs);
				if (col === null) ctx.warnings.add("XLV_W_INVALID_VALUE", "A column definition (col) has invalid min, max or width attributes; it is ignored.", loc());
				else {
					const s = attrs.get("style");
					const index = s === void 0 ? null : parseUint(s, 10);
					if (s !== void 0 && index === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The column style index "${s.slice(0, 20)}" is not a number; it is ignored.`, loc());
					columns.push(index === null ? col : {
						...col,
						style: ctx.styles.cellStyle(index, numberFormat(index).code)
					});
				}
			} else if (depth === 3 && local === "row" && parent === "sheetData") {
				const r = attrs.get("r");
				lastCol = 0;
				const n = r === void 0 ? currentRow + 1 : parseUint(r, 7);
				if (n === null || n < 1 || n > 1048576) {
					skipRow = true;
					ctx.warnings.add("XLV_W_INVALID_CELL_REF", "A row has an invalid row number; the row and its cells are skipped.", loc());
					return;
				}
				skipRow = false;
				currentRow = n;
				const ht = attrs.get("ht");
				const height = ht === void 0 ? null : parseXsdDouble(ht);
				const hiddenRaw = attrs.get("hidden");
				const hidden = hiddenRaw === void 0 ? false : parseXsdBoolean(hiddenRaw) === true;
				if (ht !== void 0 && (height === null || height < 0)) ctx.warnings.add("XLV_W_INVALID_VALUE", `The row height "${ht.slice(0, 20)}" is not a valid number; it is ignored.`, loc());
				const validHeight = height !== null && height >= 0 ? height : null;
				const outline = parseOutline(attrs);
				const customFormat = attrs.get("customFormat");
				let style;
				if (customFormat !== void 0 && parseXsdBoolean(customFormat) === true) {
					const s = attrs.get("s") ?? "0";
					const index = parseUint(s, 10);
					if (index === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The row style index "${s.slice(0, 20)}" is not a number; it is ignored.`, loc());
					else style = ctx.styles.cellStyle(index, numberFormat(index).code);
				}
				if (validHeight !== null || hidden || outline.outlineLevel !== void 0 || outline.collapsed !== void 0 || style !== void 0) rowInfo.set(currentRow, style === void 0 ? {
					height: validHeight,
					hidden,
					...outline
				} : {
					height: validHeight,
					hidden,
					...outline,
					style
				});
			} else if (depth === 4 && local === "c" && parent === "row" && path[1] === "sheetData") {
				if (skipRow) return;
				const r = attrs.get("r");
				let row = currentRow;
				let col;
				if (r === void 0 || r === "") col = lastCol + 1;
				else {
					const address = parseCellRef(r);
					if (address === null) {
						ctx.warnings.add("XLV_W_INVALID_CELL_REF", `The cell reference "${r.slice(0, 20)}" is not valid inside the sheet grid; the cell is skipped.`, loc());
						return;
					}
					row = address.row;
					col = address.col;
				}
				if (col > 16384 || row < 1) {
					ctx.warnings.add("XLV_W_INVALID_CELL_REF", "A cell without a reference lies past the last column; the cell is skipped.", loc());
					return;
				}
				lastCol = col;
				const s = attrs.get("s");
				let styleIndex = 0;
				if (s !== void 0) {
					const parsed = parseUint(s, 10);
					if (parsed === null) ctx.warnings.add("XLV_W_INVALID_VALUE", `The style index "${s.slice(0, 20)}" is not a number; using 0.`, loc());
					else styleIndex = parsed;
				}
				cell = {
					row,
					col,
					ref: formatCellRef(row, col),
					t: attrs.get("t") ?? "n",
					styleIndex,
					v: null,
					inline: null,
					inlineRuns: null,
					formulaAttrs: null,
					formulaText: ""
				};
			}
		},
		end(elementNs, local) {
			path.pop();
			if (cell === null && rules.end(elementNs, local)) return;
			if (elementNs === ns) print.end(local);
			if (onRow && local === "row" && elementNs === ns && path.length === 2) flushRows();
			if (elementNs !== ns || cell === null) return;
			if (phoneticDepth > 0) {
				phoneticDepth--;
				return;
			}
			if (local === "c" && path.length === 3) {
				const done = cell;
				cell = null;
				target = null;
				finishCell(done);
			} else if (local === "t" && target === "is") {
				const text = decodeXstring(inlineText);
				cell.inline = (cell.inline ?? "") + text;
				if (inlineRun !== null) inlineRun.text += text;
				else cell.inlineRuns?.push({ text });
				inlineText = "";
				target = null;
			} else if (local === "rPr" && path.length === 6 && inlineRun?.properties) {
				const done = cell;
				const font = parseRunFont({
					ns: elementNs,
					local,
					attrs: /* @__PURE__ */ new Map(),
					children: inlineRun.properties,
					text: ""
				}, elementNs, (message) => ctx.warnings.add("XLV_W_INVALID_VALUE", message, loc(done.ref)));
				if (Object.keys(font).length > 0) inlineRun.font = font;
				inlineRun.properties = null;
			} else if (local === "r" && path.length === 5 && inlineRun !== null) {
				cell.inlineRuns?.push(inlineRun.font === void 0 ? Object.freeze({ text: inlineRun.text }) : Object.freeze({
					text: inlineRun.text,
					font: inlineRun.font
				}));
				inlineRun = null;
			} else if ((local === "v" || local === "f") && path.length === 4) target = null;
		},
		text(value) {
			if (rules.text(value)) return;
			if (print.text(value)) return;
			if (cell === null || target === null) return;
			if (target === "v") cell.v = (cell.v ?? "") + value;
			else if (target === "f") cell.formulaText += value;
			else inlineText += value;
		}
	};
	const finish = () => {
		if (onRow) flushRows();
		else expandSharedFormulas(rows, masters, compiled, dependents, warnNotExpanded, spendFormulaText);
		const activeCell = selections.get(activePane);
		if (activeCell !== void 0) view = Object.freeze({
			...view,
			activeCell
		});
		return {
			rows,
			styleOnly,
			styledCell,
			dimension,
			merges,
			cellCount,
			columns,
			rowInfo,
			freeze,
			autoFilter,
			hyperlinks,
			tabColor,
			view,
			properties,
			print: print.result(),
			...rules.result(),
			drawingRelId,
			pictureRelId
		};
	};
	return {
		handler,
		finish
	};
}
/**
* Gives every dependent cell of a shared formula the master formula moved to its position
* (docs/decisions.md D-107, docs/facts/formulas.md). Each master is tokenized once; a dependent
* gets a formula whose `text` is rendered on first access, so a long master formula reused by
* many cells costs memory only for the texts actually read. Dependents that cannot be expanded
* keep `text: null` and are reported through `warn`.
*/
function expandSharedFormulas(rows, masters, compiled, dependents, warn, spend) {
	for (const d of dependents) {
		const rowCells = rows.get(d.row);
		const cell = rowCells?.get(d.col);
		if (rowCells === void 0 || cell === void 0 || cell.formula?.kind !== "shared" || cell.formula.ref !== null || cell.formula.sharedIndex !== d.si) continue;
		const master = masters.get(d.si);
		const range = master?.range;
		if (master === void 0 || range == null || d.row < range.start.row || d.row > range.end.row || d.col < range.start.col || d.col > range.end.col) {
			warn(d.ref);
			continue;
		}
		let formula = compiled.get(d.si);
		if (formula === void 0) {
			formula = compileFormula(master.text);
			compiled.set(d.si, formula);
		}
		const dRow = d.row - master.row;
		const dCol = d.col - master.col;
		if (formula === null || !formula.canShift(dRow, dCol)) {
			warn(d.ref);
			continue;
		}
		spend(formula.renderedLength(dRow, dCol), d.ref);
		rowCells.set(d.col, {
			...cell,
			formula: lazySharedFormula(formula, dRow, dCol, d.si)
		});
	}
}
function lazySharedFormula(compiled, dRow, dCol, si) {
	let text;
	return {
		kind: "shared",
		get text() {
			text ??= compiled.render(dRow, dCol);
			return text;
		},
		ref: null,
		sharedIndex: si
	};
}
/** Parses a `col` element (CT_Col, W2-03); `null` if min/max are missing or out of range. */
function parseColumn(attrs) {
	const min = parseUint(attrs.get("min") ?? "", 5);
	const max = parseUint(attrs.get("max") ?? "", 5);
	if (min === null || max === null || min < 1 || max < min || max > 16384) return null;
	const rawWidth = attrs.get("width");
	const width = rawWidth === void 0 ? null : parseXsdDouble(rawWidth);
	if (rawWidth !== void 0 && (width === null || width < 0)) return null;
	const hidden = attrs.get("hidden");
	return {
		min,
		max,
		width,
		hidden: hidden !== void 0 && parseXsdBoolean(hidden) === true,
		...parseOutline(attrs)
	};
}
/** `outlineLevel` (1–7 kept; 0 and invalid values omitted) and `collapsed` of a row or column (W2-65). */
function parseOutline(attrs) {
	const out = {};
	const raw = attrs.get("outlineLevel");
	const level = raw === void 0 ? null : parseUint(raw.trim(), 3);
	if (level !== null && level >= 1 && level <= 7) out.outlineLevel = level;
	const collapsed = attrs.get("collapsed");
	if (collapsed !== void 0 && parseXsdBoolean(collapsed) === true) out.collapsed = true;
	return out;
}
/** Reads `sheetFormatPr` (CT_SheetFormatPr, sml.xsd L2222; W2-65); invalid values give a warning and keep the default. */
function parseSheetFormat(attrs, base, warn) {
	let defaultRowHeight = base.defaultRowHeight;
	const raw = attrs.get("defaultRowHeight");
	if (raw !== void 0) {
		const h = parseXsdDouble(raw);
		if (h === null || h < 0) warn(`The default row height "${raw.slice(0, 20)}" is not a valid number; 15 is used.`);
		else defaultRowHeight = h;
	}
	const level = (name) => {
		const v = attrs.get(name);
		const n = v === void 0 ? 0 : parseUint(v.trim(), 3);
		return n !== null && n <= 7 ? n : 0;
	};
	return Object.freeze({
		...base,
		defaultRowHeight,
		outlineLevelRow: level("outlineLevelRow"),
		outlineLevelCol: level("outlineLevelCol")
	});
}
const LAYOUTS = [
	"normal",
	"pageBreakPreview",
	"pageLayout"
];
/**
* Reads the display settings of a `sheetView` (CT_SheetView, W2-51). Absent attributes take the
* schema defaults; invalid ones give a warning and the default. `zoomScale` is a percentage from
* 10 to 400 (Part 1 §18.3.1.87).
*/
function parseSheetView(attrs, warn) {
	const flag = (name, fallback) => {
		const raw = attrs.get(name);
		if (raw === void 0) return fallback;
		const v = parseXsdBoolean(raw);
		if (v === null) warn(`The sheet view attribute ${name}="${raw.slice(0, 20)}" is not a boolean; the default is used.`);
		return v ?? fallback;
	};
	let zoom = 100;
	const rawZoom = attrs.get("zoomScale");
	if (rawZoom !== void 0) {
		const n = parseUint(rawZoom.trim(), 10);
		if (n === null || n < 10 || n > 400) warn(`The sheet zoom "${rawZoom.slice(0, 20)}" is not a percentage from 10 to 400; 100 is used.`);
		else zoom = n;
	}
	let zoomScaleNormal = 0;
	const rawNormal = attrs.get("zoomScaleNormal");
	if (rawNormal !== void 0) {
		const n = parseUint(rawNormal.trim(), 10);
		if (n === null || n !== 0 && (n < 10 || n > 400)) warn(`The normal-view zoom "${rawNormal.slice(0, 20)}" is not 0 or a percentage from 10 to 400; 0 is used.`);
		else zoomScaleNormal = n;
	}
	const topLeft = parseCellRef((attrs.get("topLeftCell") ?? "").trim());
	const rawLayout = (attrs.get("view") ?? "normal").trim();
	const layout = LAYOUTS.find((l) => l === rawLayout);
	if (layout === void 0) warn(`The sheet view type "${rawLayout.slice(0, 30)}" is not known; "normal" is used.`);
	return Object.freeze({
		zoom,
		showGridLines: flag("showGridLines", true),
		showRowColHeaders: flag("showRowColHeaders", true),
		rightToLeft: flag("rightToLeft", false),
		layout: layout ?? "normal",
		split: null,
		topLeftCell: topLeft === null ? null : formatCellRef(topLeft.row, topLeft.col),
		activeCell: null,
		zoomScaleNormal,
		showRuler: flag("showRuler", true)
	});
}
/**
* Reads a split (not frozen) window from a `pane` whose `state` is `split`, the default (CT_Pane,
* ST_PaneState): xSplit and ySplit are then positions in 1/20 point, and xSplit is the vertical
* split bar ([MS-OI29500] §2.1.640 d; W2-51). Returned in points.
*/
function parseSplit(attrs, warn) {
	if ((attrs.get("state") ?? "split").trim() !== "split") return null;
	const position = (name) => {
		const raw = attrs.get(name);
		if (raw === void 0) return 0;
		const n = parseXsdDouble(raw);
		if (n === null || n < 0) {
			warn(`The split pane attribute ${name}="${raw.slice(0, 20)}" is not a valid position; the split is ignored.`);
			return null;
		}
		return n / 20;
	};
	const x = position("xSplit");
	const y = position("ySplit");
	if (x === null || y === null || x === 0 && y === 0) return null;
	const topLeft = parseCellRef((attrs.get("topLeftCell") ?? "").trim());
	return Object.freeze(topLeft === null ? {
		x,
		y
	} : {
		x,
		y,
		topLeftCell: formatCellRef(topLeft.row, topLeft.col)
	});
}
/**
* Reads frozen panes from a `pane` element (CT_Pane, W2-05): only the states `frozen` and
* `frozenSplit` freeze rows and columns; xSplit and ySplit then count the frozen columns and rows.
*/
function parsePane(attrs, warn) {
	const state = attrs.get("state") ?? "split";
	if (state !== "frozen" && state !== "frozenSplit") return null;
	const count = (name, max) => {
		const raw = attrs.get(name);
		if (raw === void 0) return 0;
		const n = parseXsdDouble(raw);
		if (n === null || !Number.isInteger(n) || n < 0 || n >= max) {
			warn(`The frozen pane attribute ${name}="${raw.slice(0, 20)}" is not a valid count; the frozen panes are ignored.`);
			return null;
		}
		return n;
	};
	const cols = count("xSplit", MAX_COLUMNS);
	const rows = count("ySplit", MAX_ROWS);
	if (cols === null || rows === null || rows === 0 && cols === 0) return null;
	return {
		rows,
		cols
	};
}
// src/util/sheet-not-found.ts
/**
* The `XLV_E_SHEET_NOT_FOUND` error for a sheet name, shared by every API that looks sheets up by
* name: the closest names, a note when only the case differs, and the available names.
*/
function sheetNotFound(name, names) {
	const shown = String(name).slice(0, 100);
	const sameIgnoringCase = names.filter((n) => n.toLowerCase() === shown.toLowerCase());
	const close = sameIgnoringCase.length > 0 ? sameIgnoringCase : suggest(shown, names);
	const listed = names.slice(0, 50).map((n) => `"${n}"`).join(", ");
	return new XlvError("XLV_E_SHEET_NOT_FOUND", `The workbook has no sheet named "${shown}".`, {
		location: { sheet: shown },
		hint: (close.length > 0 ? `Did you mean ${close.map((n) => `"${n}"`).join(" or ")}? ` : "") + `Available sheets: ${listed || "(none)"}${names.length > 50 ? ", …" : ""}. ` + (sameIgnoringCase.length > 0 ? "Sheet names are case-sensitive here: write the name exactly as shown." : "Sheet names are case-sensitive."),
		details: {
			name: shown,
			available: [...names]
		}
	});
}
// src/util/warnings.ts
const MAX_EXAMPLE_CELLS = 10;
/**
* Collects warnings, aggregating repeated occurrences of the same code in the same sheet or
* part into one warning with a count and up to ten example cells. Entries are found through
* nested maps (code, then sheet, then part), so no key is built from file data: a sheet name of
* millions of characters costs nothing per warning.
*/
var WarningCollector = class {
	byCode = /* @__PURE__ */ new Map();
	order = [];
	add(code, message, location = {}) {
		let bySheet = this.byCode.get(code);
		if (bySheet === void 0) {
			bySheet = /* @__PURE__ */ new Map();
			this.byCode.set(code, bySheet);
		}
		const sheet = location.sheet ?? "";
		let byPart = bySheet.get(sheet);
		if (byPart === void 0) {
			byPart = /* @__PURE__ */ new Map();
			bySheet.set(sheet, byPart);
		}
		const part = location.part ?? "";
		const entry = byPart.get(part);
		if (entry === void 0) {
			const { cell, ...rest } = location;
			const created = {
				code,
				message,
				location: rest,
				count: 1,
				cells: cell === void 0 ? [] : [cell]
			};
			byPart.set(part, created);
			this.order.push(created);
			return;
		}
		entry.count++;
		if (location.cell !== void 0 && entry.cells.length < MAX_EXAMPLE_CELLS) entry.cells.push(location.cell);
	}
	/** Warnings in the order their first occurrence was reported. */
	list() {
		return this.order.map((e) => e.cells.length > 0 ? {
			code: e.code,
			message: e.message,
			location: e.location,
			count: e.count,
			cells: [...e.cells]
		} : {
			code: e.code,
			message: e.message,
			location: e.location,
			count: e.count
		});
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
// src/zip/reader.ts
const SIG_LOCAL = 67324752;
const SIG_CENTRAL = 33639248;
const SIG_EOCD = 101010256;
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
const CORRUPT_HINT = "The file is damaged or is not a valid XLSX file. Re-save it in a spreadsheet application or ask the sender for a new copy.";
function corrupt(message, reason, offset, extra = {}) {
	return new XlvError("XLV_E_ZIP_CORRUPT", message, {
		location: offset === void 0 ? {} : { offset },
		hint: CORRUPT_HINT,
		details: {
			reason,
			...extra
		}
	});
}
function unsupported(message, reason, hint, name) {
	return new XlvError("XLV_E_ZIP_UNSUPPORTED", message, {
		location: name === void 0 ? {} : { part: `/${name}` },
		hint,
		details: { reason }
	});
}
function limitExceeded(limit, value, message, name) {
	return new XlvError("XLV_E_LIMIT_EXCEEDED", message, {
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
* `XLV_E_NOT_ZIP`, `XLV_E_UNSUPPORTED_CONTAINER`, `XLV_E_ZIP_CORRUPT`,
* `XLV_E_ZIP_UNSUPPORTED` or `XLV_E_LIMIT_EXCEEDED`.
*/
function openZip(data, limits) {
	checkSource(data.length, data.subarray(0, CFB_SIGNATURE.length), limits);
	const b = Bytes.of(data);
	const dir = locateCentralDirectory(b, limits);
	const entries = readCentralDirectory(b, dir.centralOffset, dir.centralLimit, dir.entryCount);
	validateLocalHeaders(b, entries, dir.centralOffset);
	return new Reader(b, entries, limits);
}
/** Bytes needed at the end of a file to find the end record and a directly preceding ZIP64 locator. */
const TAIL_BYTES = 65577;
/** Bytes needed at the start of a file to recognise CFB containers and local headers. */
const HEAD_BYTES = CFB_SIGNATURE.length;
/** Input size limit and CFB detection, from the size and the first bytes of the input. */
function checkSource(size, head, limits) {
	if (size > limits.maxInputBytes) throw limitExceeded("maxInputBytes", limits.maxInputBytes, `The input is ${size} bytes, more than the limit.`);
	if (CFB_SIGNATURE.every((b, i) => head[i] === b)) throw new XlvError("XLV_E_UNSUPPORTED_CONTAINER", "The file is a Compound File Binary container, not a ZIP-based XLSX file.", { hint: "This is either a legacy .xls workbook or a password-protected (encrypted) .xlsx workbook. Open it in a spreadsheet application and save it as an unencrypted .xlsx file." });
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
		if (b.u32(locator + 4) !== 0 || b.u32(locator + 16) !== 1) throw unsupported("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the workbook as a single file.");
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
	if (thisDisk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount) throw unsupported("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the workbook as a single file.");
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
		if (data[i] === 80 && data[i + 1] === 75 && b.u32(at) === SIG_EOCD) {
			if (at + EOCD_SIZE + b.u16(at + 20) === size) return at;
		}
	}
	if (size >= 4 && b.u32(0) === SIG_LOCAL) throw corrupt("The end of the ZIP archive is missing; the file is probably truncated.", "eocd-missing");
	if (size >= EOCD_SIZE && data.indexOf(80) !== -1 && containsSignature(data, 0)) throw corrupt("The end of central directory record is followed by unexpected data.", "trailing-data");
	throw new XlvError("XLV_E_NOT_ZIP", "The input is not a ZIP archive, so it cannot be an XLSX file.", { hint: "XLSX files are ZIP archives. Check that the file is an .xlsx (or .xlsm) workbook; for CSV or text files use a CSV reader." });
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
		if (p + CENTRAL_SIZE > end || b.u32(p) !== SIG_CENTRAL) throw corrupt("A central directory entry is missing or damaged.", "central-directory-entry", p);
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
		if (flags & 8257 || method === 99) throw unsupported(`ZIP entry "${name}" is encrypted.`, "encrypted", "Encrypted ZIP entries are not supported. Save the workbook again without ZIP encryption.", name);
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
		if (diskStart !== 0) throw unsupported("The archive spans several disks.", "multi-disk", "Multi-volume ZIP archives are not supported. Save the workbook as a single file.");
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
		if (at + 30 > centralOffset || b.u32(at) !== SIG_LOCAL) throw corrupt(`The local header of "${entry.name}" is missing or damaged.`, "local-header", at, { name: entry.name });
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
		if (entry === void 0) throw new XlvError("XLV_E_PART_MISSING", `The archive has no entry "${String(entryOrName)}".`, {
			location: { part: `/${String(entryOrName)}` },
			hint: "The workbook is incomplete. Re-save it in a spreadsheet application."
		});
		const { limits } = this;
		if (entry.uncompressedSize > limits.maxEntryBytes) throw limitExceeded("maxEntryBytes", limits.maxEntryBytes, `ZIP entry "${entry.name}" declares ${entry.uncompressedSize} bytes, more than the limit.`, entry.name);
		if (this.bytesRead + entry.uncompressedSize > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
		const compressed = this.b.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize);
		let out;
		if (entry.method === 0) {
			if (entry.compressedSize !== entry.uncompressedSize) throw corrupt(`Stored ZIP entry "${entry.name}" has inconsistent sizes.`, "size-mismatch", entry.headerOffset, { name: entry.name });
			out = compressed;
		} else if (entry.method === 8) out = this.inflate(entry, compressed);
		else throw unsupported(`ZIP entry "${entry.name}" uses compression method ${entry.method}.`, "compression-method", "Only stored and Deflate-compressed entries are supported. Re-save the workbook in a spreadsheet application.", entry.name);
		if (out.length !== entry.uncompressedSize) throw corrupt(`ZIP entry "${entry.name}" is ${out.length} bytes, but its header says ${entry.uncompressedSize}.`, "size-mismatch", entry.headerOffset, { name: entry.name });
		if (crc32(out) !== entry.crc32) throw corrupt(`ZIP entry "${entry.name}" failed its CRC-32 check.`, "crc-mismatch", entry.headerOffset, { name: entry.name });
		this.bytesRead += out.length;
		return out;
	}
	inflate(entry, compressed) {
		const { limits } = this;
		const chunks = [];
		let produced = 0;
		let ended = false;
		const inflater = new Inflate();
		inflater.ondata = (chunk, final) => {
			produced += chunk.length;
			if (produced > limits.maxEntryBytes) throw limitExceeded("maxEntryBytes", limits.maxEntryBytes, `ZIP entry "${entry.name}" decompresses to more than the limit.`, entry.name);
			if (this.bytesRead + produced > limits.maxTotalBytes) throw limitExceeded("maxTotalBytes", limits.maxTotalBytes, "The decompressed parts of the file are larger than the limit.", entry.name);
			if (produced > limits.compressionRatioMinBytes && produced / Math.max(1, entry.compressedSize) > limits.maxCompressionRatio) throw limitExceeded("maxCompressionRatio", limits.maxCompressionRatio, `ZIP entry "${entry.name}" expands more than ${String(limits.maxCompressionRatio)} times, which is typical of a zip bomb.`, entry.name);
			if (chunk.length > 0) chunks.push(chunk);
			if (final) ended = true;
		};
		try {
			if (compressed.length === 0) inflater.push(compressed, true);
			for (let at = 0; at < compressed.length; at += INFLATE_CHUNK) inflater.push(compressed.subarray(at, at + INFLATE_CHUNK), at + INFLATE_CHUNK >= compressed.length);
		} catch (error) {
			if (error instanceof XlvError) throw error;
			throw new XlvError("XLV_E_ZIP_CORRUPT", `ZIP entry "${entry.name}" is not a valid Deflate stream.`, {
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
// src/read.ts
const OPTION_NAMES$1 = ["limits"];
/** Content types of the workbook part (ECMA-376 Part 1 §12.3.23; [MS-OFFMACRO2] §2.2.1.4). */
const FORMATS = /* @__PURE__ */ new Map([
	["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "xlsx"],
	["application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml", "xltx"],
	["application/vnd.ms-excel.sheet.macroEnabled.main+xml", "xlsm"],
	["application/vnd.ms-excel.template.macroEnabled.main+xml", "xltm"]
]);
var ReadSheet = class {
	name;
	index;
	kind;
	state;
	part;
	data;
	sheetId;
	hyperlinks;
	notes;
	names;
	sortedRows = null;
	/** Pictures of the sheet's drawing and its background image (TD-30 … TD-39). */
	images = [];
	backgroundImage = null;
	/** Charts of the sheet's drawing (D-K01). */
	charts = [];
	/** Tables of the sheet (TD-01). */
	tables = [];
	constructor(name, index, kind, state, part, data, sheetId, hyperlinks = [], notes = [], names = []) {
		this.name = name;
		this.index = index;
		this.kind = kind;
		this.state = state;
		this.part = part;
		this.data = data;
		this.sheetId = sheetId;
		this.hyperlinks = hyperlinks;
		this.notes = notes;
		this.names = names;
	}
	get protection() {
		return this.data?.print.protection ?? null;
	}
	get pageSetup() {
		return this.data?.print.pageSetup ?? DEFAULT_PAGE_SETUP;
	}
	get headerFooter() {
		return this.data?.print.headerFooter ?? null;
	}
	get pageBreaks() {
		return {
			rows: this.data?.print.rowBreaks ?? [],
			columns: this.data?.print.colBreaks ?? []
		};
	}
	get printArea() {
		const n = this.names.find((d) => d.sheet === this.name && d.name.toLowerCase() === "_xlnm.print_area");
		return n === void 0 ? null : parsePrintArea(n.refersTo, this.name);
	}
	get printTitles() {
		const n = this.names.find((d) => d.sheet === this.name && d.name.toLowerCase() === "_xlnm.print_titles");
		return n === void 0 ? null : parsePrintTitles(n.refersTo, this.name);
	}
	get dimension() {
		return this.data?.dimension ?? null;
	}
	get merges() {
		return this.data?.merges ?? [];
	}
	get cellCount() {
		return this.data?.cellCount ?? 0;
	}
	get columns() {
		return this.data?.columns ?? [];
	}
	get freeze() {
		return this.data?.freeze ?? null;
	}
	get autoFilter() {
		return this.data?.autoFilter ?? null;
	}
	get tabColor() {
		return this.data?.tabColor ?? null;
	}
	get dataValidations() {
		return this.data?.dataValidations ?? [];
	}
	get conditionalFormats() {
		return this.data?.conditionalFormats ?? [];
	}
	get properties() {
		return this.data?.properties ?? DEFAULT_SHEET_PROPERTIES;
	}
	get view() {
		return this.data?.view ?? DEFAULT_SHEET_VIEW;
	}
	rowInfo(row) {
		return this.data?.rowInfo.get(row);
	}
	cell(ref) {
		const address = typeof ref === "string" ? parseCellRef(ref.toUpperCase()) : null;
		if (address === null) throw new XlvError("XLV_E_INVALID_CELL_REF", `"${String(ref).slice(0, 40)}" is not a valid cell reference.`, {
			location: { sheet: this.name },
			hint: "Pass an A1-style reference such as \"B7\" (column letters A to XFD, rows 1 to 1048576).",
			details: { ref: String(ref).slice(0, 40) }
		});
		return this.data?.rows.get(address.row)?.get(address.col);
	}
	*rows() {
		if (this.sortedRows === null) {
			const rows = this.data?.rows ?? /* @__PURE__ */ new Map();
			this.sortedRows = [...rows.keys()].sort((a, b) => a - b).map((r) => {
				const cells = rows.get(r) ?? /* @__PURE__ */ new Map();
				return [...cells.keys()].sort((a, b) => a - b).map((c) => cells.get(c));
			});
		}
		yield* this.sortedRows;
	}
	*cells() {
		for (const row of this.rows()) yield* row;
	}
	/** Number of cells that hold only a style (D-E06), without building them. */
	get styledCellCount() {
		let n = 0;
		for (const list of this.data?.styleOnly.values() ?? []) for (let i = 1; i < list.length; i += 2) if (list[i] !== -1) n++;
		return n;
	}
	*styledCells() {
		const data = this.data;
		if (!data) return;
		for (const r of [...data.styleOnly.keys()].sort((a, b) => a - b)) yield* styledCellsOf(r, data.styleOnly.get(r) ?? [], data.styledCell);
	}
};
/** Number of cells of a sheet that hold only a style (they count against `maxCells` but not `cellCount`). */
function styledCellCount(sheet) {
	if (sheet instanceof ReadSheet) return sheet.styledCellCount;
	let n = 0;
	for (const _cell of sheet.styledCells()) n++;
	return n;
}
var ReadWorkbook = class {
	sheets;
	dateSystem;
	conformance;
	format;
	hasMacros;
	externalLinkCount;
	warnings;
	properties;
	definedNames;
	protection;
	views;
	calculation;
	defaultFont;
	/** Image parts used by pictures and backgrounds, each once (TD-31). */
	images = [];
	constructor(sheets, dateSystem, conformance, format, hasMacros, externalLinkCount, warnings, properties, definedNames, protection = null, views = [], calculation = null, defaultFont = null) {
		this.sheets = sheets;
		this.dateSystem = dateSystem;
		this.conformance = conformance;
		this.format = format;
		this.hasMacros = hasMacros;
		this.externalLinkCount = externalLinkCount;
		this.warnings = warnings;
		this.properties = properties;
		this.definedNames = definedNames;
		this.protection = protection;
		this.views = views;
		this.calculation = calculation;
		this.defaultFont = defaultFont;
	}
	getSheet(name) {
		const sheet = this.sheets.find((s) => s.name === name);
		if (sheet !== void 0) return sheet;
		throw sheetNotFound(name, this.sheets.map((s) => s.name));
	}
};
function checkOptions(options) {
	if (options === void 0) return {};
	if (options === null || typeof options !== "object" || Array.isArray(options)) throw new XlvError("XLV_E_INVALID_OPTION", "Options must be an object.", { hint: "Pass an object such as { limits: { maxCells: 1000000 } }, or leave the options out." });
	for (const key of Object.keys(options)) if (!OPTION_NAMES$1.includes(key)) {
		const close = suggest(key, OPTION_NAMES$1);
		throw new XlvError("XLV_E_INVALID_OPTION", `Unknown option "${key}".`, {
			hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known options: ${OPTION_NAMES$1.join(", ")}.`,
			details: {
				option: key,
				known: OPTION_NAMES$1
			}
		});
	}
	return options;
}
/** Converts the accepted input types to a Uint8Array. */
function toBytes(data) {
	if (data instanceof Uint8Array) return data;
	if (data instanceof ArrayBuffer) return new Uint8Array(data);
	if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
	throw new XlvError("XLV_E_INVALID_INPUT", `Expected the file contents as a Uint8Array or ArrayBuffer, got ${data === null ? "null" : typeof data}.`, { hint: typeof data === "string" ? "Pass the bytes of the file, not its path. In Node.js use readWorkbookFile() from \"xlvane/node\", or read the file with fs.readFile() first." : "Pass the bytes of the file, for example from fetch(...).arrayBuffer(), File.arrayBuffer() or fs.readFile()." });
}
function findRelationship(rels, name) {
	return rels.find((r) => !r.external && classifyRelationship(r.type)?.name === name);
}
/**
* Reads an XLSX workbook (also .xlsm, .xltx, .xltm) from memory. The input is treated as
* untrusted: ZIP and XML limits apply, DTDs are rejected, macros are never run and external
* links are never followed. Formulas are not calculated; formula cells hold the cached value
* stored in the file.
*
* @example
* ```ts
* import { readWorkbook } from 'xlvane';
*
* const bytes = new Uint8Array(await (await fetch('/report.xlsx')).arrayBuffer());
* const workbook = readWorkbook(bytes);
* const sheet = workbook.getSheet('Sales');
* for (const row of sheet.rows()) console.log(row.map((cell) => cell.value));
* ```
*/
function readWorkbook(data, options) {
	const bytes = toBytes(data);
	const limits = resolveLimits(checkOptions(options).limits);
	const pkg = openPackage(openZip(bytes, limits), limits);
	const warnings = new WarningCollector();
	const book = analyzeWorkbook(pkg, limits, warnings);
	const budget = {
		remaining: limits.maxCells,
		sharedText: limits.maxSharedStringChars,
		formulaText: limits.maxSharedStringChars
	};
	const formatCache = /* @__PURE__ */ new Map();
	const media = /* @__PURE__ */ new Map();
	const chartCache = {
		parts: /* @__PURE__ */ new Map(),
		budget
	};
	const drawingIds = [];
	const sheets = book.sheets.map((entry) => {
		let data = null;
		if (entry.kind === "worksheet" && entry.part !== null) data = parseWorksheet(pkg.readXml(entry.part) ?? "", {
			part: entry.part,
			sheet: entry.name,
			limits,
			warnings,
			sharedStrings: book.sharedStrings,
			richStrings: book.richStrings,
			styles: book.styles,
			dateSystem: book.dateSystem,
			conformance: book.conformance,
			budget,
			formatCache
		});
		const links = data === null || entry.part === null ? [] : resolveHyperlinks(pkg, entry.part, entry.name, data, warnings);
		const notes = data === null || entry.part === null ? [] : readNotes(pkg, entry.part, entry.name, limits, warnings, entry.foreignParts);
		if (data !== null) attachNotes(data.rows, notes);
		const sheet = new ReadSheet(entry.name, entry.index, entry.kind, entry.state, entry.part, data, entry.sheetId, links, notes, book.definedNames);
		if (data !== null && entry.part !== null) sheet.tables = readTables(pkg, entry.part, entry.name, limits, warnings, entry.foreignParts);
		drawingIds.push(data === null || entry.part === null ? null : {
			drawing: data.drawingRelId ?? null,
			picture: data.pictureRelId ?? null
		});
		return sheet;
	});
	const drawingOwners = /* @__PURE__ */ new Map();
	book.sheets.forEach((entry, i) => {
		const ids = drawingIds[i];
		const sheet = sheets[i];
		if (ids === null || ids === void 0 || sheet === void 0 || entry.part === null) return;
		const drawing = readSheetDrawing(pkg, entry.part, entry.name, ids, {
			limits,
			warnings
		}, media, chartCache, drawingOwners);
		sheet.images = drawing.images;
		sheet.charts = drawing.charts;
		sheet.backgroundImage = drawing.background;
	});
	const workbook = new ReadWorkbook(sheets, book.dateSystem, book.conformance, book.format, book.hasMacros, book.externalLinkCount, warnings.list(), book.properties, book.definedNames, book.protection, book.views, book.calculation, book.styles.defaultFont() ?? null);
	workbook.images = [...media.values()].filter((m) => m !== null);
	return workbook;
}
/**
* Resolves the `r:id` of each hyperlink against the sheet's relationships (W2-42, W2-44) and
* attaches the links to the cells they cover. Targets are only read as text, never opened.
*/
function resolveHyperlinks(pkg, part, sheet, data, warnings) {
	if (data.hyperlinks.length === 0) return [];
	const rels = /* @__PURE__ */ new Map();
	const relsList = pkg.relationships(part);
	for (const r of relsList) if (!rels.has(r.id)) rels.set(r.id, r);
	const links = [];
	let unusual = null;
	for (const raw of data.hyperlinks) {
		let target = null;
		if (raw.relId !== null) {
			const rel = rels.get(raw.relId);
			if (rel === void 0 || classifyRelationship(rel.type)?.name !== "hyperlink") warnings.add("XLV_W_INVALID_VALUE", `The hyperlink of ${raw.ref} refers to relationship "${raw.relId.slice(0, 40)}", which is not a hyperlink relationship of the sheet; its target is unknown.`, {
				part,
				sheet,
				cell: raw.ref
			});
			else {
				target = rel.target;
				if (unusual === null && isUnusualTarget(target)) unusual = target;
			}
		}
		links.push(Object.freeze({
			ref: raw.ref,
			target,
			location: raw.location,
			tooltip: raw.tooltip,
			display: raw.display
		}));
	}
	if (unusual !== null) warnings.add("XLV_W_UNUSUAL_HYPERLINK", `The sheet has hyperlinks to targets other than web pages or e-mail addresses, for example "${unusual.slice(0, 80)}". xlvane does not open them.`, {
		part,
		sheet
	});
	let total = 0;
	for (const cells of data.rows.values()) total += cells.size;
	let linked = 0;
	let visits = 0;
	const maxVisits = 16 * total + 1e6;
	for (const link of links) {
		if (linked === total) break;
		if (visits > maxVisits) {
			warnings.add("XLV_W_LIMIT_REACHED", `The sheet's hyperlinks cover far more positions than it has cells; cell.hyperlink is not set for the cells of some links. sheet.hyperlinks lists every link.`, {
				part,
				sheet
			});
			break;
		}
		const range = parseRange(link.ref);
		if (range === null) continue;
		const rowSpan = range.end.row - range.start.row + 1;
		const colSpan = range.end.col - range.start.col + 1;
		const visit = (cells) => {
			const mark = (c, cell) => {
				if (cell.hyperlink === void 0) {
					cells.set(c, {
						...cell,
						hyperlink: link
					});
					linked++;
				}
			};
			if (colSpan <= cells.size) {
				visits += colSpan;
				for (let c = range.start.col; c <= range.end.col; c++) {
					const cell = cells.get(c);
					if (cell) mark(c, cell);
				}
			} else {
				visits += cells.size;
				for (const [c, cell] of cells) if (c >= range.start.col && c <= range.end.col) mark(c, cell);
			}
		};
		if (rowSpan <= data.rows.size) {
			visits += rowSpan;
			for (let r = range.start.row; r <= range.end.row; r++) {
				const cells = data.rows.get(r);
				if (cells) visit(cells);
			}
		} else {
			visits += data.rows.size;
			for (const [r, cells] of data.rows) if (r >= range.start.row && r <= range.end.row) visit(cells);
		}
	}
	return links;
}
/**
* The notes of a worksheet: the Comments part named by an internal relationship of the sheet
* (W2-46), read as text, with the settings of their boxes from the sheet's VML parts (W2-69). Threaded comments
* ([MS-XLSX] §2.3.7) are reported with `XLV_W_THREADED_COMMENTS`; their legacy placeholders are
* ordinary notes.
*/
function readNotes(pkg, part, sheet, limits, warnings, foreign = NO_PARTS) {
	const rels = pkg.relationships(part).filter((r) => r.external || !foreign.has(partKey(r.target)));
	if (rels.some((r) => r.type === "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment")) warnings.add("XLV_W_THREADED_COMMENTS", "The sheet has threaded comments (conversations); xlvane reads only the notes stored with them, whose author is \"tc={id}\" and whose text may not include the replies.", {
		part,
		sheet
	});
	const rel = rels.find((r) => !r.external && classifyRelationship(r.type)?.name === "comments");
	if (rel === void 0) return [];
	const xml = pkg.readXml(rel.target);
	if (xml === void 0) {
		warnings.add("XLV_W_PART_MISSING", "The sheet refers to a comments part that does not exist; its notes are not read.", {
			part: rel.target,
			sheet
		});
		return [];
	}
	const notes = parseComments(xml, {
		part: rel.target,
		sheet,
		limits,
		warnings
	});
	if (notes.length === 0) return notes;
	const seen = /* @__PURE__ */ new Set();
	const byRef = /* @__PURE__ */ new Map();
	notes.forEach((note, i) => {
		const list = byRef.get(note.ref);
		if (list === void 0) byRef.set(note.ref, [i]);
		else list.push(i);
	});
	for (const vml of rels) {
		if (vml.external || classifyRelationship(vml.type)?.name !== "vmlDrawing") continue;
		const key = partKey(vml.target);
		if (seen.has(key)) continue;
		seen.add(key);
		const text = pkg.readXml(vml.target);
		if (text === void 0) continue;
		const boxes = parseNoteBoxes(text, {
			part: vml.target,
			sheet,
			limits,
			warnings
		});
		if (boxes.size === 0) continue;
		for (const [ref, box] of boxes) for (const i of byRef.get(ref) ?? []) if (notes[i].box === void 0) notes[i] = Object.freeze({
			...notes[i],
			box
		});
	}
	return notes;
}
/**
* The tables of a worksheet: the Table Definition parts named by the sheet's internal `table`
* relationships (Part 1 §12.3.21, §18.3.1.94; TD-01), in relationship order. A missing or
* invalid part gives a warning, never an error.
*/
function readTables(pkg, part, sheet, limits, warnings, foreign = NO_PARTS) {
	const tables = [];
	const seen = /* @__PURE__ */ new Set();
	for (const rel of pkg.relationships(part)) {
		if (rel.external || classifyRelationship(rel.type)?.name !== "table") continue;
		const key = partKey(rel.target);
		if (seen.has(key) || foreign.has(key)) continue;
		seen.add(key);
		if (tables.length >= 4096) {
			warnings.add("XLV_W_INVALID_VALUE", `The sheet has more than ${String(MAX_TABLES_PER_SHEET)} tables; the rest are not read.`, {
				part,
				sheet
			});
			break;
		}
		const xml = pkg.readXml(rel.target);
		if (xml === void 0) {
			warnings.add("XLV_W_PART_MISSING", "The sheet refers to a table part that does not exist; that table is not read.", {
				part: rel.target,
				sheet
			});
			continue;
		}
		const table = parseTable(xml, {
			part: rel.target,
			sheet,
			limits,
			warnings
		});
		if (table !== null) tables.push(table);
	}
	return tables;
}
/** Sets `note` on the existing cells that have one (the first note of a cell wins). */
function attachNotes(rows, notes) {
	for (const note of notes) {
		const address = parseCellRef(note.ref);
		if (address === null) continue;
		const cells = rows.get(address.row);
		const cell = cells?.get(address.col);
		if (cells && cell && cell.note === void 0) cells.set(address.col, {
			...cell,
			note
		});
	}
}
const NO_PARTS = /* @__PURE__ */ new Set();
/** Relationship types whose parts are read per sheet, besides the worksheet itself. */
const SHEET_RELATED = /* @__PURE__ */ new Set([
	"comments",
	"vmlDrawing",
	"table"
]);
/** The Core Properties relationship of the package, if any (D-A02). */
function corePropertiesRel(pkg) {
	return pkg.relationships(null).find((r) => !r.external && isCorePropertiesRel(r.type));
}
/** The Extended File Properties relationship of the package, if any (D-A10). */
function appPropertiesRel(pkg) {
	return pkg.relationships(null).find((r) => !r.external && isAppPropertiesRel(r.type));
}
/** The workbook part's relationships that the reader loads in full (workbook, styles, shared strings). */
function workbookPartsToLoad(pkg) {
	const officeDocument = findRelationship(pkg.relationships(null), "officeDocument");
	if (officeDocument === void 0) return [];
	const rels = pkg.relationships(officeDocument.target);
	const out = [officeDocument.target];
	for (const name of ["sharedStrings", "styles"]) {
		const rel = findRelationship(rels, name);
		if (rel !== void 0) out.push(rel.target);
	}
	const core = corePropertiesRel(pkg);
	if (core !== void 0) out.push(core.target);
	const app = appPropertiesRel(pkg);
	if (app !== void 0) out.push(app.target);
	return out;
}
/** Reads the workbook structure, shared strings and styles (D-401). */
function analyzeWorkbook(pkg, limits, warnings) {
	const officeDocument = findRelationship(pkg.relationships(null), "officeDocument");
	if (officeDocument === void 0 || !pkg.hasPart(officeDocument.target)) throw new XlvError("XLV_E_PART_MISSING", "The package has no workbook part.", {
		location: officeDocument === void 0 ? { part: "/_rels/.rels" } : { part: officeDocument.target },
		hint: "This is not a complete spreadsheet: the package does not point to a workbook. Check that the file is an Excel workbook and re-save it in a spreadsheet application."
	});
	const workbookPart = officeDocument.target;
	const format = FORMATS.get(pkg.contentType(workbookPart) ?? "") ?? "unknown";
	const info = parseWorkbookPart(pkg.readXml(workbookPart) ?? "", {
		part: workbookPart,
		limits,
		warnings
	});
	const rels = pkg.relationships(workbookPart);
	const relsById = new Map(rels.map((r) => [r.id, r]));
	const context = {
		limits,
		warnings
	};
	const readOptionalPart = (name, what) => {
		const rel = findRelationship(rels, name);
		if (rel === void 0) return void 0;
		const xml = pkg.readXml(rel.target);
		if (xml === void 0) {
			warnings.add("XLV_W_PART_MISSING", `The workbook refers to a ${what} part that does not exist; reading continues without it.`, { part: rel.target });
			return;
		}
		return {
			part: rel.target,
			xml
		};
	};
	const sst = readOptionalPart("sharedStrings", "shared strings");
	const richStrings = /* @__PURE__ */ new Map();
	const sharedStrings = sst === void 0 ? [] : parseSharedStrings(sst.xml, {
		part: sst.part,
		...context
	}, richStrings);
	const stylesPart = readOptionalPart("styles", "styles");
	const styles = stylesPart === void 0 ? EMPTY_STYLES : parseStyles(stylesPart.xml, {
		part: stylesPart.part,
		...context
	});
	const hasMacros = detectMacros(pkg, rels);
	if (hasMacros) warnings.add("XLV_W_MACROS_PRESENT", "The workbook contains macros (a VBA project or Excel 4.0 macro sheets). xlvane never runs them.", { part: workbookPart });
	const externalLinkCount = rels.filter((r) => classifyRelationship(r.type)?.name === "externalLink").length;
	if (externalLinkCount > 0) warnings.add("XLV_W_EXTERNAL_LINKS", `The workbook links to ${externalLinkCount} external workbook(s). xlvane never follows the links; linked values are the cached copies.`, { part: workbookPart });
	const dateSystem = info.date1904 ? "1904" : "1900";
	const owners = /* @__PURE__ */ new Map();
	const sheets = info.sheets.map((entry, index) => {
		if (entry.state === "veryHidden") warnings.add("XLV_W_VERY_HIDDEN_SHEET", `Sheet "${entry.name}" is very hidden: spreadsheet applications do not offer a way to unhide it.`, { sheet: entry.name });
		const rel = entry.relId === null ? void 0 : relsById.get(entry.relId);
		const relName = rel === void 0 ? void 0 : classifyRelationship(rel.type)?.name;
		const kind = rel !== void 0 && (rel.type === "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" || rel.type === "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet") ? "macrosheet" : relName === "chartsheet" || relName === "dialogsheet" ? relName : "worksheet";
		if (rel === void 0 || rel.external || !pkg.hasPart(rel.target)) {
			warnings.add("XLV_W_SHEET_PART_MISSING", `The part of sheet "${entry.name}" is missing; the sheet is returned without cells.`, {
				sheet: entry.name,
				part: workbookPart
			});
			return {
				name: entry.name,
				index,
				sheetId: entry.sheetId,
				kind,
				state: entry.state,
				part: null,
				foreignParts: NO_PARTS
			};
		}
		const owner = owners.get(partKey(rel.target));
		if (owner !== void 0) {
			warnings.add("XLV_W_SHEET_PART_SHARED", `Sheet "${entry.name}" has the same part as sheet "${owner}"; the part is read for "${owner}" only, and "${entry.name}" is returned without cells.`, {
				sheet: entry.name,
				part: rel.target
			});
			return {
				name: entry.name,
				index,
				sheetId: entry.sheetId,
				kind,
				state: entry.state,
				part: null,
				foreignParts: NO_PARTS
			};
		}
		owners.set(partKey(rel.target), entry.name);
		const foreign = /* @__PURE__ */ new Set();
		for (const r of pkg.relationships(rel.target)) {
			const related = classifyRelationship(r.type)?.name ?? "";
			if (r.external || !SHEET_RELATED.has(related)) continue;
			const key = partKey(r.target);
			const first = owners.get(key);
			if (first === void 0) owners.set(key, entry.name);
			else if (first !== entry.name && !foreign.has(key)) {
				foreign.add(key);
				warnings.add("XLV_W_SHEET_PART_SHARED", `Sheet "${entry.name}" relates to the part ${r.target}, which belongs to sheet "${first}"; it is read for "${first}" only.`, {
					sheet: entry.name,
					part: r.target
				});
			}
		}
		return {
			name: entry.name,
			index,
			sheetId: entry.sheetId,
			kind,
			state: entry.state,
			part: rel.target,
			foreignParts: foreign
		};
	});
	const definedNames = info.definedNames.map((n) => {
		let sheet = null;
		if (n.localSheetId !== null) {
			const entry = info.sheets[n.localSheetId];
			if (entry === void 0) warnings.add("XLV_W_INVALID_VALUE", `The defined name "${n.name.slice(0, 100)}" belongs to sheet number ${String(n.localSheetId)}, which does not exist; it is read as a workbook-level name.`, { part: workbookPart });
			else sheet = entry.name;
		}
		return {
			name: n.name,
			refersTo: n.refersTo,
			sheet,
			hidden: n.hidden,
			comment: n.comment
		};
	});
	let properties = {};
	const core = corePropertiesRel(pkg);
	if (core !== void 0) {
		const xml = pkg.readXml(core.target);
		if (xml === void 0) warnings.add("XLV_W_PART_MISSING", "The package refers to a document properties part that does not exist; reading continues without it.", { part: core.target });
		else properties = parseCoreProperties(xml, {
			part: core.target,
			...context
		});
	}
	const app = appPropertiesRel(pkg);
	if (app !== void 0) {
		const xml = pkg.readXml(app.target);
		if (xml === void 0) warnings.add("XLV_W_PART_MISSING", "The package refers to an extended document properties part that does not exist; reading continues without it.", { part: app.target });
		else properties = {
			...properties,
			...parseAppProperties(xml, {
				part: app.target,
				...context
			})
		};
	}
	return {
		workbookPart,
		sheets,
		sharedStrings,
		richStrings,
		styles,
		dateSystem,
		conformance: info.conformance,
		format,
		hasMacros,
		externalLinkCount,
		properties,
		definedNames,
		protection: info.protection,
		views: info.views,
		calculation: info.calculation
	};
}
function detectMacros(pkg, workbookRels) {
	if (workbookRels.some((r) => r.type === "http://schemas.microsoft.com/office/2006/relationships/vbaProject" || r.type === "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" || r.type === "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet")) return true;
	return pkg.partNames.some((name) => {
		const type = pkg.contentType(name);
		return type === "application/vnd.ms-office.vbaProject" || type !== void 0 && MACROSHEET_CONTENT_TYPES.includes(type);
	});
}
// src/util/iso-date.ts
function pad2(n) {
	return n < 10 ? `0${String(n)}` : String(n);
}
/**
* ISO 8601 without a time zone, because cell dates have none: the date alone at midnight
* (`2024-01-31`), otherwise date and time (`2024-01-31T13:45:00`, with `.sss` only when there are
* milliseconds). The UTC fields of the `Date` are the cell's wall-clock date and time (V-* facts).
*/
function isoDate(d) {
	const date = `${String(d.getUTCFullYear()).padStart(4, "0")}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
	const ms = d.getUTCMilliseconds();
	if (d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && ms === 0) return date;
	return `${date}T${`${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}`}${ms === 0 ? "" : `.${String(ms).padStart(3, "0")}`}`;
}
// src/inspect.ts
const OPTION_NAMES = [
	"limits",
	"maxColumns",
	"headerScanRows",
	"formulaSamples"
];
/** Upper bound on the steps spent looking for values hidden under merged ranges, per sheet. */
const MAX_MERGE_CHECK_WORK = 5e6;
/** Bounds of the lists that are not controlled by an option. */
const MAX_DEFINED_NAMES = 100;
const MAX_RICH_TEXT_CELLS = 10;
const MAX_HYPERLINK_SAMPLES = 10;
/** Merged ranges listed per sheet; `mergedRangeCount` gives the total. */
const MAX_MERGED_RANGES = 100;
/** Chart group types listed per chart; `typesTruncated` says whether there are more. */
const MAX_CHART_TYPES = 10;
/** Characters shown of a defined name's formula (`refersTo`), of a chart title and of a text value. */
const MAX_SHOWN_TEXT = 100;
/** Cuts to `MAX_SHOWN_TEXT` characters and `…`, before a surrogate pair that would be split. */
const cut = (text) => text.length > MAX_SHOWN_TEXT ? `${cutAt(text, MAX_SHOWN_TEXT)}…` : text;
const TYPE_ORDER = [
	"number",
	"string",
	"date",
	"boolean",
	"error"
];
function resolveOptions(options) {
	if (options === void 0) return {
		limits: void 0,
		resolved: {
			maxColumns: 200,
			headerScanRows: 20,
			formulaSamples: 10
		}
	};
	if (options === null || typeof options !== "object" || Array.isArray(options)) throw new XlvError("XLV_E_INVALID_OPTION", "Options must be an object.", { hint: "Pass an object such as { maxColumns: 50 }, or leave the options out." });
	const record = options;
	for (const key of Object.keys(record)) if (!OPTION_NAMES.includes(key)) {
		const close = suggest(key, OPTION_NAMES);
		throw new XlvError("XLV_E_INVALID_OPTION", `Unknown option "${key}".`, {
			hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known options: ${OPTION_NAMES.join(", ")}.`,
			details: {
				option: key,
				known: [...OPTION_NAMES]
			}
		});
	}
	const count = (name, fallback, min) => {
		if (!Object.hasOwn(record, name)) return fallback;
		const value = record[name];
		if (typeof value !== "number" || !Number.isSafeInteger(value) || value < min) throw new XlvError("XLV_E_INVALID_OPTION", `Option "${name}" must be a whole number of at least ${String(min)}.`, {
			hint: `Pass for example { ${name}: ${String(fallback)} }, or leave it out to use the default.`,
			details: { option: name }
		});
		return value;
	};
	return {
		limits: Object.hasOwn(record, "limits") ? record["limits"] : void 0,
		resolved: {
			maxColumns: count("maxColumns", 200, 1),
			headerScanRows: count("headerScanRows", 20, 1),
			formulaSamples: count("formulaSamples", 10, 0)
		}
	};
}
function isWorkbook$1(input) {
	return input !== null && typeof input === "object" && !ArrayBuffer.isView(input) && !(input instanceof ArrayBuffer) && typeof input.getSheet === "function" && Array.isArray(input.sheets);
}
/** Formats a value for JSON: dates as ISO 8601 without a time zone (`isoDate`). */
function jsonValue(value) {
	if (value instanceof Date) return isoDate(value);
	return typeof value === "string" ? cut(value) : value;
}
/**
* Describes the structure of a workbook as plain JSON data: sheets and their visibility, used
* ranges, detected header rows, column types with example values, merged cells, formulas and
* warnings. Accepts the file contents or a workbook returned by `readWorkbook`.
*
* @example
* ```ts
* import { inspect } from 'xlvane';
*
* declare const bytes: Uint8Array; // the contents of an .xlsx file
* const report = inspect(bytes);
* for (const sheet of report.sheets) {
*   console.log(sheet.name, sheet.usedRange, sheet.columns.map((c) => `${c.header}: ${c.dominantType}`));
* }
* console.log(JSON.stringify(report, null, 2));
* ```
*/
function inspect(input, options) {
	const { limits, resolved } = resolveOptions(options);
	const workbook = isWorkbook$1(input) ? input : readWorkbook(input, limits === void 0 ? {} : { limits });
	const warnings = new WarningCollector();
	const sheets = workbook.sheets.map((sheet) => inspectSheet(sheet, resolved, warnings));
	return {
		schemaVersion: 1,
		format: workbook.format,
		conformance: workbook.conformance,
		dateSystem: workbook.dateSystem,
		hasMacros: workbook.hasMacros,
		externalLinkCount: workbook.externalLinkCount,
		properties: propertiesJson(workbook),
		definedNames: workbook.definedNames.slice(0, MAX_DEFINED_NAMES).map((n) => ({
			name: n.name,
			refersTo: cut(n.refersTo),
			sheet: n.sheet,
			hidden: n.hidden
		})),
		definedNamesTruncated: workbook.definedNames.length > MAX_DEFINED_NAMES,
		sheetCount: workbook.sheets.length,
		sheets,
		warnings: [...workbook.warnings, ...warnings.list()]
	};
}
function propertiesJson(workbook) {
	const out = {};
	for (const [key, value] of Object.entries(workbook.properties)) if (value instanceof Date) out[key] = value.toISOString();
	else if (typeof value === "string") out[key] = value;
	return out;
}
/** Column ranges (`"B"`, `"D:F"`) that have a default style. */
function styledColumns(sheet, max) {
	const out = [];
	for (const c of sheet.columns) {
		if (c.style === void 0 || out.length >= max) continue;
		out.push(c.min === c.max ? columnLetters(c.min) : `${columnLetters(c.min)}:${columnLetters(c.max)}`);
	}
	return out;
}
function inspectSheet(sheet, options, warnings) {
	const rows = [...sheet.rows()];
	let minRow = Number.POSITIVE_INFINITY;
	let maxRow = 0;
	let minCol = Number.POSITIVE_INFINITY;
	let maxCol = 0;
	let formulaCount = 0;
	let richTextCellCount = 0;
	const richTextCells = [];
	const formulas = [];
	for (const cells of rows) for (const cell of cells) {
		minRow = Math.min(minRow, cell.row);
		maxRow = Math.max(maxRow, cell.row);
		minCol = Math.min(minCol, cell.col);
		maxCol = Math.max(maxCol, cell.col);
		if (cell.richText !== void 0) {
			richTextCellCount++;
			if (richTextCells.length < MAX_RICH_TEXT_CELLS) richTextCells.push(cell.ref);
		}
		if (cell.formula !== null) {
			formulaCount++;
			if (formulas.length < options.formulaSamples) formulas.push({
				cell: cell.ref,
				formula: cell.formula.text
			});
		}
	}
	const empty = maxRow === 0;
	const usedRange = empty ? null : formatRange({
		start: {
			row: minRow,
			col: minCol
		},
		end: {
			row: maxRow,
			col: maxCol
		}
	});
	checkDimension(sheet, empty ? null : {
		minRow,
		maxRow,
		minCol,
		maxCol
	}, warnings);
	checkMergedValues(sheet, rows, warnings);
	const base = {
		name: sheet.name,
		index: sheet.index,
		sheetId: sheet.sheetId,
		kind: sheet.kind,
		state: sheet.state,
		dimension: sheet.dimension,
		usedRange,
		cellCount: sheet.cellCount,
		formulaCount,
		mergedRanges: sheet.merges.slice(0, MAX_MERGED_RANGES),
		mergedRangeCount: sheet.merges.length,
		mergedRangesTruncated: sheet.merges.length > MAX_MERGED_RANGES,
		formulas,
		richTextCellCount,
		richTextCells,
		styledColumns: styledColumns(sheet, options.maxColumns),
		hyperlinkCount: sheet.hyperlinks.length,
		hyperlinks: sheet.hyperlinks.slice(0, MAX_HYPERLINK_SAMPLES).map((l) => ({
			ref: l.ref,
			target: l.target,
			location: l.location
		})),
		noteCount: sheet.notes.length,
		notes: sheet.notes.slice(0, MAX_HYPERLINK_SAMPLES).map((n) => n.ref),
		imageCount: sheet.images.length,
		images: sheet.images.slice(0, MAX_HYPERLINK_SAMPLES).map((i) => ({
			cell: i.anchor.type === "absolute" ? null : i.anchor.from.cell,
			format: i.image?.format ?? null
		})),
		chartCount: sheet.charts.length,
		charts: sheet.charts.slice(0, MAX_HYPERLINK_SAMPLES).map((c) => ({
			cell: c.anchor.type === "absolute" ? null : c.anchor.from.cell,
			types: c.types.slice(0, MAX_CHART_TYPES),
			typesTruncated: c.types.length > MAX_CHART_TYPES,
			title: c.title === null ? null : "text" in c.title ? { text: cut(c.title.text) } : { ref: cut(c.title.ref) },
			seriesCount: c.series.length
		})),
		protected: sheet.protection?.sheet === true,
		printArea: sheet.printArea === null ? null : sheet.printArea.slice(0, MAX_HYPERLINK_SAMPLES),
		pageBreakCount: sheet.pageBreaks.rows.length + sheet.pageBreaks.columns.length,
		dataValidationCount: sheet.dataValidations.length,
		conditionalFormatCount: sheet.conditionalFormats.reduce((n, f) => n + f.rules.length, 0),
		tables: sheet.tables.slice(0, 20).map((t) => ({
			name: t.displayName,
			ref: t.ref,
			columns: t.columns.slice(0, 50).map((c) => c.name)
		}))
	};
	if (empty) return {
		...base,
		rowCount: 0,
		columnCount: 0,
		headerRow: null,
		columns: [],
		columnsTruncated: false
	};
	const width = maxCol - minCol + 1;
	const headerRow = detectHeader(rows, width, options.headerScanRows);
	const headerCells = headerRow === null ? void 0 : rows.find((r) => r[0]?.row === headerRow.row);
	const firstDataRow = headerRow === null ? minRow : headerRow.row + 1;
	const dataRowCount = Math.max(0, maxRow - firstDataRow + 1);
	const shownColumns = Math.min(width, options.maxColumns);
	const stats = Array.from({ length: shownColumns }, () => ({
		counts: {
			number: 0,
			string: 0,
			boolean: 0,
			date: 0,
			error: 0
		},
		samples: []
	}));
	for (const cells of rows) {
		if (cells.length === 0 || (cells[0]?.row ?? 0) < firstDataRow) continue;
		for (const cell of cells) {
			const stat = stats[cell.col - minCol];
			if (stat === void 0 || cell.type === "empty" || cell.value === null) continue;
			stat.counts[cell.type]++;
			const sample = jsonValue(cell.value);
			if (stat.samples.length < 3 && !stat.samples.includes(sample)) stat.samples.push(sample);
		}
	}
	const columns = stats.map((stat, i) => {
		const col = minCol + i;
		const headerCell = headerCells?.find((cell) => cell.col === col);
		const nonEmpty = TYPE_ORDER.reduce((sum, t) => sum + stat.counts[t], 0);
		let dominantType = "empty";
		let best = 0;
		for (const t of TYPE_ORDER) if (stat.counts[t] > best) {
			best = stat.counts[t];
			dominantType = t;
		}
		return {
			index: col,
			letter: columnLetters(col),
			header: headerCell === void 0 || headerCell.formula !== null ? null : typeof headerCell.value === "string" ? headerCell.value : isYear(headerCell.value) ? String(headerCell.value) : null,
			dominantType,
			types: {
				...stat.counts,
				empty: dataRowCount - nonEmpty
			},
			samples: stat.samples
		};
	});
	return {
		...base,
		rowCount: maxRow - minRow + 1,
		columnCount: width,
		headerRow,
		columns,
		columnsTruncated: width > shownColumns
	};
}
/**
* Header heuristic (docs/decisions.md D-R02): the first of the leading non-empty rows in which
* - cells with a formula are left out (a total such as `=SUM(B3:B9)` may sit next to the names);
* - every other cell is non-blank text or a year: a whole number from 1900 to 2100, with the
*   years of the row consecutive (`2023 | 2024 | 2025`) and, in the next non-empty row, at least
*   one value below them that is not such a year (so a row of data with years is not taken);
* - at least one cell is text, and the cells cover at least half of the used width.
* Repeated names are allowed (`readTable` numbers them `name_2`, …). Confidence is high when the
* names are distinct and the next non-empty row holds a value that is not text.
*/
function detectHeader(rows, width, scanRows) {
	const nonEmpty = rows.filter((r) => r.length > 0);
	const candidates = nonEmpty.slice(0, scanRows);
	for (let i = 0; i < candidates.length; i++) {
		const cells = candidates[i] ?? [];
		if (cells.length < Math.max(1, Math.ceil(width / 2))) continue;
		const labels = cells.filter((cell) => cell.formula === null);
		const years = labels.filter((cell) => isYear(cell.value));
		const texts = labels.filter((cell) => cell.type === "string" && typeof cell.value === "string" && cell.value.trim() !== "");
		if (texts.length === 0 || texts.length + years.length !== labels.length) continue;
		const next = nonEmpty[i + 1];
		if (years.length > 0) {
			const consecutive = years.every((cell, k) => k === 0 || Math.abs(cell.value - years[k - 1].value) === 1);
			const yearCols = new Set(years.map((y) => y.col));
			const below = next?.filter((cell) => yearCols.has(cell.col)) ?? [];
			if (!consecutive || !below.some((cell) => cell.value !== null && !isYear(cell.value))) continue;
		}
		const distinct = new Set(texts.map((cell) => cell.value.trim().toLowerCase())).size === texts.length;
		const typed = next !== void 0 && next.some((cell) => cell.type !== "string" && cell.type !== "empty");
		return {
			row: cells[0]?.row ?? 0,
			confidence: distinct && typed ? "high" : "low"
		};
	}
	return null;
}
/** A whole number from 1900 to 2100, read as a year in a header row. */
function isYear(value) {
	return typeof value === "number" && Number.isInteger(value) && value >= 1900 && value <= 2100;
}
/** Warns when cells lie outside the declared dimension (a dimension larger than the cells is normal). */
function checkDimension(sheet, used, warnings) {
	if (sheet.dimension === null || used === null) return;
	const declared = parseRange(sheet.dimension);
	if (declared === null) return;
	if (!(used.minRow >= declared.start.row && used.maxRow <= declared.end.row && used.minCol >= declared.start.col && used.maxCol <= declared.end.col)) warnings.add("XLV_W_DIMENSION_MISMATCH", `The sheet declares the range ${sheet.dimension}, but it has cells outside it; programs that trust the declared range may miss data.`, { sheet: sheet.name });
}
/** Warns about values in merged ranges other than the top-left cell; applications do not show them. */
function checkMergedValues(sheet, rows, warnings) {
	const merges = [...new Set(sheet.merges)].map((m) => parseRange(m)).filter((r) => r !== null).sort((a, b) => a.start.row - b.start.row);
	if (merges.length === 0) return;
	let work = 0;
	let next = 0;
	let active = [];
	for (const cells of rows) {
		const r = cells[0]?.row;
		if (r === void 0) continue;
		while (next < merges.length && (merges[next]?.start.row ?? Number.POSITIVE_INFINITY) <= r) {
			const m = merges[next++];
			if (m !== void 0) active.push(m);
		}
		work += active.length;
		active = active.filter((m) => m.end.row >= r);
		if (active.length === 0) continue;
		for (const cell of cells) {
			if (cell.value === null && cell.formula === null) continue;
			for (const m of active) {
				work++;
				if (!(cell.row === m.start.row && cell.col === m.start.col) && cell.col >= m.start.col && cell.col <= m.end.col) {
					warnings.add("XLV_W_VALUES_IN_MERGED_CELLS", "Some merged ranges have values in cells other than the top-left one; spreadsheet applications do not display them.", {
						sheet: sheet.name,
						cell: cell.ref
					});
					break;
				}
			}
		}
		if (work > MAX_MERGE_CHECK_WORK) return;
	}
}
// src/util/range-index.ts
/** Block heights of the row levels: 4, 16, …, 4^10 = 1 048 576 rows (the whole sheet). */
const ROW_SIZES = [
	4,
	16,
	64,
	256,
	1024,
	4096,
	16384,
	65536,
	262144,
	1048576
];
/** Block widths of the column levels: 4, 16, …, 4^7 = 16 384 columns (the whole sheet). */
const COL_SIZES = [
	4,
	16,
	64,
	256,
	1024,
	4096,
	16384
];
function levelOf(sizes, length) {
	for (let i = 0; i < sizes.length; i++) if (length <= sizes[i]) return i;
	return sizes.length - 1;
}
function blockOf(index, size) {
	return Math.floor((index - 1) / size);
}
function overlap(a, b) {
	return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}
/**
* A grid of blocks at several sizes. A range is stored in the level whose block height is the
* smallest at least as tall as the range and whose block width is the smallest at least as wide,
* in every block of that level it touches (at most 2 × 2), so a huge range costs no more than a
* small one. A query visits, in each level that holds ranges, the blocks it touches that hold
* ranges (the smaller of the blocks it spans and the blocks the level holds). Ranges of a level
* are at least about a quarter of its block size in each dimension, so when the stored ranges do
* not overlap each other (merged ranges, tables), each block holds a bounded number of them.
*/
var RangeIndex = class {
	/** Level (row level × 16 + column level) → column block → row block → ranges. */
	levels = /* @__PURE__ */ new Map();
	/** Adds a range. */
	add(range) {
		const a = levelOf(ROW_SIZES, range.end.row - range.start.row + 1);
		const b = levelOf(COL_SIZES, range.end.col - range.start.col + 1);
		const rows = ROW_SIZES[a];
		const cols = COL_SIZES[b];
		let level = this.levels.get(a * 16 + b);
		if (level === void 0) {
			level = /* @__PURE__ */ new Map();
			this.levels.set(a * 16 + b, level);
		}
		for (let cb = blockOf(range.start.col, cols); cb <= blockOf(range.end.col, cols); cb++) {
			let column = level.get(cb);
			if (column === void 0) {
				column = /* @__PURE__ */ new Map();
				level.set(cb, column);
			}
			for (let rb = blockOf(range.start.row, rows); rb <= blockOf(range.end.row, rows); rb++) {
				const list = column.get(rb);
				if (list === void 0) column.set(rb, [range]);
				else list.push(range);
			}
		}
	}
	/** Whether `range` shares a cell with a range added before. */
	overlaps(range) {
		return this.find(range) !== void 0;
	}
	/**
	* A range added before that shares a cell with `range` and that `accept` (when given) accepts,
	* or `undefined`. Which one is returned when several qualify is not defined; ranges that do not
	* overlap each other (merged ranges, tables, array formulas of a valid file) qualify at most
	* once for a single cell.
	*/
	find(range, accept) {
		const pick = (list) => {
			if (list === void 0) return void 0;
			for (const r of list) if (overlap(r, range) && (accept === void 0 || accept(r))) return r;
		};
		for (const [key, level] of this.levels) {
			const rows = ROW_SIZES[key >> 4];
			const cols = COL_SIZES[key & 15];
			const rb1 = blockOf(range.start.row, rows);
			const rb2 = blockOf(range.end.row, rows);
			const cb1 = blockOf(range.start.col, cols);
			const cb2 = blockOf(range.end.col, cols);
			const inColumn = (column) => {
				if (rb2 - rb1 + 1 <= column.size) {
					for (let rb = rb1; rb <= rb2; rb++) {
						const found = pick(column.get(rb));
						if (found !== void 0) return found;
					}
					return;
				}
				for (const [rb, list] of column) {
					if (rb < rb1 || rb > rb2) continue;
					const found = pick(list);
					if (found !== void 0) return found;
				}
			};
			if (cb2 - cb1 + 1 <= level.size) for (let cb = cb1; cb <= cb2; cb++) {
				const column = level.get(cb);
				const found = column === void 0 ? void 0 : inColumn(column);
				if (found !== void 0) return found;
			}
			else for (const [cb, column] of level) {
				if (cb < cb1 || cb > cb2) continue;
				const found = inColumn(column);
				if (found !== void 0) return found;
			}
		}
	}
};
// src/write.ts
/**
* Error values accepted in cells (ECMA-376 Part 1 §18.17.3, fact E-33).
*
* @example
* ```ts
* import { ERROR_VALUES } from 'xlvane';
*
* console.log(ERROR_VALUES.includes('#DIV/0!')); // true
* ```
*/
const ERROR_VALUES = [
	"#NULL!",
	"#DIV/0!",
	"#VALUE!",
	"#REF!",
	"#NAME?",
	"#NUM!",
	"#N/A",
	"#GETTING_DATA"
];
/** The view of a new sheet: the defaults of CT_SheetView (W2-51). */
const DEFAULT_VIEW = {
	zoom: 100,
	showGridLines: true,
	showRowColHeaders: true,
	rightToLeft: false,
	layout: "normal",
	split: null,
	topLeftCell: null,
	activeCell: null,
	zoomScaleNormal: 0,
	showRuler: true
};
/** Reads `outlineLevel` (0–7, [MS-OI29500] §2.1.598 e, §2.1.644 b) and `collapsed` of a row or column. */
function readOutline(o, current, what) {
	const level = o.get("outlineLevel");
	if (level !== void 0) {
		if (typeof level !== "number" || !Number.isInteger(level) || level < 0 || level > 7) throw invalidOption(`${what} option "outlineLevel" must be a whole number from 0 to 7.`, "For example { outlineLevel: 1 }.");
		current.outlineLevel = level;
	}
	const collapsed = o.get("collapsed");
	if (collapsed !== void 0) {
		if (typeof collapsed !== "boolean") throw invalidOption(`${what} option "collapsed" must be true or false.`, "Use { collapsed: true }.");
		current.collapsed = collapsed;
	}
}
/** Largest split position accepted, in points: an implementation bound, not a documented limit. */
const MAX_SPLIT_POINTS = 1e6;
/** Office uses sheet ids from 1 to 65534 ([MS-OI29500] §2.1.578; W2-20). */
const MAX_SHEET_ID = 65534;
/** Most characters a cell can contain (Excel specifications and limits, W-04). */
const MAX_CELL_TEXT = 32767;
/** Widest column, in characters (Excel specifications and limits, W-04). */
/** Widest column Office accepts ([MS-OI29500] §2.1.598 c). */
const MAX_COLUMN_WIDTH = 255;
/**
* Tallest row, in points. The Excel specifications and limits page says 409 (W-04); a file saved
* by Microsoft Excel in the POI corpus has rows of 409.5 points, so 409.5 is accepted (W-04a,
* docs/decisions.md D-R01).
*/
const MAX_ROW_HEIGHT = 409.5;
/** Most cells one call to {@link NewSheet.style} may touch. Implementation cap (D-206). */
const MAX_STYLE_RANGE_CELLS = 1048576;
/** Longest sheet name ([MS-OI29500] §2.1.578 a; docs/facts/sml-write.md W2-20). */
const MAX_SHEET_NAME = 31;
/** Characters not allowed in sheet names (ECMA-376 Part 1 §18.17 sheet-name-special; [MS-OI29500] §2.1.578 a, W2-20). */
const SHEET_NAME_FORBIDDEN = /[\\/?*:[\]]/;
/** Formula text limit (Excel specifications and limits: 8,192 characters; to be verified, W2-21). */
const MAX_FORMULA_LENGTH = 8192;
const DATE_FORMAT = "yyyy-mm-dd";
const DATE_TIME_FORMAT = "yyyy-mm-dd hh:mm:ss";
const DAY_MS = 864e5;
/** Most runs in one rich text value ([MS-OI29500] §2.1.675, F-SST-09). */
const MAX_RICH_TEXT_RUNS = 32767;
function invalidValue(sheet, cell, message, hint) {
	return new XlvError("XLV_E_INVALID_VALUE", message, {
		location: {
			sheet,
			cell
		},
		hint
	});
}
/**
* Runs `fn` and adds `location` to an `XlvError` it throws without a sheet, so that errors of
* shared validators (styles, colours) say where the bad value was given (D-R03).
*/
function located(location, fn) {
	try {
		return fn();
	} catch (error) {
		if (error instanceof XlvError && error.location.sheet === void 0) throw new XlvError(error.code, error.message, {
			location: {
				...location,
				...error.location
			},
			hint: error.hint,
			details: error.details,
			cause: error
		});
		throw error;
	}
}
function invalidOption(message, hint, details) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		hint,
		...details ? { details } : {}
	});
}
/** Reads the own enumerable properties of an options object, rejecting unknown ones. */
function readOptions$1(value, known, what) {
	const out = /* @__PURE__ */ new Map();
	if (value === void 0) return out;
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw invalidOption(`${what} must be an object.`, `Pass an object with some of these properties: ${known.join(", ")}.`);
	for (const key of Object.keys(value)) {
		if (!known.includes(key)) {
			const close = suggest(key, known);
			throw invalidOption(`Unknown option "${key.slice(0, 40)}" in ${what}.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known options: ${known.join(", ")}.`, {
				option: key.slice(0, 40),
				known
			});
		}
		const v = value[key];
		if (v !== void 0) out.set(key, v);
	}
	return out;
}
function invalidRef(ref, sheet, what = "cell reference", example = "\"B7\"") {
	const shown = String(typeof ref === "symbol" ? "symbol" : ref).slice(0, 40);
	return new XlvError("XLV_E_INVALID_CELL_REF", `"${shown}" is not a valid ${what}.`, {
		location: { sheet },
		hint: `Pass an A1-style ${what} such as ${example} (column letters A to XFD, rows 1 to 1048576).`,
		details: { ref: shown }
	});
}
/** Validates a sheet name; throws `XLV_E_INVALID_SHEET_NAME` with the reason. */
function checkSheetName(name) {
	const fail = (reason) => {
		throw new XlvError("XLV_E_INVALID_SHEET_NAME", `Invalid sheet name ${JSON.stringify(String(name).slice(0, 40))}: ${reason}`, {
			hint: `Use 1 to ${String(MAX_SHEET_NAME)} characters without \\ / ? * : [ ] and without an apostrophe at the start or end.`,
			details: { name: String(name).slice(0, 40) }
		});
	};
	if (typeof name !== "string") return fail("it must be a string.");
	if (name.length === 0) return fail("it is empty.");
	if (name.length > MAX_SHEET_NAME) return fail(`it has ${String(name.length)} characters, more than ${String(MAX_SHEET_NAME)}.`);
	const bad = SHEET_NAME_FORBIDDEN.exec(name);
	if (bad) return fail(`it contains "${bad[0]}".`);
	if (name.startsWith("'") || name.endsWith("'")) return fail("it starts or ends with an apostrophe.");
	if (firstLoneSurrogate(name) !== -1) return fail("it contains an unpaired UTF-16 surrogate.");
	if (/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/.test(name)) return fail("it contains a control character.");
	return name;
}
/**
* A sheet of a workbook created with {@link createWorkbook}. Methods return the sheet, so calls
* can be chained.
*
* @example
* ```ts
* import { createWorkbook } from 'xlvane';
*
* const sheet = createWorkbook().addSheet('Sales');
* sheet.set('A1', 'Product', { font: { bold: true } }).set('B1', 'Amount', { font: { bold: true } });
* sheet.set('A2', 'Tea').set('B2', 12.5, { numFmt: '#,##0.00' });
* sheet.column('A', { width: 20 }).freeze({ rows: 1 }).autoFilter('A1:B2');
* ```
*/
var NewSheet = class {
	name;
	book;
	/** Visibility of the sheet. */
	state;
	/** The id given to {@link NewWorkbook.addSheet}, or `null` when the lowest unused id is taken when writing. */
	sheetId;
	/** @internal */
	cellMap = /* @__PURE__ */ new Map();
	/** @internal */
	columnMap = /* @__PURE__ */ new Map();
	/** @internal */
	rowMap = /* @__PURE__ */ new Map();
	/** @internal Sheet-wide formatting (W2-65). */
	sheetProps = {
		defaultRowHeight: null,
		summaryBelow: true,
		summaryRight: true,
		outlineLevelRow: 0,
		outlineLevelCol: 0
	};
	/** @internal */
	mergeList = [];
	/** @internal */
	frozen = null;
	/** @internal */
	filter = null;
	/** @internal */
	tab = null;
	/** @internal */
	viewSettings = DEFAULT_VIEW;
	/** @internal Hyperlinks by normalised range, in the order they were added. */
	linkMap = /* @__PURE__ */ new Map();
	/** @internal Notes by cell reference. */
	noteMap = /* @__PURE__ */ new Map();
	/** @internal Pictures in the order they were added (TD-35). */
	imageList = [];
	/** @internal Background image (index of `NewWorkbook.addImage`), or `null`. */
	backgroundImage = null;
	/** @internal Charts in the order they were added (D-K01). */
	chartList = [];
	/** @internal Data validations in the order they were added (SR-01). */
	validationList = [];
	/** @internal Conditional formats in the order they were added (SR-07). */
	conditionalList = [];
	/** Priorities of the conditional formatting rules of the sheet (unique, SR-10). */
	priorities = /* @__PURE__ */ new Set();
	/** @internal Protection, page setup, headers and footers, page breaks, print area and titles. */
	print = EMPTY_PRINT_DATA;
	/** @internal Tables in the order they were added (TD-01); add them with {@link NewSheet.pushTable}. */
	tableList = [];
	count = 0;
	mergeIndex = new RangeIndex();
	mergesIndexed = 0;
	tableIndex = new RangeIndex();
	tablesIndexed = 0;
	/** @internal Use {@link NewWorkbook.addSheet}. */
	constructor(name, state, book, sheetId = null) {
		this.name = name;
		this.book = book;
		this.state = state;
		this.sheetId = sheetId;
	}
	/** Number of cells that have a value or a style. */
	get cellCount() {
		return this.count;
	}
	address(ref) {
		const address = typeof ref === "string" ? parseCellRef(ref) : null;
		if (address === null) throw invalidRef(ref, this.name);
		return {
			...address,
			ref: formatCellRef(address.row, address.col)
		};
	}
	range(ref, what = "range", example = "\"A1:C3\"") {
		const range = typeof ref === "string" ? parseRange(ref) : null;
		if (range === null) throw invalidRef(ref, this.name, what, example);
		return range;
	}
	getCell(row, col) {
		return this.cellMap.get(row)?.get(col);
	}
	putCell(row, col, cell) {
		let cells = this.cellMap.get(row);
		if (cell === null) {
			if (cells?.delete(col) === true) {
				this.count--;
				if (cells.size === 0) this.cellMap.delete(row);
			}
			return;
		}
		if (!cells) {
			cells = /* @__PURE__ */ new Map();
			this.cellMap.set(row, cells);
		}
		if (!cells.has(col)) this.count++;
		cells.set(col, cell);
	}
	/**
	* Sets the value of a cell, and its style when `style` is given (otherwise the cell keeps its
	* current style). Dates get a date number format (`yyyy-mm-dd`, or `yyyy-mm-dd hh:mm:ss` when
	* they have a time) unless the style sets one. Setting `null` without a style removes an
	* unstyled cell.
	*
	* Throws `XLV_E_INVALID_CELL_REF` for a bad reference, `XLV_E_INVALID_VALUE` for values a
	* spreadsheet cannot hold (NaN, infinities, text over 32 767 characters, dates outside the
	* date system) and `XLV_E_INVALID_STYLE` for a bad style.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Data');
	* sheet.set('A1', 'Total').set('B1', { formula: 'SUM(B2:B10)', result: 0 });
	* sheet.set('C1', new Date(Date.UTC(2024, 0, 31)), { font: { italic: true } });
	* sheet.set('D1', { error: '#N/A' });
	* ```
	*/
	set(ref, value, style) {
		const at = this.address(ref);
		const normalized = style === void 0 ? null : located({
			sheet: this.name,
			cell: at.ref
		}, () => normalizeStyle(style));
		const { stored, isDate, hasTime } = this.convert(value, at.ref);
		let base = normalized ?? this.book.styleOf(this.getCell(at.row, at.col)?.style ?? 0);
		if (isDate && base.numFmt === void 0) base = {
			...base,
			numFmt: hasTime ? DATE_TIME_FORMAT : DATE_FORMAT
		};
		const styleId = this.book.intern(base);
		this.putCell(at.row, at.col, stored === null && styleId === 0 ? null : {
			value: stored,
			style: styleId
		});
		return this;
	}
	convert(value, cell) {
		return convertCellInput(value, {
			sheet: this.name,
			dateSystem: this.book.dateSystem
		}, cell);
	}
	/**
	* Applies a style to every cell of a range, merging it with each cell's current style: given
	* font, border and alignment properties replace the cell's ones, a given fill or number format
	* replaces the cell's. Cells without a value are created as empty styled cells. Throws
	* `XLV_E_LIMIT_EXCEEDED` for ranges of more than 1 048 576 cells.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Data').set('A1', 'Name').set('B1', 'Amount');
	* sheet.style('A1:B1', { font: { bold: true }, fill: { color: 'D9E1F2' } });
	* ```
	*/
	style(range, style) {
		const r = this.range(range);
		const patch = located({
			sheet: this.name,
			cell: formatRange(r)
		}, () => normalizeStyle(style));
		const cells = (r.end.row - r.start.row + 1) * (r.end.col - r.start.col + 1);
		if (cells > 1048576) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The range ${range} has ${String(cells)} cells, more than ${String(MAX_STYLE_RANGE_CELLS)} can be styled at once.`, {
			location: { sheet: this.name },
			hint: "Style only the cells that hold data, for example the header row and the used range.",
			details: {
				limit: "styleRangeCells",
				value: MAX_STYLE_RANGE_CELLS
			}
		});
		const cache = /* @__PURE__ */ new Map();
		for (let row = r.start.row; row <= r.end.row; row++) for (let col = r.start.col; col <= r.end.col; col++) {
			const cell = this.getCell(row, col);
			const old = cell?.style ?? 0;
			let id = cache.get(old);
			if (id === void 0) {
				id = this.book.intern(mergeStyles(this.book.styleOf(old), patch));
				cache.set(old, id);
			}
			if (cell) cell.style = id;
			else if (id !== 0) this.putCell(row, col, {
				value: null,
				style: id
			});
		}
		return this;
	}
	/**
	* Sets the width, visibility or default style of a column, given by letters (`"C"`) or number
	* (3). Options not given keep their current value.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').column('A', { width: 30 }).column(2, { hidden: true, style: { font: { bold: true } } });
	* ```
	*/
	column(col, options) {
		return located({ sheet: this.name }, () => {
			const index = typeof col === "number" ? col : typeof col === "string" ? columnNumber(col.toUpperCase()) ?? 0 : 0;
			if (!Number.isInteger(index) || index < 1 || index > 16384) throw invalidRef(col, this.name, "column", "\"C\" or 3");
			const o = readOptions$1(options, [
				"width",
				"hidden",
				"style",
				"outlineLevel",
				"collapsed"
			], "column options");
			const current = this.columnMap.get(index) ?? {
				width: null,
				hidden: false,
				style: 0,
				outlineLevel: 0,
				collapsed: false
			};
			const width = o.get("width");
			if (width !== void 0) {
				if (typeof width !== "number" || !(width >= 0 && width <= MAX_COLUMN_WIDTH)) throw invalidValue(this.name, `${columnLetters(index)}:${columnLetters(index)}`, `Column width ${String(width)} of column ${columnLetters(index)} is not valid.`, `Use a width from 0 to ${String(MAX_COLUMN_WIDTH)} characters.`);
				current.width = width;
			}
			const hidden = o.get("hidden");
			if (hidden !== void 0) {
				if (typeof hidden !== "boolean") throw invalidOption("Column option \"hidden\" must be true or false.", "Use { hidden: true }.");
				current.hidden = hidden;
			}
			const style = o.get("style");
			if (style !== void 0) current.style = this.book.intern(located({
				sheet: this.name,
				cell: `${columnLetters(index)}:${columnLetters(index)}`
			}, () => normalizeStyle(style)));
			readOutline(o, current, "Column");
			this.columnMap.set(index, current);
			return this;
		});
	}
	/**
	* Sets the height, visibility, outline or default style of a row (1-based).
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').row(1, { height: 30, style: { font: { bold: true } } }).row(5, { hidden: true });
	* ```
	*/
	row(row, options) {
		return located({ sheet: this.name }, () => {
			if (typeof row !== "number" || !Number.isInteger(row) || row < 1 || row > 1048576) throw invalidRef(row, this.name, "row number", "1");
			const o = readOptions$1(options, [
				"height",
				"hidden",
				"style",
				"outlineLevel",
				"collapsed"
			], "row options");
			const current = this.rowMap.get(row) ?? {
				height: null,
				hidden: false,
				style: 0,
				outlineLevel: 0,
				collapsed: false
			};
			const height = o.get("height");
			if (height !== void 0) {
				if (typeof height !== "number" || !(height >= 0 && height <= MAX_ROW_HEIGHT)) throw invalidValue(this.name, `${String(row)}:${String(row)}`, `Row height ${String(height)} of row ${String(row)} is not valid.`, `Use a height from 0 to ${String(MAX_ROW_HEIGHT)} points.`);
				current.height = height;
			}
			const hidden = o.get("hidden");
			if (hidden !== void 0) {
				if (typeof hidden !== "boolean") throw invalidOption("Row option \"hidden\" must be true or false.", "Use { hidden: true }.");
				current.hidden = hidden;
			}
			const style = o.get("style");
			if (style !== void 0) current.style = this.book.intern(located({
				sheet: this.name,
				cell: `${String(row)}:${String(row)}`
			}, () => normalizeStyle(style)));
			readOutline(o, current, "Row");
			this.rowMap.set(row, current);
			return this;
		});
	}
	/**
	* Sets sheet-wide formatting: the default row height and where outline groups put their
	* summary rows and columns (W2-65).
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').properties({ defaultRowHeight: 18, summaryBelow: false });
	* ```
	*/
	properties(options) {
		return located({ sheet: this.name }, () => {
			const o = readOptions$1(options, [
				"defaultRowHeight",
				"summaryBelow",
				"summaryRight",
				"outlineLevelRow",
				"outlineLevelCol"
			], "sheet properties");
			const next = { ...this.sheetProps };
			const h = o.get("defaultRowHeight");
			if (h !== void 0) {
				if (typeof h !== "number" || !(h >= 0 && h <= MAX_ROW_HEIGHT)) throw invalidOption(`The default row height ${String(h)} is not valid.`, `Use a height from 0 to ${String(MAX_ROW_HEIGHT)} points, for example { defaultRowHeight: 18 }.`);
				next.defaultRowHeight = h;
			}
			for (const key of ["summaryBelow", "summaryRight"]) {
				const v = o.get(key);
				if (v === void 0) continue;
				if (typeof v !== "boolean") throw invalidOption(`The sheet property "${key}" must be true or false.`, `For example { ${key}: false }.`);
				next[key] = v;
			}
			for (const key of ["outlineLevelRow", "outlineLevelCol"]) {
				const v = o.get(key);
				if (v === void 0) continue;
				if (typeof v !== "number" || !Number.isInteger(v) || v < 0 || v > 7) throw invalidOption(`The sheet property "${key}" must be a whole number from 0 to 7.`, `For example { ${key}: 2 }.`);
				next[key] = v;
			}
			this.sheetProps = next;
			return this;
		});
	}
	/**
	* Merges a range of at least two cells. Throws `XLV_E_OVERLAPPING_MERGE` if it overlaps a
	* range merged before.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').set('A1', 'Quarterly report').merge('A1:D1');
	* ```
	*/
	merge(range) {
		const r = this.range(range);
		if (r.start.row === r.end.row && r.start.col === r.end.col) throw invalidRef(range, this.name, "range of two or more cells", "\"A1:C1\"");
		if (this.overlapsMerge(r)) {
			for (const m of this.mergeList) if (r.start.row <= m.end.row && m.start.row <= r.end.row && r.start.col <= m.end.col && m.start.col <= r.end.col) {
				const other = `${formatCellRef(m.start.row, m.start.col)}:${formatCellRef(m.end.row, m.end.col)}`;
				throw new XlvError("XLV_E_OVERLAPPING_MERGE", `The range ${range} overlaps the merged range ${other}.`, {
					location: { sheet: this.name },
					hint: "Merged ranges must not overlap. Choose ranges that share no cells.",
					details: {
						range: String(range),
						other
					}
				});
			}
		}
		if (this.overlapsTable(r)) {
			for (const t of this.tableList) if (r.start.row <= t.range.end.row && t.range.start.row <= r.end.row && r.start.col <= t.range.end.col && t.range.start.col <= r.end.col) throw new XlvError("XLV_E_INVALID_TABLE", `The range ${range} overlaps table "${t.name}".`, {
				location: { sheet: this.name },
				hint: "Tables cannot contain merged cells; merge cells outside the table.",
				details: { name: t.name }
			});
		}
		this.mergeList.push(r);
		return this;
	}
	/** @internal Whether `range` shares a cell with a merged range of the sheet. */
	overlapsMerge(range) {
		while (this.mergesIndexed < this.mergeList.length) this.mergeIndex.add(this.mergeList[this.mergesIndexed++]);
		return this.mergeIndex.overlaps(range);
	}
	/** @internal Whether `range` shares a cell with a table of the sheet. */
	overlapsTable(range) {
		while (this.tablesIndexed < this.tableList.length) this.tableIndex.add(this.tableList[this.tablesIndexed++].range);
		return this.tableIndex.overlaps(range);
	}
	/** @internal Adds a table checked by {@link NewWorkbook.checkTable} to the sheet. */
	pushTable(t) {
		this.tableList.push(t);
		this.book.tableAdded(t);
	}
	/**
	* Freezes the top rows and left columns so that they stay visible when scrolling. Pass
	* `{ rows: 0, cols: 0 }` to remove the frozen panes.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').freeze({ rows: 1 }); // keep the header row visible
	* ```
	*/
	freeze(options) {
		return located({ sheet: this.name }, () => {
			const o = readOptions$1(options, ["rows", "cols"], "freeze options");
			const count = (key, max) => {
				const v = o.get(key) ?? 0;
				if (typeof v !== "number" || !Number.isInteger(v) || v < 0 || v >= max) throw invalidOption(`freeze option "${key}" must be a whole number from 0 to ${String(max - 1)}.`, `For example { ${key}: 1 }.`);
				return v;
			};
			const rows = count("rows", MAX_ROWS);
			const cols = count("cols", MAX_COLUMNS);
			this.frozen = rows === 0 && cols === 0 ? null : {
				rows,
				cols
			};
			if (this.frozen !== null && this.viewSettings.split !== null) this.viewSettings = {
				...this.viewSettings,
				split: null
			};
			return this;
		});
	}
	/**
	* Sets the colour of the sheet tab (`sheetPr/tabColor`, W2-50), or removes it with `null`.
	* Colours are given as for cell styles.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Summary').tabColor('00B050');
	* ```
	*/
	tabColor(color) {
		this.tab = color === null ? null : located({ sheet: this.name }, () => normalizeColor(color, "tabColor"));
		return this;
	}
	/**
	* Sets how the sheet is displayed when opened: zoom, gridlines, headings, direction, layout
	* and a split window (W2-51). Settings not given keep their value.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').view({ zoom: 120, showGridLines: false });
	* ```
	*/
	view(options) {
		return located({ sheet: this.name }, () => {
			const o = readOptions$1(options, [
				"zoom",
				"showGridLines",
				"showRowColHeaders",
				"rightToLeft",
				"layout",
				"split",
				"topLeftCell",
				"activeCell",
				"zoomScaleNormal",
				"showRuler"
			], "view options");
			const next = { ...this.viewSettings };
			const zoom = o.get("zoom");
			if (zoom !== void 0) {
				if (typeof zoom !== "number" || !Number.isInteger(zoom) || zoom < 10 || zoom > 400) throw invalidOption("The view option \"zoom\" must be a whole number from 10 to 400 (percent).", "For example { zoom: 80 }.");
				next.zoom = zoom;
			}
			const zoomNormal = o.get("zoomScaleNormal");
			if (zoomNormal !== void 0) {
				if (typeof zoomNormal !== "number" || !Number.isInteger(zoomNormal) || zoomNormal !== 0 && (zoomNormal < 10 || zoomNormal > 400)) throw invalidOption("The view option \"zoomScaleNormal\" must be 0 (automatic) or a whole number from 10 to 400.", "For example { zoomScaleNormal: 100 }.");
				next.zoomScaleNormal = zoomNormal;
			}
			for (const key of ["topLeftCell", "activeCell"]) {
				if (!o.has(key)) continue;
				const v = o.get(key);
				if (v === null) {
					next[key] = null;
					continue;
				}
				const cell = typeof v === "string" ? parseCellRef(v.toUpperCase()) : null;
				if (cell === null) throw invalidOption(`The view option "${key}" must be a cell reference such as "B2".`, `For example { ${key}: "B2" }.`);
				next[key] = formatCellRef(cell.row, cell.col);
			}
			for (const key of [
				"showGridLines",
				"showRowColHeaders",
				"rightToLeft",
				"showRuler"
			]) {
				const v = o.get(key);
				if (v === void 0) continue;
				if (typeof v !== "boolean") throw invalidOption(`The view option "${key}" must be true or false.`, `For example { ${key}: false }.`);
				next[key] = v;
			}
			const layout = o.get("layout");
			if (layout !== void 0) {
				if (layout !== "normal" && layout !== "pageBreakPreview" && layout !== "pageLayout") throw invalidOption(`The view option "layout" is ${JSON.stringify(String(layout).slice(0, 30))}.`, "Use \"normal\", \"pageBreakPreview\" or \"pageLayout\".");
				next.layout = layout;
			}
			let unfreeze = false;
			if (o.has("split")) {
				const split = o.get("split");
				if (split === null) next.split = null;
				else {
					const p = readOptions$1(split, [
						"x",
						"y",
						"topLeftCell"
					], "view option \"split\"");
					const at = (key) => {
						const v = p.get(key) ?? 0;
						if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > MAX_SPLIT_POINTS) throw invalidOption(`The split position "${key}" must be a number of points from 0 to ${String(MAX_SPLIT_POINTS)}.`, "For example { split: { x: 100, y: 50 } }.");
						return v;
					};
					const x = at("x");
					const y = at("y");
					if (x === 0 && y === 0) throw invalidOption("A split needs \"x\" or \"y\" greater than 0.", "For example { split: { y: 50 } }, or { split: null } to remove it.");
					const corner = p.get("topLeftCell");
					if (corner === void 0 || corner === null) next.split = {
						x,
						y
					};
					else {
						const cell = typeof corner === "string" ? parseCellRef(corner.toUpperCase()) : null;
						if (cell === null) throw invalidOption("The split option \"topLeftCell\" must be a cell reference such as \"F20\".", "For example { split: { x: 100, y: 50, topLeftCell: \"F20\" } }: the first column shown right of the vertical bar and the first row below the horizontal bar.");
						next.split = {
							x,
							y,
							topLeftCell: formatCellRef(cell.row, cell.col)
						};
					}
					unfreeze = true;
				}
			}
			this.viewSettings = next;
			if (unfreeze) this.frozen = null;
			return this;
		});
	}
	/**
	* Adds filter buttons to the header row of a range, or removes them with `null`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').set('A1', 'Name').set('B1', 'Amount').autoFilter('A1:B100');
	* ```
	*/
	autoFilter(range) {
		const filter = range === null ? null : this.range(range);
		const table = filter === null ? void 0 : this.tableList.find((t) => rangesOverlap(t.range, filter));
		if (filter !== null && table !== void 0) throw new XlvError("XLV_E_INVALID_OPTION", `The filter ${formatRange(filter)} overlaps table "${table.name}" (${formatRange(table.range)}).`, {
			location: { sheet: this.name },
			hint: "A table has its own filter, and the sheet filter must not share cells with a table ([MS-OI29500] §2.1.680 h); filter a range outside the table.",
			details: {
				option: "autoFilter",
				table: table.name
			}
		});
		this.filter = filter;
		return this;
	}
	/**
	* Attaches a hyperlink to a cell or range, or removes it with `null`. `{ url }` links to a web
	* page, `mailto:` address or file; `{ location }` to a place in the workbook such as
	* `"'Sales 2026'!A1"` or a defined name. The cell text is set separately with `set`; a link
	* added again for the same range replaces the earlier one. Throws `XLV_E_INVALID_HYPERLINK`
	* (see `HyperlinkTarget` for the limits). xlvane never opens or checks the target.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Links');
	* sheet.set('A1', 'Example').link('A1', { url: 'https://example.com', tooltip: 'Open example.com' });
	* sheet.set('A2', 'Back to the top').link('A2', { location: 'Links!A1' });
	* ```
	*/
	link(range, target) {
		const r = this.range(range, "cell or range", "\"A1\" or \"A1:B2\"");
		const ref = r.start.row === r.end.row && r.start.col === r.end.col ? formatCellRef(r.start.row, r.start.col) : `${formatCellRef(r.start.row, r.start.col)}:${formatCellRef(r.end.row, r.end.col)}`;
		this.linkMap.delete(ref);
		if (target !== null) this.linkMap.set(ref, {
			ref,
			...normalizeHyperlink(target, {
				sheet: this.name,
				cell: ref
			})
		});
		return this;
	}
	/**
	* Attaches a note (comment) to a cell, or removes it with `null`. The note is plain text or
	* `{ text | runs, author? }` (see `NoteInput` for the limits); the cell does not need a value.
	* Notes are hidden until the pointer rests on the cell, as in spreadsheet applications. Throws
	* `XLV_E_INVALID_CELL_REF` or `XLV_E_INVALID_VALUE`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Budget');
	* sheet.set('B2', 1250).note('B2', 'Estimate from March');
	* sheet.note('C2', { runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'check the total' }], author: 'Ann' });
	* ```
	*/
	note(ref, note) {
		const at = this.address(ref);
		this.noteMap.delete(at.ref);
		if (note !== null) this.noteMap.set(at.ref, {
			row: at.row,
			col: at.col,
			...normalizeNote(note, {
				sheet: this.name,
				dateSystem: this.book.dateSystem
			}, at.ref)
		});
		return this;
	}
	/**
	* Protects the sheet (`sheetProtection`, SR-20), or removes the protection with `null`. Each
	* flag is `true` when the action is locked and `false` when it stays allowed; flags not given
	* keep the schema defaults (formatting, inserting, deleting, sorting, autofilter and pivot
	* tables locked; selecting cells, objects and scenarios allowed). A password is not supported
	* yet and throws `XLV_E_NOT_SUPPORTED`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Form').set('A1', 'Name').protect({ sort: false, autoFilter: false });
	* ```
	*/
	protect(options = {}) {
		this.print = {
			...this.print,
			protection: options === null ? null : normalizeProtection(options, this.name)
		};
		return this;
	}
	/**
	* Sets page setup for printing: margins, orientation, paper size, scale or fit to pages, page
	* order, print options and more (SR-23 … SR-26). Settings not given keep their value. Throws
	* `XLV_E_INVALID_OPTION` for values outside the limits of `PageSetupOptions`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').pageSetup({
	*   orientation: 'landscape',
	*   paperSize: 9,
	*   margins: { left: 0.5, right: 0.5 },
	*   fitToPage: true,
	*   fitToWidth: 1,
	*   fitToHeight: 0,
	*   printOptions: { gridLines: true },
	* });
	* ```
	*/
	pageSetup(options) {
		this.print = {
			...this.print,
			pageSetup: normalizePageSetup(options, this.print.pageSetup),
			pageSetupSet: true
		};
		return this;
	}
	/**
	* Sets printed headers and footers (SR-27), or removes them with `null`. Texts use the
	* formatting codes of `HeaderFooter`; settings not given keep their value.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').headerFooter({ oddHeader: '&C&A', oddFooter: '&RPage &P of &N' });
	* ```
	*/
	headerFooter(options) {
		this.print = {
			...this.print,
			headerFooter: options === null ? null : normalizeHeaderFooter(options, this.print.headerFooter)
		};
		return this;
	}
	/**
	* Sets the manual page breaks (SR-28): the rows and columns that start a new printed page.
	* Replaces earlier breaks; `{}` removes them.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').pageBreaks({ rows: [41, 81] });
	* ```
	*/
	pageBreaks(breaks) {
		const { rows, cols } = normalizePageBreaks(breaks);
		this.print = {
			...this.print,
			rowBreaks: rows,
			colBreaks: cols
		};
		return this;
	}
	/**
	* Sets the print area, one range or several, or removes it with `null`. It is stored as the
	* defined name `_xlnm.Print_Area` of the sheet (SR-30).
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').printArea('A1:H40');
	* ```
	*/
	printArea(ranges) {
		this.print = {
			...this.print,
			printArea: normalizePrintArea(ranges)
		};
		return this;
	}
	/**
	* Sets rows and columns repeated on every printed page, or removes them with `null`. They are
	* stored as the defined name `_xlnm.Print_Titles` of the sheet (SR-30).
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Report').printTitles({ rows: '1:2', columns: 'A' });
	* ```
	*/
	printTitles(titles) {
		this.print = {
			...this.print,
			printTitles: normalizePrintTitles(titles)
		};
		return this;
	}
	/**
	* Adds a data validation to a cell or ranges (`"B2:B50"`, or several separated by spaces), or
	* with `null` removes the validations added for exactly that range. Each call writes one
	* `dataValidation` element. See `DataValidationRule` for the rules and the limits; throws
	* `XLV_E_INVALID_OPTION` or `XLV_E_INVALID_CELL_REF`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Orders');
	* sheet.validate('C2:C500', { type: 'list', formula1: '"Open,Shipped,Paid"', showErrorMessage: true, error: 'Pick a status from the list.' });
	* sheet.validate('D2:D500', { type: 'whole', operator: 'greaterThan', formula1: 0, showInputMessage: true, prompt: 'Quantity, at least 1' });
	* ```
	*/
	validate(sqref, rule) {
		const where = {
			sheet: this.name,
			dateSystem: this.book.dateSystem
		};
		if (rule === null) {
			const range = normalizeDataValidation(sqref, { type: "none" }, where).sqref;
			for (let i = this.validationList.length - 1; i >= 0; i--) if (this.validationList[i]?.sqref === range) this.validationList.splice(i, 1);
			return this;
		}
		const dv = normalizeDataValidation(sqref, rule, where);
		if (this.validationList.length >= 65534) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The sheet "${this.name}" already has ${String(MAX_DATA_VALIDATIONS)} data validations, the most Office accepts.`, {
			location: { sheet: this.name },
			hint: "Give one validation a list of ranges (\"A1:A9 C1:C9\") instead of adding one per range.",
			details: {
				limit: "dataValidations",
				value: MAX_DATA_VALIDATIONS
			}
		});
		this.validationList.push(dv);
		return this;
	}
	/**
	* Adds conditional formatting to a cell or ranges: one or more rules, each with the format
	* applied where it is true. Rules get unique priorities on the sheet (1 is applied first). See
	* `ConditionalFormatRuleInput` for the rule types; throws `XLV_E_INVALID_OPTION`,
	* `XLV_E_INVALID_STYLE` or `XLV_E_INVALID_CELL_REF`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const sheet = createWorkbook().addSheet('Sales');
	* sheet.conditionalFormat('B2:B100', [
	*   { type: 'cellIs', operator: 'lessThan', formulas: [0], style: { font: { color: '9C0006' }, fill: { pattern: 'solid', fgColor: 'FFC7CE' } } },
	*   { type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6' } },
	* ]);
	* ```
	*/
	conditionalFormat(sqref, rules) {
		const cf = normalizeConditionalFormat(sqref, rules, {
			sheet: this.name,
			dateSystem: this.book.dateSystem
		}, this.priorities);
		this.conditionalList.push({
			sqref: cf.sqref,
			rules: cf.rules,
			dxfIds: cf.rules.map((r) => r.style === null ? null : this.book.internDxf(r.style))
		});
		return this;
	}
	/**
	* Adds a table (Part 1 §18.5): writes the header row (the column names), the data rows and the
	* totals row into the cells, and the table itself. `ref` is the top-left cell or the whole
	* range. Throws `XLV_E_INVALID_TABLE` for invalid or duplicate names (unique in the workbook,
	* also against defined names), duplicate column names, or a range that overlaps another table
	* or a merged range; see `TableInput`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().addSheet('Data').addTable({
	*   name: 'Sales',
	*   ref: 'A1',
	*   totalsRow: true,
	*   style: { theme: 'TableStyleMedium9', showRowStripes: true },
	*   columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }],
	*   rows: [['North', 1200], ['South', 980]],
	* });
	* ```
	*/
	addTable(table) {
		const t = normalizeTable(table, this.name);
		this.book.checkTable(t, this);
		const { start, end } = t.range;
		let row = start.row;
		if (t.headerRow) {
			t.columns.forEach((c, i) => this.set(formatCellRef(row, start.col + i), c.name));
			row++;
		}
		const dataEnd = end.row - (t.totalsRow ? 1 : 0);
		for (let r = row; r <= dataEnd; r++) {
			const values = t.rows[r - row] ?? [];
			t.columns.forEach((c, i) => {
				const value = values[i];
				const ref = formatCellRef(r, start.col + i);
				if (value !== void 0 && value !== null) this.set(ref, value);
				else if (c.calculatedColumnFormula !== null) this.set(ref, { formula: c.calculatedColumnFormula });
				else if (value === null) this.set(ref, null);
			});
		}
		if (t.totalsRow) t.columns.forEach((c, i) => {
			const ref = formatCellRef(end.row, start.col + i);
			const f = totalsFormula(t.displayName, c);
			if (f !== null) {
				const old = this.getCell(end.row, start.col + i)?.value;
				this.set(ref, { formula: f });
				if (old?.kind === "formula" && old.text === f && old.result !== null && old.group === void 0) {
					const cell = this.getCell(end.row, start.col + i);
					this.putCell(end.row, start.col + i, {
						...cell,
						value: {
							kind: "formula",
							text: f,
							result: old.result
						}
					});
				}
			} else if (c.totalsRowLabel !== null) this.set(ref, c.totalsRowLabel);
		});
		this.pushTable(t);
		return this;
	}
	/**
	* Places a picture on the sheet. `image` is the number `NewWorkbook.addImage` returned; the same
	* image can be placed many times and is stored once. `anchor` is a range such as `"B2:D6"`,
	* two corners `{ from, to }` (with `editAs`), or a corner and a size in pixels
	* (`{ from, size }`), see `ImageAnchorInput`. Fractional positions and pixel sizes use the
	* sheet's column widths and row heights at the time the workbook is written, at 96 pixels per
	* inch (TD-33). The anchor is copied, so changing its object afterwards does not move the
	* picture.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* declare const png: Uint8Array;
	* const workbook = createWorkbook();
	* const logo = workbook.addImage(png);
	* workbook.addSheet('Report').addImage(logo, 'B2:D6', { description: 'Company logo' });
	* ```
	*/
	addImage(image, anchor, options) {
		this.imageList.push(normalizeImage(image, this.book.imageCount, anchor, options));
		return this;
	}
	/**
	* Places a chart on the sheet: a column, bar, line, pie, scatter or area chart over ranges of
	* this or another sheet of the workbook (see `ChartInput`). `anchor` places it like a picture.
	* Each series gets an explicit colour from the default palette unless it has `color`, and the
	* file holds the current cell values of the ranges as caches, as Excel writes them
	* (docs/facts/charts.md CH-19, CH-21). The sheets named in the ranges must exist when the
	* workbook is written. Charts are drawn above the sheet's pictures. The caches of all charts
	* of a workbook cover at most 4 000 000 cells and hold at most 8 388 608 characters of text as
	* written, after escaping (limits of xlvane, checked when the workbook is written;
	* `writeWorkbookStream` writes no caches). The chart options and the anchor are copied, so changing the objects afterwards
	* does not change the chart. Throws `XLV_E_INVALID_CHART`.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook();
	* const sales = workbook.addSheet('Sales').set('A1', 'Month').set('B1', 'Sales');
	* ['Jan', 'Feb', 'Mar'].forEach((m, i) => sales.set(`A${String(i + 2)}`, m).set(`B${String(i + 2)}`, (i + 1) * 10));
	* sales.addChart({ type: 'column', anchor: 'D2:K16', title: 'Sales', series: [{ name: 'Sales!B1', categories: 'Sales!A2:A4', values: 'Sales!B2:B4' }] });
	* ```
	*/
	addChart(chart) {
		const normalized = normalizeChart(chart, this.name);
		const anchor = copyAnchor(chart.anchor);
		try {
			checkAnchor(anchor, "chart");
		} catch (error) {
			if (error instanceof XlvError) throw new XlvError(error.code, error.message, {
				location: { sheet: this.name },
				hint: error.hint,
				details: error.details
			});
			throw error;
		}
		this.chartList.push({
			anchor,
			chart: normalized
		});
		return this;
	}
	/**
	* Sets the picture shown behind the cells of the sheet (`picture` element, Part 1 §18.3.1.67),
	* or removes it with `null`. `image` is the number `NewWorkbook.addImage` returned.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* declare const png: Uint8Array;
	* const workbook = createWorkbook();
	* workbook.addSheet('Cover').background(workbook.addImage(png));
	* ```
	*/
	background(image) {
		if (image !== null) normalizeImage(image, this.book.imageCount, "A1", void 0);
		this.backgroundImage = image;
		return this;
	}
	/** @internal */
	data() {
		return {
			name: this.name,
			state: this.state,
			sheetId: this.sheetId,
			cells: this.cellMap,
			columns: this.columnMap,
			rows: this.rowMap,
			merges: this.mergeList,
			freeze: this.frozen,
			autoFilter: this.filter,
			tabColor: this.tab,
			view: this.viewSettings,
			properties: this.sheetProps,
			hyperlinks: [...this.linkMap.values()],
			notes: sortedNotes(this.noteMap),
			conditionalFormats: this.conditionalList,
			dataValidations: this.validationList,
			print: this.print,
			tables: this.tableList,
			images: this.imageList.map((i) => resolveImage(i, metricsOf(this.columnMap, this.rowMap))),
			background: this.backgroundImage,
			charts: this.chartList.map((c) => ({
				...resolveAnchor(c.anchor, metricsOf(this.columnMap, this.rowMap)),
				chart: c.chart
			}))
		};
	}
};
/** The box of a note without settings: automatic margins, locked, not moved or sized with cells. */
const DEFAULT_NOTE_BOX = Object.freeze({
	autoMargins: true,
	margins: null,
	locked: true,
	lockText: true,
	anchor: "absolute"
});
/** Most millimetres of a margin: the largest VML length in millimetres Office accepts ([MS-OI29500] §2.1.1789). */
const MAX_NOTE_MARGIN_MM = 59652;
function normalizeNoteBox(box, ctx, cell) {
	const usage = "For example box: { margins: { left: 2, top: 1, right: 2, bottom: 1 }, lockText: false, anchor: \"twoCell\" } (margins in millimetres).";
	const bad = (what) => {
		throw new XlvError("XLV_E_INVALID_OPTION", `The box of the note for ${cell}: ${what}`, {
			location: {
				sheet: ctx.sheet,
				cell
			},
			hint: usage,
			details: { option: "box" }
		});
	};
	if (typeof box !== "object" || box === null || Array.isArray(box)) return bad("it must be an object.");
	const o = box;
	for (const k of Object.keys(o)) if (![
		"autoMargins",
		"margins",
		"locked",
		"lockText",
		"anchor"
	].includes(k)) bad(`"${k.slice(0, 40)}" is not a box setting (autoMargins, margins, locked, lockText, anchor).`);
	for (const k of [
		"autoMargins",
		"locked",
		"lockText"
	]) if (o[k] !== void 0 && typeof o[k] !== "boolean") bad(`"${k}" must be true or false.`);
	if (o["anchor"] !== void 0 && o["anchor"] !== "twoCell" && o["anchor"] !== "oneCell" && o["anchor"] !== "absolute") bad("\"anchor\" must be \"twoCell\", \"oneCell\" or \"absolute\".");
	let margins = null;
	if (o["margins"] !== void 0) {
		const m = o["margins"];
		if (typeof m !== "object" || m === null || Array.isArray(m)) return bad("\"margins\" must be { left, top, right, bottom } in millimetres.");
		const mm = m;
		for (const k of Object.keys(mm)) if (![
			"left",
			"top",
			"right",
			"bottom"
		].includes(k)) bad(`"margins.${k.slice(0, 40)}" is not a margin (left, top, right, bottom).`);
		const side = (k) => {
			const v = mm[k];
			if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > MAX_NOTE_MARGIN_MM) return bad(`"margins.${k}" must be a number of millimetres from 0 to ${String(MAX_NOTE_MARGIN_MM)}.`);
			return v;
		};
		margins = Object.freeze({
			left: side("left"),
			top: side("top"),
			right: side("right"),
			bottom: side("bottom")
		});
	}
	return Object.freeze({
		autoMargins: o["autoMargins"] ?? margins === null,
		margins,
		locked: o["locked"] ?? true,
		lockText: o["lockText"] ?? true,
		anchor: o["anchor"] ?? "absolute"
	});
}
/** The notes of a sheet in row, then column order (deterministic output). */
function sortedNotes(notes) {
	return [...notes.values()].sort((a, b) => a.row - b.row || a.col - b.col);
}
/**
* Validates a {@link NoteInput}: text as a cell string, runs as `{ richText }`, the author up to
* 65 535 characters ([MS-OI29500] §2.1.689 a). Throws `XLV_E_INVALID_VALUE` naming the cell.
*/
function normalizeNote(note, ctx, cell) {
	const usage = "Pass the text, or { text: \"…\", author: \"Ann\" }, or { runs: [{ text: \"…\", font: { bold: true } }] }.";
	let input = note;
	let author = "";
	let box = DEFAULT_NOTE_BOX;
	if (typeof note === "object" && note !== null && !Array.isArray(note)) {
		const keys = Object.keys(note);
		const o = note;
		if (keys.some((k) => k !== "text" && k !== "runs" && k !== "author" && k !== "box") || o.text === void 0 === (o.runs === void 0)) throw invalidValue(ctx.sheet, cell, `The note for ${cell} must have exactly one of "text" and "runs", and optionally "author" and "box".`, usage);
		if (o.box !== void 0) box = normalizeNoteBox(o.box, ctx, cell);
		if (o.author !== void 0) {
			if (typeof o.author !== "string" || o.author.length > 65535 || firstLoneSurrogate(o.author) !== -1) throw invalidValue(ctx.sheet, cell, `The author of the note for ${cell} must be text of at most ${String(MAX_AUTHOR_LENGTH)} characters.`, "Pass the author as a string, for example author: \"Ann\".");
			author = o.author;
		}
		input = o.text !== void 0 ? o.text : { richText: o.runs };
		if (o.text !== void 0 && typeof o.text !== "string") throw invalidValue(ctx.sheet, cell, `The text of the note for ${cell} must be a string.`, usage);
	} else if (typeof note !== "string") throw invalidValue(ctx.sheet, cell, `The note for ${cell} must be text or an object.`, usage);
	let stored;
	try {
		stored = convertCellInput(input, ctx, cell).stored;
	} catch (error) {
		if (!(error instanceof XlvError)) throw error;
		throw new XlvError(error.code, `The note for ${cell}: ${error.message}`, {
			location: {
				sheet: ctx.sheet,
				cell
			},
			hint: error.hint,
			details: error.details
		});
	}
	return {
		value: stored?.kind === "string" ? stored : {
			kind: "string",
			s: ""
		},
		author,
		box
	};
}
/**
* Validates and converts a {@link CellInput} as `NewSheet.set` does (shared with the streaming
* writer). Throws `XLV_E_INVALID_VALUE` with the sheet and cell.
*/
function convertCellInput(value, ctx, cell) {
	if (value === null || value === void 0) return {
		stored: null,
		isDate: false,
		hasTime: false
	};
	if (typeof value === "object" && !(value instanceof Date)) {
		const keys = Object.keys(value);
		if (keys.includes("richText")) return {
			stored: readRichText(value, ctx.sheet, cell),
			isDate: false,
			hasTime: false
		};
		if (keys.includes("formula") || keys.includes("sharedFormula")) {
			const o = keys.includes("formula") ? readFormula(value, ctx.sheet, cell) : readSharedDependent(value, ctx.sheet, cell);
			const extra = {
				...o.group === void 0 ? {} : { group: o.group },
				...o.master === void 0 ? {} : { master: o.master }
			};
			if (o.result === void 0) return {
				stored: {
					kind: "formula",
					text: o.text,
					result: null,
					...extra
				},
				isDate: false,
				hasTime: false
			};
			const r = scalarValue(ctx, o.result, cell, "formula result");
			return {
				stored: {
					kind: "formula",
					text: o.text,
					result: r.stored,
					...extra
				},
				isDate: r.isDate,
				hasTime: r.hasTime
			};
		}
	}
	return scalarValue(ctx, value, cell, "value");
}
function scalarValue(ctx, value, cell, what) {
	const plain = (stored) => ({
		stored,
		isDate: false,
		hasTime: false
	});
	if (typeof value === "number") {
		if (!Number.isFinite(value)) throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} is ${String(value)}.`, "Spreadsheet cells hold finite numbers only. Write the value as text or leave the cell empty.");
		return plain({
			kind: "number",
			n: Object.is(value, -0) ? 0 : value
		});
	}
	if (typeof value === "string") {
		if (value.length > MAX_CELL_TEXT) throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} has ${String(value.length)} characters.`, `A cell can contain at most ${String(MAX_CELL_TEXT)} characters. Split the text over several cells.`);
		if (firstLoneSurrogate(value) !== -1) throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} contains an unpaired UTF-16 surrogate.`, "The text is not valid Unicode; it may have been cut in the middle of a character such as an emoji. Fix the text before writing it.");
		return plain({
			kind: "string",
			s: value
		});
	}
	if (typeof value === "boolean") return plain({
		kind: "boolean",
		b: value
	});
	if (value instanceof Date) {
		const time = value.getTime();
		let serial;
		try {
			serial = dateToSerial(value, ctx.dateSystem);
		} catch {
			serial = NaN;
		}
		if (!(serial >= 0)) throw invalidValue(ctx.sheet, cell, Number.isNaN(time) ? `The ${what} for ${cell} is an invalid Date.` : `The date ${value.toISOString()} for ${cell} is outside the ${ctx.dateSystem} date system.`, ctx.dateSystem === "1904" ? "The 1904 date system covers 1904-01-01 to 9999-12-31. Write earlier dates as text." : "The 1900 date system covers 1899-12-31 (time-only values) to 9999-12-31. Write earlier dates as text.");
		return {
			stored: {
				kind: "number",
				n: serial
			},
			isDate: true,
			hasTime: time % DAY_MS !== 0
		};
	}
	if (typeof value === "object" && value !== null && Object.keys(value).includes("error")) {
		const keys = Object.keys(value);
		const e = value.error;
		if (keys.length === 1 && typeof e === "string" && ERROR_VALUES.includes(e)) return plain({
			kind: "error",
			e
		});
		throw invalidValue(ctx.sheet, cell, `The error value for ${cell} is not valid.`, `Use { error: code } with one of: ${ERROR_VALUES.join(", ")}.`);
	}
	throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} has an unsupported type (${typeof value === "object" ? "object" : typeof value}).`, "Use a number, string, boolean, Date, null, { error: \"#N/A\" }, { formula: \"A1*2\", result: 2 } or { richText: [{ text: \"a\" }] }.");
}
/**
* Validates `{ richText: [{ text, font? }, …] }` (F-SST-08, F-SST-09): at least one run, at most
* 32,767 runs and 32,767 characters in all, fonts as in cell styles. Runs with the same font next
* to each other are kept apart, as given.
*/
function readRichText(value, sheet, cell) {
	const usage = "Use { richText: [{ text: \"Total: \", font: { bold: true } }, { text: \"42\" }] }.";
	const keys = Object.keys(value);
	const runsIn = value.richText;
	if (keys.length !== 1 || !Array.isArray(runsIn) || runsIn.length === 0) throw invalidValue(sheet, cell, `The rich text for ${cell} must be an object with only "richText", a non-empty array of runs.`, usage);
	if (runsIn.length > MAX_RICH_TEXT_RUNS) throw invalidValue(sheet, cell, `The rich text for ${cell} has ${String(runsIn.length)} runs.`, `A rich text value can have at most ${String(MAX_RICH_TEXT_RUNS)} runs. Join runs that have the same font.`);
	const runs = [];
	let text = "";
	for (let i = 0; i < runsIn.length; i++) {
		const run = runsIn[i];
		const runKeys = typeof run === "object" && run !== null && !Array.isArray(run) ? Object.keys(run) : null;
		const t = runKeys === null ? void 0 : run.text;
		if (runKeys === null || runKeys.some((k) => k !== "text" && k !== "font") || typeof t !== "string") throw invalidValue(sheet, cell, `Run ${String(i)} of the rich text for ${cell} is not { text, font? } with a string text.`, usage);
		if (firstLoneSurrogate(t) !== -1) throw invalidValue(sheet, cell, `Run ${String(i)} of the rich text for ${cell} contains an unpaired UTF-16 surrogate.`, "The text is not valid Unicode; it may have been cut in the middle of a character such as an emoji. Fix the text before writing it.");
		text += t;
		const fontIn = runKeys.includes("font") ? run.font : void 0;
		let font;
		if (fontIn !== void 0) try {
			font = normalizeStyle({ font: fontIn }).font;
		} catch (error) {
			if (!(error instanceof XlvError)) throw error;
			throw new XlvError(error.code, `Run ${String(i)} of the rich text for ${cell}: ${error.message}`, {
				location: {
					sheet,
					cell
				},
				hint: error.hint,
				details: {
					...error.details,
					run: i
				}
			});
		}
		runs.push(font === void 0 ? { text: t } : {
			text: t,
			font
		});
	}
	if (text.length > MAX_CELL_TEXT) throw invalidValue(sheet, cell, `The rich text for ${cell} has ${String(text.length)} characters.`, `A cell can contain at most ${String(MAX_CELL_TEXT)} characters. Split the text over several cells.`);
	return {
		kind: "string",
		s: text,
		runs
	};
}
function readFormula(value, sheet, cell) {
	const keys = Object.keys(value);
	for (const key of keys) if (key !== "formula" && key !== "result" && key !== "array" && key !== "shared") throw invalidValue(sheet, cell, `Unknown property "${key.slice(0, 40)}" in the formula for ${cell}.`, "Use { formula: \"A1*2\" } or { formula: \"A1*2\", result: 2 }; add array: \"A1:B2\" or shared: \"A1:A10\" for array or shared formulas.");
	let group;
	if (keys.includes("array") && keys.includes("shared")) throw invalidValue(sheet, cell, `The formula for ${cell} cannot be both an array formula and a shared formula.`, "Give either array: \"A1:B2\" or shared: \"A1:A10\", not both.");
	for (const t of ["array", "shared"]) {
		if (!keys.includes(t)) continue;
		const raw = value[t];
		const range = typeof raw === "string" ? parseRange(raw) : null;
		const at = parseCellRef(cell);
		if (range === null || at === null || range.start.row !== at.row || range.start.col !== at.col) throw invalidValue(sheet, cell, `The ${t} range of the formula for ${cell} must be a range whose top-left cell is ${cell}, got ${JSON.stringify(String(raw).slice(0, 40))}.`, `Put the formula in the top-left cell of its range, for example set("${cell}", { formula: "…", ${t}: "${cell}:${cell.replace(/\d+$/, (n) => String(Number(n) + 2))}" }).`);
		group = {
			t,
			ref: formatRange(range)
		};
	}
	const raw = value.formula;
	const text = typeof raw === "string" && raw.startsWith("=") ? raw.slice(1) : raw;
	if (typeof text !== "string" || text.trim().length === 0 || text.length > MAX_FORMULA_LENGTH || firstLoneSurrogate(text) !== -1) throw invalidValue(sheet, cell, `The formula for ${cell} is empty, too long or not valid text.`, `Pass the formula text, with or without "=", for example { formula: "SUM(A1:A10)" }; at most ${String(MAX_FORMULA_LENGTH)} characters.`);
	return {
		text,
		result: keys.includes("result") ? value.result : void 0,
		...group === void 0 ? {} : { group }
	};
}
/** `{ sharedFormula: "C1", result? }`: a cell that shares the formula of the master in C1 (E-28). */
function readSharedDependent(value, sheet, cell) {
	const keys = Object.keys(value);
	for (const key of keys) if (key !== "sharedFormula" && key !== "result") throw invalidValue(sheet, cell, `Unknown property "${key.slice(0, 40)}" in the shared formula for ${cell}.`, "Use { sharedFormula: \"C1\" } or { sharedFormula: \"C1\", result: 2 }, where C1 holds { formula, shared: \"C1:C10\" }.");
	const raw = value.sharedFormula;
	const master = typeof raw === "string" ? parseCellRef(raw) : null;
	if (master === null) throw invalidValue(sheet, cell, `sharedFormula for ${cell} must name one cell, the master of the shared formula, got ${JSON.stringify(String(raw).slice(0, 40))}.`, "Name the cell that holds { formula, shared: \"C1:C10\" }, for example { sharedFormula: \"C1\" }.");
	return {
		text: "",
		result: keys.includes("result") ? value.result : void 0,
		master: formatCellRef(master.row, master.col)
	};
}
/**
* A workbook being built in memory, created by {@link createWorkbook} and written with
* {@link writeWorkbook}.
*
* @example
* ```ts
* import { createWorkbook, writeWorkbook } from 'xlvane';
*
* const workbook = createWorkbook();
* workbook.addSheet('Summary').set('A1', 'Hello');
* const bytes = writeWorkbook(workbook);
* ```
*/
var NewWorkbook = class {
	/** Date system used to store dates. */
	dateSystem;
	/** Document properties given to {@link createWorkbook}. */
	properties;
	list = [];
	names = [];
	/** Sheets by lower-case name. */
	sheetsByName = /* @__PURE__ */ new Map();
	/** Sheet names as given, in sheet order. */
	sheetNames = /* @__PURE__ */ new Set();
	/** Sheets by the sheetId given to addSheet. */
	sheetsById = /* @__PURE__ */ new Map();
	/** Defined names by nameScopeKey. */
	namesByScope = /* @__PURE__ */ new Map();
	/** Lower-case defined names, of any scope. */
	lowerNames = /* @__PURE__ */ new Set();
	/** Lower-case names and display names of the tables of all sheets. */
	tableNames = /* @__PURE__ */ new Set();
	styleIds = /* @__PURE__ */ new Map();
	styleList = [{}];
	/** @internal `workbookProtection`, set by {@link NewWorkbook.protect}. */
	workbookProtection = null;
	/** @internal Workbook windows set by {@link NewWorkbook.views}; empty for the default window. */
	bookViews = [];
	/** @internal `calcPr`, set by {@link NewWorkbook.calculation}. */
	calcProps = null;
	/** @internal Font 0 given to {@link createWorkbook}, or `null` for the default (W2-12). */
	fontDefault;
	/** @internal Use {@link createWorkbook}. */
	constructor(dateSystem, properties = {}, defaultFont = null) {
		this.dateSystem = dateSystem;
		this.properties = properties;
		this.fontDefault = defaultFont;
		this.styleIds.set(stableKey({}), 0);
	}
	/**
	* Protects the workbook (`workbookProtection`, SR-22), by default its structure: sheets cannot
	* be added, deleted, moved or renamed in a spreadsheet application without the password.
	* `null` removes the protection. A password string is hashed with `hashPassword` defaults
	* (D-C08); it is not encryption, and the content stays readable.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook().protect({ password: 'secret' });
	* workbook.addSheet('Data').set('A1', 1);
	* ```
	*/
	protect(options = {}) {
		this.workbookProtection = options === null ? null : normalizeWorkbookProtection(options);
		return this;
	}
	/**
	* Sets the workbook windows (`bookViews`, W2-66). Without a call one default window is written,
	* showing the first visible sheet.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook().views([{ activeTab: 1, firstSheet: 0 }]);
	* workbook.addSheet('A');
	* workbook.addSheet('B');
	* ```
	*/
	views(views) {
		if (!Array.isArray(views)) throw new XlvError("XLV_E_INVALID_OPTION", "views() expects an array of window settings.", {
			hint: "For example views([{ activeTab: 0 }]).",
			details: { option: "views" }
		});
		const out = [];
		views.forEach((v, i) => {
			const o = readOptions$1(v, [
				"x",
				"y",
				"width",
				"height",
				"firstSheet",
				"activeTab",
				"visibility"
			], `views[${String(i)}]`);
			const view = {};
			for (const key of [
				"x",
				"y",
				"width",
				"height",
				"firstSheet",
				"activeTab"
			]) {
				const n = o.get(key);
				if (n === void 0) continue;
				const signed = key === "x" || key === "y";
				if (typeof n !== "number" || !Number.isInteger(n) || (signed ? n < -2147483648 || n > 2147483647 : n < 0 || n > 4294967295)) throw new XlvError("XLV_E_INVALID_OPTION", `views[${String(i)}].${key} must be a whole number${signed ? "" : " of 0 or more"}.`, {
					hint: `For example { ${key}: 0 }.`,
					details: { option: key }
				});
				view[key] = n;
			}
			const visibility = o.get("visibility");
			if (visibility !== void 0) {
				if (visibility !== "visible" && visibility !== "hidden" && visibility !== "veryHidden") throw new XlvError("XLV_E_INVALID_OPTION", `views[${String(i)}].visibility must be "visible", "hidden" or "veryHidden".`, {
					hint: "For example { visibility: \"visible\" }.",
					details: { option: "visibility" }
				});
				view.visibility = visibility;
			}
			out.push(view);
		});
		this.bookViews = out;
		return this;
	}
	/**
	* Sets calculation settings (`calcPr`, W2-67).
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* createWorkbook().calculation({ fullCalcOnLoad: true });
	* ```
	*/
	calculation(options) {
		if (options === null) {
			this.calcProps = null;
			return this;
		}
		const full = readOptions$1(options, ["fullCalcOnLoad"], "calculation options").get("fullCalcOnLoad") ?? false;
		if (typeof full !== "boolean") throw invalidOption("The calculation option \"fullCalcOnLoad\" must be true or false.", "Use { fullCalcOnLoad: true }.");
		this.calcProps = { fullCalcOnLoad: full };
		return this;
	}
	/** Sheets in workbook order. */
	get sheets() {
		return [...this.list];
	}
	/**
	* Adds a sheet at the end. Throws `XLV_E_INVALID_SHEET_NAME` for names Excel does not accept
	* and `XLV_E_DUPLICATE_SHEET_NAME` if a sheet with the same name (ignoring case) exists.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook();
	* workbook.addSheet('Data');
	* workbook.addSheet('Lookup', { state: 'hidden', sheetId: 5 });
	* ```
	*/
	addSheet(name, options) {
		const o = readOptions$1(options, ["state", "sheetId"], "sheet options");
		const state = o.get("state") ?? "visible";
		if (state !== "visible" && state !== "hidden" && state !== "veryHidden") throw invalidOption(`Sheet state ${JSON.stringify(String(state).slice(0, 40))} is not valid.`, "Use \"visible\", \"hidden\" or \"veryHidden\".");
		const checked = checkSheetName(name);
		const lower = checked.toLowerCase();
		const existing = this.sheetsByName.get(lower);
		if (existing) throw new XlvError("XLV_E_DUPLICATE_SHEET_NAME", `The workbook already has a sheet named "${existing.name}".`, {
			hint: "Sheet names must be unique, ignoring upper and lower case. Choose another name.",
			details: {
				name: checked,
				existing: existing.name
			}
		});
		let sheetId = null;
		if (o.has("sheetId")) {
			const id = o.get("sheetId");
			if (typeof id !== "number" || !Number.isInteger(id) || id < 1 || id > MAX_SHEET_ID) throw invalidOption(`The sheet option "sheetId" must be a whole number from 1 to ${String(MAX_SHEET_ID)}.`, "For example { sheetId: 3 }, or leave it out to take the lowest unused id.", { name: checked });
			const owner = this.sheetsById.get(id);
			if (owner) throw invalidOption(`Sheet "${owner.name}" already has the sheetId ${String(id)}.`, "Sheet ids must be unique in the workbook. Give another id, or leave it out to take the lowest unused id.", {
				name: checked,
				existing: owner.name
			});
			sheetId = id;
		}
		const sheet = new NewSheet(checked, state, this, sheetId);
		this.list.push(sheet);
		this.sheetsByName.set(lower, sheet);
		this.sheetNames.add(checked);
		if (sheetId !== null) this.sheetsById.set(sheetId, sheet);
		return sheet;
	}
	/**
	* Returns the sheet with this name (exact match); throws `XLV_E_SHEET_NOT_FOUND` otherwise.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook();
	* workbook.addSheet('Data');
	* workbook.getSheet('Data').set('A1', 1);
	* ```
	*/
	getSheet(name) {
		const found = this.list.find((s) => s.name === name);
		if (found) return found;
		throw sheetNotFound(name, this.list.map((s) => s.name));
	}
	/**
	* Defines a name for a cell, range, constant or formula. `refersTo` is formula text with or
	* without `=`, for example `"Data!$A$2:$A$100"`; it is stored as given. A name is local to a
	* sheet when `options.sheet` is set. Throws `XLV_E_INVALID_DEFINED_NAME` for a name Excel does
	* not accept or a duplicate in the same scope (names are not case-sensitive), and
	* `XLV_E_SHEET_NOT_FOUND` for an unknown sheet.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* const workbook = createWorkbook();
	* workbook.addSheet('Data').set('A1', 0.19);
	* workbook.defineName('VatRate', 'Data!$A$1');
	* workbook.defineName('Rows', 'Data!$A$2:$A$100', { sheet: 'Data', hidden: true });
	* ```
	*/
	defineName(name, refersTo, options) {
		const n = checkDefinedName(name, refersTo, options, this.sheetNames);
		checkUniqueName(n, this.namesByScope);
		if (this.tableNames.has(n.name.toLowerCase())) throw new XlvError("XLV_E_INVALID_DEFINED_NAME", `"${n.name}" is already the name of a table.`, {
			hint: "Table names and defined names share one list in a workbook (ECMA-376 Part 1 §18.5.1.2); choose another name.",
			details: { name: n.name }
		});
		this.names.push(n);
		this.namesByScope.set(nameScopeKey(n), n);
		this.lowerNames.add(n.name.toLowerCase());
		return this;
	}
	/** @internal Records the names of a table added to a sheet ({@link NewSheet.pushTable}). */
	tableAdded(t) {
		this.tableNames.add(t.name.toLowerCase());
		this.tableNames.add(t.displayName.toLowerCase());
	}
	/**
	* @internal Checks that a table's names are unique in the workbook (tables and defined names,
	* ignoring case; Part 1 §18.5.1.2) and that it overlaps no table or merged range of its sheet.
	*/
	checkTable(t, sheet) {
		const where = { sheet: sheet.name };
		const names = /* @__PURE__ */ new Set([t.name.toLowerCase(), t.displayName.toLowerCase()]);
		const clashes = (known) => [...names].some((name) => known.has(name));
		if (clashes(this.tableNames)) {
			for (const other of this.list.flatMap((s) => s.tableList)) if (names.has(other.name.toLowerCase()) || names.has(other.displayName.toLowerCase())) throw new XlvError("XLV_E_INVALID_TABLE", `The workbook already has a table named "${other.displayName}".`, {
				location: where,
				hint: "Table names are unique in a workbook, ignoring case; choose another name, for example \"Sales_2\".",
				details: { name: t.name }
			});
		}
		if (clashes(this.lowerNames)) {
			for (const n of this.names) if (names.has(n.name.toLowerCase())) throw new XlvError("XLV_E_INVALID_TABLE", `"${n.name}" is already a defined name of the workbook.`, {
				location: where,
				hint: "Table names and defined names share one list in a workbook (ECMA-376 Part 1 §18.5.1.2); choose another table name.",
				details: { name: t.name }
			});
		}
		const ref = (r) => `${formatCellRef(r.start.row, r.start.col)}:${formatCellRef(r.end.row, r.end.col)}`;
		if (sheet.overlapsTable(t.range)) {
			for (const other of sheet.tableList) if (rangesOverlap(t.range, other.range)) throw new XlvError("XLV_E_INVALID_TABLE", `Table "${t.name}" (${ref(t.range)}) overlaps table "${other.name}" (${ref(other.range)}).`, {
				location: where,
				hint: "Tables must not share cells ([MS-OI29500] §2.1.680 h); move one of them.",
				details: { name: t.name }
			});
		}
		if (sheet.filter !== null && rangesOverlap(t.range, sheet.filter)) throw new XlvError("XLV_E_INVALID_TABLE", `Table "${t.name}" (${ref(t.range)}) overlaps the sheet filter ${ref(sheet.filter)}.`, {
			location: where,
			hint: "A table has its own filter, and the sheet filter must not share cells with a table ([MS-OI29500] §2.1.680 h); remove the sheet filter (autoFilter(null)) or move the table.",
			details: { name: t.name }
		});
		if (sheet.overlapsMerge(t.range)) {
			for (const m of sheet.mergeList) if (rangesOverlap(t.range, m)) throw new XlvError("XLV_E_INVALID_TABLE", `Table "${t.name}" (${ref(t.range)}) overlaps the merged range ${ref(m)}.`, {
				location: where,
				hint: "Tables cannot contain merged cells; unmerge the range or move the table.",
				details: { name: t.name }
			});
		}
	}
	/** Defined names in the order they were added. */
	get definedNames() {
		return [...this.names];
	}
	/** @internal Returns the id of a normalised style, adding it if new. */
	intern(style) {
		const key = stableKey(style);
		let id = this.styleIds.get(key);
		if (id === void 0) {
			id = this.styleList.length;
			this.styleList.push(style);
			this.styleIds.set(key, id);
		}
		return id;
	}
	/** @internal */
	styleOf(id) {
		return this.styleList[id] ?? {};
	}
	/** @internal */
	styles() {
		return this.styleList;
	}
	dxfIds = /* @__PURE__ */ new Map();
	dxfList = [];
	/** @internal Index of a differential format (conditional formatting), deduplicated (SR-16). */
	internDxf(style) {
		const key = stableKey(style);
		let id = this.dxfIds.get(key);
		if (id === void 0) {
			id = this.dxfList.length;
			this.dxfList.push(style);
			this.dxfIds.set(key, id);
		}
		return id;
	}
	/** @internal */
	dxfStyles() {
		return this.dxfList;
	}
	imageData = [];
	/**
	* Adds an image (the bytes of a PNG, JPEG or GIF file, recognised from its first bytes; EMF and
	* WMF files, as Excel stores them, are written unchanged) and
	* returns its number for `NewSheet.addImage` and `NewSheet.background`. Other formats throw
	* `XLV_E_INVALID_IMAGE`. Identical images are stored once in the file.
	*
	* @example
	* ```ts
	* import { createWorkbook } from 'xlvane';
	*
	* declare const jpeg: Uint8Array;
	* const workbook = createWorkbook();
	* const photo = workbook.addImage(jpeg);
	* workbook.addSheet('Team').addImage(photo, { from: 'A1', size: { width: 120, height: 160 } });
	* ```
	*/
	addImage(bytes) {
		this.imageData.push(checkImage(bytes));
		return this.imageData.length - 1;
	}
	/** @internal */
	get imageCount() {
		return this.imageData.length;
	}
	/** @internal */
	imageList() {
		return this.imageData;
	}
};
/**
* Creates an empty workbook. Add sheets with {@link NewWorkbook.addSheet} and write it with
* {@link writeWorkbook}.
*
* @example
* ```ts
* import { createWorkbook, writeWorkbook } from 'xlvane';
*
* const workbook = createWorkbook();
* const sheet = workbook.addSheet('Sales');
* sheet.set('A1', 'Region', { font: { bold: true } }).set('B1', 'Revenue', { font: { bold: true } });
* sheet.set('A2', 'North').set('B2', 1250.5, { numFmt: '#,##0.00' });
* const bytes = writeWorkbook(workbook); // Uint8Array with the .xlsx file
* ```
*/
function createWorkbook(options) {
	const o = readOptions$1(options, [
		"dateSystem",
		"properties",
		"defaultFont"
	], "workbook options");
	const dateSystem = o.get("dateSystem") ?? "1900";
	if (dateSystem !== "1900" && dateSystem !== "1904") throw invalidOption(`Date system ${JSON.stringify(String(dateSystem).slice(0, 20))} is not valid.`, "Use \"1900\" (the default) or \"1904\".");
	const font = o.get("defaultFont");
	const defaultFont = font === void 0 ? null : normalizeStyle({ font }).font ?? null;
	return new NewWorkbook(dateSystem, normalizeProperties(o.get("properties")), defaultFont);
}
/**
* Checks that `activeTab` and `firstSheet` of every workbook window name an existing sheet: both
* are 0-based sheet indexes (CT_BookView; the default window written without `views()` uses the
* active sheet's index the same way). Throws `XLV_E_INVALID_OPTION`.
*/
function checkBookViews(views, sheetNames) {
	const count = sheetNames.length;
	views.forEach((view, i) => {
		for (const key of ["activeTab", "firstSheet"]) {
			const n = view[key];
			if (n === void 0 || n < count) continue;
			const listed = sheetNames.slice(0, 10).map((name, k) => `${String(k)} = "${name}"`).join(", ");
			throw new XlvError("XLV_E_INVALID_OPTION", `views[${String(i)}].${key} is ${String(n)}, but the workbook has ${String(count)} sheet${count === 1 ? "" : "s"} (indexes 0 to ${String(count - 1)}).`, {
				hint: `Use a sheet index from 0 to ${String(count - 1)}; ${key === "activeTab" ? "activeTab is the sheet shown when the file opens" : "firstSheet is the first sheet tab shown"}. Sheets: ${listed}${count > 10 ? ", …" : ""}.`,
				details: {
					option: key,
					value: n
				}
			});
		}
	});
}
/**
* Writes a workbook as an XLSX file (Office Open XML, Transitional) and returns its bytes. The
* output is deterministic: the same workbook always gives the same bytes. Throws
* `XLV_E_INVALID_INPUT` for a workbook without sheets or without a visible sheet.
*
* @example
* ```ts
* import { createWorkbook, writeWorkbook, readWorkbook } from 'xlvane';
*
* const workbook = createWorkbook();
* workbook.addSheet('Data').set('A1', 42);
* const bytes = writeWorkbook(workbook);
* console.log(readWorkbook(bytes).getSheet('Data').cell('A1')?.value); // 42
* ```
*/
function writeWorkbook(workbook, options = {}) {
	if (!(workbook instanceof NewWorkbook)) throw new XlvError("XLV_E_INVALID_INPUT", "writeWorkbook expects a workbook created with createWorkbook().", { hint: "Create the workbook with createWorkbook(), add sheets with addSheet() and pass it to writeWorkbook()." });
	const sheets = workbook.sheets;
	if (sheets.length === 0) throw new XlvError("XLV_E_INVALID_INPUT", "The workbook has no sheets.", { hint: "Add at least one sheet with workbook.addSheet(\"Sheet1\") before writing." });
	if (!sheets.some((s) => s.state === "visible")) throw new XlvError("XLV_E_INVALID_INPUT", "All sheets of the workbook are hidden.", { hint: "A workbook needs at least one visible sheet. Add one or create a sheet with { state: \"visible\" }." });
	checkBookViews(workbook.bookViews, sheets.map((s) => s.name));
	const o = readOptions$1(options, ["sharedStrings", "compressionLevel"], "write options");
	const shared = o.get("sharedStrings") ?? true;
	if (typeof shared !== "boolean") throw invalidOption("The write option \"sharedStrings\" must be true or false.", "Use { sharedStrings: false } for inline strings.");
	const level = o.get("compressionLevel") ?? 6;
	if (typeof level !== "number" || !Number.isInteger(level) || level < 0 || level > 9) throw invalidOption("The write option \"compressionLevel\" must be a whole number from 0 to 9.", "Use 0 for no compression or 9 for the smallest file; the default is 6.");
	return writePackage({
		sharedStrings: shared,
		compressionLevel: level,
		bookViews: workbook.bookViews,
		calculation: workbook.calcProps,
		dateSystem: workbook.dateSystem,
		sheets: sheets.map((s) => s.data()),
		styles: workbook.styles(),
		properties: workbook.properties,
		definedNames: [...workbook.definedNames, ...sheets.flatMap((s) => printNames(s.name, s.print))],
		dxfs: workbook.dxfStyles(),
		workbookProtection: workbook.workbookProtection,
		images: workbook.imageList(),
		defaultFont: workbook.fontDefault
	});
}
// src/util/options.ts
/** Reads `value` as an options object with the given known names. `undefined` gives an empty map. */
function readOptions(value, known, what) {
	const out = /* @__PURE__ */ new Map();
	if (value === void 0) return out;
	if (typeof value !== "object" || value === null || Array.isArray(value)) throw new XlvError("XLV_E_INVALID_OPTION", `${what} must be an object.`, { hint: `Pass an object with some of these properties: ${known.join(", ")}, or leave it out.` });
	for (const key of Object.keys(value)) {
		if (!known.includes(key)) {
			const close = suggest(key, known);
			throw new XlvError("XLV_E_INVALID_OPTION", `Unknown option "${key.slice(0, 40)}" in ${what}.`, {
				hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Known options: ${known.join(", ")}.`,
				details: {
					option: key.slice(0, 40),
					known
				}
			});
		}
		const v = value[key];
		if (v !== void 0) out.set(key, v);
	}
	return out;
}
/** Throws `XLV_E_INVALID_OPTION` for an option value outside the allowed set. */
function invalidOptionValue(name, value, allowed) {
	return new XlvError("XLV_E_INVALID_OPTION", `The option "${name}" has the value ${typeof value === "string" ? JSON.stringify(value.slice(0, 40)) : typeof value === "object" ? typeof value : String(value)}.`, {
		hint: `Use ${allowed}.`,
		details: { option: name }
	});
}
/** Returns `value` if it is one of `allowed`, otherwise throws `XLV_E_INVALID_OPTION`. */
function oneOfOption(name, value, allowed, fallback) {
	if (value === void 0) return fallback;
	if (typeof value === "string" && allowed.includes(value)) return value;
	throw invalidOptionValue(name, value, allowed.map((a) => `"${a}"`).join(", "));
}
/** Returns `value` if it is a boolean, otherwise throws `XLV_E_INVALID_OPTION`. */
function booleanOption(name, value, fallback) {
	if (value === void 0) return fallback;
	if (typeof value === "boolean") return value;
	throw invalidOptionValue(name, value, "true or false");
}
/** `true` for a workbook returned by `readWorkbook` (duck-typed, as in `inspect`). */
function isWorkbook(input) {
	return input !== null && typeof input === "object" && !ArrayBuffer.isView(input) && !(input instanceof ArrayBuffer) && typeof input.getSheet === "function" && Array.isArray(input.sheets);
}
/**
* Accepts bytes or an already read workbook. With `onWarning`, every warning of the workbook is
* passed to it once, so functions that return their own result type (`readTable`, `toCSV`,
* `toJSON`) still let callers see what was skipped or repaired while reading (D-R02).
*/
function openInput(input, limits, onWarning) {
	if (onWarning !== void 0 && typeof onWarning !== "function") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"onWarning\" must be a function.", {
		hint: "For example { onWarning: (w) => console.warn(w.code, w.message) }.",
		details: { option: "onWarning" }
	});
	const workbook = isWorkbook(input) ? input : readWorkbook(input, limits === void 0 ? void 0 : { limits });
	if (onWarning !== void 0) for (const warning of workbook.warnings) onWarning(warning);
	return workbook;
}
export { fillXml as $, metricsOf as $t, checkSource as A, decodeXml as An, parseStyles as At, createWorksheetParser as B, isXlvError as Bn, resolveLimits as Bt, toBytes as C, CORE_PROPERTIES_PART as Cn, relationshipsPartName as Ct, HEAD_BYTES as D, childrenNamed as Dn, checkUniqueName as Dt, CORRUPT_HINT as E, childNamed as En, checkDefinedName as Et, readCentralDirectory as F, escapeAttr as Fn, stableKey as Ft, DEFAULT_FONT as G, hashPassword as Gt, X14_DV_URI as H, printNames as Ht, unsupported as I, escapeText as In, isDateFormat as It, borderXml as J, DRAWING_CONTENT_TYPE as Jt, SheetRels as K, legacyPasswordHash as Kt, validateLocalHeaders as L, ERROR_CODES as Ln, DEFAULT_READ_LIMITS as Lt, limitExceeded as M, parseXml as Mn, checkStyleCount as Mt, locateCentralDirectory as N, clip as Nn, mergeStyles as Nt, NotLoaded as O, parseXmlTree as On, nameScopeKey as Ot, openZip as P, encodeXstring as Pn, normalizeStyle as Pt, createSharedStrings as Q, imageFormat as Qt, WarningCollector as R, WARNING_CODES as Rn, checkKeyText as Rt, styledCellCount as S, APP_PROPERTIES_PART as Sn, openPackage as St, Bytes as T, corePropertiesXml as Tn, normalizeTable as Tt, X14_NS as U, printXml as Ut, X14_CF_URI as V, suggest as Vt, XM_NS as W, sheetProtectionXml as Wt, cellXml as X, MediaParts as Xt, buildStyles as Y, IMAGE_CONTENT_TYPES as Yt, contentTypesXml as Z, drawingParts as Zt, inspect as _, formatRange as _n, FormulaGroups as _t, DATE_FORMAT as a, relationshipTypes as an, reusableBuiltinNumFmt as at, readNotes as b, partKey as bn, parseIsoDateTime as bt, NewSheet as c, chartPartXml as cn, sheetHead as ct, convertCellInput as d, shiftFormula as dn, sheetTail as dt, resolveAnchor as en, fontXml as et, createWorkbook as f, tokenize as fn, stringItemXml as ft, detectHeader as g, formatCellRef as gn, workbookXml as gt, RangeIndex as h, columnLetters as hn, workbookRelsXml as ht, readOptions as i, classifyRelationship as in, noteParts as it, corrupt as j, normalizeLineEnds as jn, READING_ORDERS as jt, TAIL_BYTES as k, createXmlByteStreamParser as kn, EMPTY_STYLES as kt, NewWorkbook as l, chartSheetNames as ln, sheetRelsPartName as lt, writeWorkbook as m, MAX_ROWS as mn, textElement as mt, oneOfOption as n, REL_NS as nn, hyperlinksXml as nt, DATE_TIME_FORMAT as o, CHART_CONTENT_TYPE as on, rootRelsXml as ot, sortedNotes as p, MAX_COLUMNS as pn, tableParts as pt, attr as q, verifyPassword as qt, openInput as r, SML_NS as rn, nameTitles as rt, ERROR_VALUES as s, chartPartName as sn, rowStart as st, booleanOption as t, resolveImage as tn, formatNumber as tt, checkBookViews as u, resolveChartSheets as un, sheetRulesXml as ut, isoDate as v, parseCellRef as vn, writeZip as vt, workbookPartsToLoad as w, appPropertiesXml as wn, TABLE_CONTENT_TYPE as wt, readWorkbook as x, zipNameToPartName as xn, parseSharedStrings as xt, analyzeWorkbook as y, parseRange as yn, crc32 as yt, sheetNotFound as z, XlvError as zn, checkOutputPositions as zt };
