// Demo: fill a Word, PowerPoint or Excel template in the browser with docvane. Text from the template
// and from the data is only ever inserted with textContent, never as HTML, and it is cut to a few
// hundred characters, so a template with huge strings cannot freeze the page while it is laid out.
import { inspect, isDvError, preview, render, validate } from '../vendor/docvane/index.js';

const $ = (id) => document.getElementById(id);
/** Longest text shown in one table cell or message. */
const MAX_TEXT = 300;
/** Tags and problems listed on the page. */
const MAX_ROWS = 500;
/** Longest preview text shown. */
const MAX_PREVIEW = 200_000;
/** A transparent 1 × 1 PNG, the placeholder for images in data made from a schema. */
const PIXEL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const TYPES = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};
let current = null;
/** Placeholders in other notations, looked for in a document without tags: [start, end, name]. */
const NOTATIONS = [
  ['{{', '}}'],
  ['[[', ']]'],
  ['<<', '>>'],
  ['${', '}'],
  ['«', '»'],
];

/** The delimiters of the two fields, or nothing for the default { and }. */
function delimiterOptions() {
  const start = $('start').value.trim();
  const end = $('end').value.trim();
  return start === '{' && end === '}' ? {} : { delimiters: { start, end } };
}

const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** A sentence about placeholders in another notation that the text of a document seems to have, or null. */
function notationHint(bytes) {
  let text = '';
  try {
    text = preview(bytes).text;
  } catch {
    return null;
  }
  const current = delimiterOptions().delimiters ?? { start: '{', end: '}' };
  for (const [start, end] of NOTATIONS) {
    if (start === current.start && end === current.end) continue;
    const found = text.match(new RegExp(`${escapeRegExp(start)}\\s*[\\p{L}_][\\p{L}\\p{N}_.]{0,60}\\s*${escapeRegExp(end)}`, 'gu'));
    if (found) {
      const examples = [...new Set(found)].slice(0, 3).join(', ');
      return `Its text has ${String(found.length)} placeholder${found.length === 1 ? '' : 's'} such as ${examples}. To use them as tags, set the delimiters above to ${start} and ${end}.`;
    }
  }
  return null;
}

/** Says that a document has no tags, with the placeholders in another notation that its text seems to have. */
function showNoTags(bytes) {
  const box = $('notags');
  box.replaceChildren(el('p', 'This document has no tags, so there is nothing to fill: the result would be the same document. Tags are text between the delimiters, such as {name} for a value or {#items}…{/items} for a list, typed into the document in Word, PowerPoint, Excel or LibreOffice.'));
  const hint = notationHint(bytes);
  if (hint !== null) box.append(el('p', hint));
  const link = el('a', 'Read the template syntax');
  link.href = '../guide/syntax.html';
  const p = el('p');
  p.append(link, document.createTextNode(', or try the sample offer to see a template with tags.'));
  box.append(p);
  box.hidden = false;
}

function el(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = cut(text);
  if (className) node.className = className;
  return node;
}

function cut(text, max = MAX_TEXT) {
  const s = String(text);
  return s.length > max ? `${s.slice(0, max)}…` : s;
}

function setStatus(text) {
  $('status').textContent = cut(text, 400);
}

/** One problem: what, where (in the data and in the template) and what to do. */
function problem(error) {
  const item = el('li');
  item.append(el('span', `${String(error.code ?? 'Error')}: ${error instanceof Error ? error.message : String(error)}`));
  const dataPath = error?.details?.dataPath;
  const place = error?.location?.description;
  const where = [dataPath === undefined || dataPath === '' ? '' : `${String(dataPath)} in the data`, place ?? ''].filter((s) => s !== '').join('; ');
  if (where !== '') item.append(el('span', `Where: ${where}`, 'where'));
  // Hints about options (such as raising a limit) are for programs; the demo uses the defaults.
  if (error?.hint && error.code !== 'DV_E_LIMIT_EXCEEDED') item.append(el('span', `Fix: ${String(error.hint)}`, 'fix'));
  return item;
}

function showError(error, keepResult = false) {
  const box = $('error');
  box.replaceChildren();
  const code = isDvError(error) ? error.code : 'Error';
  box.append(el('p', `${code}: ${error instanceof Error ? error.message : String(error)}`));
  if (isDvError(error) && error.errors.length > 0) {
    const list = el('ol', undefined, 'problems');
    for (const e of error.errors.slice(0, MAX_ROWS)) list.append(problem(e));
    box.append(list);
  } else if (isDvError(error) && error.hint && error.code !== 'DV_E_LIMIT_EXCEEDED') box.append(el('p', `Hint: ${error.hint}`));
  box.hidden = false;
  if (!keepResult) $('result').hidden = true;
}

/** Lets the browser paint the status line before the template is read on the main thread. */
const paint = () => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));

/** Data with the shape of a schema: empty text (0, a date or an address where a formatter needs one), false, one item per list, a placeholder per image. */
function skeleton(schema, hints, name = '', depth = 0) {
  if (!schema || depth > 12) return '';
  if (schema.anyOf) {
    const image = schema.anyOf.some((s) => s.type === 'object' && s.properties?.data);
    return image ? PIXEL : skeleton(schema.anyOf[0], hints, name, depth + 1);
  }
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  if (types[0] === 'object') return Object.fromEntries(Object.entries(schema.properties ?? {}).map(([k, v]) => [k, skeleton(v, hints, k, depth + 1)]));
  if (types[0] === 'array') return [skeleton(schema.items, hints, name, depth + 1)];
  if (types.includes('number') && (types[0] === 'number' || hints.numbers.has(name))) return 0;
  if (hints.dates.has(name)) return new Date().toISOString().slice(0, 10);
  if (hints.links.has(name)) return 'https://example.com/';
  if (types[0] === 'boolean') return false;
  return '';
}

/** Names of fields shown with a number, date or link formatter, which need such values rather than empty text. */
function formatterHints(report) {
  const numbers = new Set();
  const dates = new Set();
  const links = new Set();
  for (const tag of report.tags) {
    const names = tag.formatters.map((f) => f.name);
    const last = tag.path.split('.').pop() ?? '';
    if (names.some((n) => n === 'number' || n === 'currency' || n === 'percent')) numbers.add(last);
    if (names.includes('date')) dates.add(last);
    if (names.includes('link')) links.add(last);
  }
  return { numbers, dates, links };
}

async function open(bytes, name, data) {
  $('error').hidden = true;
  $('problems').hidden = true;
  $('output').hidden = true;
  $('notags').hidden = true;
  setStatus(`Reading ${name}…`);
  await paint();
  const started = performance.now();
  try {
    const report = inspect(bytes, delimiterOptions());
    current = { name, bytes, report };
    if (report.tags.length === 0) showNoTags(bytes);
    $('data').value = JSON.stringify(data ?? skeleton(report.schema, formatterHints(report)), null, 2);
    show(report);
    setStatus(`${name}: ${report.format.toUpperCase()}, ${String(report.tags.length)} tags, read in ${String(Math.round(performance.now() - started))} ms.`);
  } catch (error) {
    current = null;
    setStatus(`${name}: could not be read.`);
    showError(error);
    // Placeholders such as {{name}} are not tags with the delimiters { and }; say which delimiters would read them.
    const hint = isDvError(error) && error.code === 'DV_E_TEMPLATE_INVALID' ? notationHint(bytes) : null;
    if (hint !== null) {
      const box = $('notags');
      box.replaceChildren(el('p', hint));
      box.hidden = false;
      current = { name, bytes, report: null };
    }
  }
}

function show(report) {
  const body = $('tags').tBodies[0];
  body.replaceChildren();
  for (const tag of report.tags.slice(0, MAX_ROWS)) {
    const row = el('tr');
    const raw = el('td');
    raw.append(el('code', tag.raw));
    row.append(raw, el('td', tag.kind), el('td', tag.location.description ?? tag.location.part ?? ''));
    body.append(row);
  }
  const warnings = $('warnings');
  warnings.replaceChildren(...report.warnings.slice(0, 100).map((w) => el('p', `${w.code}: ${w.message}`)));
  warnings.hidden = report.warnings.length === 0;
  $('schema').textContent = cut(JSON.stringify(report.schema, null, 2), 100_000);
  $('result').hidden = false;
}

/** The data of the text box, or null after showing why it is not JSON. */
function data() {
  try {
    return JSON.parse($('data').value);
  } catch (error) {
    showProblems(`The data is not JSON: ${error instanceof Error ? error.message : String(error)}`, []);
    return null;
  }
}

function showProblems(title, errors, ok = false) {
  const box = $('problems');
  box.className = `notice ${ok ? 'ok' : 'error'}`;
  box.replaceChildren(el('p', title));
  if (errors.length > 0) {
    const list = el('ol', undefined, 'problems');
    for (const e of errors.slice(0, MAX_ROWS)) list.append(problem(e));
    box.append(list);
  }
  box.hidden = false;
}

function check() {
  if (!current?.report) return;
  if (current.report.tags.length === 0) {
    showProblems('Nothing to check: the document has no tags.', []);
    return;
  }
  const value = data();
  if (value === null) return;
  try {
    const result = validate(current.bytes, value, { locale: $('locale').value, ...delimiterOptions() });
    if (result.valid) showProblems('Valid: the data has everything the template needs.', [], true);
    else showProblems(`The data has ${String(result.errors.length)} problem${result.errors.length === 1 ? '' : 's'}:`, result.errors);
  } catch (error) {
    showError(error, true);
  }
}

function fill() {
  if (!current?.report) return;
  if (current.report.tags.length === 0) {
    showProblems('Nothing to fill: the document has no tags, so the result would be the same document.', []);
    return;
  }
  const value = data();
  if (value === null) return;
  $('problems').hidden = true;
  try {
    const started = performance.now();
    const result = render(current.bytes, value, { locale: $('locale').value, ...delimiterOptions() });
    const format = current.report.format;
    const name = current.name.replace(/\.[^.]*$/, '') + `-filled.${format}`;
    const url = URL.createObjectURL(new Blob([result], { type: TYPES[format] }));
    const link = el('a');
    link.href = url;
    link.download = name;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
    $('preview').textContent = cut(preview(result).text, MAX_PREVIEW);
    $('output').hidden = false;
    setStatus(`${name}: ${(result.byteLength / 1024).toFixed(1)} KiB, made in ${String(Math.round(performance.now() - started))} ms and downloaded.`);
  } catch (error) {
    if (isDvError(error) && error.errors.length > 0) showProblems(`${error.code}: ${error.message}`, error.errors);
    else showError(error, true);
  }
}

async function sample() {
  setStatus('Loading the sample…');
  try {
    const get = async (path) => {
      const response = await fetch(path);
      if (!response.ok) throw new Error(`${path}: HTTP ${String(response.status)}`);
      return new Uint8Array(await response.arrayBuffer());
    };
    const [template, json, logo] = await Promise.all([get('sample/offer.docx'), get('sample/data.json'), get('sample/logo.png')]);
    const values = JSON.parse(new TextDecoder().decode(json));
    // In data.json the logo is a file name for the command line; the library takes a data URL.
    let binary = '';
    for (const b of logo) binary += String.fromCharCode(b);
    values.logo = `data:image/png;base64,${btoa(binary)}`;
    $('start').value = '{';
    $('end').value = '}';
    await open(template, 'offer.docx', values);
  } catch (error) {
    setStatus('The sample could not be loaded.');
    showError(error);
  }
}

async function openFile(file) {
  if (!file) return;
  await open(new Uint8Array(await file.arrayBuffer()), file.name);
}

$('file').addEventListener('change', () => void openFile($('file').files?.[0]));
$('sample').addEventListener('click', () => void sample());
$('validate').addEventListener('click', check);
for (const id of ['start', 'end']) {
  $(id).addEventListener('change', () => {
    if (current) void open(current.bytes, current.name);
  });
}
$('render').addEventListener('click', fill);
const drop = $('drop');
drop.addEventListener('dragover', (e) => {
  e.preventDefault();
  drop.classList.add('over');
});
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('over');
  void openFile(e.dataTransfer?.files?.[0]);
});
