// The Node binding: `npm test` in node/ after `scripts/build_node.sh`.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';

const dxpdf = createRequire(import.meta.url)('../index.js');
const fixture = (name) => readFileSync(new URL(`../../test-files/${name}`, import.meta.url));

const small = fixture('sample-docx-files-sample-4.docx');
const other = fixture('oversized-word.docx');
const large = fixture('sample-docx-files-sample4.docx');

/** Page objects in a PDF (`/Type /Page`, not the `/Pages` tree nodes). */
const pages = (pdf) => pdf.toString('latin1').match(/\/Type\s*\/Page(?![s\w])/g).length;

test('convert resolves to the PDF of one document', async () => {
  const pdf = await dxpdf.convert(small);
  assert.ok(Buffer.isBuffer(pdf));
  assert.equal(pdf.subarray(0, 5).toString(), '%PDF-');
});

test('convert rejects what is not a DOCX', async () => {
  await assert.rejects(dxpdf.convert(Buffer.from('not a zip archive')), Error);
});

test('convertBatch puts every document in one PDF, in order of pages', async () => {
  const [a, b] = await Promise.all([dxpdf.convert(small), dxpdf.convert(other)]);
  const batch = await dxpdf.convertBatch([small, other, small]);
  assert.equal(pages(batch), 2 * pages(a) + pages(b));
});

test('a batch of one is the single conversion', async () => {
  assert.deepEqual(await dxpdf.convertBatch([small]), await dxpdf.convert(small));
});

test('a document that fails is named by its index', async () => {
  await assert.rejects(
    dxpdf.convertBatch([small, small, Buffer.from('garbage')]),
    (err) => err instanceof dxpdf.DocumentError && err.index === 2 && /^document 2: /.test(err.message),
  );
});

test('an empty batch is refused, and not as a document error', async () => {
  await assert.rejects(
    dxpdf.convertBatch([]),
    (err) => !(err instanceof dxpdf.DocumentError) && /at least one document/.test(err.message),
  );
});

test('imageDpi reaches the engine', async () => {
  const coarse = await dxpdf.convert(large, { imageDpi: 20 });
  const fine = await dxpdf.convert(large, { imageDpi: dxpdf.DEFAULT_IMAGE_DPI });
  assert.ok(coarse.length < fine.length, `${coarse.length} >= ${fine.length}`);
  assert.equal(dxpdf.DEFAULT_IMAGE_DPI, 220);
  assert.ok(dxpdf.MIN_IMAGE_DPI > 0);
});

test('a conversion leaves the event loop free', async () => {
  let ticks = 0;
  const timer = setInterval(() => ticks++, 1);
  await dxpdf.convert(large);
  clearInterval(timer);
  assert.ok(ticks > 0, 'the event loop ran while the document converted');
});

test('conversions run concurrently and each gets its own result', async () => {
  const inputs = [small, other, small, other, small, other];
  const results = await Promise.all(inputs.map((d) => dxpdf.convert(d)));
  const [a, b] = [await dxpdf.convert(small), await dxpdf.convert(other)];
  results.forEach((pdf, i) => assert.deepEqual(pdf, i % 2 ? b : a));
});

// convertPartsBatch: the same documents given as their parts, as a templating
// engine holds them after filling, rather than as ZIPs. XML parts may be
// strings; anything may be bytes.
const minimal = (text) => ({
  '[Content_Types].xml':
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  '_rels/.rels':
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  'word/document.xml': Buffer.from(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
  ),
});

test('convertPartsBatch renders documents given as their parts, in order', async () => {
  const pdf = await dxpdf.convertPartsBatch([minimal('first'), minimal('second')]);
  assert.equal(pdf.subarray(0, 5).toString(), '%PDF-');
  assert.equal(pages(pdf), 2);
});

test('a package given as parts that does not parse is named by its index', async () => {
  const { '_rels/.rels': _, ...broken } = minimal('broken');
  await assert.rejects(
    dxpdf.convertPartsBatch([minimal('ok'), broken]),
    (err) => err instanceof dxpdf.DocumentError && err.index === 1,
  );
});

test('an empty batch of parts is refused, and not as a document error', async () => {
  await assert.rejects(
    dxpdf.convertPartsBatch([]),
    (err) => !(err instanceof dxpdf.DocumentError) && /at least one document/.test(err.message),
  );
});
