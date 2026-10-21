'use strict';

// The native half is `src/node.rs`, built by `scripts/build_node.sh` into
// `dxpdf.node` next to this file. Both conversions run on libuv's thread pool,
// so they never block the event loop and run in parallel up to
// UV_THREADPOOL_SIZE.
const native = require('./dxpdf.node');

/** A document of a `convertBatch` call failed to convert; `index` is its position. */
class DocumentError extends Error {
  constructor(index, message) {
    super(`document ${index}: ${message}`);
    this.name = 'DocumentError';
    this.index = index;
  }
}

function convert(docx, options) {
  return native.convert(docx, options);
}

// The native call resolves to `{ pdf }` or `{ failedIndex, error }` rather than
// rejecting for a document's failure, so the index crosses the boundary as a
// number and not inside a message string.
async function convertBatch(documents, options) {
  const result = await native.convertBatch(documents, options);
  if (result.pdf) return result.pdf;
  throw new DocumentError(result.failedIndex, result.error);
}

// convertBatch for documents given as their parts — { path: string | Uint8Array } —
// rather than as ZIPs, as a templating engine holds them after filling.
async function convertPartsBatch(documents, options) {
  const result = await native.convertPartsBatch(documents, options);
  if (result.pdf) return result.pdf;
  throw new DocumentError(result.failedIndex, result.error);
}

module.exports = {
  convert,
  convertBatch,
  convertPartsBatch,
  DocumentError,
  DEFAULT_IMAGE_DPI: native.DEFAULT_IMAGE_DPI,
  MIN_IMAGE_DPI: native.MIN_IMAGE_DPI,
};
