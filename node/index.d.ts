/// <reference types="node" />

export interface ConvertOptions {
  /**
   * Target resolution (pixels per inch) embedded raster images are downsampled
   * to. Defaults to `DEFAULT_IMAGE_DPI` (220, matching Word); non-finite or
   * non-positive values are clamped up to `MIN_IMAGE_DPI`.
   */
  imageDpi?: number;
}

/** Convert DOCX bytes to PDF bytes, off the main thread. */
export function convert(docx: Uint8Array, options?: ConvertOptions): Promise<Buffer>;

/**
 * Convert several DOCX documents into one PDF, in order. Each is laid out as it
 * would be alone (its own page numbers and NUMPAGES); a face they share is
 * embedded once. Rejects with `DocumentError` naming the first document that
 * fails to parse, and with a plain `Error` for an empty batch.
 */
export function convertBatch(documents: Uint8Array[], options?: ConvertOptions): Promise<Buffer>;

export class DocumentError extends Error {
  readonly name: 'DocumentError';
  /** Position of the failing document in the batch. */
  readonly index: number;
}

export const DEFAULT_IMAGE_DPI: number;
export const MIN_IMAGE_DPI: number;
