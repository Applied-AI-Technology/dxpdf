//! Node bindings (`--features node`), loaded by `node/index.js`.
//!
//! Both conversions are [`AsyncTask`]s: the input is copied out while the
//! JavaScript caller still owns it, and the conversion runs on libuv's thread
//! pool, so the event loop never waits on a render and conversions run in
//! parallel up to `UV_THREADPOOL_SIZE` — there is no lock to serialize them on,
//! unlike the Python binding's GIL.
//!
//! A panic is caught at this boundary, as in `src/capi.rs`: unwinding into
//! libuv would abort the whole Node process, where a caller can recover from a
//! rejected promise.
//!
//! napi-derive registers the exports only outside `cfg(test)`, so a test build
//! of the library sees every item here as unused.
#![cfg_attr(test, allow(dead_code))]

use std::panic::{self, AssertUnwindSafe};

use napi::bindgen_prelude::{AsyncTask, Buffer, Either, Env, Error, Result, Task, Uint8Array};
use napi_derive::napi;

/// `node/index.d.ts` documents these for JavaScript callers.
#[napi(object)]
pub struct ConvertOptions {
    pub image_dpi: Option<f64>,
}

/// What [`convert_batch`] resolves to: the PDF, or which document failed and
/// why. A document's failure resolves rather than rejects so its index crosses
/// the boundary as a number; `node/index.js` turns it into a `DocumentError`.
#[napi(object)]
pub struct BatchResult {
    pub pdf: Option<Buffer>,
    pub failed_index: Option<u32>,
    pub error: Option<String>,
}

#[napi]
pub const DEFAULT_IMAGE_DPI: f64 = crate::DEFAULT_IMAGE_DPI as f64;

#[napi]
pub const MIN_IMAGE_DPI: f64 = crate::MIN_IMAGE_DPI as f64;

fn render_options(options: Option<ConvertOptions>) -> crate::RenderOptions {
    let dpi = options
        .and_then(|o| o.image_dpi)
        .unwrap_or(crate::DEFAULT_IMAGE_DPI as f64);
    crate::RenderOptions::default().with_image_dpi(dpi as f32)
}

/// Runs `f`, turning a panic into the error a caught panic reports in
/// `src/capi.rs`.
fn guarded<T>(f: impl FnOnce() -> T) -> Result<T> {
    panic::catch_unwind(AssertUnwindSafe(f)).map_err(|payload| {
        let message = if let Some(s) = payload.downcast_ref::<&str>() {
            (*s).to_string()
        } else if let Some(s) = payload.downcast_ref::<String>() {
            s.clone()
        } else {
            "unknown panic".to_string()
        };
        Error::from_reason(format!("dxpdf: internal error: {message}"))
    })
}

pub struct Convert {
    docx: Vec<u8>,
    options: crate::RenderOptions,
}

impl Task for Convert {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        guarded(|| crate::convert_with_options(&self.docx, &self.options))?
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    fn resolve(&mut self, _env: Env, pdf: Self::Output) -> Result<Self::JsValue> {
        Ok(pdf.into())
    }
}

/// Convert DOCX bytes to PDF bytes.
#[napi]
pub fn convert(docx: Uint8Array, options: Option<ConvertOptions>) -> AsyncTask<Convert> {
    AsyncTask::new(Convert {
        docx: docx.to_vec(),
        options: render_options(options),
    })
}

pub struct ConvertBatch {
    documents: Vec<Vec<u8>>,
    options: crate::RenderOptions,
}

impl Task for ConvertBatch {
    type Output = std::result::Result<Vec<u8>, (usize, String)>;
    type JsValue = BatchResult;

    fn compute(&mut self) -> Result<Self::Output> {
        match guarded(|| crate::convert_batch_with_options(&self.documents, &self.options))? {
            Ok(pdf) => Ok(Ok(pdf)),
            Err(crate::Error::Document { index, source }) => Ok(Err((index, source.to_string()))),
            Err(e) => Err(Error::from_reason(e.to_string())),
        }
    }

    fn resolve(&mut self, env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Self::resolve_output(env, output)
    }
}

impl ConvertBatch {
    fn resolve_output(
        _env: Env,
        output: std::result::Result<Vec<u8>, (usize, String)>,
    ) -> Result<BatchResult> {
        Ok(match output {
            Ok(pdf) => BatchResult {
                pdf: Some(pdf.into()),
                failed_index: None,
                error: None,
            },
            Err((index, error)) => BatchResult {
                pdf: None,
                failed_index: Some(index as u32),
                error: Some(error),
            },
        })
    }
}

pub struct ConvertPartsBatch {
    packages: Vec<crate::docx::zip::PackageContents>,
    options: crate::RenderOptions,
}

impl Task for ConvertPartsBatch {
    type Output = std::result::Result<Vec<u8>, (usize, String)>;
    type JsValue = BatchResult;

    fn compute(&mut self) -> Result<Self::Output> {
        let packages = std::mem::take(&mut self.packages);
        match guarded(|| crate::convert_packages_batch_with_options(packages, &self.options))? {
            Ok(pdf) => Ok(Ok(pdf)),
            Err(crate::Error::Document { index, source }) => Ok(Err((index, source.to_string()))),
            Err(e) => Err(Error::from_reason(e.to_string())),
        }
    }

    fn resolve(&mut self, env: Env, output: Self::Output) -> Result<Self::JsValue> {
        ConvertBatch::resolve_output(env, output)
    }
}

/// [`convert_batch`] for documents given as their parts rather than as ZIPs:
/// each an object from part path to its content, a string for XML or bytes
/// for anything else — what a templating engine holds after filling. See
/// [`crate::convert_packages_batch_with_options`].
#[napi]
pub fn convert_parts_batch(
    documents: Vec<std::collections::HashMap<String, Either<String, Uint8Array>>>,
    options: Option<ConvertOptions>,
) -> AsyncTask<ConvertPartsBatch> {
    let packages = documents
        .into_iter()
        .map(|parts| {
            crate::docx::zip::PackageContents::from_parts(parts.into_iter().map(
                |(path, content)| {
                    let bytes = match content {
                        Either::A(text) => text.into_bytes(),
                        Either::B(bytes) => bytes.to_vec(),
                    };
                    (path, bytes)
                },
            ))
        })
        .collect();
    AsyncTask::new(ConvertPartsBatch {
        packages,
        options: render_options(options),
    })
}

/// Convert several DOCX documents into one PDF, in order — see
/// [`crate::convert_batch_with_options`].
#[napi]
pub fn convert_batch(
    documents: Vec<Uint8Array>,
    options: Option<ConvertOptions>,
) -> AsyncTask<ConvertBatch> {
    AsyncTask::new(ConvertBatch {
        documents: documents.iter().map(|d| d.to_vec()).collect(),
        options: render_options(options),
    })
}
