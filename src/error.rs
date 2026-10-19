use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Parse error: {0}")]
    Parse(#[from] crate::docx::error::ParseError),

    #[error("Render error: {0}")]
    Render(#[from] crate::render::error::RenderError),

    /// One document of a [`crate::convert_batch`] failed; `index` is its
    /// position in the batch, so a caller can drop it and convert the rest.
    #[error("document {index}: {source}")]
    Document {
        index: usize,
        #[source]
        source: Box<Error>,
    },

    /// [`crate::convert_batch`] was given no documents. A PDF with no pages is
    /// not one a reader will open, so there is nothing to return.
    #[error("a batch needs at least one document")]
    EmptyBatch,
}
