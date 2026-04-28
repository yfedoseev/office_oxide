/// Map items in parallel (with `parallel` feature) or sequentially.
///
/// This centralizes the `cfg(feature = "parallel")` dispatch used by
/// XLSX worksheet parsing and PPTX slide parsing.
pub fn map_collect<T, R, E, F>(items: Vec<T>, f: F) -> std::result::Result<Vec<R>, E>
where
    T: Send,
    R: Send,
    F: Fn(T) -> std::result::Result<R, E> + Send + Sync,
    E: Send,
{
    #[cfg(feature = "parallel")]
    {
        use rayon::prelude::*;
        items.into_par_iter().map(f).collect()
    }
    #[cfg(not(feature = "parallel"))]
    {
        items.into_iter().map(f).collect()
    }
}
