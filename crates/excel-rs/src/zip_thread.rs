//! Compresses ZIP entries on a thread of their own, so that the writers can
//! render the next parts meanwhile. Entries are written in the order they
//! are started. If no thread can be started, the work is done inline.

use std::io::Write;
use std::sync::mpsc::{sync_channel, SyncSender};
use std::thread::JoinHandle;

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::{Error, Result};

/// The archive a [`ZipThread`] writes into.
pub(crate) trait Archive: Send + 'static {
    /// What the archive gives back when finished.
    type Output: Send + 'static;
    fn start(&mut self, name: String) -> Result<()>;
    fn data(&mut self, data: &[u8]) -> Result<()>;
    fn finish(self) -> Result<Self::Output>;
}

/// A [`ZipWriter`] with the entry options, as an [`Archive`].
pub(crate) struct ZipArchiveOut<W: Write + std::io::Seek + Send + 'static> {
    pub zip: ZipWriter<W>,
    pub options: SimpleFileOptions,
}

impl<W: Write + std::io::Seek + Send + 'static> Archive for ZipArchiveOut<W> {
    type Output = W;
    fn start(&mut self, name: String) -> Result<()> {
        Ok(self.zip.start_file(name, self.options)?)
    }
    fn data(&mut self, data: &[u8]) -> Result<()> {
        Ok(self.zip.write_all(data)?)
    }
    fn finish(self) -> Result<W> {
        Ok(self.zip.finish()?)
    }
}

/// A streaming [`ZipWriter`] (no seeking: data descriptors) as an
/// [`Archive`].
pub(crate) struct ZipStreamOut<W: Write + Send + 'static> {
    pub zip: ZipWriter<zip::write::StreamWriter<W>>,
    pub options: SimpleFileOptions,
}

impl<W: Write + Send + 'static> Archive for ZipStreamOut<W> {
    type Output = ();
    fn start(&mut self, name: String) -> Result<()> {
        Ok(self.zip.start_file(name, self.options)?)
    }
    fn data(&mut self, data: &[u8]) -> Result<()> {
        Ok(self.zip.write_all(data)?)
    }
    fn finish(self) -> Result<()> {
        self.zip.finish()?;
        Ok(())
    }
}

enum Msg {
    Start(String),
    Data(Vec<u8>),
}

/// An [`Archive`] written by another thread.
pub(crate) struct ZipThread<A: Archive> {
    tx: Option<SyncSender<Msg>>,
    thread: Option<JoinHandle<Result<A::Output>>>,
    /// Used when no thread could be started.
    inline: Option<A>,
}

impl<A: Archive> ZipThread<A> {
    pub fn new(archive: A) -> Self {
        // A few megabytes in flight at most.
        let (tx, rx) = sync_channel::<Msg>(8);
        // Shared so that it can be taken back if no thread starts.
        let slot = std::sync::Arc::new(std::sync::Mutex::new(Some(archive)));
        let theirs = std::sync::Arc::clone(&slot);
        let spawned = std::thread::Builder::new()
            .name("excel-rs zip".into())
            .spawn(move || {
                let taken = theirs.lock().ok().and_then(|mut a| a.take());
                let mut a = taken.ok_or_else(stopped)?;
                for msg in rx {
                    match msg {
                        Msg::Start(name) => a.start(name)?,
                        Msg::Data(data) => a.data(&data)?,
                    }
                }
                a.finish()
            });
        match spawned {
            Ok(thread) => ZipThread {
                tx: Some(tx),
                thread: Some(thread),
                inline: None,
            },
            Err(_) => ZipThread {
                tx: None,
                thread: None,
                inline: slot.lock().ok().and_then(|mut a| a.take()),
            },
        }
    }

    fn send(&mut self, msg: Msg) -> Result<()> {
        if let Some(a) = self.inline.as_mut() {
            return match msg {
                Msg::Start(name) => a.start(name),
                Msg::Data(data) => a.data(&data),
            };
        }
        if let Some(tx) = &self.tx {
            if tx.send(msg).is_ok() {
                return Ok(());
            }
        }
        // the thread stopped: its error
        self.tx = None;
        match self.join() {
            Err(e) => Err(e),
            Ok(_) => Err(stopped()),
        }
    }

    fn join(&mut self) -> Result<A::Output> {
        match self.thread.take().map(JoinHandle::join) {
            Some(Ok(result)) => result,
            _ => Err(stopped()),
        }
    }

    /// Starts the next entry.
    pub fn start(&mut self, name: String) -> Result<()> {
        self.send(Msg::Start(name))
    }

    /// Appends to the current entry.
    pub fn data(&mut self, data: Vec<u8>) -> Result<()> {
        if data.is_empty() {
            return Ok(());
        }
        self.send(Msg::Data(data))
    }

    /// Ends the archive, once everything sent is written.
    pub fn finish(mut self) -> Result<A::Output> {
        if let Some(a) = self.inline.take() {
            return a.finish();
        }
        self.tx = None;
        self.join()
    }
}

fn stopped() -> Error {
    Error::invalid("zip", "the compressing thread stopped")
}
