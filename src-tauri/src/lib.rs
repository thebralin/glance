use std::path::{Path, PathBuf};
use std::sync::Mutex;

use serde::Serialize;
use tauri::{
    ipc::{InvokeBody, Request, Response},
    Emitter, Manager, State,
};

/// Extensions the viewer understands. Everything else passed on the command line is ignored.
const SUPPORTED: &[&str] = &[
    "pdf", "docx", "docm", "dotx", "pptx", "pptm", "ppsx", "potx", "xlsx", "xlsm", "xltx", "csv", "tsv", "txt", "md", "markdown",
    "json",
];

/// Extensions the app is allowed to overwrite (formats with an editor).
const WRITABLE: &[&str] = &["txt", "md", "markdown", "json", "docx", "docm", "dotx"];

/// Files passed on the command line at first launch ("Open with…", double click on an associated file).
#[derive(Default)]
struct StartupFiles(Mutex<Vec<String>>);

#[derive(Serialize)]
struct FileInfo {
    path: String,
    name: String,
    size: u64,
    /// Unix epoch milliseconds.
    modified: Option<u64>,
}

fn has_extension(path: &Path, list: &[&str]) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| list.contains(&e.to_ascii_lowercase().as_str()))
        .unwrap_or(false)
}

fn is_supported(path: &Path) -> bool {
    has_extension(path, SUPPORTED)
}

/// Decode a percent-encoded (encodeURIComponent) header value.
fn percent_decode(s: &str) -> Option<String> {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' {
            let hex = std::str::from_utf8(b.get(i + 1..i + 3)?).ok()?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(b[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// Extract document paths from a process argument list (first element is the executable).
fn files_from_args(args: &[String], cwd: Option<&Path>) -> Vec<String> {
    args.iter()
        .skip(1)
        .filter(|a| !a.starts_with('-'))
        .map(|a| {
            let p = PathBuf::from(a);
            match (p.is_relative(), cwd) {
                (true, Some(cwd)) => cwd.join(p),
                _ => p,
            }
        })
        .filter(|p| p.is_file() && is_supported(p))
        .map(|p| p.to_string_lossy().into_owned())
        .collect()
}

/// Read a whole file and hand it to the webview as raw bytes (ArrayBuffer on the JS side).
#[tauri::command]
async fn read_file(path: String) -> Result<Response, String> {
    let bytes = tauri::async_runtime::spawn_blocking(move || std::fs::read(&path))
        .await
        .map_err(|e| e.to_string())?
        .map_err(|e| e.to_string())?;
    Ok(Response::new(bytes))
}

/// Overwrite an editable file. The body is the raw bytes; the target path comes percent-encoded
/// in the `path` header (headers must be ASCII, paths often are not).
#[tauri::command]
async fn write_file(request: Request<'_>) -> Result<(), String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("ожидались байты файла".into());
    };
    let path = request
        .headers()
        .get("path")
        .and_then(|v| v.to_str().ok())
        .and_then(percent_decode)
        .ok_or("не указан путь")?;
    let path = PathBuf::from(path);
    if !has_extension(&path, WRITABLE) {
        return Err("сохранение этого формата не поддерживается".into());
    }
    let bytes = bytes.clone();
    tauri::async_runtime::spawn_blocking(move || std::fs::write(&path, bytes))
        .await
        .map_err(|e| e.to_string())?
        .map_err(|e| e.to_string())
}

#[tauri::command]
fn file_info(path: String) -> Result<FileInfo, String> {
    let p = PathBuf::from(&path);
    let meta = std::fs::metadata(&p).map_err(|e| e.to_string())?;
    let modified = meta
        .modified()
        .ok()
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_millis() as u64);
    Ok(FileInfo {
        name: p.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default(),
        path,
        size: meta.len(),
        modified,
    })
}

/// Files from the initial launch. Returned once, then cleared.
#[tauri::command]
fn take_startup_files(state: State<'_, StartupFiles>) -> Vec<String> {
    std::mem::take(&mut *state.0.lock().unwrap())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let args: Vec<String> = std::env::args().collect();
    let startup = files_from_args(&args, std::env::current_dir().ok().as_deref());

    let mut builder = tauri::Builder::default();

    // A second launch (e.g. double click on another file) opens the file in the running window.
    #[cfg(desktop)]
    {
        builder = builder.plugin(tauri_plugin_single_instance::init(|app, args, cwd| {
            let files = files_from_args(&args, Some(Path::new(&cwd)));
            if let Some(window) = app.get_webview_window("main") {
                let _ = window.unminimize();
                let _ = window.show();
                let _ = window.set_focus();
            }
            if !files.is_empty() {
                let _ = app.emit("open-files", files);
            }
        }));
    }

    builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(StartupFiles(Mutex::new(startup)))
        .invoke_handler(tauri::generate_handler![read_file, write_file, file_info, take_startup_files])
        .run(tauri::generate_context!())
        .expect("error while running Glance");
}
