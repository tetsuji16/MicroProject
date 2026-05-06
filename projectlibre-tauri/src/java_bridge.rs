use base64::Engine;
use serde::Serialize;
use serde_json::Value;
use std::env;
use std::fs;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::Mutex;
use std::time::SystemTime;

#[derive(Debug, Clone, Serialize)]
pub struct JavaBridgeStatus {
    pub running: bool,
    pub compiled: bool,
    pub backend_mode: String,
    pub uses_java: bool,
    pub java_executable: Option<String>,
    pub javac_executable: Option<String>,
    pub bridge_source: String,
    pub class_dir: String,
    pub sample_files: Vec<String>,
    pub opened_project: Option<String>,
    pub last_command: Option<String>,
    pub last_response: Option<Value>,
    pub last_error: Option<String>,
}

#[derive(Clone)]
pub struct JavaBridgeState {
    config: JavaBridgeConfig,
    runtime: std::sync::Arc<Mutex<Option<JavaBridgeRuntime>>>,
    rust_runtime: std::sync::Arc<Mutex<RustBridgeRuntime>>,
}

#[derive(Clone)]
struct JavaBridgeConfig {
    bridge_source: PathBuf,
    class_dir: PathBuf,
    java_executable: PathBuf,
    javac_executable: PathBuf,
    sample_files: Vec<PathBuf>,
    prefer_java_bridge: bool,
}

struct JavaBridgeRuntime {
    child: Child,
    stdin: ChildStdin,
    stdout: BufReader<ChildStdout>,
    compiled: bool,
    opened_project: Option<String>,
    last_command: Option<String>,
    last_response: Option<Value>,
    last_error: Option<String>,
}

struct RustBridgeRuntime {
    opened_project: Option<String>,
    last_command: Option<String>,
    last_response: Option<Value>,
    last_error: Option<String>,
    sample_files: Vec<String>,
}

impl RustBridgeRuntime {
    fn new(sample_files: Vec<String>) -> Self {
        Self {
            opened_project: None,
            last_command: None,
            last_response: None,
            last_error: None,
            sample_files,
        }
    }
}

enum BridgeBackend<'a> {
    Java(&'a mut JavaBridgeRuntime),
    Rust(&'a mut RustBridgeRuntime),
}

impl<'a> BridgeBackend<'a> {
    fn ping(&mut self) -> Result<Value, String> {
        match self {
            BridgeBackend::Java(runtime) => runtime.send("ping", &[]),
            BridgeBackend::Rust(runtime) => {
                runtime.last_command = Some("ping".to_string());
                runtime.last_error = None;
                let payload = serde_json::json!({
                    "message": "bridge ready",
                    "backend": "rust",
                });
                runtime.last_response = Some(payload.clone());
                Ok(payload)
            }
        }
    }

    fn snapshot(&mut self) -> Result<Value, String> {
        match self {
            BridgeBackend::Java(runtime) => runtime.send("snapshot", &[]),
            BridgeBackend::Rust(runtime) => {
                runtime.last_command = Some("snapshot".to_string());
                runtime.last_error = None;
                let payload = serde_json::json!({
                    "opened_project": runtime.opened_project.clone(),
                    "last_command": runtime.last_command.clone(),
                    "sample_files": runtime.sample_files.clone(),
                    "bridge_mode": "rust-first",
                });
                runtime.last_response = Some(payload.clone());
                Ok(payload)
            }
        }
    }

    fn open(&mut self, path: &str) -> Result<Value, String> {
        match self {
            BridgeBackend::Java(runtime) => runtime.send("open", &[path.to_owned()]),
            BridgeBackend::Rust(runtime) => {
                runtime.last_command = Some("open".to_string());
                runtime.opened_project = Some(path.to_string());
                runtime.last_error = None;
                let payload = serde_json::json!({
                    "opened_project": path,
                    "opened_name": Path::new(path).file_name().and_then(|value| value.to_str()).unwrap_or(""),
                    "sample_files": runtime.sample_files.clone(),
                    "bridge_mode": "rust-first",
                });
                runtime.last_response = Some(payload.clone());
                Ok(payload)
            }
        }
    }

    fn import_mpp(&mut self, path: &str) -> Result<Value, String> {
        match self {
            BridgeBackend::Java(runtime) => runtime.send("import_mpp", &[path.to_owned()]),
            BridgeBackend::Rust(runtime) => {
                runtime.last_command = Some("import_mpp".to_string());
                runtime.opened_project = Some(path.to_string());
                runtime.last_error = None;
                let payload = serde_json::json!({
                    "opened_project": path,
                    "imported": true,
                    "opened_name": Path::new(path).file_name().and_then(|value| value.to_str()).unwrap_or(""),
                    "sample_files": runtime.sample_files.clone(),
                    "bridge_mode": "rust-first",
                });
                runtime.last_response = Some(payload.clone());
                Ok(payload)
            }
        }
    }

    fn export_mpp(&mut self, path: &str) -> Result<Value, String> {
        match self {
            BridgeBackend::Java(runtime) => runtime.send("export_mpp", &[path.to_owned()]),
            BridgeBackend::Rust(runtime) => {
                runtime.last_command = Some("export_mpp".to_string());
                if runtime.opened_project.is_none() {
                    runtime.last_error = Some("No project is open".to_string());
                    return Err("No project is open".to_string());
                }
                runtime.last_error = None;
                let source_project = runtime.opened_project.clone();
                let opened_name = source_project
                    .as_ref()
                    .and_then(|value| Path::new(value).file_name().and_then(|name| name.to_str()))
                    .unwrap_or("")
                    .to_string();
                let payload = serde_json::json!({
                    "source_project": source_project,
                    "target_path": path,
                    "opened_name": opened_name,
                    "sample_files": runtime.sample_files.clone(),
                    "bridge_mode": "rust-first",
                });
                runtime.last_response = Some(payload.clone());
                Ok(payload)
            }
        }
    }
}

#[derive(Debug, serde::Deserialize)]
struct BridgeEnvelope {
    ok: bool,
    command: String,
    #[serde(default)]
    data: Option<Value>,
    #[serde(default)]
    error: Option<String>,
    #[serde(default)]
    message: Option<String>,
}

impl JavaBridgeState {
    pub fn new() -> Self {
        let config = JavaBridgeConfig::discover();
        let rust_sample_files = config
            .sample_files
            .iter()
            .map(|path| path.display().to_string())
            .collect();
        Self {
            config,
            runtime: std::sync::Arc::new(Mutex::new(None)),
            rust_runtime: std::sync::Arc::new(Mutex::new(RustBridgeRuntime::new(
                rust_sample_files,
            ))),
        }
    }

    pub fn status(&self) -> JavaBridgeStatus {
        let mut java_guard = self.runtime.lock().unwrap();
        let java_running = java_guard
            .as_mut()
            .map(|runtime| runtime.is_alive())
            .unwrap_or(false);
        let java_runtime = java_guard.as_ref();
        let rust_runtime = self.rust_runtime.lock().unwrap();
        let backend_mode = if self.config.prefer_java_bridge && java_running {
            "java"
        } else {
            "rust"
        };
        JavaBridgeStatus {
            running: if self.config.prefer_java_bridge {
                java_running
            } else {
                true
            },
            compiled: if self.config.prefer_java_bridge {
                java_runtime
                    .map(|runtime| runtime.compiled)
                    .unwrap_or_else(|| self.config.is_compiled())
            } else {
                true
            },
            backend_mode: backend_mode.to_string(),
            uses_java: self.config.prefer_java_bridge && java_running,
            java_executable: if self.config.prefer_java_bridge {
                Some(self.config.java_executable.display().to_string())
            } else {
                None
            },
            javac_executable: Some(self.config.javac_executable.display().to_string()),
            bridge_source: self.config.bridge_source.display().to_string(),
            class_dir: self.config.class_dir.display().to_string(),
            sample_files: self
                .config
                .sample_files
                .iter()
                .map(|path| path.display().to_string())
                .collect(),
            opened_project: if self.config.prefer_java_bridge && java_running {
                java_runtime.and_then(|runtime| runtime.opened_project.clone())
            } else {
                rust_runtime.opened_project.clone()
            },
            last_command: if self.config.prefer_java_bridge && java_running {
                java_runtime.and_then(|runtime| runtime.last_command.clone())
            } else {
                rust_runtime.last_command.clone()
            },
            last_response: if self.config.prefer_java_bridge && java_running {
                java_runtime.and_then(|runtime| runtime.last_response.clone())
            } else {
                rust_runtime.last_response.clone()
            },
            last_error: if self.config.prefer_java_bridge && java_running {
                java_runtime.and_then(|runtime| runtime.last_error.clone())
            } else {
                rust_runtime.last_error.clone()
            },
        }
    }

    pub fn ping(&self) -> Result<Value, String> {
        self.with_backend(|backend| backend.ping())
    }

    pub fn snapshot(&self) -> Result<Value, String> {
        self.with_backend(|backend| backend.snapshot())
    }

    pub fn open_mpp(&self, path: &str) -> Result<Value, String> {
        self.with_backend(|backend| backend.open(path))
    }

    pub fn import_mpp(&self, path: &str) -> Result<Value, String> {
        self.with_backend(|backend| backend.import_mpp(path))
    }

    pub fn export_mpp(&self, path: &str) -> Result<Value, String> {
        self.with_backend(|backend| backend.export_mpp(path))
    }

    fn with_backend<T>(
        &self,
        f: impl FnOnce(&mut BridgeBackend<'_>) -> Result<T, String>,
    ) -> Result<T, String> {
        if self.config.prefer_java_bridge {
            let mut guard = self
                .runtime
                .lock()
                .map_err(|_| "java bridge lock poisoned".to_string())?;
            let running = guard
                .as_mut()
                .map(|runtime| runtime.is_alive())
                .unwrap_or(false);
            if !running {
                *guard = Some(JavaBridgeRuntime::launch(&self.config)?);
            }
            let runtime = guard.as_mut().expect("runtime exists after launch");
            return f(&mut BridgeBackend::Java(runtime));
        }

        let mut guard = self
            .rust_runtime
            .lock()
            .map_err(|_| "rust bridge lock poisoned".to_string())?;
        f(&mut BridgeBackend::Rust(&mut *guard))
    }
}

impl JavaBridgeConfig {
    fn discover() -> Self {
        let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let bridge_source = manifest_dir
            .join("java-bridge")
            .join("src")
            .join("ProjectLibreBridge.java");
        let class_dir = manifest_dir
            .join("target")
            .join("java-bridge")
            .join("classes");
        let java_executable = locate_executable("java").unwrap_or_else(|| PathBuf::from("java"));
        let javac_executable = locate_javac().unwrap_or_else(|| PathBuf::from("javac"));
        let sample_files = discover_sample_files(&manifest_dir);
        let prefer_java_bridge = env::var("MICROPROJECT_USE_JAVA_BRIDGE")
            .map(|value| matches!(value.as_str(), "1" | "true" | "yes" | "on"))
            .unwrap_or(false);

        Self {
            bridge_source,
            class_dir,
            java_executable,
            javac_executable,
            sample_files,
            prefer_java_bridge,
        }
    }

    fn is_compiled(&self) -> bool {
        let class_file = self.class_dir.join("ProjectLibreBridge.class");
        let source_modified = metadata_modified(&self.bridge_source);
        let class_modified = metadata_modified(&class_file);

        match (source_modified, class_modified) {
            (Some(source), Some(class)) => class >= source,
            (_, Some(_)) => true,
            _ => false,
        }
    }
}

impl JavaBridgeRuntime {
    fn launch(config: &JavaBridgeConfig) -> Result<Self, String> {
        compile_bridge(config)?;
        fs::create_dir_all(&config.class_dir)
            .map_err(|error| format!("Unable to create Java bridge class dir: {error}"))?;

        let mut command = Command::new(&config.java_executable);
        command
            .arg("-cp")
            .arg(&config.class_dir)
            .arg("ProjectLibreBridge")
            .args(config.sample_files.iter().map(|path| path.as_os_str()))
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit());

        let mut child = command
            .spawn()
            .map_err(|error| format!("Unable to launch Java bridge: {error}"))?;
        let stdin = child
            .stdin
            .take()
            .ok_or_else(|| "Java bridge stdin is unavailable".to_string())?;
        let stdout = child
            .stdout
            .take()
            .ok_or_else(|| "Java bridge stdout is unavailable".to_string())?;
        let mut runtime = Self {
            child,
            stdin,
            stdout: BufReader::new(stdout),
            compiled: true,
            opened_project: None,
            last_command: None,
            last_response: None,
            last_error: None,
        };

        let ready = runtime.read_response()?;
        if !ready.ok {
            return Err(ready
                .error
                .unwrap_or_else(|| "Java bridge reported a startup error".to_string()));
        }

        runtime.last_command = Some(ready.command);
        runtime.last_response = ready
            .data
            .clone()
            .or_else(|| ready.message.map(Value::String));
        Ok(runtime)
    }

    fn is_alive(&mut self) -> bool {
        match self.child.try_wait() {
            Ok(Some(_)) => false,
            Ok(None) => true,
            Err(_) => false,
        }
    }

    fn send(&mut self, command: &str, args: &[String]) -> Result<Value, String> {
        let line = encode_command(command, args);
        self.stdin
            .write_all(line.as_bytes())
            .map_err(|error| format!("Unable to send {command} to Java bridge: {error}"))?;
        self.stdin
            .write_all(b"\n")
            .and_then(|_| self.stdin.flush())
            .map_err(|error| format!("Unable to flush {command} to Java bridge: {error}"))?;

        let envelope = self.read_response()?;
        self.last_command = Some(envelope.command.clone());
        self.last_response = envelope
            .data
            .clone()
            .or_else(|| envelope.message.clone().map(Value::String));
        if !envelope.ok {
            self.last_error = envelope.error.clone();
            return Err(envelope
                .error
                .unwrap_or_else(|| format!("Java bridge command {command} failed")));
        }
        if command == "open" || command == "import_mpp" {
            if let Some(path) = args.first() {
                self.opened_project = Some(path.clone());
            }
        }
        Ok(envelope.data.unwrap_or_else(|| Value::Null))
    }

    fn read_response(&mut self) -> Result<BridgeEnvelope, String> {
        let mut line = String::new();
        let bytes = self
            .stdout
            .read_line(&mut line)
            .map_err(|error| format!("Unable to read Java bridge response: {error}"))?;
        if bytes == 0 {
            return Err("Java bridge exited before responding".to_string());
        }
        serde_json::from_str(line.trim()).map_err(|error| {
            format!("Unable to parse Java bridge response: {error}; line={line:?}")
        })
    }
}

fn compile_bridge(config: &JavaBridgeConfig) -> Result<(), String> {
    fs::create_dir_all(&config.class_dir)
        .map_err(|error| format!("Unable to create Java bridge class dir: {error}"))?;
    if config.is_compiled() {
        return Ok(());
    }

    let status = Command::new(&config.javac_executable)
        .arg("-source")
        .arg("8")
        .arg("-target")
        .arg("8")
        .arg("-Xlint:-options")
        .arg("-encoding")
        .arg("UTF-8")
        .arg("-d")
        .arg(&config.class_dir)
        .arg(&config.bridge_source)
        .status()
        .map_err(|error| format!("Unable to invoke javac: {error}"))?;

    if !status.success() {
        return Err(format!(
            "javac failed for {}",
            config.bridge_source.display()
        ));
    }

    Ok(())
}

fn encode_command(command: &str, args: &[String]) -> String {
    let mut parts = vec![command.to_string()];
    parts.extend(args.iter().map(|arg| base64_encode(arg)));
    parts.join("\t")
}

fn base64_encode(value: &str) -> String {
    base64::engine::general_purpose::STANDARD.encode(value.as_bytes())
}

fn locate_executable(name: &str) -> Option<PathBuf> {
    if let Ok(path) = locate_with_where(name) {
        return Some(path);
    }

    let mut candidates = Vec::new();
    if let Ok(java_home) = env::var("JAVA_HOME") {
        candidates.push(
            PathBuf::from(java_home)
                .join("bin")
                .join(executable_name(name)),
        );
    }

    if let Ok(program_files) = env::var("ProgramFiles") {
        candidates.push(
            PathBuf::from(&program_files)
                .join("Java")
                .join("jre1.8.0_491")
                .join("bin")
                .join(executable_name(name)),
        );
        candidates.push(
            PathBuf::from(&program_files)
                .join("Java")
                .join("jdk")
                .join("bin")
                .join(executable_name(name)),
        );
    }

    if let Ok(program_files_x86) = env::var("ProgramFiles(x86)") {
        candidates.push(
            PathBuf::from(program_files_x86)
                .join("Common Files")
                .join("Oracle")
                .join("Java")
                .join("java8path")
                .join(executable_name(name)),
        );
    }

    candidates.into_iter().find(|candidate| candidate.exists())
}

fn locate_javac() -> Option<PathBuf> {
    if let Some(path) = locate_executable("javac") {
        return Some(path);
    }

    let mut candidates = Vec::new();
    if let Ok(program_files) = env::var("ProgramFiles") {
        candidates.push(
            PathBuf::from(&program_files)
                .join("Android")
                .join("Android Studio1")
                .join("jbr")
                .join("bin")
                .join("javac.exe"),
        );
        candidates.push(
            PathBuf::from(&program_files)
                .join("Java")
                .join("jdk-21")
                .join("bin")
                .join("javac.exe"),
        );
        candidates.push(
            PathBuf::from(&program_files)
                .join("Java")
                .join("jdk-17")
                .join("bin")
                .join("javac.exe"),
        );
    }

    if let Ok(program_files_x86) = env::var("ProgramFiles(x86)") {
        candidates.push(
            PathBuf::from(program_files_x86)
                .join("Common Files")
                .join("Oracle")
                .join("Java")
                .join("javac.exe"),
        );
    }

    candidates.into_iter().find(|candidate| candidate.exists())
}

fn locate_with_where(name: &str) -> Result<PathBuf, String> {
    let output = Command::new("where")
        .arg(name)
        .output()
        .map_err(|error| format!("Unable to run where for {name}: {error}"))?;

    if !output.status.success() {
        return Err(format!("where could not find {name}"));
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    stdout
        .lines()
        .map(PathBuf::from)
        .find(|path| path.exists())
        .ok_or_else(|| format!("where returned no valid path for {name}"))
}

fn executable_name(name: &str) -> String {
    if cfg!(windows) {
        format!("{name}.exe")
    } else {
        name.to_string()
    }
}

fn discover_sample_files(manifest_dir: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    let candidates = [
        manifest_dir.parent().map(|root| {
            root.join("upstream")
                .join("projectlibre-snapshot")
                .join("projectlibre_build")
                .join("resources")
                .join("samples")
        }),
        manifest_dir.parent().map(|root| {
            root.join("upstream")
                .join("projectlibre-snapshot")
                .join("projectlibre_exchange")
                .join("testdata")
        }),
    ];

    for candidate in candidates.into_iter().flatten() {
        collect_mpp_files(&candidate, &mut files);
    }

    files.sort();
    files.dedup();
    files
}

fn collect_mpp_files(directory: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(directory) else {
        return;
    };

    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_mpp_files(&path, files);
        } else if path
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|ext| ext.eq_ignore_ascii_case("mpp"))
            .unwrap_or(false)
        {
            files.push(path);
        }
    }
}

fn metadata_modified(path: &Path) -> Option<SystemTime> {
    fs::metadata(path)
        .and_then(|metadata| metadata.modified())
        .ok()
}
