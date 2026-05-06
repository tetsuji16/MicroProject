fn main() {
    tauri_build::build();
    println!("cargo:rerun-if-changed=src/main.rs");
    println!("cargo:rerun-if-changed=src/commands.rs");
    println!("cargo:rerun-if-changed=src/session.rs");
    println!("cargo:rerun-if-changed=src/mspdi.rs");
    println!("cargo:rerun-if-changed=frontend/index.html");
    println!("cargo:rerun-if-changed=frontend/app.js");
    println!("cargo:rerun-if-changed=frontend/styles.css");
}
