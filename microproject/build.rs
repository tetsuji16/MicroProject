fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=src/app.rs");
    println!("cargo:rerun-if-changed=src/mspdi.rs");
    println!("cargo:rerun-if-changed=src/main.rs");
}
