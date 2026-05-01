use tauri::Manager;

fn main() {
    tauri::Builder::default()
        .run(tauri::generate_context!("src-tauri/tauri.conf.json"))
        .expect("error while running tauri application");
}
