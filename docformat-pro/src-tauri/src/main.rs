use tauri;

#[tauri::command]
fn read_file(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(&path).map_err(|e| e.to_string())
}

#[tauri::command]
fn write_file(path: String, data: Vec<u8>) -> Result<(), String> {
    std::fs::write(&path, data).map_err(|e| e.to_string())
}

#[tauri::command]
fn get_system_fonts() -> Result<Vec<String>, String> {
    // Placeholder for system font discovery
    Ok(vec!["宋体".into(), "黑体".into(), "仿宋".into(), "楷体".into()])
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![read_file, write_file, get_system_fonts])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
