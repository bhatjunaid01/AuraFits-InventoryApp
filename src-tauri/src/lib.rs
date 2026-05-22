mod db;

use tauri::Manager;

use db::{
    add_category, add_expense, add_product, delete_category, delete_expense, delete_product,
    find_user, get_categories, get_expenses, get_products, get_sales_stats, get_today_sales,
    get_statement, login, record_sale, reset_password, signup, update_category, update_expense,
    update_product, AppState,
};

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            let state = AppState::new(app.handle())?;
            app.manage(state);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            login,
            signup,
            find_user,
            reset_password,
            get_products,
            add_product,
            update_product,
            delete_product,
            get_categories,
            add_category,
            update_category,
            delete_category,
            get_expenses,
            add_expense,
            update_expense,
            delete_expense,
            record_sale,
            get_today_sales,
            get_sales_stats,
            get_statement
        ])
        .run(tauri::generate_context!())
        .expect("error while running Aura Fits");
}
