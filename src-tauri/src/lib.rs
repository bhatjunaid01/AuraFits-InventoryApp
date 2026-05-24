mod db;

use tauri::Manager;
use tauri_plugin_updater::UpdaterExt;

use db::{
    add_category, add_expense, add_product, delete_category, delete_expense, delete_product,
    find_user, get_categories, get_expenses, get_products, get_sales_stats, get_today_sales,
    get_statement, login, record_sale, reset_password, signup, update_category, update_expense,
    update_product, delete_sale, update_sale, get_pending_balances, update_balance,
    verify_security_answer, AppState,
};

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .setup(|app| {
            let state = AppState::new(app.handle())?;
            app.manage(state);

            // Check for updates on launch, prompt user before installing
            let handle = app.handle().clone();
            tauri::async_runtime::spawn(async move {
                if let Ok(updater) = handle.updater() {
                    if let Ok(Some(update)) = updater.check().await {
                        let version = update.version.clone();
                        let dialog = tauri_plugin_dialog::DialogExt::dialog(&handle)
                            .message(format!(
                                "Aura Fits {} is available. Install now?\nThe app will restart after updating.",
                                version
                            ))
                            .title("Update Available")
                            .ok_button_label("Install")
                            .cancel_button_label("Later");
                        if dialog.blocking_show() {
                            let _ = update.download_and_install(|_, _| {}, || {}).await;
                            tauri_plugin_process::restart(&handle.env());
                        }
                    }
                }
            });

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            login,
            signup,
            find_user,
            verify_security_answer,
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
            get_statement,
            delete_sale,
            update_sale,
            get_pending_balances,
            update_balance
        ])
        .run(tauri::generate_context!())
        .expect("error while running Aura Fits");
}