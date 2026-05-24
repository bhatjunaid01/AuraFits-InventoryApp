use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use std::fs;
use std::sync::Mutex;
use tauri::{AppHandle, Manager, State};

pub struct AppState {
    conn: Mutex<Connection>,
}

impl AppState {
    pub fn new(app: &AppHandle) -> Result<Self, Box<dyn std::error::Error>> {
        let app_data_dir = app.path().app_data_dir()?;
        fs::create_dir_all(&app_data_dir)?;
        let db_path = app_data_dir.join("aurafits.db");
        let conn = Connection::open(db_path)?;
        conn.execute_batch("PRAGMA foreign_keys = ON; PRAGMA journal_mode = WAL;")?;
        init_db(&conn)?;
        Ok(Self {
            conn: Mutex::new(conn),
        })
    }
}

fn init_db(conn: &Connection) -> rusqlite::Result<()> {
    conn.execute_batch(
        "
        CREATE TABLE IF NOT EXISTS products (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          category TEXT,
          brand TEXT,
          size TEXT,
          color TEXT,
          cost REAL DEFAULT 0,
          price REAL DEFAULT 0,
          stock INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS categories (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          icon TEXT DEFAULT 'shirt',
          count INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS expenses (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          category TEXT,
          amount REAL DEFAULT 0,
          date TEXT,
          notes TEXT
        );

        CREATE TABLE IF NOT EXISTS sales (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          total REAL DEFAULT 0,
          discount REAL DEFAULT 0,
          payment TEXT,
          customer_name TEXT,
          customer_phone TEXT,
          amount_paid REAL DEFAULT 0,
          balance REAL DEFAULT 0,
          created_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS sale_items (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          sale_id INTEGER,
          product_id INTEGER,
          product_name TEXT,
          category TEXT,
          size TEXT,
          color TEXT,
          cost REAL DEFAULT 0,
          qty INTEGER,
          price REAL,
          FOREIGN KEY(sale_id) REFERENCES sales(id)
        );

        CREATE TABLE IF NOT EXISTS users (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT,
          email TEXT UNIQUE,
          password TEXT,
          security_question TEXT,
          security_answer TEXT
        );
        ",
    )?;

    for migration in [
        "ALTER TABLE sales ADD COLUMN customer_name TEXT",
        "ALTER TABLE sale_items ADD COLUMN category TEXT",
        "ALTER TABLE sale_items ADD COLUMN size TEXT",
        "ALTER TABLE sale_items ADD COLUMN color TEXT",
        "ALTER TABLE sale_items ADD COLUMN cost REAL DEFAULT 0",
        "ALTER TABLE sales ADD COLUMN customer_phone TEXT",
        "ALTER TABLE sales ADD COLUMN amount_paid REAL DEFAULT 0",
        "ALTER TABLE sales ADD COLUMN balance REAL DEFAULT 0",
    ] {
        let _ = conn.execute(migration, []);
    }

    let existing_admin: Option<i64> = conn
        .query_row(
            "SELECT id FROM users WHERE email = ?1",
            ["admin@aura.fits"],
            |row| row.get(0),
        )
        .optional()?;

    if existing_admin.is_none() {
        conn.execute(
            "INSERT INTO users (name, email, password, security_question, security_answer)
             VALUES (?1, ?2, ?3, ?4, ?5)",
            params![
                "Admin",
                "admin@aura.fits",
                "admin123",
                "What was the name of your first pet?",
                "buddy"
            ],
        )?;
    }

    Ok(())
}

fn lock_conn<'a, 'r>(
    state: &'a State<'r, AppState>,
) -> Result<std::sync::MutexGuard<'a, Connection>, String> {
    state
        .conn
        .lock()
        .map_err(|_| "Database connection is unavailable".to_string())
}

#[derive(Debug, Serialize)]
pub struct OkResponse {
    ok: bool,
}

#[derive(Debug, Serialize)]
pub struct LoginResponse {
    ok: bool,
    name: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct SignupResponse {
    ok: bool,
    error: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SecurityUser {
    security_question: String,
    security_answer: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SignupInput {
    name: String,
    email: String,
    password: String,
    security_question: String,
    security_answer: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResetPasswordInput {
    email: String,
    new_password: String,
    security_answer: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct Product {
    id: Option<i64>,
    name: String,
    category: Option<String>,
    brand: Option<String>,
    size: Option<String>,
    color: Option<String>,
    cost: f64,
    price: f64,
    stock: i64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct Category {
    id: Option<i64>,
    name: String,
    icon: Option<String>,
    count: Option<i64>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct Expense {
    id: Option<i64>,
    name: String,
    category: Option<String>,
    amount: f64,
    date: Option<String>,
    notes: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct CartItem {
    id: i64,
    name: String,
    category: Option<String>,
    size: Option<String>,
    color: Option<String>,
    cost: f64,
    qty: i64,
    price: f64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct IdResponse {
    id: i64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SaleResponse {
    ok: bool,
    sale_id: i64,
}

#[derive(Debug, Serialize)]
pub struct TodaySale {
    id: i64,
    total: f64,
    discount: f64,
    payment: Option<String>,
    created_at: Option<String>,
    items_summary: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct SalesTotal {
    total: f64,
    count: Option<i64>,
}

#[derive(Debug, Serialize)]
pub struct WeeklySale {
    dow: String,
    sales: f64,
}

#[derive(Debug, Serialize)]
pub struct SalesStats {
    today: SalesTotal,
    month: SalesTotal,
    weekly: Vec<WeeklySale>,
    real_profit: f64,
}

#[derive(Debug, Serialize)]
pub struct StatementRow {
    serial_no: i64,
    sale_id: i64,
    product_id: i64,
    created_at: String,
    customer_name: String,
    customer_phone: String,
    payment: String,
    product_name: String,
    category: String,
    size: String,
    color: String,
    qty: i64,
    cost: f64,
    price: f64,
    line_total: f64,
    profit: f64,
    sale_total: f64,
    discount: f64,
    amount_paid: f64,
    balance: f64,
}

#[tauri::command]
pub fn login(
    state: State<AppState>,
    email: String,
    password: String,
) -> Result<LoginResponse, String> {
    let conn = lock_conn(&state)?;
    let user: Option<(String, String)> = conn
        .query_row(
            "SELECT name, password FROM users WHERE email = ?1",
            [email],
            |row| Ok((row.get(0)?, row.get(1)?)),
        )
        .optional()
        .map_err(|e| e.to_string())?;

    Ok(match user {
        Some((name, saved_password)) if saved_password == password => LoginResponse {
            ok: true,
            name: Some(name),
        },
        _ => LoginResponse {
            ok: false,
            name: None,
        },
    })
}

#[tauri::command]
pub fn signup(state: State<AppState>, input: SignupInput) -> Result<SignupResponse, String> {
    let conn = lock_conn(&state)?;
    let exists: Option<i64> = conn
        .query_row(
            "SELECT id FROM users WHERE email = ?1",
            [&input.email],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;

    if exists.is_some() {
        return Ok(SignupResponse {
            ok: false,
            error: Some("Email already exists".to_string()),
        });
    }

    conn.execute(
        "INSERT INTO users (name, email, password, security_question, security_answer)
         VALUES (?1, ?2, ?3, ?4, ?5)",
        params![
            input.name,
            input.email,
            input.password,
            input.security_question,
            input.security_answer.to_lowercase().trim()
        ],
    )
    .map_err(|e| e.to_string())?;

    Ok(SignupResponse {
        ok: true,
        error: None,
    })
}

#[tauri::command]
pub fn find_user(state: State<AppState>, email: String) -> Result<Option<String>, String> {
    let conn = lock_conn(&state)?;
    conn.query_row(
        "SELECT security_question FROM users WHERE email = ?1",
        [email],
        |row| row.get(0),
    )
    .optional()
    .map_err(|e| e.to_string())
}

#[tauri::command]
pub fn verify_security_answer(
    state: State<AppState>,
    email: String,
    answer: String,
) -> Result<bool, String> {
    let conn = lock_conn(&state)?;
    let stored: Option<String> = conn
        .query_row(
            "SELECT security_answer FROM users WHERE email = ?1",
            [&email],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    Ok(stored.map(|s| s.to_lowercase().trim().to_string()) == Some(answer.to_lowercase().trim().to_string()))
}

#[tauri::command]
pub fn reset_password(
    state: State<AppState>,
    input: ResetPasswordInput,
) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    // Verify security answer before allowing password reset
    let stored_answer: Option<String> = conn
        .query_row(
            "SELECT security_answer FROM users WHERE email = ?1",
            [&input.email],
            |row| row.get(0),
        )
        .optional()
        .map_err(|e| e.to_string())?;
    let answer_ok = stored_answer
        .map(|s| s.to_lowercase().trim().to_string())
        == Some(input.security_answer.to_lowercase().trim().to_string());
    if !answer_ok {
        return Ok(OkResponse { ok: false });
    }
    conn.execute(
        "UPDATE users SET password = ?1 WHERE email = ?2",
        params![input.new_password, input.email],
    )
    .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn get_products(state: State<AppState>) -> Result<Vec<Product>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn
        .prepare(
            "SELECT id,
                    COALESCE(name, ''),
                    COALESCE(category, ''),
                    COALESCE(brand, ''),
                    COALESCE(size, ''),
                    COALESCE(color, ''),
                    COALESCE(cost, 0),
                    COALESCE(price, 0),
                    COALESCE(stock, 0)
             FROM products ORDER BY name",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            Ok(Product {
                id: Some(row.get(0)?),
                name: row.get(1)?,
                category: row.get(2)?,
                brand: row.get(3)?,
                size: row.get(4)?,
                color: row.get(5)?,
                cost: row.get(6)?,
                price: row.get(7)?,
                stock: row.get(8)?,
            })
        })
        .map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub fn add_product(state: State<AppState>, product: Product) -> Result<IdResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "INSERT INTO products (name, category, brand, size, color, cost, price, stock)
         VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
        params![
            product.name,
            product.category,
            product.brand,
            product.size,
            product.color,
            product.cost,
            product.price,
            product.stock
        ],
    )
    .map_err(|e| e.to_string())?;
    Ok(IdResponse {
        id: conn.last_insert_rowid(),
    })
}

#[tauri::command]
pub fn update_product(state: State<AppState>, product: Product) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "UPDATE products SET name=?1, category=?2, brand=?3, size=?4, color=?5, cost=?6, price=?7, stock=?8 WHERE id=?9",
        params![
            product.name,
            product.category,
            product.brand,
            product.size,
            product.color,
            product.cost,
            product.price,
            product.stock,
            product.id
        ],
    )
    .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn delete_product(state: State<AppState>, id: i64) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute("DELETE FROM products WHERE id = ?1", [id])
        .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn get_categories(state: State<AppState>) -> Result<Vec<Category>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn
        .prepare("SELECT id, COALESCE(name, ''), COALESCE(icon, 'shirt'), COALESCE(count, 0) FROM categories ORDER BY name")
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            Ok(Category {
                id: Some(row.get(0)?),
                name: row.get(1)?,
                icon: row.get(2)?,
                count: row.get(3)?,
            })
        })
        .map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub fn add_category(state: State<AppState>, category: Category) -> Result<IdResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "INSERT INTO categories (name, icon) VALUES (?1, ?2)",
        params![category.name, category.icon],
    )
    .map_err(|e| e.to_string())?;
    Ok(IdResponse {
        id: conn.last_insert_rowid(),
    })
}

#[tauri::command]
pub fn update_category(state: State<AppState>, category: Category) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "UPDATE categories SET name=?1, icon=?2 WHERE id=?3",
        params![category.name, category.icon, category.id],
    )
    .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn delete_category(state: State<AppState>, id: i64) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute("DELETE FROM categories WHERE id = ?1", [id])
        .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn get_expenses(state: State<AppState>) -> Result<Vec<Expense>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn
        .prepare(
            "SELECT id,
                    COALESCE(name, ''),
                    COALESCE(category, ''),
                    COALESCE(amount, 0),
                    COALESCE(date, ''),
                    COALESCE(notes, '')
             FROM expenses ORDER BY date DESC",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            Ok(Expense {
                id: Some(row.get(0)?),
                name: row.get(1)?,
                category: row.get(2)?,
                amount: row.get(3)?,
                date: row.get(4)?,
                notes: row.get(5)?,
            })
        })
        .map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub fn add_expense(state: State<AppState>, expense: Expense) -> Result<IdResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "INSERT INTO expenses (name, category, amount, date, notes) VALUES (?1, ?2, ?3, ?4, ?5)",
        params![
            expense.name,
            expense.category,
            expense.amount,
            expense.date,
            expense.notes
        ],
    )
    .map_err(|e| e.to_string())?;
    Ok(IdResponse {
        id: conn.last_insert_rowid(),
    })
}

#[tauri::command]
pub fn update_expense(state: State<AppState>, expense: Expense) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "UPDATE expenses SET name=?1, category=?2, amount=?3, date=?4, notes=?5 WHERE id=?6",
        params![
            expense.name,
            expense.category,
            expense.amount,
            expense.date,
            expense.notes,
            expense.id
        ],
    )
    .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn delete_expense(state: State<AppState>, id: i64) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute("DELETE FROM expenses WHERE id = ?1", [id])
        .map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[tauri::command]
pub fn record_sale(
    state: State<AppState>,
    cart: Vec<CartItem>,
    total: f64,
    discount: f64,
    payment: String,
    customer_name: String,
    customer_phone: Option<String>,
    amount_paid: Option<f64>,
    balance: Option<f64>,
) -> Result<SaleResponse, String> {
    let mut conn = lock_conn(&state)?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    tx.execute(
        "INSERT INTO sales (total, discount, payment, customer_name, customer_phone, amount_paid, balance)
         VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
        params![
            total,
            discount,
            payment,
            customer_name,
            customer_phone.unwrap_or_default(),
            amount_paid.unwrap_or(0.0),
            balance.unwrap_or(0.0)
        ],
    )
    .map_err(|e| e.to_string())?;
    let sale_id = tx.last_insert_rowid();

    for item in cart {
        tx.execute(
            "INSERT INTO sale_items (sale_id, product_id, product_name, category, size, color, cost, qty, price)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)",
            params![
                sale_id,
                item.id,
                item.name,
                item.category,
                item.size,
                item.color,
                item.cost,
                item.qty,
                item.price
            ],
        )
        .map_err(|e| e.to_string())?;
        let rows_affected = tx.execute(
            "UPDATE products SET stock = stock - ?1 WHERE id = ?2 AND stock >= ?1",
            params![item.qty, item.id],
        )
        .map_err(|e| e.to_string())?;
        if rows_affected == 0 {
            return Err(format!("Insufficient stock for product id {}", item.id));
        }
    }

    tx.commit().map_err(|e| e.to_string())?;
    Ok(SaleResponse { ok: true, sale_id })
}

#[tauri::command]
pub fn get_today_sales(state: State<AppState>) -> Result<Vec<TodaySale>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn
        .prepare(
            "
            SELECT s.id,
                   COALESCE(s.total, 0),
                   COALESCE(s.discount, 0),
                   COALESCE(s.payment, ''),
                   COALESCE(s.created_at, ''),
                   COALESCE(GROUP_CONCAT(COALESCE(si.product_name, '') || ' x' || COALESCE(si.qty, 0)), '') as items_summary
            FROM sales s
            LEFT JOIN sale_items si ON s.id = si.sale_id
            WHERE date(s.created_at) = date('now','localtime')
            GROUP BY s.id
            ORDER BY s.created_at DESC
            ",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |row| {
            Ok(TodaySale {
                id: row.get(0)?,
                total: row.get(1)?,
                discount: row.get(2)?,
                payment: row.get(3)?,
                created_at: row.get(4)?,
                items_summary: row.get(5)?,
            })
        })
        .map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub fn get_sales_stats(state: State<AppState>) -> Result<SalesStats, String> {
    let conn = lock_conn(&state)?;
    let today = conn
        .query_row(
            "SELECT COALESCE(SUM(total),0) as total, COUNT(*) as count
             FROM sales WHERE date(created_at)=date('now','localtime')",
            [],
            |row| {
                Ok(SalesTotal {
                    total: row.get(0)?,
                    count: Some(row.get(1)?),
                })
            },
        )
        .map_err(|e| e.to_string())?;

    let month = conn
        .query_row(
            "SELECT COALESCE(SUM(total),0) as total
             FROM sales WHERE strftime('%Y-%m', created_at)=strftime('%Y-%m','now','localtime')",
            [],
            |row| {
                Ok(SalesTotal {
                    total: row.get(0)?,
                    count: None,
                })
            },
        )
        .map_err(|e| e.to_string())?;

    let mut stmt = conn
        .prepare(
            "
            SELECT strftime('%w', created_at) as dow, COALESCE(SUM(total),0) as sales
            FROM sales WHERE date(created_at) >= date('now','-6 days','localtime')
            GROUP BY dow
            ",
        )
        .map_err(|e| e.to_string())?;
    let weekly = stmt
        .query_map([], |row| {
            Ok(WeeklySale {
                dow: row.get(0)?,
                sales: row.get(1)?,
            })
        })
        .map_err(|e| e.to_string())?
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())?;

    let real_profit: f64 = conn
        .query_row(
            "SELECT COALESCE(SUM(
                (si.price - COALESCE(NULLIF(si.cost,0), p.cost, 0)) * si.qty
                * (CASE WHEN s.total > 0 THEN (s.total - COALESCE(s.balance,0)) / s.total ELSE 1 END)
             ), 0)
             FROM sale_items si
             JOIN sales s ON s.id = si.sale_id
             LEFT JOIN products p ON p.id = si.product_id
             WHERE strftime('%Y-%m', s.created_at) = strftime('%Y-%m','now','localtime')",
            [],
            |row| row.get(0),
        )
        .map_err(|e| e.to_string())?;

    Ok(SalesStats {
        today,
        month,
        weekly,
        real_profit,
    })
}

#[tauri::command]
pub fn get_statement(
    state: State<AppState>,
    from_date: String,
    to_date: String,
) -> Result<Vec<StatementRow>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn
        .prepare(
            "
            SELECT s.id,
                   COALESCE(si.product_id, 0),
                   COALESCE(s.created_at, ''),
                   COALESCE(s.customer_name, ''),
                   COALESCE(s.customer_phone, ''),
                   COALESCE(s.payment, ''),
                   COALESCE(si.product_name, ''),
                   COALESCE(si.category, p.category, ''),
                   COALESCE(si.size, p.size, ''),
                   COALESCE(si.color, p.color, ''),
                   COALESCE(si.qty, 0),
                   COALESCE(NULLIF(si.cost, 0), p.cost, 0),
                   COALESCE(si.price, 0),
                   COALESCE(si.qty, 0) * COALESCE(si.price, 0) as line_total,
                   (COALESCE(si.price, 0) - COALESCE(NULLIF(si.cost, 0), p.cost, 0)) * COALESCE(si.qty, 0) as profit,
                   COALESCE(s.total, 0),
                   COALESCE(s.discount, 0),
                   COALESCE(s.amount_paid, 0),
                   COALESCE(s.balance, 0)
            FROM sales s
            LEFT JOIN sale_items si ON s.id = si.sale_id
            LEFT JOIN products p ON p.id = si.product_id
            WHERE date(s.created_at) BETWEEN date(?1) AND date(?2)
            ORDER BY s.created_at DESC, s.id DESC, si.id ASC
            ",
        )
        .map_err(|e| e.to_string())?;

    let rows = stmt
        .query_map(params![from_date, to_date], |row| {
            Ok(StatementRow {
                serial_no: 0,
                sale_id: row.get(0)?,
                product_id: row.get(1)?,
                created_at: row.get(2)?,
                customer_name: row.get(3)?,
                customer_phone: row.get(4)?,
                payment: row.get(5)?,
                product_name: row.get(6)?,
                category: row.get(7)?,
                size: row.get(8)?,
                color: row.get(9)?,
                qty: row.get(10)?,
                cost: row.get(11)?,
                price: row.get(12)?,
                line_total: row.get(13)?,
                profit: row.get(14)?,
                sale_total: row.get(15)?,
                discount: row.get(16)?,
                amount_paid: row.get(17)?,
                balance: row.get(18)?,
            })
        })
        .map_err(|e| e.to_string())?;

    let mut rows = rows
        .collect::<rusqlite::Result<Vec<_>>>()
        .map_err(|e| e.to_string())?;
    for (index, row) in rows.iter_mut().enumerate() {
        row.serial_no = (index + 1) as i64;
    }
    Ok(rows)
}

#[tauri::command]
pub fn delete_sale(state: State<AppState>, id: i64) -> Result<OkResponse, String> {
    let mut conn = lock_conn(&state)?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    // Return stock for all items in this sale
    {
        let mut stmt = tx.prepare(
            "SELECT product_id, qty FROM sale_items WHERE sale_id = ?1"
        ).map_err(|e| e.to_string())?;
        let items: Vec<(i64, i64)> = stmt.query_map([id], |row| {
            Ok((row.get(0)?, row.get(1)?))
        }).map_err(|e| e.to_string())?
        .filter_map(|r| r.ok())
        .collect();
        for (product_id, qty) in items {
            tx.execute(
                "UPDATE products SET stock = stock + ?1 WHERE id = ?2",
                params![qty, product_id],
            ).map_err(|e| e.to_string())?;
        }
    }
    tx.execute("DELETE FROM sale_items WHERE sale_id = ?1", [id])
        .map_err(|e| e.to_string())?;
    tx.execute("DELETE FROM sales WHERE id = ?1", [id])
        .map_err(|e| e.to_string())?;
    tx.commit().map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct UpdateSaleInput {
    pub id: i64,
    pub customer_name: String,
    pub customer_phone: String,
    pub payment: String,
    pub discount: f64,
    pub amount_paid: f64,
    pub balance: f64,
    pub total: f64,
    pub cart: Vec<CartItem>,
}

#[tauri::command]
pub fn update_sale(state: State<AppState>, input: UpdateSaleInput) -> Result<OkResponse, String> {
    let mut conn = lock_conn(&state)?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    // Restore stock for old items
    {
        let mut stmt = tx.prepare(
            "SELECT product_id, qty FROM sale_items WHERE sale_id = ?1"
        ).map_err(|e| e.to_string())?;
        let old_items: Vec<(i64, i64)> = stmt.query_map([input.id], |row| {
            Ok((row.get(0)?, row.get(1)?))
        }).map_err(|e| e.to_string())?
        .filter_map(|r| r.ok())
        .collect();
        for (product_id, qty) in old_items {
            tx.execute(
                "UPDATE products SET stock = stock + ?1 WHERE id = ?2",
                params![qty, product_id],
            ).map_err(|e| e.to_string())?;
        }
    }
    // Delete old items
    tx.execute("DELETE FROM sale_items WHERE sale_id = ?1", [input.id])
        .map_err(|e| e.to_string())?;
    // Update sale header
    tx.execute(
        "UPDATE sales SET customer_name=?1, customer_phone=?2, payment=?3, discount=?4, amount_paid=?5, balance=?6, total=?7 WHERE id=?8",
        params![input.customer_name, input.customer_phone, input.payment, input.discount, input.amount_paid, input.balance, input.total, input.id],
    ).map_err(|e| e.to_string())?;
    // Insert new items and deduct stock
    for item in &input.cart {
        tx.execute(
            "INSERT INTO sale_items (sale_id, product_id, product_name, category, size, color, cost, qty, price)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)",
            params![input.id, item.id, item.name, item.category, item.size, item.color, item.cost, item.qty, item.price],
        ).map_err(|e| e.to_string())?;
        let rows_affected = tx.execute(
            "UPDATE products SET stock = stock - ?1 WHERE id = ?2 AND stock >= ?1",
            params![item.qty, item.id],
        ).map_err(|e| e.to_string())?;
        if rows_affected == 0 {
            return Err(format!("Insufficient stock for product id {}", item.id));
        }
    }
    tx.commit().map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}
#[derive(Debug, Serialize)]
pub struct PendingBalance {
    sale_id: i64,
    customer_name: String,
    customer_phone: String,
    sale_total: f64,
    amount_paid: f64,
    balance: f64,
    payment: String,
    created_at: String,
    items_summary: String,
}

#[tauri::command]
pub fn get_pending_balances(state: State<AppState>) -> Result<Vec<PendingBalance>, String> {
    let conn = lock_conn(&state)?;
    let mut stmt = conn.prepare(
        "SELECT s.id,
                COALESCE(s.customer_name, 'Walk-in Customer'),
                COALESCE(s.customer_phone, ''),
                COALESCE(s.total, 0),
                COALESCE(s.amount_paid, 0),
                COALESCE(s.balance, 0),
                COALESCE(s.payment, ''),
                COALESCE(s.created_at, ''),
                COALESCE(GROUP_CONCAT(si.product_name || ' x' || si.qty), '') as items_summary
         FROM sales s
         LEFT JOIN sale_items si ON s.id = si.sale_id
         WHERE s.balance > 0
         GROUP BY s.id
         ORDER BY s.created_at DESC"
    ).map_err(|e| e.to_string())?;
    let rows = stmt.query_map([], |row| {
        Ok(PendingBalance {
            sale_id: row.get(0)?,
            customer_name: row.get(1)?,
            customer_phone: row.get(2)?,
            sale_total: row.get(3)?,
            amount_paid: row.get(4)?,
            balance: row.get(5)?,
            payment: row.get(6)?,
            created_at: row.get(7)?,
            items_summary: row.get(8)?,
        })
    }).map_err(|e| e.to_string())?;
    rows.collect::<rusqlite::Result<Vec<_>>>().map_err(|e| e.to_string())
}

#[tauri::command]
pub fn update_balance(state: State<AppState>, sale_id: i64, amount_paid: f64, balance: f64) -> Result<OkResponse, String> {
    let conn = lock_conn(&state)?;
    conn.execute(
        "UPDATE sales SET amount_paid = ?1, balance = ?2 WHERE id = ?3",
        params![amount_paid, balance, sale_id],
    ).map_err(|e| e.to_string())?;
    Ok(OkResponse { ok: true })
}