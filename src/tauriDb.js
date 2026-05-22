import { invoke } from "@tauri-apps/api/core";

const call = (command, payload = {}) => invoke(command, payload);

const text = (value) => (value == null ? "" : String(value));
const money = (value) => Number(value) || 0;
const integer = (value) => Number.parseInt(value, 10) || 0;

const normalizeProduct = (product) => ({
  ...product,
  id: product.id,
  name: text(product.name),
  category: text(product.category),
  brand: text(product.brand),
  size: text(product.size),
  color: text(product.color),
  cost: money(product.cost),
  price: money(product.price),
  stock: integer(product.stock),
});

const normalizeCategory = (category) => ({
  ...category,
  name: text(category.name),
  icon: text(category.icon || "shirt"),
  count: integer(category.count),
});

const normalizeExpense = (expense) => ({
  ...expense,
  name: text(expense.name),
  category: text(expense.category),
  amount: money(expense.amount),
  date: text(expense.date),
  notes: text(expense.notes),
});

export const db = {
  login: (data) => call("login", data),
  signup: (data) => call("signup", { input: data }),
  findUser: (email) => call("find_user", { email }),
  resetPassword: (data) => call("reset_password", { input: data }),

  getProducts: async () => (await call("get_products")).map(normalizeProduct),
  addProduct: (product) => call("add_product", { product }),
  updateProduct: (product) => call("update_product", { product }),
  deleteProduct: (id) => call("delete_product", { id }),

  getCategories: async () => (await call("get_categories")).map(normalizeCategory),
  addCategory: (category) => call("add_category", { category }),
  updateCategory: (category) => call("update_category", { category }),
  deleteCategory: (id) => call("delete_category", { id }),

  getExpenses: async () => (await call("get_expenses")).map(normalizeExpense),
  addExpense: (expense) => call("add_expense", { expense }),
  updateExpense: (expense) => call("update_expense", { expense }),
  deleteExpense: (id) => call("delete_expense", { id }),

  recordSale: (data) => call("record_sale", data),
  getTodaySales: () => call("get_today_sales"),
  getSalesStats: () => call("get_sales_stats"),
  getStatement: (fromDate, toDate) => call("get_statement", { fromDate, toDate }),
};

export function installDbBridge() {
  window.db = db;
}
