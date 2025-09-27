import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import traceback
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import os
import json

DB_NAME = "clients.db"
SHEET_ID = "1_DfTT8yzCjP0VH0PZu1Fz6FYMm1eRr7c0TmZU2DrH_w"

# ================== База данных ==================
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fio TEXT NOT NULL,
                dob TEXT,
                phone TEXT,
                contract_number TEXT,
                ippcu_start TEXT,
                ippcu_end TEXT,
                group_name TEXT
            )
            """
        )
        conn.commit()


def add_client(fio, dob, phone, contract_number, ippcu_start, ippcu_end, group):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO clients (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group),
        )
        conn.commit()


def get_all_clients(limit=200):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            ORDER BY fio
            LIMIT ?
            """,
            (limit,),
        )
        return cur.fetchall()


def search_clients(query="", date_from=None, date_to=None, limit=200):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        like = f"%{query}%"

        sql = """
            SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE (fio LIKE ? OR contract_number LIKE ? OR phone LIKE ? 
                   OR ippcu_start LIKE ? OR ippcu_end LIKE ? OR group_name LIKE ?)
        """
        params = [like, like, like, like, like, like]

        if date_from:
            sql += " AND DATE(ippcu_end) >= DATE(?) "
            params.append(date_from)
        if date_to:
            sql += " AND DATE(ippcu_end) <= DATE(?) "
            params.append(date_to)

        sql += " ORDER BY fio LIMIT ?"
        params.append(limit)

        cur.execute(sql, params)
        return cur.fetchall()


def update_client(cid, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE clients
            SET fio=?, dob=?, phone=?, contract_number=?, ippcu_start=?, ippcu_end=?, group_name=?
            WHERE id=?
            """,
            (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group, cid),
        )
        conn.commit()


def delete_client(cid):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM clients WHERE id=?", (cid,))
        conn.commit()


# ================== Google Sheets ==================
def get_gsheet(sheet_id, sheet_name="Лист1"):
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise RuntimeError("Секрет GOOGLE_CREDENTIALS не найден!")

    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
    return sheet


def import_from_gsheet():
    try:
        sheet = get_gsheet(SHEET_ID)
        data = sheet.get_all_records()

        for row in data:
            add_client(
                row.get("ФИО", ""),
                row.get("Дата рождения", ""),
                row.get("Телефон", ""),
                row.get("Номер договора", ""),
                row.get("Дата начала ИППСУ", ""),
                row.get("Дата окончания ИППСУ", ""),
                row.get("Группа", ""),
            )
        refresh_tree()
        messagebox.showinfo("Успех", "Импорт из Google Sheets завершён!")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")


# ================== UI ==================
def refresh_tree(results=None):
    for row in tree.get_children():
        tree.delete(row)

    if results is None:
        results = get_all_clients(limit=200)

    today = datetime.today().date()
    soon = today + timedelta(days=30)

    for row in results:
        cid, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group = row

        tag = ""
        try:
            if ippcu_end:
                end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
                if end_date < today:
                    tag = "expired"
                elif today <= end_date <= soon:
                    tag = "soon"
                else:
                    tag = "active"
        except:
            pass

        tree.insert("", "end", values=row, tags=(tag,))


def add_window():
    win = tk.Toplevel()
    win.title("Добавить обслуживаемого")

    tk.Label(win, text="ФИО").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=30)
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=30)
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=30)
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_start.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата окончания ИППСУ").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_end.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=30)
    e_group.grid(row=6, column=1, padx=10, pady=5)

    def save_client():
        fio = e_fio.get().strip()
        dob = e_dob.get_date().strftime("%Y-%m-%d")
        phone = e_phone.get().strip()
        contract_number = e_contract.get().strip()
        ippcu_start = e_ippcu_start.get_date().strftime("%Y-%m-%d")
        ippcu_end = e_ippcu_end.get_date().strftime("%Y-%m-%d")
        group = e_group.get().strip()

        if not fio:
            messagebox.showerror("Ошибка", "Поле 'ФИО' обязательно для заполнения!")
            return

        try:
            add_client(fio, dob, phone, contract_number, ippcu_start, ippcu_end, group)
            refresh_tree()
            win.destroy()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось добавить:\n{e}")

    tk.Button(win, text="Сохранить", command=save_client).grid(row=7, column=0, columnspan=2, pady=10)


def edit_client():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Ошибка", "Выберите клиента для редактирования")
        return

    item = tree.item(selected[0])
    cid, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group = item["values"]

    win = tk.Toplevel()
    win.title("Редактировать обслуживаемого")

    tk.Label(win, text="ФИО").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=30)
    e_fio.insert(0, fio)
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        e_dob.set_date(datetime.strptime(dob, "%Y-%m-%d"))
    except:
        pass
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=30)
    e_phone.insert(0, phone)
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=30)
    e_contract.insert(0, contract_number)
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        e_ippcu_start.set_date(datetime.strptime(ippcu_start, "%Y-%m-%d"))
    except:
        pass
    e_ippcu_start.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата окончания ИППСУ").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        e_ippcu_end.set_date(datetime.strptime(ippcu_end, "%Y-%m-%d"))
    except:
        pass
    e_ippcu_end.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=30)
    e_group.insert(0, group)
    e_group.grid(row=6, column=1, padx=10, pady=5)

    def save_edit():
        new_fio = e_fio.get().strip()
        new_dob = e_dob.get_date().strftime("%Y-%m-%d")
        new_phone = e_phone.get().strip()
        new_contract = e_contract.get().strip()
        new_ippcu_start = e_ippcu_start.get_date().strftime("%Y-%m-%d")
        new_ippcu_end = e_ippcu_end.get_date().strftime("%Y-%m-%d")
        new_group = e_group.get().strip()

        if not new_fio:
            messagebox.showerror("Ошибка", "Поле 'ФИО' обязательно!")
            return

        try:
            update_client(cid, new_fio, new_dob, new_phone, new_contract,
                          new_ippcu_start, new_ippcu_end, new_group)
            refresh_tree()
            win.destroy()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    tk.Button(win, text="Сохранить изменения", command=save_edit).grid(row=7, column=0, columnspan=2, pady=10)


def delete_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Ошибка", "Выберите клиента для удаления")
        return
    item = tree.item(selected[0])
    cid = item["values"][0]
    if messagebox.askyesno("Удалить", "Точно удалить выбранного клиента?"):
        delete_client(cid)
        refresh_tree()


def do_search():
    query = search_entry.get().strip()
    date_from = date_from_entry.get_date().strftime("%Y-%m-%d") if date_from_entry.get() else None
    date_to = date_to_entry.get_date().strftime("%Y-%m-%d") if date_to_entry.get() else None

    results = search_clients(query=query, date_from=date_from, date_to=date_to, limit=200)
    refresh_tree(results)


# ================== MAIN ==================
root = tk.Tk()
root.title("База клиентов")

# Поиск + фильтры
search_entry = tk.Entry(root, width=40)
search_entry.grid(row=0, column=0, padx=5, pady=5, sticky="w")
tk.Button(root, text="Поиск", command=do_search).grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Дата окончания ИППСУ:").grid(row=0, column=2, padx=5)
date_from_entry = DateEntry(root, width=12, date_pattern="dd.mm.yyyy")
date_from_entry.grid(row=0, column=3, padx=5)
date_to_entry = DateEntry(root, width=12, date_pattern="dd.mm.yyyy")
date_to_entry.grid(row=0, column=4, padx=5)
tk.Button(root, text="Фильтр", command=do_search).grid(row=0, column=5, padx=5)

# Таблица
tree = ttk.Treeview(root, columns=("ID", "ФИО", "Дата рождения", "Телефон",
                                   "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"),
                    show="headings", height=20)
tree.grid(row=1, column=0, columnspan=7, padx=5, pady=5, sticky="nsew")

for col in tree["columns"]:
    tree.heading(col, text=col)

# Настройка цветов
tree.tag_configure("expired", background="#ffcccc")
tree.tag_configure("soon", background="#fff2cc")
tree.tag_configure("active", background="#ccffcc")

# Кнопки
tk.Button(root, text="Добавить", command=add_window).grid(row=2, column=0, padx=5, pady=5)
tk.Button(root, text="Редактировать", command=edit_client).grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Удалить", command=delete_selected).grid(row=2, column=2, padx=5, pady=5)
tk.Button(root, text="Импорт Google Sheets", command=import_from_gsheet).grid(row=2, column=3, padx=5, pady=5)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

init_db()
root.after(200, refresh_tree)

root.mainloop()
