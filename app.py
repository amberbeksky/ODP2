import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import traceback
from tkcalendar import DateEntry

DB_NAME = "clients.db"

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


def search_clients(query="", limit=200):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        like = f"%{query}%"
        cur.execute(
            """
            SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE fio LIKE ? OR contract_number LIKE ? OR phone LIKE ? 
                  OR ippcu_start LIKE ? OR ippcu_end LIKE ? OR group_name LIKE ?
            ORDER BY fio
            LIMIT ?
            """,
            (like, like, like, like, like, like, limit),
        )
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


# ================== UI ==================
def refresh_tree(results=None):
    for row in tree.get_children():
        tree.delete(row)

    if results is None:
        results = get_all_clients(limit=200)

    for row in results:
        tree.insert("", "end", values=row)


def add_window():
    win = tk.Toplevel()
    win.title("Добавить обслуживаемого")

    # ===== Поля ввода =====
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

    # ===== Сохранение =====
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

    # Форма
    tk.Label(win, text="ФИО").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=30)
    e_fio.insert(0, fio)
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        from datetime import datetime
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

    # Сохранение
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
    results = search_clients(query=query, limit=200)
    refresh_tree(results)


def import_excel():
    import pandas as pd
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not path:
        return
    try:
        df = pd.read_excel(path)
        for _, row in df.iterrows():
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
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")


def export_excel():
    import pandas as pd
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not path:
        return
    try:
        rows = get_all_clients(limit=1000000)  # выгружаем всех
        df = pd.DataFrame(rows, columns=["ID", "ФИО", "Дата рождения", "Телефон",
                                         "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"])
        df.to_excel(path, index=False)
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось экспортировать:\n{e}")


def show_statistics():
    import pandas as pd
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt

    rows = get_all_clients(limit=1000000)
    df = pd.DataFrame(rows, columns=["ID", "ФИО", "Дата рождения", "Телефон",
                                     "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"])
    if df.empty:
        messagebox.showinfo("Статистика", "Нет данных")
        return

    counts = df["Группа"].value_counts()
    counts.plot(kind="bar")
    plt.title("Количество клиентов по группам")
    plt.show()


# ================== MAIN ==================
root = tk.Tk()
root.title("База клиентов")

search_entry = tk.Entry(root, width=40)
search_entry.grid(row=0, column=0, padx=5, pady=5)
tk.Button(root, text="Поиск", command=do_search).grid(row=0, column=1, padx=5, pady=5)

tree = ttk.Treeview(root, columns=("ID", "ФИО", "Дата рождения", "Телефон",
                                   "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"),
                    show="headings", height=20)
tree.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="nsew")

for col in tree["columns"]:
    tree.heading(col, text=col)

tk.Button(root, text="Добавить", command=add_window).grid(row=2, column=0, padx=5, pady=5)
tk.Button(root, text="Редактировать", command=edit_client).grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Удалить", command=delete_selected).grid(row=2, column=2, padx=5, pady=5)
tk.Button(root, text="Импорт Excel", command=import_excel).grid(row=2, column=3, padx=5, pady=5)
tk.Button(root, text="Экспорт Excel", command=export_excel).grid(row=2, column=4, padx=5, pady=5)
tk.Button(root, text="Статистика", command=show_statistics).grid(row=2, column=5, padx=5, pady=5)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

init_db()
root.after(200, refresh_tree)  # загружаем список чуть позже

root.mainloop()
