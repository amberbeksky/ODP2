# clients_app_full.py
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from datetime import datetime
import os
import traceback

DB_NAME = "clients.db"
BACKUP_FOLDER = "backup"

# -----------------------------
# --- база данных / utils -----
# -----------------------------
def init_db():
    """Создаёт таблицы, а если старая таблица без ippcu_end — добавляет колонку."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("""
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
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER,
            action TEXT,
            timestamp TEXT
        )
        """)
        # если по каким-то причинам старая таблица не имела ippcu_end, добавим
        cur.execute("PRAGMA table_info(clients)")
        cols = [row[1] for row in cur.fetchall()]
        if "ippcu_end" not in cols:
            try:
                cur.execute("ALTER TABLE clients ADD COLUMN ippcu_end TEXT")
            except Exception:
                # если не получилось — не фатально
                pass
        cur.execute("CREATE INDEX IF NOT EXISTS idx_contract ON clients(contract_number)")
        conn.commit()

def log_action(client_id, action):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO logs (client_id, action, timestamp) VALUES (?, ?, ?)",
                (client_id, action, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
            conn.commit()
    except Exception:
        traceback.print_exc()

def add_client(fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name=""):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO clients (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
            )
            client_id = cur.lastrowid
            conn.commit()
        log_action(client_id, "Добавлен обслуживаемый")
        return client_id
    except Exception:
        traceback.print_exc()
        raise

def update_client(client_id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name=""):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE clients
                SET fio=?, dob=?, phone=?, contract_number=?, ippcu_start=?, ippcu_end=?, group_name=?
                WHERE id=?
            """, (fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name, client_id))
            conn.commit()
        log_action(client_id, "Обновлён обслуживаемый")
    except Exception:
        traceback.print_exc()
        raise

def delete_client(client_id):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM clients WHERE id=?", (client_id,))
            conn.commit()
        log_action(client_id, "Удалён обслуживаемый")
    except Exception:
        traceback.print_exc()
        raise

def search_clients(query=""):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        like = f"%{query}%"
        cur.execute("""
            SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE fio LIKE ? OR contract_number LIKE ? OR phone LIKE ? OR ippcu_start LIKE ? OR ippcu_end LIKE ? OR group_name LIKE ?
            ORDER BY fio
        """, (like, like, like, like, like, like))
        rows = cur.fetchall()
    return rows

def get_client_by_id(client_id):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE id=?
        """, (client_id,))
        return cur.fetchone()

# -----------------------------
# --- утилиты работы с датой --
# -----------------------------
def normalize_date_for_db(value):
    """
    Принимает строку/None/pandas.Timestamp и возвращает строку 'YYYY-MM-DD' либо ''.
    Без исключения — в случае некорректной даты возвращает пустую строку.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return ""
        # пробуем стандартный ISO
        try:
            d = datetime.strptime(value, "%Y-%m-%d")
            return d.strftime("%Y-%m-%d")
        except Exception:
            pass
        # пробуем pandas parsing (устойчивее к разным форматам)
        try:
            parsed = pd.to_datetime(value, dayfirst=False, errors='coerce')
            if pd.isna(parsed):
                return ""
            return parsed.strftime("%Y-%m-%d")
        except Exception:
            return ""
    # pandas.Timestamp or datetime
    try:
        return pd.to_datetime(value, errors='coerce').strftime("%Y-%m-%d")
    except Exception:
        return ""

def parse_date_for_display(value):
    """Пытается вернуть красивую строку для отображения, если пустая или некорректная — возвращает ''. """
    if not value:
        return ""
    try:
        d = pd.to_datetime(value, errors='coerce')
        if pd.isna(d):
            return value  # вернём оригинал, на случай пользователь ввёл нестандартно
        return d.strftime("%Y-%m-%d")
    except Exception:
        return value

# -----------------------------
# --- интерфейс -------------
# -----------------------------
def refresh_tree(results=None):
    for row in tree.get_children():
        tree.delete(row)
    if results is None:
        results = search_clients()
    for r in results:
        # r: (id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
        ippcu_date = None
        tag = ""
        try:
            if r[5]:
                # используем pandas для парсинга (более гибко)
                parsed = pd.to_datetime(r[5], errors='coerce')
                if not pd.isna(parsed):
                    delta = (parsed - pd.Timestamp.now()).days
                    if delta < 0:
                        tag = "red"
                    elif delta <= 30:
                        tag = "orange"
                    else:
                        tag = "green"
        except Exception:
            tag = ""
        values = (
            r[1] or "",
            parse_date_for_display(r[2]),
            r[3] or "",
            r[4] or "",
            parse_date_for_display(r[5]),
            parse_date_for_display(r[6]),
            r[7] or ""
        )
        tree.insert("", "end", text=str(r[0]), values=values, tags=(tag,))

def add_window():
    win = tk.Toplevel()
    win.title("Добавить обслуживаемого")

    # ФИО
    tk.Label(win, text="ФИО").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=30)
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    # Дата рождения
    tk.Label(win, text="Дата рождения").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    # Телефон
    tk.Label(win, text="Телефон").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=30)
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    # Номер договора
    tk.Label(win, text="Номер договора").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=30)
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    # Дата начала ИППСУ
    tk.Label(win, text="Дата начала ИППСУ").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_start.grid(row=4, column=1, padx=10, pady=5)

    # Дата окончания ИППСУ
    tk.Label(win, text="Дата окончания ИППСУ").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_end.grid(row=5, column=1, padx=10, pady=5)

    # Группа
    tk.Label(win, text="Группа").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=30)
    e_group.grid(row=6, column=1, padx=10, pady=5)

    # Сохранить
    def save_client():
        fio = e_fio.get().strip()
        dob = e_dob.get_date().strftime("%d.%m.%Y")
        phone = e_phone.get().strip()
        contract = e_contract.get().strip()
        ippcu_start = e_ippcu_start.get_date().strftime("%d.%m.%Y")
        ippcu_end = e_ippcu_end.get_date().strftime("%d.%m.%Y")
        group = e_group.get().strip()

        if not fio:
            messagebox.showerror("Ошибка", "Поле 'ФИО' обязательно для заполнения!")
            return

        c.execute(
            "INSERT INTO clients (fio, dob, phone, contract, ippcu_start, ippcu_end, group_name) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (fio, dob, phone, contract, ippcu_start, ippcu_end, group)
        )
        conn.commit()
        refresh_clients()
        win.destroy()

    tk.Button(win, text="Сохранить", command=save_client).grid(row=7, column=0, columnspan=2, pady=10)


def edit_client(event=None):
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Выберите обслуживаемого", "Сначала выберите обслуживаемого для редактирования")
        return
    client_id = tree.item(selected[0], "text")
    client = get_client_by_id(client_id)
    if not client:
        messagebox.showerror("Ошибка", "Не удалось получить данные клиента из базы.")
        return
    # client: (id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
    def save():
        fio_val = e_fio.get().strip()
        if not fio_val:
            messagebox.showerror("Ошибка", "Введите ФИО")
            return
        try:
            dob_val = normalize_date_for_db(e_dob.get())
            ippcu_start_val = normalize_date_for_db(e_ippcu_start.get())
            ippcu_end_val = normalize_date_for_db(e_ippcu_end.get())
            phone_val = e_phone.get().strip()
            contract_val = e_contract.get().strip()
            group_val = e_group.get().strip()

            update_client(client_id, fio_val, dob_val, phone_val, contract_val, ippcu_start_val, ippcu_end_val, group_val)
            messagebox.showinfo("Успех", "Данные обслуживаемого обновлены!")
            status_var.set(f"Обновлён ID {client_id}")
            win.destroy()
            refresh_tree()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось обновить:\n{e}")

    win = tk.Toplevel(root)
    win.title("Редактировать обслуживаемого")
    win.configure(bg="#f5f5f7")

    tk.Label(win, text="ФИО", bg="#f5f5f7").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=40)
    e_fio.insert(0, client[1] or "")
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения", bg="#f5f5f7").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=37, date_pattern='yyyy-mm-dd')
    try:
        if client[2]:
            e_dob.set_date(parse_date_for_display(client[2]))
    except Exception:
        pass
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон", bg="#f5f5f7").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=40)
    e_phone.insert(0, client[3] or "")
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора", bg="#f5f5f7").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=40)
    e_contract.insert(0, client[4] or "")
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ", bg="#f5f5f7").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=37, date_pattern='yyyy-mm-dd')
    try:
        if client[5]:
            e_ippcu_start.set_date(parse_date_for_display(client[5]))
    except Exception:
        pass
    e_ippcu_start.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата окончания ИППСУ", bg="#f5f5f7").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=37, date_pattern='yyyy-mm-dd')
    try:
        if client[6]:
            e_ippcu_end.set_date(parse_date_for_display(client[6]))
    except Exception:
        pass
    e_ippcu_end.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа", bg="#f5f5f7").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=40)
    e_group.insert(0, client[7] or "")
    e_group.grid(row=6, column=1, padx=10, pady=5)

    tk.Button(win, text="Сохранить", command=save, bg="#007aff", fg="#ffffff").grid(row=7, columnspan=2, pady=10)

def remove_client():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Выберите обслуживаемого", "Сначала выберите обслуживаемого для удаления")
        return
    client_id = tree.item(selected[0], "text")
    if messagebox.askyesno("Удалить", "Вы уверены, что хотите удалить этого обслуживаемого?"):
        try:
            delete_client(client_id)
            status_var.set(f"Удалён ID {client_id}")
            refresh_tree()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось удалить:\n{e}")

def do_search():
    query = search_entry.get().strip()
    results = search_clients(query)
    refresh_tree(results)
    status_var.set(f"Найдено: {len(results)}")

def import_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path)
        added = 0
        for _, row in df.iterrows():
            fio = str(row.get('ФИО', '')).strip()
            if not fio:
                continue
            dob = normalize_date_for_db(row.get('Дата рождения', ''))
            phone = str(row.get('Телефон', '')).strip()
            contract = str(row.get('Номер договора', '')).strip()
            ippcu_start = normalize_date_for_db(row.get('Дата начала ИППСУ', ''))
            ippcu_end = normalize_date_for_db(row.get('Дата окончания ИППСУ', ''))
            group = str(row.get('Группа', '')).strip()

            add_client(fio, dob, phone, contract, ippcu_start, ippcu_end, group)
            added += 1
        messagebox.showinfo("Успех", f"Импортировано записей: {added}")
        status_var.set(f"Импортировано {added} записей")
        refresh_tree()
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось импортировать файл:\n{e}")

def export_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    try:
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql_query(
                "SELECT fio AS 'ФИО', dob AS 'Дата рождения', phone AS 'Телефон', contract_number AS 'Номер договора', ippcu_start AS 'Дата начала ИППСУ', ippcu_end AS 'Дата окончания ИППСУ', group_name AS 'Группа' FROM clients",
                conn
            )
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
        status_var.set(f"Экспорт в {os.path.basename(file_path)}")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось экспортировать данные:\n{e}")

def backup_database():
    if not os.path.exists(BACKUP_FOLDER):
        os.makedirs(BACKUP_FOLDER)
    backup_file = os.path.join(BACKUP_FOLDER, f"clients_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    try:
        with sqlite3.connect(DB_NAME) as conn:
            with sqlite3.connect(backup_file) as bck:
                conn.backup(bck)
        messagebox.showinfo("Успех", f"База скопирована в {backup_file}")
        status_var.set(f"Резервная копия: {os.path.basename(backup_file)}")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось создать резервную копию:\n{e}")

def about_window():
    win = tk.Toplevel(root)
    win.title("О программе")
    win.geometry("500x250")
    win.configure(bg="#f5f5f7")
    win.resizable(False, False)

    tk.Label(win, text="База обслуживаемых (полустационар)", font=("San Francisco", 16, "bold"),
             bg="#f5f5f7", fg="#1d1d1f").pack(pady=(20,10))
    tk.Label(win, text="👨‍💻", font=("San Francisco", 40), bg="#f5f5f7").pack(pady=5)

    info_text = ("Разработчик: Зеленков Данил Вадимович\n"
                 "Младший администратор базы данных\n"
                 "ГБУ КЦСОН Варнавинского муниципального округа\n"
                 "Версия программы: 1.0 (обновлено)")
    tk.Label(win, text=info_text, font=("San Francisco", 12), bg="#f5f5f7", fg="#1d1d1f", justify="center").pack(pady=10)

def show_statistics():
    results = search_clients()
    df = pd.DataFrame(results, columns=["id","ФИО","Дата рождения","Телефон","Номер договора","Дата начала ИППСУ","Дата окончания ИППСУ","Группа"])
    if df.empty:
        messagebox.showinfo("Статистика", "Нет данных для построения графика")
        return
    df['Дата начала ИППСУ'] = pd.to_datetime(df['Дата начала ИППСУ'], errors='coerce')
    df = df.dropna(subset=['Дата начала ИППСУ'])
    if df.empty:
        messagebox.showinfo("Статистика", "Нет корректных дат для построения графика")
        return
    df['Месяц'] = df['Дата начала ИППСУ'].dt.to_period('M').astype(str)
    counts = df.groupby('Месяц').size().sort_index()
    counts.plot(kind='bar')
    plt.title("Количество обслуживаемых по месяцам начала ИППСУ")
    plt.xlabel("Месяц")
    plt.ylabel("Количество")
    plt.tight_layout()
    plt.show()
    status_var.set("Показана статистика")

# -----------------------------
# --- главное окно ----------
# -----------------------------
root = tk.Tk()
root.title("База обслуживаемых (полустационар)")
root.geometry("1100x700")
root.configure(bg="#f5f5f7")

frame = tk.Frame(root, bg="#f5f5f7")
frame.pack(pady=10, fill="x", padx=10)

search_entry = tk.Entry(frame, width=50)
search_entry.grid(row=0, column=0, padx=5, sticky="w")
tk.Button(frame, text="Поиск", command=do_search, bg="#007aff", fg="#ffffff", width=12).grid(row=0, column=1, padx=5)
tk.Button(frame, text="Добавить обслуживаемого", command=add_window, bg="#34c759", fg="#ffffff", width=18).grid(row=0, column=2, padx=5)
tk.Button(frame, text="Редактировать", command=edit_client, bg="#ff9500", fg="#ffffff", width=12).grid(row=0, column=3, padx=5)
tk.Button(frame, text="Удалить", command=remove_client, bg="#ff3b30", fg="#ffffff", width=10).grid(row=0, column=4, padx=5)
tk.Button(frame, text="Импорт Excel", command=import_excel, bg="#5856d6", fg="#ffffff", width=12).grid(row=0, column=5, padx=5)
tk.Button(frame, text="Экспорт Excel", command=export_excel, bg="#5ac8fa", fg="#ffffff", width=12).grid(row=0, column=6, padx=5)
tk.Button(frame, text="Резервная копия", command=backup_database, bg="#ffcc00", fg="#ffffff", width=14).grid(row=0, column=7, padx=5)
tk.Button(frame, text="Статистика", command=show_statistics, bg="#ff2d55", fg="#ffffff", width=12).grid(row=0, column=8, padx=5)
tk.Button(frame, text="О программе", command=about_window, bg="#8e8e93", fg="#ffffff", width=10).grid(row=0, column=9, padx=5)

cols = ("ФИО", "Дата рождения", "Телефон", "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа")
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

# scrollbars
vsb = ttk.Scrollbar(tree_frame, orient="vertical")
hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
tree = ttk.Treeview(tree_frame, columns=cols, show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set)
vsb.config(command=tree.yview)
hsb.config(command=tree.xview)
vsb.pack(side="right", fill="y")
hsb.pack(side="bottom", fill="x")
tree.pack(fill="both", expand=True)

for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor="w")

# теги подсветки
tree.tag_configure('red', background='#ffcccc')
tree.tag_configure('orange', background='#ffe5b4')
tree.tag_configure('green', background='#ccffcc')

# bind double click to edit
tree.bind("<Double-1>", edit_client)

# статус-бар
status_var = tk.StringVar(value="Готово")
status_bar = tk.Label(root, textvariable=status_var, anchor="w", bg="#f5f5f7")
status_bar.pack(fill="x", side="bottom")

# инициализация
init_db()
refresh_tree()

root.mainloop()
