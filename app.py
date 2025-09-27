import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import traceback

DB_NAME = "clients.db"
BACKUP_FOLDER = "backup"

# --- база данных ---
def init_db():
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
        # индекс для быстрого поиска по договору
        cur.execute("CREATE INDEX IF NOT EXISTS idx_contract ON clients(contract_number)")
        conn.commit()

def log_action(client_id, action):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO logs (client_id, action, timestamp) VALUES (?, ?, ?)",
                        (client_id, action, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
    except Exception:
        # не фатально — просто напечатаем в консоль
        traceback.print_exc()

def add_client(fio, dob, phone, contract_number, ippcu_start, group_name=""):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO clients (fio, dob, phone, contract_number, ippcu_start, group_name) VALUES (?, ?, ?, ?, ?, ?)",
                (fio, dob, phone, contract_number, ippcu_start, group_name)
            )
            client_id = cur.lastrowid
            conn.commit()
        log_action(client_id, "Добавлен обслуживаемый")
        return client_id
    except Exception:
        traceback.print_exc()
        raise

def update_client(client_id, fio, dob, phone, contract_number, ippcu_start, group_name=""):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE clients
                SET fio=?, dob=?, phone=?, contract_number=?, ippcu_start=?, group_name=?
                WHERE id=?
            """, (fio, dob, phone, contract_number, ippcu_start, group_name, client_id))
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
            SELECT id, fio, dob, phone, contract_number, ippcu_start, group_name
            FROM clients
            WHERE fio LIKE ? OR contract_number LIKE ? OR phone LIKE ? OR ippcu_start LIKE ? OR group_name LIKE ?
            ORDER BY fio
        """, (like, like, like, like, like))
        rows = cur.fetchall()
    return rows

# --- интерфейс ---
def refresh_tree(results=None):
    for row in tree.get_children():
        tree.delete(row)
    if results is None:
        results = search_clients()
    for r in results:
        # Цветовые метки по дате начала ИППСУ
        ippcu_date = None
        tag = ""
        try:
            if r[5]:
                ippcu_date = datetime.strptime(r[5], "%Y-%m-%d")
                delta = (ippcu_date - datetime.now()).days
                if delta < 0:
                    tag = "red"
                elif delta <= 30:
                    tag = "orange"
                else:
                    tag = "green"
        except Exception:
            # не фатально, оставим без тега
            tag = ""
        # text хранит id (скрытая колонка), values — отображаемые поля
        tree.insert("", "end", text=str(r[0]), values=r[1:], tags=(tag,))

def add_window():
    def save():
        fio_val = e_fio.get().strip()
        if not fio_val:
            messagebox.showerror("Ошибка", "Введите ФИО")
            return
        try:
            dob_val = e_dob.get()
            ippcu_val = e_ippcu.get()
            client_id = add_client(fio_val, dob_val, e_phone.get().strip(), e_contract.get().strip(), ippcu_val, e_group.get().strip())
            messagebox.showinfo("Успех", "Обслуживаемый добавлен!")
            status_var.set(f"Добавлен ID {client_id}")
            win.destroy()
            refresh_tree()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось добавить:\n{e}")

    win = tk.Toplevel(root)
    win.title("Добавить обслуживаемого")
    win.configure(bg="#f5f5f7")

    tk.Label(win, text="ФИО", bg="#f5f5f7").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=35)
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения", bg="#f5f5f7").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=32, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон", bg="#f5f5f7").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=35)
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора", bg="#f5f5f7").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=35)
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ", bg="#f5f5f7").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu = DateEntry(win, width=32, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
    e_ippcu.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа", bg="#f5f5f7").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=35)
    e_group.grid(row=5, column=1, padx=10, pady=5)

    tk.Button(win, text="Сохранить", command=save, bg="#007aff", fg="#ffffff").grid(row=6, columnspan=2, pady=10)

def edit_client(event=None):
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Выберите обслуживаемого", "Сначала выберите обслуживаемого для редактирования")
        return
    client_data = tree.item(selected[0], "values")
    client_id = tree.item(selected[0], "text")

    def save():
        fio_val = e_fio.get().strip()
        if not fio_val:
            messagebox.showerror("Ошибка", "Введите ФИО")
            return
        try:
            update_client(client_id, e_fio.get(), e_dob.get(), e_phone.get(), e_contract.get(), e_ippcu.get(), e_group.get())
            messagebox.showinfo("Успех", "Данные обслуживаемого обновлены!")
            status_var.set(f"Обновлён ID {client_id}")
            win.destroy()
            refresh_tree()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить:\n{e}")

    win = tk.Toplevel(root)
    win.title("Редактировать обслуживаемого")
    win.configure(bg="#f5f5f7")

    tk.Label(win, text="ФИО", bg="#f5f5f7").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_fio = tk.Entry(win, width=35)
    e_fio.insert(0, client_data[0])
    e_fio.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения", bg="#f5f5f7").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=32, date_pattern='yyyy-mm-dd')
    try:
        e_dob.set_date(client_data[1])
    except Exception:
        pass
    e_dob.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон", bg="#f5f5f7").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=35)
    e_phone.insert(0, client_data[2])
    e_phone.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора", bg="#f5f5f7").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=35)
    e_contract.insert(0, client_data[3])
    e_contract.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ", bg="#f5f5f7").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_ippcu = DateEntry(win, width=32, date_pattern='yyyy-mm-dd')
    try:
        e_ippcu.set_date(client_data[4])
    except Exception:
        pass
    e_ippcu.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа", bg="#f5f5f7").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=35)
    e_group.insert(0, client_data[5])
    e_group.grid(row=5, column=1, padx=10, pady=5)

    tk.Button(win, text="Сохранить", command=save, bg="#007aff", fg="#ffffff").grid(row=6, columnspan=2, pady=10)

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
        required_cols = {
            'ФИО': 'fio',
            'Дата рождения': 'dob',
            'Телефон': 'phone',
            'Номер договора': 'contract_number',
            'Дата начала ИППСУ': 'ippcu_start',
            'Группа': 'group_name'
        }
        added = 0
        for _, row in df.iterrows():
            fio = str(row.get('ФИО', '')).strip()
            if not fio:
                continue
            dob_raw = row.get('Дата рождения', '')
            ippcu_raw = row.get('Дата начала ИППСУ', '')
            # корректная строковая форма даты в формате yyyy-mm-dd, если возможно
            def norm_date(val):
                if pd.isna(val):
                    return ""
                if isinstance(val, pd.Timestamp):
                    return val.strftime("%Y-%m-%d")
                if isinstance(val, str):
                    try:
                        parsed = pd.to_datetime(val, dayfirst=False, errors='coerce')
                        if pd.isna(parsed):
                            return val
                        return parsed.strftime("%Y-%m-%d")
                    except Exception:
                        return val
                try:
                    parsed = pd.to_datetime(val, errors='coerce')
                    if pd.isna(parsed):
                        return ""
                    return parsed.strftime("%Y-%m-%d")
                except Exception:
                    return str(val)
            dob = norm_date(dob_raw)
            ippcu = norm_date(ippcu_raw)
            add_client(
                fio,
                dob,
                str(row.get('Телефон', '')).strip(),
                str(row.get('Номер договора', '')).strip(),
                ippcu,
                str(row.get('Группа', '')).strip()
            )
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
            df = pd.read_sql_query("SELECT fio AS 'ФИО', dob AS 'Дата рождения', phone AS 'Телефон', contract_number AS 'Номер договора', ippcu_start AS 'Дата начала ИППСУ', group_name AS 'Группа' FROM clients", conn)
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
    df = pd.DataFrame(results, columns=["id","ФИО","Дата рождения","Телефон","Номер договора","Дата начала ИППСУ","Группа"])
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
    # простой столбчатый график
    counts.plot(kind='bar')
    plt.title("Количество обслуживаемых по месяцам начала ИППСУ")
    plt.xlabel("Месяц")
    plt.ylabel("Количество")
    plt.tight_layout()
    plt.show()
    status_var.set("Показана статистика")

# --- главное окно ---
root = tk.Tk()
root.title("База обслуживаемых (полустационар)")
root.geometry("1000x600")
root.configure(bg="#f5f5f7")

frame = tk.Frame(root, bg="#f5f5f7")
frame.pack(pady=10, fill="x", padx=10)

search_entry = tk.Entry(frame, width=40)
search_entry.grid(row=0, column=0, padx=5, sticky="w")
tk.Button(frame, text="Поиск", command=do_search, bg="#007aff", fg="#ffffff").grid(row=0, column=1, padx=5)
tk.Button(frame, text="Добавить обслуживаемого", command=add_window, bg="#34c759", fg="#ffffff").grid(row=0, column=2, padx=5)
tk.Button(frame, text="Редактировать", command=edit_client, bg="#ff9500", fg="#ffffff").grid(row=0, column=3, padx=5)
tk.Button(frame, text="Удалить", command=remove_client, bg="#ff3b30", fg="#ffffff").grid(row=0, column=4, padx=5)
tk.Button(frame, text="Импорт Excel", command=import_excel, bg="#5856d6", fg="#ffffff").grid(row=0, column=5, padx=5)
tk.Button(frame, text="Экспорт Excel", command=export_excel, bg="#5ac8fa", fg="#ffffff").grid(row=0, column=6, padx=5)
tk.Button(frame, text="Резервная копия", command=backup_database, bg="#ffcc00", fg="#ffffff").grid(row=0, column=7, padx=5)
tk.Button(frame, text="Статистика", command=show_statistics, bg="#ff2d55", fg="#ffffff").grid(row=0, column=8, padx=5)
tk.Button(frame, text="О программе", command=about_window, bg="#8e8e93", fg="#ffffff").grid(row=0, column=9, padx=5)

cols = ("ФИО", "Дата рождения", "Телефон", "Номер договора", "Дата начала ИППСУ", "Группа")
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
