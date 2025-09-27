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
import sys

# ================== Пути ==================
APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
os.makedirs(APP_DIR, exist_ok=True)

DB_NAME = os.path.join(APP_DIR, "clients.db")
SHEET_ID = "1_DfTT8yzCjP0VH0PZu1Fz6FYMm1eRr7c0TmZU2DrH_w"


# ----------------------
# --- Утилиты ФИО ------
# ----------------------
def split_fio(fio: str):
    """Простое разделение ФИО на фамилию, имя, отчество.
    Возвращает кортеж (last, first, middle). Если частей меньше — middle = ''.
    """
    if not fio:
        return "", "", ""
    parts = fio.strip().split()
    if len(parts) == 1:
        return parts[0], "", ""
    if len(parts) == 2:
        return parts[0], parts[1], ""
    # >=3 — считаем первые три, остальные присоединим к middle через пробел (на случай составных отчеств)
    last = parts[0]
    first = parts[1]
    middle = " ".join(parts[2:])
    return last, first, middle


def join_fio(last, first, middle):
    parts = [p for p in (last or "", first or "", middle or "") if p and p.strip()]
    return " ".join(parts)


# ================== База данных ==================
def init_db():
    """Создаёт новую схему или мигрирует старую (если есть колонка fio)."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        # проверим структуру
        cur.execute("PRAGMA table_info(clients)")
        cols = [r[1] for r in cur.fetchall()]

        if not cols:
            # таблицы нет — создаём новый вариант
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS clients (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    last_name TEXT NOT NULL,
                    first_name TEXT NOT NULL,
                    middle_name TEXT,
                    dob TEXT NOT NULL,
                    phone TEXT,
                    contract_number TEXT,
                    ippcu_start TEXT,
                    ippcu_end TEXT,
                    group_name TEXT
                )
                """
            )
            # создаём индекс для проверки дублей
            cur.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS idx_clients_unique
                ON clients (lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)
                """
            )
            conn.commit()
            return

        # Если есть старая схема с fio — мигрируем
        if "fio" in cols and "last_name" not in cols:
            # создаём временную таблицу с новой схемой
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS clients_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    last_name TEXT NOT NULL,
                    first_name TEXT NOT NULL,
                    middle_name TEXT,
                    dob TEXT NOT NULL,
                    phone TEXT,
                    contract_number TEXT,
                    ippcu_start TEXT,
                    ippcu_end TEXT,
                    group_name TEXT
                )
                """
            )
            # переносим данные, разбивая fio
            cur.execute("SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name FROM clients")
            rows = cur.fetchall()
            for r in rows:
                _, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name = r
                last, first, middle = split_fio(fio or "")
                dob_val = dob or ""
                try:
                    cur.execute(
                        """
                        INSERT OR IGNORE INTO clients_new
                        (last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (last, first, middle, dob_val, phone, contract_number, ippcu_start, ippcu_end, group_name)
                    )
                except Exception:
                    cur.execute(
                        "INSERT OR IGNORE INTO clients_new (last_name, first_name, middle_name, dob) VALUES (?, ?, ?, ?)",
                        (last or "", first or "", middle or "", dob_val)
                    )
            # удаляем старую таблицу и переименовываем новую
            cur.execute("DROP TABLE clients")
            cur.execute("ALTER TABLE clients_new RENAME TO clients")
            # добавляем индекс
            cur.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS idx_clients_unique
                ON clients (lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)
                """
            )
            conn.commit()
            return

               # если уже новая схема — просто убеждаемся, что индекс есть
        if "last_name" in cols and "dob" in cols:
            # Удалим старый индекс, если он есть
            try:
                cur.execute("DROP INDEX IF EXISTS idx_clients_unique")
            except Exception:
                pass

            # Удалим дубли (оставляем запись с минимальным id)
            cur.execute(
                """
                DELETE FROM clients
                WHERE id NOT IN (
                    SELECT MIN(id) FROM clients
                    GROUP BY lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob
                )
                """
            )

            # Создадим новый уникальный индекс
            cur.execute(
                """
                CREATE UNIQUE INDEX idx_clients_unique
                ON clients (lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)
                """
            )
            conn.commit()
            return




def add_client(last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group):
    """Добавление с проверкой дублей (по ФИО+дата рождения, без учёта регистра)."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        # normalise empty middle to ''
        middle_name = middle_name or ""
        dob_val = dob or ""

        # проверка дубля
        cur.execute(
            """
            SELECT id FROM clients
            WHERE lower(last_name)=lower(?) AND lower(first_name)=lower(?) AND lower(COALESCE(middle_name,''))=lower(?) AND dob=?
            """,
            (last_name, first_name, middle_name, dob_val)
        )
        if cur.fetchone():
            raise ValueError(f"Клиент '{join_fio(last_name, first_name, middle_name)}' с датой рождения {dob_val} уже есть в базе.")

        cur.execute(
            """
            INSERT INTO clients (last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (last_name, first_name, middle_name, dob_val, phone, contract_number, ippcu_start, ippcu_end, group),
        )
        conn.commit()


def get_all_clients(limit=200):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            ORDER BY lower(last_name), lower(first_name)
            LIMIT ?
            """,
            (limit,),
        )
        return cur.fetchall()


def search_clients(query="", date_from=None, date_to=None, limit=200):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        q = (query or "").strip().lower()
        like = f"%{q}%"

        sql = """
            SELECT id, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE ( lower(last_name) LIKE ? OR lower(first_name) LIKE ? OR lower(COALESCE(middle_name,'')) LIKE ?
                   OR lower(contract_number) LIKE ? OR lower(phone) LIKE ? OR lower(COALESCE(group_name,'')) LIKE ? )
        """
        params = [like, like, like, like, like, like]

        if date_from:
            sql += " AND DATE(ippcu_end) >= DATE(?) "
            params.append(date_from)
        if date_to:
            sql += " AND DATE(ippcu_end) <= DATE(?) "
            params.append(date_to)

        sql += " ORDER BY lower(last_name), lower(first_name) LIMIT ?"
        params.append(limit)

        cur.execute(sql, params)
        return cur.fetchall()


def update_client(cid, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE clients
            SET last_name=?, first_name=?, middle_name=?, dob=?, phone=?, contract_number=?, ippcu_start=?, ippcu_end=?, group_name=?
            WHERE id=?
            """,
            (last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group, cid),
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

    # Если нет env — пробуем файл рядом с exe/скриптом
    if not creds_json:
        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        creds_path = os.path.join(exe_dir, "credentials.json")
        if not os.path.exists(creds_path):
            raise RuntimeError("Не найден GOOGLE_CREDENTIALS и нет файла credentials.json рядом с программой!")
        with open(creds_path, "r", encoding="utf-8") as f:
            creds_json = f.read()

    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
    return sheet


def import_from_gsheet():
    try:
        sheet = get_gsheet(SHEET_ID)
        data = sheet.get_all_records()

        added = 0
        for row in data:
            fio_raw = row.get("ФИО", "") or ""
            last, first, middle = split_fio(fio_raw)
            dob = row.get("Дата рождения", "") or ""
            phone = row.get("Телефон", "") or ""
            contract = row.get("Номер договора", "") or ""
            ippcu_start = row.get("Дата начала ИППСУ", "") or ""
            ippcu_end = row.get("Дата окончания ИППСУ", "") or ""
            group = row.get("Группа", "") or ""
            try:
                add_client(last, first, middle, dob, phone, contract, ippcu_start, ippcu_end, group)
                added += 1
            except ValueError:
                # дубликат — пропускаем
                continue
        refresh_tree()
        messagebox.showinfo("Успех", f"Импорт из Google Sheets завершён! Добавлено: {added}")
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
        # row: id, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
        cid, last, first, middle, dob, phone, contract, ippcu_start, ippcu_end, group = row
        fio = join_fio(last, first, middle)
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
        except Exception:
            pass

        tree.insert("", "end", values=(cid, fio, dob or "", phone or "", contract or "", ippcu_start or "", ippcu_end or "", group or ""), tags=(tag,))


def add_window():
    win = tk.Toplevel()
    win.title("Добавить обслуживаемого")

    # Фамилия / Имя / Отчество
    tk.Label(win, text="Фамилия").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_last = tk.Entry(win, width=30)
    e_last.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Имя").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_first = tk.Entry(win, width=30)
    e_first.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Отчество").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_middle = tk.Entry(win, width=30)
    e_middle.grid(row=2, column=1, padx=10, pady=5)

    # Дата рождения
    tk.Label(win, text="Дата рождения").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_dob.grid(row=3, column=1, padx=10, pady=5)

    # Остальные поля
    tk.Label(win, text="Телефон").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=30)
    e_phone.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=30)
    e_contract.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_start.grid(row=6, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата окончания ИППСУ").grid(row=7, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    e_ippcu_end.grid(row=7, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа").grid(row=8, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=30)
    e_group.grid(row=8, column=1, padx=10, pady=5)

    def save_client():
        last = e_last.get().strip()
        first = e_first.get().strip()
        middle = e_middle.get().strip()
        dob = e_dob.get_date().strftime("%Y-%m-%d")
        phone = e_phone.get().strip()
        contract_number = e_contract.get().strip()
        ippcu_start = e_ippcu_start.get_date().strftime("%Y-%m-%d")
        ippcu_end = e_ippcu_end.get_date().strftime("%Y-%m-%d")
        group = e_group.get().strip()

        if not last or not first or not dob:
            messagebox.showerror("Ошибка", "Поля 'Фамилия', 'Имя' и 'Дата рождения' обязательны!")
            return

        try:
            add_client(last, first, middle, dob, phone, contract_number, ippcu_start, ippcu_end, group)
            refresh_tree()
            win.destroy()
        except ValueError as ve:
            messagebox.showwarning("Дубликат", str(ve))
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось добавить:\n{e}")

    tk.Button(win, text="Сохранить", command=save_client).grid(row=9, column=0, columnspan=2, pady=10)


def edit_client():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Ошибка", "Выберите клиента для редактирования")
        return

    item = tree.item(selected[0])
    cid = item["values"][0]
    fio_display = item["values"][1]
    dob_display = item["values"][2]
    phone = item["values"][3]
    contract = item["values"][4]
    ippcu_start = item["values"][5]
    ippcu_end = item["values"][6]
    group = item["values"][7]

    # Получим полные поля из БД (чтобы точно знать last/first/middle)
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("SELECT last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name FROM clients WHERE id=?", (cid,))
        row = cur.fetchone()
        if not row:
            messagebox.showerror("Ошибка", "Клиент не найден в базе")
            return
        last, first, middle, dob_db, phone_db, contract_db, ippcu_start_db, ippcu_end_db, group_db = row

    win = tk.Toplevel()
    win.title("Редактировать обслуживаемого")

    tk.Label(win, text="Фамилия").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    e_last = tk.Entry(win, width=30)
    e_last.insert(0, last or "")
    e_last.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(win, text="Имя").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    e_first = tk.Entry(win, width=30)
    e_first.insert(0, first or "")
    e_first.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(win, text="Отчество").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    e_middle = tk.Entry(win, width=30)
    e_middle.insert(0, middle or "")
    e_middle.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата рождения").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    e_dob = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        e_dob.set_date(datetime.strptime(dob_db, "%Y-%m-%d"))
    except Exception:
        pass
    e_dob.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(win, text="Телефон").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    e_phone = tk.Entry(win, width=30)
    e_phone.insert(0, phone_db or "")
    e_phone.grid(row=4, column=1, padx=10, pady=5)

    tk.Label(win, text="Номер договора").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    e_contract = tk.Entry(win, width=30)
    e_contract.insert(0, contract_db or "")
    e_contract.grid(row=5, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата начала ИППСУ").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_start = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        if ippcu_start_db:
            e_ippcu_start.set_date(datetime.strptime(ippcu_start_db, "%Y-%m-%d"))
    except Exception:
        pass
    e_ippcu_start.grid(row=6, column=1, padx=10, pady=5)

    tk.Label(win, text="Дата окончания ИППСУ").grid(row=7, column=0, padx=10, pady=5, sticky="w")
    e_ippcu_end = DateEntry(win, width=27, date_pattern="dd.mm.yyyy")
    try:
        if ippcu_end_db:
            e_ippcu_end.set_date(datetime.strptime(ippcu_end_db, "%Y-%m-%d"))
    except Exception:
        pass
    e_ippcu_end.grid(row=7, column=1, padx=10, pady=5)

    tk.Label(win, text="Группа").grid(row=8, column=0, padx=10, pady=5, sticky="w")
    e_group = tk.Entry(win, width=30)
    e_group.insert(0, group_db or "")
    e_group.grid(row=8, column=1, padx=10, pady=5)

    def save_edit():
        new_last = e_last.get().strip()
        new_first = e_first.get().strip()
        new_middle = e_middle.get().strip()
        new_dob = e_dob.get_date().strftime("%Y-%m-%d")
        new_phone = e_phone.get().strip()
        new_contract = e_contract.get().strip()
        new_ippcu_start = e_ippcu_start.get_date().strftime("%Y-%m-%d")
        new_ippcu_end = e_ippcu_end.get_date().strftime("%Y-%m-%d")
        new_group = e_group.get().strip()

        if not new_last or not new_first or not new_dob:
            messagebox.showerror("Ошибка", "Поля 'Фамилия', 'Имя' и 'Дата рождения' обязательны!")
            return

        try:
            # Обновление — перед этим можно проверить на дубль (если изменилось ФИО/ДР)
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    SELECT id FROM clients
                    WHERE lower(last_name)=lower(?) AND lower(first_name)=lower(?) AND lower(COALESCE(middle_name,''))=lower(?) AND dob=? AND id<>?
                    """,
                    (new_last, new_first, new_middle or "", new_dob, cid)
                )
                if cur.fetchone():
                    messagebox.showwarning("Дубликат", "Есть другой клиент с таким же ФИО и датой рождения.")
                    return

            update_client(cid, new_last, new_first, new_middle, new_dob, new_phone, new_contract, new_ippcu_start, new_ippcu_end, new_group)
            refresh_tree()
            win.destroy()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    tk.Button(win, text="Сохранить изменения", command=save_edit).grid(row=9, column=0, columnspan=2, pady=10)


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

# Таблица (ID, ФИО, ДР, Телефон, Договор, Начало, Окончание, Группа)
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
