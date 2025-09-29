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
from docx import Document   # <--- добавил для экспорта в Word

# ================== Пути ==================
APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
os.makedirs(APP_DIR, exist_ok=True)

DB_NAME = os.path.join(APP_DIR, "clients.db")
SHEET_ID = "1_DfTT8yzCjP0VH0PZu1Fz6FYMm1eRr7c0TmZU2DrH_w"


# ----------------------
# --- Утилиты ФИО ------
# ----------------------
def split_fio(fio: str):
    if not fio:
        return "", "", ""
    parts = fio.strip().split()
    if len(parts) == 1:
        return parts[0], "", ""
    if len(parts) == 2:
        return parts[0], parts[1], ""
    last = parts[0]
    first = parts[1]
    middle = " ".join(parts[2:])
    return last, first, middle


def join_fio(last, first, middle):
    parts = [p for p in (last or "", first or "", middle or "") if p and p.strip()]
    return " ".join(parts)


# ================== Экспорт в Word ==================
def export_selected_to_word():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Ошибка", "Выберите хотя бы одного клиента")
        return

    doc = Document()
    doc.add_heading("Список обслуживаемых", level=1)

    for i, item in enumerate(selected, start=1):
        values = tree.item(item)["values"]
        # values = (ID, Фамилия, Имя, Отчество, ДР, Телефон, Договор, Начало, Окончание, Группа)
        fio = " ".join(v for v in [values[1], values[2], values[3]] if v)
        dob = values[4]
        doc.add_paragraph(f"{i}. {fio} – {dob}")

    file_path = os.path.join(APP_DIR, "список.docx")
    doc.save(file_path)

    messagebox.showinfo("Готово", f"Список сохранён:\n{file_path}")

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
                    group_name TEXT,
                    UNIQUE(lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)
                )
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
                    group_name TEXT,
                    UNIQUE(lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)
                )
                """
            )
            # переносим данные, разбивая fio
            cur.execute("SELECT id, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name FROM clients")
            rows = cur.fetchall()
            for r in rows:
                _, fio, dob, phone, contract_number, ippcu_start, ippcu_end, group_name = r
                last, first, middle = split_fio(fio or "")
                # Для переносимых строк, если dob пуст — ставим '' (требуем dob NOT NULL, но чтобы не упало)
                dob_val = dob or ""
                try:
                    cur.execute(
                        """
                        INSERT OR IGNORE INTO clients_new
                        (id, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (None, last, first, middle, dob_val, phone, contract_number, ippcu_start, ippcu_end, group_name)
                    )
                except Exception:
                    # если какие-то данные некорректны — вставим минимально
                    cur.execute(
                        "INSERT OR IGNORE INTO clients_new (last_name, first_name, middle_name, dob) VALUES (?, ?, ?, ?)",
                        (last or "", first or "", middle or "", dob_val)
                    )
            # удаляем старую таблицу и переименовываем новую
            cur.execute("DROP TABLE clients")
            cur.execute("ALTER TABLE clients_new RENAME TO clients")
            conn.commit()
            return

        # если уже новая схема — ничего не делаем
        if "last_name" in cols and "dob" in cols:
            # убедимся, что UNIQUE индекс существует (на случай более старых версий)
            try:
                cur.execute(
                    "CREATE UNIQUE INDEX IF NOT EXISTS idx_clients_unique ON clients (lower(last_name), lower(first_name), lower(COALESCE(middle_name,'')), dob)"
                )
            except Exception:
                pass
            conn.commit()
            return

        # В иных случаях — попытка создать недостающие колонки (на всякий случай)
        # (этот блок — запасной; чаще всего не нужен)
        try:
            if "last_name" not in cols:
                cur.execute("ALTER TABLE clients ADD COLUMN last_name TEXT")
            if "first_name" not in cols:
                cur.execute("ALTER TABLE clients ADD COLUMN first_name TEXT")
            if "middle_name" not in cols:
                cur.execute("ALTER TABLE clients ADD COLUMN middle_name TEXT")
            conn.commit()
        except Exception:
            pass


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
        cid, last, first, middle, dob, phone, contract, ippcu_start, ippcu_end, group = row
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

        tree.insert("", "end", values=(
            cid, last or "", first or "", middle or "",
            dob or "", phone or "", contract or "",
            ippcu_start or "", ippcu_end or "", group or ""
        ), tags=(tag,))



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
    """Окно редактирования клиента"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Ошибка", "Выберите клиента для редактирования")
        return

    values = tree.item(selected[0], "values")
    cid = values[0]  # ID клиента

    # Теперь берем Фамилия / Имя / Отчество из отдельных колонок
    last, first, middle = values[1], values[2], values[3]
    dob, phone, contract = values[4], values[5], values[6]
    ippcu_start, ippcu_end, group = values[7], values[8], values[9]

    win = tk.Toplevel(root)
    win.title("Редактировать клиента")

    tk.Label(win, text="Фамилия").grid(row=0, column=0)
    last_entry = tk.Entry(win)
    last_entry.insert(0, last)
    last_entry.grid(row=0, column=1)

    tk.Label(win, text="Имя").grid(row=1, column=0)
    first_entry = tk.Entry(win)
    first_entry.insert(0, first)
    first_entry.grid(row=1, column=1)

    tk.Label(win, text="Отчество").grid(row=2, column=0)
    middle_entry = tk.Entry(win)
    middle_entry.insert(0, middle)
    middle_entry.grid(row=2, column=1)

    tk.Label(win, text="Дата рождения (ГГГГ-ММ-ДД)").grid(row=3, column=0)
    dob_entry = tk.Entry(win)
    dob_entry.insert(0, dob)
    dob_entry.grid(row=3, column=1)

    tk.Label(win, text="Телефон").grid(row=4, column=0)
    phone_entry = tk.Entry(win)
    phone_entry.insert(0, phone)
    phone_entry.grid(row=4, column=1)

    tk.Label(win, text="Номер договора").grid(row=5, column=0)
    contract_entry = tk.Entry(win)
    contract_entry.insert(0, contract)
    contract_entry.grid(row=5, column=1)

    tk.Label(win, text="Дата начала ИППСУ").grid(row=6, column=0)
    start_entry = tk.Entry(win)
    start_entry.insert(0, ippcu_start)
    start_entry.grid(row=6, column=1)

    tk.Label(win, text="Дата окончания ИППСУ").grid(row=7, column=0)
    end_entry = tk.Entry(win)
    end_entry.insert(0, ippcu_end)
    end_entry.grid(row=7, column=1)

    tk.Label(win, text="Группа").grid(row=8, column=0)
    group_entry = tk.Entry(win)
    group_entry.insert(0, group)
    group_entry.grid(row=8, column=1)

    def save_changes():
        update_client(cid,
                      last_entry.get(), first_entry.get(), middle_entry.get(),
                      dob_entry.get(), phone_entry.get(), contract_entry.get(),
                      start_entry.get(), end_entry.get(), group_entry.get())
        refresh_tree()
        win.destroy()

    tk.Button(win, text="Сохранить", command=save_changes).grid(row=9, column=0, columnspan=2, pady=10)


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

    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        q = (query or "").strip().lower()
        like = f"%{q}%"

        # Основной запрос
        sql = """
            SELECT id, last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name
            FROM clients
            WHERE (
                lower(last_name) LIKE ?
                OR lower(first_name) LIKE ?
                OR lower(COALESCE(middle_name,'')) LIKE ?
                OR lower(last_name || ' ' || first_name || ' ' || COALESCE(middle_name,'')) LIKE ?
                OR lower(contract_number) LIKE ?
                OR lower(phone) LIKE ?
                OR lower(COALESCE(group_name,'')) LIKE ?
            )
        """
        params = [like, like, like, like, like, like, like]

        # Фильтры по датам окончания
        if date_from:
            sql += " AND DATE(ippcu_end) >= DATE(?) "
            params.append(date_from)
        if date_to:
            sql += " AND DATE(ippcu_end) <= DATE(?) "
            params.append(date_to)

        sql += " ORDER BY lower(last_name), lower(first_name) LIMIT ?"
        params.append(200)

        cur.execute(sql, params)
        results = cur.fetchall()

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
tree = ttk.Treeview(
    root,
    columns=("ID", "Фамилия", "Имя", "Отчество", "Дата рождения", "Телефон",
             "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"),
    show="headings",
    height=20
)
tree.grid(row=1, column=0, columnspan=7, padx=5, pady=5, sticky="nsew")

for col in tree["columns"]:
    tree.heading(col, text=col)

tree.tag_configure("expired", background="#ffcccc")
tree.tag_configure("soon", background="#fff2cc")
tree.tag_configure("active", background="#ccffcc")

# Кнопки
tk.Button(root, text="Добавить", command=add_window).grid(row=2, column=0, padx=5, pady=5)
tk.Button(root, text="Редактировать", command=edit_client).grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Удалить", command=delete_selected).grid(row=2, column=2, padx=5, pady=5)
tk.Button(root, text="Импорт Google Sheets", command=import_from_gsheet).grid(row=2, column=3, padx=5, pady=5)
tk.Button(root, text="Экспорт в Word", command=export_selected_to_word).grid(row=2, column=4, padx=5, pady=5)  # <--- новая кнопка

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

init_db()
root.after(200, refresh_tree)

root.mainloop()
