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
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import simpledialog


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


def export_selected_to_word():
    selected_items = []
    for row_id in tree.get_children():
        values = tree.item(row_id, "values")
        if values and values[0] == "X":  # галочка стоит
            selected_items.append(values)

    if not selected_items:
        messagebox.showerror("Ошибка", "Отметьте галочками хотя бы одного клиента")
        return

    # Запрос названия смены
    shift_name = simpledialog.askstring("Смена", "Введите название смены (например: 11 смена)")
    if not shift_name:
        return

    # Запрос дат
    date_range = simpledialog.askstring("Даты", "Введите период (например: с 01.10.2024 по 15.10.2024)")
    if not date_range:
        return

    doc = Document()

    # === Заголовок ===
    heading = doc.add_paragraph(f"{shift_name} {date_range}")
    heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = heading.runs[0]
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph("")  # отступ после заголовка

    # === Нумерованный список выбранных клиентов ===
    for i, values in enumerate(selected_items, start=1):
        last = values[2]
        first = values[3]
        middle = values[4]
        dob = values[5]

        fio = " ".join(v for v in [last, first, middle] if v)
        p = doc.add_paragraph(f"{i}. {fio} – {dob} г.р.")
        p.runs[0].font.size = Pt(12)

    # === «Пружина» для смещения итогов и подписи вниз ===
    spacer = doc.add_paragraph("\n")
    spacer.paragraph_format.space_after = Pt(300)

    # === Итог ===
    total = len(selected_items)
    total_p = doc.add_paragraph(f"Итого: {total} человек")
    total_p.runs[0].bold = True
    total_p.runs[0].font.size = Pt(12)

    # === Подпись заведующей ===
    podpis = doc.add_paragraph()
    podpis.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    run_role = podpis.add_run("Заведующая отделением дневного пребывания ")
    run_role.font.size = Pt(12)

    run_line = podpis.add_run("__________________ ")
    run_line.font.size = Pt(12)

    run_name = podpis.add_run("Дурандина А.В.")
    run_name.font.size = Pt(12)

    # === Сохраняем файл на рабочий стол ===
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    safe_shift = shift_name.replace(" ", "_")
    safe_date = date_range.replace(" ", "_").replace(":", "-").replace(".", "-")
    file_name = f"{safe_shift}_{safe_date}.docx"
    file_path = os.path.join(desktop, file_name)

    try:
        doc.save(file_path)
        messagebox.showinfo("Готово", f"Список сохранён на рабочем столе:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    # === ДОБАВЬТЕ ЭТОТ БЛОК В КОНЕЦ ФУНКЦИИ ===
    # После экспорта сбрасываем галочки и обновляем счетчик
    for row_id in tree.get_children():
        values = list(tree.item(row_id, "values"))
        if values[0] == "X":
            values[0] = " "
            tree.item(row_id, values=values)
    
    if hasattr(root, 'update_word_count'):
        root.update_word_count()




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
                    UNIQUE(last_name, first_name, middle_name, dob)
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
                    UNIQUE(last_name, first_name, middle_name, dob)
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
            conn.commit()
            return

        # В иных случаях — добавляем недостающие колонки
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
            
# ================== Контекстное меню ==================
def show_context_menu(event):
    """Показать контекстное меню по правому клику"""
    item = tree.identify_row(event.y)
    if not item:
        return
    
    tree.selection_set(item)
    context_menu = tk.Menu(root, tearoff=0)
    
    values = tree.item(item, "values")
    client_id = values[1]
    last_name = values[2]
    first_name = values[3]
    client_name = f"{last_name} {first_name}"
    
    context_menu.add_command(
        label=f"Редактировать: {client_name}", 
        command=edit_client
    )
    context_menu.add_command(
        label=f"Удалить: {client_name}", 
        command=delete_selected
    )
    context_menu.add_separator()
    context_menu.add_command(
        label="Быстрый просмотр", 
        command=lambda: quick_view(client_id)
    )
    context_menu.add_command(
        label="Скопировать ФИО", 
        command=lambda: copy_to_clipboard(f"{last_name} {first_name} {values[4] or ''}".strip())
    )
    context_menu.add_command(
        label="Скопировать телефон", 
        command=lambda: copy_to_clipboard(values[6] or "")
    )
    context_menu.add_separator()
    context_menu.add_command(
        label="Добавить в список Word", 
        command=lambda: add_to_word_list(item)
    )
    
    try:
        context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        context_menu.grab_release()

def quick_view(client_id):
    """Быстрый просмотр информации о клиенте"""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT last_name, first_name, middle_name, dob, phone, contract_number, ippcu_start, ippcu_end, group_name FROM clients WHERE id=?",
            (client_id,)
        )
        client = cur.fetchone()
    
    if not client:
        messagebox.showerror("Ошибка", "Клиент не найден")
        return
    
    last, first, middle, dob, phone, contract, ippcu_start, ippcu_end, group = client
    
    info_text = f"""👤 {last} {first} {middle or ''}

📅 Дата рождения: {dob or 'не указана'}
📞 Телефон: {phone or 'не указан'}
📄 Договор: {contract or 'не указан'}
🏷️ Группа: {group or 'не указана'}

📋 ИППСУ:
   Начало: {ippcu_start or 'не указано'}
   Окончание: {ippcu_end or 'не указано'}"""
    
    if ippcu_end:
        try:
            end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
            today = datetime.today().date()
            days_left = (end_date - today).days
            
            if days_left < 0:
                info_text += f"\n\n⚠️ ИППСУ ПРОСРОЧЕН на {abs(days_left)} дн."
            elif days_left <= 30:
                info_text += f"\n\n⚠️ ИППСУ истекает через {days_left} дн."
            else:
                info_text += f"\n\n✅ ИППСУ активен ({days_left} дн. осталось)"
        except:
            pass
    
    messagebox.showinfo("Информация о клиенте", info_text)

def copy_to_clipboard(text):
    """Копировать текст в буфер обмена"""
    if text:
        root.clipboard_clear()
        root.clipboard_append(text)
        show_status_message(f"Скопировано: {text[:20]}..." if len(text) > 20 else f"Скопировано: {text}")

def add_to_word_list(item):
    """Добавить/убрать клиента из списка для Word"""
    values = list(tree.item(item, "values"))
    values[0] = "X" if values[0].strip() == "" else " "
    tree.item(item, values=values)
    
    action = "добавлен в" if values[0] == "X" else "удален из"
    show_status_message(f"Клиент {action} списка для Word")

def show_status_message(message, duration=3000):
    """Показать временное сообщение в статусной строке"""
    if hasattr(root, 'status_label'):
        root.status_label.config(text=message)
        root.after(duration, lambda: root.status_label.config(text="Готово"))

def create_status_bar():
    """Создать строку статуса"""
    status_frame = tk.Frame(root, relief=tk.SUNKEN, bd=1)
    status_frame.grid(row=10, column=0, columnspan=8, sticky="we")
    
    status_label = tk.Label(status_frame, text="Готово", anchor="w")
    status_label.pack(side=tk.LEFT, fill=tk.X, padx=5)
    
    word_count_label = tk.Label(status_frame, text="Выбрано для Word: 0", anchor="e")
    word_count_label.pack(side=tk.RIGHT, padx=5)
    
    root.status_label = status_label
    root.word_count_label = word_count_label
    
    def update_word_count():
        count = sum(1 for row_id in tree.get_children() 
                   if tree.item(row_id, "values")[0] == "X")
        word_count_label.config(text=f"Выбрано для Word: {count}")
    
    root.update_word_count = update_word_count
    root.after(100, update_word_count)

# ================== Автоподбор колонок ==================
def auto_resize_columns(tree, max_width=400):
    """Автоподбор ширины колонок с ограничением по максимальной ширине"""
    tree.update_idletasks()
    
    # Определяем приоритеты колонок (какие должны быть шире)
    column_priority = {
        "Фамилия": 2, "Имя": 2, "Отчество": 2, 
        "Дата рождения": 1, "Телефон": 1, "Номер договора": 1,
        "Дата начала ИППСУ": 1, "Дата окончания ИППСУ": 1, "Группа": 1,
        "✓": 0, "ID": 0  # Эти колонки делаем узкими
    }
    
    for col in tree["columns"]:
        # Ширина заголовка
        header_text = tree.heading(col)["text"]
        header_width = tk.font.Font().measure(header_text) + 30
        
        # Ширина по содержимому
        content_width = header_width
        for item in tree.get_children():
            cell_value = str(tree.set(item, col))
            cell_width = tk.font.Font().measure(cell_value) + 20
            if cell_width > content_width:
                content_width = cell_width
        
        # Учитываем приоритет колонки
        priority = column_priority.get(header_text, 1)
        if priority == 0:  # Узкие колонки
            final_width = min(content_width, 80)
        elif priority == 2:  # Широкие колонки (ФИО)
            final_width = min(content_width, max_width)
        else:  # Средние колонки
            final_width = min(content_width, 150)
        
        tree.column(col, width=final_width, minwidth=30)

def setup_tree_behavior(tree):
    """Настройка поведения таблицы"""
    
    # Двойной клик на разделитель колонок - автоподбор
    def on_header_click(event):
        # Определяем, был ли клик на разделителе колонок
        region = tree.identify("region", event.x, event.y)
        if region == "separator":
            column = tree.identify_column(event.x)
            col_id = column.replace("#", "")
            columns = tree["columns"]
            if col_id.isdigit() and int(col_id) <= len(columns):
                col_name = columns[int(col_id)-1]
                auto_resize_single_column(tree, col_name)
    
    tree.bind("<Double-1>", on_header_click)

def auto_resize_single_column(tree, col_name):
    """Автоподбор ширины для одной колонки"""
    tree.update_idletasks()
    
    # Ширина заголовка
    header_text = tree.heading(col_name)["text"]
    header_width = tk.font.Font().measure(header_text) + 30
    
    # Ширина по содержимому
    content_width = header_width
    for item in tree.get_children():
        cell_value = str(tree.set(item, col_name))
        cell_width = tk.font.Font().measure(cell_value) + 20
        if cell_width > content_width:
            content_width = cell_width
    
    final_width = min(content_width, 400)
    tree.column(col_name, width=final_width)

def setup_initial_columns(tree):
    """Начальная настройка колонок"""
    # Устанавливаем разумные начальные ширины
    tree.column("✓", width=30, minwidth=20, stretch=False)
    tree.column("ID", width=40, minwidth=30, stretch=False)
    tree.column("Фамилия", width=120, minwidth=80)
    tree.column("Имя", width=120, minwidth=80)
    tree.column("Отчество", width=120, minwidth=80)
    tree.column("Дата рождения", width=100, minwidth=80)
    tree.column("Телефон", width=120, minwidth=80)
    tree.column("Номер договора", width=120, minwidth=80)
    tree.column("Дата начала ИППСУ", width=120, minwidth=80)
    tree.column("Дата окончания ИППСУ", width=120, minwidth=80)
    tree.column("Группа", width=100, minwidth=80)


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
            " ",  # 👈 первая колонка для галочки всегда пустая
            cid, last or "", first or "", middle or "",
            dob or "", phone or "", contract or "",
            ippcu_start or "", ippcu_end or "", group or ""
        ), tags=(tag,))
    
    # Автоподбор ширины колонок после обновления данных
    root.after(100, lambda: auto_resize_columns(tree))




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
root.title("Отделение дневного пребывания — Полустационарное обслуживание")
root.iconbitmap("icon.ico")


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

# ================== Таблица ==================
tree = ttk.Treeview(
    root,
    columns=("✓", "ID", "Фамилия", "Имя", "Отчество", "Дата рождения", "Телефон",
             "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"),
    show="headings",
    height=20
)
tree.grid(row=1, column=0, columnspan=7, padx=5, pady=5, sticky="nsew")

for col in tree["columns"]:
    tree.heading(col, text=col)

# Цветовые теги
tree.tag_configure("expired", background="#ffcccc")
tree.tag_configure("soon", background="#fff2cc")
tree.tag_configure("active", background="#ccffcc")

# ==== ДОБАВЬТЕ ЭТИ 2 СТРОКИ ====
# КОНТЕКСТНОЕ МЕНЮ (правый клик)
tree.bind("<Button-3>", show_context_menu)

# ==== ИНИЦИАЛИЗАЦИЯ КОЛОНОК ====
setup_initial_columns(tree)
setup_tree_behavior(tree)

# ==== обработчик для чекбоксов ====
def toggle_check(event):
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return
    col = tree.identify_column(event.x)
    if col != "#1":  # первая колонка ("✓")
        return

    row_id = tree.identify_row(event.y)
    if not row_id:
        return

    values = list(tree.item(row_id, "values"))
    current = values[0]  # колонка "✓"
    values[0] = "X" if current.strip() == "" else " "
    tree.item(row_id, values=values)
    if hasattr(root, 'update_word_count'):
        root.update_word_count()

tree.bind("<Button-1>", toggle_check)

# Кнопки
tk.Button(root, text="Добавить", command=add_window).grid(row=2, column=0, padx=5, pady=5)
tk.Button(root, text="Редактировать", command=edit_client).grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Удалить", command=delete_selected).grid(row=2, column=2, padx=5, pady=5)
tk.Button(root, text="Быстрый просмотр", command=lambda: quick_view(tree.item(tree.selection()[0], "values")[1] if tree.selection() else None)).grid(row=2, column=3, padx=5, pady=5)
tk.Button(root, text="Импорт Google Sheets", command=import_from_gsheet).grid(row=2, column=4, padx=5, pady=5)
tk.Button(root, text="Экспорт в Word", command=export_selected_to_word).grid(row=2, column=5, padx=5, pady=5)
tk.Button(root, text="Автоподбор колонок", command=lambda: auto_resize_columns(tree)).grid(row=2, column=6, padx=5, pady=5)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

init_db()
root.after(200, refresh_tree)

create_status_bar()

root.mainloop()
