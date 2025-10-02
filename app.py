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
import updater
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import simpledialog


# ================== Пути ==================
APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
os.makedirs(APP_DIR, exist_ok=True)

DB_NAME = os.path.join(APP_DIR, "clients.db")
SHEET_ID = "1_DfTT8yzCjP0VH0PZu1Fz6FYMm1eRr7c0TmZU2DrH_w"

# ================== СОВРЕМЕННЫЙ СТИЛЬ ==================
class ModernStyle:
    COLORS = {
        'primary': '#2E86AB',
        'primary_dark': '#1A5A7A',
        'secondary': '#A23B72',
        'accent': '#F18F01',
        'success': '#4CAF50',
        'warning': '#FF9800',
        'error': '#F44336',
        'background': '#F8F9FA',
        'surface': '#FFFFFF',
        'text_primary': '#212529',
        'text_secondary': '#6C757D',
        'border': '#DEE2E6'
    }
    
    FONTS = {
        'h1': ('Segoe UI', 20, 'bold'),
        'h2': ('Segoe UI', 16, 'bold'),
        'h3': ('Segoe UI', 14, 'bold'),
        'body': ('Segoe UI', 11),
        'small': ('Segoe UI', 10),
        'button': ('Segoe UI', 11, 'bold')
    }

def setup_modern_style():
    """Настройка современного стиля"""
    style = ttk.Style()
    
    try:
        style.theme_use('vista')
    except:
        try:
            style.theme_use('clam')
        except:
            pass
    
    # Настраиваем стили
    style.configure('Modern.TFrame', background=ModernStyle.COLORS['background'])
    style.configure('Modern.TLabel', background=ModernStyle.COLORS['background'], 
                   foreground=ModernStyle.COLORS['text_primary'], font=ModernStyle.FONTS['body'])
    style.configure('Primary.TButton', background=ModernStyle.COLORS['primary'], 
                   foreground='white', font=ModernStyle.FONTS['button'], borderwidth=0)
    style.configure('Secondary.TButton', background=ModernStyle.COLORS['surface'], 
                   foreground=ModernStyle.COLORS['primary'], font=ModernStyle.FONTS['button'])
    
    style.map('Primary.TButton',
              background=[('active', ModernStyle.COLORS['primary_dark']),
                         ('pressed', ModernStyle.COLORS['primary_dark'])])
    
    style.map('Secondary.TButton',
              background=[('active', ModernStyle.COLORS['border']),
                         ('pressed', ModernStyle.COLORS['border'])])
    
    # Стиль для Treeview
    style.configure('Modern.Treeview', 
                   background=ModernStyle.COLORS['surface'],
                   fieldbackground=ModernStyle.COLORS['surface'],
                   foreground=ModernStyle.COLORS['text_primary'],
                   font=ModernStyle.FONTS['body'],
                   rowheight=25)
    
    style.configure('Modern.Treeview.Heading', 
                   background=ModernStyle.COLORS['primary'],
                   foreground='white',
                   font=ModernStyle.FONTS['button'],
                   relief='flat')
    
    style.map('Modern.Treeview', 
              background=[('selected', ModernStyle.COLORS['primary'])],
              foreground=[('selected', 'white')])

def create_modern_header(root):
    """Создание современного заголовка"""
    header_frame = tk.Frame(root, bg=ModernStyle.COLORS['primary'], height=80)
    header_frame.pack(fill='x', padx=0, pady=0)
    
    # Основной заголовок
    title_frame = tk.Frame(header_frame, bg=ModernStyle.COLORS['primary'])
    title_frame.pack(fill='x', padx=20, pady=12)
    
    title_label = tk.Label(title_frame, 
                          text="Отделение дневного пребывания",
                          bg=ModernStyle.COLORS['primary'],
                          fg='white',
                          font=ModernStyle.FONTS['h1'])
    title_label.pack(side='left')
    
    subtitle_label = tk.Label(title_frame,
                             text="Полустационарное обслуживание",
                             bg=ModernStyle.COLORS['primary'],
                             fg='white',
                             font=ModernStyle.FONTS['h3'])
    subtitle_label.pack(side='left', padx=(15, 0))
    
    return header_frame

def create_search_panel(root):
    """Создание панели поиска"""
    search_frame = tk.Frame(root, bg=ModernStyle.COLORS['background'], padx=20, pady=15)
    search_frame.pack(fill='x', padx=0, pady=0)
    
    # Поисковая строка
    search_container = tk.Frame(search_frame, bg=ModernStyle.COLORS['surface'], 
                               relief='solid', bd=1, padx=10, pady=8)
    search_container.pack(fill='x', padx=0, pady=0)
    
    tk.Label(search_container, text="🔍 Поиск клиентов:", 
             bg=ModernStyle.COLORS['surface'],
             fg=ModernStyle.COLORS['text_primary'],
             font=ModernStyle.FONTS['h3']).pack(side='left', padx=(0, 10))
    
    search_entry = tk.Entry(search_container, width=40, font=ModernStyle.FONTS['body'],
                           relief='flat', bg=ModernStyle.COLORS['background'], bd=0)
    search_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
    
    search_btn = ttk.Button(search_container, text="Найти", style='Primary.TButton',
                           command=lambda: do_search())
    search_btn.pack(side='left', padx=(0, 20))
    
    # Фильтры по датам
    filters_frame = tk.Frame(search_container, bg=ModernStyle.COLORS['surface'])
    filters_frame.pack(side='left')
    
    tk.Label(filters_frame, text="ИППСУ до:", 
             bg=ModernStyle.COLORS['surface'],
             fg=ModernStyle.COLORS['text_secondary'],
             font=ModernStyle.FONTS['small']).pack(side='left', padx=(0, 5))
    
    date_from_entry = DateEntry(filters_frame, width=10, date_pattern="dd.mm.yyyy",
                               font=ModernStyle.FONTS['small'], background=ModernStyle.COLORS['primary'],
                               foreground='white', borderwidth=0)
    date_from_entry.pack(side='left', padx=(0, 10))
    
    tk.Label(filters_frame, text="–", 
             bg=ModernStyle.COLORS['surface'],
             fg=ModernStyle.COLORS['text_secondary'],
             font=ModernStyle.FONTS['small']).pack(side='left', padx=(0, 10))
    
    date_to_entry = DateEntry(filters_frame, width=10, date_pattern="dd.mm.yyyy",
                             font=ModernStyle.FONTS['small'], background=ModernStyle.COLORS['primary'],
                             foreground='white', borderwidth=0)
    date_to_entry.pack(side='left', padx=(0, 10))
    
    filter_btn = ttk.Button(filters_frame, text="Применить", style='Secondary.TButton',
                           command=lambda: do_search())
    filter_btn.pack(side='left')
    
    return search_entry, date_from_entry, date_to_entry, search_frame

def create_toolbar(root):
    """Создание панели инструментов"""
    toolbar_frame = tk.Frame(root, bg=ModernStyle.COLORS['surface'], padx=20, pady=10)
    toolbar_frame.pack(fill='x', padx=0, pady=0)
    
    buttons = [
        ("➕ Добавить клиента", add_window, 'Primary.TButton', "Ctrl+N"),
        ("✏️ Редактировать", edit_client, 'Secondary.TButton', "Ctrl+E"),
        ("🗑️ Удалить", delete_selected, 'Secondary.TButton', "Delete"),
        ("👁️ Просмотр", lambda: quick_view_wrapper(), 'Secondary.TButton', "Ctrl+Q"),
        ("📥 Импорт", import_from_gsheet, 'Secondary.TButton', "Ctrl+I"),
        ("📄 Экспорт в Word", export_selected_to_word, 'Secondary.TButton', "Ctrl+W"),
        ("📊 Статистика", show_statistics, 'Secondary.TButton', ""),
        ("🔔 Уведомления", show_notifications, 'Secondary.TButton', "F2"),
        ("⚙️ Настройки", settings_window, 'Secondary.TButton', "")  # Новая кнопка
    ]
    
    for text, command, style_name, shortcut in buttons:
        btn = ttk.Button(toolbar_frame, text=text, command=command, style=style_name)
        btn.pack(side='left', padx=(0, 8))
        
        # Добавляем подсказку с горячей клавишей
        if shortcut:
            tooltip_text = f"{text} ({shortcut})"
            create_tooltip(btn, tooltip_text)

    # Кнопка справки
    help_btn = ttk.Button(toolbar_frame, text="❓ Справка", 
                         command=show_help, style='Secondary.TButton')
    help_btn.pack(side='right')
    create_tooltip(help_btn, "Справка по горячим клавишам (F1)")
    
    return toolbar_frame

def create_tooltip(widget, text):
    """Создание всплывающей подсказки"""
    def on_enter(event):
        tooltip = tk.Toplevel()
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        
        label = tk.Label(tooltip, text=text, background="#ffffe0", 
                        relief='solid', borderwidth=1, font=ModernStyle.FONTS['small'])
        label.pack()
        
        widget.tooltip = tooltip
    
    def on_leave(event):
        if hasattr(widget, 'tooltip'):
            widget.tooltip.destroy()
    
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)

def create_modern_table(root):
    """Создание современной таблицы"""
    table_container = tk.Frame(root, bg=ModernStyle.COLORS['background'], padx=20, pady=15)
    table_container.pack(fill='both', expand=True, padx=0, pady=0)
    
    # Заголовок таблицы
    table_header = tk.Frame(table_container, bg=ModernStyle.COLORS['background'])
    table_header.pack(fill='x', pady=(0, 10))
    
    tk.Label(table_header, text="Список клиентов", 
             bg=ModernStyle.COLORS['background'],
             fg=ModernStyle.COLORS['text_primary'],
             font=ModernStyle.FONTS['h2']).pack(side='left')
    
    # Контейнер для таблицы с тенью
    table_wrapper = tk.Frame(table_container, bg=ModernStyle.COLORS['border'], 
                            relief='solid', bd=1, padx=1, pady=1)
    table_wrapper.pack(fill='both', expand=True)
    
    # Создаем таблицу с прокруткой
    tree_scroll = ttk.Scrollbar(table_wrapper)
    tree_scroll.pack(side='right', fill='y')
    
    tree = ttk.Treeview(
        table_wrapper,
        columns=("✓", "ID", "Фамилия", "Имя", "Отчество", "Дата рождения", "Телефон",
                 "Номер договора", "Дата начала ИППСУ", "Дата окончания ИППСУ", "Группа"),
        show="headings",
        height=15,
        style='Modern.Treeview',
        yscrollcommand=tree_scroll.set
    )
    tree.pack(side='left', fill='both', expand=True)
    tree_scroll.config(command=tree.yview)
    
    # Настраиваем колонки
    for col in tree["columns"]:
        tree.heading(col, text=col)
    
    return tree, table_container

def create_status_bar(root):
    """Создание строки статуса"""
    status_frame = tk.Frame(root, bg=ModernStyle.COLORS['primary'], height=30)
    status_frame.pack(fill='x', side='bottom', padx=0, pady=0)
    status_frame.pack_propagate(False)
    
    status_label = tk.Label(status_frame, text="Готово", 
                           bg=ModernStyle.COLORS['primary'],
                           fg='white', font=ModernStyle.FONTS['small'])
    status_label.pack(side='left', padx=10, pady=5)
    
    word_count_label = tk.Label(status_frame, text="Выбрано для Word: 0", 
                               bg=ModernStyle.COLORS['primary'],
                               fg='white', font=ModernStyle.FONTS['small'])
    word_count_label.pack(side='right', padx=10, pady=5)
    
    root.status_label = status_label
    root.word_count_label = word_count_label
    
    def update_word_count():
        count = sum(1 for row_id in tree.get_children() 
                   if tree.item(row_id, "values")[0] == "X")
        word_count_label.config(text=f"Выбрано для Word: {count}")
    
    root.update_word_count = update_word_count
    return status_frame

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
        label=f"Редактировать: {client_name} (Ctrl+E)", 
        command=edit_client
    )
    context_menu.add_command(
        label=f"Удалить: {client_name} (Delete)", 
        command=delete_selected
    )
    context_menu.add_separator()
    context_menu.add_command(
        label="Быстрый просмотр (Ctrl+Q)", 
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
    context_menu.add_separator()
    context_menu.add_command(
        label="Справка по горячим клавишам (F1)", 
        command=show_help
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

# ================== Автоподбор колонок ==================
def auto_resize_columns(tree, max_width=400):
    """Автоподбор ширины колонок с ограничением по максимальной ширине"""
    tree.update_idletasks()
    
    column_priority = {
        "Фамилия": 2, "Имя": 2, "Отчество": 2, 
        "Дата рождения": 1, "Телефон": 1, "Номер договора": 1,
        "Дата начала ИППСУ": 1, "Дата окончания ИППСУ": 1, "Группа": 1,
        "✓": 0, "ID": 0
    }
    
    for col in tree["columns"]:
        header_text = tree.heading(col)["text"]
        header_width = tk.font.Font().measure(header_text) + 30
        
        content_width = header_width
        for item in tree.get_children():
            cell_value = str(tree.set(item, col))
            cell_width = tk.font.Font().measure(cell_value) + 20
            if cell_width > content_width:
                content_width = cell_width
        
        priority = column_priority.get(header_text, 1)
        if priority == 0:
            final_width = min(content_width, 80)
        elif priority == 2:
            final_width = min(content_width, max_width)
        else:
            final_width = min(content_width, 150)
        
        tree.column(col, width=final_width, minwidth=30)

def setup_tree_behavior(tree):
    """Настройка поведения таблицы"""
    def on_header_click(event):
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
    
    header_text = tree.heading(col_name)["text"]
    header_width = tk.font.Font().measure(header_text) + 30
    
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

def show_statistics():
    """Показать статистику по клиентам"""
    clients = get_all_clients(limit=10000)
    total = len(clients)
    
    today = datetime.today().date()
    active = 0
    expired = 0
    soon = 0
    groups = {}
    
    for client in clients:
        ippcu_end = client[8]
        group = client[9] or "Без группы"
        
        if group not in groups:
            groups[group] = 0
        groups[group] += 1
        
        if ippcu_end:
            try:
                end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
                if end_date < today:
                    expired += 1
                elif end_date <= today + timedelta(days=30):
                    soon += 1
                else:
                    active += 1
            except:
                pass
    
    stats_text = f"""📊 СТАТИСТИКА

Всего клиентов: {total}
├─ Активные ИППСУ: {active}
├─ Истекают в течение 30 дней: {soon}
└─ Просроченные ИППСУ: {expired}

📂 РАСПРЕДЕЛЕНИЕ ПО ГРУППАМ:"""
    
    for group, count in sorted(groups.items()):
        percentage = (count / total) * 100 if total > 0 else 0
        stats_text += f"\n├─ {group}: {count} чел. ({percentage:.1f}%)"
    
    messagebox.showinfo("📊 Статистика", stats_text)

def check_expiring_ippcu():
    """Проверка истекающих ИППСУ при запуске"""
    clients = get_all_clients(limit=10000)
    today = datetime.today().date()
    
    expiring = []
    expired = []
    
    for client in clients:
        ippcu_end = client[8]
        if ippcu_end:
            try:
                end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
                days_left = (end_date - today).days
                
                if 0 <= days_left <= 7:
                    expiring.append((client, days_left))
                elif days_left < 0:
                    expired.append((client, abs(days_left)))
            except:
                pass
    
    messages = []
    
    if expired:
        messages.append(f"❌ ПРОСРОЧЕНЫ {len(expired)} ИППСУ!")
        for client, days in expired[:3]:
            messages.append(f"   {client[1]} {client[2]} - просрочено {days} дн. назад")
    
    if expiring:
        messages.append(f"⚠️ ИСТЕКАЮТ {len(expiring)} ИППСУ в течение недели!")
        for client, days in expiring[:3]:
            messages.append(f"   {client[1]} {client[2]} - осталось {days} дн.")
    
    if messages:
        messagebox.showwarning("Внимание!", "\n".join(messages))

def export_selected_to_word():
    selected_items = []
    for row_id in tree.get_children():
        values = tree.item(row_id, "values")
        if values and values[0] == "X":
            selected_items.append(values)

    if not selected_items:
        messagebox.showerror("Ошибка", "Отметьте галочками хотя бы одного клиента")
        return

    shift_name = simpledialog.askstring("Смена", "Введите название смены (например: 11 смена)")
    if not shift_name:
        return

    date_range = simpledialog.askstring("Даты", "Введите период (например: с 01.10.2024 по 15.10.2024)")
    if not date_range:
        return

    doc = Document()

    heading = doc.add_paragraph(f"{shift_name} {date_range}")
    heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = heading.runs[0]
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph("")

    for i, values in enumerate(selected_items, start=1):
        last = values[2]
        first = values[3]
        middle = values[4]
        dob = values[5]

        fio = " ".join(v for v in [last, first, middle] if v)
        p = doc.add_paragraph(f"{i}. {fio} – {dob} г.р.")
        p.runs[0].font.size = Pt(12)

    spacer = doc.add_paragraph("\n")
    spacer.paragraph_format.space_after = Pt(300)

    total = len(selected_items)
    total_p = doc.add_paragraph(f"Итого: {total} человек")
    total_p.runs[0].bold = True
    total_p.runs[0].font.size = Pt(12)

    podpis = doc.add_paragraph()
    podpis.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    run_role = podpis.add_run("Заведующая отделением дневного пребывания ")
    run_role.font.size = Pt(12)

    run_line = podpis.add_run("__________________ ")
    run_line.font.size = Pt(12)

    run_name = podpis.add_run("Дурандина А.В.")
    run_name.font.size = Pt(12)

    # Используем путь из настроек или рабочий стол по умолчанию
    export_path = settings_manager.get('default_export_path', os.path.join(os.path.expanduser("~"), "Desktop"))
    
    safe_shift = shift_name.replace(" ", "_")
    safe_date = date_range.replace(" ", "_").replace(":", "-").replace(".", "-")
    file_name = f"{safe_shift}_{safe_date}.docx"
    file_path = os.path.join(export_path, file_name)

    try:
        doc.save(file_path)
        messagebox.showinfo("Готово", f"Список сохранён:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    for row_id in tree.get_children():
        values = list(tree.item(row_id, "values"))
        if values[0] == "X":
            values[0] = " "
            tree.item(row_id, values=values)
    
    if hasattr(root, 'update_word_count'):
        root.update_word_count()

# ================== СИСТЕМА УВЕДОМЛЕНИЙ ==================
class NotificationSystem:
    def __init__(self):
        self.notifications = []
        self.setup_daily_checks()
    
    def setup_daily_checks(self):
        """Настройка ежедневных проверок"""
        self.check_birthdays()
        self.check_ippcu_expiry()
        self.check_empty_contracts()
    
    def check_birthdays(self):
        """Проверка ближайших дней рождений"""
        today = datetime.today().date()
        next_week = today + timedelta(days=7)
        
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT last_name, first_name, middle_name, dob 
                FROM clients 
                WHERE substr(dob, 6, 5) BETWEEN ? AND ?
            """, (today.strftime("%m-%d"), next_week.strftime("%m-%d")))
            
            birthdays = cur.fetchall()
        
        for last, first, middle, dob in birthdays:
            try:
                bday = datetime.strptime(dob, "%Y-%m-%d").date()
                bday_this_year = bday.replace(year=today.year)
                days_until = (bday_this_year - today).days
                if days_until >= 0:
                    self.add_notification(
                        "birthday", 
                        f"День рождения у {last} {first} {middle or ''} через {days_until} дн. ({bday.strftime('%d.%m.%Y')})",
                        "info" if days_until > 3 else "warning"
                    )
            except:
                continue
    
    def check_ippcu_expiry(self):
        """Проверка истекающих ИППСУ"""
        today = datetime.today().date()
        
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT last_name, first_name, ippcu_end 
                FROM clients 
                WHERE ippcu_end IS NOT NULL AND ippcu_end != ''
            """)
            
            clients = cur.fetchall()
        
        for last, first, ippcu_end in clients:
            try:
                end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
                days_left = (end_date - today).days
                
                if 0 < days_left <= 7:
                    self.add_notification(
                        "ippcu_warning",
                        f"ИППСУ {last} {first} истекает через {days_left} дн.",
                        "warning"
                    )
                elif days_left == 0:
                    self.add_notification(
                        "ippcu_urgent",
                        f"СРОЧНО: ИППСУ {last} {first} истекает сегодня!",
                        "error"
                    )
                elif days_left < 0:
                    self.add_notification(
                        "ippcu_expired",
                        f"ПРОСРОЧЕНО: ИППСУ {last} {first} ({abs(days_left)} дн. назад)",
                        "error"
                    )
            except:
                continue
    
    def check_empty_contracts(self):
        """Проверка клиентов без договоров"""
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT last_name, first_name 
                FROM clients 
                WHERE contract_number IS NULL OR contract_number = '' OR contract_number = 'не указан'
            """)
            
            empty_contracts = cur.fetchall()
        
        if empty_contracts:
            self.add_notification(
                "empty_contracts",
                f"Найдено {len(empty_contracts)} клиентов без номера договора",
                "warning"
            )
    
    def add_notification(self, category, message, level="info"):
        """Добавление уведомления"""
        self.notifications.append({
            "timestamp": datetime.now(),
            "category": category,
            "message": message,
            "level": level,
            "read": False
        })
    
    def show_daily_reminders(self):
        """Показать ежедневные напоминания"""
        if not self.notifications:
            return
        
        unread = [n for n in self.notifications if not n['read']]
        if unread:
            self.show_notification_window()
    
    def show_notification_window(self):
        """Окно уведомлений"""
        if not hasattr(self, 'notification_window') or not self.notification_window.winfo_exists():
            self.create_notification_window()
        
        self.update_notification_list()
    
    def create_notification_window(self):
        """Создание окна уведомлений"""
        self.notification_window = tk.Toplevel(root)
        self.notification_window.title("Уведомления")
        self.notification_window.geometry("500x400")
        self.notification_window.configure(bg=ModernStyle.COLORS['background'])
        
        # Заголовок
        header = tk.Frame(self.notification_window, bg=ModernStyle.COLORS['primary'], height=50)
        header.pack(fill='x', padx=0, pady=0)
        
        tk.Label(header, text="🔔 Уведомления", 
                bg=ModernStyle.COLORS['primary'],
                fg='white',
                font=ModernStyle.FONTS['h2']).pack(pady=10)
        
        # Список уведомлений
        notification_frame = tk.Frame(self.notification_window, bg=ModernStyle.COLORS['background'])
        notification_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.notification_list = tk.Listbox(notification_frame, 
                                          font=ModernStyle.FONTS['body'],
                                          bg=ModernStyle.COLORS['surface'],
                                          relief='flat',
                                          selectmode='single')
        self.notification_list.pack(fill='both', expand=True)
        
        # Кнопки
        button_frame = tk.Frame(self.notification_window, bg=ModernStyle.COLORS['background'])
        button_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(button_frame, text="Пометить как прочитанные", 
                  style='Primary.TButton',
                  command=self.mark_all_read).pack(side='left', padx=(0, 10))
        
        ttk.Button(button_frame, text="Очистить все", 
                  style='Secondary.TButton',
                  command=self.clear_all).pack(side='left')
        
        ttk.Button(button_frame, text="Закрыть", 
                  style='Secondary.TButton',
                  command=self.notification_window.destroy).pack(side='right')
        
        # Двойной клик для пометки как прочитанного
        self.notification_list.bind('<Double-1>', lambda e: self.mark_selected_read())
    
    def update_notification_list(self):
        """Обновление списка уведомлений"""
        if hasattr(self, 'notification_list'):
            self.notification_list.delete(0, tk.END)
            
            for notification in sorted(self.notifications, 
                                     key=lambda x: x['timestamp'], reverse=True):
                level_icon = {
                    'info': 'ℹ️',
                    'warning': '⚠️', 
                    'error': '❌'
                }.get(notification['level'], '📌')
                
                status_icon = '✅' if notification['read'] else '🔔'
                time_str = notification['timestamp'].strftime("%H:%M")
                
                display_text = f"{status_icon} {level_icon} [{time_str}] {notification['message']}"
                self.notification_list.insert(tk.END, display_text)
    
    def mark_selected_read(self):
        """Пометить выбранное уведомление как прочитанное"""
        selection = self.notification_list.curselection()
        if selection:
            index = selection[0]
            # Находим соответствующее уведомление (учитываем обратный порядок)
            actual_index = len(self.notifications) - 1 - index
            if 0 <= actual_index < len(self.notifications):
                self.notifications[actual_index]['read'] = True
            self.update_notification_list()
    
    def mark_all_read(self):
        """Пометить все как прочитанные"""
        for notification in self.notifications:
            notification['read'] = True
        self.update_notification_list()
    
    def clear_all(self):
        """Очистить все уведомления"""
        self.notifications = []
        self.update_notification_list()

# Глобальный экземпляр системы уведомлений
notification_system = NotificationSystem()

    def show_notifications():
    """Показать уведомления (вызывается из меню)"""
    notification_system.show_notification_window()

# ================== ГОРЯЧИЕ КЛАВИШИ ==================
    def setup_keyboard_shortcuts():
    """Настройка горячих клавиш"""
    
    # Основные команды
    root.bind('<Control-n>', lambda e: add_window())
    root.bind('<Control-f>', lambda e: root.search_entry.focus())
    root.bind('<Control-s>', lambda e: do_search())
    root.bind('<Delete>', lambda e: delete_selected())
    root.bind('<F5>', lambda e: refresh_tree())
    root.bind('<F1>', lambda e: show_help())
    
    # Навигация
    root.bind('<Control-q>', lambda e: quick_view_wrapper())
    root.bind('<Control-e>', lambda e: edit_client())
    root.bind('<Control-i>', lambda e: import_from_gsheet())
    root.bind('<Control-w>', lambda e: export_selected_to_word())
    
    # Уведомления
    root.bind('<F2>', lambda e: show_notifications())
    
    # Сообщение в статусной строке о горячих клавишах
    show_status_message("Горячие клавиши активированы. Нажмите F1 для справки.")

def quick_view_wrapper():
    """Обертка для быстрого просмотра с горячей клавишей"""
    selected = tree.selection()
    if selected:
        client_id = tree.item(selected[0], "values")[1]
        quick_view(client_id)
    else:
        messagebox.showinfo("Подсказка", "Выберите клиента для быстрого просмотра")

def show_help():
    """Окно справки по горячим клавишам"""
    help_text = """
📋 ГОРЯЧИЕ КЛАВИШИ:

Основные команды:
Ctrl+N - Добавить клиента
Ctrl+F - Перейти в поиск
Ctrl+S - Выполнить поиск
Delete - Удалить выбранного
F5 - Обновить список

Навигация:
Ctrl+Q - Быстрый просмотр
Ctrl+E - Редактировать
Ctrl+I - Импорт из Google Sheets  
Ctrl+W - Экспорт в Word

Уведомления:
F2 - Показать уведомления

Справка:
F1 - Показать эту справку

Управление таблицей:
←/→ - Изменить ширину колонки
Double Click - Автоподбор колонки
Правый клик - Контекстное меню
"""
    
    help_window = tk.Toplevel(root)
    help_window.title("Справка по горячим клавишам")
    help_window.geometry("500x500")
    help_window.configure(bg=ModernStyle.COLORS['background'])
    
    # Заголовок
    header = tk.Frame(help_window, bg=ModernStyle.COLORS['primary'], height=50)
    header.pack(fill='x', padx=0, pady=0)
    
    tk.Label(header, text="⌨️ Горячие клавиши", 
            bg=ModernStyle.COLORS['primary'],
            fg='white',
            font=ModernStyle.FONTS['h2']).pack(pady=10)
    
    # Текст справки
    text_frame = tk.Frame(help_window, bg=ModernStyle.COLORS['background'])
    text_frame.pack(fill='both', expand=True, padx=20, pady=20)
    
    help_text_widget = tk.Text(text_frame, 
                              font=ModernStyle.FONTS['body'],
                              bg=ModernStyle.COLORS['surface'],
                              fg=ModernStyle.COLORS['text_primary'],
                              wrap='word',
                              padx=10,
                              pady=10)
    help_text_widget.pack(fill='both', expand=True)
    
    help_text_widget.insert('1.0', help_text)
    help_text_widget.config(state='disabled')  # Только для чтения
    
    # Кнопка закрытия
    button_frame = tk.Frame(help_window, bg=ModernStyle.COLORS['background'])
    button_frame.pack(fill='x', padx=20, pady=10)
    
    ttk.Button(button_frame, text="Закрыть", 
              style='Primary.TButton',
              command=help_window.destroy).pack(side='right')

# ================== База данных ==================
def init_db():
    """Создаёт новую схему или мигрирует старую (если есть колонка fio)."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(clients)")
        cols = [r[1] for r in cur.fetchall()]

        if not cols:
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

        if "fio" in cols and "last_name" not in cols:
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
                    cur.execute(
                        "INSERT OR IGNORE INTO clients_new (last_name, first_name, middle_name, dob) VALUES (?, ?, ?, ?)",
                        (last or "", first or "", middle or "", dob_val)
                    )
            cur.execute("DROP TABLE clients")
            cur.execute("ALTER TABLE clients_new RENAME TO clients")
            conn.commit()
            return

        if "last_name" in cols and "dob" in cols:
            conn.commit()
            return

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
        middle_name = middle_name or ""
        dob_val = dob or ""

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
                continue
        refresh_tree()
        messagebox.showinfo("Успех", f"Импорт из Google Sheets завершён! Добавлено: {added}")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")

# ================== UI ==================
def refresh_tree(results=None):
    # очищаем таблицу
    for row in tree.get_children():
        tree.delete(row)

    # если нет результатов — берём все записи
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
                    tag = "expired"   # срок истёк
                elif end_date <= soon:
                    tag = "soon"      # истекает скоро
                else:
                    tag = "active"    # ещё действует
        except Exception:
            tag = ""

        tree.insert(
            "",
            "end",
            values=(" ", cid, last, first, middle, dob, phone, contract, ippcu_start, ippcu_end, group),
            tags=(tag,)
        )

    # оформление цветом
    tree.tag_configure("expired", background="#F8D7DA")   # красный (просрочен)
    tree.tag_configure("soon", background="#FFF3CD")      # жёлтый (скоро истечёт)
    tree.tag_configure("active", background="#D4EDDA")    # зелёный (активный)

    
    root.after(100, lambda: auto_resize_columns(tree))

def add_window():
    win = tk.Toplevel()
    win.title("Добавить обслуживаемого")
    win.configure(bg=ModernStyle.COLORS['background'])

    fields = [
        ("Фамилия", 0), ("Имя", 1), ("Отчество", 2),
        ("Дата рождения", 3), ("Телефон", 4), ("Номер договора", 5),
        ("Дата начала ИППСУ", 6), ("Дата окончания ИППСУ", 7), ("Группа", 8)
    ]

    entries = {}
    for field, row in fields:
        tk.Label(win, text=field, bg=ModernStyle.COLORS['background'],
                fg=ModernStyle.COLORS['text_primary'], font=ModernStyle.FONTS['body']).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        if "Дата" in field:
            entry = DateEntry(win, width=27, date_pattern="dd.mm.yyyy",
                            font=ModernStyle.FONTS['body'])
        else:
            entry = tk.Entry(win, width=30, font=ModernStyle.FONTS['body'])
        
        entry.grid(row=row, column=1, padx=10, pady=5)
        entries[field] = entry

    def save_client():
        last = entries["Фамилия"].get().strip()
        first = entries["Имя"].get().strip()
        middle = entries["Отчество"].get().strip()
        dob = entries["Дата рождения"].get_date().strftime("%Y-%m-%d")
        phone = entries["Телефон"].get().strip()
        contract_number = entries["Номер договора"].get().strip()
        ippcu_start = entries["Дата начала ИППСУ"].get_date().strftime("%Y-%m-%d")
        ippcu_end = entries["Дата окончания ИППСУ"].get_date().strftime("%Y-%m-%d")
        group = entries["Группа"].get().strip()

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

    save_btn = ttk.Button(win, text="Сохранить", style='Primary.TButton', command=save_client)
    save_btn.grid(row=9, column=0, columnspan=2, pady=10)

def edit_client():
    """Окно редактирования клиента"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Ошибка", "Выберите клиента для редактирования")
        return

    values = tree.item(selected[0], "values")
    cid = values[1]

    last, first, middle = values[2], values[3], values[4]
    dob, phone, contract = values[5], values[6], values[7]
    ippcu_start, ippcu_end, group = values[8], values[9], values[10]

    win = tk.Toplevel(root)
    win.title("Редактировать клиента")
    win.configure(bg=ModernStyle.COLORS['background'])

    fields = [
        ("Фамилия", last, 0), ("Имя", first, 1), ("Отчество", middle, 2),
        ("Дата рождения", dob, 3), ("Телефон", phone, 4), ("Номер договора", contract, 5),
        ("Дата начала ИППСУ", ippcu_start, 6), ("Дата окончания ИППСУ", ippcu_end, 7), ("Группа", group, 8)
    ]

    entries = {}
    for field, value, row in fields:
        tk.Label(win, text=field, bg=ModernStyle.COLORS['background'],
                fg=ModernStyle.COLORS['text_primary'], font=ModernStyle.FONTS['body']).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        if "Дата" in field:
            entry = tk.Entry(win, width=30, font=ModernStyle.FONTS['body'])
            entry.insert(0, value)
        else:
            entry = tk.Entry(win, width=30, font=ModernStyle.FONTS['body'])
            entry.insert(0, value)
        
        entry.grid(row=row, column=1, padx=10, pady=5)
        entries[field] = entry

    def save_changes():
        update_client(cid,
                      entries["Фамилия"].get(), entries["Имя"].get(), entries["Отчество"].get(),
                      entries["Дата рождения"].get(), entries["Телефон"].get(), entries["Номер договора"].get(),
                      entries["Дата начала ИППСУ"].get(), entries["Дата окончания ИППСУ"].get(), entries["Группа"].get())
        refresh_tree()
        win.destroy()

    save_btn = ttk.Button(win, text="Сохранить", style='Primary.TButton', command=save_changes)
    save_btn.grid(row=9, column=0, columnspan=2, pady=10)

def delete_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showerror("Ошибка", "Выберите клиента для удаления")
        return
    item = tree.item(selected[0])
    cid = item["values"][1]
    if messagebox.askyesno("Удалить", "Точно удалить выбранного клиента?"):
        delete_client(cid)
        refresh_tree()

def do_search():
    query = root.search_entry.get().strip()
    date_from = root.date_from_entry.get_date().strftime("%Y-%m-%d") if root.date_from_entry.get() else None
    date_to = root.date_to_entry.get_date().strftime("%Y-%m-%d") if root.date_to_entry.get() else None

    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        q = (query or "").strip().lower()
        like = f"%{q}%"

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

def toggle_check(event):
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return
    col = tree.identify_column(event.x)
    if col != "#1":
        return

    row_id = tree.identify_row(event.y)
    if not row_id:
        return

    values = list(tree.item(row_id, "values"))
    current = values[0]
    values[0] = "X" if current.strip() == "" else " "
    tree.item(row_id, values=values)
    if hasattr(root, 'update_word_count'):
        root.update_word_count()

# ================== MAIN ==================
def main():
    global root, tree
    
    root = tk.Tk()
    root.title("Отделение дневного пребывания - Полустационарное обслуживание")
    root.geometry("1400x900")
    
    # Настройка современного стиля
    setup_modern_style()
    root.configure(bg=ModernStyle.COLORS['background'])
    
    # Создание интерфейса
    header = create_modern_header(root)
    search_entry, date_from_entry, date_to_entry, search_frame = create_search_panel(root)
    toolbar = create_toolbar(root)
    tree, table_container = create_modern_table(root)
    status_bar = create_status_bar(root)
    
    # Сохраняем ссылки на элементы
    root.search_entry = search_entry
    root.date_from_entry = date_from_entry
    root.date_to_entry = date_to_entry
    
    # Настройка таблицы
    setup_initial_columns(tree)
    setup_tree_behavior(tree)
    
    # Настройка горячих клавиш
    setup_keyboard_shortcuts()
    
    # Привязка событий
    tree.bind("<Button-3>", show_context_menu)
    tree.bind("<Button-1>", toggle_check)
    
    # Инициализация
    init_db()
    root.after(200, refresh_tree)
    root.after(1000, check_expiring_ippcu)
    
    # Показать уведомления при запуске (через 2 секунды)
    root.after(2000, notification_system.show_daily_reminders)

    # при старте проверяем обновления
    root.after(100, updater.auto_update)

    root.mainloop()

if __name__ == "__main__":
    main()
