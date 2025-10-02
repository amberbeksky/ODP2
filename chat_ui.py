import tkinter as tk
from tkinter import ttk
from datetime import datetime
from chat_manager import ChatManager

class ChatUI:
    def __init__(self, parent, chat_manager, style_colors, style_fonts):
        self.parent = parent
        self.chat_manager = chat_manager
        self.colors = style_colors
        self.fonts = style_fonts
        self.setup_ui()
        
    def setup_ui(self):
        """Создание интерфейса чата"""
        self.chat_frame = tk.Frame(self.parent, bg=self.colors['background'])
        
        # Заголовок чата
        self.create_chat_header()
        
        # Основная область чата
        self.create_chat_main_area()
        
        # Панель ввода
        self.create_input_panel()
        
        # Панель быстрых действий
        self.create_quick_actions()
        
    def create_chat_header(self):
        """Создание заголовка чата"""
        chat_header = tk.Frame(self.chat_frame, bg=self.colors['primary'], height=40)
        chat_header.pack(fill='x', padx=0, pady=0)
        
        self.chat_title = tk.Label(chat_header, text="💬 Чат сотрудников", 
                                  bg=self.colors['primary'],
                                  fg='white',
                                  font=self.fonts['h3'])
        self.chat_title.pack(side='left', padx=10, pady=8)
        
        # Счетчик непрочитанных
        self.unread_label = tk.Label(chat_header, text="", 
                                    bg=self.colors['primary'],
                                    fg='yellow',
                                    font=self.fonts['small'])
        self.unread_label.pack(side='right', padx=10, pady=8)
        
        # Кнопка обновления
        refresh_btn = ttk.Button(chat_header, text="🔄", 
                                style='Secondary.TButton',
                                command=self.refresh_chat,
                                width=3)
        refresh_btn.pack(side='right', padx=5)
    
    def create_chat_main_area(self):
        """Создание основной области сообщений"""
        chat_main = tk.Frame(self.chat_frame, bg=self.colors['background'])
        chat_main.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Список сообщений с прокруткой
        message_frame = tk.Frame(chat_main, bg=self.colors['surface'])
        message_frame.pack(fill='both', expand=True, pady=(0, 5))
        
        # Прокрутка
        message_scroll = ttk.Scrollbar(message_frame)
        message_scroll.pack(side='right', fill='y')
        
        # Текстовое поле для сообщений
        self.messages_text = tk.Text(message_frame, 
                                   height=15,
                                   wrap='word',
                                   font=self.fonts['body'],
                                   bg=self.colors['surface'],
                                   fg=self.colors['text_primary'],
                                   yscrollcommand=message_scroll.set,
                                   state='disabled')
        self.messages_text.pack(side='left', fill='both', expand=True)
        message_scroll.config(command=self.messages_text.yview)
        
        # Настройка цветов для разных типов сообщений
        self.setup_text_tags()
    
    def setup_text_tags(self):
        """Настройка тегов для форматирования текста"""
        # Обычные сообщения
        self.messages_text.tag_configure("header_text", 
                                       foreground=self.colors['primary'],
                                       font=(self.fonts['body'][0], self.fonts['body'][1], 'bold'))
        self.messages_text.tag_configure("message_text", 
                                       foreground=self.colors['text_primary'])
        
        # Системные сообщения
        self.messages_text.tag_configure("header_system", 
                                       foreground=self.colors['secondary'],
                                       font=(self.fonts['body'][0], self.fonts['body'][1], 'bold'))
        self.messages_text.tag_configure("message_system", 
                                       foreground=self.colors['text_secondary'])
        
        # Предупреждения
        self.messages_text.tag_configure("header_alert", 
                                       foreground=self.colors['error'],
                                       font=(self.fonts['body'][0], self.fonts['body'][1], 'bold'))
        self.messages_text.tag_configure("message_alert", 
                                       foreground=self.colors['error'])
    
    def create_input_panel(self):
        """Создание панели ввода сообщения"""
        input_frame = tk.Frame(self.chat_frame, bg=self.colors['background'])
        input_frame.pack(fill='x', pady=5)
        
        self.input_entry = tk.Entry(input_frame, 
                                  font=self.fonts['body'],
                                  bg=self.colors['surface'])
        self.input_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.input_entry.bind('<Return>', lambda e: self.send_message())
        
        self.send_btn = ttk.Button(input_frame, text="Отправить", 
                                 style='Primary.TButton',
                                 command=self.send_message)
        self.send_btn.pack(side='right')
    
    def create_quick_actions(self):
        """Создание панели быстрых действий"""
        quick_actions = tk.Frame(self.chat_frame, bg=self.colors['background'])
        quick_actions.pack(fill='x', pady=5)
        
        quick_buttons = [
            ("📋 Справка", "Нужна помощь с оформлением справки"),
            ("❓ Вопрос", "Есть вопрос по клиенту"),
            ("📅 Встреча", "Нужно обсудить план работы"),
        ]
        
        for text, template in quick_buttons:
            btn = ttk.Button(quick_actions, text=text,
                            style='Secondary.TButton',
                            command=lambda t=template: self.input_entry.insert(0, t))
            btn.pack(side='left', padx=(0, 5))
    
    def send_message(self):
        """Отправка сообщения в чат"""
        message = self.input_entry.get().strip()
        if message and self.chat_manager.send_message(message):
            self.input_entry.delete(0, tk.END)
            self.refresh_chat()
            self.update_unread_count()
    
    def refresh_chat(self):
        """Обновление отображения сообщений"""
        self.messages_text.config(state='normal')
        self.messages_text.delete(1.0, tk.END)
        
        messages = self.chat_manager.get_messages(limit=50)
        messages.reverse()  # Показываем старые сверху, новые снизу
        
        for msg_id, username, fullname, message, timestamp, msg_type in messages:
            # Форматируем время
            try:
                msg_time = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").strftime("%H:%M")
            except:
                msg_time = timestamp
            
            # Определяем префикс в зависимости от типа сообщения
            if msg_type == "system":
                prefix = f"⚡ {msg_time} "
                header_tag = "header_system"
                message_tag = "message_system"
            elif msg_type == "alert":
                prefix = f"🚨 {msg_time} "
                header_tag = "header_alert"
                message_tag = "message_alert"
            else:
                # Определяем текущий пользователь или другой
                if username == self.chat_manager.current_user:
                    prefix = f"👤 Вы ({msg_time}): "
                else:
                    prefix = f"👤 {fullname or username} ({msg_time}): "
                header_tag = "header_text"
                message_tag = "message_text"
            
            self.messages_text.insert(tk.END, prefix, header_tag)
            self.messages_text.insert(tk.END, f"{message}\n\n", message_tag)
        
        self.messages_text.config(state='disabled')
        self.messages_text.see(tk.END)  # Прокрутка к последнему сообщению
        
        # Помечаем сообщения как прочитанные
        self.chat_manager.mark_as_read()
        self.update_unread_count()
    
    def update_unread_count(self):
        """Обновление счетчика непрочитанных"""
        unread_count = self.chat_manager.get_unread_count()
        
        if unread_count > 0:
            self.unread_label.config(text=f"Непрочитанных: {unread_count}")
        else:
            self.unread_label.config(text="")
    
    def flash_notification(self):
        """Мигающее уведомление о новых сообщениях"""
        original_color = self.chat_title.cget('background')
        for i in range(3):
            self.parent.after(i * 500, lambda: self.chat_title.config(bg='yellow'))
            self.parent.after(i * 500 + 250, lambda: self.chat_title.config(bg=original_color))
    
    def get_widget(self):
        """Возвращает основной фрейм чата"""
        return self.chat_frame
    
    def set_current_user(self, username):
        """Установка текущего пользователя"""
        self.chat_manager.current_user = username
        self.refresh_chat()