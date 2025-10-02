import sqlite3
import os
from datetime import datetime
from tkinter import messagebox

# Пути
APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
DB_NAME = os.path.join(APP_DIR, "clients.db")

class ChatManager:
    def __init__(self):
        self.current_user = "user1"  # Можно сделать выбор пользователя
        self.unread_count = 0
        self.init_chat_tables()
        
    def init_chat_tables(self):
        """Инициализация таблиц чата"""
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            
            # Таблица для сообщений чата
            cur.execute("""
                CREATE TABLE IF NOT EXISTS chat_messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_name TEXT NOT NULL,
                    message TEXT NOT NULL,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                    message_type TEXT DEFAULT 'text',
                    is_read INTEGER DEFAULT 0
                )
            """)
            
            # Таблица для пользователей чата
            cur.execute("""
                CREATE TABLE IF NOT EXISTS chat_users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_name TEXT UNIQUE NOT NULL,
                    full_name TEXT NOT NULL,
                    role TEXT DEFAULT 'employee',
                    is_online INTEGER DEFAULT 0,
                    last_seen DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Добавляем стандартных пользователей
            default_users = [
                ("admin", "Зеленков Д.В.", "admin"),
                ("manager", "Дурандина А.В.", "manager"),
                ("social1", "Социальный работник 1", "employee"),
                ("social2", "Социальный работник 2", "employee"),
                ("social3", "Социальный работник 3", "employee")
            ]
            
            for username, fullname, role in default_users:
                try:
                    cur.execute(
                        "INSERT OR IGNORE INTO chat_users (user_name, full_name, role) VALUES (?, ?, ?)",
                        (username, fullname, role)
                    )
                except Exception as e:
                    print(f"Ошибка добавления пользователя {username}: {e}")
            
            conn.commit()
    
    def send_message(self, message, message_type="text"):
        """Отправка сообщения в чат"""
        if not message.strip():
            return False
            
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute(
                    "INSERT INTO chat_messages (user_name, message, message_type) VALUES (?, ?, ?)",
                    (self.current_user, message.strip(), message_type)
                )
                conn.commit()
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отправить сообщение: {e}")
            return False
    
    def get_messages(self, limit=100, offset=0):
        """Получение сообщений из чата"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT cm.id, cm.user_name, cu.full_name, cm.message, cm.timestamp, cm.message_type
                    FROM chat_messages cm
                    LEFT JOIN chat_users cu ON cm.user_name = cu.user_name
                    ORDER BY cm.timestamp DESC
                    LIMIT ? OFFSET ?
                """, (limit, offset))
                return cur.fetchall()
        except Exception as e:
            print(f"Ошибка получения сообщений: {e}")
            return []
    
    def get_unread_count(self):
        """Получение количества непрочитанных сообщений"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM chat_messages WHERE is_read = 0 AND user_name != ?", 
                           (self.current_user,))
                return cur.fetchone()[0]
        except:
            return 0
    
    def mark_as_read(self):
        """Пометить все сообщения как прочитанные"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("UPDATE chat_messages SET is_read = 1 WHERE user_name != ?", 
                           (self.current_user,))
                conn.commit()
        except Exception as e:
            print(f"Ошибка пометки сообщений как прочитанных: {e}")
    
    def get_online_users(self):
        """Получить список онлайн пользователей"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT full_name, role, last_seen 
                    FROM chat_users 
                    WHERE is_online = 1 
                    ORDER BY role, full_name
                """)
                return cur.fetchall()
        except:
            return []
    
    def set_user_online(self, online=True):
        """Установить статус пользователя"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("""
                    UPDATE chat_users 
                    SET is_online = ?, last_seen = CURRENT_TIMESTAMP 
                    WHERE user_name = ?
                """, (1 if online else 0, self.current_user))
                conn.commit()
        except Exception as e:
            print(f"Ошибка установки статуса пользователя: {e}")
    
    def get_user_info(self, username=None):
        """Получить информацию о пользователе"""
        if username is None:
            username = self.current_user
            
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("SELECT full_name, role FROM chat_users WHERE user_name = ?", (username,))
                return cur.fetchone()
        except:
            return None
    
    def send_system_message(self, message):
        """Отправить системное сообщение"""
        return self.send_message(message, "system")
    
    def send_alert_message(self, message):
        """Отправить сообщение-предупреждение"""
        return self.send_message(message, "alert")
    
    def clear_chat_history(self):
        """Очистить историю чата (только для админа)"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cur = conn.cursor()
                cur.execute("DELETE FROM chat_messages")
                conn.commit()
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось очистить историю: {e}")
            return False