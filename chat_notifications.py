import sqlite3
from datetime import datetime, timedelta
import os
from chat_manager import ChatManager

class ChatNotifications:
    def __init__(self, chat_manager):
        self.chat_manager = chat_manager
        self.setup_automatic_messages()
    
    def setup_automatic_messages(self):
        """Настройка автоматических сообщений"""
        self.send_daily_greeting()
        self.send_ippcu_reminders()
        self.send_birthday_reminders()
    
    def send_daily_greeting(self):
        """Ежедневное приветственное сообщение"""
        # Проверяем, не отправляли ли сегодня уже приветствие
        today = datetime.now().strftime("%Y-%m-%d")
        try:
            with sqlite3.connect(os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp", "clients.db")) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT COUNT(*) FROM chat_messages 
                    WHERE message LIKE 'Доброе утро%' 
                    AND DATE(timestamp) = ?
                """, (today,))
                count = cur.fetchone()[0]
                
                if count == 0:
                    greeting = f"Доброе утро! Сегодня {datetime.now().strftime('%d.%m.%Y')}. Удачи в работе! 🌞"
                    self.chat_manager.send_system_message(greeting)
        except Exception as e:
            print(f"Ошибка отправки приветствия: {e}")
    
    def send_ippcu_reminders(self):
        """Напоминания о ИППСУ через чат"""
        today = datetime.today().date()
        soon = today + timedelta(days=3)
        
        try:
            with sqlite3.connect(os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp", "clients.db")) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT last_name, first_name, ippcu_end 
                    FROM clients 
                    WHERE ippcu_end BETWEEN ? AND ?
                """, (today.strftime("%Y-%m-%d"), soon.strftime("%Y-%m-%d")))
                
                expiring_clients = cur.fetchall()
                
                for last_name, first_name, ippcu_end in expiring_clients:
                    end_date = datetime.strptime(ippcu_end, "%Y-%m-%d").date()
                    days_left = (end_date - today).days
                    
                    if days_left <= 3:
                        message = f"⚠️ Внимание! ИППСУ клиента {last_name} {first_name} истекает через {days_left} дн. ({end_date.strftime('%d.%m.%Y')})"
                        self.chat_manager.send_alert_message(message)
                        
        except Exception as e:
            print(f"Ошибка отправки напоминаний ИППСУ: {e}")
    
    def send_birthday_reminders(self):
        """Напоминания о днях рождения"""
        today = datetime.today().date()
        next_week = today + timedelta(days=7)
        
        try:
            with sqlite3.connect(os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp", "clients.db")) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT last_name, first_name, middle_name, dob 
                    FROM clients 
                    WHERE substr(dob, 6, 5) BETWEEN ? AND ?
                """, (today.strftime("%m-%d"), next_week.strftime("%m-%d")))
                
                birthdays = cur.fetchall()
                
                for last_name, first_name, middle_name, dob in birthdays:
                    bday = datetime.strptime(dob, "%Y-%m-%d").date()
                    bday_this_year = bday.replace(year=today.year)
                    days_until = (bday_this_year - today).days
                    
                    if days_until >= 0:
                        middle = f" {middle_name}" if middle_name else ""
                        message = f"🎂 Через {days_until} дн. день рождения у {last_name} {first_name}{middle} ({bday.strftime('%d.%m')})"
                        self.chat_manager.send_system_message(message)
                        
        except Exception as e:
            print(f"Ошибка отправки напоминаний о днях рождения: {e}")
    
    def send_custom_alert(self, message, alert_type="system"):
        """Отправка пользовательского уведомления"""
        if alert_type == "alert":
            self.chat_manager.send_alert_message(message)
        else:
            self.chat_manager.send_system_message(message)