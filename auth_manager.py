import sqlite3
import hashlib
import os
import json
import secrets
from datetime import datetime, timedelta

class AuthManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.current_user = None
        self.remember_me = False
        self.init_auth_db()
        self.load_remembered_user()
    
    def init_auth_db(self):
        """Инициализация базы данных пользователей"""
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    full_name TEXT NOT NULL,
                    role TEXT NOT NULL,
                    permissions TEXT DEFAULT 'basic',
                    is_active INTEGER DEFAULT 1,
                    last_login TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            cur.execute("""
                CREATE TABLE IF NOT EXISTS remember_tokens (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    token_hash TEXT NOT NULL,
                    expires_at TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                )
            """)
            
            # Добавляем стандартных пользователей
            default_users = [
                {
                    'username': 'admin',
                    'password': 'admin',
                    'full_name': 'Зеленков Д.В.',
                    'role': 'младший администратор БД (Главный)',
                    'permissions': 'all'
                },
                {
                    'username': 'ДУРАНДИНА',
                    'password': '12345',
                    'full_name': 'Дурандина А.В.',
                    'role': 'Заведующая',
                    'permissions': 'all'
                },
                {
                    'username': 'ЛАВРОВА',
                    'password': '12345', 
                    'full_name': 'Лаврова А.А.',
                    'role': 'Сотрудник',
                    'permissions': 'all'
                }
            ]
            
            for user in default_users:
                password_hash = self.hash_password(user['password'])
                try:
                    cur.execute("""
                        INSERT OR IGNORE INTO users 
                        (username, password_hash, full_name, role, permissions) 
                        VALUES (?, ?, ?, ?, ?)
                    """, (user['username'], password_hash, user['full_name'], 
                          user['role'], user['permissions']))
                except sqlite3.IntegrityError:
                    pass
            
            conn.commit()
    
    def hash_password(self, password):
        """Хеширование пароля"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verify_password(self, password, password_hash):
        """Проверка пароля"""
        return self.hash_password(password) == password_hash
    
    def generate_remember_token(self):
        """Генерация токена для запоминания"""
        return secrets.token_urlsafe(32)
    
    def hash_token(self, token):
        """Хеширование токена"""
        return hashlib.sha256(token.encode()).hexdigest()
    
    def create_remember_token(self, user_id):
        """Создание токена запоминания на 30 дней"""
        token = self.generate_remember_token()
        token_hash = self.hash_token(token)
        expires_at = datetime.now() + timedelta(days=30)
        
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            # Удаляем старые токены пользователя
            cur.execute("DELETE FROM remember_tokens WHERE user_id = ?", (user_id,))
            # Добавляем новый токен
            cur.execute("""
                INSERT INTO remember_tokens (user_id, token_hash, expires_at)
                VALUES (?, ?, ?)
            """, (user_id, token_hash, expires_at.isoformat()))
            conn.commit()
        
        return token
    
    def verify_remember_token(self, token):
        """Проверка токена запоминания"""
        if not token:
            return None
        
        token_hash = self.hash_token(token)
        
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT rt.user_id, u.username, u.full_name, u.role, u.permissions
                FROM remember_tokens rt
                JOIN users u ON rt.user_id = u.id
                WHERE rt.token_hash = ? AND rt.expires_at > ? AND u.is_active = 1
            """, (token_hash, datetime.now().isoformat()))
            
            result = cur.fetchone()
            
            if result:
                user_id, username, full_name, role, permissions = result
                return {
                    'id': user_id,
                    'username': username,
                    'full_name': full_name,
                    'role': role,
                    'permissions': permissions.split(',') if permissions else ['basic']
                }
        
        return None
    
    def save_remember_token_to_file(self, token):
        """Сохранение токена в файл"""
        APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
        os.makedirs(APP_DIR, exist_ok=True)
        token_file = os.path.join(APP_DIR, "remember_token.json")
        token_data = {
            'token': token,
            'saved_at': datetime.now().isoformat()
        }
        
        try:
            with open(token_file, 'w', encoding='utf-8') as f:
                json.dump(token_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения токена: {e}")
    
    def load_remember_token_from_file(self):
        """Загрузка токена из файла"""
        APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
        token_file = os.path.join(APP_DIR, "remember_token.json")
        
        if not os.path.exists(token_file):
            return None
        
        try:
            with open(token_file, 'r', encoding='utf-8') as f:
                token_data = json.load(f)
            
            # Проверяем, не старше ли токен 25 дней (обновляем заранее)
            saved_at = datetime.fromisoformat(token_data['saved_at'])
            if datetime.now() - saved_at > timedelta(days=25):
                self.clear_remember_token()
                return None
            
            return token_data['token']
        except Exception as e:
            print(f"Ошибка загрузки токена: {e}")
            return None
    
    def clear_remember_token(self):
        """Очистка токена"""
        APP_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "MyApp")
        token_file = os.path.join(APP_DIR, "remember_token.json")
        if os.path.exists(token_file):
            try:
                os.remove(token_file)
            except Exception as e:
                print(f"Ошибка удаления токена: {e}")
    
    def load_remembered_user(self):
        """Загрузка запомненного пользователя"""
        token = self.load_remember_token_from_file()
        if token:
            user_data = self.verify_remember_token(token)
            if user_data:
                self.current_user = user_data
                self.remember_me = True
                # Обновляем время последнего входа
                with sqlite3.connect(self.db_path) as conn:
                    cur = conn.cursor()
                    cur.execute("UPDATE users SET last_login = ? WHERE id = ?", 
                               (datetime.now().isoformat(), user_data['id']))
                    conn.commit()
                return True
        return False
    
    def login(self, username, password, remember_me=False):
        """Аутентификация пользователя"""
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, username, password_hash, full_name, role, permissions 
                FROM users 
                WHERE username = ? AND is_active = 1
            """, (username,))
            
            user_data = cur.fetchone()
            
            if user_data and self.verify_password(password, user_data[2]):
                user_id, username, _, full_name, role, permissions = user_data
                
                # Обновляем время последнего входа
                cur.execute("UPDATE users SET last_login = ? WHERE id = ?", 
                           (datetime.now().isoformat(), user_id))
                
                # Создаем токен запоминания если нужно
                if remember_me:
                    token = self.create_remember_token(user_id)
                    self.save_remember_token_to_file(token)
                    self.remember_me = True
                else:
                    self.clear_remember_token()
                    self.remember_me = False
                
                conn.commit()
                
                self.current_user = {
                    'id': user_id,
                    'username': username,
                    'full_name': full_name,
                    'role': role,
                    'permissions': permissions.split(',') if permissions else ['basic']
                }
                
                return True, "Успешный вход"
            else:
                return False, "Неверный логин или пароль"
    
    def logout(self):
        """Выход пользователя"""
        self.clear_remember_token()
        self.current_user = None
        self.remember_me = False
    
    def has_permission(self, permission):
        """Проверка прав доступа"""
        if not self.current_user:
            return False
        
        if 'all' in self.current_user['permissions']:
            return True
        
        return permission in self.current_user['permissions']
    
    def get_current_user(self):
        """Получить текущего пользователя"""
        return self.current_user
    
    def get_user_display_name(self):
        """Получить отображаемое имя пользователя"""
        if self.current_user:
            return f"{self.current_user['full_name']} ({self.current_user['role']})"
        return "Не авторизован"
    
    def cleanup_expired_tokens(self):
        """Очистка просроченных токенов"""
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM remember_tokens WHERE expires_at < ?", 
                       (datetime.now().isoformat(),))
            conn.commit()
