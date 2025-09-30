import requests
import os
import sys
import subprocess
import time
import threading
import tkinter as tk
from tkinter import messagebox, ttk

APP_VERSION = "2.0.0"  # <<< измени на номер своей текущей версии
VERSION_URL = "VERSION_URL = "https://raw.githubusercontent.com/amberbeksky/ODP2/main/version.txt"
APP_URL = "https://github.com/amberbeksky/ODP2/releases/latest/download/app.exe"
APP_PATH = sys.argv[0]  # путь к текущему exe


def check_for_update():
    try:
        resp = requests.get(VERSION_URL, timeout=5)
        latest_version = resp.text.strip()
        if latest_version > APP_VERSION:
            return latest_version
    except Exception as e:
        print("Ошибка проверки обновления:", e)
    return None


def threaded_download(url, filepath, win, progress, percent_label, speed_label, eta_label):
    try:
        resp = requests.get(url, stream=True, timeout=30)
        total = int(resp.headers.get("content-length", 0))

        progress["maximum"] = total
        downloaded = 0
        start_time = time.time()

        with open(filepath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)

                    def update_labels(d=downloaded, t=total, st=start_time):
                        progress["value"] = d
                        percent = (d / t) * 100 if t > 0 else 0
                        percent_label.config(text=f"{percent:.1f}%")

                        elapsed = time.time() - st
                        speed = d / 1024 / elapsed if elapsed > 0 else 0
                        speed_label.config(text=f"Скорость: {speed:.1f} КБ/с")

                        if speed > 0:
                            remaining_kb = (t - d) / 1024
                            eta_sec = remaining_kb / speed
                            mins, secs = divmod(int(eta_sec), 60)
                            eta_label.config(text=f"Оставшееся время: {mins:02d}:{secs:02d}")
                        else:
                            eta_label.config(text="Оставшееся время: --:--")

                    win.after(0, update_labels)

        win.after(0, win.destroy)

    except Exception as e:
        win.after(0, lambda: messagebox.showerror("Ошибка", f"Не удалось скачать обновление: {e}"))


def download_with_progress(url, filepath):
    win = tk.Toplevel()
    win.title("Обновление программы")
    win.geometry("420x180")
    win.resizable(False, False)

    tk.Label(win, text="Загружается новая версия...").pack(pady=10)

    progress = ttk.Progressbar(win, length=380, mode="determinate")
    progress.pack(pady=5)

    percent_label = tk.Label(win, text="0%")
    percent_label.pack()

    speed_label = tk.Label(win, text="Скорость: 0 КБ/с")
    speed_label.pack()

    eta_label = tk.Label(win, text="Оставшееся время: --:--")
    eta_label.pack()

    thread = threading.Thread(
        target=threaded_download,
        args=(url, filepath, win, progress, percent_label, speed_label, eta_label),
        daemon=True
    )
    thread.start()

    win.grab_set()
    win.wait_window()


def download_and_replace():
    try:
        new_path = APP_PATH + ".new"
        download_with_progress(APP_URL, new_path)

        backup = APP_PATH + ".old"
        if os.path.exists(backup):
            os.remove(backup)
        os.rename(APP_PATH, backup)
        os.rename(new_path, APP_PATH)

        messagebox.showinfo("Обновление", "Обновление завершено. Программа будет перезапущена.")
        subprocess.Popen([APP_PATH])
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обновить: {e}")


def auto_update():
    new_ver = check_for_update()
    if new_ver:
        download_and_replace()
