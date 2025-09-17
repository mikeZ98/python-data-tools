import shutil
import os
import datetime
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import winreg
import time
import sys

class BackupHandler(FileSystemEventHandler):
    def __init__(self, src_folder, dst_folder, status_label, backup_mode, backup_time=None):
        self.src_folder = src_folder
        self.dst_folder = dst_folder
        self.status_label = status_label
        self.backup_mode = backup_mode
        self.backup_time = backup_time
        self.backup_thread = None
        self.stop_event = threading.Event()

    def start_backup_thread(self):
        self.stop_event.clear()
        self.backup_thread = threading.Thread(target=self.run_backup_loop)
        self.backup_thread.start()

    def stop_backup_thread(self):
        self.stop_event.set()
        if self.backup_thread:
            self.backup_thread.join()
        time.sleep(0.5)

    def run_backup_loop(self):
        if self.backup_mode == 'interval':
            while not self.stop_event.is_set():
                now = datetime.datetime.now()
                next_backup = now.replace(hour=self.backup_time.hour, minute=self.backup_time.minute, second=0, microsecond=0)
                if now > next_backup:
                    next_backup += datetime.timedelta(days=1)
                sleep_time = (next_backup - now).total_seconds()
                self.status_label.config(text=f"Oczekiwanie do {next_backup.strftime('%H:%M:%S')} na następne kopiowanie...")
                self.stop_event.wait(sleep_time)
                if not self.stop_event.is_set():
                    self.perform_backup()
        else:
            self.perform_backup()
            self.stop_event.wait(60)

    def on_any_event(self, event):
        if self.backup_mode == 'automatic':
            if event.event_type in ['modified', 'created', 'deleted']:
                self.perform_backup()

    def perform_backup(self):
        now = datetime.datetime.now()
        backup_folder = f'{self.dst_folder}/backup_{now.strftime("%Y%m%d_%H%M%S")}'
        copy_files(self.src_folder, backup_folder, self.status_label)
        self.status_label.config(text="Kopia zapasowa utworzona.")

def copy_files(src_folder, dst_folder, status_label):
    if not os.path.exists(src_folder):
        status_label.config(text="Folder źródłowy nie istnieje!")
        return

    try:
        if not os.path.exists(dst_folder):
            os.makedirs(dst_folder)

        for item in os.listdir(src_folder):
            src_path = os.path.join(src_folder, item)
            dst_path = os.path.join(dst_folder, item)

            try:
                if os.path.isdir(src_path):
                    if os.path.exists(dst_path):
                        shutil.rmtree(dst_path)
                    shutil.copytree(src_path, dst_path)
                elif os.path.isfile(src_path):
                    if os.access(src_path, os.R_OK):
                        shutil.copy2(src_path, dst_path)
            except PermissionError:
                status_label.config(text=f"Brak dostępu do pliku: {src_path}")
            except Exception as e:
                status_label.config(text=f"Błąd przy kopiowaniu {src_path} do {dst_path}: {e}")
    except Exception as e:
        status_label.config(text=f"Błąd podczas tworzenia kopii zapasowej: {e}")

def on_close():
    if handler:
        stop_monitoring()
    root.destroy()

def choose_src_folder():
    if not monitoring_active:
        global src_folder_path
        src_folder_path = filedialog.askdirectory()
        src_folder_label.config(text=f"Folder źródłowy: {src_folder_path}")

def choose_dst_folder():
    if not monitoring_active:
        global dst_folder_path
        dst_folder_path = filedialog.askdirectory()
        dst_folder_label.config(text=f"Folder docelowy: {dst_folder_path}")

def start_monitoring():
    global observer, handler, monitoring_active
    if not src_folder_path or not dst_folder_path:
        messagebox.showwarning("Uwaga", "Musisz wybrać zarówno folder źródłowy, jak i docelowy!")
        return

    backup_mode = mode_var.get()
    backup_time = None
    if backup_mode == 'interval':
        try:
            hour = int(hour_entry.get())
            minute = int(minute_entry.get())
            if hour < 0 or hour > 23 or minute < 0 or minute > 59:
                raise ValueError
            backup_time = datetime.time(hour, minute)
        except ValueError:
            messagebox.showwarning("Uwaga", "Wprowadź poprawny czas w formacie HH:MM!")
            return

    if handler:
        stop_monitoring()

    handler = BackupHandler(src_folder_path, dst_folder_path, status_label, backup_mode, backup_time)
    observer = Observer()
    observer.schedule(handler, path=src_folder_path, recursive=True)
    observer.start()
    handler.start_backup_thread()
    status_label.config(text="Tworzenie kopii zapasowej uruchomione.")
    monitoring_active = True
    toggle_controls(False)

def stop_monitoring():
    global observer, handler, monitoring_active
    if handler:
        handler.stop_backup_thread()
        handler = None
    if observer:
        observer.stop()
        observer.join()
        observer = None
    status_label.config(text="Tworzenie kopii zapasowej zatrzymane.")
    monitoring_active = False
    toggle_controls(True)

def toggle_controls(state):
    src_button.config(state=tk.NORMAL if state else tk.DISABLED)
    dst_button.config(state=tk.NORMAL if state else tk.DISABLED)
    auto_radio.config(state=tk.NORMAL if state else tk.DISABLED)
    interval_radio.config(state=tk.NORMAL if state else tk.DISABLED)
    hour_entry.config(state=tk.NORMAL if state else tk.DISABLED)
    minute_entry.config(state=tk.NORMAL if state else tk.DISABLED)
    start_button.config(state=tk.NORMAL if state else tk.DISABLED)
    stop_button.config(state=tk.NORMAL if not state else tk.DISABLED)

def add_to_autostart():
    try:
        if not os.access(__file__, os.W_OK):
            raise PermissionError("Brak dostępu do pliku lub wymagane są uprawnienia administratora.")

        key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(reg, "BackupApp", 0, winreg.REG_SZ, f'"{os.path.abspath(__file__)}"')
        winreg.CloseKey(reg)
        messagebox.showinfo("Autostart", "Program został dodany do autostartu.")
    except PermissionError as e:
        messagebox.showerror("Błąd uprawnień", str(e))
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się dodać programu do autostartu: {e}")

stop_event = threading.Event()
handler = None
observer = None
monitoring_active = False

root = tk.Tk()
root.title("Monitorowanie i Kopiowanie Plików")
root.geometry("600x500")

# Wybieranie folderów przez użytkownika
src_folder_path = ''
dst_folder_path = ''

# Ustawienie etykiet
src_folder_label = tk.Label(root, text="Folder źródłowy: (nie wybrano)")
src_folder_label.pack(pady=5)
dst_folder_label = tk.Label(root, text="Folder docelowy: (nie wybrano)")
dst_folder_label.pack(pady=5)

src_button = tk.Button(root, text="Wybierz folder źródłowy", command=choose_src_folder)
src_button.pack(pady=10)
dst_button = tk.Button(root, text="Wybierz folder docelowy", command=choose_dst_folder)
dst_button.pack(pady=10)

mode_var = tk.StringVar(value='interval')
auto_radio = tk.Radiobutton(root, text="Tryb obserwacji zmian plików", variable=mode_var, value='automatic')
auto_radio.pack(pady=5)
interval_radio = tk.Radiobutton(root, text="Interwał godzinowy", variable=mode_var, value='interval')
interval_radio.pack(pady=5)

time_frame = tk.Frame(root)
time_frame.pack(pady=10)
tk.Label(time_frame, text="Godzina:").pack(side=tk.LEFT)
hour_entry = tk.Entry(time_frame, width=5)
hour_entry.insert(0, "17")
hour_entry.pack(side=tk.LEFT)
tk.Label(time_frame, text="Minuta:").pack(side=tk.LEFT)
minute_entry = tk.Entry(time_frame, width=5)
minute_entry.insert(0, "00")
minute_entry.pack(side=tk.LEFT)

start_button = tk.Button(root, text="Rozpocznij backup", command=start_monitoring)
start_button.pack(pady=10)
stop_button = tk.Button(root, text="Zatrzymaj backup", command=stop_monitoring)
stop_button.pack(pady=10)
end_button = tk.Button(root, text="Zamknij", command=on_close)
end_button.pack(pady=10)

status_label = tk.Label(root, text="Status: Backup nie uruchomiony.")
status_label.pack(pady=10)

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
