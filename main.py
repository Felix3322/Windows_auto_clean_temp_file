import os
import shutil
import tempfile
import ctypes
import sys
from send2trash import send2trash
import tkinter as tk
from tkinter import messagebox
import win32com.shell.shell as shell


class CleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Cleaner -by BUFF")

        self.temp_var = tk.BooleanVar()
        self.recycle_var = tk.BooleanVar()
        self.downloads_var = tk.BooleanVar()
        self.browser_cache_var = tk.BooleanVar()
        self.system_logs_var = tk.BooleanVar()
        self.app_logs_var = tk.BooleanVar()
        self.prefetch_var = tk.BooleanVar()
        self.update_cache_var = tk.BooleanVar()

        tk.Label(root, text="Select items to clean:").pack()

        tk.Checkbutton(root, text="Temporary Files", variable=self.temp_var).pack()
        tk.Checkbutton(root, text="Recycle Bin", variable=self.recycle_var).pack()
        tk.Checkbutton(root, text="Downloads Folder", variable=self.downloads_var).pack()
        tk.Checkbutton(root, text="Browser Cache", variable=self.browser_cache_var).pack()
        tk.Checkbutton(root, text="System Logs", variable=self.system_logs_var).pack()
        tk.Checkbutton(root, text="Application Logs", variable=self.app_logs_var).pack()
        tk.Checkbutton(root, text="Prefetch Files", variable=self.prefetch_var).pack()
        tk.Checkbutton(root, text="Windows Update Cache", variable=self.update_cache_var).pack()

        tk.Button(root, text="Clean", command=self.clean).pack()

        self.output_text = tk.Text(root, height=15, width=50)
        self.output_text.pack()

    def log(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update()

    def delete_temp_files(self):
        temp_dir = tempfile.gettempdir()
        count, size = 0, 0
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    size += os.path.getsize(file_path)
                    os.remove(file_path)
                    count += 1
                    self.log(f"Deleted: {file_path}")
                except Exception as e:
                    self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def empty_recycle_bin(self):
        recycle_bin = "C:\\$Recycle.Bin"
        count, size = 0, 0
        for root, dirs, files in os.walk(recycle_bin):
            for dir in dirs:
                dir_path = os.path.join(root, dir)
                try:
                    for root, dirs, files in os.walk(dir_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            size += os.path.getsize(file_path)
                            os.remove(file_path)
                            count += 1
                    shutil.rmtree(dir_path)
                    self.log(f"Emptied: {dir_path}")
                except Exception as e:
                    self.log(f"Failed to empty {dir_path}: {e}")
        return count, size

    def clean_downloads_folder(self):
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        count, size = 0, 0
        for root, dirs, files in os.walk(downloads_folder):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    size += os.path.getsize(file_path)
                    send2trash(file_path)
                    count += 1
                    self.log(f"Sent to trash: {file_path}")
                except Exception as e:
                    self.log(f"Failed to send {file_path} to trash: {e}")
        return count, size

    def clean_browser_cache(self):
        cache_paths = [
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Google", "Chrome", "User Data", "Default",
                         "Cache"),
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Mozilla", "Firefox", "Profiles")
        ]
        count, size = 0, 0
        for cache_path in cache_paths:
            if os.path.exists(cache_path):
                for root, dirs, files in os.walk(cache_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size += os.path.getsize(file_path)
                            os.remove(file_path)
                            count += 1
                            self.log(f"Deleted: {file_path}")
                        except Exception as e:
                            self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def clean_system_logs(self):
        log_paths = [
            "C:\\Windows\\System32\\winevt\\Logs"
        ]
        count, size = 0, 0
        for log_path in log_paths:
            if os.path.exists(log_path):
                for root, dirs, files in os.walk(log_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size += os.path.getsize(file_path)
                            os.remove(file_path)
                            count += 1
                            self.log(f"Deleted: {file_path}")
                        except Exception as e:
                            self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def clean_app_logs(self):
        log_paths = [
            "C:\\ProgramData\\Microsoft\\Windows\\WER\\ReportQueue"
        ]
        count, size = 0, 0
        for log_path in log_paths:
            if os.path.exists(log_path):
                for root, dirs, files in os.walk(log_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size += os.path.getsize(file_path)
                            os.remove(file_path)
                            count += 1
                            self.log(f"Deleted: {file_path}")
                        except Exception as e:
                            self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def clean_prefetch_files(self):
        prefetch_path = "C:\\Windows\\Prefetch"
        count, size = 0, 0
        for root, dirs, files in os.walk(prefetch_path):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    size += os.path.getsize(file_path)
                    os.remove(file_path)
                    count += 1
                    self.log(f"Deleted: {file_path}")
                except Exception as e:
                    self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def clean_update_cache(self):
        update_cache_path = "C:\\Windows\\SoftwareDistribution\\Download"
        count, size = 0, 0
        for root, dirs, files in os.walk(update_cache_path):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    size += os.path.getsize(file_path)
                    os.remove(file_path)
                    count += 1
                    self.log(f"Deleted: {file_path}")
                except Exception as e:
                    self.log(f"Failed to delete {file_path}: {e}")
        return count, size

    def clean(self):
        self.output_text.delete(1.0, tk.END)
        self.log("Starting cleanup...")

        total_count, total_size = 0, 0

        if self.temp_var.get():
            self.log("Cleaning temporary files...")
            count, size = self.delete_temp_files()
            total_count += count
            total_size += size

        if self.recycle_var.get():
            self.log("Emptying recycle bin...")
            count, size = self.empty_recycle_bin()
            total_count += count
            total_size += size

        if self.downloads_var.get():
            self.log("Cleaning downloads folder...")
            count, size = self.clean_downloads_folder()
            total_count += count
            total_size += size

        if self.browser_cache_var.get():
            self.log("Cleaning browser cache...")
            count, size = self.clean_browser_cache()
            total_count += count
            total_size += size

        if self.system_logs_var.get():
            self.log("Cleaning system logs...")
            count, size = self.clean_system_logs()
            total_count += count
            total_size += size

        if self.app_logs_var.get():
            self.log("Cleaning application logs...")
            count, size = self.clean_app_logs()
            total_count += count
            total_size += size

        if self.prefetch_var.get():
            self.log("Cleaning prefetch files...")
            count, size = self.clean_prefetch_files()
            total_count += count
            total_size += size

        if self.update_cache_var.get():
            self.log("Cleaning Windows Update cache...")
            count, size = self.clean_update_cache()
            total_count += count
            total_size += size

        self.log(
            f"Cleanup complete. Total files deleted: {total_count}, Total space freed: {total_size / (1024 * 1024):.2f} MB.")
        messagebox.showinfo("Completed",
                            f"Cleanup complete.\nTotal files deleted: {total_count}\nTotal space freed: {total_size / (1024 * 1024):.2f} MB.")


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if __name__ == "__main__":
    if is_admin():
        root = tk.Tk()
        app = CleanerApp(root)
        root.mainloop()
    else:
        script = os.path.abspath(__file__)
        params = ' '.join(['"' + arg + '"' for arg in sys.argv[1:]])
        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=script + ' ' + params)
