import os
import win32print
import win32api

# اسم الطابعة التي تريد جعلها افتراضية
target_printer = "XP-80C"

# اجعلها الطابعة الافتراضية
win32print.SetDefaultPrinter(target_printer)

# اطبع الملف
file_path = os.path.abspath("2.txt")
win32api.ShellExecute(0, "print", file_path, f'/d:"{target_printer}"', ".", 0)
