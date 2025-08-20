import win32print
import win32api
import os

# استخدم raw string لتفادي مشاكل الباكسلاش
file_path = os.path.abspath(r"8.docx")

# تأكد أن الطابعة الافتراضية مضبوطة
printer_name = win32print.GetDefaultPrinter()

# تنفيذ أمر الطباعة
win32api.ShellExecute(
    0,
    "print",
    file_path,
    f'/d:"{printer_name}"',
    ".",
    0
)
