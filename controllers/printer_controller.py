"""
This module centralizes all printer-related helpers used by the Wempy backend.
It provides:
    • DualPrinter – sends a document to two printers (if available).
    • print_file_single – prints to one printer with multiple fall-backs.
    • print_file_dual – prints to two printers, falling back to single.
    • print_file – public entry point that defaults to dual printing.

The code is Windows-only and depends on the `pywin32` package.
"""
import os
import time
from pathlib import Path
import platform
import win32api
import win32print
from models import ResponseMessages

class DualPrinter:
    def __init__(self):
        self.printers = self._find_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None

    def _find_xprinters(self):
        """Find all XPrinter devices"""
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        xprinters = []
        for printer in printers:
            if 'xprinter' in printer.lower() or 'xp-' in printer.lower():
                xprinters.append(printer)
        return sorted(xprinters)

    def print_to_both(self, file_path):
        """Print file to both printers - main function"""
        if not self.printer1 or not self.printer2:
            print("❌ Both printers not available")
            return False
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            return False
        print(f"🖨️ Printing to both printers: {os.path.basename(file_path)}")
        result1 = self._print_to_printer(file_path, self.printer1, "Printer 1")
        time.sleep(0.5)
        result2 = self._print_to_printer(file_path, self.printer2, "Printer 2")
        return result1 and result2

    def _print_to_printer(self, file_path, printer_name, printer_label):
        """Internal print function"""
        try:
            current_default = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            time.sleep(1)
            win32print.SetDefaultPrinter(current_default)
            print(f"✅ Sent to {printer_label}: {os.path.basename(file_path)}")
            return True
        except Exception as e:
            print(f"❌ Error printing to {printer_label}: {e}")
            return False

    def get_status(self):
        return {
            'printer1': self.printer1,
            'printer2': self.printer2,
            'both_available': bool(self.printer1 and self.printer2)
        }

def get_xprinter():
    """البحث عن أول طابعة XPrinter متاحة"""
    try:
        printers = win32print.EnumPrinters(2)
        for printer in printers:
            printer_name = printer[2]
            printer_name_lower = printer_name.lower()
            if 'xprinter' in printer_name_lower or 'xp-' in printer_name_lower:
                return printer[2]
    except Exception as e:
        print(f"Error finding XPrinter: {e}")
    return None

def print_file_single(filepath: Path) -> None:
    """طباعة الملف على طابعة واحدة (XPrinter أو الافتراضية)"""
    xprinter = get_xprinter()
    if xprinter and win32api:
        try:
            print(f"Printing to XPrinter: {xprinter}")
            win32api.ShellExecute(0, "print", str(filepath), f'/d:"{xprinter}"', ".", 0)
            print(f"✅ تم إرسال الطلب للطباعة على {xprinter}")
            return
        except Exception as e:
            print(f"XPrinter printing failed: {e}. Trying default method.")
    try:
        os.startfile(str(filepath), "print")  # type: ignore[attr-defined]
        print("✅ تم إرسال الطلب للطباعة على الطابعة الافتراضية")
    except Exception as e:
        print(f"os.startfile failed: {e}. Trying win32api with default printer.")
        if win32api and win32print:
            try:
                default_printer = win32print.GetDefaultPrinter()
                win32api.ShellExecute(0, "print", str(filepath), f'/d:"{default_printer}"', ".", 0)
                print(f"✅ تم إرسال الطلب للطباعة على {default_printer}")
            except Exception as e2:
                print(f"win32api printing failed: {e2}")
                raise Exception(f"فشل في الطباعة: {e2}")

def print_file_dual(filepath: Path, root_path: Path) -> None:
    """طباعة الملف على طابعتين (DualPrinter)، مع معالجة الأخطاء"""
    if platform.system() != "Windows":
        print(ResponseMessages.PRINTER_SUPPORTED_WINDOWS.value)
        return
    if not filepath.exists():
        raise FileNotFoundError(f"{ResponseMessages.PRINTER_FILE_NOT_FOUND.value} : {filepath}")
    try:
        dual_printer = None
        try:
            dual_printer = DualPrinter()
        except Exception as e:
            print(f"DualPrinter init failed: {e}")
        success = dual_printer.print_to_both(str(filepath)) if dual_printer else False
        if success:
            print("✅ تم إرسال الطلب للطباعة على الطابعتين")
        else:
            print("⚠️ فشل في الطباعة المزدوجة، جاري المحاولة بالطريقة الافتراضية...")
            print_file_single(filepath)
    except Exception as e:
        print(f"DualPrinter failed: {e}. Using fallback method.")
        print_file_single(filepath)


def print_file(filepath: Path, root_path: Path) -> None:  # noqa: ARG002
    """دالة الطباعة الرئيسية - تستخدم الطباعة المزدوجة افتراضيًا"""
    print_file_dual(filepath, root_path)

__all__ = [
    "DualPrinter",
    "print_file_single",
    "print_file_dual",
    "print_file",
]
