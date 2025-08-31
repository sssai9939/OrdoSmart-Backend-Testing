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
from models.enums import ResponseMessages

# --- Conditional Import for Windows-specific modules ---
IS_WINDOWS = platform.system() == "Windows"
win32api = None
win32print = None

if IS_WINDOWS:
    try:
        import win32api
        import win32print
    except ImportError:
        print("⚠️ pywin32 not found. Printing functionality will be disabled.")
        IS_WINDOWS = False  # Force disable if imports fail

# --- Platform-specific Implementations ---

if IS_WINDOWS:
    # --- Windows-specific (real) implementation ---
    class DualPrinter:
        def __init__(self):
            self.printers = self._find_xprinters()
            self.printer1 = self.printers[0] if len(self.printers) > 0 else None
            self.printer2 = self.printers[1] if len(self.printers) > 1 else None

        def _find_xprinters(self):
            printers = [p[2] for p in win32print.EnumPrinters(2)]
            return sorted([p for p in printers if 'xprinter' in p.lower() or 'xp-' in p.lower()])

        def print_to_both(self, file_path):
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
            return {'printer1': self.printer1, 'printer2': self.printer2, 'both_available': bool(self.printer1 and self.printer2)}

    def print_file(filepath: Path, root_path: Path) -> None:  # noqa: ARG001
        if not filepath.exists():
            raise FileNotFoundError(f"{ResponseMessages.PRINTER_FILE_NOT_FOUND.value} : {filepath}")
        try:
            dual_printer = DualPrinter()
            if not dual_printer.print_to_both(str(filepath)):
                print("⚠️ Dual printing failed, falling back to default print.")
                os.startfile(str(filepath), "print")
        except Exception as e:
            print(f"🚨 An error occurred during printing: {e}. Trying default print.")
            os.startfile(str(filepath), "print")

else:
    # --- Non-Windows (dummy) implementation ---
    class DualPrinter:
        def __init__(self):
            print("Printer support disabled (non-Windows environment).")
        def print_to_both(self, file_path):
            print(f"(Skipping print for {file_path} - non-Windows)")
            return False
        def get_status(self):
            return {'printer1': None, 'printer2': None, 'both_available': False}

    def print_file(filepath: Path, root_path: Path) -> None:  # noqa: ARG001
        print(f"(Skipping print for {filepath} - non-Windows environment)")
        pass

__all__ = [
    "DualPrinter",
    "print_file",
]
