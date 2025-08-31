import win32print
import win32api
import time
from pathlib import Path

class DualPrinter:
    """Manages finding and printing to two XP-80C printers."""
    def __init__(self):
        self.printers = self._find_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None
        if self.printer1 and self.printer2:
            print(f"✅ Printer 1 (Kitchen): {self.printer1}")
            print(f"✅ Printer 2 (Cashier): {self.printer2}")
        else:
            print(f"❌ Error: Could not find both 'XP-80C' printers. Check connections and driver names.")

    def _find_xprinters(self) -> list[str]:
        """Finds all printers with 'XP-80C' in their name."""
        printers = [p[2] for p in win32print.EnumPrinters(2)]
        xprinters = sorted([p for p in printers if 'XP-80C' in p])
        return xprinters

    def print_file(self, file_path: str | Path) -> bool:
        """Prints a single file to both printers."""
        if not self.printer1 or not self.printer2:
            print("❌ Error: One or both printers are not configured. Cannot print.")
            return False
        
        abs_path = str(Path(file_path).resolve())
        print(f"\n--- Printing file: {abs_path} ---")

        # Print to both printers and return True only if both succeed
        success1 = self._print_to_printer(abs_path, self.printer1, "Kitchen")
        success2 = self._print_to_printer(abs_path, self.printer2, "Cashier")
        
        return success1 and success2

    def _print_to_printer(self, file_path: str, printer_name: str, printer_label: str) -> bool:
        """Temporarily sets the default printer and sends a print job."""
        try:
            current_default = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            time.sleep(3)  # Give time for the job to be spooled
            print(f"✅ Sent to {printer_label} ({printer_name})")
            return True
        except Exception as e:
            print(f"❌ CRITICAL ERROR printing to {printer_label}: {e}")
            return False
        finally:
            # Always try to restore the original default printer
            if 'current_default' in locals():
                win32print.SetDefaultPrinter(current_default)
