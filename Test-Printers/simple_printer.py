import win32print
import win32api
import os
from pathlib import Path

class SimplePrinter:
    """كلاس بسيط للطباعة باستخدام XPrinter"""
    
    def __init__(self):
        self.printer_name = self.find_xprinter()
    
    def find_xprinter(self):
        """البحث عن طابعة XPrinter"""
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        for printer in printers:
            if 'XP' in printer or 'xprinter' in printer.lower():
                return printer
        return printers[0] if printers else None
    
    def print_file(self, file_path):
        if not os.path.exists(file_path):
            print(f"File Not Found : {file_path}")
            return False
        
        try:
            win32api.ShellExecute(0, "print", file_path, f'/d:"{self.printer_name}"', ".", 0)
            print(f"File Sent To Printer : {os.path.basename(file_path)}")
            return True
        except Exception as e:
            print(f"Error In Printing : {e}")
            return False
    
    def list_printers_here(self):

        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        print("Printers :")
        for i, printer in enumerate(printers, 1):
            print(f"{i}. {printer}")


if __name__ == "__main__":
    printer = SimplePrinter()
    print(f"Printer Name: {printer.printer_name}")

    #file_path = Path(__file__).parent / "orders" / "wempy_order_7.docx"
    #printer.print_file(str(file_path))

    file_path = Path(__file__).parent.parent / "orders" / "wempy_order_7.docx" # .. تعني الصعود مجلد واحد للأعلى
    printer.print_file(str(file_path))