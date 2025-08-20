import win32print
import win32api
import os
import time
from pathlib import Path

class DualPrinter:
    """Simple dual XPrinter controller - prints one file to both printers"""
    
    def __init__(self):
        self.printers = self._find_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None
        
    def _find_xprinters(self):
        """Find all XPrinter devices"""
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        xprinters = []
        for printer in printers:
            if 'XP' in printer and 'copy' in printer.lower():
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
        
        # Print to first printer
        result1 = self._print_to_printer(file_path, self.printer1, "Printer 1")
        time.sleep(0.5)
        
        # Print to second printer  
        result2 = self._print_to_printer(file_path, self.printer2, "Printer 2")
        
        return result1 and result2
    
    def _print_to_printer(self, file_path, printer_name, printer_label):
        """Internal print function"""
        try:
            # Save current default printer
            current_default = win32print.GetDefaultPrinter()
            
            # Set target printer as default temporarily
            win32print.SetDefaultPrinter(printer_name)
            
            # Print
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            
            # Restore original default printer
            time.sleep(1)
            win32print.SetDefaultPrinter(current_default)
            
            print(f"✅ Sent to {printer_label}: {os.path.basename(file_path)}")
            return True
        except Exception as e:
            print(f"❌ Error printing to {printer_label}: {e}")
            return False
    
    def get_status(self):
        """Get printer status"""
        return {
            'printer1': self.printer1,
            'printer2': self.printer2,
            'both_available': bool(self.printer1 and self.printer2)
        }


def print_order(file_path):
    """Simple function to print order to both printers"""
    printer = DualPrinter()
    return printer.print_to_both(file_path)


if __name__ == "__main__":
    # Test the printer
    printer = DualPrinter()
    status = printer.get_status()
    
    print(f"Printer 1: {status['printer1']}")
    print(f"Printer 2: {status['printer2']}")
    print(f"Both Available: {status['both_available']}")
    
    if status['both_available']:
        test_file = Path(__file__).parent.parent / "orders" / "wempy_order_12.docx"
        printer.print_to_both(str(test_file))
    else:
        print("❌ Both printers not available for testing")
