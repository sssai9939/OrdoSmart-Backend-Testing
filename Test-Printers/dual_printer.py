import win32print
import win32api
import os
from pathlib import Path
import time

class DualPrinter:
    """كلاس للطباعة على طابعتين XPrinter"""
    
    def __init__(self):
        self.printers = self.find_all_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None
        
    def find_all_xprinters(self):
        """البحث عن جميع طابعات XPrinter"""
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        xprinters = []
        for printer in printers:
            if 'XP' in printer and 'copy' in printer.lower():
                xprinters.append(printer)
        return sorted(xprinters)  # ترتيب للحصول على Copy 1, Copy 2
    
    def list_available_printers(self):
        """عرض الطابعات المتاحة"""
        print("Available XPrinters:")
        for i, printer in enumerate(self.printers, 1):
            print(f"{i}. {printer}")
        print(f"Printer 1: {self.printer1}")
        print(f"Printer 2: {self.printer2}")
    
    def print_to_printer1(self, file_path):
        """طباعة على الطابعة الأولى"""
        return self._print_file(file_path, self.printer1, "Printer 1")
    
    def print_to_printer2(self, file_path):
        """طباعة على الطابعة الثانية"""
        return self._print_file(file_path, self.printer2, "Printer 2")
    
    def print_to_both(self, file_path):
        """طباعة على كلا الطابعتين"""
        print(f"Printing to both printers: {os.path.basename(file_path)}")
        
        result1 = self.print_to_printer1(file_path)
        time.sleep(0.5)  # توقف قصير بين الطباعتين
        result2 = self.print_to_printer2(file_path)
        
        return result1 and result2
    
    def _print_file(self, file_path, printer_name, printer_label):
        """وظيفة مساعدة للطباعة"""
        if not printer_name:
            print(f"❌ {printer_label} not available")
            return False
            
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            return False
        
        try:
            # حفظ الطابعة الافتراضية الحالية
            current_default = win32print.GetDefaultPrinter()
            
            # تعيين الطابعة المحددة كافتراضية مؤقتاً
            win32print.SetDefaultPrinter(printer_name)
            print(f"🔄 Set default printer to: {printer_name}")
            
            # الطباعة
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            
            # إعادة الطابعة الافتراضية الأصلية
            time.sleep(1)  # انتظار قصير لضمان إرسال الطلب
            win32print.SetDefaultPrinter(current_default)
            print(f"🔄 Restored default printer to: {current_default}")
            
            print(f"✅ File sent to {printer_label} ({printer_name}): {os.path.basename(file_path)}")
            return True
        except Exception as e:
            print(f"❌ Error printing to {printer_label}: {e}")
            return False


if __name__ == "__main__":
    # إنشاء كائن الطابعة المزدوجة
    dual_printer = DualPrinter()
    
    # عرض الطابعات المتاحة
    dual_printer.list_available_printers()
    print("-" * 50)
    
    # مسارات الملفات
    base_path = Path(__file__).parent.parent / "orders"
    file1 = base_path / "wempy_order_10.docx"
    file2 = base_path / "wempy_order_11.docx"
    shared_file = base_path / "wempy_order_12.docx"
    
    # اختبار مبسط - ملف واحد على الطابعتين
    test_file = base_path / "wempy_order_12.docx"
    
    print("🖨️ Simple Test: Print one file to both printers")
    print(f"File: {test_file.name}")
    print("-" * 30)
    
    dual_printer.print_to_both(str(test_file))
