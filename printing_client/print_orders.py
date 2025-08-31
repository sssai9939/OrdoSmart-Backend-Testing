import os
import time
from pathlib import Path
from dotenv import load_dotenv
import win32print
import win32api
from supabase import create_client, Client

# --- Load Environment Variables ---
print("Loading environment variables...")
load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
BUCKET_NAME = os.getenv("BUCKET_NAME")

# --- Validate Environment Variables ---
if not all([SUPABASE_URL, SUPABASE_KEY, BUCKET_NAME]):
    print("❌ Critical Error: SUPABASE_URL, SUPABASE_KEY, and BUCKET_NAME must be set in .env file.")
    exit()

print("Initializing Supabase client...")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- Printer Class (Windows Only) ---
class DualPrinter:
    def __init__(self):
        self.printers = self._find_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None
        print(f"Printer 1 (Kitchen): {self.printer1 or 'Not Found'}")
        print(f"Printer 2 (Cashier): {self.printer2 or 'Not Found'}")

    def _find_xprinters(self):
        printers = [p[2] for p in win32print.EnumPrinters(2)]
        return sorted([p for p in printers if 'xprinter' in p.lower() or 'xp-' in p.lower()])

    def print_file(self, file_path):
        if not self.printer1 or not self.printer2:
            print("❌ Both printers must be available to print.")
            return False
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            return False
        
        print(f"\n🖨️ Printing to both printers: {os.path.basename(file_path)}")
        result1 = self._print_to_printer(file_path, self.printer1, "Printer 1")
        time.sleep(0.5)
        result2 = self._print_to_printer(file_path, self.printer2, "Printer 2")
        return result1 and result2

    def _print_to_printer(self, file_path, printer_name, printer_label):
        try:
            current_default = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            time.sleep(1) # Wait for the print job to be sent
            win32print.SetDefaultPrinter(current_default)
            print(f"✅ Sent to {printer_label} ({printer_name})")
            return True
        except Exception as e:
            print(f"❌ Error printing to {printer_label}: {e}")
            return False

# --- Supabase Interaction ---
def download_order(order_id: int, local_dir: Path) -> Path | None:
    file_name = f"wempy_order_{order_id}.docx"
    local_path = local_dir / file_name
    
    print(f"Downloading {file_name} from Supabase...")
    try:
        with open(local_path, "wb+") as f:
            res = supabase.storage.from_(BUCKET_NAME).download(file_name)
            f.write(res)
        print(f"✅ Downloaded successfully to {local_path}")
        return local_path
    except Exception as e:
        print(f"❌ Failed to download order {order_id}. Error: {e}")
        return None

# --- Realtime Callback ---
def handle_new_order(payload):
    print(f"\n🔔 New Order Received! Payload: {payload}")
    record = payload.get('new', {})
    order_id = record.get('id')
    
    if not order_id:
        print("⚠️ Received payload without a valid order ID.")
        return

    print(f"Processing Order ID: {order_id}")
    # It's good practice to create a fresh printer and dir object inside the callback
    # to ensure thread safety and fresh state, although not strictly necessary here.
    local_orders_dir = Path("./temp_orders")
    printer = DualPrinter()

    downloaded_path = download_order(order_id, local_orders_dir)
    if downloaded_path:
        printer.print_file(str(downloaded_path))
        # Optional: Clean up the downloaded file after printing
        # os.remove(downloaded_path)

# --- Main Application ---
def main():
    print("--- Wempy Realtime Printing Client ---")
    local_orders_dir = Path("./temp_orders")
    local_orders_dir.mkdir(exist_ok=True)

    try:
        printer_check = DualPrinter()
        if not printer_check.printer1 or not printer_check.printer2:
             print("\n--- WAITING FOR PRINTERS ---")
             print("Please ensure both XPrinter devices are connected and recognized.")
             return
    except Exception as e:
        print(f"❌ Could not initialize printer object. Error: {e}")
        print("Please ensure 'pywin32' is installed and printers are connected.")
        return

    print("\n🚀 Connecting to Supabase Realtime...")
    realtime_channel = supabase.realtime
    channel = realtime_channel.channel('public:orders')
    channel.on(
        'postgres_changes',
        filter={'event': 'INSERT', 'schema': 'public', 'table': 'orders'},
        callback=handle_new_order
    ).subscribe()

    print("✅ Connected! Listening for new orders...")
    try:
        # Keep the script running to listen for events
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down listener...")
        # It's good practice to unsubscribe, though not strictly required
        # channel.unsubscribe() # This might need proper async handling
        print("Goodbye!")

if __name__ == "__main__":
    main()
