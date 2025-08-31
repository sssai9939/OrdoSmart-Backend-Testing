import os
import time
from pathlib import Path
from dotenv import load_dotenv
import win32print
import win32api
from supabase import create_client, Client

# --- Load Environment Variables & Initialize Supabase ---
load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
BUCKET_NAME = os.getenv("BUCKET_NAME")

if not all([SUPABASE_URL, SUPABASE_KEY, BUCKET_NAME]):
    print("❌ Critical Error: SUPABASE_URL, SUPABASE_KEY, and BUCKET_NAME must be set in .env file.")
    exit()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- Configuration ---
POLLING_INTERVAL_SECONDS = 10
STATE_FILE = Path("./last_order.txt")

# --- Printer Class (Same as before) ---
class DualPrinter:
    def __init__(self):
        self.printers = self._find_xprinters()
        self.printer1 = self.printers[0] if len(self.printers) > 0 else None
        self.printer2 = self.printers[1] if len(self.printers) > 1 else None
        if self.printer1 and self.printer2:
            print(f"✅ Printer 1 (Kitchen): {self.printer1}")
            print(f"✅ Printer 2 (Cashier): {self.printer2}")
        else:
            print(f"❌ Error: Could not find both 'XP-80C' printers.")

    def _find_xprinters(self):
        printers = [p[2] for p in win32print.EnumPrinters(2)]
        xprinters = [p for p in printers if 'XP-80C' in p]
        return sorted(xprinters)

    def print_file(self, file_path):
        if not self.printer1 or not self.printer2:
            print("❌ Error: One or both printers are not found. Cannot print.")
            return False
        
        print_success1 = self._print_to_printer(file_path, self.printer1, "Kitchen")
        print_success2 = self._print_to_printer(file_path, self.printer2, "Cashier")
        return print_success1 and print_success2

    def _print_to_printer(self, file_path, printer_name, printer_label):
        try:
            current_default = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)
            win32api.ShellExecute(0, "print", str(file_path), None, ".", 0)
            time.sleep(2) # Increased wait time for stability
            win32print.SetDefaultPrinter(current_default)
            print(f"✅ Sent to {printer_label} ({printer_name})")
            return True
        except Exception as e:
            print(f"❌ Error printing to {printer_label}: {e}")
            return False

# --- State Management ---
def get_next_order_id() -> int:
    if not STATE_FILE.exists():
        # If you want to start from a specific order, change the '1' here.
        print("State file not found. Starting from order ID 1.")
        return 1
    try:
        last_id = int(STATE_FILE.read_text().strip())
        return last_id + 1
    except (ValueError, FileNotFoundError):
        print("Error reading state file. Starting from order ID 1.")
        return 1

def save_last_order_id(order_id: int):
    STATE_FILE.write_text(str(order_id))

# --- Supabase Interaction ---
def download_order(order_id: int, local_dir: Path) -> Path | None:
    file_name = f"wempy_order_{order_id}.docx"
    local_path = local_dir / file_name
    
    try:
        res = supabase.storage.from_(BUCKET_NAME).download(file_name)
        with open(local_path, "wb+") as f:
            f.write(res)
        return local_path
    except Exception as e:
        # This is expected when an order doesn't exist yet.
        # We only log other unexpected errors.
        if 'Not found' not in str(e):
             print(f"\nAn unexpected error occurred while downloading order {order_id}: {e}")
        return None

# --- Main Application Loop ---
def main():
    print("--- Wempy Order Polling Client ---")
    local_orders_dir = Path("./temp_orders")
    local_orders_dir.mkdir(exist_ok=True)

    try:
        printer = DualPrinter()
        if not printer.printer1 or not printer.printer2:
            print("\n--- WAITING FOR PRINTERS ---")
            print("Please ensure both XPrinter devices are connected and recognized. Exiting.")
            return
    except Exception as e:
        print(f"❌ Could not initialize printer object. Error: {e}")
        return

    next_order_id = get_next_order_id()
    print(f"\n🚀 Starting polling from Order ID: {next_order_id}")

    while True:
        try:
            print(f"\nChecking for Order ID: {next_order_id}...")
            downloaded_path = download_order(next_order_id, local_orders_dir)

            if downloaded_path:
                print(f"✅ Found and downloaded {downloaded_path.name}")
                print_success = printer.print_file(downloaded_path)
                
                if print_success:
                    save_last_order_id(next_order_id)
                    print(f"✅ Successfully processed Order ID: {next_order_id}")
                    next_order_id += 1
                    # Don't wait, immediately check for the next order
                    continue 
                else:
                    print(f"⚠️ Failed to print Order ID: {next_order_id}. Will retry.")

            else:
                print(f"No new order found. Waiting for {POLLING_INTERVAL_SECONDS} seconds...")
            
            time.sleep(POLLING_INTERVAL_SECONDS)

        except KeyboardInterrupt:
            print("\nShutting down client...")
            break
        except Exception as e:
            print(f"An unexpected error occurred in the main loop: {e}")
            print("Restarting loop after a short delay...")
            time.sleep(15)

if __name__ == "__main__":
    main()
