import time
from pathlib import Path

from printer_manager import DualPrinter
from supabase_manager import download_order

# --- Configuration ---
POLLING_INTERVAL_SECONDS = 10
STATE_FILE = Path("./last_order.txt")

# --- State Management ---
def load_last_order_id() -> int:
    """Reads the last successfully processed order ID from the state file."""
    if not STATE_FILE.exists():
        print("State file not found. Starting from order ID 1.")
        return 1
    try:
        return int(STATE_FILE.read_text())
    except (ValueError, FileNotFoundError):
        print("Error reading state file. Starting from order ID 1.")
        return 1

def save_last_order_id(order_id: int):
    """Saves the last successfully processed order ID to the state file."""
    STATE_FILE.write_text(str(order_id))

# --- Main Application Loop ---
def main():
    """The main function to run the polling client."""
    print("--- Wempy Order Polling Client ---")
    local_orders_dir = Path("./temp_orders")
    local_orders_dir.mkdir(exist_ok=True)

    try:
        printer = DualPrinter()
        if not printer.printer1 or not printer.printer2:
            print("\nExiting: Printers not ready.")
            return
    except Exception as e:
        print(f"\nExiting: Failed to initialize printers: {e}")
        return

    last_known_id = load_last_order_id()
    print(f"\n🚀 Starting polling from Order ID: {last_known_id + 1}\n")

    while True:
        try:
            next_order_id = last_known_id + 1
            print(f"Checking for Order ID: {next_order_id}...", end="", flush=True)

            downloaded_path = download_order(next_order_id, local_orders_dir)

            if downloaded_path:
                print_success = printer.print_file(downloaded_path)
                if print_success:
                    print(f"✅ Successfully processed Order ID: {next_order_id}")
                    last_known_id = next_order_id
                    save_last_order_id(last_known_id)
                else:
                    print(f"❌ Printing failed for Order ID: {next_order_id}. Will retry on next cycle.")
                print("")  # Newline for cleaner logging
            else:
                print(" No new order found.", end="\r")
                time.sleep(POLLING_INTERVAL_SECONDS)

        except Exception as e:
            print(f"\nAn unexpected error occurred in the main loop: {e}")
            print("Restarting loop after a short delay...")
            time.sleep(15)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n👋 Exiting client.")


if __name__ == "__main__":
    main()
