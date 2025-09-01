from pathlib import Path

def get_next_order_id(root_path: Path) -> int:
    ORDERS_DIR = root_path
    ORDER_ID_FILE = ORDERS_DIR / "last_id.txt"
    ORDERS_DIR.mkdir(exist_ok=True)

    last_id = 0
    if ORDER_ID_FILE.exists():
        content = ORDER_ID_FILE.read_text(encoding="utf-8").strip()
        if content.isdigit():
            last_id = int(content)
    new_id = last_id + 1
    ORDER_ID_FILE.write_text(str(new_id), encoding="utf-8")
    return new_id

def build_order_docx_path(order_id: int, root_path: Path) -> Path:
    ORDERS_DIR = root_path
    return ORDERS_DIR / f"wempy_order_{order_id}.docx"